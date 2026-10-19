VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStateSolutionRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Решение НДС для заданных N + Mx + My
' ==========================================================================
' Настраивает CSectionSolver и получает начальное приближение через
' CStateGuessBuilder. При несходимости выполняет предусмотренные повторные
' старты, в том числе с разрешённым техническим продолжением диаграмм.
' Возвращает сходимость, использование продолжения и проверку физических
' пределов. Модели материалов заранее готовит CMaterialModelProvider.

Private Const NEAR_ZERO As Double = 0.000000001 ' Общий допуск для распознавания почти нулевого усилия/момента.

Private mSolverMethod As String ' Newton или Secant для прямого равновесия.
Private mLoadSteps As Long ' Число ступеней приложения нагрузки.
Private mMaxIterations As Long ' Максимум итераций CSectionSolver на одной ступени.
Private mToleranceN As Double ' Допуск равновесия по N, Н.
Private mToleranceMx As Double ' Допуск равновесия по Mx, Н*мм.
Private mToleranceMy As Double ' Допуск равновесия по My, Н*мм.
Private mSecantMaxRestarts As Long ' Лимит перезапуска матрицы в secant-ветке.
Private mDiagnosticsEnabled As Boolean ' Нужно ли копить подробный журнал итераций.

Private mSolver As CSectionSolver ' Объект последней попытки решения НДС, в том числе неуспешной.
Private mExtensionUsed As Boolean ' Признак использования вспомогательного продолжения диаграмм в найденном НДС.
Private mWithinPhysicalRange As Boolean ' True, если все элементы остались в физических eps-пределах.
Private mDiagnosticLog As String ' Краткий журнал повторных попыток решения НДС.
Private mInitialGuessApplied As Boolean ' True, если стартовая плоскость была задана не нулевой.
Private mSolverCallCount As Long ' Количество фактических вызовов CSectionSolver.Solve.
Private mRetryAttemptFailures As Collection ' Краткие итоги точно совпадающих неудачных повторных попыток в пределах одного поиска.
Private mRetryContextSection As CSectionModel ' Модель и материалы текущего поиска для проверки применимости сохранённых попыток; конечные НДС здесь не хранятся.
Private mRetryContextConcrete As Object ' Свойства диаграммы бетона для проверки точного совпадения повторной попытки.
Private mRetryContextSteel As Object ' Свойства диаграммы арматуры для проверки точного совпадения повторной попытки.
Private mRetrySectionRevision As Long ' Изменение модели запрещает использование прежних попыток.

' Задает безопасные настройки по умолчанию для unit-тестов и вызовов без Config.
Private Sub Class_Initialize()
    mSolverMethod = "Newton"
    mLoadSteps = 8
    mMaxIterations = 80
    mToleranceN = 5#
    mToleranceMx = 5000#
    mToleranceMy = 5000#
    mSecantMaxRestarts = 2
    mDiagnosticsEnabled = True
    mWithinPhysicalRange = True
End Sub

' Начинает учёт повторных численных попыток одного поиска на той же модели
' и материалах. Сохраняет лишь диагностику неудач, не конечные НДС.
' Другое начальное поле или численные настройки дают отдельную попытку.
Friend Sub BeginRetryAttemptSession(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object)
    If section Is Nothing Then Err.Raise vbObjectError + 3483, "CStateSolutionRunner", "Для повторных попыток решения НДС не передана модель сечения."
    If concreteMaterial Is Nothing Then Err.Raise vbObjectError + 3484, "CStateSolutionRunner", "Для повторных попыток решения НДС не передана диаграмма бетона."
    If steelMaterial Is Nothing Then Err.Raise vbObjectError + 3485, "CStateSolutionRunner", "Для повторных попыток решения НДС не передана диаграмма арматуры."
    Set mRetryContextSection = section
    Set mRetryContextConcrete = concreteMaterial
    Set mRetryContextSteel = steelMaterial
    mRetrySectionRevision = section.Revision
    Set mRetryAttemptFailures = New Collection
End Sub

' Восстанавливает данные ранее найденного НДС без нового решения равновесия.
' Проверяет совпадение модели, материалов и целевых усилий;
' счётчик запусков решателя не увеличивается.
Public Function RestoreSnapshot(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal stateResult As CSectionStateResult) As CSectionSolver
    If stateResult Is Nothing Then Exit Function
    If Not stateResult.Converged Then Exit Function

    Dim snapshot As CSectionSolver
    Set snapshot = New CSectionSolver
    ConfigureSolver snapshot
    snapshot.EvaluateStrainPlane section, concreteMaterial, steelMaterial, _
        stateResult.Epsilon0, stateResult.KappaX, stateResult.KappaY
    If snapshot.ConfirmEquilibrium(stateResult.TargetN, stateResult.TargetMx, stateResult.TargetMy) Then
        Set RestoreSnapshot = snapshot
    End If
End Function

' Решает прямое равновесие от заданного LC.
' allowExtensionWarmStart означает не сам факт включения extension в материалах,
' а разрешение на вторую серию стартов, если первая попытка CSectionSolver не
' нашла равновесие. Сами диаграммы при этом уже переданы снаружи.
Public Sub Solve(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        Optional ByVal allowExtensionWarmStart As Boolean = False)
    SolveWithInitialSolver section, concreteMaterial, steelMaterial, _
        targetN, targetMx, targetMy, Nothing, allowExtensionWarmStart
End Sub

' Решает заданную нагрузку, начиная с близкого найденного НДС.
' Так получаются, например, состояния после трещины с общими
' настройками, повторными попытками и диагностикой обычного расчёта.
Public Sub SolveWithInitialSolver(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal initialSolver As CSectionSolver, _
        Optional ByVal allowExtensionWarmStart As Boolean = False)
    If section Is Nothing Then Err.Raise vbObjectError + 3480, "CStateSolutionRunner", "Модель сечения не передана."
    If concreteMaterial Is Nothing Then Err.Raise vbObjectError + 3481, "CStateSolutionRunner", "Диаграмма бетона не передана."
    If steelMaterial Is Nothing Then Err.Raise vbObjectError + 3482, "CStateSolutionRunner", "Диаграмма арматуры не передана."

    ClearResult

    Dim service As CSectionSolver
    Set service = New CSectionSolver
    ConfigureSolver service
    ApplyInitialSolverOrGuess section, concreteMaterial, steelMaterial, _
        targetN, targetMx, targetMy, initialSolver, service

    mSolverCallCount = mSolverCallCount + 1
    service.Solve section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy
    Set mSolver = service

    If CanRetrySolver(service) Then
        If TrySolveWithDirectStateGuess(section, concreteMaterial, steelMaterial, _
                targetN, targetMx, targetMy, service) Then
            Set mSolver = service
        End If
    End If

    If CanRetrySolver(service) And allowExtensionWarmStart Then
        If TrySolveWithExtensionWarmStart(section, concreteMaterial, steelMaterial, _
                targetN, targetMx, targetMy, service) Then
            Set mSolver = service
        End If
    End If

    Set mSolver = service
    If Not mSolver Is Nothing Then
        If mSolver.Converged Then
            mWithinPhysicalRange = StateWithinPhysicalRange(section, mSolver, concreteMaterial, steelMaterial)
            mExtensionUsed = StateUsesExtension(section, mSolver, concreteMaterial, steelMaterial)
        Else
            mWithinPhysicalRange = False
            mExtensionUsed = False
        End If
    End If
End Sub

' Решает НДС с явно указанной стартовой нагрузкой и её полем деформаций.
' При поиске несущей способности это позволяет перейти от соседнего lambda
' к новой нагрузке. Выбор старта не меняет ключ физического запроса
' для повторного использования конечных состояний.
Public Sub SolveWithInitialLoadState(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByVal initialN As Double, ByVal initialMx As Double, ByVal initialMy As Double, _
        Optional ByVal allowExtensionWarmStart As Boolean = False)
    If section Is Nothing Then Err.Raise vbObjectError + 3483, "CStateSolutionRunner", "Модель сечения не передана."
    If concreteMaterial Is Nothing Then Err.Raise vbObjectError + 3484, "CStateSolutionRunner", "Диаграмма бетона не передана."
    If steelMaterial Is Nothing Then Err.Raise vbObjectError + 3485, "CStateSolutionRunner", "Диаграмма арматуры не передана."

    ClearResult

    Dim service As CSectionSolver
    Set service = New CSectionSolver
    ConfigureSolver service
    service.SetInitialLoadState eps0, kx, ky, initialN, initialMx, initialMy
    mInitialGuessApplied = True
    AppendDiagnostic "initialGuess=warmStartLoadState"

    mSolverCallCount = mSolverCallCount + 1
    service.Solve section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy
    Set mSolver = service

    If CanRetrySolver(service) Then
        If TrySolveWithDirectStateGuess(section, concreteMaterial, steelMaterial, _
                targetN, targetMx, targetMy, service) Then
            Set mSolver = service
        End If
    End If

    If CanRetrySolver(service) And allowExtensionWarmStart Then
        If TrySolveWithExtensionWarmStart(section, concreteMaterial, steelMaterial, _
                targetN, targetMx, targetMy, service) Then
            Set mSolver = service
        End If
    End If

    Set mSolver = service
    If Not mSolver Is Nothing Then
        If mSolver.Converged Then
            mWithinPhysicalRange = StateWithinPhysicalRange(section, mSolver, concreteMaterial, steelMaterial)
            mExtensionUsed = StateUsesExtension(section, mSolver, concreteMaterial, steelMaterial)
        Else
            mWithinPhysicalRange = False
            mExtensionUsed = False
        End If
    End If
End Sub

' Решает НДС из начального поля, подготовленного сразу для полной целевой нагрузки.
' Нужно для первых проб чистого изгиба, чтобы не проходить от нуля
' через излом диаграммы. Остальные настройки и диагностика остаются общими.
Public Sub SolveWithDirectTargetGuess(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        Optional ByVal diagnosticLine As String = vbNullString, _
        Optional ByVal allowExtensionWarmStart As Boolean = False)
    If section Is Nothing Then Err.Raise vbObjectError + 3486, "CStateSolutionRunner", "Модель сечения не передана."
    If concreteMaterial Is Nothing Then Err.Raise vbObjectError + 3487, "CStateSolutionRunner", "Диаграмма бетона не передана."
    If steelMaterial Is Nothing Then Err.Raise vbObjectError + 3488, "CStateSolutionRunner", "Диаграмма арматуры не передана."

    ClearResult

    Dim service As CSectionSolver
    Set service = New CSectionSolver
    ConfigureSolver service

    Dim concreteDiagram As CMaterialDiagram
    Dim steelDiagram As CMaterialDiagram
    Dim guessBuilder As CStateGuessBuilder
    Dim eps0 As Double
    Dim kx As Double
    Dim ky As Double
    If TryGetMaterialDiagrams(concreteMaterial, steelMaterial, concreteDiagram, steelDiagram) Then
        Set guessBuilder = New CStateGuessBuilder
        If guessBuilder.BuildForDirectState(section, concreteDiagram, steelDiagram, _
                targetN, targetMx, targetMy, eps0, kx, ky) Then
            service.SetInitialLoadState eps0, kx, ky, targetN, targetMx, targetMy
            mInitialGuessApplied = True
            If Len(diagnosticLine) > 0 Then
                AppendDiagnostic diagnosticLine
            Else
                AppendDiagnostic "initialGuess=directTargetActiveSet"
            End If
        End If
    End If

    mSolverCallCount = mSolverCallCount + 1
    service.Solve section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy
    Set mSolver = service

    If CanRetrySolver(service) And allowExtensionWarmStart Then
        If TrySolveWithExtensionWarmStart(section, concreteMaterial, steelMaterial, _
                targetN, targetMx, targetMy, service) Then
            Set mSolver = service
        End If
    End If

    Set mSolver = service
    If Not mSolver Is Nothing Then
        If mSolver.Converged Then
            mWithinPhysicalRange = StateWithinPhysicalRange(section, mSolver, concreteMaterial, steelMaterial)
            mExtensionUsed = StateUsesExtension(section, mSolver, concreteMaterial, steelMaterial)
        Else
            mWithinPhysicalRange = False
            mExtensionUsed = False
        End If
    End If
End Sub

' Ошибку модели, настройки или программы другое начальное поле не исправляет.
' Повтор разрешён только после численной неудачи.
Private Function CanRetrySolver(ByVal solver As CSectionSolver) As Boolean
    If solver Is Nothing Then Exit Function
    If solver.Converged Then Exit Function
    Select Case solver.FailureCode
        Case sfcNumericalFailure, sfcSingularTangent
            CanRetrySolver = True
    End Select
End Function

' После несходимости повторяет НДС с начальным полем по полной нагрузке.
' Это отличается от ступенчатого нагружения с той же формой деформаций.
' Общий способ применяется к изгибу и внецентренному растяжению или сжатию;
' для успешно сошедшихся сочетаний дополнительная попытка не запускается.
Private Function TrySolveWithDirectStateGuess(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByRef service As CSectionSolver) As Boolean
    If mMaxIterations <= 1 Then Exit Function
    If Abs(targetN) + Abs(targetMx) + Abs(targetMy) <= NEAR_ZERO Then Exit Function

    Dim concreteDiagram As CMaterialDiagram
    Dim steelDiagram As CMaterialDiagram
    If Not TryGetMaterialDiagrams(concreteMaterial, steelMaterial, concreteDiagram, steelDiagram) Then Exit Function

    Dim guessBuilder As CStateGuessBuilder
    Set guessBuilder = New CStateGuessBuilder

    Dim eps0 As Double
    Dim kx As Double
    Dim ky As Double
    If Not guessBuilder.BuildForDirectState(section, concreteDiagram, steelDiagram, _
            targetN, targetMx, targetMy, eps0, kx, ky) Then
        Exit Function
    End If

    If TryRunRetry(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
            eps0, kx, ky, targetN, targetMx, targetMy, 1, _
            "directStateActiveSet / direct target", service, "directStateRetry") Then
        mInitialGuessApplied = True
        TrySolveWithDirectStateGuess = True
        Exit Function
    End If

End Function

' Выбирает начальное поле: пригодное близкое НДС, переданное вызывающим расчётом,
' или новую оценку общего CStateGuessBuilder.
Private Sub ApplyInitialSolverOrGuess(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal initialSolver As CSectionSolver, ByVal solver As CSectionSolver)
    If Not initialSolver Is Nothing Then
        If initialSolver.Converged Then
            solver.SetInitialLoadState initialSolver.Epsilon0, initialSolver.KappaX, initialSolver.KappaY, _
                initialSolver.Nint, initialSolver.Mxint, initialSolver.Myint
        Else
            solver.SetInitialState initialSolver.Epsilon0, initialSolver.KappaX, initialSolver.KappaY
        End If
        mInitialGuessApplied = True
        AppendDiagnostic "initialGuess=warmStartSolver"
        Exit Sub
    End If

    ApplyInitialGuess section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, solver
End Sub

' Передаёт настройки новому CSectionSolver. Каждая попытка решения НДС
' использует тот же набор пользовательских допусков.
Private Sub ConfigureSolver(ByVal solver As CSectionSolver)
    solver.LoadSteps = mLoadSteps
    solver.SolverMethod = mSolverMethod
    solver.MaxIterations = mMaxIterations
    solver.ToleranceN = mToleranceN
    solver.ToleranceMx = mToleranceMx
    solver.ToleranceMy = mToleranceMy
    solver.SecantMaxRestarts = mSecantMaxRestarts
    solver.DiagnosticsEnabled = mDiagnosticsEnabled
End Sub

' Строит начальную плоскость для случаев, где нулевая плоскость является плохой
' стартовой точкой. Чаще всего это чистый изгиб: при N=0 первая
' итерация попадает ровно в излом бетонной диаграммы около нуля.
Private Sub ApplyInitialGuess(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal solver As CSectionSolver)
    If Abs(targetN) > NEAR_ZERO Then Exit Sub
    ' Квадраты безопасны только в малой окрестности нуля. Если хотя бы
    ' одна компонента больше допуска, норма заведомо ненулевая; ее вычислять
    ' не нужно. Внутри малой окрестности применяется евклидова норма момента.
    If Abs(targetMx) <= NEAR_ZERO And Abs(targetMy) <= NEAR_ZERO Then
        If Sqr(targetMx * targetMx + targetMy * targetMy) <= NEAR_ZERO Then Exit Sub
    End If

    Dim concreteDiagram As CMaterialDiagram
    Dim steelDiagram As CMaterialDiagram
    If Not TryGetMaterialDiagrams(concreteMaterial, steelMaterial, concreteDiagram, steelDiagram) Then Exit Sub

    Dim guessBuilder As CStateGuessBuilder
    Dim eps0 As Double
    Dim kx As Double
    Dim ky As Double
    Set guessBuilder = New CStateGuessBuilder
    If guessBuilder.BuildForDirectState(section, concreteDiagram, steelDiagram, _
            targetN, targetMx, targetMy, eps0, kx, ky) Then
        solver.SetInitialState eps0, kx, ky
        mInitialGuessApplied = True
        AppendDiagnostic "initialGuess=directStateActiveSet"
    End If
End Sub

' Перебирает поля для повторных попыток на численном продолжении диаграмм.
' CStateGuessBuilder готовит epsilon0/kappaX/kappaY; этот класс
' запускает решатель и проверяет полученное равновесие.
Private Function TrySolveWithExtensionWarmStart(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByRef service As CSectionSolver) As Boolean
    Dim concreteDiagram As CMaterialDiagram
    Dim steelDiagram As CMaterialDiagram
    If Not TryGetMaterialDiagrams(concreteMaterial, steelMaterial, concreteDiagram, steelDiagram) Then Exit Function

    Dim guessBuilder As CStateGuessBuilder
    Set guessBuilder = New CStateGuessBuilder

    Dim eps0 As Double
    Dim kx As Double
    Dim ky As Double
    Dim attemptIndex As Long

    For attemptIndex = 1 To guessBuilder.ExtensionAttemptCount
        If guessBuilder.BuildAttemptForExtensionState(attemptIndex, section, service, _
                concreteDiagram, steelDiagram, targetN, targetMx, targetMy, eps0, kx, ky) Then
            If TrySolveFromExtensionGuess(section, concreteMaterial, steelMaterial, _
                    targetN, targetMx, targetMy, eps0, kx, ky, _
                    guessBuilder.ExtensionAttemptName(attemptIndex), service) Then
                TrySolveWithExtensionWarmStart = True
                Exit Function
            End If
        End If
    Next attemptIndex
End Function

' Проверяет одно начальное поле: сначала при полной нагрузке, затем при неудаче
' прикладывает нагрузку ступенями от усилий этого поля. Величину ступени
' определяет CSectionSolver.
Private Function TrySolveFromExtensionGuess(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByVal attemptName As String, ByRef service As CSectionSolver) As Boolean
    Dim startState As CSectionSolver
    Set startState = New CSectionSolver
    startState.EvaluateStrainPlane section, concreteMaterial, steelMaterial, eps0, kx, ky

    Dim retryLoadSteps As Long
    retryLoadSteps = mLoadSteps
    If retryLoadSteps < 4 Then retryLoadSteps = 4

    If TryRunRetry(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
            eps0, kx, ky, targetN, targetMx, targetMy, 1, _
            attemptName & " / direct target", service) Then
        TrySolveFromExtensionGuess = True
        Exit Function
    End If

    TrySolveFromExtensionGuess = TryRunRetry(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
        eps0, kx, ky, startState.Nint, startState.Mxint, startState.Myint, retryLoadSteps, _
        attemptName & " / stepped", service)
End Function

' Повторяет решение из подготовленного поля и нагрузок initialN/Mx/My.
' От этого уровня CSectionSolver прикладывает ступени до цели.
' Это позволяет начать с близкого состояния на продолжении диаграммы
' вместо нулевой нагрузки с несоответствующей ей формой деформаций.
Private Function TryRunRetry(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByVal initialN As Double, ByVal initialMx As Double, ByVal initialMy As Double, _
        ByVal loadSteps As Long, ByVal attemptName As String, ByRef service As CSectionSolver, _
        Optional ByVal retryKind As String = "extensionRetry") As Boolean
    Dim retry As CSectionSolver
    Set retry = New CSectionSolver
    ConfigureSolver retry
    retry.LoadSteps = loadSteps
    retry.SetInitialLoadState eps0, kx, ky, initialN, initialMx, initialMy
    Dim attemptKey As Variant, previousFailure As Variant
    If RetryAttemptContextMatches(section, concreteMaterial, steelMaterial) Then
        attemptKey = Array(targetN, targetMx, targetMy, eps0, kx, ky, initialN, initialMx, initialMy, _
            loadSteps, mMaxIterations, mToleranceN, mToleranceMx, mToleranceMy, _
            mSolverMethod, mSecantMaxRestarts)
        If TryRestoreFailedRetryAttempt(attemptKey, previousFailure) Then
            If mDiagnosticsEnabled Then
                AppendDiagnostic retryKind & "=" & attemptName & _
                    "; идентичная неудачная попытка уже выполнена в этом поиске; stop=" & CStr(previousFailure(2))
            End If
            Exit Function
        End If
    End If
    mSolverCallCount = mSolverCallCount + 1
    retry.Solve section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy

    If retry.Converged Then
        Set service = retry
        If mDiagnosticsEnabled Then AppendDiagnostic retryKind & "=" & attemptName & "; result=converged"
        TryRunRetry = True
    Else
        If IsArray(attemptKey) Then StoreFailedRetryAttempt attemptKey, retry
        If mDiagnosticsEnabled Then
            AppendDiagnostic retryKind & "=" & attemptName & _
                "; result=failed; stop=" & retry.StopReason & _
                "; residualN=" & FormatNumberInvariant(retry.ResidualN) & _
                "; residualMx=" & FormatNumberInvariant(retry.ResidualMx) & _
                "; residualMy=" & FormatNumberInvariant(retry.ResidualMy)
        End If
        If FailedRetryImprovesAllResiduals(retry, service) Then
            ' Сохраняем более близкое поле только для следующих стартов.
            ' Converged остаётся False; подтверждённого НДС для общего хранилища нет.
            Set service = retry
            AppendDiagnostic retryKind & "=" & attemptName & "; сохранена более близкая несошедшаяся плоскость."
        End If
    End If
End Function

' Сравнивает только фактические невязки численно неуспешных попыток. Новый
' старт принимается, если не ухудшил ни одной компоненты и улучшил хотя бы одну.
' Деления/суммы норм здесь нет: сравнение не создает переполнения и не ослабляет
' критерий равновесия. Terminal причины и непригодные числа не заменяются.
Private Function FailedRetryImprovesAllResiduals(ByVal candidate As CSectionSolver, _
        ByVal previous As CSectionSolver) As Boolean
    If Not CanRetrySolver(candidate) Then Exit Function
    If Not CanRetrySolver(previous) Then Exit Function
    If Not RetryResidualIsComparable(candidate.ResidualN) Or Not RetryResidualIsComparable(candidate.ResidualMx) Or _
            Not RetryResidualIsComparable(candidate.ResidualMy) Then Exit Function
    If Not RetryResidualIsComparable(previous.ResidualN) Or Not RetryResidualIsComparable(previous.ResidualMx) Or _
            Not RetryResidualIsComparable(previous.ResidualMy) Then Exit Function
    If Abs(candidate.ResidualN) > Abs(previous.ResidualN) Then Exit Function
    If Abs(candidate.ResidualMx) > Abs(previous.ResidualMx) Then Exit Function
    If Abs(candidate.ResidualMy) > Abs(previous.ResidualMy) Then Exit Function
    FailedRetryImprovesAllResiduals = Abs(candidate.ResidualN) < Abs(previous.ResidualN) Or _
        Abs(candidate.ResidualMx) < Abs(previous.ResidualMx) Or Abs(candidate.ResidualMy) < Abs(previous.ResidualMy)
End Function

' Отклоняет непредставимую или чрезмерную невязку при выборе нового старта.
' Эта численная защита не подтверждает сходимость или физический предел.
Private Function RetryResidualIsComparable(ByVal value As Double) As Boolean
    RetryResidualIsComparable = (value = value And Abs(value) < 1E+100)
End Function

' Проверяет модель, её версию и материалы перед чтением прежних повторных попыток.
' При несовпадении завершает их учёт, но разрешает обычное решение новой задачи.
Private Function RetryAttemptContextMatches(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object) As Boolean
    If mRetryAttemptFailures Is Nothing Then Exit Function
    If Not section Is mRetryContextSection Or Not concreteMaterial Is mRetryContextConcrete Or _
            Not steelMaterial Is mRetryContextSteel Then
        Set mRetryAttemptFailures = Nothing
        Exit Function
    End If
    If section.Revision <> mRetrySectionRevision Then
        Set mRetryAttemptFailures = Nothing
        Exit Function
    End If
    RetryAttemptContextMatches = True
End Function

' Сравнивает все численные параметры попытки без округления и без lambda-only
' ключа. Запись доказывает только повтор той же неудачной процедуры, не разрушение.
Private Function TryRestoreFailedRetryAttempt(ByVal attemptKey As Variant, ByRef failureRecord As Variant) As Boolean
    Dim record As Variant, storedKey As Variant, index As Long, sameAttempt As Boolean
    For Each record In mRetryAttemptFailures
        storedKey = record(0)
        sameAttempt = True
        For index = LBound(attemptKey) To UBound(attemptKey)
            If attemptKey(index) <> storedKey(index) Then
                sameAttempt = False
                Exit For
            End If
        Next index
        If sameAttempt Then
            failureRecord = record
            TryRestoreFailedRetryAttempt = True
            Exit Function
        End If
    Next record
End Function

' Сохраняет краткую диагностику выполненной неудачной попытки без решателя
' и массивов элементов. Число записей ограничено. Окончательные ошибки
' и физические НДС здесь не сохраняются; следующий независимый поиск
' получает новый объект запуска.
Private Sub StoreFailedRetryAttempt(ByVal attemptKey As Variant, ByVal retry As CSectionSolver)
    If retry.FailureCode <> sfcNumericalFailure And retry.FailureCode <> sfcSingularTangent Then Exit Sub
    If mRetryAttemptFailures.Count >= 64 Then mRetryAttemptFailures.Remove 1
    mRetryAttemptFailures.Add Array(attemptKey, retry.FailureCode, retry.StopReason)
End Sub

' Проверяет, использует ли хотя бы один элемент добавленное численное
' продолжение диаграммы. Признак отдельно сохраняется в Results
' для схемы, AutoCAD и расчёта трещин.
Public Function StateUsesExtension(ByVal section As CSectionModel, ByVal solver As CSectionSolver, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object) As Boolean
    If section Is Nothing Then Exit Function
    If solver Is Nothing Then Exit Function
    If Not solver.Converged Then Exit Function
    Dim concreteDiagram As CMaterialDiagram
    Dim steelDiagram As CMaterialDiagram
    If TypeName(concreteMaterial) = "CMaterialDiagram" Then Set concreteDiagram = concreteMaterial
    If TypeName(steelMaterial) = "CMaterialDiagram" Then Set steelDiagram = steelMaterial

    Dim i As Long
    Dim strain As Double

    If Not concreteDiagram Is Nothing Then
        For i = 1 To section.ConcreteCount
            strain = ElementStrain(solver, section.ConcreteX(i), section.ConcreteY(i))
            If concreteDiagram.IsInExtensionRange(strain) Then
                StateUsesExtension = True
                Exit Function
            End If
        Next i
    End If

    If Not steelDiagram Is Nothing Then
        For i = 1 To section.RebarCount
            strain = ElementStrain(solver, section.RebarX(i), section.RebarY(i))
            If steelDiagram.IsInExtensionRange(strain) Then
                StateUsesExtension = True
                Exit Function
            End If
        Next i
    End If
End Function

' Проверяет физические eps-пределы всех бетонных волокон и стержней.
' Если растянутая ветвь бетона отключена и у диаграммы нет растянутого
' предела, положительные деформации бетона считаются допустимыми для проверки
' range, но дают нулевое напряжение через саму диаграмму.
Public Function StateWithinPhysicalRange(ByVal section As CSectionModel, ByVal solver As CSectionSolver, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object) As Boolean
    If section Is Nothing Then Exit Function
    If solver Is Nothing Then Exit Function
    If Not solver.Converged Then Exit Function
    If concreteMaterial Is Nothing Then Exit Function
    If steelMaterial Is Nothing Then Exit Function
    Dim concreteDiagram As CMaterialDiagram
    Dim steelDiagram As CMaterialDiagram
    If TypeName(concreteMaterial) = "CMaterialDiagram" Then Set concreteDiagram = concreteMaterial
    If TypeName(steelMaterial) = "CMaterialDiagram" Then Set steelDiagram = steelMaterial

    Dim i As Long
    Dim strain As Double

    If Not concreteDiagram Is Nothing Then
        For i = 1 To section.ConcreteCount
            strain = ElementStrain(solver, section.ConcreteX(i), section.ConcreteY(i))
            If Not concreteDiagram.IsInPhysicalRange(strain) Then Exit Function
        Next i
    End If

    If Not steelDiagram Is Nothing Then
        For i = 1 To section.RebarCount
            strain = ElementStrain(solver, section.RebarX(i), section.RebarY(i))
            If Not steelDiagram.IsInPhysicalRange(strain) Then Exit Function
        Next i
    End If

    StateWithinPhysicalRange = True
End Function

' Проверяет наличие технических ветвей в фактически переданных материалах.
' Численные повторные старты не разрешают продолжение сами по себе: при Off
' и для линейного Object material API этот метод всегда возвращает False.
Public Function MaterialsHaveExtension(ByVal concreteMaterial As Object, ByVal steelMaterial As Object) As Boolean
    Dim concreteDiagram As CMaterialDiagram
    Dim steelDiagram As CMaterialDiagram
    If TypeName(concreteMaterial) = "CMaterialDiagram" Then
        Set concreteDiagram = concreteMaterial
        MaterialsHaveExtension = concreteDiagram.HasCompressionExtension Or concreteDiagram.HasTensionExtension
    End If
    If TypeName(steelMaterial) = "CMaterialDiagram" Then
        Set steelDiagram = steelMaterial
        MaterialsHaveExtension = MaterialsHaveExtension Or _
            steelDiagram.HasCompressionExtension Or steelDiagram.HasTensionExtension
    End If
End Function

' Возвращает CMaterialDiagram, если переданы именно такие материалы.
' Контрольные линейные материалы решаются тем же CSectionSolver,
' но у них нет физических пределов для специальных стартов на продолжении диаграмм.
Private Function TryGetMaterialDiagrams(ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByRef concreteDiagram As CMaterialDiagram, ByRef steelDiagram As CMaterialDiagram) As Boolean
    If concreteMaterial Is Nothing Or steelMaterial Is Nothing Then Exit Function
    If TypeName(concreteMaterial) <> "CMaterialDiagram" Then Exit Function
    If TypeName(steelMaterial) <> "CMaterialDiagram" Then Exit Function

    Set concreteDiagram = concreteMaterial
    Set steelDiagram = steelMaterial
    TryGetMaterialDiagrams = True
End Function

' Вычисляет деформацию произвольной точки по найденным epsilon0/kappaX/kappaY
' для проверки диаграмм после решения.
Private Function ElementStrain(ByVal solver As CSectionSolver, ByVal x As Double, ByVal y As Double) As Double
    ElementStrain = solver.StrainAt(x, y)
End Function

' Добавляет строку в диагностический журнал только когда диагностика включена.
' Так рабочий расчет не тратит время на лишний текстовый вывод.
Private Sub AppendDiagnostic(ByVal lineText As String)
    If Not mDiagnosticsEnabled Then Exit Sub
    If Len(lineText) = 0 Then Exit Sub
    mDiagnosticLog = mDiagnosticLog & lineText & vbCrLf
End Sub

' Очищает результат одной попытки перед новым НДС; численные настройки сохраняются.
Private Sub ClearResult()
    Set mSolver = Nothing
    mExtensionUsed = False
    mWithinPhysicalRange = False
    mDiagnosticLog = vbNullString
    mInitialGuessApplied = False
    mSolverCallCount = 0
End Sub

' Форматирует числа диагностики с точкой как десятичным разделителем, чтобы
' отчет одинаково читался при русской и английской локали Excel/Windows.
Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

' Возвращает объект последней попытки решения НДС, в том числе неуспешной.
Public Property Get ResultSolver() As CSectionSolver
    Set ResultSolver = mSolver
End Property

' Показывает, найдено ли равновесие последней попыткой; до первого запуска возвращает False.
Public Property Get Converged() As Boolean
    If Not mSolver Is Nothing Then Converged = mSolver.Converged
End Property

' Возвращает причину остановки последней попытки решения НДС.
Public Property Get StopReason() As String
    If Not mSolver Is Nothing Then StopReason = mSolver.StopReason
End Property

' Возвращает число итераций последней попытки решения НДС.
Public Property Get Iterations() As Long
    If Not mSolver Is Nothing Then Iterations = mSolver.Iterations
End Property

' Количество фактических вызовов CSectionSolver.Solve.
Public Property Get SolverCallCount() As Long
    SolverCallCount = mSolverCallCount
End Property

' Показывает, использовано ли вспомогательное продолжение диаграмм в найденном НДС.
Public Property Get ExtensionUsed() As Boolean
    ExtensionUsed = mExtensionUsed
End Property

' True, если все элементы остались в физических eps-пределах.
Public Property Get WithinPhysicalRange() As Boolean
    WithinPhysicalRange = mWithinPhysicalRange
End Property

' True, если стартовая плоскость была задана не нулевой.
Public Property Get InitialGuessApplied() As Boolean
    InitialGuessApplied = mInitialGuessApplied
End Property

' Возвращает краткий журнал повторных попыток решения НДС.
Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

' Newton или Secant для прямого равновесия.
Public Property Let SolverMethod(ByVal value As String)
    mSolverMethod = value
End Property

' Число ступеней приложения нагрузки.
Public Property Let LoadSteps(ByVal value As Long)
    mLoadSteps = value
End Property

' Максимум итераций CSectionSolver на одной ступени.
Public Property Let MaxIterations(ByVal value As Long)
    mMaxIterations = value
End Property

' Допуск равновесия по N, Н.
Public Property Let ToleranceN(ByVal value As Double)
    mToleranceN = value
End Property

' Допуск равновесия по Mx, Н*мм.
Public Property Let ToleranceMx(ByVal value As Double)
    mToleranceMx = value
End Property

' Допуск равновесия по My, Н*мм.
Public Property Let ToleranceMy(ByVal value As Double)
    mToleranceMy = value
End Property

' Лимит перезапуска матрицы в secant-ветке.
Public Property Let SecantMaxRestarts(ByVal value As Long)
    mSecantMaxRestarts = value
End Property

' Нужно ли копить подробный журнал итераций.
Public Property Let DiagnosticsEnabled(ByVal value As Boolean)
    mDiagnosticsEnabled = value
End Property

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================

' ДЛЯ ТЕСТОВ: вызывает настоящую процедуру повторного решения с явно заданными числами.
' Проверяет учёт точно совпадающих попыток без создания конечного НДС
' или записи в общее хранилище состояний.
Friend Function TestRunRetryAttempt(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, Optional ByVal epsilon0 As Double = 0#, _
        Optional ByVal loadSteps As Long = 1) As Boolean
    Dim service As CSectionSolver
    TestRunRetryAttempt = TryRunRetry(section, concreteMaterial, steelMaterial, _
        targetN, 0#, 0#, epsilon0, 0#, 0#, targetN, 0#, 0#, loadSteps, _
        "проверка эквивалентности попыток", service, "testRetry")
End Function

' ДЛЯ ТЕСТОВ: возвращает число сохранённых кратких итогов попыток для проверки их очистки после поиска.
Friend Property Get TestRetryAttemptCount() As Long
    If Not mRetryAttemptFailures Is Nothing Then TestRetryAttemptCount = mRetryAttemptFailures.Count
End Property
