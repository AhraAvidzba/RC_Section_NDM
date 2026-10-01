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
' Запуск прямого состояния НДС
' ==========================================================================
' CStateSolutionRunner управляет обычным равновесием для заданных N + Mx + My,
' включая named-state и промежуточные пробы общего поиска. Он создает
' CSectionSolver, задает ему настройки,
' подбирает стартовую плоскость через CStateGuessBuilder и, если первая попытка
' не сошлась, пробует согласованные повторные старты в technical extension.
'
' В классе нет выбора расчетного профиля, диаграмм Strength/CrackedNDS и нет записи
' Results. Материалы уже должны быть подготовлены CMaterialModelProvider, а
' инженерный orchestration-слой после этого только читает готовый результат: сошелся
' ли solver, вышло ли состояние в extension и находится ли оно в физических
' пределах диаграммы.

Private Const NEAR_ZERO As Double = 0.000000001 ' Общий допуск для распознавания почти нулевого усилия/момента.

Private mSolverMethod As String ' Newton или Secant для прямого равновесия.
Private mLoadSteps As Long ' Число ступеней приложения нагрузки.
Private mMaxIterations As Long ' Максимум итераций CSectionSolver на одной ступени.
Private mToleranceN As Double ' Допуск равновесия по N, Н.
Private mToleranceMx As Double ' Допуск равновесия по Mx, Н*мм.
Private mToleranceMy As Double ' Допуск равновесия по My, Н*мм.
Private mLineSearchEnabled As Boolean ' Разрешает line search внутри Newton/Secant.
Private mDampingInitial As Double ' Начальный коэффициент дробления шага.
Private mMinLineSearchAlpha As Double ' Минимальный alpha перед признанием line search неудачным.
Private mMaxDeltaEpsilon0 As Double ' Ограничение одного шага по epsilon0.
Private mMaxDeltaKappa As Double ' Ограничение одного шага по кривизнам.
Private mSecantMaxRestarts As Long ' Лимит перезапуска матрицы в secant-ветке.
Private mSecantMinStepNorm As Double ' Минимальная норма шага secant до остановки.
Private mDiagnosticsEnabled As Boolean ' Нужно ли копить подробный журнал итераций.

Private mSolver As CSectionSolver ' Последний solver прямого НДС, успешный или диагностический.
Private mExtensionUsed As Boolean ' True, если найденное состояние попало в technical extension.
Private mWithinPhysicalRange As Boolean ' True, если все элементы остались в физических eps-пределах.
Private mDiagnosticLog As String ' Короткие строки о retry-попытках прямого НДС.
Private mInitialGuessApplied As Boolean ' True, если стартовая плоскость была задана не нулевой.
Private mSolverCallCount As Long ' Сколько раз runner действительно запускал CSectionSolver.Solve.

' Задает безопасные настройки по умолчанию для unit-тестов и вызовов без Config.
Private Sub Class_Initialize()
    mSolverMethod = "Newton"
    mLoadSteps = 8
    mMaxIterations = 80
    mToleranceN = 5#
    mToleranceMx = 5000#
    mToleranceMy = 5000#
    mLineSearchEnabled = True
    mDampingInitial = 1#
    mMinLineSearchAlpha = 0.03125
    mMaxDeltaEpsilon0 = 0.0005
    mMaxDeltaKappa = 0.00001
    mSecantMaxRestarts = 2
    mSecantMinStepNorm = 0.000000000001
    mDiagnosticsEnabled = True
    mWithinPhysicalRange = True
End Sub

' Восстанавливает вычисляемый снимок найденного named-state без поиска
' равновесия. Материалы и сечение должны принадлежать тому же контексту;
' сохраненные целевые усилия проверяются повторно, число solve не увеличивается.
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

' Решает прямое равновесие от заданного LC, начиная не с нулевой плоскости,
' а с уже найденного solver-а. Это общий путь для служебных состояний вроде
' PostCrackState: они получают ту же централизованную настройку solver-а,
' диагностику и fallback-логику, что и обычные Strength/CrackedState, но могут
' стартовать от близкого физического состояния на том же уровне нагрузки.
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

' Решает прямое равновесие от заданного LC с явной стартовой нагрузкой.
' Этот вход нужен search-probe-ам Capacity: предыдущая lambda-точка задает не
' только eps0/kappa, но и уровень нагрузки, от которого solver должен перейти к
' новой target-точке. Warm-start остается solve-option и не попадает в
' repository/reuse-ключи.
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

' Решает direct-state из active-set плоскости, сразу привязанной к текущей
' целевой нагрузке. Этот вход нужен capacity probe-ам чистого изгиба: такая
' lambda-точка не должна идти ступенями от нулевой нагрузки через излом
' диаграммы около нуля, но все остальные настройки solver-а и диагностика
' остаются в общем CStateSolutionRunner.
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

' Ошибки модели, настроек и программного контракта не исправляются сменой
' начальной плоскости; повтор допустим только после численной неудачи.
Private Function CanRetrySolver(ByVal solver As CSectionSolver) As Boolean
    If solver Is Nothing Then Exit Function
    If solver.Converged Then Exit Function
    Select Case solver.FailureCode
        Case sfcNumericalFailure, sfcSingularTangent
            CanRetrySolver = True
    End Select
End Function

' Повторяет прямой solve от active-set плоскости для общего случая N+M.
' Первый штатный старт сохраняется прежним; этот retry включается только после
' несходимости. Старт на полной нагрузке отличается от применения такой же
' плоскости при ступенчатом нагружении и не должен блокироваться его флагом. Так
' внецентренное растяжение/сжатие получает тот же устойчивый старт, который уже
' помогал чистому изгибу, без изменения результата успешно сходящихся LC.
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
            "directStateActiveSet / direct target", service, False, "directStateRetry") Then
        mInitialGuessApplied = True
        TrySolveWithDirectStateGuess = True
        Exit Function
    End If

    If TryRunRetry(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
            eps0, kx, ky, targetN, targetMx, targetMy, 1, _
            "directStateActiveSet / direct target without line search", service, True, "directStateRetry") Then
        mInitialGuessApplied = True
        TrySolveWithDirectStateGuess = True
    End If
End Function

' Выбирает старт для прямого solve-а. Если вызывающий слой передал близкое
' состояние, используем его как warm start. Иначе включается обычный общий
' подбор стартовой плоскости через CStateGuessBuilder.
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

' Передает настройки в новый экземпляр CSectionSolver.
' Runner не меняет алгоритм solver-а: он только гарантирует, что все попытки
' прямого НДС используют один и тот же набор пользовательских допусков.
Private Sub ConfigureSolver(ByVal solver As CSectionSolver)
    solver.LoadSteps = mLoadSteps
    solver.SolverMethod = mSolverMethod
    solver.MaxIterations = mMaxIterations
    solver.ToleranceN = mToleranceN
    solver.ToleranceMx = mToleranceMx
    solver.ToleranceMy = mToleranceMy
    solver.LineSearchEnabled = mLineSearchEnabled
    solver.DampingInitial = mDampingInitial
    solver.MinLineSearchAlpha = mMinLineSearchAlpha
    solver.MaxDeltaEpsilon0 = mMaxDeltaEpsilon0
    solver.MaxDeltaKappa = mMaxDeltaKappa
    solver.SecantMaxRestarts = mSecantMaxRestarts
    solver.SecantMinStepNorm = mSecantMinStepNorm
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
    If Sqr(targetMx * targetMx + targetMy * targetMy) <= NEAR_ZERO Then Exit Sub

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

' Перебирает стартовые плоскости, которые CStateGuessBuilder строит для
' technical extension. Здесь находится управление retry: builder только дает
' eps0/kx/ky, а runner запускает новый CSectionSolver и проверяет сходимость.
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

' Проверяет одну стартовую плоскость retry-сценария.
' Сначала пробуем сразу попасть в целевую нагрузку, затем тот же старт без
' line search, и только после этого идем ступенями от усилий, которые уже
' создает стартовая плоскость.
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

    If TryRunRetry(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
            eps0, kx, ky, targetN, targetMx, targetMy, 1, _
            attemptName & " / direct target without line search", service, True) Then
        TrySolveFromExtensionGuess = True
        Exit Function
    End If

    TrySolveFromExtensionGuess = TryRunRetry(section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, _
        eps0, kx, ky, startState.Nint, startState.Mxint, startState.Myint, retryLoadSteps, _
        attemptName & " / stepped", service)
End Function

' Выполняет один повторный solve из подготовленной плоскости.
' initialN/Mx/My задают уровень нагрузки, от которого CSectionSolver будет
' прикладывать ступени до полного target. Это позволяет идти от близкого
' extension-состояния, а не от нулевой нагрузки с чужой деформированной формой.
Private Function TryRunRetry(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByVal initialN As Double, ByVal initialMx As Double, ByVal initialMy As Double, _
        ByVal loadSteps As Long, ByVal attemptName As String, ByRef service As CSectionSolver, _
        Optional ByVal disableLineSearch As Boolean = False, _
        Optional ByVal retryKind As String = "extensionRetry") As Boolean
    Dim retry As CSectionSolver
    Set retry = New CSectionSolver
    ConfigureSolver retry
    If disableLineSearch Then retry.LineSearchEnabled = False
    retry.LoadSteps = loadSteps
    retry.SetInitialLoadState eps0, kx, ky, initialN, initialMx, initialMy
    mSolverCallCount = mSolverCallCount + 1
    retry.Solve section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy

    If retry.Converged Then
        Set service = retry
        AppendDiagnostic retryKind & "=" & attemptName & "; result=converged"
        TryRunRetry = True
    Else
        AppendDiagnostic retryKind & "=" & attemptName & _
            "; result=failed; stop=" & retry.StopReason & _
            "; residualN=" & FormatNumberInvariant(retry.ResidualN) & _
            "; residualMx=" & FormatNumberInvariant(retry.ResidualMx) & _
            "; residualMy=" & FormatNumberInvariant(retry.ResidualMy)
    End If
End Function

' Проверяет, попала ли деформация хотя бы одного элемента в добавленную
' provider-ом technical extension. Этот признак сохраняется в Results отдельно
' от пользовательского статуса и нужен схеме, AutoCAD и расчету трещин.
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

' Возвращает штатные диаграммы материалов, если они действительно переданы.
' Тестовые линейные материалы решаются тем же CSectionSolver, но без
' специализированных active-set и extension-подсказок, потому что у них нет
' физических границ диаграммы.
Private Function TryGetMaterialDiagrams(ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByRef concreteDiagram As CMaterialDiagram, ByRef steelDiagram As CMaterialDiagram) As Boolean
    If concreteMaterial Is Nothing Or steelMaterial Is Nothing Then Exit Function
    If TypeName(concreteMaterial) <> "CMaterialDiagram" Then Exit Function
    If TypeName(steelMaterial) <> "CMaterialDiagram" Then Exit Function

    Set concreteDiagram = concreteMaterial
    Set steelDiagram = steelMaterial
    TryGetMaterialDiagrams = True
End Function

' Вычисляет деформацию произвольной точки по найденной solver-плоскости
' epsilon0/kappaX/kappaY. Используется для проверки диаграмм после решения.
Private Function ElementStrain(ByVal solver As CSectionSolver, ByVal x As Double, ByVal y As Double) As Double
    ElementStrain = solver.Epsilon0 + solver.KappaX * y + solver.KappaY * x
End Function

' Добавляет строку в диагностический журнал только когда диагностика включена.
' Так рабочий расчет не тратит время на лишний текстовый вывод.
Private Sub AppendDiagnostic(ByVal lineText As String)
    If Not mDiagnosticsEnabled Then Exit Sub
    If Len(lineText) = 0 Then Exit Sub
    mDiagnosticLog = mDiagnosticLog & lineText & vbCrLf
End Sub

' Сбрасывает результат одного запуска runner-а перед новым решением НДС.
' Настройки solver-а при этом не меняются.
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

Public Property Get ResultSolver() As CSectionSolver
    Set ResultSolver = mSolver
End Property

Public Property Get Converged() As Boolean
    If Not mSolver Is Nothing Then Converged = mSolver.Converged
End Property

Public Property Get StopReason() As String
    If Not mSolver Is Nothing Then StopReason = mSolver.StopReason
End Property

Public Property Get Iterations() As Long
    If Not mSolver Is Nothing Then Iterations = mSolver.Iterations
End Property

Public Property Get SolverCallCount() As Long
    SolverCallCount = mSolverCallCount
End Property

Public Property Get ExtensionUsed() As Boolean
    ExtensionUsed = mExtensionUsed
End Property

Public Property Get WithinPhysicalRange() As Boolean
    WithinPhysicalRange = mWithinPhysicalRange
End Property

Public Property Get InitialGuessApplied() As Boolean
    InitialGuessApplied = mInitialGuessApplied
End Property

Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

Public Property Let SolverMethod(ByVal value As String)
    mSolverMethod = value
End Property

Public Property Let LoadSteps(ByVal value As Long)
    mLoadSteps = value
End Property

Public Property Let MaxIterations(ByVal value As Long)
    mMaxIterations = value
End Property

Public Property Let ToleranceN(ByVal value As Double)
    mToleranceN = value
End Property

Public Property Let ToleranceMx(ByVal value As Double)
    mToleranceMx = value
End Property

Public Property Let ToleranceMy(ByVal value As Double)
    mToleranceMy = value
End Property

Public Property Let LineSearchEnabled(ByVal value As Boolean)
    mLineSearchEnabled = value
End Property

Public Property Let DampingInitial(ByVal value As Double)
    mDampingInitial = value
End Property

Public Property Let MinLineSearchAlpha(ByVal value As Double)
    mMinLineSearchAlpha = value
End Property

Public Property Let MaxDeltaEpsilon0(ByVal value As Double)
    mMaxDeltaEpsilon0 = value
End Property

Public Property Let MaxDeltaKappa(ByVal value As Double)
    mMaxDeltaKappa = value
End Property

Public Property Let SecantMaxRestarts(ByVal value As Long)
    mSecantMaxRestarts = value
End Property

Public Property Let SecantMinStepNorm(ByVal value As Double)
    mSecantMinStepNorm = value
End Property

Public Property Let DiagnosticsEnabled(ByVal value As Boolean)
    mDiagnosticsEnabled = value
End Property

