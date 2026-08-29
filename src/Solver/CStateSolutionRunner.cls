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
' CStateSolutionRunner управляет только пользовательским состоянием от заданного
' сочетания N + Mx + My. Он создает CSectionSolver, задает ему настройки,
' подбирает стартовую плоскость через CStateGuessBuilder и, если первая попытка
' не сошлась, пробует согласованные повторные старты в technical extension.
'
' В классе нет выбора Group1/Group2, диаграмм Strength/CrackedNDS и нет записи
' Results. Материалы уже должны быть подготовлены CMaterialModelProvider, а
' CBatchSectionCalculator после этого только читает готовый результат: сошелся
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

' Решает прямое равновесие от заданного LC.
' allowExtensionWarmStart означает не сам факт включения extension в материалах,
' а разрешение на вторую серию стартов, если первая попытка CSectionSolver не
' нашла равновесие. Сами диаграммы при этом уже переданы снаружи.
Public Sub Solve(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        Optional ByVal allowExtensionWarmStart As Boolean = False)
    If section Is Nothing Then Err.Raise vbObjectError + 3480, "CStateSolutionRunner", "Модель сечения не передана."
    If concreteMaterial Is Nothing Then Err.Raise vbObjectError + 3481, "CStateSolutionRunner", "Диаграмма бетона не передана."
    If steelMaterial Is Nothing Then Err.Raise vbObjectError + 3482, "CStateSolutionRunner", "Диаграмма арматуры не передана."

    ClearResult

    Dim service As CSectionSolver
    Set service = New CSectionSolver
    ConfigureSolver service
    ApplyInitialGuess section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy, service

    service.Solve section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy
    Set mSolver = service

    If Not service.Converged And allowExtensionWarmStart Then
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
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal solver As CSectionSolver)
    If Abs(targetN) > NEAR_ZERO Then Exit Sub
    If Sqr(targetMx * targetMx + targetMy * targetMy) <= NEAR_ZERO Then Exit Sub

    Dim guessBuilder As CStateGuessBuilder
    Dim eps0 As Double
    Dim kx As Double
    Dim ky As Double
    Set guessBuilder = New CStateGuessBuilder
    If guessBuilder.BuildForDirectState(section, concreteMaterial, steelMaterial, _
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
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByRef service As CSectionSolver) As Boolean
    Dim guessBuilder As CStateGuessBuilder
    Set guessBuilder = New CStateGuessBuilder

    Dim eps0 As Double
    Dim kx As Double
    Dim ky As Double
    Dim attemptIndex As Long

    For attemptIndex = 1 To guessBuilder.ExtensionAttemptCount
        If guessBuilder.BuildAttemptForExtensionState(attemptIndex, section, service, _
                concreteMaterial, steelMaterial, targetN, targetMx, targetMy, eps0, kx, ky) Then
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
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
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
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByVal initialN As Double, ByVal initialMx As Double, ByVal initialMy As Double, _
        ByVal loadSteps As Long, ByVal attemptName As String, ByRef service As CSectionSolver, _
        Optional ByVal disableLineSearch As Boolean = False) As Boolean
    Dim retry As CSectionSolver
    Set retry = New CSectionSolver
    ConfigureSolver retry
    If disableLineSearch Then retry.LineSearchEnabled = False
    retry.LoadSteps = loadSteps
    retry.SetInitialLoadState eps0, kx, ky, initialN, initialMx, initialMy
    retry.Solve section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy

    If retry.Converged Then
        Set service = retry
        AppendDiagnostic "extensionRetry=" & attemptName & "; result=converged"
        TryRunRetry = True
    Else
        AppendDiagnostic "extensionRetry=" & attemptName & _
            "; result=failed; stop=" & retry.StopReason & _
            "; residualN=" & FormatNumberInvariant(retry.ResidualN) & _
            "; residualMx=" & FormatNumberInvariant(retry.ResidualMx) & _
            "; residualMy=" & FormatNumberInvariant(retry.ResidualMy)
    End If
End Function

' Проверяет, попала ли деформация хотя бы одного элемента в добавленную
' provider-ом technical extension. Этот признак сохраняется в Results отдельно
' от пользовательского статуса и нужен схеме, AutoCAD и расчету трещин.
Private Function StateUsesExtension(ByVal section As CSectionModel, ByVal solver As CSectionSolver, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram) As Boolean
    Dim i As Long
    Dim strain As Double

    For i = 1 To section.ConcreteCount
        strain = ElementStrain(solver, section.ConcreteX(i), section.ConcreteY(i))
        If concreteMaterial.IsInExtensionRange(strain) Then
            StateUsesExtension = True
            Exit Function
        End If
    Next i

    For i = 1 To section.RebarCount
        strain = ElementStrain(solver, section.RebarX(i), section.RebarY(i))
        If steelMaterial.IsInExtensionRange(strain) Then
            StateUsesExtension = True
            Exit Function
        End If
    Next i
End Function

' Проверяет физические eps-пределы всех бетонных волокон и стержней.
' Если растянутая ветвь бетона отключена и у диаграммы нет растянутого
' предела, положительные деформации бетона считаются допустимыми для проверки
' range, но дают нулевое напряжение через саму диаграмму.
Private Function StateWithinPhysicalRange(ByVal section As CSectionModel, ByVal solver As CSectionSolver, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram) As Boolean
    Dim i As Long
    Dim strain As Double

    For i = 1 To section.ConcreteCount
        strain = ElementStrain(solver, section.ConcreteX(i), section.ConcreteY(i))
        If Not concreteMaterial.IsInPhysicalRange(strain) Then Exit Function
    Next i

    For i = 1 To section.RebarCount
        strain = ElementStrain(solver, section.RebarX(i), section.RebarY(i))
        If Not steelMaterial.IsInPhysicalRange(strain) Then Exit Function
    Next i

    StateWithinPhysicalRange = True
End Function

Private Function ElementStrain(ByVal solver As CSectionSolver, ByVal x As Double, ByVal y As Double) As Double
    ElementStrain = solver.Epsilon0 + solver.KappaX * y + solver.KappaY * x
End Function

Private Sub AppendDiagnostic(ByVal lineText As String)
    If Len(lineText) = 0 Then Exit Sub
    mDiagnosticLog = mDiagnosticLog & lineText & vbCrLf
End Sub

Private Sub ClearResult()
    Set mSolver = Nothing
    mExtensionUsed = False
    mWithinPhysicalRange = False
    mDiagnosticLog = vbNullString
    mInitialGuessApplied = False
End Sub

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

