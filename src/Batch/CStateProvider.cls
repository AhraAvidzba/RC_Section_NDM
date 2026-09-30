VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStateProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Единая точка получения named-state
' ==========================================================================
' CStateProvider стоит между инженерными расчетами и CStateSolutionRunner. Он
' сначала проверяет CStateRepository, затем при необходимости запускает общий
' state-solve pipeline и сохраняет полученный конечный state. Search-probes
' сюда не складываются: provider работает только с именованными состояниями,
' которые имеют самостоятельный смысл в Results.

Private mSection As CSectionModel ' Расчетная модель сечения для solve.
Private mMaterialProvider As CMaterialModelProvider ' Поставщик физических и extended-диаграмм.
Private mRepository As CStateRepository ' Хранилище named-state текущего LC.

Private mSolverMethod As String ' Newton или Secant.
Private mLoadSteps As Long ' Число ступеней приложения нагрузки.
Private mMaxIterations As Long ' Максимум итераций solver-а.
Private mToleranceN As Double ' Допуск по N.
Private mToleranceMx As Double ' Допуск по Mx.
Private mToleranceMy As Double ' Допуск по My.
Private mLineSearchEnabled As Boolean ' Разрешение line search.
Private mDampingInitial As Double ' Начальный alpha line search.
Private mMinLineSearchAlpha As Double ' Нижняя граница alpha.
Private mMaxDeltaEpsilon0 As Double ' Ограничение шага epsilon0.
Private mMaxDeltaKappa As Double ' Ограничение шага кривизн.
Private mSecantMaxRestarts As Long ' Перезапуски secant-матрицы.
Private mSecantMinStepNorm As Double ' Минимальная норма secant-шага.
Private mDiagnosticsEnabled As Boolean ' Нужно ли копить диагностический журнал.

Private mLastRunner As CStateSolutionRunner ' Runner последнего реального solve; Nothing при cache-hit.
Private mLastStateWasReused As Boolean ' True, если state взят из repository.

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
End Sub

' Подключает расчетный контекст. Repository может быть Nothing; тогда provider
' работает как простой solve-wrapper без переиспользования.
Public Sub Initialize(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        Optional ByVal repository As CStateRepository = Nothing)
    If section Is Nothing Then Err.Raise vbObjectError + 3990, "CStateProvider", "Модель сечения не передана."
    If materialProvider Is Nothing Then Err.Raise vbObjectError + 3991, "CStateProvider", "Поставщик материалов не передан."
    Set mSection = section
    Set mMaterialProvider = materialProvider
    Set mRepository = repository
End Sub

' Возвращает эквивалентный state из repository или решает новый. initialSolver
' является только warm-start и не входит в ключ reuse.
Public Function GetOrSolve(ByVal request As CStateRequest, _
        Optional ByVal initialSolver As CSectionSolver = Nothing) As CSectionStateResult
    If request Is Nothing Then Err.Raise vbObjectError + 3992, "CStateProvider", "Запрос named-state не передан."
    Set mLastRunner = Nothing
    mLastStateWasReused = False

    If Not mRepository Is Nothing Then
        Dim cached As CSectionStateResult
        Set cached = mRepository.FindEquivalent(request)
        If Not cached Is Nothing Then
            mLastStateWasReused = True
            Set GetOrSolve = cached
            Exit Function
        End If
    End If

    Dim runner As CStateSolutionRunner
    Set runner = New CStateSolutionRunner
    ConfigureRunner runner

    Dim concreteMaterial As CMaterialDiagram
    Dim steelMaterial As CMaterialDiagram
    Set concreteMaterial = ConcreteMaterialForRequest(request)
    Set steelMaterial = SteelMaterialForRequest(request)

    If initialSolver Is Nothing Then
        runner.Solve mSection, concreteMaterial, steelMaterial, _
            request.TargetN, request.TargetMx, request.TargetMy, request.AllowRetryStarts
    Else
        runner.SolveWithInitialSolver mSection, concreteMaterial, steelMaterial, _
            request.TargetN, request.TargetMx, request.TargetMy, initialSolver, request.AllowRetryStarts
    End If
    Set mLastRunner = runner

    Dim service As CSectionSolver
    Set service = runner.ResultSolver
    If service Is Nothing Then Exit Function

    Dim stateResult As CSectionStateResult
    Set stateResult = New CSectionStateResult
    stateResult.InitializeFromSolver request.StateType, request.MaterialRole, request.MaterialSpec, _
        service, runner.ExtensionUsed, runner.WithinPhysicalRange, _
        runner.SolverCallCount, runner.DiagnosticLog

    If Not mRepository Is Nothing Then mRepository.StoreForRequest request, stateResult
    Set GetOrSolve = stateResult
End Function

' Передает сохраненные настройки в новый runner.
Private Sub ConfigureRunner(ByVal runner As CStateSolutionRunner)
    runner.LoadSteps = mLoadSteps
    runner.MaxIterations = mMaxIterations
    runner.ToleranceN = mToleranceN
    runner.ToleranceMx = mToleranceMx
    runner.ToleranceMy = mToleranceMy
    runner.LineSearchEnabled = mLineSearchEnabled
    runner.DampingInitial = mDampingInitial
    runner.MinLineSearchAlpha = mMinLineSearchAlpha
    runner.SolverMethod = mSolverMethod
    runner.SecantMaxRestarts = mSecantMaxRestarts
    runner.SecantMinStepNorm = mSecantMinStepNorm
    runner.MaxDeltaEpsilon0 = mMaxDeltaEpsilon0
    runner.MaxDeltaKappa = mMaxDeltaKappa
    runner.DiagnosticsEnabled = mDiagnosticsEnabled
End Sub

' Выбирает бетонную диаграмму по физическому режиму request.
Private Function ConcreteMaterialForRequest(ByVal request As CStateRequest) As CMaterialDiagram
    If request.UseDirectStateExtension Then
        Set ConcreteMaterialForRequest = mMaterialProvider.ConcreteStateMaterialFromSpec(request.MaterialSpec)
    Else
        Set ConcreteMaterialForRequest = mMaterialProvider.ConcreteMaterialFromSpec(request.MaterialSpec)
    End If
End Function

' Выбирает диаграмму арматуры по физическому режиму request.
Private Function SteelMaterialForRequest(ByVal request As CStateRequest) As CMaterialDiagram
    If request.UseDirectStateExtension Then
        Set SteelMaterialForRequest = mMaterialProvider.SteelStateMaterialFromSpec(request.MaterialSpec)
    Else
        Set SteelMaterialForRequest = mMaterialProvider.SteelMaterialFromSpec(request.MaterialSpec)
    End If
End Function

Public Property Get LastRunner() As CStateSolutionRunner
    Set LastRunner = mLastRunner
End Property

Public Property Get LastStateWasReused() As Boolean
    LastStateWasReused = mLastStateWasReused
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
