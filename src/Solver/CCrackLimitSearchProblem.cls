VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCrackLimitSearchProblem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILimitSearchProblem

' ==========================================================================
' Адаптер CrackFormation для общего LimitSearch
' ==========================================================================
' Класс прячет CCrackFormationCalculator за общим ILimitSearchProblem. Доменный
' калькулятор по-прежнему хранит критерий образования трещины, Pre/Post-state и
' crack-диагностику, но общий search обращается к нему только через контракт.

Private Const STATE_ACCEPTABLE As String = "Acceptable" ' Трещина еще не образована.
Private Const STATE_CRACK_LIMIT As String = "CrackFormationLimit" ' Достигнут критерий eps_bt,ult.
Private Const STATE_NUMERICAL_FAILURE As String = "NumericalFailure" ' НДС для пробной lambda не найдено.

Private mCrack As CCrackFormationCalculator ' Доменный расчетчик трещинообразования и владелец crack-диагностики.

' Подключает существующий расчетчик трещин без изменения его настроек.
Public Sub Initialize(ByVal crackCalculator As CCrackFormationCalculator)
    If crackCalculator Is Nothing Then Err.Raise vbObjectError + 4250, "CCrackLimitSearchProblem", _
        "Не передан расчетчик трещинообразования."
    Set mCrack = crackCalculator
End Sub

Private Property Get ILimitSearchProblem_ResultKind() As EResultKind
    ILimitSearchProblem_ResultKind = rkCrackFormation
End Property

Private Property Get ILimitSearchProblem_InitialLambdaStep() As Double
    ILimitSearchProblem_InitialLambdaStep = 1#
End Property

Private Property Get ILimitSearchProblem_MaxLambda() As Double
    ILimitSearchProblem_MaxLambda = mCrack.LimitSearchCrackBracketLimit
End Property

Private Property Get ILimitSearchProblem_LambdaTolerance() As Double
    ILimitSearchProblem_LambdaTolerance = mCrack.LimitSearchCrackLambdaTolerance
End Property

Private Property Get ILimitSearchProblem_MaxIterations() As Long
    ILimitSearchProblem_MaxIterations = mCrack.LimitSearchCrackMaxIterations
End Property

Private Property Get ILimitSearchProblem_SearchMethod() As String
    ILimitSearchProblem_SearchMethod = "Bisection"
End Property

Private Property Get ILimitSearchProblem_UsesRetainedForceCapacity() As Boolean
    ILimitSearchProblem_UsesRetainedForceCapacity = False
End Property

Private Property Get ILimitSearchProblem_Converged() As Boolean
    ILimitSearchProblem_Converged = mCrack.Converged
End Property

Private Property Get ILimitSearchProblem_LimitState() As String
    ILimitSearchProblem_LimitState = mCrack.StopReason
End Property

Private Property Get ILimitSearchProblem_DiagnosticLog() As String
    If Not mCrack Is Nothing Then ILimitSearchProblem_DiagnosticLog = mCrack.DiagnosticLog
End Property

' Для CrackFormation специальной подготовки перед LoadMultiplier нет:
' материалы и путь уже переданы вызывающим слоем.
Private Function ILimitSearchProblem_Prepare(ByVal request As CLimitSearchRequest) As Boolean
    ILimitSearchProblem_Prepare = True
End Function

' Материалы и инженерный критерий уже подготовлены владельцем Formation.
Private Function ILimitSearchProblem_PrepareUltimate(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector) As Boolean
    ILimitSearchProblem_PrepareUltimate = Not loadPath Is Nothing
End Function

' Сохраняет выбор стартов CrackFormation, но все попытки решаются общим
' Newton-циклом через типизированный ILimitSearchProblem.
Private Function ILimitSearchProblem_RunUltimateAttempts(ByVal request As CLimitSearchRequest) As Boolean
    ILimitSearchProblem_RunUltimateAttempts = mCrack.LimitSearchRunCrackFormationUltimate( _
        request.Section, request.ConcreteMaterial, request.SteelMaterial, request.TargetN, _
        request.TargetMx, request.TargetMy, request.MomentMxAtReference, _
        request.MomentMyAtReference, request.ReferenceX, request.ReferenceY)
End Function

' Formation сохраняет точку непосредственно при принятии; отдельная
' финализация инженерных чисел после общего цикла ему не требуется.
Private Sub ILimitSearchProblem_CompleteUltimate(ByVal loadPath As CLoadPathVector)
End Sub

' Передает критерии и численные настройки Formation общему Newton-циклу.
' Ни шаг Newton, ни Якобиан, ни line search здесь не реализуются.
Private Sub ILimitSearchProblem_LimitSearchInitializeUltimateGuess(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double)
    mCrack.LimitSearchInitializeUltimateGuess section, concreteMaterial, steelMaterial, _
        loadPath, startSolver, eps0, kx, ky
End Sub

Private Function ILimitSearchProblem_LimitSearchEvaluateUltimateResidual(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal eps0 As Double, _
        ByVal kx As Double, ByVal ky As Double, ByRef solver As CSectionSolver, _
        ByRef rPath1 As Double, ByRef rPath2 As Double, ByRef rLimit As Double, _
        ByRef lambdaValue As Double, ByRef state As String) As Boolean
    ILimitSearchProblem_LimitSearchEvaluateUltimateResidual = mCrack.LimitSearchEvaluateUltimateResidual( _
        section, concreteMaterial, steelMaterial, loadPath, eps0, kx, ky, solver, _
        rPath1, rPath2, rLimit, lambdaValue, state)
End Function

Private Function ILimitSearchProblem_LimitSearchUltimateResidualReached(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Boolean
    ILimitSearchProblem_LimitSearchUltimateResidualReached = _
        mCrack.LimitSearchUltimateResidualReached(rPath1, rPath2, rLimit)
End Function

Private Function ILimitSearchProblem_LimitSearchUltimateResidualNorm(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Double
    ILimitSearchProblem_LimitSearchUltimateResidualNorm = _
        mCrack.LimitSearchUltimateResidualNorm(rPath1, rPath2, rLimit)
End Function

Private Function ILimitSearchProblem_LimitSearchFinalizeUltimateResult(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal solver As CSectionSolver, ByVal state As String, ByVal iterations As Long) As Boolean
    ILimitSearchProblem_LimitSearchFinalizeUltimateResult = mCrack.LimitSearchFinalizeUltimateResult( _
        section, concreteMaterial, steelMaterial, loadPath, lambdaValue, solver, state, iterations)
End Function

Private Sub ILimitSearchProblem_LimitSearchSetUltimateNumericalFailure(ByVal reason As String)
    mCrack.LimitSearchSetUltimateNumericalFailure reason
End Sub

' Однотипные настройки общего Newton из численной конфигурации Formation.
Private Property Get ILimitSearchProblem_LimitSearchUltimateMaxIterations() As Long
    ILimitSearchProblem_LimitSearchUltimateMaxIterations = mCrack.LimitSearchUltimateMaxIterations
End Property

Private Property Get ILimitSearchProblem_LimitSearchUltimateDampingInitial() As Double
    ILimitSearchProblem_LimitSearchUltimateDampingInitial = mCrack.LimitSearchUltimateDampingInitial
End Property

Private Property Get ILimitSearchProblem_LimitSearchUltimateMinLineSearchAlpha() As Double
    ILimitSearchProblem_LimitSearchUltimateMinLineSearchAlpha = mCrack.LimitSearchUltimateMinLineSearchAlpha
End Property

Private Property Get ILimitSearchProblem_LimitSearchUltimateLineSearchEnabled() As Boolean
    ILimitSearchProblem_LimitSearchUltimateLineSearchEnabled = mCrack.LimitSearchUltimateLineSearchEnabled
End Property

Private Property Get ILimitSearchProblem_LimitSearchUltimateMaxDeltaEpsilon0() As Double
    ILimitSearchProblem_LimitSearchUltimateMaxDeltaEpsilon0 = mCrack.LimitSearchUltimateMaxDeltaEpsilon0
End Property

Private Property Get ILimitSearchProblem_LimitSearchUltimateMaxDeltaKappa() As Double
    ILimitSearchProblem_LimitSearchUltimateMaxDeltaKappa = mCrack.LimitSearchUltimateMaxDeltaKappa
End Property

' Создает crack-траекторию с учетом выбранного LambdaMxy/LambdaN/LambdaNMxy.
Private Function ILimitSearchProblem_BuildLoadPath(ByVal request As CLimitSearchRequest) As CLoadPathVector
    Set ILimitSearchProblem_BuildLoadPath = mCrack.LimitSearchCrackFormationLoadPath( _
        request.TargetN, request.MomentMxAtReference, request.MomentMyAtReference, _
        request.ReferenceX, request.ReferenceY)
End Function

' Оценивает одну lambda-точку crack-formation и возвращает solver, если
' равновесие найдено. Solver нужен для warm-start и сохранения PreCrackState.
Private Function ILimitSearchProblem_EvaluatePoint(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByVal lambdaValue As Double, ByRef functionValue As Double, _
        ByRef probeState As String, ByRef solver As CSectionSolver) As Boolean
    ILimitSearchProblem_EvaluatePoint = mCrack.LimitSearchEvaluateCrackFormationLoadMultiplier( _
        request.Section, request.ConcreteMaterial, request.SteelMaterial, _
        loadPath, startSolver, lambdaValue, functionValue, probeState, solver)
End Function

Private Function ILimitSearchProblem_IsAcceptableState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsAcceptableState = (probeState = STATE_ACCEPTABLE)
End Function

Private Function ILimitSearchProblem_IsLimitState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsLimitState = (probeState = STATE_CRACK_LIMIT)
End Function

Private Function ILimitSearchProblem_IsNumericalFailureState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsNumericalFailureState = (probeState = STATE_NUMERICAL_FAILURE)
End Function

Private Function ILimitSearchProblem_IsInvalidInputState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsInvalidInputState = False
End Function

Private Function ILimitSearchProblem_ReduceLambdaAfterNumericalFailure(ByVal lowLambda As Double, _
        ByVal highLambda As Double, ByVal probeState As String) As Double
    ILimitSearchProblem_ReduceLambdaAfterNumericalFailure = 0.5 * (lowLambda + highLambda)
End Function

Private Sub ILimitSearchProblem_RecordAcceptable(ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double)
End Sub

Private Sub ILimitSearchProblem_HandleInitialFailure(ByVal probeState As String)
    mCrack.LimitSearchSetCrackFormationStopReason _
        "Не удалось найти равновесие в начальной точке пути трещинообразования."
End Sub

Private Sub ILimitSearchProblem_HandleBracketNotFound(ByVal lowLambda As Double)
    mCrack.LimitSearchSetCrackFormationSearchBoundReached _
        "Поиск трещинообразования достиг технической границы lambda = " & _
        FormatNumberInvariant(mCrack.LimitSearchCrackBracketLimit) & _
        ", но критерий eps_bt,ult не был достигнут. Это не доказывает отсутствие трещины."
End Sub

Private Sub ILimitSearchProblem_HandleBracketNumericalFailure(ByVal lowLambda As Double, _
        ByVal failedLambda As Double, ByVal probeState As String)
    mCrack.LimitSearchSetCrackFormationStopReason _
        "Не удалось найти равновесие при lambda = " & _
        FormatNumberInvariant(failedLambda) & " в поиске трещинообразования."
End Sub

Private Sub ILimitSearchProblem_HandleIterationLimit(ByVal maxIterations As Long, _
        ByVal remainingWidth As Double, ByVal tolerance As Double)
    mCrack.LimitSearchSetCrackFormationStopReason _
        "Поиск Bisection достиг лимита итераций до заданной точности: ширина скобки " & _
        FormatNumberInvariant(remainingWidth) & ", допуск " & FormatNumberInvariant(tolerance) & "."
End Sub

Private Sub ILimitSearchProblem_HandleUnexpectedState(ByVal probeState As String)
    mCrack.LimitSearchSetCrackFormationStopReason _
        "Поиск трещинообразования остановлен: " & probeState & "."
End Sub

Private Sub ILimitSearchProblem_SetInvalidInput(ByVal message As String)
    mCrack.LimitSearchSetCrackFormationStopReason message
End Sub

Private Sub ILimitSearchProblem_SetNumericalFailure(ByVal message As String)
    mCrack.LimitSearchSetCrackFormationStopReason message
End Sub

Private Sub ILimitSearchProblem_AppendDiagnostic(ByVal message As String)
    mCrack.LimitSearchAppendDiagnostic message
End Sub

Private Function ILimitSearchProblem_FinalizeAt(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal probeState As String, ByVal methodName As String, _
        ByVal solver As CSectionSolver, ByVal iterations As Long) As Boolean
    Dim concreteMaterial As CMaterialDiagram
    Dim steelMaterial As CMaterialDiagram
    Set concreteMaterial = request.ConcreteMaterial
    Set steelMaterial = request.SteelMaterial
    ILimitSearchProblem_FinalizeAt = _
        mCrack.LimitSearchFinalizeCrackFormationLoadMultiplier(request.Section, _
            concreteMaterial, steelMaterial, loadPath, lambdaValue, solver, iterations)
End Function

Private Sub ILimitSearchProblem_FinalizeFromRetainedAcceptable(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal methodName As String)
End Sub

Private Sub ILimitSearchProblem_AcceptRetainedCapacity(ByVal request As CLimitSearchRequest)
End Sub

Private Function ILimitSearchProblem_BuildResult(ByVal request As CLimitSearchRequest, _
        ByVal searchSucceeded As Boolean) As CLimitSearchResult
    Dim result As CLimitSearchResult
    Set result = New CLimitSearchResult
    result.InitializeFromCrackFormation mCrack, request.SolutionStrategy, searchSucceeded
    Set ILimitSearchProblem_BuildResult = result
End Function

Private Sub ILimitSearchProblem_BeginFormationAttempt(ByVal pathName As String)
    mCrack.LimitSearchBeginCrackFormationAttempt pathName
End Sub

Private Sub ILimitSearchProblem_SetResolvedFormationPath(ByVal pathName As String)
    mCrack.LimitSearchSetResolvedCrackFormationPath pathName
End Sub

Private Sub ILimitSearchProblem_SetCurrentFormationPath(ByVal pathName As String)
    mCrack.LimitSearchSetCurrentFormationPath pathName
End Sub

Private Property Get ILimitSearchProblem_FormationFallbackOnly() As Boolean
    ILimitSearchProblem_FormationFallbackOnly = mCrack.LimitSearchFormationFallbackOnly
End Property

Private Property Get ILimitSearchProblem_StoppedTerminally() As Boolean
    ILimitSearchProblem_StoppedTerminally = mCrack.LimitSearchStoppedTerminally()
End Property

Private Property Get ILimitSearchProblem_StopReason() As String
    ILimitSearchProblem_StopReason = mCrack.StopReason
End Property

Private Sub ILimitSearchProblem_ClearStopReason()
    mCrack.LimitSearchClearStopReason
End Sub

Private Sub ILimitSearchProblem_SetStopReasonIfEmpty(ByVal reason As String)
    mCrack.LimitSearchSetStopReasonIfEmpty reason
End Sub

Private Function ILimitSearchProblem_FormationPathDisplayName() As String
    ILimitSearchProblem_FormationPathDisplayName = mCrack.LimitSearchCrackFormationPathDisplayName()
End Function

' Форматирует число для диагностического журнала с точкой как разделителем.
Private Function ILimitSearchProblem_HasScaledFormationLoad(ByVal request As CLimitSearchRequest) As Boolean
    ILimitSearchProblem_HasScaledFormationLoad = _
        mCrack.LimitSearchCrackFormationHasScaledLoad(request.TargetN, _
            request.MomentMxAtReference, request.MomentMyAtReference)
End Function

Private Sub ILimitSearchProblem_MarkFormationPathFallback(ByVal reason As String)
    mCrack.LimitSearchMarkFormationPathFallback reason
End Sub

Private Sub ILimitSearchProblem_FinishNoCrack(ByVal reason As String)
    mCrack.LimitSearchFinishNoCrack reason
End Sub

Private Sub ILimitSearchProblem_MergeFallback(ByVal firstLimitState As String, ByVal firstStopReason As String, _
        ByVal firstDiagnostics As String, ByVal firstAttemptMessage As String, _
        ByVal transitionMessage As String, ByVal successMessage As String)
End Sub

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function
