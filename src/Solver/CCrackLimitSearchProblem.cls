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
' Класс прячет CCrackWidthCalculator за общим ILimitSearchProblem. Доменный
' калькулятор по-прежнему хранит критерий образования трещины, Pre/Post-state и
' crack-диагностику, но общий search обращается к нему только через контракт.

Private Const STATE_ACCEPTABLE As String = "Acceptable" ' Трещина еще не образована.
Private Const STATE_CRACK_LIMIT As String = "CrackFormationLimit" ' Достигнут критерий eps_bt,ult.
Private Const STATE_NUMERICAL_FAILURE As String = "NumericalFailure" ' НДС для пробной lambda не найдено.

Private mCrack As CCrackWidthCalculator ' Доменный расчетчик трещинообразования и владелец crack-диагностики.

' Подключает существующий расчетчик трещин без изменения его настроек.
Public Sub Initialize(ByVal crackCalculator As CCrackWidthCalculator)
    If crackCalculator Is Nothing Then Err.Raise vbObjectError + 4250, "CCrackLimitSearchProblem", _
        "Не передан расчетчик трещинообразования."
    Set mCrack = crackCalculator
End Sub

Private Property Get ILimitSearchProblem_DomainContext() As Object
    Set ILimitSearchProblem_DomainContext = mCrack
End Property

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
    ILimitSearchProblem_FinalizeAt = _
        mCrack.LimitSearchFinalizeCrackFormationLoadMultiplier(loadPath, lambdaValue, solver, iterations)
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
