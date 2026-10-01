VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCapacityLimitSearchProblem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILimitSearchProblem

' ==========================================================================
' Адаптер Capacity для общего LimitSearch
' ==========================================================================
' Класс прячет CCapacitySolver за общим ILimitSearchProblem. Одномерный
' search работает с этим адаптером как с источником probe-функции и приемником
' результата, не зная о внутренних полях несущей способности.

Private mCapacity As CCapacitySolver ' Доменный расчетчик несущей способности и владелец capacity-диагностики.

' Подключает уже настроенный CCapacitySolver. Настройки материалов, допусков и
' solver-а остаются ровно там же, где были до выделения adapter-а.
Public Sub Initialize(ByVal capacitySolver As CCapacitySolver)
    If capacitySolver Is Nothing Then Err.Raise vbObjectError + 4240, "CCapacityLimitSearchProblem", _
        "Не передан расчетчик несущей способности."
    Set mCapacity = capacitySolver
End Sub

Private Property Get ILimitSearchProblem_ResultKind() As EResultKind
    ILimitSearchProblem_ResultKind = rkCapacity
End Property

Private Property Get ILimitSearchProblem_InitialLambdaStep() As Double
    ILimitSearchProblem_InitialLambdaStep = mCapacity.LimitSearchInitialLambdaStep
End Property

Private Property Get ILimitSearchProblem_MaxLambda() As Double
    ILimitSearchProblem_MaxLambda = mCapacity.LimitSearchMaxLambda
End Property

Private Property Get ILimitSearchProblem_LambdaTolerance() As Double
    ILimitSearchProblem_LambdaTolerance = mCapacity.LimitSearchLambdaTolerance
End Property

Private Property Get ILimitSearchProblem_MaxIterations() As Long
    ILimitSearchProblem_MaxIterations = mCapacity.LimitSearchMaxIterations
End Property

Private Property Get ILimitSearchProblem_SearchMethod() As String
    ILimitSearchProblem_SearchMethod = mCapacity.LimitSearchCapacitySearchMethod()
End Property

Private Property Get ILimitSearchProblem_UsesRetainedForceCapacity() As Boolean
    ILimitSearchProblem_UsesRetainedForceCapacity = True
End Property

Private Property Get ILimitSearchProblem_Converged() As Boolean
    ILimitSearchProblem_Converged = mCapacity.Converged
End Property

Private Property Get ILimitSearchProblem_LimitState() As String
    ILimitSearchProblem_LimitState = mCapacity.LimitState
End Property

Private Property Get ILimitSearchProblem_DiagnosticLog() As String
    If Not mCapacity Is Nothing Then ILimitSearchProblem_DiagnosticLog = mCapacity.DiagnosticLog
End Property

' Подготавливает CCapacitySolver к LoadMultiplier-поиску.
Private Function ILimitSearchProblem_Prepare(ByVal request As CLimitSearchRequest) As Boolean
    ILimitSearchProblem_Prepare = mCapacity.LimitSearchPrepareCapacityLoadMultiplier( _
        request.Section, request.ConcreteMaterial, request.SteelMaterial, _
        request.NBase, request.MxBase, request.MyBase)
End Function

' Собирает capacity-траекторию Offset + lambda*Base.
Private Function ILimitSearchProblem_BuildLoadPath(ByVal request As CLimitSearchRequest) As CLoadPathVector
    Dim loadPath As CLoadPathVector
    Set loadPath = New CLoadPathVector
    loadPath.Initialize request.NOffset, request.NBase, request.MxOffset, request.MxBase, _
        request.MyOffset, request.MyBase
    Set ILimitSearchProblem_BuildLoadPath = loadPath
End Function

' Оценивает одну lambda-точку capacity через существующий CCapacitySolver.
Private Function ILimitSearchProblem_EvaluatePoint(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByVal lambdaValue As Double, ByRef functionValue As Double, _
        ByRef probeState As String, ByRef solver As CSectionSolver) As Boolean
    Set solver = Nothing
    ILimitSearchProblem_EvaluatePoint = mCapacity.LimitSearchEvaluateCapacityLoadMultiplier( _
        request.Section, request.ConcreteMaterial, request.SteelMaterial, _
        loadPath, lambdaValue, functionValue, probeState)
    If ILimitSearchProblem_EvaluatePoint Then
        ' Возвращаем общий search-слой ровно тот solver-снимок, который capacity
        ' использовал для оценки этой lambda-точки. Без этого финальный result
        ' может получить устаревшее LastSolver-состояние из предыдущей probe.
        Set solver = mCapacity.LastSolver
    End If
End Function

Private Function ILimitSearchProblem_IsAcceptableState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsAcceptableState = mCapacity.LimitSearchCapacityStateIsAcceptable(probeState)
End Function

Private Function ILimitSearchProblem_IsLimitState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsLimitState = mCapacity.LimitSearchCapacityStateIsPhysicalLimit(probeState)
End Function

Private Function ILimitSearchProblem_IsNumericalFailureState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsNumericalFailureState = mCapacity.LimitSearchCapacityStateIsNumericalFailure(probeState)
End Function

Private Function ILimitSearchProblem_IsInvalidInputState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsInvalidInputState = mCapacity.LimitSearchCapacityStateIsInvalidInput(probeState)
End Function

Private Function ILimitSearchProblem_ReduceLambdaAfterNumericalFailure(ByVal lowLambda As Double, _
        ByVal highLambda As Double, ByVal probeState As String) As Double
    ILimitSearchProblem_ReduceLambdaAfterNumericalFailure = _
        mCapacity.LimitSearchReduceCapacityLambdaAfterNumericalFailure(lowLambda, highLambda, probeState)
End Function

Private Sub ILimitSearchProblem_RecordAcceptable(ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double)
    mCapacity.LimitSearchRecordCapacityAcceptable loadPath.TargetN(lambdaValue), _
        loadPath.TargetMx(lambdaValue), loadPath.TargetMy(lambdaValue)
End Sub

Private Sub ILimitSearchProblem_HandleInitialFailure(ByVal probeState As String)
    mCapacity.LimitSearchHandleCapacityInitialFailure probeState
End Sub

Private Sub ILimitSearchProblem_HandleBracketNotFound(ByVal lowLambda As Double)
    mCapacity.LimitSearchHandleCapacityBracketNotFound lowLambda
End Sub

Private Sub ILimitSearchProblem_HandleBracketNumericalFailure(ByVal lowLambda As Double, _
        ByVal failedLambda As Double, ByVal probeState As String)
    mCapacity.LimitSearchHandleCapacityBracketNumericalFailure lowLambda, probeState
End Sub

Private Sub ILimitSearchProblem_HandleIterationLimit(ByVal maxIterations As Long, _
        ByVal remainingWidth As Double, ByVal tolerance As Double)
    mCapacity.LimitSearchSetCapacityNumericalFailure _
        "Поиск Bisection достиг лимита итераций до заданной точности: ширина скобки " & _
        FormatNumberInvariant(remainingWidth) & ", допуск " & FormatNumberInvariant(tolerance) & "."
End Sub

Private Sub ILimitSearchProblem_HandleUnexpectedState(ByVal probeState As String)
    mCapacity.LimitSearchSetCapacityNumericalFailure "Поиск Bisection остановлен: " & probeState & "."
End Sub

Private Sub ILimitSearchProblem_SetInvalidInput(ByVal message As String)
    mCapacity.LimitSearchSetCapacityInvalidInput message
End Sub

Private Sub ILimitSearchProblem_SetNumericalFailure(ByVal message As String)
    mCapacity.LimitSearchSetCapacityNumericalFailure message
End Sub

Private Sub ILimitSearchProblem_AppendDiagnostic(ByVal message As String)
    mCapacity.LimitSearchAppendCapacityDiagnostic message
End Sub

Private Function ILimitSearchProblem_FinalizeAt(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal probeState As String, ByVal methodName As String, _
        ByVal solver As CSectionSolver, ByVal iterations As Long) As Boolean
    mCapacity.LimitSearchFinalizeCapacityLoadMultiplier lambdaValue, loadPath, probeState, methodName, solver
    ILimitSearchProblem_FinalizeAt = mCapacity.Converged
End Function

Private Sub ILimitSearchProblem_FinalizeFromRetainedAcceptable(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal methodName As String)
    mCapacity.LimitSearchFinalizeCapacityFromRetainedAcceptable lambdaValue, loadPath, methodName
End Sub

Private Sub ILimitSearchProblem_AcceptRetainedCapacity(ByVal request As CLimitSearchRequest)
    mCapacity.LimitSearchAcceptPureAxialRetainedCapacity request.Section, _
        request.ConcreteMaterial, request.SteelMaterial, ILimitSearchProblem_BuildLoadPath(request)
End Sub

Private Function ILimitSearchProblem_BuildResult(ByVal request As CLimitSearchRequest, _
        ByVal searchSucceeded As Boolean) As CLimitSearchResult
    Dim result As CLimitSearchResult
    Set result = New CLimitSearchResult
    result.InitializeFromCapacitySolver mCapacity, request.SolutionStrategy
    Set ILimitSearchProblem_BuildResult = result
End Function

' Передает общему Newton уже проверенную Capacity-постановку.
Private Function ILimitSearchProblem_PrepareUltimate(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector) As Boolean
    ILimitSearchProblem_PrepareUltimate = mCapacity.LimitSearchPrepareCapacityUltimate( _
        request.Section, request.ConcreteMaterial, request.SteelMaterial, loadPath)
End Function

' Выполняет один общий поиск; дополнительные старты выбираются владельцем
' Capacity, а не отдельной реализацией численного цикла в адаптере.
Private Function ILimitSearchProblem_RunUltimateAttempts(ByVal request As CLimitSearchRequest) As Boolean
    Dim search As CUltimateStrainSearch
    Set search = New CUltimateStrainSearch
    ILimitSearchProblem_RunUltimateAttempts = search.RunCapacityUltimateNewton( _
        request, ILimitSearchProblem_BuildLoadPath(request))
End Function

Private Sub ILimitSearchProblem_CompleteUltimate(ByVal loadPath As CLoadPathVector)
    mCapacity.LimitSearchCompleteCapacityUltimate loadPath
End Sub

' Реализует численный callback-контракт критериями и настройками Capacity.
' Адаптер не решает систему и не выбирает Newton/line-search шаг.
Private Sub ILimitSearchProblem_LimitSearchInitializeUltimateGuess(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double)
    mCapacity.LimitSearchInitializeUltimateGuess section, concreteMaterial, steelMaterial, _
        loadPath, startSolver, eps0, kx, ky
End Sub

Private Function ILimitSearchProblem_LimitSearchEvaluateUltimateResidual(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal eps0 As Double, _
        ByVal kx As Double, ByVal ky As Double, ByRef solver As CSectionSolver, _
        ByRef rPath1 As Double, ByRef rPath2 As Double, ByRef rLimit As Double, _
        ByRef lambdaValue As Double, ByRef state As String) As Boolean
    ILimitSearchProblem_LimitSearchEvaluateUltimateResidual = mCapacity.LimitSearchEvaluateUltimateResidual( _
        section, concreteMaterial, steelMaterial, loadPath, eps0, kx, ky, solver, _
        rPath1, rPath2, rLimit, lambdaValue, state)
End Function

Private Function ILimitSearchProblem_LimitSearchUltimateResidualReached(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Boolean
    ILimitSearchProblem_LimitSearchUltimateResidualReached = _
        mCapacity.LimitSearchUltimateResidualReached(rPath1, rPath2, rLimit)
End Function

Private Function ILimitSearchProblem_LimitSearchUltimateResidualNorm(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Double
    ILimitSearchProblem_LimitSearchUltimateResidualNorm = _
        mCapacity.LimitSearchUltimateResidualNorm(rPath1, rPath2, rLimit)
End Function

Private Function ILimitSearchProblem_LimitSearchFinalizeUltimateResult(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal solver As CSectionSolver, ByVal state As String, ByVal iterations As Long) As Boolean
    ILimitSearchProblem_LimitSearchFinalizeUltimateResult = mCapacity.LimitSearchFinalizeUltimateResult( _
        section, concreteMaterial, steelMaterial, loadPath, lambdaValue, solver, state, iterations)
End Function

Private Sub ILimitSearchProblem_LimitSearchSetUltimateNumericalFailure(ByVal reason As String)
    mCapacity.LimitSearchSetUltimateNumericalFailure reason
End Sub

' Однотипные настройки общего Newton из численной конфигурации Capacity.
Private Property Get ILimitSearchProblem_LimitSearchUltimateMaxIterations() As Long
    ILimitSearchProblem_LimitSearchUltimateMaxIterations = mCapacity.LimitSearchUltimateMaxIterations
End Property

Private Property Get ILimitSearchProblem_LimitSearchUltimateDampingInitial() As Double
    ILimitSearchProblem_LimitSearchUltimateDampingInitial = mCapacity.LimitSearchUltimateDampingInitial
End Property

Private Property Get ILimitSearchProblem_LimitSearchUltimateMinLineSearchAlpha() As Double
    ILimitSearchProblem_LimitSearchUltimateMinLineSearchAlpha = mCapacity.LimitSearchUltimateMinLineSearchAlpha
End Property

Private Property Get ILimitSearchProblem_LimitSearchUltimateLineSearchEnabled() As Boolean
    ILimitSearchProblem_LimitSearchUltimateLineSearchEnabled = mCapacity.LimitSearchUltimateLineSearchEnabled
End Property

Private Property Get ILimitSearchProblem_LimitSearchUltimateMaxDeltaEpsilon0() As Double
    ILimitSearchProblem_LimitSearchUltimateMaxDeltaEpsilon0 = mCapacity.LimitSearchUltimateMaxDeltaEpsilon0
End Property

Private Property Get ILimitSearchProblem_LimitSearchUltimateMaxDeltaKappa() As Double
    ILimitSearchProblem_LimitSearchUltimateMaxDeltaKappa = mCapacity.LimitSearchUltimateMaxDeltaKappa
End Property

' Capacity не использует crack-операции Auto-пути; методы оставлены no-op по
' контракту ILimitSearchProblem.
Private Sub ILimitSearchProblem_BeginFormationAttempt(ByVal pathName As String)
End Sub

Private Sub ILimitSearchProblem_SetResolvedFormationPath(ByVal pathName As String)
End Sub

Private Sub ILimitSearchProblem_SetCurrentFormationPath(ByVal pathName As String)
End Sub

Private Property Get ILimitSearchProblem_FormationFallbackOnly() As Boolean
End Property

Private Property Get ILimitSearchProblem_StoppedTerminally() As Boolean
End Property

Private Property Get ILimitSearchProblem_StopReason() As String
    If Not mCapacity Is Nothing Then ILimitSearchProblem_StopReason = mCapacity.StopReason
End Property

Private Sub ILimitSearchProblem_ClearStopReason()
End Sub

Private Sub ILimitSearchProblem_SetStopReasonIfEmpty(ByVal reason As String)
End Sub

Private Function ILimitSearchProblem_FormationPathDisplayName() As String
    ILimitSearchProblem_FormationPathDisplayName = vbNullString
End Function

' Форматирует число для диагностического журнала с точкой как разделителем.
Private Function ILimitSearchProblem_HasScaledFormationLoad(ByVal request As CLimitSearchRequest) As Boolean
    ILimitSearchProblem_HasScaledFormationLoad = True
End Function

Private Sub ILimitSearchProblem_MarkFormationPathFallback(ByVal reason As String)
End Sub

Private Sub ILimitSearchProblem_FinishNoCrack(ByVal reason As String)
End Sub

Private Sub ILimitSearchProblem_MergeFallback(ByVal firstLimitState As String, ByVal firstStopReason As String, _
        ByVal firstDiagnostics As String, ByVal firstAttemptMessage As String, _
        ByVal transitionMessage As String, ByVal successMessage As String)
    mCapacity.LimitSearchMergeCapacityFallback firstLimitState, firstStopReason, _
        firstDiagnostics, firstAttemptMessage, transitionMessage, successMessage
End Sub

Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function
