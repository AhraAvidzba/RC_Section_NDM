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

' ==========================================================================
' Адаптер Capacity для общего LimitSearch
' ==========================================================================
' Класс прячет CCapacitySolver за общим ILimitSearchProblem. Одномерный
' search работает с этим адаптером как с источником probe-функции и приемником
' результата, не зная о внутренних полях несущей способности.

Implements ILimitSearchProblem

Private mCapacity As CCapacitySolver ' Доменный расчетчик несущей способности и владелец capacity-диагностики.

' Начальное состояние каждой capacity search-session очищает ее численный
' solver. Общий coordinator уже передал новый Offset/Base через request.
Private Sub ILimitSearchProblem_BeginLoadPath(ByVal request As CLimitSearchRequest)
End Sub

' Историю путей сохраняет CLimitSearchResult; Capacity трактует его готовую
' машинную причину и не склеивает собственную параллельную сводку попыток.
Private Sub ILimitSearchProblem_FinishLoadPath(ByVal request As CLimitSearchRequest, _
        ByVal previousPathComments As String)
End Sub

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

' Несущая уточняет уже подтвержденную скобку до допуска lambda.
' Малый лимит внутреннего solve не ограничивает количество делений скобки.
Private Property Get ILimitSearchProblem_BisectionMaxIterations() As Long
    ILimitSearchProblem_BisectionMaxIterations = 0
End Property

Private Property Get ILimitSearchProblem_SearchMethod() As String
    ILimitSearchProblem_SearchMethod = mCapacity.LimitSearchCapacitySearchMethod()
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

' Начало capacity-пути должно удовлетворять исходному физическому критерию.
' Эта доменная оценка не зависит от алгоритма расширения численной скобки.
Private Function ILimitSearchProblem_AcceptInitialPoint(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal functionValue As Double, _
        ByVal probeState As String, ByVal solver As CSectionSolver) As Boolean
    If Not mCapacity.LimitSearchCapacityStateIsAcceptable(probeState) Then
        mCapacity.LimitSearchHandleCapacityInitialFailure probeState
        Exit Function
    End If
    ILimitSearchProblem_RecordAcceptable loadPath, 0#
    ILimitSearchProblem_AcceptInitialPoint = True
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
        ' На осевом плато повторные силовые пробы за максимумом не дают
        ' физической верхней точки. Уже найденная допустимая проба позволяет
        ' проверить отдельную предельную плоскость без длительного дробления.
        ' Успех здесь означает проверенные N/Mx/My и физический критерий,
        ' а не признание несошедшейся пробной плоскости предельной.
        If request.PureAxialForMethod And lambdaValue > 0# Then
            If mCapacity.LimitSearchCapacityStateIsNumericalFailure(probeState) Then
                mCapacity.LimitSearchAcceptPureAxialRetainedCapacity request.Section, _
                    request.ConcreteMaterial, request.SteelMaterial, loadPath
                If mCapacity.Converged Then ILimitSearchProblem_EvaluatePoint = False
            End If
        End If
    End If
End Function

' Однотипные предикаты передают классификацию probe владельцу Capacity.
' Общий Search отличает допустимую пробу, предел и техническую неудачу,
' но не назначает инженерный BaseFail по строкам внешних статусов.
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

' Запрашивает восстановительную lambda у Capacity после неуспешной пробы.
' Численная неудача не становится физической верхней границей автоматически.
Private Function ILimitSearchProblem_ReduceLambdaAfterNumericalFailure(ByVal lowLambda As Double, _
        ByVal highLambda As Double, ByVal probeState As String) As Double
    ILimitSearchProblem_ReduceLambdaAfterNumericalFailure = _
        mCapacity.LimitSearchReduceCapacityLambdaAfterNumericalFailure(lowLambda, highLambda, probeState)
End Function

' Сохраняет усилия подтвержденной допустимой точки для последующей финализации.
' Все компоненты вычисляются на той же общей траектории нагрузки.
Private Sub ILimitSearchProblem_RecordAcceptable(ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double)
    mCapacity.LimitSearchRecordCapacityAcceptable loadPath.TargetN(lambdaValue), _
        loadPath.TargetMx(lambdaValue), loadPath.TargetMy(lambdaValue)
End Sub

' Передает исход начальной пробы владельцу инженерной интерпретации Capacity.
' Несходимость и подтвержденное превышение критерия имеют разные причины.
Private Sub ILimitSearchProblem_HandleInitialFailure(ByVal probeState As String)
    mCapacity.LimitSearchHandleCapacityInitialFailure probeState
End Sub

' Завершает поиск без верхней физической точки, сохраняя последнюю допустимую
' lambda. Достижение технической границы не объявляется найденной capacity.
Private Sub ILimitSearchProblem_HandleBracketNotFound(ByVal lowLambda As Double)
    mCapacity.LimitSearchHandleCapacityBracketNotFound lowLambda
End Sub

' Сохраняет численную причину невозможности построить скобку после recovery.
' Последняя допустимая точка остается диагностикой, а не доказанным пределом.
Private Sub ILimitSearchProblem_HandleBracketNumericalFailure(ByVal lowLambda As Double, _
        ByVal failedLambda As Double, ByVal probeState As String)
    mCapacity.LimitSearchHandleCapacityBracketNumericalFailure lowLambda, probeState
End Sub

' Объясняет исчерпание бюджета до требуемой точности скобки. Допуск пользователя
' не увеличивается для превращения незавершенного поиска в успешный.
Private Sub ILimitSearchProblem_HandleIterationLimit(ByVal maxIterations As Long, _
        ByVal remainingWidth As Double, ByVal tolerance As Double)
    mCapacity.LimitSearchSetCapacityNumericalFailure _
        "Поиск Bisection достиг лимита итераций до заданной точности: ширина скобки " & _
        FormatNumberInvariant(remainingWidth) & ", допуск " & FormatNumberInvariant(tolerance) & "."
End Sub

' Останавливает поиск при неподдержанном исходе probe и сохраняет его имя
' для диагностики; такой исход не подтверждает физическое разрушение.
Private Sub ILimitSearchProblem_HandleUnexpectedState(ByVal probeState As String)
    mCapacity.LimitSearchSetCapacityNumericalFailure "Поиск Bisection остановлен: " & probeState & "."
End Sub

' Передает ошибку параметров поиска без попытки представить ее как несходимость.
Private Sub ILimitSearchProblem_SetInvalidInput(ByVal message As String)
    mCapacity.LimitSearchSetCapacityInvalidInput message
End Sub

' Передает численную причину общего поиска в результат Capacity.
Private Sub ILimitSearchProblem_SetNumericalFailure(ByVal message As String)
    mCapacity.LimitSearchSetCapacityNumericalFailure message
End Sub

' Дополняет журнал текущей search-session, не меняя инженерный статус результата.
Private Sub ILimitSearchProblem_AppendDiagnostic(ByVal message As String)
    mCapacity.LimitSearchAppendCapacityDiagnostic message
End Sub

' Подтверждает конкретную lambda как конечный физический предел. Для осевого
' плато дополнительно проверяется предельная плоскость и равновесие N/Mx/My.
Private Function ILimitSearchProblem_FinalizeAt(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal probeState As String, ByVal methodName As String, _
        ByVal solver As CSectionSolver, ByVal iterations As Long) As Boolean
    If Not request.PureAxialForMethod Then
        Dim runner As CStateSolutionRunner
        Set runner = New CStateSolutionRunner
        If Not runner.StateWithinPhysicalRange(request.Section, solver, request.ConcreteMaterial, request.SteelMaterial) Then
            mCapacity.LimitSearchAppendCapacityDiagnostic _
                methodName & ": кандидат конечной точки вне физического диапазона; продолжаем уточнение скобки."
            Exit Function
        End If
    End If
    mCapacity.LimitSearchFinalizeCapacityLoadMultiplier lambdaValue, loadPath, probeState, methodName, solver
    ' На осевом плато одна и та же сила соответствует разным деформациям.
    ' Нижняя проба силовой скобки не определяет предельную плоскость: доводим
    ' ее до исходного физического предела и отдельно подтверждаем N/Mx/My.
    If request.PureAxialForMethod Then mCapacity.LimitSearchAcceptPureAxialRetainedCapacity _
        request.Section, request.ConcreteMaterial, request.SteelMaterial, loadPath
    ILimitSearchProblem_FinalizeAt = mCapacity.Converged
End Function

' Финализирует сохраненную допустимую сторону уточненной скобки; проверка
' физического критерия остается у Capacity, включая неоднозначное осевое плато.
Private Sub ILimitSearchProblem_FinalizeBracket(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lowLambda As Double, ByVal lowSolver As CSectionSolver, _
        ByVal highLambda As Double, ByVal highSolver As CSectionSolver, _
        ByVal methodName As String, ByVal iterations As Long)
    If lowSolver Is Nothing Then
        mCapacity.LimitSearchFinalizeCapacityFromRetainedAcceptable lowLambda, loadPath, methodName
    Else
        mCapacity.LimitSearchFinalizeCapacityLoadMultiplier lowLambda, loadPath, vbNullString, methodName, lowSolver
    End If
    If request.PureAxialForMethod Then mCapacity.LimitSearchAcceptPureAxialRetainedCapacity _
        request.Section, request.ConcreteMaterial, request.SteelMaterial, loadPath
End Sub

' Для силового осевого пути пробует физически подтвердить сохраненную точку
' после неудачи общего поиска. LambdaN с постоянными моментами тоже может
' выйти на силовое плато: готовая плоскость проверяется на том же Offset/Base,
' без замены фиксированных моментов и без признания failed probe пределом.
Private Sub ILimitSearchProblem_TryFinalizeAfterSearchFailure(ByVal request As CLimitSearchRequest)
    If mCapacity.Converged Then Exit Sub
    If mCapacity.LimitSearchHasTerminalFailure Then Exit Sub
    If Not request.ForceOnly And Not request.PureAxialForMethod Then
        If request.SelectedLoadPath Is Nothing Then Exit Sub
        If Not request.SelectedLoadPath.ScalesN Then Exit Sub
        If request.SelectedLoadPath.ScalesMx Or request.SelectedLoadPath.ScalesMy Then Exit Sub
    End If
    mCapacity.LimitSearchAcceptPureAxialRetainedCapacity request.Section, _
        request.ConcreteMaterial, request.SteelMaterial, ILimitSearchProblem_BuildLoadPath(request)
End Sub

' Собирает итог из фактических данных Capacity, а не только Boolean общего
' алгоритма. Это сохраняет точную причину остановки и конечное named-state.
Private Function ILimitSearchProblem_BuildResult(ByVal request As CLimitSearchRequest, _
        ByVal searchSucceeded As Boolean) As CLimitSearchResult
    Set ILimitSearchProblem_BuildResult = BuildSnapshot(request.SolutionStrategy, request.Section, _
        request.ConcreteMaterial, request.SteelMaterial, request.MaterialSpec)
End Function

' Интерпретирует машинный исход capacity-задачи и фиксирует конечное НДС.
' Общий Search-result не знает ни правил BaseFail, ни полей CCapacitySolver.
' Lambda ниже единицы остается успехом поиска; проверку текущего LC делает CapacityResult.
Public Function BuildSnapshot(ByVal requestedStrategy As String, _
        Optional ByVal section As CSectionModel = Nothing, _
        Optional ByVal concreteMaterial As Object = Nothing, _
        Optional ByVal steelMaterial As Object = Nothing, _
        Optional ByVal materialSpec As CMaterialModelSpec = Nothing) As CLimitSearchResult
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    Dim searchExecuted As Boolean
    If Not mCapacity Is Nothing Then searchExecuted = mCapacity.SearchExecuted
    If mCapacity Is Nothing Then
        meta.SetResult rsInternalError, rcInternalError, rkCapacity, "Не передана задача несущей способности.", vbNullString, True, False
    ElseIf mCapacity.Converged Then
        meta.SetResult rsSuccess, rcCheckPassed, rkCapacity, mCapacity.StopReason
    ElseIf mCapacity.FailureCode = sfcInvalidConfiguration Then
        meta.SetResult rsInvalidConfiguration, rcInvalidConfiguration, rkCapacity, mCapacity.StopReason
    ElseIf mCapacity.FailureCode = sfcInternalError Then
        meta.SetResult rsInternalError, rcInternalError, rkCapacity, mCapacity.StopReason, _
            vbNullString, True, searchExecuted
    ElseIf mCapacity.SearchBoundReached Then
        meta.SetResult rsNumericalFailure, rcSearchBoundReached, rkCapacity, mCapacity.StopReason
    ElseIf mCapacity.LimitSearchInitialLambdaRejected And _
            Not mCapacity.LimitSearchCapacityStateIsInvalidInput(mCapacity.LimitState) Then
        meta.SetResult rsCheckFailed, rcInitialStateBeyondLimit, rkCapacity, _
            mCapacity.StopReason
    ElseIf mCapacity.LimitSearchCapacityStateIsInvalidInput(mCapacity.LimitState) Then
        meta.SetResult rsInvalidInput, rcInvalidInput, rkCapacity, mCapacity.StopReason
    ElseIf mCapacity.LimitSearchCapacityStateIsNumericalFailure(mCapacity.LimitState) Then
        Dim failureResultCode As EResultCode
        failureResultCode = rcNumericalFailure
        If mCapacity.FailureCode = sfcSingularTangent Then failureResultCode = rcSingularTangent
        meta.SetResult rsNumericalFailure, failureResultCode, rkCapacity, mCapacity.StopReason
    Else
        meta.SetResult rsInternalError, rcInternalError, rkCapacity, mCapacity.StopReason, _
            vbNullString, True, searchExecuted
    End If

    Dim pointState As CSectionStateResult
    Dim hasPoint As Boolean
    Dim finalDiagnostic As String
    Dim lambdaValue As Double
    Dim nValue As Double, mxValue As Double, myValue As Double
    Dim methodName As String, criterionName As String, diagnostics As String
    If Not mCapacity Is Nothing Then
        hasPoint = mCapacity.Converged And mCapacity.LambdaUltimate > 0#
        lambdaValue = mCapacity.LambdaUltimate
        nValue = mCapacity.NUltimate
        mxValue = mCapacity.MxUltimate
        myValue = mCapacity.MyUltimate
        methodName = mCapacity.SolutionMethod
        criterionName = mCapacity.LimitState
        diagnostics = mCapacity.DiagnosticLog
        If Not mCapacity.LastSolver Is Nothing Then
            finalDiagnostic = mCapacity.LastSolver.DiagnosticLog
            If hasPoint And Not materialSpec Is Nothing Then
                Dim runner As CStateSolutionRunner
                Set runner = New CStateSolutionRunner
                Set pointState = New CSectionStateResult
                pointState.InitializeFromSolver sstCapacityState, cpStrength, materialSpec, mCapacity.LastSolver, _
                    runner.StateUsesExtension(section, mCapacity.LastSolver, concreteMaterial, steelMaterial), _
                    runner.StateWithinPhysicalRange(section, mCapacity.LastSolver, concreteMaterial, steelMaterial), _
                    0, diagnostics
            End If
        End If
    End If
    Dim result As CLimitSearchResult
    Set result = New CLimitSearchResult
    result.Initialize meta, requestedStrategy, methodName, hasPoint, lambdaValue, _
        nValue, mxValue, myValue, pointState, criterionName, diagnostics, finalDiagnostic, searchExecuted
    Set BuildSnapshot = result
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

' Завершает инженерные величины Capacity после принятого UltimateStrain-решения
' на исходной траектории; численный цикл здесь не повторяется.
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

' Вычисляет две невязки принадлежности траектории и невязку предельной
' деформации через Capacity. Общий Newton получает тот же solver этой плоскости.
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

' Передает проверку трех невязок с исходными допусками Capacity.
Private Function ILimitSearchProblem_LimitSearchUltimateResidualReached(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Boolean
    ILimitSearchProblem_LimitSearchUltimateResidualReached = _
        mCapacity.LimitSearchUltimateResidualReached(rPath1, rPath2, rLimit)
End Function

' Возвращает норму невязок для выбора шага общего Newton, не заменяя ею
' отдельную проверку достижения критерия и компонентных допусков.
Private Function ILimitSearchProblem_LimitSearchUltimateResidualNorm(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Double
    ILimitSearchProblem_LimitSearchUltimateResidualNorm = _
        mCapacity.LimitSearchUltimateResidualNorm(rPath1, rPath2, rLimit)
End Function

' Передает найденную плоскость на окончательное подтверждение Capacity.
' Вспомогательное равновесие на extension не считается физическим пределом.
Private Function ILimitSearchProblem_LimitSearchFinalizeUltimateResult(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal solver As CSectionSolver, ByVal state As String, ByVal iterations As Long) As Boolean
    ILimitSearchProblem_LimitSearchFinalizeUltimateResult = mCapacity.LimitSearchFinalizeUltimateResult( _
        section, concreteMaterial, steelMaterial, loadPath, lambdaValue, solver, state, iterations)
End Function

' Сохраняет численную остановку UltimateStrain у владельца Capacity-результата.
Private Sub ILimitSearchProblem_LimitSearchSetUltimateNumericalFailure(ByVal reason As String)
    mCapacity.LimitSearchSetUltimateNumericalFailure reason
End Sub

' Передает машинную причину общего алгоритма владельцу Capacity-результата.
Private Sub ILimitSearchProblem_SetFailure(ByVal failureCode As ESolverFailureCode, ByVal reason As String)
    mCapacity.LimitSearchSetFailure failureCode, reason
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

' Сообщает общему coordinator, что получен окончательный успех или терминальная
' ошибка Capacity. Отсутствующая задача не выдает ложный признак завершения.
Private Property Get ILimitSearchProblem_StoppedTerminally() As Boolean
    If Not mCapacity Is Nothing Then _
        ILimitSearchProblem_StoppedTerminally = mCapacity.Converged Or mCapacity.LimitSearchHasTerminalFailure
End Property

' Возвращает причину текущей задачи; без Capacity причина остается пустой.
Private Property Get ILimitSearchProblem_StopReason() As String
    If Not mCapacity Is Nothing Then ILimitSearchProblem_StopReason = mCapacity.StopReason
End Property

' Capacity очищает рабочий итог при подготовке своей попытки; отдельный сброс
' coordinator не должен стирать сохраненную typed-причину.
Private Sub ILimitSearchProblem_ClearStopReason()
End Sub

' Общий fallback не подменяет причину Capacity общим текстом: результат
' формируется из машинного исхода и инженерной интерпретации самой задачи.
Private Sub ILimitSearchProblem_SetStopReasonIfEmpty(ByVal reason As String)
End Sub

' Объединяет диагностику первой попытки и перехода к fallback у Capacity.
' Комментарий конечного результата формируется владельцем, не writer-ом.
Private Sub ILimitSearchProblem_MergeFallback(ByVal firstLimitState As String, ByVal firstStopReason As String, _
        ByVal firstDiagnostics As String, ByVal firstAttemptMessage As String, _
        ByVal transitionMessage As String, ByVal successMessage As String, _
        Optional ByVal firstSearchExecuted As Boolean = False)
    mCapacity.LimitSearchMergeCapacityFallback firstLimitState, firstStopReason, _
        firstDiagnostics, firstAttemptMessage, transitionMessage, successMessage, firstSearchExecuted
End Sub

Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

' Форматирует число для диагностического журнала с точкой как разделителем.
Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function
