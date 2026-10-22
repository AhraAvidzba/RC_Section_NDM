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
' Связь расчёта несущей способности с общим поиском предела
' ==========================================================================
' Реализует ILimitSearchProblem через CCapacitySolver: передаёт оценку пробных
' точек, критерии и приём результата. Общий численный поиск использует
' эти операции независимо от внутренних данных расчёта несущей способности.

Implements ILimitSearchProblem

Private mCapacity As CCapacitySolver ' Расчёт несущей способности, который оценивает нагрузки, принимает физический предел и хранит диагностику.

' В начале нового пути очищает рабочие данные решателя.
' Общий координатор уже передал постоянную и масштабируемую части нагрузки.
Private Sub ILimitSearchProblem_BeginLoadPath(ByVal request As CLimitSearchRequest)
End Sub

' Передаёт расчёту несущей способности итоговую причину и историю пробовавшихся путей.
' История сохраняется в CLimitSearchResult без параллельного составления того же текста.
Private Sub ILimitSearchProblem_FinishLoadPath(ByVal request As CLimitSearchRequest, _
        ByVal previousPathComments As String)
End Sub

' Подключает настроенный CCapacitySolver для оценки пробных нагрузок
' и принятия результата общего поиска. Материалы и численные настройки не меняет.
Public Sub Initialize(ByVal capacitySolver As CCapacitySolver)
    If capacitySolver Is Nothing Then Err.Raise vbObjectError + 4240, "CCapacityLimitSearchProblem", _
        "Не передан расчетчик несущей способности."
    Set mCapacity = capacitySolver
End Sub

' Тип результата, публикуемого данной задачей предельного поиска.
Private Property Get ILimitSearchProblem_ResultKind() As EResultKind
    ILimitSearchProblem_ResultKind = rkCapacity
End Property

' Возвращает начальный шаг множителя нагрузки lambda для поиска предела.
Private Property Get ILimitSearchProblem_InitialLambdaStep() As Double
    ILimitSearchProblem_InitialLambdaStep = mCapacity.LimitSearchInitialLambdaStep
End Property

' Возвращает заданную техническую верхнюю границу множителя lambda.
Private Property Get ILimitSearchProblem_MaxLambda() As Double
    ILimitSearchProblem_MaxLambda = mCapacity.LimitSearchMaxLambda
End Property

' Возвращает допуск уточнения предельного множителя нагрузки lambda.
Private Property Get ILimitSearchProblem_LambdaTolerance() As Double
    ILimitSearchProblem_LambdaTolerance = mCapacity.LimitSearchLambdaTolerance
End Property

' Возвращает максимальное число итераций общего поиска предела.
Private Property Get ILimitSearchProblem_MaxIterations() As Long
    ILimitSearchProblem_MaxIterations = mCapacity.LimitSearchMaxIterations
End Property

' Несущая уточняет уже подтвержденную скобку до допуска lambda.
' Малый лимит внутреннего solve не ограничивает количество делений скобки.
Private Property Get ILimitSearchProblem_BisectionMaxIterations() As Long
    ILimitSearchProblem_BisectionMaxIterations = 0
End Property

' Возвращает выбранный численный метод поиска предельной нагрузки.
Private Property Get ILimitSearchProblem_SearchMethod() As String
    ILimitSearchProblem_SearchMethod = mCapacity.LimitSearchCapacitySearchMethod()
End Property

' Показывает, подтверждена ли сходимость поиска предельной точки.
Private Property Get ILimitSearchProblem_Converged() As Boolean
    ILimitSearchProblem_Converged = mCapacity.Converged
End Property

' Возвращает сохранённый код физического предела или исхода поиска.
Private Property Get ILimitSearchProblem_LimitState() As String
    ILimitSearchProblem_LimitState = mCapacity.LimitState
End Property

' Возвращает накопленный диагностический журнал поиска предела.
Private Property Get ILimitSearchProblem_DiagnosticLog() As String
    If Not mCapacity Is Nothing Then ILimitSearchProblem_DiagnosticLog = mCapacity.DiagnosticLog
End Property

' Подготавливает CCapacitySolver к LoadMultiplier-поиску.
Private Function ILimitSearchProblem_Prepare(ByVal request As CLimitSearchRequest) As Boolean
    ILimitSearchProblem_Prepare = mCapacity.LimitSearchPrepareCapacityLoadMultiplier( _
        request.Section, request.ConcreteMaterial, request.SteelMaterial, _
        request.NBase, request.MxBase, request.MyBase)
End Function

' Готовит зависимость усилий Offset + lambda*Base для поиска несущей способности.
Private Function ILimitSearchProblem_BuildLoadPath(ByVal request As CLimitSearchRequest) As CLoadPathVector
    Dim loadPath As CLoadPathVector
    Set loadPath = New CLoadPathVector
    loadPath.Initialize request.NOffset, request.NBase, request.MxOffset, request.MxBase, _
        request.MyOffset, request.MyBase
    Set ILimitSearchProblem_BuildLoadPath = loadPath
End Function

' Проверяет физическую допустимость начальной нагрузки выбранного пути.
' Эта проверка выполняется независимо от способа увеличения и уточнения lambda.
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

' Оценивает НДС и критерий несущей способности при одном пробном lambda через CCapacitySolver.
Private Function ILimitSearchProblem_EvaluatePoint(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByVal lambdaValue As Double, ByRef functionValue As Double, _
        ByRef probeState As String, ByRef solver As CSectionSolver) As Boolean
    Set solver = Nothing
    ILimitSearchProblem_EvaluatePoint = mCapacity.LimitSearchEvaluateCapacityLoadMultiplier( _
        request.Section, request.ConcreteMaterial, request.SteelMaterial, _
        loadPath, lambdaValue, functionValue, probeState, startSolver)
    If ILimitSearchProblem_EvaluatePoint Then
        Set solver = mCapacity.LastSolver ' Возвращаем общему поиску именно то НДС, по которому оценено текущее lambda. Иначе конечный результат мог бы получить устаревшее состояние предыдущей пробы.
        ' На осевом плато повторные силовые пробы за максимумом не дают
        ' физической верхней точки. Уже найденная допустимая проба позволяет
        ' проверить отдельную предельную плоскость без длительного дробления.
        ' Успех здесь означает проверенные N/Mx/My и физический критерий,
        ' а не признание несошедшейся пробной плоскости предельной.
        If request.PureAxialForMethod And lambdaValue > 0# Then
            If mCapacity.LimitSearchCapacityStateIsNumericalFailure(probeState) Then
                mCapacity.LimitSearchAcceptRetainedPhysicalCapacity request.Section, _
                    request.ConcreteMaterial, request.SteelMaterial, loadPath
                If mCapacity.Converged Then ILimitSearchProblem_EvaluatePoint = False
            End If
        End If
    End If
End Function

' Различает допустимую нагрузку, достигнутый предел и численную неудачу
' по признакам расчёта несущей способности. Инженерный BaseFail
' не выводится из текста внешнего статуса.
Private Function ILimitSearchProblem_IsAcceptableState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsAcceptableState = mCapacity.LimitSearchCapacityStateIsAcceptable(probeState)
End Function

' Проверяет физический предельный статус задачи; численная несходимость сама по себе не означает
' разрушение.
Private Function ILimitSearchProblem_IsLimitState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsLimitState = mCapacity.LimitSearchCapacityStateIsPhysicalLimit(probeState)
End Function

' Распознаёт численную неудачу отдельно от физического предела.
Private Function ILimitSearchProblem_IsNumericalFailureState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsNumericalFailureState = mCapacity.LimitSearchCapacityStateIsNumericalFailure(probeState)
End Function

' Распознаёт ошибку исходных данных, при которой смена поискового пути недопустима.
Private Function ILimitSearchProblem_IsInvalidInputState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsInvalidInputState = mCapacity.LimitSearchCapacityStateIsInvalidInput(probeState)
End Function

' Запрашивает новое уменьшенное lambda после неудачного решения НДС.
' Сама несходимость ещё не подтверждает физическую верхнюю границу нагрузки.
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

' Передаёт исход первой пробы расчёту несущей способности.
' Численная несходимость и подтверждённое превышение физического критерия различаются.
Private Sub ILimitSearchProblem_HandleInitialFailure(ByVal probeState As String)
    mCapacity.LimitSearchHandleCapacityInitialFailure probeState
End Sub

' Завершает поиск, если физическая верхняя точка не найдена, сохраняя последнее
' допустимое lambda. Техническая граница поиска не считается несущей способностью.
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

' Останавливает поиск при неизвестном исходе пробы и записывает его имя.
' Такой исход сам по себе не доказывает физическое разрушение.
Private Sub ILimitSearchProblem_HandleUnexpectedState(ByVal probeState As String)
    mCapacity.LimitSearchSetCapacityNumericalFailure "Поиск Bisection остановлен: " & probeState & "."
End Sub

' Передает ошибку параметров поиска без попытки представить ее как несходимость.
Private Sub ILimitSearchProblem_SetInvalidInput(ByVal message As String)
    mCapacity.LimitSearchSetCapacityInvalidInput message
End Sub

' Передаёт фактическую численную причину остановки в результат несущей способности.
Private Sub ILimitSearchProblem_SetNumericalFailure(ByVal message As String)
    mCapacity.LimitSearchSetCapacityNumericalFailure message
End Sub

' Добавляет сообщение текущей попытки в журнал, не меняя инженерный статус.
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
    If request.PureAxialForMethod Then mCapacity.LimitSearchAcceptRetainedPhysicalCapacity _
        request.Section, request.ConcreteMaterial, request.SteelMaterial, loadPath
    ILimitSearchProblem_FinalizeAt = mCapacity.Converged
End Function

' Передаёт допустимую сторону уточнённого интервала на окончательную проверку
' несущей способности, включая неоднозначный случай осевого плато диаграмм.
Private Function ILimitSearchProblem_FinalizeBracket(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lowLambda As Double, ByVal lowSolver As CSectionSolver, _
        ByVal highLambda As Double, ByVal highSolver As CSectionSolver, _
        ByVal methodName As String, ByVal iterations As Long) As Boolean
    If lowSolver Is Nothing Then
        mCapacity.LimitSearchFinalizeCapacityFromRetainedAcceptable lowLambda, loadPath, methodName
    Else
        mCapacity.LimitSearchFinalizeCapacityLoadMultiplier lowLambda, loadPath, vbNullString, methodName, lowSolver
    End If
    If request.PureAxialForMethod Then mCapacity.LimitSearchAcceptRetainedPhysicalCapacity _
        request.Section, request.ConcreteMaterial, request.SteelMaterial, loadPath
    ILimitSearchProblem_FinalizeBracket = mCapacity.Converged
End Function

' После численной неудачи уточняет предельную плоскость из сохраненного НДС.
' Приемка требует исходного Offset/Base, равновесия всех трех усилий и
' физического критерия; несошедшаяся lambda не становится верхней границей.
Private Sub ILimitSearchProblem_TryFinalizeAfterSearchFailure(ByVal request As CLimitSearchRequest)
    If mCapacity.Converged Then Exit Sub
    If mCapacity.LimitSearchHasTerminalFailure Then Exit Sub
    If Not request.ForceOnly And Not request.PureAxialForMethod Then
        If Not mCapacity.LimitSearchHasFailedEquilibriumProbe Then Exit Sub
    End If
    mCapacity.LimitSearchAcceptRetainedPhysicalCapacity request.Section, _
        request.ConcreteMaterial, request.SteelMaterial, ILimitSearchProblem_BuildLoadPath(request)
End Sub

' Собирает результат по фактическому исходу расчёта: причине остановки
' и конечному НДС. Одного признака завершения общего алгоритма недостаточно.
Private Function ILimitSearchProblem_BuildResult(ByVal request As CLimitSearchRequest, _
        ByVal searchSucceeded As Boolean) As CLimitSearchResult
    Set ILimitSearchProblem_BuildResult = BuildSnapshot(request.SolutionStrategy, request.Section, _
        request.ConcreteMaterial, request.SteelMaterial, request.MaterialSpec)
End Function

' Сохраняет конечное НДС и физическую причину исхода. Lambda ниже единицы
' остаётся успешно найденным пределом; непрохождение заданной нагрузки
' отдельно определяет CCapacityResult.
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

    If Not mCapacity Is Nothing Then meta.SourceKeys = mCapacity.FailureSettingKeys
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

' Передаёт общему методу Ньютона проверенные данные задачи несущей способности.
Private Function ILimitSearchProblem_PrepareUltimate(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector) As Boolean
    If Not mCapacity.LimitSearchPrepareCapacityUltimate( _
            request.Section, request.ConcreteMaterial, request.SteelMaterial, loadPath) Then Exit Function
    If loadPath.NOffset <> 0# Or loadPath.MxOffset <> 0# Or loadPath.MyOffset <> 0# Then
        ' У аффинного пути сначала проверяем постоянную нагрузку. Равновесная
        ' точка за пределом деформаций отклоняется тем же правилом, что в
        ' LoadMultiplier; искать ее границу через Newton уже не требуется.
        Dim value As Double, state As String, solver As CSectionSolver
        If mCapacity.LimitSearchEvaluateCapacityLoadMultiplier(request.Section, _
                request.ConcreteMaterial, request.SteelMaterial, loadPath, 0#, value, state, Nothing, False) Then
            If ILimitSearchProblem_IsLimitState(state) Then
                Set solver = mCapacity.LastSolver
                ILimitSearchProblem_PrepareUltimate = ILimitSearchProblem_AcceptInitialPoint( _
                    request, loadPath, value, state, solver)
                mCapacity.LimitSearchAppendCapacityDiagnostic "UltimateStrain: lambda=0 уже за физическим пределом; итерации предельной плоскости не запускаются."
                Exit Function
            End If
        End If
        If mCapacity.LimitSearchHasTerminalFailure Then Exit Function
        ' Численная неудача начальной пробы не доказывает физический отказ.
        ' Сохраняем прежний Newton и его старты, без переноса пробного состояния.
        If Not mCapacity.LimitSearchPrepareCapacityUltimate( _
                request.Section, request.ConcreteMaterial, request.SteelMaterial, loadPath) Then Exit Function
    End If
    ILimitSearchProblem_PrepareUltimate = True
End Function

' Повторяет общий метод Ньютона с новым начальным приближением.
' Нагрузки, физический критерий и численные настройки между попытками сохраняются.
Private Function ILimitSearchProblem_RunUltimateAttempts(ByVal request As CLimitSearchRequest) As Boolean
    Dim search As CUltimateStrainSearch
    Set search = New CUltimateStrainSearch
    Dim loadPath As CLoadPathVector, alternate As CSectionSolver
    Set loadPath = ILimitSearchProblem_BuildLoadPath(request)
    ILimitSearchProblem_RunUltimateAttempts = search.RunNewton(Me, request.Section, _
        request.ConcreteMaterial, request.SteelMaterial, loadPath, Nothing, _
        "Поиск по предельной деформации (UltimateStrain) не завершён: итерации Newton не нашли равновесие.")
    If ILimitSearchProblem_RunUltimateAttempts Or mCapacity.LimitSearchHasTerminalFailure Then Exit Function
    Set alternate = mCapacity.LimitSearchAlternativeUltimateStart(request.Section, _
        request.ConcreteMaterial, request.SteelMaterial, loadPath)
    If alternate Is Nothing Then Exit Function
    ILimitSearchProblem_RunUltimateAttempts = search.RunNewton(Me, request.Section, _
        request.ConcreteMaterial, request.SteelMaterial, loadPath, alternate, _
        "UltimateStrain: повторный старт не дал подтвержденной предельной плоскости.")
End Function

' После принятого решения UltimateStrain сохраняет инженерные величины
' на исходной траектории без повторного выполнения численного поиска.
Private Sub ILimitSearchProblem_CompleteUltimate(ByVal loadPath As CLoadPathVector)
    mCapacity.LimitSearchCompleteCapacityUltimate loadPath
End Sub

' Передаёт критерии и настройки несущей способности общему методу Ньютона.
' Систему уравнений и подбор шага этот класс самостоятельно не решает.
Private Sub ILimitSearchProblem_LimitSearchInitializeUltimateGuess(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double)
    mCapacity.LimitSearchInitializeUltimateGuess section, concreteMaterial, steelMaterial, _
        loadPath, startSolver, eps0, kx, ky
End Sub

' Получает от расчёта две невязки траектории нагрузки и одну невязку
' предельной деформации. Передаёт общему методу Ньютона то же оценённое НДС.
Private Function ILimitSearchProblem_LimitSearchEvaluateUltimateResidual(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal eps0 As Double, _
        ByVal kx As Double, ByVal ky As Double, ByRef solver As CSectionSolver, _
        ByRef rPath1 As Double, ByRef rPath2 As Double, ByRef rLimit As Double, _
        ByRef lambdaValue As Double, ByRef state As String, Optional ByVal centeredPlane As Boolean = False) As Boolean
    ILimitSearchProblem_LimitSearchEvaluateUltimateResidual = mCapacity.LimitSearchEvaluateUltimateResidual( _
        section, concreteMaterial, steelMaterial, loadPath, eps0, kx, ky, solver, _
        rPath1, rPath2, rLimit, lambdaValue, state, centeredPlane)
End Function

' Проверяет три невязки по исходным допускам расчёта несущей способности.
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

' Владелец численной системы выбирает ее масштабы, не меняя допуски приемки.
Private Property Get ILimitSearchProblem_LimitSearchUseJacobianMeritScaling() As Boolean
    ILimitSearchProblem_LimitSearchUseJacobianMeritScaling = mCapacity.LimitSearchJacobianMeritForPath
End Property

' Передаёт найденное поле деформаций на окончательную физическую проверку.
' Вспомогательное равновесие на продолжении диаграмм не подтверждает предел материалов.
Private Function ILimitSearchProblem_LimitSearchFinalizeUltimateResult(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal solver As CSectionSolver, ByVal state As String, ByVal iterations As Long) As EUltimateFinalization
    If mCapacity.LimitSearchFinalizeUltimateResult(section, concreteMaterial, steelMaterial, _
            loadPath, lambdaValue, solver, state, iterations) Then
        ILimitSearchProblem_LimitSearchFinalizeUltimateResult = ufAccepted
    ElseIf lambdaValue <= 0# Then
        ILimitSearchProblem_LimitSearchFinalizeUltimateResult = ufRejectedPoint
    Else
        ILimitSearchProblem_LimitSearchFinalizeUltimateResult = ufNeedsRefinement
    End If
End Function

' Сохраняет причину численной остановки UltimateStrain в расчёте несущей способности.
Private Sub ILimitSearchProblem_LimitSearchSetUltimateNumericalFailure(ByVal reason As String)
    mCapacity.LimitSearchSetUltimateNumericalFailure reason
End Sub

' Передаёт код причины общего алгоритма в результат несущей способности.
Private Sub ILimitSearchProblem_SetFailure(ByVal failureCode As ESolverFailureCode, ByVal reason As String)
    mCapacity.LimitSearchSetFailure failureCode, reason
End Sub

' Возвращает численные настройки общего метода Ньютона для задачи несущей способности.
Private Property Get ILimitSearchProblem_LimitSearchUltimateMaxIterations() As Long
    ILimitSearchProblem_LimitSearchUltimateMaxIterations = mCapacity.LimitSearchUltimateMaxIterations
End Property

' Сообщает координатору об окончательном успехе или ошибке, запрещающей продолжение.
' Отсутствие объекта задачи не выдаётся за успешное завершение.
Private Property Get ILimitSearchProblem_StoppedTerminally() As Boolean
    If Not mCapacity Is Nothing Then _
        ILimitSearchProblem_StoppedTerminally = mCapacity.Converged Or mCapacity.LimitSearchHasTerminalFailure
End Property

' Возвращает сохранённое пояснение причины остановки поиска.
Private Property Get ILimitSearchProblem_StopReason() As String
    If Not mCapacity Is Nothing Then ILimitSearchProblem_StopReason = mCapacity.StopReason
End Property

' При подготовке своей попытки расчёт уже очищает рабочий итог.
' Дополнительный сброс координатора не должен стереть сохранённый код причины.
Private Sub ILimitSearchProblem_ClearStopReason()
End Sub

' Сохраняет собственную причину расчёта несущей способности
' вместо замены общим текстом дополнительного способа поиска.
Private Sub ILimitSearchProblem_SetStopReasonIfEmpty(ByVal reason As String)
End Sub

' Объединяет диагностику первой попытки и перехода к дополнительному способу поиска.
' Окончательное пояснение определяет расчёт несущей способности, а не запись на лист.
Private Sub ILimitSearchProblem_MergeFallback(ByVal firstLimitState As String, ByVal firstStopReason As String, _
        ByVal firstDiagnostics As String, ByVal firstAttemptMessage As String, _
        ByVal transitionMessage As String, ByVal successMessage As String, _
        Optional ByVal firstSearchExecuted As Boolean = False)
    mCapacity.LimitSearchMergeCapacityFallback firstLimitState, firstStopReason, _
        firstDiagnostics, firstAttemptMessage, transitionMessage, successMessage, firstSearchExecuted
End Sub

' Форматирует число для диагностического журнала с точкой как разделителем.
Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function
