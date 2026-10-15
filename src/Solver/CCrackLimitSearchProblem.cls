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

' ==========================================================================
' Связь расчёта образования трещины с общим поиском предела
' ==========================================================================
' Реализует ILimitSearchProblem через CCrackFormationCalculator.
' Передаёт оценку пробной нагрузки и приём результата; критерий бетона,
' состояния до и после трещины и диагностику хранит расчёт образования трещины.

Implements ILimitSearchProblem

Private Const STATE_ACCEPTABLE As String = "Acceptable" ' Трещина еще не образована.
Private Const STATE_CRACK_LIMIT As String = "CrackFormationLimit" ' Достигнут критерий eps_bt,ult.
Private Const STATE_OTHER_LIMIT As String = "OtherPhysicalLimit" ' Другой физический предел достигнут до трещины.
Private Const STATE_NUMERICAL_FAILURE As String = "NumericalFailure" ' НДС для пробной lambda не найдено.

Private mCrack As CCrackFormationCalculator ' Расчёт образования трещины, который оценивает критерий бетона и сохраняет диагностику.
Private mLastAcceptable As CSectionSolver ' Подтвержденный НДС текущей поисковой сессии.
Private mEquilibriumProbeFailed As Boolean ' Признак неудачи пробного решения равновесия при поиске образования трещины.

' Готовит критерий и рабочие данные образования трещины для выбранного пути.
' Постоянная и масштабируемая части нагрузки задаются тем же способом, что для прочности.
Private Sub ILimitSearchProblem_BeginLoadPath(ByVal request As CLimitSearchRequest)
    mCrack.LimitSearchBeginCrackFormationAttempt request.SelectedLoadPath.Key, request.SelectedLoadPath
End Sub

' Передаёт расчёту образования трещины выбранный путь и причины прежних попыток
' для сохранения в окончательном результате.
Private Sub ILimitSearchProblem_FinishLoadPath(ByVal request As CLimitSearchRequest, _
        ByVal previousPathComments As String)
    mCrack.LimitSearchFinishLoadPath request.SelectedLoadPath.Key, previousPathComments
End Sub

' Подключает существующий расчетчик трещин без изменения его настроек.
Public Sub Initialize(ByVal crackCalculator As CCrackFormationCalculator)
    If crackCalculator Is Nothing Then Err.Raise vbObjectError + 4250, "CCrackLimitSearchProblem", _
        "Не передан расчетчик трещинообразования."
    Set mCrack = crackCalculator
End Sub

' Тип результата, публикуемого данной задачей предельного поиска.
Private Property Get ILimitSearchProblem_ResultKind() As EResultKind
    ILimitSearchProblem_ResultKind = rkCrackFormation
End Property

' Возвращает начальный шаг множителя нагрузки lambda для поиска предела.
Private Property Get ILimitSearchProblem_InitialLambdaStep() As Double
    ILimitSearchProblem_InitialLambdaStep = 1#
End Property

' Возвращает заданную техническую верхнюю границу множителя lambda.
Private Property Get ILimitSearchProblem_MaxLambda() As Double
    ILimitSearchProblem_MaxLambda = mCrack.LimitSearchCrackBracketLimit
End Property

' Возвращает допуск уточнения предельного множителя нагрузки lambda.
Private Property Get ILimitSearchProblem_LambdaTolerance() As Double
    ILimitSearchProblem_LambdaTolerance = mCrack.LimitSearchCrackLambdaTolerance
End Property

' Возвращает максимальное число итераций общего поиска предела.
Private Property Get ILimitSearchProblem_MaxIterations() As Long
    ILimitSearchProblem_MaxIterations = mCrack.LimitSearchCrackMaxIterations
End Property

' Сохраняет заданный предел уточнения скобки трещинообразования.
Private Property Get ILimitSearchProblem_BisectionMaxIterations() As Long
    ILimitSearchProblem_BisectionMaxIterations = mCrack.LimitSearchCrackMaxIterations
End Property

' Возвращает выбранный численный метод поиска предельной нагрузки.
Private Property Get ILimitSearchProblem_SearchMethod() As String
    ILimitSearchProblem_SearchMethod = "Bisection"
End Property

' Показывает, подтверждена ли сходимость поиска предельной точки.
Private Property Get ILimitSearchProblem_Converged() As Boolean
    ILimitSearchProblem_Converged = mCrack.Converged
End Property

' Возвращает сохранённый код физического предела или исхода поиска.
Private Property Get ILimitSearchProblem_LimitState() As String
    ILimitSearchProblem_LimitState = mCrack.StopReason
End Property

' Возвращает накопленный диагностический журнал поиска предела.
Private Property Get ILimitSearchProblem_DiagnosticLog() As String
    If Not mCrack Is Nothing Then ILimitSearchProblem_DiagnosticLog = mCrack.DiagnosticLog
End Property

' Для CrackFormation специальной подготовки перед LoadMultiplier нет:
' материалы и путь уже переданы вызывающим слоем.
Private Function ILimitSearchProblem_Prepare(ByVal request As CLimitSearchRequest) As Boolean
    Set mLastAcceptable = Nothing
    mEquilibriumProbeFailed = False
    mCrack.LimitSearchAppendDiagnostic "LoadMultiplier: начинаем подбор lambda_crc пробными расчетами НДС; путь " & _
        mCrack.LimitSearchCrackFormationPathDisplayName() & "."
    ILimitSearchProblem_Prepare = True
End Function

' Материалы и критерий уже подготовлены расчётом образования трещины.
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

' Расчёт образования трещины сохраняет точку при её подтверждении.
' Отдельная повторная запись инженерных величин после общего цикла не нужна.
Private Sub ILimitSearchProblem_CompleteUltimate(ByVal loadPath As CLoadPathVector)
End Sub

' Передаёт критерий образования трещины и настройки общему методу Ньютона.
' Вычисление Якобиана и подбор шага выполняет общий численный класс.
Private Sub ILimitSearchProblem_LimitSearchInitializeUltimateGuess(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double)
    mCrack.LimitSearchInitializeUltimateGuess section, concreteMaterial, steelMaterial, _
        loadPath, startSolver, eps0, kx, ky
End Sub

' Получает невязки траектории нагрузки и критерия растяжения бетона
' от расчёта образования трещины. Передаёт общему методу Ньютона то же НДС.
Private Function ILimitSearchProblem_LimitSearchEvaluateUltimateResidual(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal eps0 As Double, _
        ByVal kx As Double, ByVal ky As Double, ByRef solver As CSectionSolver, _
        ByRef rPath1 As Double, ByRef rPath2 As Double, ByRef rLimit As Double, _
        ByRef lambdaValue As Double, ByRef state As String, Optional ByVal centeredPlane As Boolean = False) As Boolean
    ILimitSearchProblem_LimitSearchEvaluateUltimateResidual = mCrack.LimitSearchEvaluateUltimateResidual( _
        section, concreteMaterial, steelMaterial, loadPath, eps0, kx, ky, solver, _
        rPath1, rPath2, rLimit, lambdaValue, state, centeredPlane)
End Function

' Проверяет три невязки по действующим допускам расчёта образования трещины.
Private Function ILimitSearchProblem_LimitSearchUltimateResidualReached(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Boolean
    ILimitSearchProblem_LimitSearchUltimateResidualReached = _
        mCrack.LimitSearchUltimateResidualReached(rPath1, rPath2, rLimit)
End Function

' Возвращает норму для выбора численного шага; она не заменяет независимое
' подтверждение конечного физического состояния и растягивающего критерия.
Private Function ILimitSearchProblem_LimitSearchUltimateResidualNorm(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Double
    ILimitSearchProblem_LimitSearchUltimateResidualNorm = _
        mCrack.LimitSearchUltimateResidualNorm(rPath1, rPath2, rLimit)
End Function

' Для образования трещины размерные и деформационные невязки уравновешивает Якобиан.
Private Property Get ILimitSearchProblem_LimitSearchUseJacobianMeritScaling() As Boolean
    ILimitSearchProblem_LimitSearchUseJacobianMeritScaling = True
End Property

' Передаёт найденную точку на окончательную проверку и сохранение.
' Решение на продолжении диаграмм не принимается автоматически
' как физически допустимое состояние перед трещиной.
Private Function ILimitSearchProblem_LimitSearchFinalizeUltimateResult(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal solver As CSectionSolver, ByVal state As String, ByVal iterations As Long) As EUltimateFinalization
    If mCrack.LimitSearchFinalizeUltimateResult(section, concreteMaterial, steelMaterial, _
            loadPath, lambdaValue, solver, state, iterations) Then
        ILimitSearchProblem_LimitSearchFinalizeUltimateResult = ufAccepted
    ElseIf lambdaValue < 0# Then
        ILimitSearchProblem_LimitSearchFinalizeUltimateResult = ufRejectedPoint
    Else
        ILimitSearchProblem_LimitSearchFinalizeUltimateResult = ufNeedsRefinement
    End If
End Function

' Сохраняет причину численной остановки UltimateStrain в проверке образования трещины.
Private Sub ILimitSearchProblem_LimitSearchSetUltimateNumericalFailure(ByVal reason As String)
    mCrack.LimitSearchSetUltimateNumericalFailure reason
End Sub

' Передаёт фактический код причины общего поиска в результат образования трещины.
Private Sub ILimitSearchProblem_SetFailure(ByVal failureCode As ESolverFailureCode, ByVal reason As String)
    mCrack.LimitSearchSetFailure failureCode, reason
End Sub

' Возвращает численные настройки общего метода Ньютона для образования трещины.
Private Property Get ILimitSearchProblem_LimitSearchUltimateMaxIterations() As Long
    ILimitSearchProblem_LimitSearchUltimateMaxIterations = mCrack.LimitSearchUltimateMaxIterations
End Property

' Возвращает выбранное изменение нагрузки LambdaMx/My/Mxy/N/NMxy.
' Если отдельный объект траектории не подготовлен, использует проверенные
' исходные данные расчёта образования трещины.
Private Function ILimitSearchProblem_BuildLoadPath(ByVal request As CLimitSearchRequest) As CLoadPathVector
    If request.HasLoadPathSelection Then
        Set ILimitSearchProblem_BuildLoadPath = request.SelectedLoadPath.Vector
        Exit Function
    End If
    Set ILimitSearchProblem_BuildLoadPath = mCrack.LimitSearchCrackFormationLoadPath( _
        request.TargetN, request.MomentMxAtReference, request.MomentMyAtReference, _
        request.ReferenceX, request.ReferenceY)
End Function

' Передаёт начальную нагрузку на физическую проверку образования трещины.
' Сам общий поиск не определяет наличие трещины в постоянной нагрузке
' или возможность принять psi_s = 1.
Private Function ILimitSearchProblem_AcceptInitialPoint(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal functionValue As Double, _
        ByVal probeState As String, ByVal solver As CSectionSolver) As Boolean
    ILimitSearchProblem_AcceptInitialPoint = mCrack.LimitSearchAcceptFormationInitialPoint( _
        request.TargetN, request.MomentMxAtReference, request.MomentMyAtReference, functionValue)
End Function

' Оценивает образование трещины при одном lambda и возвращает найденное НДС.
' Оно используется как начальное приближение следующей пробы
' и для сохранения состояния перед трещиной.
Private Function ILimitSearchProblem_EvaluatePoint(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByVal lambdaValue As Double, ByRef functionValue As Double, _
        ByRef probeState As String, ByRef solver As CSectionSolver) As Boolean
    ILimitSearchProblem_EvaluatePoint = mCrack.LimitSearchEvaluateCrackFormationLoadMultiplier( _
        request.Section, request.ConcreteMaterial, request.SteelMaterial, _
        loadPath, startSolver, lambdaValue, functionValue, probeState, solver)
    If ILimitSearchProblem_EvaluatePoint Then
        If probeState = STATE_ACCEPTABLE And Not solver Is Nothing Then
            If solver.Converged Then Set mLastAcceptable = solver
        End If
    ElseIf probeState = STATE_NUMERICAL_FAILURE Then
        mEquilibriumProbeFailed = True
    End If
End Function

' Различает допустимую пробу, достижение критерия трещины и численную неудачу.
' Ошибка ввода передаётся отдельным кодом причины и не определяется
' по произвольной строке результата пробы.
Private Function ILimitSearchProblem_IsAcceptableState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsAcceptableState = (probeState = STATE_ACCEPTABLE)
End Function

' Проверяет физический предельный статус задачи; численная несходимость сама по себе не означает
' разрушение.
Private Function ILimitSearchProblem_IsLimitState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsLimitState = (probeState = STATE_CRACK_LIMIT Or probeState = STATE_OTHER_LIMIT)
End Function

' Распознаёт численную неудачу отдельно от физического предела.
Private Function ILimitSearchProblem_IsNumericalFailureState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsNumericalFailureState = (probeState = STATE_NUMERICAL_FAILURE)
End Function

' Распознаёт ошибку исходных данных, при которой смена поискового пути недопустима.
Private Function ILimitSearchProblem_IsInvalidInputState(ByVal probeState As String) As Boolean
    ILimitSearchProblem_IsInvalidInputState = False
End Function

' После численной неудачи предлагает середину рабочего неотрицательного интервала lambda.
' Общий поиск отдельно проверяет представимость числа
' и реальное изменение следующей пробной нагрузки.
Private Function ILimitSearchProblem_ReduceLambdaAfterNumericalFailure(ByVal lowLambda As Double, _
        ByVal highLambda As Double, ByVal probeState As String) As Double
    ILimitSearchProblem_ReduceLambdaAfterNumericalFailure = lowLambda + (highLambda - lowLambda) / 2#
End Function

' Допустимые пробы сохраняются как рабочие данные текущего поиска.
' Отдельное конечное НДС для каждой нижней нагрузки не создаётся.
Private Sub ILimitSearchProblem_RecordAcceptable(ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double)
End Sub

' Записывает причину неудачи начальной пробы, не объявляя ее отсутствием трещин.
Private Sub ILimitSearchProblem_HandleInitialFailure(ByVal probeState As String)
    mCrack.LimitSearchSetCrackFormationStopReason _
        "Не удалось найти равновесие в начальной точке пути трещинообразования."
End Sub

' Сохраняет достижение технической верхней границы поиска. Это SEARCH_BOUND_REACHED,
' а не доказанная недостижимость критерия образования трещины.
Private Sub ILimitSearchProblem_HandleBracketNotFound(ByVal lowLambda As Double)
    mCrack.LimitSearchSetCrackFormationSearchBoundReached _
        "Поиск трещинообразования достиг технической границы lambda = " & _
        FormatNumberInvariant(mCrack.LimitSearchCrackBracketLimit) & _
        ", но условие образования нормальной трещины не выполнено. Это не доказывает отсутствие трещины."
End Sub

' Сохраняет, при каком lambda не подтверждено равновесие при поиске границ интервала.
' Состояния до и после трещины не подставляются. Возможность расчёта ширины
' отдельно определяется пригодностью текущего НДС.
Private Sub ILimitSearchProblem_HandleBracketNumericalFailure(ByVal lowLambda As Double, _
        ByVal failedLambda As Double, ByVal probeState As String)
    mCrack.LimitSearchSetCrackFormationStopReason _
        "Не удалось найти равновесие при lambda = " & _
        FormatNumberInvariant(failedLambda) & " в поиске трещинообразования."
End Sub

' Записывает незавершенное уточнение скобки с фактической шириной и допуском.
' Исчерпание бюджета не признается достижением требуемой точности.
Private Sub ILimitSearchProblem_HandleIterationLimit(ByVal maxIterations As Long, _
        ByVal remainingWidth As Double, ByVal tolerance As Double)
    mCrack.LimitSearchSetCrackFormationStopReason _
        "Поиск Bisection достиг лимита итераций до заданной точности: ширина скобки " & _
        FormatNumberInvariant(remainingWidth) & ", допуск " & FormatNumberInvariant(tolerance) & "."
End Sub

' Останавливает проверку при неизвестном исходе пробы и сохраняет его имя
' для диагностики без догадки о физическом состоянии сечения.
Private Sub ILimitSearchProblem_HandleUnexpectedState(ByVal probeState As String)
    mCrack.LimitSearchSetFailure sfcInternalError, _
        "Поиск образования трещины остановлен из-за неизвестного результата контрольной пробы: " & probeState & "."
End Sub

' Передаёт ошибку ввода общего поиска кодом InvalidInput.
' Текст поясняет причину, но не определяет статус.
Private Sub ILimitSearchProblem_SetInvalidInput(ByVal message As String)
    mCrack.LimitSearchSetFailure sfcInvalidInput, message
End Sub

' Передаёт причину численной остановки в расчёт образования трещины.
' Обозначение внешнего статуса формируется отдельно.
Private Sub ILimitSearchProblem_SetNumericalFailure(ByVal message As String)
    mCrack.LimitSearchSetCrackFormationStopReason message
End Sub

' Добавляет техническую строку текущей попытки, не меняя ее инженерный исход.
Private Sub ILimitSearchProblem_AppendDiagnostic(ByVal message As String)
    mCrack.LimitSearchAppendDiagnostic message
End Sub

' Подтверждает выбранную нагрузку lambda по её НДС и материалам.
' Окончательное состояние сохраняет расчёт образования трещины.
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

' Передаёт верхнюю нагрузку уточнённого интервала для проверки критерия растяжения.
' Перед сохранением физическая пригодность её НДС подтверждается повторно.
Private Function ILimitSearchProblem_FinalizeBracket(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lowLambda As Double, ByVal lowSolver As CSectionSolver, _
        ByVal highLambda As Double, ByVal highSolver As CSectionSolver, _
        ByVal methodName As String, ByVal iterations As Long) As Boolean
    ILimitSearchProblem_FinalizeBracket = ILimitSearchProblem_FinalizeAt(request, loadPath, _
        highLambda, STATE_CRACK_LIMIT, methodName, highSolver, iterations)
End Function

' На плато последняя допустимая силовая проба служит стартом общего Newton.
' Принимается только реальный первый предел на исходном пути с равновесием;
' техническая граница поиска и ошибки конфигурации сюда не направляются.
Private Sub ILimitSearchProblem_TryFinalizeAfterSearchFailure(ByVal request As CLimitSearchRequest)
    If Not mEquilibriumProbeFailed Or mLastAcceptable Is Nothing Then Exit Sub
    If mCrack.Converged Then Exit Sub
    If mCrack.LimitSearchStoppedTerminally() Then Exit Sub
    Dim search As CUltimateStrainSearch
    Set search = New CUltimateStrainSearch
    mCrack.LimitSearchAppendDiagnostic "LoadMultiplier: уточняем первый физический предел от последнего подтвержденного НДС."
    search.RunNewton Me, request.Section, request.ConcreteMaterial, request.SteelMaterial, _
        ILimitSearchProblem_BuildLoadPath(request), mLastAcceptable, _
        "Не удалось подтвердить первый физический предел из последней допустимой пробы."
End Sub

' Сохраняет фактический результат образования трещины, его код причины
' и доступность пороговой точки. Одного завершения общего алгоритма недостаточно.
Private Function ILimitSearchProblem_BuildResult(ByVal request As CLimitSearchRequest, _
        ByVal searchSucceeded As Boolean) As CLimitSearchResult
    Set ILimitSearchProblem_BuildResult = mCrack.BuildSearchSnapshot(request.SolutionStrategy)
End Function

' Сообщает координатору об окончательном исходе текущей попытки образования трещины.
Private Property Get ILimitSearchProblem_StoppedTerminally() As Boolean
    ILimitSearchProblem_StoppedTerminally = mCrack.LimitSearchStoppedTerminally()
End Property

' Возвращает сохранённое пояснение причины остановки поиска.
Private Property Get ILimitSearchProblem_StopReason() As String
    ILimitSearchProblem_StopReason = mCrack.StopReason
End Property

' Очищает рабочую причину перед следующей попыткой поиска образования трещины.
Private Sub ILimitSearchProblem_ClearStopReason()
    mCrack.LimitSearchClearStopReason
End Sub

' Добавляет общее пояснение остановки, только если своей подробной причины ещё нет.
Private Sub ILimitSearchProblem_SetStopReasonIfEmpty(ByVal reason As String)
    mCrack.LimitSearchSetStopReasonIfEmpty reason
End Sub

' Сведения о дополнительных стартах и путях уже объединяет сам расчёт образования трещины.
' Этот вызов общего координатора не добавляет те же сообщения повторно.
Private Sub ILimitSearchProblem_MergeFallback(ByVal firstLimitState As String, ByVal firstStopReason As String, _
        ByVal firstDiagnostics As String, ByVal firstAttemptMessage As String, _
        ByVal transitionMessage As String, ByVal successMessage As String, _
        Optional ByVal firstSearchExecuted As Boolean = False)
End Sub

' Форматирует lambda для журнала и инженерного комментария. У целого числа
' убирает пустой десятичный разделитель перед последующей пунктуацией.
Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
    If Right$(FormatNumberInvariant, 1) = "." Then _
        FormatNumberInvariant = Left$(FormatNumberInvariant, Len(FormatNumberInvariant) - 1)
End Function
