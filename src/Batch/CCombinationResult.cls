VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCombinationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Канонический результат одного сочетания нагрузок
' ==========================================================================
' Содержит typed subtrees прочности, трещин и устойчивости и repository named
' states. Численные данные принадлежат конечным result-объектам и не копируются
' в плоскую таблицу. Здесь собираются только итоговые meta и комментарии LC;
' класс не решает равновесие, не читает Config и не пишет Excel.

Public ProfileId As String ' Снимок выбранного профиля одной строки сочетаний.
Public DiagramExtensionEnabled As Boolean ' Глобальная настройка на момент запуска LC.

Private mStateRepository As CStateRepository ' Конечные named states текущего LC.
Private mResultPolicy As CResultStatusPolicy ' Единственное преобразование meta в display.
Private mStrengthResult As CStrengthResult ' DirectState и Capacity.
Private mCrackResult As CCrackResult ' Formation, CurrentState, Width, Longitudinal.
Private mStabilityResult As CStabilityResult ' Проверка устойчивости.

' Создает нейтральное дерево результатов без фиктивных физических состояний.
Private Sub Class_Initialize()
    Clear vbNullString
End Sub

' Сбрасывает все ветви LC, не затрагивая исходные нагрузки и геометрию.
Public Sub Clear(ByVal defaultCapacityPath As String)
    Set mStateRepository = New CStateRepository
    Set mStrengthResult = New CStrengthResult
    mStrengthResult.Clear defaultCapacityPath
    Set mCrackResult = New CCrackResult
    Set mStabilityResult = New CStabilityResult
    ProfileId = vbNullString
    DiagramExtensionEnabled = False
End Sub

' Display-статусы являются производными от typed meta и не имеют setter-а.
Public Property Get Status() As String
    Status = Policy.ExternalStatus(OverallMeta)
End Property

Public Property Get NormalCrackStatus() As String
    NormalCrackStatus = Policy.ExternalStatus(CrackMeta)
End Property

Public Property Get CrackSummaryStatus() As String
    CrackSummaryStatus = Policy.ExternalStatus(CrackSummaryMeta)
End Property

' Определяет использование расширения по сохраненным конечным состояниям,
' а не по последнему живому solver-у или промежуточным Search probes.
Public Property Get ExtensionUsed() As Boolean
    Dim i As Long
    For i = 1 To mStateRepository.StateCount
        If mStateRepository.StateAt(i).ExtensionUsed Then
            ExtensionUsed = True
            Exit Property
        End If
    Next i
End Property

' Возвращает наибольшее использование по готовым применимым ветвям.
Public Property Get GoverningUtil() As Double
    GoverningUtil = mStrengthResult.Capacity.UtilCapacity
    If mCrackResult.Width.Utilization > GoverningUtil Then GoverningUtil = mCrackResult.Width.Utilization
    If mCrackResult.Longitudinal.Utilization > GoverningUtil Then GoverningUtil = mCrackResult.Longitudinal.Utilization
    If mStabilityResult.Utilization > GoverningUtil Then GoverningUtil = mStabilityResult.Utilization
End Property

' Возвращает ResultMeta итогового LC, построенный из агрегированного статуса
' и комментариев проблемных дочерних результатов. Writer-ы берут уже готовый
' текст из result-tree и не склеивают пояснения самостоятельно.
Public Property Get OverallMeta() As CResultMeta
    Dim commentText As String
    commentText = BuildOverallResultComment()
    Set OverallMeta = AggregateResultMeta(rkGeneric, commentText, DirectStateMeta, _
        CapacityMeta, CrackMeta, LongitudinalCrackMeta, StabilityMeta)
End Property

' Возвращает сводную ResultMeta блока прочности: сначала ищет
' проблемный direct/capacity-result, затем оставляет агрегированный
' статус без попытки заново интерпретировать численные результаты.
Public Property Get StrengthMeta() As CResultMeta
    Set StrengthMeta = AggregateResultMeta(rkGeneric, BuildStrengthResultComment(), _
        DirectStateMeta, CapacityMeta)
End Property

' Возвращает meta непосредственно из канонического result-subtree.
Public Property Get DirectStateMeta() As CResultMeta
    Set DirectStateMeta = mStrengthResult.DirectMeta
End Property

Public Property Get CapacityMeta() As CResultMeta
    Set CapacityMeta = mStrengthResult.CapacityMeta
End Property

' Возвращает отображаемое имя фактически принятой capacity-траектории.
' Это только presentation-слой результата LC: расчетный ключ остается в
' mStrengthResult.Capacity.PathResolved, а преобразование в lambda*... делается единообразно
' через CCapacityLoadPath.
Public Property Get CapacityPathDisplayName() As String
    Dim loadPath As CCapacityLoadPath
    Set loadPath = New CCapacityLoadPath
    CapacityPathDisplayName = loadPath.DisplayNameForKey(mStrengthResult.Capacity.PathResolved)
End Property

' Собирает показатель для выбора определяющего LC из готовых результатов.
' Сам выбор строки остается у Batch; физические расчеты и поиск здесь не выполняются.
Public Function ReserveForCriterion(ByVal criterion As String) As Double
    Select Case criterion
        Case "StrengthStrain"
            ReserveForCriterion = mStrengthResult.DirectState.StrainReserve
        Case "StrengthCapacity"
            ReserveForCriterion = mStrengthResult.Capacity.ReserveFactor
        Case "Cracks"
            ReserveForCriterion = PositiveMin(mCrackResult.Width.ReserveFactor, mCrackResult.Longitudinal.ReserveFactor)
        Case "Stability"
            ReserveForCriterion = mStabilityResult.SummaryReserve
    End Select
End Function

' Объединяет только доступные положительные запасы, не выдавая отсутствие за ноль проверки.
Private Function PositiveMin(ByVal firstValue As Double, ByVal secondValue As Double) As Double
    If firstValue > 0# And secondValue > 0# Then
        If firstValue < secondValue Then PositiveMin = firstValue Else PositiveMin = secondValue
    ElseIf firstValue > 0# Then
        PositiveMin = firstValue
    ElseIf secondValue > 0# Then
        PositiveMin = secondValue
    End If
End Function

' Возвращает ResultMeta расчета нормальных трещин.
Public Property Get CrackMeta() As CResultMeta
    Dim formationMeta As CResultMeta
    Set formationMeta = CrackFormationMeta

    Dim currentStateProblem As CResultMeta
    Set currentStateProblem = CrackCurrentProblemMeta()

    If IsProblemStatus(Policy.ExternalStatus(formationMeta)) Then
        Set CrackMeta = AggregateResultMeta(rkCrackFormation, BuildCrackResultComment(), formationMeta)
    ElseIf CrackFormationNotReached(formationMeta) Then
        ' Отсутствие нормальной трещины не скрывает отказ обязательного НДС.
        ' Формула ширины неприменима, но сводный результат сохраняет причину State.
        If currentStateProblem Is Nothing Then
            Set CrackMeta = AggregateResultMeta(rkCrackFormation, BuildCrackFormationResultComment(), formationMeta)
        Else
            Set CrackMeta = AggregateResultMeta(rkCrackFormation, BuildCrackResultComment(), _
                formationMeta, currentStateProblem)
        End If
    ElseIf mCrackResult.Formation.CrackFormed Then
        If currentStateProblem Is Nothing Then
            Set CrackMeta = AggregateResultMeta(rkCrackWidth, BuildCrackResultComment(), CrackWidthMeta)
        Else
            Set CrackMeta = AggregateResultMeta(rkCrackWidth, BuildCrackResultComment(), _
                currentStateProblem, CrackWidthMeta)
        End If
    Else
        If currentStateProblem Is Nothing Then
            Set CrackMeta = AggregateResultMeta(rkCrackWidth, BuildCrackResultComment(), _
                formationMeta, CrackWidthMeta)
        Else
            Set CrackMeta = AggregateResultMeta(rkCrackWidth, BuildCrackResultComment(), _
                formationMeta, currentStateProblem, CrackWidthMeta)
        End If
    End If
End Property

' Возвращает ResultMeta всего detailed-блока трещин, который на Results
' показывает нормальные и продольные трещины одной таблицей. Writer получает
' уже готовый статус и комментарий: он не должен сам агрегировать normal/longitudinal.
Public Property Get CrackSummaryMeta() As CResultMeta
    Set CrackSummaryMeta = AggregateResultMeta(rkGeneric, BuildCrackSummaryResultComment(), _
        CrackMeta, LongitudinalCrackMeta)
End Property

' Возвращает meta соответствующего узла без параллельного кэша статусов.
Public Property Get CrackFormationMeta() As CResultMeta
    Set CrackFormationMeta = mCrackResult.Formation.ResultMeta
End Property

' Возвращает meta соответствующего узла без параллельного кэша статусов.
Public Property Get CrackCurrentStateMeta() As CResultMeta
    Set CrackCurrentStateMeta = mCrackResult.CurrentStateMeta
End Property

' Возвращает meta соответствующего узла без параллельного кэша статусов.
Public Property Get CrackWidthMeta() As CResultMeta
    Set CrackWidthMeta = mCrackResult.Width.ResultMeta
End Property

' Возвращает meta соответствующего узла без параллельного кэша статусов.
Public Property Get LongitudinalCrackMeta() As CResultMeta
    Set LongitudinalCrackMeta = mCrackResult.Longitudinal.ResultMeta
End Property

' Возвращает meta соответствующего узла без параллельного кэша статусов.
Public Property Get StabilityMeta() As CResultMeta
    Set StabilityMeta = mStabilityResult.Meta
End Property

' Возвращает structured-subtree расчетов по модели прочности.
Public Property Get StrengthResult() As CStrengthResult
    Set StrengthResult = mStrengthResult
End Property

' Возвращает structured-subtree расчета трещин.
Public Property Get CrackResult() As CCrackResult
    Set CrackResult = mCrackResult
End Property

' Возвращает structured-subtree проверки устойчивости.
Public Property Get StabilityResult() As CStabilityResult
    Set StabilityResult = mStabilityResult
End Property

' Собирает сводный meta-результат из дочерних typed meta. Побеждает худший
' применимый дочерний статус; комментарий передается уже централизованно
' собранным ResultComment без повторного склеивания во writer-ах.
Private Function AggregateResultMeta(ByVal fallbackKind As EResultKind, ByVal resultComment As String, _
        ParamArray metas() As Variant) As CResultMeta
    Dim selected As CResultMeta
    Dim candidate As CResultMeta
    Dim i As Long
    For i = LBound(metas) To UBound(metas)
        Set candidate = Nothing
        If IsObject(metas(i)) Then Set candidate = metas(i)
        If Not candidate Is Nothing Then
            Set selected = WorseMeta(selected, candidate)
        End If
    Next i

    If selected Is Nothing Then
        Set AggregateResultMeta = NotApplicableMeta(fallbackKind, resultComment)
    Else
        Set AggregateResultMeta = selected.Clone
        AggregateResultMeta.ResultComment = resultComment
    End If
End Function

' Выбирает худший из двух meta-результатов по утвержденному внешнему словарю
' статусов. Здесь не анализируются комментарии или StopReason.
Private Function WorseMeta(ByVal currentMeta As CResultMeta, ByVal candidateMeta As CResultMeta) As CResultMeta
    If candidateMeta Is Nothing Then
        Set WorseMeta = currentMeta
    ElseIf currentMeta Is Nothing Then
        Set WorseMeta = candidateMeta
    ElseIf StatusPriority(Policy.ExternalStatus(candidateMeta)) > _
            StatusPriority(Policy.ExternalStatus(currentMeta)) Then
        Set WorseMeta = candidateMeta
    Else
        Set WorseMeta = currentMeta
    End If
End Function

' Приоритет внешних display-статусов нужен только для сводных узлов Results.
Private Function StatusPriority(ByVal statusText As String) As Long
    StatusPriority = Policy.StatusPriority(statusText)
End Function

' Создает typed N/A для веток, которые не запускались или неприменимы.
Private Function NotApplicableMeta(ByVal resultKind As EResultKind, _
        Optional ByVal resultComment As String = vbNullString) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    meta.SetNotApplicable resultKind, resultComment
    Set NotApplicableMeta = meta
End Function

' Возвращает копию meta только для результата, который должен ухудшить общий статус.
' Нейтральные OK и N/A остаются в своих подробных блоках и не блокируют остальные проверки.
Private Function FirstProblemMeta(ByVal meta As CResultMeta) As CResultMeta
    If meta Is Nothing Then Exit Function

    Dim statusText As String
    statusText = Policy.ExternalStatus(meta)
    Select Case statusText
        Case Policy.Fail, Policy.BaseFail, Policy.NumFail, Policy.InputErr, Policy.CalcErr
            Set FirstProblemMeta = meta.Clone
            EnsureMetaComment FirstProblemMeta
    End Select
End Function

' Собирает пользовательский комментарий для batch summary из всех дочерних
' result-subtree. В подробных таблицах каждый writer продолжает получать только
' свой meta-комментарий, а верхняя сводка показывает полный набор причин LC.
Private Function BuildOverallResultComment() As String
    Dim resultText As String
    AppendMetaComment resultText, DirectStateMeta
    AppendMetaComment resultText, CapacityMeta
    AppendMetaComment resultText, CrackFormationMeta
    AppendMetaComment resultText, CrackCurrentStateMeta
    AppendMetaComment resultText, CrackWidthMeta
    AppendMetaComment resultText, LongitudinalCrackMeta
    AppendMetaComment resultText, StabilityMeta
    BuildOverallResultComment = resultText
End Function

' Собирает комментарий только для detailed-блока прочности. В него входят
' причины прямого НДС по модели прочности и поиска несущей способности; crack
' и stability-комментарии остаются в своих блоках и в общей batch-сводке.
Private Function BuildStrengthResultComment() As String
    Dim resultText As String
    AppendMetaComment resultText, DirectStateMeta
    AppendMetaComment resultText, CapacityMeta
    BuildStrengthResultComment = resultText
End Function

' Собирает комментарий только для detailed-блока трещин: образование трещины,
' равновесное CrackedState и формульная проверка ширины раскрытия.
Private Function BuildCrackResultComment() As String
    Dim resultText As String
    AppendMetaComment resultText, CrackFormationMeta
    AppendMetaComment resultText, CrackCurrentStateMeta
    AppendMetaComment resultText, CrackWidthMeta
    BuildCrackResultComment = resultText
End Function

' Собирает комментарий для всей detailed-таблицы трещин. В отличие от
' CrackMeta сюда входит и проверка продольных трещин, потому что она выводится
' в том же пользовательском блоке Results.
Private Function BuildCrackSummaryResultComment() As String
    Dim resultText As String
    AppendMetaComment resultText, CrackFormationMeta
    AppendMetaComment resultText, CrackCurrentStateMeta
    AppendMetaComment resultText, CrackWidthMeta
    AppendMetaComment resultText, LongitudinalCrackMeta
    BuildCrackSummaryResultComment = resultText
End Function

' Собирает комментарий ветки образования нормальной трещины без примешивания
' CurrentCrackedState. Эта ветка используется, когда search корректно показал,
' что нормальная трещина при текущем сочетании еще не образована.
Private Function BuildCrackFormationResultComment() As String
    Dim resultText As String
    AppendMetaComment resultText, CrackFormationMeta
    BuildCrackFormationResultComment = resultText
End Function

' True, если search образования нормальной трещины корректно показал,
' что при текущем сочетании критерий трещинообразования еще не достигнут.
Private Function CrackFormationNotReached(ByVal meta As CResultMeta) As Boolean
    If meta Is Nothing Then Exit Function
    CrackFormationNotReached = (meta.InternalStatus = rsSuccess And _
        (meta.ResultCode = rcCriterionNotReached Or meta.ResultCode = rcCrackNotFormed))
End Function

' Возвращает проблемный meta текущего CrackedState, если обязательное НДС для
' трещин не получено или найдено только за пределами физической диаграммы.
' Формульная ширина при этом остается отдельным result-subtree, но общий crack-
' статус не должен скрывать ошибку обязательного состояния.
Private Function CrackCurrentProblemMeta() As CResultMeta
    Dim meta As CResultMeta
    Set meta = CrackCurrentStateMeta
    If meta Is Nothing Then Exit Function

    Set CrackCurrentProblemMeta = FirstProblemMeta(meta)
End Function

' Добавляет комментарий для проблемного результата или успешного результата с
' предупреждением. Обычные OK и N/A остаются нейтральными, но fallback-комментарий
' rsSuccessWithWarning нельзя терять только потому, что внешний статус равен OK.
Private Sub AppendMetaComment(ByRef targetText As String, ByVal meta As CResultMeta)
    If meta Is Nothing Then Exit Sub

    If Not ShouldIncludeMetaComment(meta) Then Exit Sub

    Dim commentText As String
    commentText = Trim$(meta.ResultComment)
    If Len(commentText) = 0 Then commentText = DefaultProblemComment(Policy.ExternalStatus(meta))
    If Len(commentText) = 0 Then Exit Sub

    AppendUniqueComment targetText, ResultKindLabel(meta.ResultKind) & ": " & commentText
End Sub

' Решает, должен ли комментарий дочернего result попасть в сводный текст.
' Логика опирается на InternalStatus/ExternalStatus, но не анализирует сам текст
' комментария и не назначает по нему новые статусы.
Private Function ShouldIncludeMetaComment(ByVal meta As CResultMeta) As Boolean
    If meta Is Nothing Then Exit Function
    If IsProblemStatus(Policy.ExternalStatus(meta)) Then
        ShouldIncludeMetaComment = True
    ElseIf meta.InternalStatus = rsSuccessWithWarning Then
        ShouldIncludeMetaComment = True
    End If
End Function

' Добавляет фразу без дублей; batch summary может собрать одну причину из
' нескольких result-subtree одного LC.
Private Sub AppendUniqueComment(ByRef targetText As String, ByVal commentText As String)
    commentText = Trim$(commentText)
    If Len(commentText) = 0 Then Exit Sub
    If InStr(1, targetText, commentText, vbTextCompare) > 0 Then Exit Sub

    If Len(targetText) > 0 Then targetText = targetText & "; "
    targetText = targetText & commentText
End Sub

' Гарантирует, что problem-meta, который попадет в подробный output-блок,
' имеет понятное пояснение даже при неполном результате без комментария.
Private Sub EnsureMetaComment(ByVal meta As CResultMeta)
    If meta Is Nothing Then Exit Sub
    If Len(Trim$(meta.ResultComment)) > 0 Then Exit Sub
    meta.ResultComment = DefaultProblemComment(Policy.ExternalStatus(meta))
End Sub

' Возвращает короткое пользовательское пояснение для типовых проблем, когда
' ответственный расчетный объект не передал более конкретную причину.
Private Function DefaultProblemComment(ByVal statusText As String) As String
    Select Case statusText
        Case Policy.BaseFail
            DefaultProblemComment = "Несущая способность не проходит уже для исходной части выбранного пути."
        Case Policy.Fail
            DefaultProblemComment = "Проверка не проходит."
        Case Policy.NumFail
            DefaultProblemComment = "Численная процедура не сошлась."
        Case Policy.InputErr
            DefaultProblemComment = "Ошибка исходных данных или настроек."
        Case Policy.CalcErr
            DefaultProblemComment = "Внутренняя ошибка расчета или несформированный результат."
    End Select
End Function

' Отделяет проблемные пользовательские статусы от нейтральных OK/N/A.
Private Function IsProblemStatus(ByVal statusText As String) As Boolean
    Select Case statusText
        Case Policy.Fail, Policy.BaseFail, Policy.NumFail, Policy.InputErr, Policy.CalcErr
            IsProblemStatus = True
    End Select
End Function

' Дает короткую русскую метку result-subtree для агрегированного комментария
' batch summary, чтобы пользователь видел, какой раздел сформировал причину.
Private Function ResultKindLabel(ByVal resultKind As EResultKind) As String
    Select Case resultKind
        Case rkDirectState
            ResultKindLabel = "НДС"
        Case rkCapacity
            ResultKindLabel = "Несущая"
        Case rkCrackFormation
            ResultKindLabel = "Трещинообразование"
        Case rkCrackWidth
            ResultKindLabel = "Ширина трещин"
        Case rkLongitudinalCrack
            ResultKindLabel = "Продольные трещины"
        Case rkStability
            ResultKindLabel = "Устойчивость"
        Case Else
            ResultKindLabel = "Расчет"
    End Select
End Function

' Лениво создает policy, чтобы CCombinationResult оставался простым контейнером
' и не зависел от порядка вызова Clear/Class_Initialize.
Private Function Policy() As CResultStatusPolicy
    If mResultPolicy Is Nothing Then Set mResultPolicy = New CResultStatusPolicy
    Set Policy = mResultPolicy
End Function

' Устанавливает фактическую инженерную meta прямого НДС.
Public Sub SetDirectStateMeta(ByVal meta As CResultMeta)
    mStrengthResult.SetDirectMeta meta
End Sub

' Объединяет несколько причин прямого расчета в один итоговый result.
Public Sub MergeDirectStateMeta(ByVal meta As CResultMeta)
    SetDirectStateMeta AggregateResultMeta(rkDirectState, vbNullString, DirectStateMeta, meta)
End Sub

' Устанавливает meta устойчивости, если полный calculator-result отсутствует.
Public Sub SetStabilityMeta(ByVal meta As CResultMeta, Optional ByVal resultComment As String = vbNullString)
    mStabilityResult.SetMeta meta, resultComment
End Sub

' Сохраняет единственный самостоятельный result устойчивости.
Public Sub StoreStabilityResult(ByVal meta As CResultMeta, ByVal stability As CStabilityCalculator, _
        Optional ByVal referenceX As Double = 0#, Optional ByVal referenceY As Double = 0#)
    Set mStabilityResult = New CStabilityResult
    mStabilityResult.InitializeFromCalculator meta, stability, referenceX, referenceY
End Sub

' Обновляет окончательные усилия после общего фильтра малых моментов.
Public Sub UpdateStabilityDesignLoads(ByVal designN As Double, ByVal designMx As Double, _
        ByVal designMy As Double, ByVal principalAngleRad As Double)
    mStabilityResult.UpdateDesignLoads designN, designMx, designMy, principalAngleRad
End Sub

' Сбрасывает незапрошенную ветвь устойчивости.
Public Sub ClearStabilityResult()
    Set mStabilityResult = New CStabilityResult
End Sub

' Добавляет конечное named-state LC. Промежуточные итерации solver-а и
' lambda-точки поиска сюда не пишутся; состояния PreCrackState/PostCrackState сохраняются,
' если обязательная проверка образования трещины реально дошла до lambda_crc.
Public Sub AddState(ByVal stateResult As CSectionStateResult)
    If stateResult Is Nothing Then Exit Sub
    If mStateRepository Is Nothing Then Set mStateRepository = New CStateRepository
    mStateRepository.StoreState stateResult
End Sub




' Возвращает repository текущего LC. Его используют только слои orchestration,
' которые имеют право добавлять или переиспользовать конечные named-state.
Public Property Get StateRepository() As CStateRepository
    If mStateRepository Is Nothing Then Set mStateRepository = New CStateRepository
    Set StateRepository = mStateRepository
End Property

' Сохраняет готовый capacity-result и его конечное физическое состояние.
Public Sub StoreCapacityResult(ByVal capacity As CCapacityResult)
    If capacity Is Nothing Then Err.Raise vbObjectError + 3941, "CCombinationResult", "Не передан результат несущей способности."
    mStrengthResult.StoreCapacityResult capacity
    If Not capacity.StateResult Is Nothing Then AddState capacity.StateResult
End Sub

' Привязывает прямой result к уже сохраненному каноническому named state.
Public Sub StoreSectionStateResult(ByVal stateResult As CSectionStateResult, _
        Optional ByVal physicalConcrete As CMaterialDiagram = Nothing, _
        Optional ByVal physicalSteel As CMaterialDiagram = Nothing)
    mStrengthResult.StoreDirectFromStateResult stateResult, DiagramExtensionEnabled, physicalConcrete, physicalSteel
End Sub

' Сохраняет самостоятельное дерево трещин без копирования численных полей.
Public Sub StoreCrackAggregateResult(ByVal crackResult As CCrackResult)
    If crackResult Is Nothing Then Err.Raise vbObjectError + 3942, "CCombinationResult", "Не передан результат расчета трещин."
    Set mCrackResult = crackResult
End Sub

' Полностью сбрасывает незапрошенную ветвь трещин.
Public Sub ClearCrackResult()
    Set mCrackResult = New CCrackResult
End Sub
