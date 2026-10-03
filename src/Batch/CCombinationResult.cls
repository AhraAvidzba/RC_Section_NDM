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

Private mFrozen As Boolean ' После публикации заполнение/reset запрещены; численные данные не копируются.
Private mProfileId As String ' Снимок выбранного профиля одной строки сочетаний.
Private mDiagramExtensionEnabled As Boolean ' Глобальная настройка на момент запуска LC.

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
    AssertWritable
    Set mStateRepository = New CStateRepository
    Set mStrengthResult = New CStrengthResult
    mStrengthResult.Clear defaultCapacityPath
    Set mCrackResult = New CCrackResult
    Set mStabilityResult = New CStabilityResult
    mProfileId = vbNullString
    mDiagramExtensionEnabled = False
End Sub

' Display-статусы являются производными от typed meta и не имеют setter-а.
Public Property Get Status() As String
    Status = Policy.ExternalStatus(OverallMeta)
End Property

' Возвращает display-итог образования/раскрытия нормальной трещины;
' независимая проверка продольных трещин в этот статус не включается.
Public Property Get NormalCrackStatus() As String
    NormalCrackStatus = Policy.ExternalStatus(CrackMeta)
End Property

' Возвращает display-итог всего crack subtree, включая продольные трещины.
' Источником является готовая агрегированная meta, а не строки соседних ячеек.
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
    Dim strength As CResultMeta, crack As CResultMeta, stability As CResultMeta
    Set strength = mStrengthResult.ResultMeta
    Set crack = mCrackResult.ResultMeta
    Set stability = mStabilityResult.Meta
    Dim comment As String
    strength.AppendCommentTo comment, True
    crack.AppendCommentTo comment, True
    stability.AppendCommentTo comment
    Set OverallMeta = Policy.WorstResultMeta(rkGeneric, strength, crack, stability)
    OverallMeta.ResultComment = comment
End Property

' Возвращает сводную ResultMeta блока прочности: сначала ищет
' проблемный direct/capacity-result, затем оставляет агрегированный
' статус без попытки заново интерпретировать численные результаты.
Public Property Get StrengthMeta() As CResultMeta
    Set StrengthMeta = mStrengthResult.ResultMeta
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
' через CLoadPathDescriptor.
Public Property Get CapacityPathDisplayName() As String
    Dim loadPath As CLoadPathDescriptor
    Set loadPath = New CLoadPathDescriptor
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
    Set CrackMeta = mCrackResult.NormalMeta
End Property

' Возвращает ResultMeta всего detailed-блока трещин, который на Results
' показывает нормальные и продольные трещины одной таблицей. Writer получает
' уже готовый статус и комментарий: он не должен сам агрегировать normal/longitudinal.
Public Property Get CrackSummaryMeta() As CResultMeta
    Set CrackSummaryMeta = mCrackResult.ResultMeta
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


' Лениво создает policy, чтобы CCombinationResult оставался простым контейнером
' и не зависел от порядка вызова Clear/Class_Initialize.
Private Function Policy() As CResultStatusPolicy
    If mResultPolicy Is Nothing Then Set mResultPolicy = New CResultStatusPolicy
    Set Policy = mResultPolicy
End Function

' Устанавливает фактическую инженерную meta прямого НДС.
Public Sub SetDirectStateMeta(ByVal meta As CResultMeta)
    AssertWritable
    mStrengthResult.SetDirectMeta meta
End Sub

' Объединяет несколько причин прямого расчета в один итоговый result.
Public Sub MergeDirectStateMeta(ByVal meta As CResultMeta)
    AssertWritable
    mStrengthResult.MergeDirectMeta meta
End Sub

' Устанавливает meta устойчивости, если полный calculator-result отсутствует.
Public Sub SetStabilityMeta(ByVal meta As CResultMeta, Optional ByVal resultComment As String = vbNullString)
    AssertWritable
    mStabilityResult.SetMeta meta, resultComment
End Sub

' Сохраняет единственный самостоятельный result устойчивости.
Public Sub StoreStabilityResult(ByVal meta As CResultMeta, ByVal stability As CStabilityCalculator, _
        Optional ByVal referenceX As Double = 0#, Optional ByVal referenceY As Double = 0#)
    AssertWritable
    Set mStabilityResult = New CStabilityResult
    mStabilityResult.InitializeFromCalculator meta, stability, referenceX, referenceY
End Sub

' Обновляет окончательные усилия после общего фильтра малых моментов.
Public Sub UpdateStabilityDesignLoads(ByVal designN As Double, ByVal designMx As Double, _
        ByVal designMy As Double, ByVal principalAngleRad As Double)
    AssertWritable
    mStabilityResult.UpdateDesignLoads designN, designMx, designMy, principalAngleRad
End Sub

' Сбрасывает незапрошенную ветвь устойчивости.
Public Sub ClearStabilityResult()
    AssertWritable
    Set mStabilityResult = New CStabilityResult
End Sub

' Добавляет конечное named-state LC. Промежуточные итерации solver-а и
' lambda-точки поиска сюда не пишутся; состояния PreCrackState/PostCrackState сохраняются,
' если обязательная проверка образования трещины реально дошла до lambda_crc.
Public Sub AddState(ByVal stateResult As CSectionStateResult)
    AssertWritable
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
    AssertWritable
    If capacity Is Nothing Then Err.Raise vbObjectError + 3941, "CCombinationResult", "Не передан результат несущей способности."
    mStrengthResult.StoreCapacityResult capacity
    If Not capacity.StateResult Is Nothing Then AddState capacity.StateResult
End Sub

' Привязывает прямой result к уже сохраненному каноническому named state.
Public Sub StoreSectionStateResult(ByVal stateResult As CSectionStateResult, _
        Optional ByVal physicalConcrete As CMaterialDiagram = Nothing, _
        Optional ByVal physicalSteel As CMaterialDiagram = Nothing)
    AssertWritable
    mStrengthResult.StoreDirectFromStateResult stateResult, DiagramExtensionEnabled, physicalConcrete, physicalSteel
End Sub

' Сохраняет самостоятельное дерево трещин без копирования численных полей.
Public Sub StoreCrackAggregateResult(ByVal crackResult As CCrackResult)
    AssertWritable
    If crackResult Is Nothing Then Err.Raise vbObjectError + 3942, "CCombinationResult", "Не передан результат расчета трещин."
    Set mCrackResult = crackResult
End Sub

' Полностью сбрасывает незапрошенную ветвь трещин.
Public Sub ClearCrackResult()
    AssertWritable
    Set mCrackResult = New CCrackResult
End Sub

' Контекст LC читается потребителями, а меняется только внутри workflow до
' публикации. Friend методы недоступны внешнему Automation-интерфейсу.
Public Property Get ProfileId() As String
    ProfileId = mProfileId
End Property

' Запоминает профиль до публикации LC; внешним потребителям доступно только чтение.
Friend Sub SetProfileId(ByVal value As String)
    AssertWritable
    mProfileId = value
End Sub

Public Property Get DiagramExtensionEnabled() As Boolean
    DiagramExtensionEnabled = mDiagramExtensionEnabled
End Property

' Запоминает effective Extension запуска, не меняя признаки уже полученных State.
Friend Sub SetDiagramExtensionEnabled(ByVal value As Boolean)
    AssertWritable
    mDiagramExtensionEnabled = value
End Sub

' Завершает жизненный цикл заполнения перед передачей snapshot потребителю.
' Повторный вызов безопасен; общие ссылки на named-state остаются теми же,
' но их fill/reset API больше не могут изменить опубликованный результат.
Friend Sub Freeze()
    If mFrozen Then Exit Sub
    If Not mStrengthResult Is Nothing Then mStrengthResult.Freeze
    If Not mCrackResult Is Nothing Then mCrackResult.Freeze
    If Not mStabilityResult Is Nothing Then mStabilityResult.Freeze
    If Not mStateRepository Is Nothing Then mStateRepository.Freeze
    mFrozen = True
End Sub

' Явно отклоняет изменение уже выданного результата вместо тихой порчи снимка.
' Для повторного расчета workflow создает новый объект, сохраняя прежние ссылки.
Private Sub AssertWritable()
    If mFrozen Then Err.Raise vbObjectError + 4220, "CCombinationResult", _
        "Опубликованный результат нельзя изменить; для нового расчета нужен новый объект."
End Sub
