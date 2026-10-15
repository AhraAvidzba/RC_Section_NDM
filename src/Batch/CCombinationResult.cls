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
' Общий результат одного сочетания нагрузок
' ==========================================================================
' Объединяет результаты прочности, трещин и устойчивости и хранит полученные
' расчётные состояния. Численные данные остаются в соответствующих объектах;
' здесь собираются итоговый статус сочетания и пояснения к нему.
' Класс получает готовые результаты и передаёт их средствам записи на лист Results.

Private mFrozen As Boolean ' После передачи готового результата его заполнение и очистка запрещены; сохранённые числа доступны для чтения.
Private mProfileId As String ' Снимок выбранного профиля одной строки сочетаний.
Private mDiagramExtensionEnabled As Boolean ' Разрешение численного продолжения диаграмм на момент запуска этого сочетания.

Private mStateRepository As CStateRepository ' Конечные расчётные состояния этого сочетания нагрузок.
Private mResultPolicy As CResultStatusPolicy ' Общие правила преобразования внутреннего статуса в обозначение для пользователя.
Private mStrengthResult As CStrengthResult ' Результаты НДС при заданной нагрузке и предельной несущей способности.
Private mCrackResult As CCrackResult ' Результаты образования и раскрытия трещин, текущего НДС и проверки продольных трещин.
Private mStabilityResult As CStabilityResult ' Проверка устойчивости.
Private mWorkflowMeta As CResultMeta ' Ошибка общей проверки исходных данных или остановка выполнения сочетания до получения инженерного результата.

' Создаёт пустые результаты отдельных проверок. Несуществующее НДС при этом не подставляется.
Private Sub Class_Initialize()
    Clear vbNullString
End Sub

' Очищает результаты всех проверок сочетания; исходные нагрузки и геометрия сохраняются.
Public Sub Clear(ByVal defaultCapacityPath As String)
    AssertWritable
    Set mStateRepository = New CStateRepository
    Set mStrengthResult = New CStrengthResult
    mStrengthResult.Clear defaultCapacityPath
    Set mCrackResult = New CCrackResult
    Set mStabilityResult = New CStabilityResult
    Set mWorkflowMeta = New CResultMeta
    mWorkflowMeta.SetNotRequested rkGeneric
    mProfileId = vbNullString
    mDiagramExtensionEnabled = False
End Sub

' Возвращает внешний статус сочетания по сохранённым статусам отдельных проверок.
Public Property Get Status() As String
    Status = Policy.ExternalStatus(OverallMeta)
End Property

' Возвращает общий статус образования и раскрытия нормальных трещин.
' Продольные трещины проверяются отдельно и в этот статус не входят.
Public Property Get NormalCrackStatus() As String
    NormalCrackStatus = Policy.ExternalStatus(CrackMeta)
End Property

' Возвращает общий статус проверок трещин, включая продольные.
Public Property Get CrackSummaryStatus() As String
    CrackSummaryStatus = Policy.ExternalStatus(CrackSummaryMeta)
End Property

' Проверяет, использовано ли численное продолжение диаграмм в сохранённых конечных состояниях.
' Промежуточные пробы поиска и последняя попытка решателя на этот признак не влияют.
Public Property Get ExtensionUsed() As Boolean
    Dim i As Long
    For i = 1 To mStateRepository.StateCount
        If mStateRepository.StateAt(i).ExtensionUsed Then
            ExtensionUsed = True
            Exit Property
        End If
    Next i
End Property

' Возвращает наибольший коэффициент использования среди выполненных и применимых проверок.
Public Property Get GoverningUtil() As Double
    GoverningUtil = mStrengthResult.Capacity.UtilCapacity
    If mCrackResult.Width.Utilization > GoverningUtil Then GoverningUtil = mCrackResult.Width.Utilization
    If mCrackResult.Longitudinal.Utilization > GoverningUtil Then GoverningUtil = mCrackResult.Longitudinal.Utilization
    If mStabilityResult.Utilization > GoverningUtil Then GoverningUtil = mStabilityResult.Utilization
End Property

' Собирает итоговый статус сочетания и пояснения к проверкам с проблемами.
' Классы записи на лист получают готовый текст из этого результата.
Public Property Get OverallMeta() As CResultMeta
    Dim strength As CResultMeta, crack As CResultMeta, stability As CResultMeta
    Set strength = mStrengthResult.ResultMeta
    Set crack = mCrackResult.ResultMeta
    Set stability = mStabilityResult.Meta
    Set OverallMeta = Policy.WorstResultMeta(rkGeneric, mWorkflowMeta, strength, crack, stability)
    ' Одну общую ошибку ввода не повторяем под именем каждой зависимой проверки.
    ' Статус выбираем через общие правила, сохраняя отдельно возникшую ошибку расчёта CalcErr.
    If mWorkflowMeta.InternalStatus = rsInvalidInput Or mWorkflowMeta.InternalStatus = rsInvalidConfiguration Then
        If OverallMeta.InternalStatus = mWorkflowMeta.InternalStatus And OverallMeta.ResultCode = mWorkflowMeta.ResultCode Then
            OverallMeta.ResultComment = mWorkflowMeta.ResultComment
            Exit Property
        End If
    End If
    Dim comment As String
    mWorkflowMeta.AppendCommentTo comment
    strength.AppendCommentTo comment, True
    crack.AppendCommentTo comment, True
    stability.AppendCommentTo comment
    OverallMeta.ResultComment = comment
End Property

' Возвращает причину общей ошибки ввода или технической остановки сочетания.
' Результаты уже выполненных инженерных проверок сохраняются отдельно.
Public Property Get WorkflowMeta() As CResultMeta
    Set WorkflowMeta = mWorkflowMeta.Clone
End Property

' Сохраняет статус и причину общей остановки сочетания, не разбирая текст пояснения.
' НДС и результаты отдельных инженерных проверок этот метод не заменяет.
Friend Sub SetWorkflowMeta(ByVal meta As CResultMeta)
    AssertWritable
    If meta Is Nothing Then Err.Raise vbObjectError + 3943, "CCombinationResult", "Не передана причина остановки сочетания."
    Set mWorkflowMeta = meta.Clone
End Sub

' Возвращает общий статус и пояснение блока прочности. Сначала учитывает проблемы
' НДС и несущей способности, затем их общий итог без нового расчёта численных величин.
Public Property Get StrengthMeta() As CResultMeta
    Set StrengthMeta = mStrengthResult.ResultMeta
End Property

' Возвращает статус и пояснение расчёта НДС при заданной нагрузке.
Public Property Get DirectStateMeta() As CResultMeta
    Set DirectStateMeta = mStrengthResult.DirectMeta
End Property

' Возвращает статус и пояснение расчёта несущей способности.
Public Property Get CapacityMeta() As CResultMeta
    Set CapacityMeta = mStrengthResult.CapacityMeta
End Property

' Возвращает понятное название фактически принятого изменения нагрузки при поиске предела.
' Сохранённый расчётный ключ PathResolved преобразуется в обозначение lambda*...
' общим CLoadPathDescriptor.
Public Property Get CapacityPathDisplayName() As String
    Dim loadPath As CLoadPathDescriptor
    Set loadPath = New CLoadPathDescriptor
    CapacityPathDisplayName = loadPath.DisplayNameForKey(mStrengthResult.Capacity.PathResolved)
End Property

' Получает из готовых результатов показатель для выбора определяющего сочетания.
' Само сочетание выбирает пакетный расчёт; повторные физические проверки здесь не выполняются.
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

' Возвращает общий статус и пояснение проверок нормальных трещин.
Public Property Get CrackMeta() As CResultMeta
    Set CrackMeta = mCrackResult.NormalMeta
End Property

' Возвращает общий статус и пояснение таблицы нормальных и продольных трещин.
' Класс записи на Results получает готовый итог обеих проверок.
Public Property Get CrackSummaryMeta() As CResultMeta
    Set CrackSummaryMeta = mCrackResult.ResultMeta
End Property

' Возвращает статус и пояснение определения нагрузки образования трещины.
Public Property Get CrackFormationMeta() As CResultMeta
    Set CrackFormationMeta = mCrackResult.Formation.ResultMeta
End Property

' Возвращает статус и пояснение текущего НДС для расчёта трещин.
Public Property Get CrackCurrentStateMeta() As CResultMeta
    Set CrackCurrentStateMeta = mCrackResult.CurrentStateMeta
End Property

' Возвращает статус и пояснение расчёта ширины раскрытия нормальных трещин.
Public Property Get CrackWidthMeta() As CResultMeta
    Set CrackWidthMeta = mCrackResult.Width.ResultMeta
End Property

' Возвращает статус и пояснение проверки продольных трещин.
Public Property Get LongitudinalCrackMeta() As CResultMeta
    Set LongitudinalCrackMeta = mCrackResult.Longitudinal.ResultMeta
End Property

' Возвращает статус и пояснение проверки устойчивости.
Public Property Get StabilityMeta() As CResultMeta
    Set StabilityMeta = mStabilityResult.Meta
End Property

' Возвращает готовые результаты НДС по прочности и несущей способности.
Public Property Get StrengthResult() As CStrengthResult
    Set StrengthResult = mStrengthResult
End Property

' Возвращает готовые результаты образования, раскрытия и продольных трещин.
Public Property Get CrackResult() As CCrackResult
    Set CrackResult = mCrackResult
End Property

' Возвращает сохранённый результат проверки устойчивости.
Public Property Get StabilityResult() As CStabilityResult
    Set StabilityResult = mStabilityResult
End Property


' При первом обращении создаёт общие правила вывода статусов.
' Порядок вызова Clear и Class_Initialize не влияет на их доступность.
Private Function Policy() As CResultStatusPolicy
    If mResultPolicy Is Nothing Then Set mResultPolicy = New CResultStatusPolicy
    Set Policy = mResultPolicy
End Function

' Сохраняет фактический статус и пояснение расчёта НДС при заданной нагрузке.
Public Sub SetDirectStateMeta(ByVal meta As CResultMeta)
    AssertWritable
    mStrengthResult.SetDirectMeta meta
End Sub

' Объединяет несколько причин завершения расчёта НДС в один итог.
Public Sub MergeDirectStateMeta(ByVal meta As CResultMeta)
    AssertWritable
    mStrengthResult.MergeDirectMeta meta
End Sub

' Сохраняет статус устойчивости, когда полного объекта её расчётных результатов нет.
Public Sub SetStabilityMeta(ByVal meta As CResultMeta, Optional ByVal resultComment As String = vbNullString)
    AssertWritable
    mStabilityResult.SetMeta meta, resultComment
End Sub

' Сохраняет готовый результат проверки устойчивости.
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

' Добавляет конечное расчётное состояние сочетания для последующего вывода.
' Промежуточные итерации и пробные нагрузки сюда не записываются. Состояния
' до и после трещины сохраняются, когда найден её множитель нагрузки lambda_crc.
Public Sub AddState(ByVal stateResult As CSectionStateResult)
    AssertWritable
    If stateResult Is Nothing Then Exit Sub
    If mStateRepository Is Nothing Then Set mStateRepository = New CStateRepository
    mStateRepository.StoreState stateResult
End Sub




' Возвращает хранилище расчётных состояний этого сочетания.
' Вызывающий расчёт может добавлять состояния или получать ранее рассчитанные.
Public Property Get StateRepository() As CStateRepository
    If mStateRepository Is Nothing Then Set mStateRepository = New CStateRepository
    Set StateRepository = mStateRepository
End Property

' Сохраняет результат несущей способности и найденное предельное НДС.
Public Sub StoreCapacityResult(ByVal capacity As CCapacityResult)
    AssertWritable
    If capacity Is Nothing Then Err.Raise vbObjectError + 3941, "CCombinationResult", "Не передан результат несущей способности."
    mStrengthResult.StoreCapacityResult capacity
    If Not capacity.StateResult Is Nothing Then AddState capacity.StateResult
End Sub

' Связывает результат проверки при заданной нагрузке с уже сохранённым НДС без копирования его
' чисел.
Public Sub StoreSectionStateResult(ByVal stateResult As CSectionStateResult, _
        Optional ByVal physicalConcrete As CMaterialDiagram = Nothing, _
        Optional ByVal physicalSteel As CMaterialDiagram = Nothing)
    AssertWritable
    mStrengthResult.StoreDirectFromStateResult stateResult, DiagramExtensionEnabled, physicalConcrete, physicalSteel
End Sub

' Сохраняет общий результат проверок трещин; численные данные остаются в результатах отдельных
' проверок.
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

' Возвращает выбранный профиль сочетания. Он задаётся до завершения расчёта;
' служебные методы изменения недоступны внешнему управлению через Automation.
Public Property Get ProfileId() As String
    ProfileId = mProfileId
End Property

' Сохраняет имя выбранного профиля до завершения результата; затем оно доступно только для
' чтения.
Friend Sub SetProfileId(ByVal value As String)
    AssertWritable
    mProfileId = value
End Sub

' Разрешение численного продолжения диаграмм на момент запуска этого сочетания.
Public Property Get DiagramExtensionEnabled() As Boolean
    DiagramExtensionEnabled = mDiagramExtensionEnabled
End Property

' Сохраняет разрешение численного продолжения диаграмм для данного запуска.
' Признаки уже полученных расчётных состояний при этом не меняются.
Friend Sub SetDiagramExtensionEnabled(ByVal value As Boolean)
    AssertWritable
    mDiagramExtensionEnabled = value
End Sub

' Завершает заполнение и запрещает изменение результатов прочности, трещин, устойчивости
' и хранилища НДС. Общие ссылки на состояния сохраняются без копирования;
' повторное завершение уже готового результата безопасно.
Friend Sub Freeze()
    If mFrozen Then Exit Sub
    If Not mStrengthResult Is Nothing Then mStrengthResult.Freeze
    If Not mCrackResult Is Nothing Then mCrackResult.Freeze
    If Not mStabilityResult Is Nothing Then mStabilityResult.Freeze
    If Not mStateRepository Is Nothing Then mStateRepository.Freeze
    mFrozen = True
End Sub

' Запрещает изменение результата после передачи потребителям.
' Для повторного расчёта создаётся новый объект, а прежний результат сохраняется.
Private Sub AssertWritable()
    If mFrozen Then Err.Raise vbObjectError + 4220, "CCombinationResult", _
        "Опубликованный результат нельзя изменить; для нового расчета нужен новый объект."
End Sub
