VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRebarRuleExecutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Расстановка арматуры по общим правилам рецепта
' ==========================================================================
' Выполняет независимую расстановку, зависимые ряды и проекцию стержней.
' Проверяет зависимости правил, попадание в бетон и совпадения позиций.
' Возвращает готовую CRebarLayout и сохраняет фактические шаги, количество
' стержней и сведения о правилах, по которым они созданы.

Private Const DIAGNOSTIC_RECORD_LIMIT As Long = 128 ' Храним не более 128 подробных сообщений; общее число событий учитываем без сокращения.

' Рабочие данные расстановки: координаты стержней, зависимости рядов, проекции
' и порядок добавления стержней в общий список.
Private mReadOnly As Boolean ' После завершения сборки повторная расстановка этим объектом запрещена.
Private mBuilt As Boolean ' Freeze разрешён только после полного успешного выполнения правил.
Private mDefinition As CSectionDefinition ' Готовое описание сечения, по которому выполнены расстановка и её проверки.
Private mRules() As TRebarRule, mState() As Long, mStart() As Long, mCount() As Long ' Правила, стадия их выполнения, первый индекс и число позиций каждого правила.
Private mPositions() As TRebarPosition, mPositionCount As Long, mCapacity As Long ' Места установки стержней до объединения дублей, их количество и вместимость массива.
Private mTotalNormalOffset() As Double ' Суммарное нормальное смещение зависимого ряда от его независимой или спроецированной линии.
Private mOffsetContours As CRebarOffsetContours ' Общие индивидуальные контуры только семейств с ненулевым пользовательским отступом.
Private mSkippedOffsets As Long ' Дополнительные позиции вне полного смещённого контура.
Private mMergedOffsetNeighbors As Long ' Пересекающиеся дополнительные стержни соседних граней, заменённые первым физическим стержнем.
Private mSkippedCoincidences As Long ' Разрешённые совпадения с обязательным предупреждением; полный счётчик не зависит от лимита подробных записей.
Private mSkippedProjections As Long ' Проекции без пересечения, пропуск которых явно разрешён.
Private mFilteredByArea As Long, mFilteredAreaByRule() As Long ' Геометрические позиции, исключённые общим порогом до поиска физического владельца.
Private mRuleArea() As Double, mRuleBelowMinArea() As Boolean ' Площадь диаметра и готовый признак фильтра, вычисленные один раз для каждого активного правила.
Private mMergedPositions As Long, mConflicts As Long ' Объединённые позиции и отклонённые конфликты.
Private mCreatedPhysical As Long ' Количество действительно созданных стержней; совпавшие позиции не считаются повторно.
Private mSkippedByRule() As Long, mMergedByRule() As Long, mCreatedByRule() As Long ' Число пропущенных позиций, объединённых позиций и созданных стержней по каждому правилу.
Private mDiagnostic As CGeometryDiagnostic, mDiagnostics As Collection ' Последнее событие и ограниченный набор подробных записей.
Private mBarBuckets As Object, mBarBinSize As Double, mBarOverlaps As Long ' Списки близких стержней, размер ячейки поиска в мм и число пар пересекающихся стержней.
Private mPhysicalRule() As Long, mPhysicalCapacity As Long ' Правило, создавшее каждый стержень, и размер массива этих связей.
Private mPhysicalOwner() As Long ' Номер фактического стержня для каждой запрошенной позиции, включая объединённые совпадения.
Private mPhysicalFirstPosition() As Long ' Первая исходная позиция физического стержня для проверки законного общего узла.
Private mPhysicalOffsetRule() As Long ' Правило проверки смещённого ряда, участвующее в физическом стержне; сохраняется после объединения.
Private mDiagnosticRecordsOmitted As Long ' Число событий сверх лимита подробных записей; счётчики остаются полными.

' Все запрошенные позиции, включая пропуски проекции и смещения; это не число физических стержней.
Public Property Get RequestedPositions() As Long: RequestedPositions = mPositionCount + mSkippedProjections + mSkippedOffsets: End Property
' Действительно добавленные физические стержни текущего Build, даже при отказе в середине.
Public Property Get CreatedPositions() As Long: CreatedPositions = mCreatedPhysical: End Property
' Возвращает число позиций без пересечения с целевой линией, пропуск которых разрешён правилом.
Public Property Get SkippedProjections() As Long: SkippedProjections = mSkippedProjections: End Property
' Возвращает число подготовленных позиций, исключённых по Mesh.MinElementArea.
Public Property Get FilteredByArea() As Long: FilteredByArea = mFilteredByArea: End Property
' Возвращает число запрошенных позиций, объединённых с уже добавленными стержнями.
Public Property Get MergedPositions() As Long: MergedPositions = mMergedPositions: End Property
' Возвращает число недопустимых совпадений и пересечений, отклонённых по настройкам правила.
Public Property Get Conflicts() As Long: Conflicts = mConflicts: End Property
' Возвращает число пар стержней, чьи поперечные сечения пересекаются по площади.
' Простое касание стержней в это число не входит.
Public Property Get BarOverlaps() As Long: BarOverlaps = mBarOverlaps: End Property
' Число событий сверх 128 сохранённых подробных образцов; полные счётчики не сокращаются.
Public Property Get DiagnosticRecordsOmitted() As Long: DiagnosticRecordsOmitted = mDiagnosticRecordsOmitted: End Property
' Последняя машинная диагностика; Nothing при отсутствии записанного события.
Public Property Get LastDiagnostic() As CGeometryDiagnostic: Set LastDiagnostic = mDiagnostic: End Property
' Возвращает число сохранённых подробных диагностических записей, от 0 до 128.
Public Property Get DiagnosticCount() As Long
    If Not mDiagnostics Is Nothing Then DiagnosticCount = mDiagnostics.Count
End Property

' Возвращает подробное сообщение по номеру от 1 до DiagnosticCount.
' Хранятся первые 128 сообщений; итоговые счётчики учитывают все события.
Public Function DiagnosticAt(ByVal index As Long) As CGeometryDiagnostic
    Set DiagnosticAt = mDiagnostics(index)
End Function

' Возвращает число подготовленных позиций. Каждая позиция связана с номером
' реального стержня в CRebarLayout; совпадающие позиции могут ссылаться на один стержень.
Public Property Get PositionCount() As Long: PositionCount = mPositionCount: End Property
' ID правила заявленной позиции, включая объединённые позиции.
Public Function PositionRuleID(ByVal index As Long) As String
    PositionRuleID = mRules(mPositions(index).RuleIndex).RuleID
End Function
' ID геометрического пути заявленной позиции.
Public Function PositionPathID(ByVal index As Long) As String: PositionPathID = mPositions(index).PathID: End Function
' Возвращает имя исходной позиции, по которой был добавлен стержень.
Public Function PositionSourceID(ByVal index As Long) As String: PositionSourceID = mPositions(index).SourcePositionID: End Function
' Возвращает номер стержня в итоговом CRebarLayout для выбранной позиции.
' У объединённых совпадающих позиций номер стержня одинаков.
Public Function PositionPhysicalOwner(ByVal index As Long) As Long: PositionPhysicalOwner = mPhysicalOwner(index): End Function

' Составляет пояснение по правилу: запрошенные и созданные стержни, совпадения и пропуски.
' StationStep означает расстояние вдоль линии расстановки, а не прямой отрезок между центрами.
Public Function RuleTrace(ByVal ruleID As String) As String
    Dim index As Long, minStep As Double, maxStep As Double, stepCount As Long, minChord As Double, maxChord As Double
    If mDefinition Is Nothing Then RuleError 0, "Правила еще не выполнялись."
    index = mDefinition.RuleIndex(ruleID)
    RuleStepBounds ruleID, stepCount, minStep, maxStep, minChord, maxChord
    RuleTrace = "Rule=" & ruleID & "; Requested=" & CStr(mCount(index) + mSkippedByRule(index)) & _
        "; Created=" & CStr(mCreatedByRule(index)) & "; Merged=" & CStr(mMergedByRule(index)) & _
        "; Skipped=" & CStr(mSkippedByRule(index) + mFilteredAreaByRule(index)) & _
        "; FilteredArea=" & CStr(mFilteredAreaByRule(index)) & "; StationSteps=" & CStr(stepCount) & _
        "; StationStepMin=" & CStr(minStep) & "; StationStepMax=" & CStr(maxStep) & _
        "; CenterStepMin=" & CStr(minChord) & "; CenterStepMax=" & CStr(maxChord)
End Function

' Определяет наименьший и наибольший шаг стержней, мм. StationStep измеряется
' вдоль линии расстановки, для зависимого ряда — вдоль исходной линии. CenterStep
' равен прямому расстоянию между центрами. Замыкание учитывается отдельно;
' разрывы между независимыми линиями группы в шаг не включаются.
Public Sub RuleStepBounds(ByVal ruleID As String, ByRef stepCount As Long, ByRef minStep As Double, ByRef maxStep As Double, _
        ByRef minChord As Double, ByRef maxChord As Double)
    If mDefinition Is Nothing Then RuleError 0, "Правила еще не выполнялись."
    Dim index As Long, i As Long, previous As Object, first As Object, id As Variant, path As CGeometryPath, step As Double
    index = mDefinition.RuleIndex(ruleID)
    Set previous = CreateObject("Scripting.Dictionary"): previous.CompareMode = vbBinaryCompare
    Set first = CreateObject("Scripting.Dictionary"): first.CompareMode = vbBinaryCompare
    stepCount = 0: minStep = 1E+100: maxStep = 0#: minChord = 1E+100: maxChord = 0#
    For i = mStart(index) To mStart(index) + mCount(index) - 1
        id = mPositions(i).PathID
        If previous.Exists(id) Then
            step = Abs(mPositions(i).Station - mPositions(CLng(previous(id))).Station)
            AccumulateSpacing i, CLng(previous(id)), step, stepCount, minStep, maxStep, minChord, maxChord
        Else
            first.Add id, i
        End If
        previous(id) = i
    Next i
    For Each id In first.Keys
        Set path = mDefinition.Path(CStr(id))
        If path.IsClosed Then
            step = path.Length - Abs(mPositions(CLng(previous(id))).Station - mPositions(CLng(first(id))).Station)
            AccumulateSpacing CLng(first(id)), CLng(previous(id)), GeomMax(0#, step), stepCount, minStep, maxStep, minChord, maxChord
        End If
    Next id
    If stepCount = 0 Then minStep = 0#: minChord = 0#
End Sub

' Учитывает отдельно расстояние вдоль линии и прямое расстояние между центрами стержней,
' чтобы для дуги не подменить длину участка длиной хорды.
Private Sub AccumulateSpacing(ByVal first As Long, ByVal last As Long, ByVal step As Double, ByRef count As Long, _
        ByRef minStep As Double, ByRef maxStep As Double, ByRef minChord As Double, ByRef maxChord As Double)
    Dim chord As Double
    count = count + 1: minStep = GeomMin(minStep, step): maxStep = GeomMax(maxStep, step)
    chord = Sqr((mPositions(first).X - mPositions(last).X) ^ 2 + (mPositions(first).Y - mPositions(last).Y) ^ 2)
    minChord = GeomMin(minChord, chord): maxChord = GeomMax(maxChord, chord)
End Sub

' Собирает текстовый отчёт по всем правилам и счётчикам.
' Длины в отчёте указаны в мм независимо от выбранных единиц вывода.
Public Function DiagnosticsText() As String
    Dim i As Long, diagnostic As CGeometryDiagnostic
    DiagnosticsText = "Requested=" & CStr(RequestedPositions) & "; Created=" & CStr(CreatedPositions) & _
        "; Merged=" & CStr(MergedPositions) & "; Skipped=" & CStr(mSkippedProjections + mSkippedOffsets + mFilteredByArea) & _
        "; FilteredArea=" & CStr(mFilteredByArea) & _
        "; Conflicts=" & CStr(Conflicts) & "; BarOverlaps=" & CStr(BarOverlaps) & _
        "; DiagnosticRecordsOmitted=" & CStr(DiagnosticRecordsOmitted) & vbCrLf
    If mSkippedOffsets > 0 Then
        DiagnosticsText = DiagnosticsText & SkippedOffsetsWarning() & vbCrLf
        For i = 1 To mDefinition.RuleCount
            If mRules(i).Mode = DependentOffset And mSkippedByRule(i) > 0 Then DiagnosticsText = DiagnosticsText & _
                mRules(i).SourceReference & ": " & mRules(i).Comment & "; пропущено " & CStr(mSkippedByRule(i)) & " стержней." & vbCrLf
        Next i
        For Each diagnostic In mDiagnostics
            If diagnostic.Code = "OFFSET_POSITION_SKIPPED" Then DiagnosticsText = DiagnosticsText & diagnostic.SourceReference & ": " & _
                diagnostic.EntityID & "; " & diagnostic.RelatedEntityID & ". " & diagnostic.Reason & vbCrLf
        Next diagnostic
    End If
    If mSkippedCoincidences > 0 Then
        DiagnosticsText = DiagnosticsText & SkippedCoincidencesWarning() & vbCrLf
        For Each diagnostic In mDiagnostics
            If diagnostic.Code = "COINCIDENCE_SKIPPED" Then DiagnosticsText = DiagnosticsText & diagnostic.SourceReference & ": " & diagnostic.EntityID & "; " & diagnostic.Reason & vbCrLf
        Next diagnostic
    End If
    If mFilteredByArea > 0 Then DiagnosticsText = DiagnosticsText & FilteredAreaWarning() & vbCrLf
    If mMergedOffsetNeighbors > 0 Then
        DiagnosticsText = DiagnosticsText & MergedOffsetNeighborsWarning() & vbCrLf
        For Each diagnostic In mDiagnostics
            If diagnostic.Code = "OFFSET_NEIGHBOR_MERGED" Then DiagnosticsText = DiagnosticsText & diagnostic.SourceReference & ": " & diagnostic.Reason & vbCrLf
        Next diagnostic
    End If
    If Not mDefinition Is Nothing Then
        For i = 1 To mDefinition.RuleCount: DiagnosticsText = DiagnosticsText & RuleTrace(mRules(i).RuleID) & vbCrLf: Next i
    End If
End Function

' Распознаёт значимые пропуски и объединения по полным счётчикам, включая события сверх лимита записей.
Public Property Get HasGeometryWarnings() As Boolean
    HasGeometryWarnings = (mSkippedOffsets > 0 Or mMergedOffsetNeighbors > 0 Or mFilteredByArea > 0 Or mSkippedCoincidences > 0)
End Property

' Готовит краткие строки по событиям из полных счётчиков без анализа подробного отчёта.
Public Property Get GeometryWarningText() As String
    If mSkippedCoincidences > 0 Then GeometryWarningText = SkippedCoincidencesWarning()
    If mSkippedOffsets > 0 Then
        If Len(GeometryWarningText) > 0 Then GeometryWarningText = GeometryWarningText & vbCrLf
        GeometryWarningText = GeometryWarningText & SkippedOffsetsWarning()
    End If
    If mMergedOffsetNeighbors > 0 Then
        If Len(GeometryWarningText) > 0 Then GeometryWarningText = GeometryWarningText & vbCrLf
        GeometryWarningText = GeometryWarningText & MergedOffsetNeighborsWarning()
    End If
    If mFilteredByArea > 0 Then
        If Len(GeometryWarningText) > 0 Then GeometryWarningText = GeometryWarningText & vbCrLf
        GeometryWarningText = GeometryWarningText & FilteredAreaWarning()
    End If
End Property

' Поясняет штатный пропуск совпавших позиций с сохранением первого физического стержня.
Private Function SkippedCoincidencesWarning() As String
    SkippedCoincidencesWarning = "Предупреждение: пропущено совпавших стержней: " & CStr(mSkippedCoincidences) & _
        ". В каждой совпавшей позиции сохранён первый добавленный стержень с его диаметром и материалом. На схеме и в расчёте учтены фактически созданные стержни. Подробности и адреса полей приведены в Geometry.RecipeDiagnostics на Results."
End Function

' Поясняет пропуск дополнительных позиций; наличие события определяет счётчик, а не эта строка.
Private Function SkippedOffsetsWarning() As String
    SkippedOffsetsWarning = "Предупреждение: после смещения не размещено дополнительных стержней: " & CStr(mSkippedOffsets) & _
        ". Их положения по нормали оказались вне полного смещённого контура. Подробности и адреса полей приведены в Geometry.RecipeDiagnostics на Results."
End Function

' Поясняет объединение соседей с сохранением одного физического стержня.
Private Function MergedOffsetNeighborsWarning() As String
    MergedOffsetNeighborsWarning = "Предупреждение: объединено дополнительных стержней соседних граней: " & CStr(mMergedOffsetNeighbors) & _
        ". В расчёте и подписях учтён только первый добавленный стержень с его координатами и диаметром."
End Function

' Сообщает число геометрических позиций, которые не становятся физическими стержнями.
' Зависимые ряды используют исходные позиции; пропуск не меняет их координаты и порядковые номера.
Private Function FilteredAreaWarning() As String
    FilteredAreaWarning = "По настройке Mesh.MinElementArea пропущено стержней: " & CStr(mFilteredByArea) & _
        ". Площадь каждого из них меньше заданного порога; в расчёте и подписях они не учитываются."
End Function

' Размещает стержни по готовому описанию в переданной бетонной области
' и возвращает их список. Бетонная сетка строится отдельно.
Public Function Build(ByVal definition As CSectionDefinition, ByVal geometry As CGeometryRegion, _
        Optional ByVal minArea As Double = 0#) As CRebarLayout
    If mReadOnly Then SectionRecipeError "CRebarRuleExecutor", "Опубликованные правила нельзя выполнять заново в том же объекте. Создайте новый исполнитель."
    mBuilt = False
    Set mDefinition = Nothing ' Даже отказ до подготовки начинает новый отчет, без счетчиков и ID предыдущего успешного запуска того же исполнителя.
    mPositionCount = 0: mCapacity = 0: mSkippedProjections = 0: mSkippedOffsets = 0: Erase mPositions
    mMergedOffsetNeighbors = 0: mFilteredByArea = 0: mSkippedCoincidences = 0
    Erase mFilteredAreaByRule: Erase mRuleArea: Erase mRuleBelowMinArea
    Erase mTotalNormalOffset
    Set mOffsetContours = Nothing
    mMergedPositions = 0: mConflicts = 0: mBarOverlaps = 0: mPhysicalCapacity = 0: mCreatedPhysical = 0
    mDiagnosticRecordsOmitted = 0
    Set mDiagnostic = Nothing: Set mDiagnostics = New Collection
    Set mBarBuckets = CreateObject("Scripting.Dictionary"): Erase mPhysicalRule
    Erase mPhysicalFirstPosition: Erase mPhysicalOffsetRule
    Erase mPhysicalOwner: Erase mRules: Erase mState: Erase mStart: Erase mCount
    Erase mSkippedByRule: Erase mMergedByRule: Erase mCreatedByRule
    If minArea < 0# Then RuleError 0, "Mesh.MinElementArea не может быть отрицательной."
    If definition Is Nothing Or geometry Is Nothing Then RuleError 0, "Определение и бетонная область обязательны."
    If Not definition.IsSealed Then RuleError 0, "Определение еще не опубликовано."
    Set mDefinition = definition
    Dim ruleCount As Long, i As Long, result As CRebarLayout
    ruleCount = definition.RuleCount
    Set result = New CRebarLayout
    If ruleCount = 0 Then mBuilt = True: Set Build = result: Exit Function
    ReDim mRules(1 To ruleCount): ReDim mState(1 To ruleCount)
    ReDim mStart(1 To ruleCount): ReDim mCount(1 To ruleCount): ReDim mTotalNormalOffset(1 To ruleCount)
    ReDim mSkippedByRule(1 To ruleCount): ReDim mMergedByRule(1 To ruleCount): ReDim mCreatedByRule(1 To ruleCount)
    ReDim mFilteredAreaByRule(1 To ruleCount): ReDim mRuleArea(1 To ruleCount): ReDim mRuleBelowMinArea(1 To ruleCount)
    For i = 1 To ruleCount: definition.CopyRule i, mRules(i): Next i
    EnableRowOffsetChecks
    mBarBinSize = 1#
    For i = 1 To ruleCount
        If mRules(i).Enabled Then mBarBinSize = GeomMax(mBarBinSize, mRules(i).Diameter) ' Ячейка не меньше любого диаметра: для пересечений достаточно своей и восьми соседних ячеек.
    Next i
    Preflight
    For i = 1 To ruleCount
        If mRules(i).Enabled And mRules(i).ValidateRowOffset Then
            Set mOffsetContours = New CRebarOffsetContours
            mOffsetContours.Initialize definition, mRules
            Exit For
        End If
    Next i
    For i = 1 To ruleCount: PrepareRule i: Next i
    If minArea > 0# Then
        For i = 1 To ruleCount
            If mCount(i) > 0 Then
                mRuleArea(i) = GEOM_PI * mRules(i).Diameter * mRules(i).Diameter / 4#
                mRuleBelowMinArea(i) = (mRuleArea(i) < minArea)
            End If
        Next i
    End If
    If mPositionCount > 0 Then ReDim mPhysicalOwner(1 To mPositionCount) Else Erase mPhysicalOwner
    Dim order() As Long, scratch() As Long, index As Long, position As TRebarPosition
    Dim region As CConcreteRegion, query As CSectionGeometryQuery, owner As Long, identical As Boolean
    Set region = definition.ConcreteDomain: Set query = New CSectionGeometryQuery
    If mPositionCount > 0 Then
        ReDim order(1 To mPositionCount): ReDim scratch(1 To mPositionCount)
        For i = 1 To mPositionCount: order(i) = i: Next i
        SortOrder order, scratch, 1, mPositionCount
        For i = 1 To mPositionCount
            position = mPositions(order(i)): index = position.RuleIndex
            If mRuleBelowMinArea(index) Then
                mFilteredByArea = mFilteredByArea + 1: mFilteredAreaByRule(index) = mFilteredAreaByRule(index) + 1
                AddDiagnostic index, "MIN_ELEMENT_AREA", position.SourcePositionID, _
                    "Стержень пропущен: площадь его диаметра меньше Mesh.MinElementArea.", mRuleArea(index), minArea
                GoTo NextPosition
            End If
            If MergeDuplicateMainSource(index, order(i), result) Then GoTo NextPosition
            owner = OffsetNeighborOwner(index, position, result)
            If owner > 0 Then
                ValidatePlacement OffsetCheckRule(index, owner), result.X(owner), result.Y(owner), result.Diameter(owner), result.BarID(owner), geometry, region, query
                MergeOffsetNeighbor index, owner, order(i), result
                GoTo NextPosition
            End If
            owner = result.FindBarAt(position.X, position.Y)
            If owner > 0 Then
                If mRules(index).ValidateRowOffset Or mPhysicalOffsetRule(owner) > 0 Then
                    If Not SharedSourceJoint(order(i), mPhysicalFirstPosition(owner)) Then _
                        RuleError OffsetCheckRule(index, owner), "Дополнительный стержень совпал с самостоятельным стержнем. Измените отступы рядов.", "ROW_OFFSET_COINCIDENCE", result.BarID(owner), position.X, position.Y
                End If
                identical = (result.Diameter(owner) = mRules(index).Diameter And _
                    StrComp(result.SteelClass(owner), mRules(index).SteelClass, vbTextCompare) = 0)
                If mRules(index).Coincidence = RejectCoincidence Then _
                    RuleError index, "Позиция совпадает со стержнем '" & result.BarID(owner) & "'; политика объединения не задана.", "COINCIDENCE", result.BarID(owner), position.X, position.Y
                If Not identical And mRules(index).Coincidence = MergeIdentical Then _
                    RuleError index, "Совпавшие позиции имеют разные диаметры или материалы.", "COINCIDENCE_CONFLICT", result.BarID(owner), result.Diameter(owner), mRules(index).Diameter
                ValidatePlacement index, result.X(owner), result.Y(owner), result.Diameter(owner), position.SourcePositionID, geometry, region, query
                If mRules(index).ValidateRowOffset Then
                    mPhysicalOffsetRule(owner) = index
                    CheckNeighbors index, position, result, owner, result.Diameter(owner)
                End If
                If mRules(index).WarnOnCoincidence Then
                    mSkippedCoincidences = mSkippedCoincidences + 1
                    AddDiagnostic index, "COINCIDENCE_SKIPPED", result.BarID(owner), _
                        "Стержень '" & position.SourcePositionID & "' пропущен: его центр совпал с ранее добавленным стержнем '" & result.BarID(owner) & _
                        "'. Сохранены диаметр " & CStr(result.Diameter(owner)) & " мм и материал " & result.SteelClass(owner) & _
                        " первого стержня; запрошенный диаметр пропущенного — " & CStr(mRules(index).Diameter) & " мм.", result.Diameter(owner), mRules(index).Diameter
                ElseIf Not identical Then
                    mConflicts = mConflicts + 1
                    AddDiagnostic index, "COINCIDENCE_OWNER", result.BarID(owner), "В совпавшей позиции сохранён стержень, которому правило отдаёт приоритет.", result.Diameter(owner), mRules(index).Diameter
                End If
                mMergedPositions = mMergedPositions + 1
                mMergedByRule(index) = mMergedByRule(index) + 1
                mPhysicalOwner(order(i)) = owner
                result.AddMembership owner, mRules(index).AnnotationGroup
                GoTo NextPosition
            End If
            ValidatePlacement index, position.X, position.Y, mRules(index).Diameter, position.SourcePositionID, geometry, region, query
            CheckNeighbors index, position, result
            result.AddBar position.SourcePositionID, position.X, position.Y, mRules(index).Diameter, 0#, _
                mRules(index).SteelClass, mRules(index).Comment, geometry, mRules(index).AnnotationGroup
            mCreatedByRule(index) = mCreatedByRule(index) + 1
            mCreatedPhysical = mCreatedPhysical + 1
            mPhysicalOwner(order(i)) = result.Count
            IndexPhysicalBar result, index, order(i)
NextPosition:
        Next i
    End If
    mBuilt = True
    Set Build = result
End Function

' Не создаёт повторные дочерние стержни по двум совпавшим позициям одного Main на той же линии.
' Дополнительные ряды разных сторон или исходных линий независимы и проходят общую проверку конфликтов.
Private Function MergeDuplicateMainSource(ByVal index As Long, ByVal positionIndex As Long, ByVal result As CRebarLayout) As Boolean
    Dim source As Long, original As Long ' Первичная позиция Main для текущего потомка и позиция, впервые создавшая физический основной стержень.
    Dim mainOwner As Long, owner As Long, rule As Long ' Физические владельцы основного и дочернего стержней и номер исходного Main-правила.
    source = mPositions(positionIndex).SourcePositionIndex
    If source = 0 Then Exit Function
    Do While mPositions(source).SourcePositionIndex > 0: source = mPositions(source).SourcePositionIndex: Loop
    rule = mPositions(source).RuleIndex
    If Not mRules(rule).WarnOnCoincidence Then Exit Function
    mainOwner = mPhysicalOwner(source)
    If mainOwner = 0 Then Exit Function
    original = mPhysicalFirstPosition(mainOwner)
    If original = source Then Exit Function
    If mPositions(original).RuleIndex <> rule Or mPositions(original).PathID <> mPositions(source).PathID Then Exit Function
    owner = result.FindBarAt(mPositions(positionIndex).X, mPositions(positionIndex).Y)
    If owner = 0 Then Exit Function
    If mPhysicalRule(owner) <> index Then Exit Function ' Слияние допустимо только внутри того же дочернего правила; диаметр и материал чужого ряда не подменяются.
    mMergedPositions = mMergedPositions + 1: mMergedByRule(index) = mMergedByRule(index) + 1
    mPhysicalOwner(positionIndex) = owner
    result.AddMembership owner, mRules(index).AnnotationGroup
    AddDiagnostic index, "DEPENDENT_SOURCE_DUPLICATE_SKIPPED", result.BarID(owner), _
        "Повторный зависимый стержень пропущен: совпавшая позиция того же основного ряда уже имеет физического владельца на той же линии.", mPositions(positionIndex).X, mPositions(positionIndex).Y
    MergeDuplicateMainSource = True
End Function

' Выбирает самый ранний физический стержень среди пересекающихся дополнительных рядов соседних граней.
' Основные стержни, разные бетонные границы и ряды одного семейства сюда не относятся.
Private Function OffsetNeighborOwner(ByVal index As Long, ByRef position As TRebarPosition, ByVal result As CRebarLayout) As Long
    Dim bx As Double, by As Double, dx As Long, dy As Long, key As String, bucket As Collection, candidate As Variant
    Dim other As Long, otherRule As Long, distanceSquared As Double, radiusSum As Double
    If mOffsetContours Is Nothing Or mRules(index).Mode <> DependentOffset Then Exit Function
    bx = Int(position.X / mBarBinSize): by = Int(position.Y / mBarBinSize)
    For dx = -1 To 1
        For dy = -1 To 1
            key = Format$(bx + dx, "0") & "|" & Format$(by + dy, "0")
            If mBarBuckets.Exists(key) Then
                Set bucket = mBarBuckets(key)
                For Each candidate In bucket
                    other = CLng(candidate): otherRule = mPhysicalRule(other)
                    If mRules(otherRule).Mode = DependentOffset Then
                        If mRules(index).ValidateRowOffset Or mRules(otherRule).ValidateRowOffset Then
                            radiusSum = (mRules(index).Diameter + result.Diameter(other)) / 2#
                            distanceSquared = (position.X - result.X(other)) ^ 2 + (position.Y - result.Y(other)) ^ 2
                            If radiusSum > GEOM_TOLERANCE And distanceSquared < (radiusSum - GEOM_TOLERANCE) ^ 2 Then
                                If mOffsetContours.AreAdjacent(index, otherRule) Then
                                    If OffsetNeighborOwner = 0 Or other < OffsetNeighborOwner Then OffsetNeighborOwner = other
                                End If
                            End If
                        End If
                    End If
                Next candidate
            End If
        Next dy
    Next dx
End Function

' Сохраняет координаты, площадь, материал, диаметр и основную группу подписи первого стержня.
' Дополнительная принадлежность второй группе остаётся в связях без повторного подсчёта площади
' и стержня в подписи; пользователь получает предупреждение с обоими рядами.
Private Sub MergeOffsetNeighbor(ByVal index As Long, ByVal owner As Long, ByVal positionIndex As Long, ByVal result As CRebarLayout)
    Dim firstRule As Long, reason As String
    firstRule = mPhysicalRule(owner)
    mMergedPositions = mMergedPositions + 1: mMergedByRule(index) = mMergedByRule(index) + 1
    mMergedOffsetNeighbors = mMergedOffsetNeighbors + 1: mPhysicalOwner(positionIndex) = owner
    If mRules(index).ValidateRowOffset Then mPhysicalOffsetRule(owner) = index
    result.AddMembership owner, mRules(index).AnnotationGroup
    reason = "Ряд '" & mRules(index).RuleID & "': пересекающийся стержень диаметром " & CStr(mRules(index).Diameter) & _
        " мм объединён с первым стержнем '" & result.BarID(owner) & "' ряда '" & mRules(firstRule).RuleID & _
        "'. Сохранены координаты (" & CStr(result.X(owner)) & "; " & CStr(result.Y(owner)) & ") мм и диаметр " & _
        CStr(result.Diameter(owner)) & " мм; вторая площадь не добавлена. Поля первого ряда: " & mRules(firstRule).SourceReference
    AddDiagnostic index, "OFFSET_NEIGHBOR_MERGED", result.BarID(owner), reason, result.Diameter(owner), mRules(index).Diameter
End Sub

' Включает проверки обоих дополнительных рядов только у семейства с активным ненулевым базовым отступом.
Private Sub EnableRowOffsetChecks()
    Dim families As Object, i As Long
    Set families = CreateObject("Scripting.Dictionary")
    families.CompareMode = vbBinaryCompare
    For i = 1 To mDefinition.RuleCount
        If mRules(i).Enabled And mRules(i).Mode = DependentOffset And mRules(i).ValidateRowOffset Then families(mRules(i).SourceRuleID) = True
    Next i
    For i = 1 To mDefinition.RuleCount
        If mRules(i).Enabled And mRules(i).Mode = DependentOffset Then
            If families.Exists(mRules(i).SourceRuleID) Then
                mRules(i).ValidateRowOffset = True
                mRules(i).PlacementCheck = CheckFullDisk
                mRules(i).OverlapPolicy = RejectOverlap
            End If
        End If
    Next i
End Sub

' Возвращает правило проверки смещённого ряда, чтобы ошибка указывала на его исходные поля.
Private Function OffsetCheckRule(ByVal index As Long, ByVal owner As Long) As Long
    OffsetCheckRule = index
    If Not mRules(index).ValidateRowOffset And mPhysicalOffsetRule(owner) > 0 Then OffsetCheckRule = mPhysicalOffsetRule(owner)
End Function

' Допускает совместное владение только при одном исходном узле на концах разных линий.
' Совпадение самостоятельных стержней внутри линии или двух рядов одного Main таким узлом не считается.
Private Function SharedSourceJoint(ByVal first As Long, ByVal last As Long) As Boolean
    Do While mPositions(first).SourcePositionIndex > 0: first = mPositions(first).SourcePositionIndex: Loop
    Do While mPositions(last).SourcePositionIndex > 0: last = mPositions(last).SourcePositionIndex: Loop
    If mPositions(first).PathID = mPositions(last).PathID Then Exit Function
    If Abs(mPositions(first).X - mPositions(last).X) > GEOM_TOLERANCE Or Abs(mPositions(first).Y - mPositions(last).Y) > GEOM_TOLERANCE Then Exit Function
    Dim firstPath As CGeometryPath, lastPath As CGeometryPath
    Set firstPath = mDefinition.Path(mPositions(first).PathID): Set lastPath = mDefinition.Path(mPositions(last).PathID)
    If firstPath.IsClosed Or lastPath.IsClosed Then Exit Function
    SharedSourceJoint = (mPositions(first).Station <= GEOM_TOLERANCE Or firstPath.Length - mPositions(first).Station <= GEOM_TOLERANCE) And _
        (mPositions(last).Station <= GEOM_TOLERANCE Or lastPath.Length - mPositions(last).Station <= GEOM_TOLERANCE)
End Function

' Сначала выбирает близкие стержни по координатным ячейкам, затем проверяет
' зазор по исходным координатам центров и радиусам. Простое касание
' стержней зависимых рядов допускается.
Private Sub CheckNeighbors(ByVal index As Long, ByRef position As TRebarPosition, ByVal result As CRebarLayout, Optional ByVal excludedOwner As Long = 0, Optional ByVal actualDiameter As Double = -1#)
    Dim bx As Double, by As Double, dx As Long, dy As Long, key As String, bucket As Collection, candidate As Variant
    Dim gap As Double, other As Long
    If actualDiameter < 0# Then actualDiameter = mRules(index).Diameter
    bx = Int(position.X / mBarBinSize): by = Int(position.Y / mBarBinSize)
    For dx = -1 To 1
        For dy = -1 To 1
            key = Format$(bx + dx, "0") & "|" & Format$(by + dy, "0")
            If mBarBuckets.Exists(key) Then
                Set bucket = mBarBuckets(key)
                For Each candidate In bucket
                    other = CLng(candidate)
                    If other = excludedOwner Then GoTo NextCandidate
                    gap = Sqr((position.X - result.X(other)) ^ 2 + (position.Y - result.Y(other)) ^ 2) - _
                        (actualDiameter + result.Diameter(other)) / 2#
                    If gap < -GEOM_TOLERANCE Then
                        mBarOverlaps = mBarOverlaps + 1
                        If mRules(index).OverlapPolicy = RejectOverlap Or mRules(mPhysicalRule(other)).OverlapPolicy = RejectOverlap Or mPhysicalOffsetRule(other) > 0 Then _
                            RuleError OffsetCheckRule(index, other), "Поперечные сечения стержней пересекаются. Измените шаг арматуры или смещения связанных рядов.", "BAR_OVERLAP", result.BarID(other), gap, actualDiameter
                        AddDiagnostic index, "BAR_OVERLAP", result.BarID(other), "Поперечные сечения стержней пересекаются: расстояние между центрами меньше суммы радиусов.", gap, mRules(index).Diameter
                    End If
NextCandidate:
                Next candidate
            End If
        Next dy
    Next dx
End Sub

' Запоминает координатную ячейку добавленного стержня для дальнейшей проверки
' совпадающих центров и пересечений поперечных сечений.
Private Sub IndexPhysicalBar(ByVal result As CRebarLayout, ByVal ruleIndex As Long, ByVal positionIndex As Long)
    Dim key As String, bucket As Collection, index As Long
    index = result.Count
    If index > mPhysicalCapacity Then
        mPhysicalCapacity = CLng(GeomMin(GEOM_MAX_REBAR_COUNT, GeomMax(16#, CDbl(mPhysicalCapacity) * 2#)))
        ReDim Preserve mPhysicalRule(1 To mPhysicalCapacity)
        ReDim Preserve mPhysicalFirstPosition(1 To mPhysicalCapacity)
        ReDim Preserve mPhysicalOffsetRule(1 To mPhysicalCapacity)
    End If
    mPhysicalRule(index) = ruleIndex
    mPhysicalFirstPosition(index) = positionIndex
    If mRules(ruleIndex).ValidateRowOffset Then mPhysicalOffsetRule(index) = ruleIndex
    key = Format$(Int(result.X(index) / mBarBinSize), "0") & "|" & Format$(Int(result.Y(index) / mBarBinSize), "0")
    If mBarBuckets.Exists(key) Then Set bucket = mBarBuckets(key) Else Set bucket = New Collection: mBarBuckets.Add key, bucket
    bucket.Add index
End Sub

' Проверяет положение и геометрическую допустимость заявленного стержня в бетонной области.
Private Sub ValidatePlacement(ByVal index As Long, ByVal x As Double, ByVal y As Double, ByVal diameter As Double, _
        ByVal positionID As String, ByVal geometry As CGeometryRegion, ByVal region As CConcreteRegion, ByVal query As CSectionGeometryQuery)
    If Not geometry.ContainsPoint(x, y) Then RuleError index, "Ось стержня '" & positionID & "' вне бетона. Уменьшите смещение или измените положение ряда.", "BAR_CENTER_OUTSIDE", positionID, x, y
    If mRules(index).PlacementCheck = CheckFullDisk Then
        Dim cover As Double
        cover = query.BoundaryDistance(region, x, y)
        If cover + GEOM_TOLERANCE < diameter / 2# Then _
            RuleError index, "Стержень '" & positionID & "' не помещается целиком внутри бетона. Уменьшите диаметр или измените положение ряда.", "BAR_DISK_OUTSIDE", positionID, cover, diameter / 2#
    End If
End Sub

' До выделения массивов проверяет общее число запрошенных стержней с учётом
' всех независимых линий и зависимых рядов. Для проекций учитывает и направления,
' которые позднее могут быть пропущены из-за отсутствия пересечения.
Private Sub Preflight()
    Dim state() As Long, counts() As Double, total As Double, i As Long
    ReDim state(1 To mDefinition.RuleCount)
    ReDim counts(1 To mDefinition.RuleCount)
    For i = 1 To mDefinition.RuleCount
        EstimateRule i, state, counts
        total = total + counts(i)
        If total > GEOM_MAX_REBAR_COUNT Then RuleError i, "Суммарное число запрошенных позиций превышает диапазон счетчиков; уменьшите количество или число активных рядов."
    Next i
End Sub

' Оценивает число позиций до выделения массивов; EverySecond учитывается отдельно на каждом
' независимом пути.
Private Sub EstimateRule(ByVal index As Long, ByRef state() As Long, ByRef counts() As Double)
    If state(index) = 2 Then Exit Sub
    If state(index) = 1 Then RuleError index, "Зависимости рядов арматуры образуют замкнутую цепочку. Уберите циклическую ссылку между правилами."
    state(index) = 1 ' Состояния предварительного обхода: 0 — не посещено, 1 — в цепочке, 2 — оценено.
    If mRules(index).Enabled Then
        ValidateOptions index
        Dim source As Long, ids As Variant, id As Variant, path As CGeometryPath
        If mRules(index).Mode = ProjectFrom Or mRules(index).Mode = DependentOffset Then
            source = mDefinition.RuleIndex(mRules(index).SourceRuleID)
            EstimateRule source, state, counts
            counts(index) = counts(source)
            If mRules(index).EverySecond And mRules(source).Mode <= ExplicitPositions And counts(source) > 0# Then
                counts(index) = 0# ' В режиме PerPath выбираем позиции 1, 3, ... заново на каждой линии. Так независимые ряды с нечётным числом стержней не теряют последний стержень.
                ids = mDefinition.TargetPaths(mRules(source).Target)
                For Each id In ids
                    Set path = mDefinition.Path(CStr(id))
                    counts(index) = counts(index) - Int(-PositionCountFor(source, path) / 2#)
                Next id
            End If
        Else
            ids = mDefinition.TargetPaths(mRules(index).Target)
            For Each id In ids
                Set path = mDefinition.Path(CStr(id))
                ValidateWorkInterval index, path
                counts(index) = counts(index) + PositionCountFor(index, path)
            Next id
        End If
        If counts(index) > GEOM_MAX_REBAR_COUNT Then RuleError index, "Число мест установки арматуры превышает предел счётчика программы. Уменьшите количество стержней или увеличьте их шаг."
    End If
    state(index) = 2
End Sub

' Определяет число мест установки стержней на одной линии с учётом шага
' и включения концов ряда. Координаты на этом этапе ещё не создаются.
Private Function PositionCountFor(ByVal index As Long, ByVal path As CGeometryPath) As Double
    Dim intervals As Double, length As Double
    length = path.Length - mRules(index).StartSetback - mRules(index).EndSetback
    Select Case mRules(index).Mode
        Case ByCount
            PositionCountFor = mRules(index).Count
        Case ExplicitPositions
            If Not IsArray(mRules(index).ExplicitStations) Then RuleError index, "Не заданы расстояния от начала линии до мест установки стержней."
            PositionCountFor = UBound(mRules(index).ExplicitStations) - LBound(mRules(index).ExplicitStations) + 1#
        Case ByMaxStep
            intervals = -Int(-length / mRules(index).Step)
            Select Case mRules(index).EndMode
                Case IncludeBoth: PositionCountFor = intervals + 1#
                Case ExcludeBoth: PositionCountFor = GeomMax(0#, intervals - 1#)
                Case Else: PositionCountFor = intervals
            End Select
        Case ByFixedStep
            intervals = length / mRules(index).Step
            If RequiresIntegralFixedStep(index, path) Then
                If Abs(intervals - Round(intervals)) > 0.0000000001 Then RuleError index, "Точный шаг несовместим с фиксированными обоими концами или полным замкнутым обходом."
                PositionCountFor = Round(intervals)
                If mRules(index).EndMode = IncludeBoth Then PositionCountFor = PositionCountFor + 1#
            Else
                PositionCountFor = Fix(intervals) + 1#
            End If
        Case Else
            RuleError index, "Не распознан способ размещения арматуры. Проверьте правило задания количества, шага или проекции."
    End Select
    If PositionCountFor < 0# Or PositionCountFor > GEOM_MAX_REBAR_COUNT Then RuleError index, "Число мест установки арматуры превышает предел счётчика программы. Уменьшите количество стержней или увеличьте их шаг."
End Function

' Подготавливает правило после его зависимостей и сохраняет исходные номера позиций.
' Порядок добавления стержней и пропуски проекций не меняют выбор каждого второго стержня.
Private Sub PrepareRule(ByVal index As Long)
    On Error GoTo InvalidGeometry
    If mState(index) = 2 Then Exit Sub
    If mState(index) = 1 Then RuleError index, "Зависимости рядов арматуры образуют замкнутую цепочку. Уберите циклическую ссылку между правилами."
    mState(index) = 1 ' Состояния обхода зависимостей: 0 — не посещено, 1 — выполняется, 2 — завершено.
    If Not mRules(index).Enabled Then mState(index) = 2: Exit Sub
    Dim source As Long ' Номер исходного правила для зависимости или проекции.
    If mRules(index).Mode = DependentOffset Or mRules(index).Mode = ProjectFrom Then
        source = mDefinition.RuleIndex(mRules(index).SourceRuleID)
        PrepareRule source
    End If
    mStart(index) = mPositionCount + 1 ' Каждое правило хранит свой непрерывный интервал подготовленных позиций.
    Select Case mRules(index).Mode
        Case ByCount, ByMaxStep, ByFixedStep, ExplicitPositions
            PrepareIndependent index
        Case DependentOffset
            PrepareDependent index, source
        Case ProjectFrom
            PrepareProjected index, source
        Case Else
            RuleError index, "Не распознан способ размещения арматуры. Проверьте правило задания количества, шага или проекции."
    End Select
    mCount(index) = mPositionCount - mStart(index) + 1 ' Пропущенные проекции/смещения не занимают места в массиве позиций.
    mState(index) = 2
    Exit Sub
InvalidGeometry:
    SectionRecipeGeometryError "CRebarRuleExecutor", mRules(index).SourceReference, mRules(index).Comment
End Sub

' Вычисляет координаты независимого ряда. Нумерация позиций начинается с 1
' на каждой линии, а уникальные имена и порядок добавления продолжаются по всей группе.
Private Sub PrepareIndependent(ByVal index As Long)
    Dim ids As Variant, id As Variant, path As CGeometryPath, stations As Variant, ordinal As Long, i As Long
    Dim position As TRebarPosition, first As Long
    ids = mDefinition.TargetPaths(mRules(index).Target)
    For Each id In ids
        Set path = mDefinition.Path(CStr(id))
        stations = StationsFor(index, path)
        If Not IsEmpty(stations) Then
            first = mPositionCount + 1
            For i = LBound(stations) To UBound(stations)
                ordinal = ordinal + 1
                path.AtDistance CDbl(stations(i)), position.X, position.Y, position.TangentX, position.TangentY, mRules(index).IncomingTangentAtJoint
                position.Station = CDbl(stations(i)): position.RuleIndex = index
                position.SourcePositionIndex = 0: position.IncomingTangentAtJoint = mRules(index).IncomingTangentAtJoint ' У независимого ряда нет родительской позиции; сохраняется только выбор касательной на стыке.
                position.RootRuleIndex = index: position.RootOrdinal = ordinal
                position.SourceOrdinal = i - LBound(stations) + 1 ' Сохраняем общий порядок добавления стержней; выбор позиций для зависимого ряда определяется отдельно по исходным номерам на каждой линии.
                position.RootPublicationRank = ordinal
                position.PathID = CStr(id)
                position.SourcePositionID = PositionID(index, ordinal, "")
                AppendPosition position
            Next i
            If mRules(index).PositionOrder = FromPathMiddle Then RankFromMiddle first, mPositionCount, path.Length
        End If
    Next id
End Sub

' Упорядочивает позиции по расстоянию от середины полной линии расстановки.
' Исходные номера позиций сохраняются, в том числе при разных отступах от концов.
Private Sub RankFromMiddle(ByVal first As Long, ByVal last As Long, ByVal length As Double)
    Dim order() As Long, scratch() As Long, i As Long, base As Long
    ReDim order(first To last): ReDim scratch(first To last)
    For i = first To last: order(i) = i: Next i
    SortMiddle order, scratch, first, last, length / 2#
    base = mPositions(first).RootOrdinal - 1
    For i = first To last: mPositions(order(i)).RootPublicationRank = base + i - first + 1: Next i
End Sub

' Стабильно сортирует номера позиций по расстоянию от середины линии.
' При равных расстояниях сохраняет исходный порядок; координаты и исходные номера не меняет.
Private Sub SortMiddle(ByRef order() As Long, ByRef scratch() As Long, ByVal first As Long, ByVal last As Long, ByVal middleStation As Double)
    If first >= last Then Exit Sub
    Dim middle As Long, a As Long, b As Long, out As Long, takeA As Boolean
    middle = first + (last - first) \ 2
    SortMiddle order, scratch, first, middle, middleStation
    SortMiddle order, scratch, middle + 1, last, middleStation
    a = first: b = middle + 1
    For out = first To last
        If a > middle Then
            takeA = False
        ElseIf b > last Then
            takeA = True
        Else
            takeA = (Abs(mPositions(order(a)).Station - middleStation) <= Abs(mPositions(order(b)).Station - middleStation))
        End If
        If takeA Then scratch(out) = order(a): a = a + 1 Else scratch(out) = order(b): b = b + 1
    Next out
    For out = first To last: order(out) = scratch(out): Next out
End Sub

' Возвращает координаты первой и последней подготовленных позиций указанного правила.
Friend Function RuleEnds(ByVal ruleID As String, ByRef startX As Double, ByRef startY As Double, ByRef endX As Double, ByRef endY As Double) As Boolean
    Dim index As Long, last As Long
    index = mDefinition.RuleIndex(ruleID)
    If mCount(index) = 0 Then Exit Function
    If mFilteredAreaByRule(index) = mCount(index) Then Exit Function
    last = mStart(index) + mCount(index) - 1
    startX = mPositions(mStart(index)).X: startY = mPositions(mStart(index)).Y
    endX = mPositions(last).X: endY = mPositions(last).Y
    RuleEnds = True
End Function

' Аннотация измеряет подготовленный ряд; зависимые дополнительные позиции
' не меняют его базовую привязку. basisX/Y задают единичную ось проекций;
' без них возвращаются X/Y-границы. Координаты стержней не пересчитываются.
Friend Function RuleBounds(ByVal ruleID As String, ByRef minX As Double, ByRef maxX As Double, _
        ByRef minY As Double, ByRef maxY As Double, Optional ByVal basisX As Double = 1#, Optional ByVal basisY As Double = 0#) As Boolean
    Dim index As Long, i As Long, along As Double, across As Double
    index = mDefinition.RuleIndex(ruleID)
    If mCount(index) = 0 Then Exit Function
    If mFilteredAreaByRule(index) = mCount(index) Then Exit Function
    minX = 1E+100: minY = 1E+100: maxX = -1E+100: maxY = -1E+100
    For i = mStart(index) To mStart(index) + mCount(index) - 1
        along = mPositions(i).X * basisX + mPositions(i).Y * basisY
        across = -mPositions(i).X * basisY + mPositions(i).Y * basisX
        minX = GeomMin(minX, along): maxX = GeomMax(maxX, along)
        minY = GeomMin(minY, across): maxY = GeomMax(maxY, across)
    Next i
    RuleBounds = True
End Function

' Опорную прямую подписи выбирает крайнее положение по заданной нормали.
' Длину и середину вдоль грани определяют все стержни основного ряда,
' включая участки скруглений: их исключение смещало подпись в сторону.
Friend Function RuleExtremeEnds(ByVal ruleID As String, ByVal nx As Double, ByVal ny As Double, _
        ByRef startX As Double, ByRef startY As Double, ByRef endX As Double, ByRef endY As Double) As Boolean
    Dim first As Double, last As Double, low As Double, support As Double
    If Not RuleBounds(ruleID, first, last, low, support, ny, -nx) Then Exit Function
    startX = first * ny + support * nx: startY = -first * nx + support * ny
    endX = last * ny + support * nx: endY = -last * nx + support * ny
    RuleExtremeEnds = True
End Function

' Получает координаты зависимого ряда смещением исходных стержней.
' Сохраняет их исходные номера на каждой линии и связь с основным рядом.
Private Sub PrepareDependent(ByVal index As Long, ByVal source As Long)
    On Error GoTo InvalidTransfer
    Dim i As Long, ordinal As Long, position As TRebarPosition, paths As Object, offset As CGeometryPath
    Dim path As CGeometryPath, boundary As CGeometryPath, query As CSectionGeometryQuery, x As Double, y As Double, tx As Double, ty As Double
    Set paths = CreateObject("Scripting.Dictionary"): Set query = New CSectionGeometryQuery ' Подготовленная смещённая линия используется повторно для всех стержней на том же пути.
    mTotalNormalOffset(index) = mTotalNormalOffset(source) + mRules(index).NormalOffset ' На цепочке зависимостей нормальный перенос проверяется относительно исходной линии.
    For i = mStart(source) To mStart(source) + mCount(source) - 1
        ordinal = ordinal + 1
        position = mPositions(i)
        If Not mRules(index).EverySecond Or position.SourceOrdinal Mod 2 = 1 Then
            If mRules(index).ValidateRowOffset Then
                Set path = mDefinition.Path(position.PathID)
                x = position.X - mRules(index).NormalOffset * position.TangentY
                y = position.Y + mRules(index).NormalOffset * position.TangentX
                If Not mOffsetContours.Accepts(index, x, y) Then
                    mSkippedOffsets = mSkippedOffsets + 1: mSkippedByRule(index) = mSkippedByRule(index) + 1
                    AddDiagnostic index, "OFFSET_POSITION_SKIPPED", position.SourcePositionID, _
                        "Дополнительный стержень пропущен: его положение после нормального смещения находится вне общего проверочного контура этого ряда.", x, y
                    GoTo NextPosition
                End If
            ElseIf Abs(mTotalNormalOffset(index)) > GEOM_TOLERANCE Then
                Set path = mDefinition.Path(position.PathID)
                Set boundary = path.OffsetBoundary
                If Not boundary Is Nothing Or path.HasCollapsedNormalArc(mTotalNormalOffset(index)) Then
                    If Not paths.Exists(path.ID) Then
                        If boundary Is Nothing Then
                            Set offset = query.OffsetPath("Offset." & mRules(index).RuleID & "." & path.ID, path, mTotalNormalOffset(index))
                        Else
                            Set offset = query.OffsetPath("Offset." & mRules(index).RuleID & "." & path.ID, boundary, mTotalNormalOffset(index) * path.BoundaryDirection)
                        End If
                        paths.Add path.ID, offset
                    End If
                    Set offset = paths.Item(path.ID)
                    path.AtDistance position.Station, x, y, tx, ty, position.IncomingTangentAtJoint
                    x = x - mTotalNormalOffset(index) * ty: y = y + mTotalNormalOffset(index) * tx
                    If Not offset.ContainsPoint(x, y) Then
                        mSkippedOffsets = mSkippedOffsets + 1: mSkippedByRule(index) = mSkippedByRule(index) + 1
                        AddDiagnostic index, "OFFSET_POSITION_SKIPPED", position.SourcePositionID, _
                            "Дополнительный стержень пропущен: его положение по нормали не попадает на полный смещённый контур.", x, y
                        GoTo NextPosition
                    End If
                End If
            End If
            position.X = position.X + mRules(index).TangentOffset * position.TangentX - mRules(index).NormalOffset * position.TangentY ' Касательная контактная добавка применяется после проверки положения по нормали к полному контуру.
            position.Y = position.Y + mRules(index).TangentOffset * position.TangentY + mRules(index).NormalOffset * position.TangentX
            position.RuleIndex = index
            position.SourcePositionIndex = i ' Связь с непосредственной исходной позицией сохраняет всю цепочку переноса.
            position.SourcePositionID = PositionID(index, ordinal, mPositions(i).SourcePositionID)
            AppendPosition position
        End If
NextPosition:
    Next i
    Exit Sub
InvalidTransfer:
    If Err.Number = 6 Then RuleError index, "Отступ ряда выходит за числовой диапазон. Уменьшите заданное расстояние."
    SectionRecipeGeometryError "CRebarRuleExecutor", mRules(index).SourceReference, mRules(index).Comment, path
End Sub

' Для каждого исходного стержня ищет ближайшее пересечение по заданному направлению
' с целевыми линиями. Учитывает отступы от концов; найденные координаты
' не заменяются новой равномерной расстановкой.
Private Sub PrepareProjected(ByVal index As Long, ByVal source As Long)
    Dim ids As Variant, id As Variant, path As CGeometryPath, i As Long, ordinal As Long, targetIndex As Long
    Dim prepared As Collection
    Dim position As TRebarPosition, station As Double, ray As Double, bestRay As Double, bestStation As Double, bestID As String
    Dim directionX As Double, directionY As Double, successful As Long
    ids = mDefinition.TargetPaths(mRules(index).Target)
    Set prepared = New Collection
    For Each id In ids
        Set path = mDefinition.Path(CStr(id))
        ValidateWorkInterval index, path
        If mRules(index).StartSetback <> 0# Or mRules(index).EndSetback <> 0# Then _
            Set path = path.Slice(mRules(index).StartSetback, path.Length - mRules(index).EndSetback)
        prepared.Add path
    Next id
    For i = mStart(source) To mStart(source) + mCount(source) - 1
        ordinal = ordinal + 1
        position = mPositions(i)
        If Not mRules(index).EverySecond Or position.SourceOrdinal Mod 2 = 1 Then
            bestID = "": bestRay = 1E+100
            directionX = mRules(index).DirectionX: directionY = mRules(index).DirectionY
            If mRules(index).SourceNormalSign <> 0# Then
                directionX = -position.TangentY * mRules(index).SourceNormalSign
                directionY = position.TangentX * mRules(index).SourceNormalSign
            End If
            For targetIndex = 1 To prepared.Count
                Set path = prepared(targetIndex)
                If path.RayHit(position.X, position.Y, directionX, directionY, station, ray) Then
                    If ray < bestRay Then
                        bestRay = ray: bestStation = station + mRules(index).StartSetback: bestID = path.ID
                    End If
                End If
            Next targetIndex
            If Len(bestID) = 0 Then
                If Not mRules(index).SkipMissingProjection Then RuleError index, "Не удалось спроецировать исходный стержень на заданный участок линии размещения. Проверьте направление проекции и концевые отступы.", "PROJECTION_MISSING", position.SourcePositionID, position.X, position.Y
                mSkippedProjections = mSkippedProjections + 1
                mSkippedByRule(index) = mSkippedByRule(index) + 1
                AddDiagnostic index, "PROJECTION_MISSING", position.SourcePositionID, "Стержень пропущен: направление его проекции не пересекает заданный участок линии размещения.", position.X, position.Y
            Else
                Set path = mDefinition.Path(bestID)
                path.AtDistance bestStation, position.X, position.Y, position.TangentX, position.TangentY, mRules(index).IncomingTangentAtJoint
                position.Station = bestStation: position.PathID = bestID: position.RuleIndex = index
                position.SourcePositionIndex = 0: position.IncomingTangentAtJoint = mRules(index).IncomingTangentAtJoint ' Проекция становится новой геометрической основой для зависимых рядов.
                successful = successful + 1 ' Обычная проекция сохраняет исходные номера позиций отдельно на каждой линии, даже при пропусках. Для рядов у отверстия можно отдельно включить нумерацию только успешно полученных проекций.
                If mRules(index).SuccessfulProjectionOrdinal Then position.SourceOrdinal = successful
                If mRules(index).NewPublicationRoot Then
                    position.RootRuleIndex = index: position.RootOrdinal = successful: position.RootPublicationRank = successful
                End If
                position.SourcePositionID = PositionID(index, ordinal, mPositions(i).SourcePositionID)
                AppendPosition position
            End If
        End If
    Next i
End Sub

' Возвращает расстояния от начала линии до мест установки стержней
' для выбранного способа расстановки и рабочих отступов.
Private Function StationsFor(ByVal index As Long, ByVal path As CGeometryPath) As Variant
    ValidateWorkInterval index, path
    Dim length As Double, first As Double, last As Double, count As Double, intervals As Double
    Dim stations() As Double, i As Long, start As Double, step As Double, kept As Long
    first = mRules(index).StartSetback: last = path.Length - mRules(index).EndSetback
    length = last - first
    If mRules(index).Mode = ExplicitPositions Then
        If Not IsArray(mRules(index).ExplicitStations) Then RuleError index, "Не заданы расстояния от начала линии до мест установки стержней."
        count = UBound(mRules(index).ExplicitStations) - LBound(mRules(index).ExplicitStations) + 1
        ReDim stations(0 To CLng(count) - 1)
        For i = 0 To CLng(count) - 1
            stations(i) = CDbl(mRules(index).ExplicitStations(LBound(mRules(index).ExplicitStations) + i))
            If stations(i) < first Or stations(i) > last Then RuleError index, "Заданное расстояние до стержня выходит за участок линии, оставшийся после концевых отступов."
            If path.IsClosed And stations(i) = path.Length Then RuleError index, "Стержень в конце полного замкнутого обхода совпадает с первым. Удалите повторную конечную позицию."
        Next i
        StationsFor = stations: Exit Function
    End If
    If length <= GEOM_TOLERANCE Then
        If mRules(index).Mode <> ByCount Or mRules(index).Count <> 1 Then RuleError index, "После концевых отступов участок сократился до точки. В ней можно задать только один стержень по количеству."
    End If
    If mRules(index).Mode = ByFixedStep Then
        StationsFor = FixedStations(index, path, first, length): Exit Function
    End If
    count = mRules(index).Count
    If mRules(index).Mode = ByMaxStep Then
        intervals = -Int(-length / mRules(index).Step)
        Select Case mRules(index).EndMode
            Case IncludeBoth: count = intervals + 1#
            Case ExcludeBoth: count = GeomMax(0#, intervals - 1#)
            Case Else: count = intervals
        End Select
    End If
    If count < 0# Or count > GEOM_MAX_REBAR_COUNT Then RuleError index, "Число мест установки арматуры превышает предел счётчика программы. Уменьшите количество стержней или увеличьте их шаг."
    If count = 0# Then Exit Function
    ReDim stations(0 To CLng(count) - 1)
    If count = 1# And mRules(index).Mode = ByCount Then
        Select Case mRules(index).EndMode
            Case IncludeStart: stations(0) = first
            Case IncludeEnd: stations(0) = first + length
            Case Else: stations(0) = first + length / 2#
        End Select
    Else
        Select Case mRules(index).EndMode
            Case IncludeBoth
                step = length / (count - 1#): start = first
            Case ExcludeBoth
                step = length / (count + 1#): start = first + step
            Case IncludeStart
                step = length / count: start = first
            Case IncludeEnd
                step = length / count: start = first + step
        End Select
        For i = 0 To CLng(count) - 1: stations(i) = start + CDbl(i) * step: Next i
    End If
    StationsFor = stations
End Function

' Оба заданных конца требуют целого числа точных шагов; на полном замкнутом пути это также относится к включению одного конца без разрыва шага на замыкании.
Private Function RequiresIntegralFixedStep(ByVal index As Long, ByVal path As CGeometryPath) As Boolean
    RequiresIntegralFixedStep = (mRules(index).EndMode = IncludeBoth) Or _
        (path.IsClosed And (mRules(index).EndMode = IncludeStart Or mRules(index).EndMode = IncludeEnd))
End Function

' Вычисляет расстояния до стержней с постоянным шагом и заданным включением концов.
' Если длина не позволяет выдержать этот шаг в выбранном режиме, сообщает об ошибке.
Private Function FixedStations(ByVal index As Long, ByVal path As CGeometryPath, ByVal first As Double, ByVal length As Double) As Variant
    Dim count As Double, intervals As Double, start As Double, step As Double, stations() As Double, i As Long, kept As Long, station As Double
    step = mRules(index).Step: intervals = length / step
    If RequiresIntegralFixedStep(index, path) Then
        If Abs(intervals - Round(intervals)) > 0.0000000001 Then RuleError index, "Точный шаг несовместим с фиксированными обоими концами или полным замкнутым обходом."
        count = Round(intervals)
        If mRules(index).EndMode = IncludeBoth Then count = count + 1#
    Else
        count = Fix(intervals) + 1#
    End If
    If count > GEOM_MAX_REBAR_COUNT Then RuleError index, "Число мест установки арматуры превышает предел счётчика программы. Уменьшите количество стержней или увеличьте их шаг."
    If count = 0# Then Exit Function
    ReDim stations(0 To CLng(count) - 1)
    start = first
    Select Case mRules(index).EndMode
        Case IncludeEnd: start = first + length - (count - 1#) * step
        Case ExcludeBoth
            Select Case mRules(index).RemainderPlacement
                Case RemainderStart: start = first + length - (count - 1#) * step
                Case RemainderCenter: start = first + (length - (count - 1#) * step) / 2#
            End Select
    End Select
    For i = 0 To CLng(count) - 1
        station = start + CDbl(i) * step
        If Not ((mRules(index).EndMode = ExcludeBoth Or mRules(index).EndMode = IncludeEnd) And station <= first + GEOM_TOLERANCE) And _
            Not ((mRules(index).EndMode = ExcludeBoth Or mRules(index).EndMode = IncludeStart) And station >= first + length - GEOM_TOLERANCE) Then
            stations(kept) = station: kept = kept + 1
        End If
    Next i
    If kept = 0 Then Exit Function
    ReDim Preserve stations(0 To kept - 1): FixedStations = stations
End Function

' Проверяет параметры активного правила до построения позиций и зависимостей.
Private Sub ValidateOptions(ByVal index As Long)
    With mRules(index)
        If .Diameter <= 0# Then RuleError index, "Диаметр активного ряда должен быть положительным."
        If .StartSetback < 0# Or .EndSetback < 0# Then RuleError index, "Концевые отступы не могут быть отрицательными."
        If .EndMode < IncludeBoth Or .EndMode > IncludeEnd Then RuleError index, "Не распознан способ установки стержней на концах линии (EndMode)."
        If .PlacementCheck <> CheckCenter And .PlacementCheck <> CheckFullDisk Then RuleError index, "Не распознана проверка положения арматуры: принадлежность центров или всего поперечного сечения стержней бетону (PlacementCheck)."
        If .OverlapPolicy <> ReportOverlap And .OverlapPolicy <> RejectOverlap Then RuleError index, "Не распознано правило обработки пересечения поперечных сечений стержней."
        If .Publication <> PublishByRule And .Publication <> PublishPerSourcePosition Then RuleError index, "Не распознан порядок присвоения расчётных имён основным и связанным стержням (Publication)."
        If .PositionOrder <> AlongPath And .PositionOrder <> FromPathMiddle Then RuleError index, "Не распознан порядок нумерации стержней вдоль линии или от её середины (PositionOrder)."
        If .Coincidence < RejectCoincidence Or .Coincidence > KeepPublishedOwner Then RuleError index, "Не распознано правило обработки стержней с совпавшими центрами (Coincidence)."
        If .RemainderPlacement < RemainderStart Or .RemainderPlacement > RemainderCenter Then RuleError index, "Не распознано положение остатка длины при размещении с фиксированным шагом (RemainderPlacement)."
        If .Mode = ByMaxStep Or .Mode = ByFixedStep Then
            If .Step <= 0# Then RuleError index, "Шаг должен быть положительным."
        End If
        If .Mode = ProjectFrom Then
            If .SourceNormalSign <> -1# And .SourceNormalSign <> 0# And .SourceNormalSign <> 1# Then RuleError index, "Знак местной нормали должен быть -1, 0 или 1."
            If .SourceNormalSign = 0# And Sqr(.DirectionX * .DirectionX + .DirectionY * .DirectionY) <= GEOM_TOLERANCE Then RuleError index, "Не задано направление проекции."
        End If
    End With
End Sub

' Проверяет рабочий интервал после концевых отступов и совместимость с замкнутым обходом.
' Нулевой интервал допустим только для одного стержня ByCount или явно заданных позиций.
Private Sub ValidateWorkInterval(ByVal index As Long, ByVal path As CGeometryPath)
    Dim remaining As Double, pointPlacement As Boolean
    remaining = path.Length - mRules(index).StartSetback - mRules(index).EndSetback
    pointPlacement = (mRules(index).Mode = ByCount And mRules(index).Count = 1) Or mRules(index).Mode = ExplicitPositions
    If remaining < -GEOM_TOLERANCE Or (remaining <= GEOM_TOLERANCE And Not pointPlacement) Then ' Один стержень по количеству или явно заданная позиция допускают точечный участок. Допустимость каждой явной позиции проверяется в StationsFor.
        RuleError index, "После концевых отступов не осталось участка для размещения арматуры. " & _
            "Длина линии — " & CStr(path.Length) & " мм; отступ от начала — " & CStr(mRules(index).StartSetback) & _
            " мм; отступ от конца — " & CStr(mRules(index).EndSetback) & _
            " мм. Уменьшите концевые отступы так, чтобы их сумма была меньше длины линии.", _
            "WORK_INTERVAL_EMPTY", path.ID, path.Length, mRules(index).StartSetback + mRules(index).EndSetback
    End If
    If path.IsClosed Then
        If mRules(index).StartSetback <> 0# Or mRules(index).EndSetback <> 0# Then RuleError index, "При равномерном размещении по всей замкнутой линии концевые отступы должны быть нулевыми. Для ненулевых отступов задайте открытый участок."
    End If
End Sub

' Создаёт уникальный ID исходной позиции из правила и сквозного номера.
Private Function PositionID(ByVal index As Long, ByVal ordinal As Long, ByVal sourceID As String) As String
    Dim prefix As String
    prefix = mRules(index).SourcePrefix
    If Len(prefix) = 0 And Len(sourceID) > 0 And Len(mRules(index).SourceSuffix) > 0 Then
        PositionID = sourceID & mRules(index).SourceSuffix
    Else
        If Len(prefix) = 0 Then prefix = mRules(index).RuleID & "."
        PositionID = prefix & CStr(ordinal) & mRules(index).SourceSuffix
    End If
End Function

' Сохраняет подготовленную позицию стержня, её правило, исходную позицию
' и порядок добавления в общий список.
Private Sub AppendPosition(ByRef value As TRebarPosition)
    If mPositionCount >= GEOM_MAX_REBAR_COUNT Then RuleError value.RuleIndex, "Суммарное число мест установки арматуры превышает предел счётчика программы. Уменьшите количество стержней или увеличьте их шаг."
    If mPositionCount = mCapacity Then
        mCapacity = CLng(GeomMin(GEOM_MAX_REBAR_COUNT, GeomMax(16#, CDbl(mCapacity) * 2#)))
        ReDim Preserve mPositions(1 To mCapacity)
    End If
    mPositionCount = mPositionCount + 1: mPositions(mPositionCount) = value
End Sub

' Сортирует номера позиций по заданному порядку добавления стержней.
' Координаты и исходная последовательность расстановки при этом не меняются.
Private Sub SortOrder(ByRef order() As Long, ByRef scratch() As Long, ByVal first As Long, ByVal last As Long)
    If first >= last Then Exit Sub
    Dim middle As Long, a As Long, b As Long, out As Long
    middle = first + (last - first) \ 2
    SortOrder order, scratch, first, middle: SortOrder order, scratch, middle + 1, last
    a = first: b = middle + 1
    For out = first To last
        If a > middle Then
            scratch(out) = order(b): b = b + 1
        ElseIf b > last Then
            scratch(out) = order(a): a = a + 1
        ElseIf BeforePosition(order(a), order(b)) Then
            scratch(out) = order(a): a = a + 1
        Else
            scratch(out) = order(b): b = b + 1
        End If
    Next out
    For out = first To last: order(out) = scratch(out): Next out
End Sub

' Определяет, какая из двух позиций должна добавить стержень раньше.
' При совпадении координат это позволяет сохранить стержень по заданному правилу,
' а не выбирать его диаметр и материал случайно.
Private Function BeforePosition(ByVal a As Long, ByVal b As Long) As Boolean
    Dim keyA As Long, keyB As Long, ordinalA As Long, ordinalB As Long
    keyA = mPositions(a).RuleIndex: keyB = mPositions(b).RuleIndex
    ordinalA = mPositions(a).RootPublicationRank: ordinalB = mPositions(b).RootPublicationRank
    If mRules(keyA).Publication = PublishPerSourcePosition Then keyA = mPositions(a).RootRuleIndex
    If mRules(keyB).Publication = PublishPerSourcePosition Then keyB = mPositions(b).RootRuleIndex
    If keyA <> keyB Then BeforePosition = (keyA < keyB): Exit Function
    If ordinalA <> ordinalB Then BeforePosition = (ordinalA < ordinalB): Exit Function
    If mPositions(a).RuleIndex <> mPositions(b).RuleIndex Then BeforePosition = (mPositions(a).RuleIndex < mPositions(b).RuleIndex): Exit Function
    BeforePosition = (a <= b) ' При равных ключах стабильный порядок исходных позиций определяет первого владельца.
End Function

' Сохраняет диагностическую запись; подробные образцы ограничены по памяти, итоговые счётчики
' остаются полными.
Private Sub AddDiagnostic(ByVal index As Long, ByVal code As String, ByVal relatedEntityID As String, ByVal reason As String, _
        Optional ByVal detail1 As Double = 0#, Optional ByVal detail2 As Double = 0#)
    Dim entityID As String, reference As String
    If index > 0 Then
        entityID = mRules(index).RuleID: reference = mRules(index).SourceReference
        If Len(reference) = 0 Then reference = mDefinition.SourceReference(entityID)
    End If
    Set mDiagnostic = New CGeometryDiagnostic
    mDiagnostic.Initialize code, entityID, relatedEntityID, reason, detail1, detail2, reference
    If mDiagnostics Is Nothing Then Set mDiagnostics = New Collection
    If mDiagnostics.Count < DIAGNOSTIC_RECORD_LIMIT Then mDiagnostics.Add mDiagnostic Else mDiagnosticRecordsOmitted = mDiagnosticRecordsOmitted + 1
End Sub

' Выдаёт ошибку арматурного правила с ID и исходным адресом параметра.
Private Sub RuleError(ByVal index As Long, ByVal reason As String, Optional ByVal code As String = "RULE_INVALID", _
        Optional ByVal relatedEntityID As String = "", Optional ByVal detail1 As Double = 0#, Optional ByVal detail2 As Double = 0#)
    If code = "COINCIDENCE" Or code = "COINCIDENCE_CONFLICT" Or code = "BAR_OVERLAP" Or code = "ROW_OFFSET_COINCIDENCE" Then mConflicts = mConflicts + 1
    AddDiagnostic index, code, relatedEntityID, reason, detail1, detail2
    Dim reference As String, context As String, entityCaption As String
    If index > 0 Then
        reference = mDiagnostic.SourceReference
        entityCaption = "Ряд '" & mRules(index).RuleID & "':"
        context = mRules(index).Comment
    End If
    SectionRecipeError "CRebarRuleExecutor", reason, reference, context, entityCaption
End Sub

' Запрещает дальнейшую расстановку и изменение результатов после успешной сборки.
' Координаты и диагностика остаются доступными для чтения без копирования массивов.
Friend Sub Freeze()
    If Not mBuilt Then RuleError 0, "Нет успешно исполненных правил для публикации."
    mReadOnly = True
End Sub

' Показывает, защищены ли данные исполненных правил от дальнейших изменений.
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = mReadOnly
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Копирует исходную позицию и правило для независимого пересчёта формулы смещения.
Friend Sub TestCopyTransfer(ByVal positionIndex As Long, ByRef source As TRebarPosition, ByRef rule As TRebarRule)
    source = mPositions(mPositions(positionIndex).SourcePositionIndex)
    rule = mRules(mPositions(positionIndex).RuleIndex)
End Sub
