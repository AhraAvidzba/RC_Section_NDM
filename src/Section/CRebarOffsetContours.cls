VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRebarOffsetContours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Общие проверочные контуры второго и третьего рядов арматуры
' ==========================================================================
' Собирает один контур на ряд и бетонную границу с индивидуальными as и
' полными нормальными смещениями. Принадлежность задают явные метаданные;
' имена правил, порядок сторон и названия типов сечений не анализируются.
' Не переставляет стержни: проверяет уже вычисленную нормальную позицию.

Private Type TBoundaryPart
    Edge As TRegionEdge ' Исходный участок фигуры, к которому однозначно относится линия Main.
    LinkIndex As Long ' Связь Main/R2/R3 с бетонной границей и исходной фигурой.
    Distance As Double ' Заданный as со знаком левой нормали исходного участка, мм.
    Direction As Double ' Направление линии Main относительно исходного участка: +1 или -1.
    Baseline As Boolean ' Участок ограничивает исходную линию и при выключенном Main; дополнительных стержней не задаёт.
End Type
Private Type TPreparedContour
    Edges() As TRegionEdge ' Подготовленные границы области; не копируются заново для каждого стержня.
    Count As Long ' Число подготовленных отрезков и дуг.
End Type
Private mDefinition As CSectionDefinition ' Завершённое описание геометрии и правил текущей сборки.
Private mRules() As TRebarRule ' Независимые копии правил после включения проверок семейства.
Private mLinks() As TRebarContourLink, mLinkCount As Long ' Явные связи семейств с исходной геометрией и их количество.
Private mLoops() As TSectionLoop, mLoopCount As Long ' Наружные границы и отверстия, включая принадлежность отверстий.
Private mLoopIndex As Object ' Номера бетонных границ по точному ID без линейного поиска для каждого стержня.
Private mParts() As TBoundaryPart, mPartCount As Long ' Исходные участки Main, заданные as и направления их обхода.
Private mPartCapacity As Long ' Выделенная ёмкость массива участков; растёт с удвоением.
Private mRuleIndex As Object ' Точное соответствие ID правила его номеру, строится один раз.
Private mRowRules() As Long ' Номера правил R2/R3 каждого семейства без повторного поиска по ID.
Private mPartsByLink() As Collection ' Индексы участков каждого семейства для проверки его соседей.
Private mPartsByLoop As Object ' Индексы участков каждой бетонной границы; другие отверстия не просматриваются.
Private mRuleLink() As Long, mRuleRow() As Long ' Явная связь дополнительного правила с семейством и номером ряда 2/3.
Private mRegions As Object ' Номера подготовленных областей по ряду и наружной границе; готовятся один раз.
Private mPrepared() As TPreparedContour, mPreparedCount As Long ' Массивы проверенных границ и количество подготовленных областей.
Private mAdjacency As Object ' Соседство явно связанных семейств по общим концам исходных участков.
Private mQuery As CSectionGeometryQuery ' Точные операции с областями, включая отверстия и самопересечения.
Private mOffset As CGeometryVariableOffset ' Отдельный построитель индивидуального контура без изменения прежних Offset.
Private mReferences As String ' Фактические адреса полей всех участвующих семейств для ошибок построения.

' Подготавливает однозначные связи; семейства без нового режима не меняет.
Friend Sub Initialize(ByVal definition As CSectionDefinition, ByRef rules() As TRebarRule)
    On Error GoTo InvalidMetadata
    Dim i As Long, j As Long, main As Long, axis As CGeometryPath, source As CGeometryPath
    Dim paths As Variant, pathID As Variant, inward As Double, area As Double, edge As TRegionEdge, reference As String
    Set mDefinition = definition: Set mQuery = New CSectionGeometryQuery: Set mOffset = New CGeometryVariableOffset
    Set mRegions = CreateObject("Scripting.Dictionary")
    Set mAdjacency = CreateObject("Scripting.Dictionary")
    Set mRuleIndex = CreateObject("Scripting.Dictionary"): mRuleIndex.CompareMode = vbBinaryCompare
    Set mPartsByLoop = CreateObject("Scripting.Dictionary"): mPartsByLoop.CompareMode = vbBinaryCompare
    Set mLoopIndex = CreateObject("Scripting.Dictionary"): mLoopIndex.CompareMode = vbBinaryCompare
    ReDim mRules(1 To definition.RuleCount): ReDim mRuleLink(1 To definition.RuleCount): ReDim mRuleRow(1 To definition.RuleCount)
    For i = 1 To definition.RuleCount: mRules(i) = rules(i): mRuleIndex.Add rules(i).RuleID, i: Next i
    mLoopCount = definition.LoopCount: ReDim mLoops(1 To mLoopCount)
    ReDim mPrepared(1 To 2 * mLoopCount)
    For i = 1 To mLoopCount: definition.CopyLoop i, mLoops(i): mLoopIndex.Add mLoops(i).LoopID, i: Next i
    mLinkCount = definition.RebarContourCount
    If mLinkCount > 0 Then
        ReDim mLinks(1 To mLinkCount): ReDim mRowRules(1 To mLinkCount, 2 To 3): ReDim mPartsByLink(1 To mLinkCount)
    End If
    For i = 1 To mLinkCount
        Set mPartsByLink(i) = New Collection
        definition.CopyRebarContour i, mLinks(i)
        main = RuleNumber(mLinks(i).MainRuleID)
        reference = mLinks(i).SourceReference
        If main > 0 Then reference = mRules(main).SourceReference
        If main = 0 And Len(mLinks(i).MainRuleID) > 0 Then Fail "В связи с контуром не найден основной ряд.", reference
        mReferences = mReferences & reference & "; "
        If mLinks(i).AxisDistance < 0# Then Fail "В связи с контуром задан отрицательный as.", reference
        If mLinks(i).NormalSign <> 0# And Abs(mLinks(i).NormalSign) <> 1# Then _
            Fail "Направление от Main в бетон должно быть -1 или 1; неизвестное направление задаётся нулём.", reference
        j = LoopIndex(mLinks(i).LoopID)
        If j = 0 Then Fail "В связи арматуры не найдена бетонная граница.", reference
        If main > 0 Then
            BindRow i, main, 2, mLinks(i).Row2RuleID
            BindRow i, main, 3, mLinks(i).Row3RuleID
        End If
        Set source = definition.Path(mLinks(i).SourcePathID)
        If Not source.IsClosed Then Fail "Исходная фигура для заданного as должна быть замкнутой.", reference
        If IsArray(mLinks(i).SourceSegments) Then
            For Each pathID In mLinks(i).SourceSegments
                If Not source.HasSegmentID(CStr(pathID)) Then Fail "В связи с контуром указан несуществующий исходный участок.", reference
            Next pathID
        End If
        area = 0#
        For j = 1 To source.SegmentCount
            source.CopySegment j, edge: area = area + mOffset.EdgeArea(edge)
        Next j
        inward = Sgn(area) * mLinks(i).AxisDistance ' Знак обхода переводит положительный as в смещение по левой нормали исходного контура.
        j = LoopIndex(mLinks(i).LoopID)
        If mLoops(j).Role = LoopOpening Then inward = -inward ' У отверстия as откладывается из пустоты в бетон.
        If main = 0 Then
            If Not IsArray(mLinks(i).SourceSegments) Then Fail "Для базового ограничения явно укажите исходные участки.", reference
            For j = 1 To source.SegmentCount
                If SelectedSource(i, source.SegmentID(j)) Then
                    AppendPart i
                    source.CopySegment j, mParts(mPartCount).Edge
                    mParts(mPartCount).LinkIndex = i: mParts(mPartCount).Distance = inward: mParts(mPartCount).Baseline = True
                End If
            Next j
        Else
            paths = definition.TargetPaths(mRules(main).Target)
            For Each pathID In paths
                Set axis = definition.Path(CStr(pathID))
                RegisterAxis i, axis, source, inward, reference
            Next pathID
        End If
    Next i
    For i = 1 To definition.RuleCount
        If rules(i).Enabled And rules(i).ValidateRowOffset Then
            If mRuleLink(i) = 0 Then Fail "Для проверки дополнительного ряда нужна явная связь Main с исходным контуром. " & _
                "Добавьте её через CSectionDefinitionBuilder.RebarContour; as и принадлежность ряда по имени не определяются.", rules(i).SourceReference
        End If
    Next i
    Exit Sub
InvalidMetadata:
    Dim reason As String, number As Long, owner As String
    reason = Err.Description: number = Err.Number: owner = Err.Source
    If Len(reference) > 0 And InStr(reason, reference) = 0 Then Fail reason, reference
    Err.Raise number, owner, reason
End Sub

' Проверяет координаты после нормального и до касательного смещения.
' Внутренние линии составного сечения допускаются внутри общей области.
Friend Function Accepts(ByVal ruleIndex As Long, ByVal x As Double, ByVal y As Double) As Boolean
    Dim link As Long, outer As Long, key As String, region As CConcreteRegion, index As Long
    link = mRuleLink(ruleIndex): outer = LoopIndex(mLinks(link).LoopID)
    If mLoops(outer).Role = LoopOpening Then outer = LoopIndex(mLoops(outer).ParentID)
    key = CStr(mRuleRow(ruleIndex)) & ":" & CStr(outer) ' Одна подготовленная область на ряд 2/3 и наружную границу, включая все её отверстия.
    If Not mRegions.Exists(key) Then
        Set region = BuildDomain(mRuleRow(ruleIndex), outer)
        mPreparedCount = mPreparedCount + 1
        region.CopyEdges mPrepared(mPreparedCount).Edges, mPrepared(mPreparedCount).Count
        mRegions.Add key, mPreparedCount
    End If
    index = CLng(mRegions(key))
    Accepts = mQuery.ContainsPreparedPoint(mPrepared(index).Edges, mPrepared(index).Count, x, y, True)
End Function

' У соседних граней разные семейства и общий конец исходных участков одной бетонной границы.
' Совпадающие участки одной прямой/окружности и разные отверстия соседями не считаются.
Friend Function AreAdjacent(ByVal firstRule As Long, ByVal lastRule As Long) As Boolean
    Dim a As Long, b As Long, i As Variant, j As Variant, key As String
    a = mRuleLink(firstRule): b = mRuleLink(lastRule)
    If a = 0 Or b = 0 Or a = b Then Exit Function
    If mLinks(a).LoopID <> mLinks(b).LoopID Then Exit Function
    key = CStr(GeomMin(a, b)) & ":" & CStr(GeomMax(a, b))
    If mAdjacency.Exists(key) Then AreAdjacent = mAdjacency(key): Exit Function
    For Each i In mPartsByLink(a)
        If Not mParts(i).Baseline Then
            For Each j In mPartsByLink(b)
                If Not mParts(j).Baseline Then
                    If Not SameSupport(mParts(i).Edge, mParts(j).Edge) Then
                        If ShareEndpoint(mParts(i).Edge, mParts(j).Edge) Then AreAdjacent = True: GoTo Found
                    End If
                End If
            Next j
        End If
    Next i
Found:
    mAdjacency.Add key, AreAdjacent
End Function

' Проверяет общий геометрический конец без разбора имён сторон и порядка их объявления.
Private Function ShareEndpoint(ByRef first As TRegionEdge, ByRef last As TRegionEdge) As Boolean
    ShareEndpoint = modGeometryPrimitives.SamePoint(first.X1, first.Y1, last.X1, last.Y1) Or _
        modGeometryPrimitives.SamePoint(first.X1, first.Y1, last.X2, last.Y2) Or _
        modGeometryPrimitives.SamePoint(first.X2, first.Y2, last.X1, last.Y1) Or _
        modGeometryPrimitives.SamePoint(first.X2, first.Y2, last.X2, last.Y2)
End Function

' Регистрирует номер дополнительного ряда, не разбирая ID и не предполагая порядок правил.
Private Sub BindRow(ByVal link As Long, ByVal main As Long, ByVal row As Long, ByVal ruleID As String)
    Dim index As Long
    index = RuleNumber(ruleID)
    If index = 0 Then Exit Sub
    If mRules(index).Mode <> DependentOffset Or mRules(index).SourceRuleID <> mRules(main).RuleID Then _
        Fail "Второй и третий ряды в связи с контуром должны независимо ссылаться на Main.", mRules(index).SourceReference
    If mRuleLink(index) <> 0 Then Fail "Дополнительный ряд связан с несколькими исходными контурами.", mRules(index).SourceReference
    mRuleLink(index) = link: mRuleRow(index) = row
    mRowRules(link, row) = index
    mReferences = mReferences & mRules(index).SourceReference & "; "
End Sub

' Создаёт место для участка с амортизированным линейным ростом массива и индексирует его принадлежность.
Private Sub AppendPart(ByVal link As Long)
    Dim items As Collection, loopID As String
    If mPartCount = mPartCapacity Then
        mPartCapacity = CLng(GeomMax(16#, 2# * mPartCapacity))
        ReDim Preserve mParts(1 To mPartCapacity)
    End If
    mPartCount = mPartCount + 1
    mPartsByLink(link).Add mPartCount
    loopID = mLinks(link).LoopID
    If Not mPartsByLoop.Exists(loopID) Then Set items = New Collection: mPartsByLoop.Add loopID, items
    Set items = mPartsByLoop(loopID): items.Add mPartCount
End Sub

' Возвращает номер по точному ID; отсутствующий дополнительный ряд даёт ноль.
Private Function RuleNumber(ByVal ruleID As String) As Long
    If mRuleIndex.Exists(ruleID) Then RuleNumber = CLng(mRuleIndex(ruleID))
End Function

' Находит исходный участок по точному геометрическому соответствию и переданному as.
' Неоднозначность является ошибкой; ближайшая сторона вместо исходной не выбирается.
Private Sub RegisterAxis(ByVal link As Long, ByVal axis As CGeometryPath, ByVal source As CGeometryPath, _
        ByVal distance As Double, ByVal sourceReference As String)
    Dim i As Long, j As Long, matches As Long, found As Long, direction As Double, tx As Double, ty As Double, sx As Double, sy As Double
    Dim part As TRegionEdge, original As TRegionEdge, shifted As TRegionEdge, radius As Double, foundDirection As Double
    For i = 1 To axis.SegmentCount
        axis.CopySegment i, part: matches = 0
        For j = 1 To source.SegmentCount
            source.CopySegment j, original
            If Not SelectedSource(link, source.SegmentID(j)) Then GoTo NextSource
            If (part.Sweep = 0#) = (original.Sweep = 0#) Then
                If original.Sweep <> 0# Then
                    radius = Sqr((original.X1 - original.CenterX) ^ 2 + (original.Y1 - original.CenterY) ^ 2)
                    If radius - Sgn(original.Sweep) * distance <= GEOM_BOUNDARY_TOLERANCE Then GoTo NextSource
                End If
                mOffset.ShiftEdge original, distance, shifted, sourceReference
                If SameSupport(part, shifted) And MidpointOn(part, shifted) Then
                    modGeometryPrimitives.EdgeTangent part, 0.5, tx, ty
                    modGeometryPrimitives.EdgeTangent shifted, 0.5, sx, sy
                    If part.Sweep = 0# Then direction = Sgn(tx * sx + ty * sy) Else direction = Sgn(part.Sweep * shifted.Sweep)
                    If mLinks(link).NormalSign <> 0# And distance <> 0# Then
                        If Sgn(distance) * direction <> mLinks(link).NormalSign Then GoTo NextSource
                    End If
                    matches = matches + 1: found = j: foundDirection = direction
                End If
            End If
NextSource:
        Next j
        If matches <> 1 Then Fail "Нельзя однозначно связать линию Main с исходной фигурой при заданном as. " & _
            "Проверьте исходный контур, переданное расстояние и связь RebarContour.", sourceReference
        source.CopySegment found, original
        AppendPart link
        mParts(mPartCount).Edge = original: mParts(mPartCount).LinkIndex = link
        mParts(mPartCount).Distance = distance: mParts(mPartCount).Direction = foundDirection
    Next i
    For j = 1 To source.SegmentCount ' Исчезнувшая при as дуга сохраняет явную привязку исходного сегмента. Она нужна для удаления скругления из общего контура, даже если Main уже не содержит дуги.
        source.CopySegment j, original
        If original.Sweep <> 0# And axis.HasCollapsedSegmentID(source.SegmentID(j)) Then
            If SelectedSource(link, source.SegmentID(j)) Then
                radius = Sqr((original.X1 - original.CenterX) ^ 2 + (original.Y1 - original.CenterY) ^ 2)
                If radius - Sgn(original.Sweep) * distance <= GEOM_BOUNDARY_TOLERANCE Then
                    If mLinks(link).NormalSign = 0# Or distance = 0# Then _
                        Fail "Для исчезнувшего скругления явно передайте направление Main в RebarContour.normalSign.", sourceReference
                    AppendPart link
                    mParts(mPartCount).Edge = original: mParts(mPartCount).LinkIndex = link
                    mParts(mPartCount).Distance = distance: mParts(mPartCount).Direction = mLinks(link).NormalSign * Sgn(distance)
                End If
            End If
        End If
    Next j
End Sub

' Строит наружную проверочную границу и все принадлежащие ей отверстия одного ряда.
' Вложенность проверяется до получения пригодной области для фильтрации позиций.
Private Function BuildDomain(ByVal row As Long, ByVal outer As Long) As CConcreteRegion
    Dim domain As CConcreteRegion, opening As CConcreteRegion, i As Long, reason As String
    On Error GoTo InvalidDomain
    Set domain = BuildLoop(row, outer)
    For i = 1 To mLoopCount
        If mLoops(i).Role = LoopOpening And mLoops(i).ParentID = mLoops(outer).LoopID Then
            Set opening = BuildLoop(row, i)
            mQuery.ValidateContainedRegion domain, opening
            Set domain = mQuery.SubtractContainedRegion(domain, opening)
        End If
    Next i
    Set BuildDomain = domain
    Exit Function
InvalidDomain:
    reason = Err.Description
    If Err.Source = "CGeometryVariableOffset" Then Err.Raise Err.Number, Err.Source, _
        "Не удалось построить допустимый контур для " & CStr(row) & "-го ряда арматуры." & vbCrLf & vbCrLf & reason ' Построитель уже указал адреса конкретных участков; общая обёртка со всеми семействами повторяла их и скрывала место ошибки.
    Fail "Не удалось построить допустимый контур для " & CStr(row) & "-го ряда арматуры: " & reason, mReferences
End Function

' Разделяет только прямые по границам исходных фигур. Это позволяет задавать
' разные отступы у совпадающих по направлению сторон двух прямоугольников.
Private Function BuildLoop(ByVal row As Long, ByVal loopIndexValue As Long) As CConcreteRegion
    Dim path As CGeometryPath, edge As TRegionEdge, part As TRegionEdge, edges() As TRegionEdge, distances() As Double
    Dim cuts() As Double, cutCount As Long, i As Long, j As Variant, k As Long, count As Long, parameter As Double, capacity As Long
    Dim references() As String, reference As String, sourceValues As Variant
    Dim parts As Collection
    Set parts = New Collection
    If mPartsByLoop.Exists(mLoops(loopIndexValue).LoopID) Then Set parts = mPartsByLoop(mLoops(loopIndexValue).LoopID)
    Set path = mDefinition.Path(mLoops(loopIndexValue).PathID)
    For i = 1 To path.SegmentCount
        path.CopySegment i, edge
        Erase cuts: cutCount = 0: AddCut cuts, cutCount, 0#: AddCut cuts, cutCount, 1#
        For Each j In parts
            If mLinks(mParts(j).LinkIndex).LoopID = mLoops(loopIndexValue).LoopID And SameSupport(edge, mParts(j).Edge) Then
                If modGeometryPrimitives.PointParameter(edge, mParts(j).Edge.X1, mParts(j).Edge.Y1, parameter) Then AddCut cuts, cutCount, parameter
                If modGeometryPrimitives.PointParameter(edge, mParts(j).Edge.X2, mParts(j).Edge.Y2, parameter) Then AddCut cuts, cutCount, parameter
            End If
        Next j
        For k = 1 To cutCount - 1
            modGeometryPrimitives.SubEdge edge, cuts(k), cuts(k + 1), part
            If count = capacity Then
                capacity = CLng(GeomMax(16#, 2# * capacity))
                ReDim Preserve edges(1 To capacity): ReDim Preserve distances(1 To capacity): ReDim Preserve references(1 To capacity)
            End If
            count = count + 1
            reference = vbNullString
            edges(count) = part: distances(count) = PartDistance(part, row, mLoops(loopIndexValue).LoopID, reference)
            references(count) = reference
        Next k
    Next i
    sourceValues = references
    Set BuildLoop = mOffset.Build(edges, distances, count, mReferences, sourceValues)
End Function

' Для участка без семейства сохраняет бетонную границу, не придумывая несуществующий as.
' У семейства без дополнительного ряда остаётся только явно заданный as Main.
Private Function PartDistance(ByRef edge As TRegionEdge, ByVal row As Long, ByVal loopID As String, ByRef sourceReference As String) As Double
    Dim i As Variant, link As Long, rule As Long, value As Double, found As Boolean, direction As Double
    Dim baseline As Double, hasBaseline As Boolean, mainBase As Double, thisBase As Double
    Dim tx As Double, ty As Double, sx As Double, sy As Double
    If Not mPartsByLoop.Exists(loopID) Then Exit Function
    For Each i In mPartsByLoop(loopID)
        link = mParts(i).LinkIndex
        If mLinks(link).LoopID = loopID And SameSupport(edge, mParts(i).Edge) Then
            If MidpointOn(edge, mParts(i).Edge) Then
                sourceReference = sourceReference & mLinks(link).SourceReference & "; "
                value = mParts(i).Distance
                If Not mParts(i).Baseline Then
                    rule = mRowRules(link, row)
                    If rule > 0 Then
                        If mRules(rule).Enabled Then
                            value = value + mRules(rule).NormalOffset * mParts(i).Direction ' Направление Main относительно исходного участка определяет знак дополнительного отступа.
                            sourceReference = sourceReference & mRules(rule).SourceReference & "; "
                        End If
                    End If
                End If
                modGeometryPrimitives.EdgeTangent edge, 0.5, tx, ty
                modGeometryPrimitives.EdgeTangent mParts(i).Edge, 0.5, sx, sy
                If edge.Sweep = 0# Then direction = Sgn(tx * sx + ty * sy) Else direction = Sgn(edge.Sweep * mParts(i).Edge.Sweep)
                value = value * direction
                thisBase = mParts(i).Distance * direction
                If mParts(i).Baseline Then
                    If hasBaseline And Abs(baseline - value) > GEOM_BOUNDARY_TOLERANCE Then Fail "Для исходного участка заданы разные базовые as.", mReferences
                    baseline = value: hasBaseline = True
                Else
                    If found And Abs(value - PartDistance) > GEOM_BOUNDARY_TOLERANCE Then _
                        Fail "Один исходный участок получил разные отступы от нескольких семейств. Уточните явные связи Main с контуром.", mReferences
                    PartDistance = value: mainBase = thisBase: found = True
                End If
                If found And hasBaseline Then
                    If Abs(mainBase - baseline) > GEOM_BOUNDARY_TOLERANCE Then Fail "as основного ряда не совпадает с явно заданным ограничителем.", mReferences
                End If
            End If
        End If
    Next i
    If Not found And hasBaseline Then PartDistance = baseline ' Отключённый ряд не отменяет явно заданный геометрический ограничитель.
End Function

' Проверяет совпадение несущей прямой либо окружности, без выбора ближайшей стороны.
Private Function SameSupport(ByRef first As TRegionEdge, ByRef second As TRegionEdge) As Boolean
    Dim ax As Double, ay As Double, bx As Double, by As Double, length As Double, other As Double
    If (first.Sweep = 0#) <> (second.Sweep = 0#) Then Exit Function
    If first.Sweep = 0# Then
        ax = first.X2 - first.X1: ay = first.Y2 - first.Y1: bx = second.X2 - second.X1: by = second.Y2 - second.Y1
        length = Sqr(ax * ax + ay * ay): other = Sqr(bx * bx + by * by)
        SameSupport = Abs(ax * by - ay * bx) <= GEOM_PARAMETER_TOLERANCE * length * other And _
            Abs(ax * (second.Y1 - first.Y1) - ay * (second.X1 - first.X1)) <= GEOM_BOUNDARY_TOLERANCE * length
    Else
        length = Sqr((first.X1 - first.CenterX) ^ 2 + (first.Y1 - first.CenterY) ^ 2)
        other = Sqr((second.X1 - second.CenterX) ^ 2 + (second.Y1 - second.CenterY) ^ 2)
        SameSupport = modGeometryPrimitives.SamePoint(first.CenterX, first.CenterY, second.CenterX, second.CenterY) And Abs(length - other) <= GEOM_BOUNDARY_TOLERANCE
    End If
End Function

' Проверяет, лежит ли середина первого участка на втором, а не только на общей несущей прямой/окружности.
Private Function MidpointOn(ByRef first As TRegionEdge, ByRef second As TRegionEdge) As Boolean
    Dim x As Double, y As Double, parameter As Double
    modGeometryPrimitives.EdgePoint first, 0.5, x, y
    MidpointOn = modGeometryPrimitives.PointParameter(second, x, y, parameter)
End Function

' Добавляет уникальный параметр разбиения в возрастающем порядке.
Private Sub AddCut(ByRef values() As Double, ByRef count As Long, ByVal value As Double)
    Dim i As Long, j As Long
    value = GeomMax(0#, GeomMin(1#, value))
    For i = 1 To count
        If Abs(values(i) - value) <= GEOM_PARAMETER_TOLERANCE Then Exit Sub
        If values(i) > value Then Exit For
    Next i
    count = count + 1: ReDim Preserve values(1 To count)
    For j = count To i + 1 Step -1: values(j) = values(j - 1): Next j
    values(i) = value
End Sub

' При заданном массиве выбирает только перечисленные исходные участки.
' Если список не передан (Empty), допускает все сегменты; точное соответствие затем проверяет RegisterAxis.
Private Function SelectedSource(ByVal link As Long, ByVal segmentID As String) As Boolean
    Dim names As Variant, name As Variant
    names = mLinks(link).SourceSegments
    If IsEmpty(names) Then SelectedSource = True: Exit Function
    If Not IsArray(names) Then Fail "Явные исходные участки RebarContour задаются массивом имён.", mReferences
    For Each name In names
        If CStr(name) = segmentID Then SelectedSource = True: Exit Function
    Next name
End Function

' Находит только явно указанный LoopID; наружная роль не выводится из имени.
Private Function LoopIndex(ByVal loopID As String) As Long
    If mLoopIndex.Exists(loopID) Then LoopIndex = CLng(mLoopIndex(loopID))
End Function

' Сообщает адресную ошибку метаданных или геометрии до успешного завершения сборки.
Private Sub Fail(ByVal reason As String, ByVal sourceReference As String)
    SectionRecipeError "CRebarOffsetContours", reason, sourceReference
End Sub
