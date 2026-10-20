VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNDMResultsWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Запись рассчитанного состояния сечения на лист Results
' ==========================================================================
' Получает готовую модель и результаты пакетного расчёта, переводит величины
' в OUTPUT-единицы и сохраняет согласованный набор данных последнего расчёта.
' Записывает геометрию и состояния StrengthState, CapacityState, CrackedState
' и другие. Подробную таблицу прочности оформляет CStrengthSummaryWriter.
' Текущие Inputs не перечитываются: сохраняется именно рассчитанное состояние.

Private Const ELEMENT_RESULT_COLUMNS As Long = 9 ' Столбцы результата каждого элемента: запуск, сочетание, профиль, вид НДС, имя элемента, материал, деформация, напряжение и состояние работы.
Private Const GEOMETRY_COLUMNS As Long = 15 ' Только расчетные бетонные элементы и арматура.
Private Const CONTOUR_COLUMNS As Long = 17 ' Компактные аналитические исходные границы, формат v1.
Private Const SECTION_PROPERTY_COLUMNS As Long = 6 ' Общие свойства сечения и LC-свойства уровня всего сечения.
Private Const ANNOTATION_COLUMNS As Long = 22 ' Аннотации и принадлежность готовой области трещин; старые поля 1..13 сохранены.
Private Const MATERIAL_DIAGRAM_COLUMNS As Long = 11 ' Список точек всех различных диаграмм, фактически использованных в расчётах.
Private Const COMMON_SECTION_PROPERTY_ROWS As Long = 72 ' Строки общих свойств сечения до LC-зависимых блоков.
Private Const LC_SECTION_PROPERTY_ROWS As Long = 23 ' Число строк общих свойств одного сочетания без дополнительных данных отдельных НДС.
Private Const STATE_SECTION_PROPERTY_ROWS As Long = 13 ' Число строк дополнительных данных одного сохранённого расчётного состояния.
Private Const SNAPSHOT_TITLE_OFFSET As Long = 2 ' Строка общего синего заголовка расположена на две строки выше начал нижних таблиц результатов.

' Возвращает фактическое число столбцов выбранной таблицы сохраняемых результатов.
Public Function OutputColumnCount(ByVal rangeName As String) As Long
    Select Case LCase$(Trim$(rangeName))
        Case "rngndmelementresults"
            OutputColumnCount = ELEMENT_RESULT_COLUMNS
        Case "rngndmsectiongeometry"
            OutputColumnCount = GEOMETRY_COLUMNS
        Case "rngndmsectioncontours"
            OutputColumnCount = CONTOUR_COLUMNS
        Case "rngndmsectionproperties"
            OutputColumnCount = SECTION_PROPERTY_COLUMNS
        Case "rngndmmaterialdiagrams"
            OutputColumnCount = MATERIAL_DIAGRAM_COLUMNS
        Case "rngndmsectionannotations"
            OutputColumnCount = ANNOTATION_COLUMNS
        Case Else
            Err.Raise vbObjectError + 4460, "CNDMResultsWriter.OutputColumnCount", _
                "Неизвестный блок Results: " & rangeName
    End Select
End Function

' Оценивает наибольшее число строк таблицы элементов rngNDMElementResults до начала расчёта.
Public Function EstimatedElementResultRows(ByVal elementCount As Long, ByVal namedStateCount As Long) As Long
    If elementCount < 0 Then elementCount = 0
    If namedStateCount < 0 Then namedStateCount = 0
    EstimatedElementResultRows = 1 + elementCount * namedStateCount
End Function

' До расчёта оценивает размер rngNDMSectionProperties по наибольшему числу сохраняемых НДС.
Public Function EstimatedSectionPropertyRows(ByVal combinationCount As Long, ByVal namedStateCount As Long) As Long
    If combinationCount < 0 Then combinationCount = 0
    If namedStateCount < 0 Then namedStateCount = 0
    EstimatedSectionPropertyRows = 1 + COMMON_SECTION_PROPERTY_ROWS + _
        combinationCount * LC_SECTION_PROPERTY_ROWS + namedStateCount * STATE_SECTION_PROPERTY_ROWS
End Function

' ------------------------------
' Полная запись Results
' ------------------------------
' Очищает прежние расчётные данные и записывает согласованные таблицы:
' геометрию, результаты элементов, свойства сечения, диаграммы материалов и аннотации.
Public Sub WriteResults(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal snapshot As CSectionPropertiesCalculator, ByVal batch As CBatchSectionCalculator, _
        Optional ByVal units As CUnitSystem = Nothing)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4450, "CNDMResultsWriter", "Книга Excel не передана."
    If section Is Nothing Then Err.Raise vbObjectError + 4451, "CNDMResultsWriter", "Модель сечения не передана."
    If snapshot Is Nothing Then Err.Raise vbObjectError + 4452, "CNDMResultsWriter", "Готовые характеристики снимка не переданы."
    If batch Is Nothing Then Err.Raise vbObjectError + 4454, "CNDMResultsWriter", "Пакетный расчетчик не передан."

    snapshot.ValidateSnapshot section, True
    ClearResults workbook

    Dim runID As Date
    runID = Now

    WriteSnapshotTitle workbook
    WriteSectionGeometry workbook, section, runID, units
    WriteSectionContours workbook, section.Contours, runID, units
    WriteElementResults workbook, section, batch, runID, units
    WriteSectionProperties workbook, section, snapshot, batch, runID, units
    WriteMaterialDiagrams workbook, batch, runID, units
    WriteSectionAnnotations workbook, section, runID, units, batch
End Sub

' После импорта AutoCAD записывает геометрию без расчёта сочетаний.
' По этим данным сразу строится схема, а следующий расчёт
' использует ту же сохранённую геометрию Results.
Public Sub WriteGeometryPreview(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal snapshot As CSectionPropertiesCalculator, _
        Optional ByVal units As CUnitSystem = Nothing)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4450, "CNDMResultsWriter", "Книга Excel не передана."
    If section Is Nothing Then Err.Raise vbObjectError + 4451, "CNDMResultsWriter", "Модель сечения не передана."
    If snapshot Is Nothing Then Err.Raise vbObjectError + 4452, "CNDMResultsWriter", "Для предварительной схемы не переданы подготовленные характеристики сечения."
    snapshot.ValidateSnapshot section

    ClearResults workbook

    Dim runID As Date
    runID = Now

    WriteSnapshotTitle workbook
    WriteSectionGeometry workbook, section, runID, units
    WriteSectionContours workbook, section.Contours, runID, units
    WriteEmptyElementResults workbook, runID, units
    WriteSectionPropertiesPreview workbook, section, snapshot, runID, units
    WriteSectionAnnotations workbook, section, runID, units
End Sub

' Очищает только таблицы сохранённых расчётных данных перед новой записью.
Public Sub ClearResults(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4455, "CNDMResultsWriter", "Книга Excel не передана."

    ClearSnapshotTitle workbook
    ClearNamedRegion workbook, "rngNDMElementResults"
    ClearNamedRegion workbook, "rngNDMSectionGeometry"
    ClearNamedRegion workbook, "rngNDMSectionContours"
    ClearNamedRegion workbook, "rngNDMSectionProperties"
    ClearNamedRegion workbook, "rngNDMSectionAnnotations"
    ClearNamedRegion workbook, "rngNDMMaterialDiagrams"
    Dim emptyContours As CSectionContours: Set emptyContours = New CSectionContours
    WriteSectionContours workbook, emptyContours, Empty
End Sub

' Очищает прямоугольник одной таблицы результатов. Ширину берёт из её структуры,
' высоту — по заполненному первому столбцу. CurrentRegion не используется,
' поскольку из-за форматирования может захватить соседние таблицы.
Private Sub ClearNamedRegion(ByVal workbook As Object, ByVal rangeName As String)
    On Error GoTo MissingRange
    Dim anchor As Object
    Set anchor = workbook.Names.Item(rangeName).RefersToRange

    Dim columnCount As Long
    columnCount = OutputColumnCount(rangeName)

    Dim rowCount As Long
    rowCount = ExistingAnchoredRowCount(anchor)
    If rowCount < 1 Then rowCount = 1

    Dim tableRange As Object
    Set tableRange = anchor.Resize(rowCount, columnCount)
    tableRange.UnMerge
    tableRange.Clear

    If anchor.Row > 1 Then
        Dim titleRange As Object
        Set titleRange = anchor.Offset(-1, 0).Resize(1, columnCount)
        titleRange.UnMerge
        titleRange.Clear
    End If
MissingRange:
End Sub

' Определяет высоту записанной таблицы по первому столбцу.
' Он заполнен в заголовке и каждой строке данных, поэтому
' первая пустая ячейка означает конец таблицы.
Private Function ExistingAnchoredRowCount(ByVal anchor As Object) As Long
    ExistingAnchoredRowCount = SnapshotAnchoredExtent(anchor, False)
End Function

' Записывает общий заголовок над нижними расчётными таблицами Results.
' Соседние пользовательские ячейки той же строки не затрагиваются.
Private Sub WriteSnapshotTitle(ByVal workbook As Object)
    On Error GoTo MissingRange
    Dim title As Object
    Set title = SnapshotTitleRange(workbook)
    If title Is Nothing Then Exit Sub
    title.Interior.Color = RGB(189, 215, 238)
    title.RowHeight = 15
    With title.Resize(1, 1)
        .Value2 = "Расчетный снимок Results: элементы, геометрия, свойства сечения, диаграммы материалов и аннотации"
        .Font.Name = "Arial"
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(255, 0, 0)
        .HorizontalAlignment = -4131
        .VerticalAlignment = -4108
        .WrapText = False
    End With
MissingRange:
End Sub

' Очищает общий заголовок расчётных таблиц, сохраняя пользовательские данные
' и оформление справа от них.
Private Sub ClearSnapshotTitle(ByVal workbook As Object)
    On Error GoTo MissingRange
    Dim title As Object
    Set title = SnapshotTitleRange(workbook)
    If title Is Nothing Then Exit Sub
    title.Clear
    title.RowHeight = 15
MissingRange:
End Sub

' Находит границы общего заголовка по именованным началам и размерам пяти таблиц.
' При перемещении таблиц заголовок следует за ними и не расширяется на весь лист.
Private Function SnapshotTitleRange(ByVal workbook As Object) As Object
    Dim anchor As Object, block As Object, name As Variant
    Set anchor = workbook.Names.Item("rngNDMElementResults").RefersToRange
    If anchor.Row <= SNAPSHOT_TITLE_OFFSET Then Exit Function

    Dim rightColumn As Long, blockRightColumn As Long
    rightColumn = anchor.Column
    For Each name In Array("rngNDMElementResults", "rngNDMSectionGeometry", _
            "rngNDMSectionContours", "rngNDMSectionProperties", "rngNDMMaterialDiagrams", "rngNDMSectionAnnotations")
        Set block = workbook.Names.Item(CStr(name)).RefersToRange
        blockRightColumn = block.Column + OutputColumnCount(CStr(name)) - 1
        If blockRightColumn > rightColumn Then rightColumn = blockRightColumn
    Next name
    Set SnapshotTitleRange = anchor.Offset(-SNAPSHOT_TITLE_OFFSET, 0).Resize(1, rightColumn - anchor.Column + 1)
End Function

' ------------------------------
' LC-зависимые результаты элементов
' ------------------------------
' В этот блок не пишутся координаты и площади: они уже есть в
' rngNDMSectionGeometry и связываются по ElementID.
Private Sub WriteElementResults(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal batch As CBatchSectionCalculator, _
        ByVal runID As Variant, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMElementResults").RefersToRange

    Dim elementCount As Long
    elementCount = section.ConcreteCount + section.RebarCount

    Dim stateCount As Long
    Dim i As Long
    ' Считаем только конечные расчётные состояния. Для одного сочетания
    ' могут сохраняться НДС по прочности, предельное НДС, текущее НДС
    ' с трещинами и состояния до и после образования трещины.
    For i = 1 To batch.Count
        stateCount = stateCount + batch.ResultAt(i).StateRepository.StateCount
    Next i

    Dim rowCount As Long
    rowCount = 1 + stateCount * elementCount
    If rowCount < 1 Then rowCount = 1

    Dim data() As Variant
    ReDim data(1 To rowCount, 1 To ELEMENT_RESULT_COLUMNS)
    PutElementResultHeaders data, units

    Dim rowIndex As Long
    rowIndex = 2
    ' Все строки собираются в массив VBA и одним присваиванием выгружаются
    ' на лист. Это принципиально быстрее и устойчивее построчной записи Cells.
    For i = 1 To batch.Count
        Dim stateIndex As Long
        For stateIndex = 1 To batch.ResultAt(i).StateRepository.StateCount
            PutCombinationElementResults data, rowIndex, section, batch, i, _
                batch.ResultAt(i).StateRepository.StateAt(stateIndex), runID, units
        Next stateIndex
    Next i

    target.Resize(rowCount, ELEMENT_RESULT_COLUMNS).Value2 = data
    FormatResultsBlock target, rowCount, ELEMENT_RESULT_COLUMNS, "Результаты элементов по сочетаниям"
End Sub

' ------------------------------
' Геометрия расчетных элементов
' ------------------------------
' Геометрия не зависит от LoadCase, поэтому записывается один раз. Схема и
' AutoCAD export затем объединяют эту таблицу с нужным LC из ElementResults.
Private Sub WriteSectionGeometry(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal runID As Variant, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMSectionGeometry").RefersToRange

    Dim rowCount As Long
    rowCount = 1 + section.ConcreteCount + section.RebarCount

    Dim data() As Variant
    ReDim data(1 To rowCount, 1 To GEOMETRY_COLUMNS)
    PutGeometryHeaders data, units

    Dim rowIndex As Long
    rowIndex = 2

    Dim i As Long
    ' Геометрия записывается один раз на расчетный снимок. Она не зависит от LoadCase,
    ' поэтому координаты и площади не дублируются в rngNDMElementResults для каждого сочетания.
    For i = 1 To section.ConcreteCount
        data(rowIndex, 1) = runID
        data(rowIndex, 2) = section.ConcreteID(i)
        data(rowIndex, 3) = "Concrete"
        data(rowIndex, 4) = OutputLength(section.ConcreteX(i), units)
        data(rowIndex, 5) = OutputLength(section.ConcreteY(i), units)
        data(rowIndex, 6) = OutputArea(section.ConcreteArea(i), units)
        data(rowIndex, 7) = section.ConcreteShapeType(i)
        data(rowIndex, 8) = OutputLength(section.ConcreteWidth(i), units)
        data(rowIndex, 9) = OutputLength(section.ConcreteHeight(i), units)
        data(rowIndex, 10) = vbNullString
        data(rowIndex, 11) = section.ConcreteRotation(i)
        data(rowIndex, 12) = OutputFourthPowerLength(section.ConcreteLocalIx(i), units)
        data(rowIndex, 13) = OutputFourthPowerLength(section.ConcreteLocalIy(i), units)
        data(rowIndex, 14) = OutputFourthPowerLength(section.ConcreteLocalIxy(i), units)
        data(rowIndex, 15) = section.ConcreteComment(i)
        rowIndex = rowIndex + 1
    Next i

    ' Бетон и арматуру записываем в одну таблицу геометрии. Для стержня указываем
    ' диаметр вместо ширины и высоты; тип элемента определяется полем MaterialType.
    For i = 1 To section.RebarCount
        data(rowIndex, 1) = runID
        data(rowIndex, 2) = section.RebarID(i)
        data(rowIndex, 3) = "Rebar"
        data(rowIndex, 4) = OutputLength(section.RebarX(i), units)
        data(rowIndex, 5) = OutputLength(section.RebarY(i), units)
        data(rowIndex, 6) = OutputArea(section.RebarArea(i), units)
        data(rowIndex, 7) = "Circle"
        data(rowIndex, 8) = vbNullString
        data(rowIndex, 9) = vbNullString
        data(rowIndex, 10) = OutputLength(section.RebarDiameter(i), units)
        data(rowIndex, 11) = 0#
        data(rowIndex, 12) = 0#
        data(rowIndex, 13) = 0#
        data(rowIndex, 14) = 0#
        data(rowIndex, 15) = section.RebarComment(i)
        If Len(section.RebarRecipeTrace(i)) > 0 Then data(rowIndex, 15) = data(rowIndex, 15) & vbLf & "RecipeTrace:" & section.RebarRecipeTrace(i)
        rowIndex = rowIndex + 1
    Next i

    target.Resize(rowCount, GEOMETRY_COLUMNS).Value2 = data
    FormatResultsBlock target, rowCount, GEOMETRY_COLUMNS, "Геометрия расчетных элементов сечения"
End Sub

' Сериализует исходные границы отдельно от элементов и LC-аннотаций.
' Дуга хранит начало, центр и знаковый угол; ее конец и радиус вычисляются
' при чтении. Радиус записывается только для окружности, концы - для линии.
' Шапка без строк означает отсутствие сохраненных достоверных границ,
' а не доказательство отсутствия отверстий. XY - координаты модели сечения.
Public Sub WriteSectionContours(ByVal workbook As Object, ByVal contours As CSectionContours, _
        ByVal runID As Variant, Optional ByVal units As CUnitSystem = Nothing)
    Dim target As Object: Set target = workbook.Names.Item("rngNDMSectionContours").RefersToRange
    ClearNamedRegion workbook, "rngNDMSectionContours"
    Dim data() As Variant, headers As Variant, i As Long, row As Long, kind As String
    headers = Array("RunID", "LoopID", "SegmentID", "Sequence", "LoopRole", "SegmentType", _
        "StartX", "StartY", "EndX", "EndY", "CenterX", "CenterY", "Radius", "SweepAngle, rad", "SourceID", "Comment", "LengthUnit (SectionXY)")
    ReDim data(1 To contours.Count + 1, 1 To CONTOUR_COLUMNS)
    For i = 0 To UBound(headers): data(1, i + 1) = headers(i): Next i
    For i = 1 To contours.Count
        row = i + 1: kind = contours.SegmentType(i)
        data(row, 1) = runID: data(row, 2) = contours.LoopID(i): data(row, 3) = contours.ContourID(i)
        data(row, 4) = i: data(row, 5) = contours.LoopRole(i): data(row, 6) = kind
        data(row, 7) = OutputLength(contours.StartX(i), units): data(row, 8) = OutputLength(contours.StartY(i), units)
        Select Case kind
            Case "CONTOUR_LINE"
                data(row, 9) = OutputLength(contours.EndX(i), units): data(row, 10) = OutputLength(contours.EndY(i), units)
            Case "CONTOUR_CIRCLE"
                data(row, 13) = OutputLength(contours.Radius(i), units)
            Case "CONTOUR_ARC"
                data(row, 11) = OutputLength(contours.CenterX(i), units): data(row, 12) = OutputLength(contours.CenterY(i), units)
                data(row, 14) = contours.SweepAngle(i)
        End Select
        data(row, 15) = contours.SourceID(i): data(row, 16) = contours.Comment(i)
        data(row, 17) = OutputLengthUnit(units)
    Next i
    target.Resize(contours.Count + 1, CONTOUR_COLUMNS).Value2 = data
    FormatResultsBlock target, contours.Count + 1, CONTOUR_COLUMNS, "Достоверные контуры и отверстия сечения"
End Sub


' ------------------------------
' Свойства сечения и расчётные состояния
' ------------------------------
' Записывает габариты, бетонные и приведённые характеристики, единицы вывода
' и поля деформаций каждого сохранённого НДС. Нагрузки пользователя заданы
' от бетонного центра; главные оси схемы и AutoCAD определяются
' по сохранённым приведённым характеристикам Transformed.*.
Private Sub WriteSectionProperties(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal snapshot As CSectionPropertiesCalculator, ByVal batch As CBatchSectionCalculator, _
        ByVal runID As Variant, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMSectionProperties").RefersToRange

    Dim rowCount As Long
    rowCount = SectionPropertyRowCount(batch)

    Dim data() As Variant
    ReDim data(1 To rowCount, 1 To SECTION_PROPERTY_COLUMNS)
    PutSectionPropertyHeaders data

    Dim rowIndex As Long
    rowIndex = 2

    ' Bounds нужны любой геометрии, Generated или AutoCAD. По ним схема может
    ' восстановить общий габарит без обращения к текущим настройкам Config.
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim bounds As Variant: bounds = snapshot.SnapshotBounds
    minX = bounds(1): maxX = bounds(2): minY = bounds(3): maxY = bounds(4)

    ' Бетонные и приведённые характеристики уже подготовлены.
    ' Переводим готовые значения в единицы вывода без пересчёта по текущему Config.
    Dim concreteProps As CSectionPropertiesCalculator
    Set concreteProps = snapshot

    Dim transformedProps As CSectionPropertiesCalculator
    Set transformedProps = snapshot.TransformedSnapshot

    PutProperty data, rowIndex, runID, "ALL", "Bounds.MinX", OutputLength(minX, units), OutputLengthUnit(units), "Минимальный X бетонной расчетной модели в координатах модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MaxX", OutputLength(maxX, units), OutputLengthUnit(units), "Максимальный X бетонной расчетной модели в координатах модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MinY", OutputLength(minY, units), OutputLengthUnit(units), "Минимальный Y бетонной расчетной модели в координатах модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MaxY", OutputLength(maxY, units), OutputLengthUnit(units), "Максимальный Y бетонной расчетной модели в координатах модели."
    PutProperty data, rowIndex, runID, "ALL", "Geometry.Source", section.SourceType, "-", "Источник геометрии расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Geometry.ConcreteMeshSource", section.ConcreteMeshSource, "-", "Источник бетонной сетки: Generated, AutoCADRegions или AutoCADContours."
    PutProperty data, rowIndex, runID, "ALL", "Geometry.Revision", section.Revision, "-", "Версия геометрии, к которой относятся состояния и расчетные области этого снимка."
    PutProperty data, rowIndex, runID, "ALL", "Geometry.RecipeDiagnostics", section.GeometryDiagnosticsText, "-", "Запрошенные/созданные позиции, совпадения, пропуски и конфликты. StationStep и CenterStep в мм; для связанного ряда StationStep относится к исходному пути."
    PutProperty data, rowIndex, runID, "ALL", "LoadReferenceX", OutputLength(batch.LoadReferenceX, units), OutputLengthUnit(units), "X точки приложения нагрузки в координатах модели; это центр тяжести бетонного сечения плюс Load.ReferenceOffsetX."
    PutProperty data, rowIndex, runID, "ALL", "LoadReferenceY", OutputLength(batch.LoadReferenceY, units), OutputLengthUnit(units), "Y точки приложения нагрузки в координатах модели; это центр тяжести бетонного сечения плюс Load.ReferenceOffsetY."
    PutProperty data, rowIndex, runID, "ALL", "ConcreteCount", section.ConcreteCount, "шт", "Количество бетонных элементов расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "RebarCount", section.RebarCount, "шт", "Количество элементов продольной арматуры расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Output.LengthUnit", OutputLengthUnit(units), "-", "Единица вывода длины на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.AreaUnit", OutputAreaUnit(units), "-", "Единица вывода площади на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.ForceUnit", OutputForceUnit(units), "-", "Единица вывода силы на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.MomentUnit", OutputMomentUnit(units), "-", "Единица вывода момента на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.StressUnit", OutputStressUnit(units), "-", "Единица вывода напряжения на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.CurvatureUnit", OutputCurvatureUnit(units), "-", "Единица вывода кривизны на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.SignConvention.N", OutputSignN(units), "-", "Пользовательский знак N на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.SignConvention.Mx", OutputSignMx(units), "-", "Пользовательский знак Mx на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.SignConvention.My", OutputSignMy(units), "-", "Пользовательский знак My на момент расчета."
    PutConcreteSectionProperties data, rowIndex, runID, section, concreteProps, units
    PutTransformedSectionProperties data, rowIndex, runID, section, transformedProps, units
    PutGeneralSectionCheckProperties data, rowIndex, runID, section, concreteProps, transformedProps, batch
    Dim i As Long
    ' Общие свойства НДС хранятся отдельно от строк элементов.
    ' Поля деформаций записываем как State.<StateType>.*, чтобы схема
    ' и AutoCAD могли выбрать нужное состояние конкретного сочетания.
    For i = 1 To batch.Count
        If batch.ResultAt(i).StrengthResult.DirectState.StateAvailable Then
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "Nint", OutputForce(batch.ResultAt(i).StrengthResult.DirectState.StateResult.Nint, units), OutputForceUnit(units), "Равнодействующая продольная сила внутреннего НДС для выбранного сочетания."
        Else
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "Nint", vbNullString, OutputForceUnit(units), "Равнодействующая продольная сила внутреннего НДС для выбранного сочетания."
        End If
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "N", OutputForce(batch.N(i), units), OutputForceUnit(units), "Заданная пользователем продольная сила сочетания."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "MxUser", OutputMx(batch.UserMx(i), units), OutputMomentUnit(units), "Пользовательский Mx в точке приложения нагрузки."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "MyUser", OutputMy(batch.UserMy(i), units), OutputMomentUnit(units), "Пользовательский My в точке приложения нагрузки."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "MxInternal", OutputMx(batch.Mx(i), units), OutputMomentUnit(units), "Mx того же сочетания после переноса к началу внутренней системы координат модели."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "MyInternal", OutputMy(batch.My(i), units), OutputMomentUnit(units), "My того же сочетания после переноса к началу внутренней системы координат модели."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "OverallStatus", batch.ResultAt(i).Status, "-", "Итоговый пользовательский статус сочетания."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "ProfileId", batch.ResultAt(i).ProfileId, "-", "Расчетный профиль, назначенный сочетанию в rngLoadCombinations."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "DirectStateStatus", batch.ResultAt(i).StrengthResult.DirectState.Status, "-", "Статус прямого решения НДС фактического сочетания."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "CapacityStatus", batch.ResultAt(i).StrengthResult.Capacity.Status, "-", "Статус расчета несущей способности."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "LoadPath", batch.ResultAt(i).CapacityPathDisplayName, "-", "Какие компоненты нагрузки масштабировались при поиске несущей способности."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "CapacitySolutionMethod", batch.ResultAt(i).StrengthResult.Capacity.SolutionMethod, "-", "Фактический численный метод, которым найдена предельная несущая способность."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "LambdaCapacity", batch.ResultAt(i).StrengthResult.Capacity.LambdaCapacity, "-", "Предельный множитель для выбранной lambda-траектории."
        If batch.ResultAt(i).StrengthResult.Capacity.ReserveFactor > 0# Then
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "Nultimate", OutputForce(batch.ResultAt(i).StrengthResult.Capacity.NUltimate, units), OutputForceUnit(units), "Полная N при lambdaUltimate в точке приложения нагрузки."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "MxUltimate", OutputMx(batch.ResultAt(i).StrengthResult.Capacity.MxUltimate, units), OutputMomentUnit(units), "Полный Mx при lambdaUltimate в точке приложения нагрузки."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "MyUltimate", OutputMy(batch.ResultAt(i).StrengthResult.Capacity.MyUltimate, units), OutputMomentUnit(units), "Полный My при lambdaUltimate в точке приложения нагрузки."
        Else
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "Nultimate", vbNullString, OutputForceUnit(units), "Заполняется после успешного поиска несущей способности."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "MxUltimate", vbNullString, OutputMomentUnit(units), "Заполняется после успешного поиска несущей способности."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "MyUltimate", vbNullString, OutputMomentUnit(units), "Заполняется после успешного поиска несущей способности."
        End If
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "CrackStatus", batch.ResultAt(i).NormalCrackStatus, "-", "Статус расчета ширины раскрытия трещин."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "LongitudinalCrackStatus", batch.ResultAt(i).CrackResult.Longitudinal.Status, "-", "Статус SLS(II)-проверки продольных трещин по сжатому бетону; выполняется как часть Calculation.Crack.Width."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "LongitudinalCrack.MaxConcreteCompressionStress", OutputStress(batch.ResultAt(i).CrackResult.Longitudinal.MaxCompressionStress, units), OutputStressUnit(units), "Максимальное по модулю сжимающее напряжение бетона в прямом НДС."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "LongitudinalCrack.RbMc2", OutputStress(batch.ResultAt(i).CrackResult.Longitudinal.RbMc2, units), OutputStressUnit(units), "Пользовательское допустимое Rb,mc2 для проверки продольных трещин."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "LongitudinalCrack.Utilization", batch.ResultAt(i).CrackResult.Longitudinal.Utilization, "-", "Использование по продольным трещинам: sigma_c,max / Rb,mc2."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "ExtensionUsed", BooleanText(batch.ResultAt(i).ExtensionUsed), "-", "True, если расчетное состояние вышло за физические пределы диаграммы в numerical extension."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "General.DiagramExtension", BooleanText(batch.ResultAt(i).DiagramExtensionEnabled), "-", "Сохраненное значение настройки General.DiagramExtension на момент расчета."
        PutNamedStateProperties data, rowIndex, runID, batch, i, units
    Next i

    target.Resize(rowCount, SECTION_PROPERTY_COLUMNS).Value2 = data
    FormatResultsBlock target, rowCount, SECTION_PROPERTY_COLUMNS, "Свойства всего сечения"
End Sub

' Записывает логический признак одинаково в любой локали Excel:
' TRUE или FALSE вместо зависящего от языка интерфейса представления.
Private Function BooleanText(ByVal value As Boolean) As String
    If value Then
        BooleanText = "True"
    Else
        BooleanText = "False"
    End If
End Function

' ------------------------------
' Сохранение использованных диаграмм
' ------------------------------
' Записывает точки различных CMaterialDiagram конечных расчётных состояний.
' Учитывает профиль, вид НДС, параметры материалов и численное продолжение.
' Разные сочетания с одинаковыми диаграммами не дублируют их точки.
Private Sub WriteMaterialDiagrams(ByVal workbook As Object, _
        ByVal batch As CBatchSectionCalculator, ByVal runID As Variant, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMMaterialDiagrams").RefersToRange

    Dim rowCapacity As Long
    rowCapacity = MaterialDiagramRowCount(batch)

    Dim data() As Variant
    ReDim data(1 To rowCapacity, 1 To MATERIAL_DIAGRAM_COLUMNS)
    PutMaterialDiagramHeaders data, units

    Dim rowIndex As Long
    rowIndex = 2
    Dim written As Object
    Set written = CreateObject("Scripting.Dictionary")
    Dim strainRanges As Object
    Set strainRanges = MaterialDiagramStrainRanges(batch)

    Dim i As Long
    For i = 1 To batch.Count
        PutCombinationMaterialDiagrams data, rowIndex, runID, batch, i, units, written, strainRanges
    Next i

    Dim rowCount As Long
    rowCount = rowIndex - 1
    If rowCount < 1 Then rowCount = 1

    target.Resize(rowCount, MATERIAL_DIAGRAM_COLUMNS).Value2 = data
    FormatResultsBlock target, rowCount, MATERIAL_DIAGRAM_COLUMNS, "Фактические диаграммы материалов"
End Sub

' Добавляет диаграммы сохранённых состояний сочетания, если их DiagramId ещё не записан.
' У разных профилей один вид НДС может использовать разные параметры материалов.
Private Sub PutCombinationMaterialDiagrams(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal runID As Variant, _
        ByVal batch As CBatchSectionCalculator, ByVal combinationIndex As Long, ByVal units As CUnitSystem, _
        ByVal written As Object, ByVal strainRanges As Object)
    Dim stateIndex As Long
    For stateIndex = 1 To batch.ResultAt(combinationIndex).StateRepository.StateCount
        Dim stateResult As CSectionStateResult
        Set stateResult = batch.ResultAt(combinationIndex).StateRepository.StateAt(stateIndex)
        If Not stateResult Is Nothing Then
            If stateResult.HasMaterialSnapshot Then
                AppendMaterialDiagramPoints data, rowIndex, runID, batch.ResultAt(combinationIndex).ProfileId, _
                    stateResult, "Concrete", _
                    stateResult.ConcreteDiagramSnapshot, _
                    units, written, strainRanges
                AppendMaterialDiagramPoints data, rowIndex, runID, batch.ResultAt(combinationIndex).ProfileId, _
                    stateResult, "Rebar", _
                    stateResult.SteelDiagramSnapshot, _
                    units, written, strainRanges
            End If
        End If
    Next stateIndex
End Sub

' Строит ID диаграммы по её фактическим параметрам. Номер сочетания в ключ
' не входит: одинаковые профиль и модель материалов могут использоваться
' несколькими сочетаниями.
Private Function MaterialDiagramId(ByVal profileId As String, ByVal stateResult As CSectionStateResult, _
        ByVal materialKind As String, ByVal diagram As CMaterialDiagram) As String
    MaterialDiagramId = profileId & "|" & stateResult.StateTypeText & "|" & _
        stateResult.MaterialModelRoleText & "|" & materialKind & "|" & _
        MaterialDiagramModeText(diagram) & "|" & _
        stateResult.MaterialSpec.SpecKey
End Function

' Показывает, была ли в точках диаграммы записана физическая модель или
' техническое продолжение активной диаграммы равновесия.
Private Function MaterialDiagramModeText(ByVal diagram As CMaterialDiagram) As String
    MaterialDiagramModeText = "Physical"
    If diagram Is Nothing Then Exit Function
    If diagram.HasCompressionExtension Or diagram.HasTensionExtension Then
        MaterialDiagramModeText = "StateExtension"
    Else
        MaterialDiagramModeText = "Physical"
    End If
End Function

' Записывает физические узлы и опорные точки продолжения. Если найденные НДС
' выходят дальше опорных точек, продлевает вывод до общего диапазона этих НДС.
' Это границы рисунка и таблицы, а не ограничение расчётной диаграммы.
Private Sub AppendMaterialDiagramPoints(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal runID As Variant, ByVal profileId As String, ByVal stateResult As CSectionStateResult, _
        ByVal materialKind As String, ByVal diagram As CMaterialDiagram, _
        ByVal units As CUnitSystem, ByVal written As Object, ByVal strainRanges As Object)
    If diagram Is Nothing Then Exit Sub

    Dim diagramId As String
    diagramId = MaterialDiagramId(profileId, stateResult, materialKind, diagram)
    If written.Exists(diagramId) Then Exit Sub
    written.Add diagramId, True

    Dim pointIndex As Long, strain As Double, stress As Double, bounds As Variant
    For pointIndex = 1 To diagram.PointCount
        strain = diagram.PointStrain(pointIndex)
        stress = diagram.PointStress(pointIndex)
        If strainRanges.Exists(diagramId) Then
            bounds = strainRanges(diagramId)
            If pointIndex = 1 And diagram.HasCompressionExtension Then strain = MinDouble(strain, CDbl(bounds(0)))
            If pointIndex = diagram.PointCount And diagram.HasTensionExtension Then strain = MaxDouble(strain, CDbl(bounds(1)))
            If strain <> diagram.PointStrain(pointIndex) Then stress = diagram.GetStress(strain)
        End If
        data(rowIndex, 1) = runID
        data(rowIndex, 2) = diagramId
        data(rowIndex, 3) = profileId
        data(rowIndex, 4) = stateResult.StateTypeText
        data(rowIndex, 5) = stateResult.MaterialModelRoleText
        data(rowIndex, 6) = stateResult.MaterialSpec.SpecKey
        data(rowIndex, 7) = MaterialDiagramModeText(diagram)
        data(rowIndex, 8) = materialKind
        data(rowIndex, 9) = pointIndex
        data(rowIndex, 10) = strain
        data(rowIndex, 11) = OutputStress(stress, units)
        rowIndex = rowIndex + 1
    Next pointIndex
End Sub

' Собирает диапазоны деформаций всех найденных НДС с одинаковой диаграммой.
' Неудачные итерации не определяют границы вывода: их НДС не подтверждено.
Private Function MaterialDiagramStrainRanges(ByVal batch As CBatchSectionCalculator) As Object
    Dim result As Object, combinationIndex As Long, stateIndex As Long, state As CSectionStateResult
    Set result = CreateObject("Scripting.Dictionary")
    For combinationIndex = 1 To batch.Count
        For stateIndex = 1 To batch.ResultAt(combinationIndex).StateRepository.StateCount
            Set state = batch.ResultAt(combinationIndex).StateRepository.StateAt(stateIndex)
            If Not state Is Nothing Then
                If state.Converged And state.HasMaterialSnapshot Then
                    IncludeMaterialDiagramRange result, MaterialDiagramId(batch.ResultAt(combinationIndex).ProfileId, _
                        state, "Concrete", state.ConcreteDiagramSnapshot), state.MinConcreteStrain, state.MaxConcreteStrain
                    IncludeMaterialDiagramRange result, MaterialDiagramId(batch.ResultAt(combinationIndex).ProfileId, _
                        state, "Rebar", state.SteelDiagramSnapshot), state.MinSteelStrain, state.MaxSteelStrain
                End If
            End If
        Next stateIndex
    Next combinationIndex
    Set MaterialDiagramStrainRanges = result
End Function

' Объединяет крайние деформации нескольких сочетаний без повторения диаграммы.
Private Sub IncludeMaterialDiagramRange(ByVal ranges As Object, ByVal diagramId As String, _
        ByVal minimumStrain As Double, ByVal maximumStrain As Double)
    Dim bounds As Variant
    If ranges.Exists(diagramId) Then
        bounds = ranges(diagramId)
        bounds(0) = MinDouble(CDbl(bounds(0)), minimumStrain)
        bounds(1) = MaxDouble(CDbl(bounds(1)), maximumStrain)
        ranges(diagramId) = bounds
    Else
        ranges.Add diagramId, Array(minimumStrain, maximumStrain)
    End If
End Sub

' Шапка диагностической таблицы диаграмм материалов.
Private Sub PutMaterialDiagramHeaders(ByRef data() As Variant, ByVal units As CUnitSystem)
    data(1, 1) = "RunID"
    data(1, 2) = "DiagramId"
    data(1, 3) = "ProfileId"
    data(1, 4) = "StateType"
    data(1, 5) = "MaterialModelRole"
    data(1, 6) = "MaterialModelSpec"
    data(1, 7) = "DiagramMode"
    data(1, 8) = "Material"
    data(1, 9) = "PointIndex"
    data(1, 10) = "Strain"
    data(1, 11) = "Stress, " & OutputStressUnit(units)
End Sub

' Записывает справочные свойства бетонной части сечения. По координатам центра
' можно проверить пользовательское смещение Load.ReferenceOffset
' и построить бетонные главные оси на схеме или в AutoCAD.
Private Sub PutConcreteSectionProperties(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal runID As Variant, ByVal section As CSectionModel, _
        ByVal props As CSectionPropertiesCalculator, ByVal units As CUnitSystem)
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Area", OutputArea(props.Area, units), OutputAreaUnit(units), "Площадь бетонной части сечения по расчетной сетке."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.CentroidX", OutputLength(props.CentroidX, units), OutputLengthUnit(units), "X центра тяжести бетонного сечения в координатах модели; от этой точки задается Load.ReferenceOffsetX."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.CentroidY", OutputLength(props.CentroidY, units), OutputLengthUnit(units), "Y центра тяжести бетонного сечения в координатах модели; от этой точки задается Load.ReferenceOffsetY."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Ix", OutputFourthPowerLength(props.Ix, units), OutputLengthUnit(units) & "4", "Момент инерции бетонного сечения относительно оси X внутренней системы координат модели."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Iy", OutputFourthPowerLength(props.Iy, units), OutputLengthUnit(units) & "4", "Момент инерции бетонного сечения относительно оси Y внутренней системы координат модели."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Ixy", OutputFourthPowerLength(props.Ixy, units), OutputLengthUnit(units) & "4", "Центробежный момент бетонного сечения относительно осей X/Y внутренней системы координат модели."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Ixc", OutputFourthPowerLength(props.Ixc, units), OutputLengthUnit(units) & "4", "Момент инерции бетонного сечения относительно центральной оси X через ц.т. бетонного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Iyc", OutputFourthPowerLength(props.Iyc, units), OutputLengthUnit(units) & "4", "Момент инерции бетонного сечения относительно центральной оси Y через ц.т. бетонного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Ixyc", OutputFourthPowerLength(props.Ixyc, units), OutputLengthUnit(units) & "4", "Центробежный момент бетонного сечения относительно центральных осей X/Y через ц.т. бетонного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.RadiusX", OutputLength(props.RadiusX, units), OutputLengthUnit(units), "Радиус инерции бетонного сечения относительно центральной оси X через ц.т. бетонного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.RadiusY", OutputLength(props.RadiusY, units), OutputLengthUnit(units), "Радиус инерции бетонного сечения относительно центральной оси Y через ц.т. бетонного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.PrincipalAngle", props.PrincipalAngleRad, "rad", "Угол от +X модели к главной центральной оси 1 бетонного сечения; положительный против часовой стрелки."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.PrincipalI1", OutputFourthPowerLength(props.PrincipalI1, units), OutputLengthUnit(units) & "4", "Главный центральный момент инерции бетонного сечения относительно главной оси 1 (больший)."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.PrincipalI2", OutputFourthPowerLength(props.PrincipalI2, units), OutputLengthUnit(units) & "4", "Главный центральный момент инерции бетонного сечения относительно главной оси 2 (меньший)."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.PrincipalRadius1", OutputLength(props.PrincipalRadius1, units), OutputLengthUnit(units), "Радиус инерции бетонного сечения относительно главной центральной оси 1."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.PrincipalRadius2", OutputLength(props.PrincipalRadius2, units), OutputLengthUnit(units), "Радиус инерции бетонного сечения относительно главной центральной оси 2."
    PutAxisCoreDistanceProperties data, rowIndex, runID, "Concrete", _
        "бетонного сечения", section, props, units
    PutPrincipalCoreDistanceProperties data, rowIndex, runID, "Concrete", _
        "бетонного сечения", section, props, units
End Sub

' Записывает свойства приведенного сечения бетон + арматура.
' Они остаются отдельным набором справочных данных и совместимы с уже
' существующей логикой главных осей на схеме и в AutoCAD export.
Private Sub PutTransformedSectionProperties(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal runID As Variant, ByVal section As CSectionModel, _
        ByVal props As CSectionPropertiesCalculator, ByVal units As CUnitSystem)
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Area", OutputArea(props.Area, units), OutputAreaUnit(units), "Площадь приведенного сечения в бетонной базе."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.CentroidX", OutputLength(props.CentroidX, units), OutputLengthUnit(units), "X центра тяжести приведенного сечения бетон + арматура в координатах модели; это смещение ц.т. приведенного сечения относительно бетонной модели."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.CentroidY", OutputLength(props.CentroidY, units), OutputLengthUnit(units), "Y центра тяжести приведенного сечения бетон + арматура в координатах модели; это смещение ц.т. приведенного сечения относительно бетонной модели."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Ix", OutputFourthPowerLength(props.Ix, units), OutputLengthUnit(units) & "4", "Момент инерции приведенного сечения относительно оси X внутренней системы координат модели."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Iy", OutputFourthPowerLength(props.Iy, units), OutputLengthUnit(units) & "4", "Момент инерции приведенного сечения относительно оси Y внутренней системы координат модели."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Ixy", OutputFourthPowerLength(props.Ixy, units), OutputLengthUnit(units) & "4", "Центробежный момент приведенного сечения относительно осей X/Y внутренней системы координат модели."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Ixc", OutputFourthPowerLength(props.Ixc, units), OutputLengthUnit(units) & "4", "Момент инерции приведенного сечения относительно центральной оси X через ц.т. приведенного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Iyc", OutputFourthPowerLength(props.Iyc, units), OutputLengthUnit(units) & "4", "Момент инерции приведенного сечения относительно центральной оси Y через ц.т. приведенного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Ixyc", OutputFourthPowerLength(props.Ixyc, units), OutputLengthUnit(units) & "4", "Центробежный момент приведенного сечения относительно центральных осей X/Y через ц.т. приведенного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.RadiusX", OutputLength(props.RadiusX, units), OutputLengthUnit(units), "Радиус инерции приведенного сечения относительно центральной оси X через ц.т. приведенного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.RadiusY", OutputLength(props.RadiusY, units), OutputLengthUnit(units), "Радиус инерции приведенного сечения относительно центральной оси Y через ц.т. приведенного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.PrincipalAngle", props.PrincipalAngleRad, "rad", "Угол от +X модели к главной центральной оси 1 приведенного сечения; положительный против часовой стрелки."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.PrincipalI1", OutputFourthPowerLength(props.PrincipalI1, units), OutputLengthUnit(units) & "4", "Главный центральный момент инерции приведенного сечения относительно главной оси 1 (больший)."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.PrincipalI2", OutputFourthPowerLength(props.PrincipalI2, units), OutputLengthUnit(units) & "4", "Главный центральный момент инерции приведенного сечения относительно главной оси 2 (меньший)."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.PrincipalRadius1", OutputLength(props.PrincipalRadius1, units), OutputLengthUnit(units), "Радиус инерции приведенного сечения относительно главной центральной оси 1."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.PrincipalRadius2", OutputLength(props.PrincipalRadius2, units), OutputLengthUnit(units), "Радиус инерции приведенного сечения относительно главной центральной оси 2."
    PutAxisCoreDistanceProperties data, rowIndex, runID, "Transformed", _
        "приведенного сечения", section, props, units
    PutPrincipalCoreDistanceProperties data, rowIndex, runID, "Transformed", _
        "приведенного сечения", section, props, units
End Sub

' Записывает ядровые расстояния по глобальным центральным осям X/Y.
' Расчет идет через тот же CoreDistanceAlong, что и устойчивость СП 35:
' A/I берутся из переданного калькулятора, а крайняя точка c - по бетонной
' оболочке section, без расширения контура арматурой. Plus/Minus относятся
' к направлению эксцентриситета, а не к стороне крайнего волокна.
Private Sub PutAxisCoreDistanceProperties(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal runID As Variant, ByVal prefix As String, ByVal description As String, _
        ByVal section As CSectionModel, ByVal props As CSectionPropertiesCalculator, _
        ByVal units As CUnitSystem)
    PutProperty data, rowIndex, runID, "ALL", prefix & ".CoreDistanceXPlus", _
        OutputLength(props.CoreDistanceAlong(section, 1#, 0#, False), units), _
        OutputLengthUnit(units), "Ядровое расстояние r " & description & " для эксцентриситета в направлении +X центральных осей X/Y."
    PutProperty data, rowIndex, runID, "ALL", prefix & ".CoreDistanceXMinus", _
        OutputLength(props.CoreDistanceAlong(section, -1#, 0#, False), units), _
        OutputLengthUnit(units), "Ядровое расстояние r " & description & " для эксцентриситета в направлении -X центральных осей X/Y."
    PutProperty data, rowIndex, runID, "ALL", prefix & ".CoreDistanceYPlus", _
        OutputLength(props.CoreDistanceAlong(section, 0#, 1#, False), units), _
        OutputLengthUnit(units), "Ядровое расстояние r " & description & " для эксцентриситета в направлении +Y центральных осей X/Y."
    PutProperty data, rowIndex, runID, "ALL", prefix & ".CoreDistanceYMinus", _
        OutputLength(props.CoreDistanceAlong(section, 0#, -1#, False), units), _
        OutputLengthUnit(units), "Ядровое расстояние r " & description & " для эксцентриситета в направлении -Y центральных осей X/Y."
End Sub

' Записывает справочные ядровые расстояния именно по главным центральным осям
' 1 и 2. Это не то же самое, что "главная плоскость" в расчете устойчивости:
' плоскость 1 изгибается вокруг оси 1, а ее габарит и r берутся по нормали,
' то есть по оси 2. Здесь же имя CoreDistance1Plus означает направление +оси 1.
Private Sub PutPrincipalCoreDistanceProperties(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal runID As Variant, ByVal prefix As String, ByVal description As String, _
        ByVal section As CSectionModel, ByVal props As CSectionPropertiesCalculator, _
        ByVal units As CUnitSystem)
    Dim axis1X As Double
    Dim axis1Y As Double
    Dim axis2X As Double
    Dim axis2Y As Double
    props.PrincipalAxisDirection 1, axis1X, axis1Y
    props.PrincipalAxisDirection 2, axis2X, axis2Y

    PutProperty data, rowIndex, runID, "ALL", prefix & ".CoreDistance1Plus", _
        OutputLength(props.CoreDistanceAlong(section, axis1X, axis1Y, False), units), _
        OutputLengthUnit(units), "Ядровое расстояние r " & description & " для эксцентриситета в направлении +главной центральной оси 1."
    PutProperty data, rowIndex, runID, "ALL", prefix & ".CoreDistance1Minus", _
        OutputLength(props.CoreDistanceAlong(section, -axis1X, -axis1Y, False), units), _
        OutputLengthUnit(units), "Ядровое расстояние r " & description & " для эксцентриситета в направлении -главной центральной оси 1."
    PutProperty data, rowIndex, runID, "ALL", prefix & ".CoreDistance2Plus", _
        OutputLength(props.CoreDistanceAlong(section, axis2X, axis2Y, False), units), _
        OutputLengthUnit(units), "Ядровое расстояние r " & description & " для эксцентриситета в направлении +главной центральной оси 2."
    PutProperty data, rowIndex, runID, "ALL", prefix & ".CoreDistance2Minus", _
        OutputLength(props.CoreDistanceAlong(section, -axis2X, -axis2Y, False), units), _
        OutputLengthUnit(units), "Ядровое расстояние r " & description & " для эксцентриситета в направлении -главной центральной оси 2."
End Sub

' Записывает общие расчетные характеристики, которые полезны не только в
' подробном блоке устойчивости: процент продольного армирования и гибкость
' l0/i по СП 35 для двух главных плоскостей приведенного сечения.
Private Sub PutGeneralSectionCheckProperties(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal runID As Variant, ByVal section As CSectionModel, _
        ByVal concreteProps As CSectionPropertiesCalculator, _
        ByVal transformedProps As CSectionPropertiesCalculator, ByVal batch As CBatchSectionCalculator)
    Dim steelArea As Double
    Dim reinforcementRatio As Variant
    steelArea = transformedProps.RebarAreaTotal(section)
    If concreteProps.Area > 0# Then
        reinforcementRatio = 100# * steelArea / concreteProps.Area
    Else
        reinforcementRatio = vbNullString
    End If

    PutProperty data, rowIndex, runID, "ALL", "Rebar.RatioPercent", reinforcementRatio, "%", _
        "Процент продольного армирования: суммарная As / gross-площадь бетонной сетки."
    PutProperty data, rowIndex, runID, "ALL", "SP35.Slenderness1", _
        SlendernessValue(batch.StabilityElementLength, batch.StabilityMu1, transformedProps.PrincipalRadius1), _
        "-", "Гибкость l0/i по СП 35 для плоскости 1: расчет изгиба вокруг главной оси 1 приведенного сечения."
    PutProperty data, rowIndex, runID, "ALL", "SP35.Slenderness2", _
        SlendernessValue(batch.StabilityElementLength, batch.StabilityMu2, transformedProps.PrincipalRadius2), _
        "-", "Гибкость l0/i по СП 35 для плоскости 2: расчет изгиба вокруг главной оси 2 приведенного сечения."
End Sub

' Возвращает безразмерную гибкость l0/i, если расчетная длина и радиус
' инерции заданы корректно; иначе оставляет справочное поле пустым.
Private Function SlendernessValue(ByVal geometricLength As Double, ByVal muValue As Double, _
        ByVal radiusValue As Double) As Variant
    If geometricLength <= 0# Or muValue <= 0# Or radiusValue <= 0# Then
        SlendernessValue = vbNullString
    Else
        SlendernessValue = geometricLength * muValue / radiusValue
    End If
End Function

' Записывает общие данные каждого сохранённого НДС: поле деформаций
' и фактически использованные параметры материалов.
' Повторять их в каждой строке бетонного элемента или стержня не требуется.
Private Sub PutNamedStateProperties(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal runID As Variant, ByVal batch As CBatchSectionCalculator, ByVal combinationIndex As Long, _
        ByVal units As CUnitSystem)
    Dim stateIndex As Long
    For stateIndex = 1 To batch.ResultAt(combinationIndex).StateRepository.StateCount
        Dim stateResult As CSectionStateResult
        Set stateResult = batch.ResultAt(combinationIndex).StateRepository.StateAt(stateIndex)
        If Not stateResult Is Nothing Then
            Dim prefix As String
            prefix = "State." & stateResult.StateTypeText & "."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "MaterialModelRole", _
                stateResult.MaterialModelRoleText, "-", "Смысловая роль материальной модели для именованного расчетного состояния."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "MaterialModelSpec", _
                stateResult.MaterialSpec.SpecKey, "-", "Фактическая спецификация ValueSet/ConcreteDiagram/ConcreteTension/SteelDiagram."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "ConcreteDiagramId", _
                MaterialDiagramId(batch.ResultAt(combinationIndex).ProfileId, stateResult, "Concrete", stateResult.ConcreteDiagramSnapshot), _
                "-", "Ссылка на бетонную диаграмму в rngNDMMaterialDiagrams."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "RebarDiagramId", _
                MaterialDiagramId(batch.ResultAt(combinationIndex).ProfileId, stateResult, "Rebar", stateResult.SteelDiagramSnapshot), _
                "-", "Ссылка на диаграмму арматуры в rngNDMMaterialDiagrams."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "Epsilon0", _
                stateResult.Epsilon0, "-", "Деформация в исходной точке X=0, Y=0 модели (для импорта AutoCAD - начало МСК/WCS). Численный перенос к центру сечения эту точку вывода не меняет."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "KappaX", _
                OutputCurvature(stateResult.KappaX, units), OutputCurvatureUnit(units), "Кривизна KappaX при координате Y внутренней системы координат модели."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "KappaY", _
                OutputCurvature(stateResult.KappaY, units), OutputCurvatureUnit(units), "Кривизна KappaY при координате X внутренней системы координат модели."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "NeutralLine.Angle", _
                NeutralLineAngleValue(stateResult.KappaX, stateResult.KappaY), "rad", _
                "Угол нейтральной линии от оси X внутренней системы координат модели против часовой стрелки."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "NeutralLine.SignedDistance", _
                OutputLengthVariant(NeutralLineSignedDistanceValue(stateResult.Epsilon0, stateResult.KappaX, stateResult.KappaY), units), _
                OutputLengthUnit(units), "Знаковое расстояние от начала внутренней системы координат модели до нейтральной линии по нормали (KappaY, KappaX)."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "ExtensionUsed", _
                BooleanText(stateResult.ExtensionUsed), "-", "True, если это расчетное состояние использовало numerical extension."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "Status", _
                stateResult.Status, "-", "Короткий статус НДС, сохранённый при расчёте сочетания."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "ResultComment", _
                CalculationResultMessage(batch, combinationIndex, stateResult.ResultMeta), "-", "Причина результата НДС, номер сочетания и сохранённый адрес исходной строки нагрузок."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "StopReason", _
                stateResult.StopReason, "-", "Служебная причина остановки solver-а для диагностики."
        End If
    Next stateIndex
End Sub

' Возвращает угол нейтральной линии для поля eps = eps0 + kx*y + ky*x.
' Если кривизны равны нулю, нейтральной линии нет и поле Results остается пустым.
Private Function NeutralLineAngleValue(ByVal kappaX As Double, ByVal kappaY As Double) As Variant
    If NeutralCurvatureNorm(kappaX, kappaY) <= 0.000000000000001 Then
        NeutralLineAngleValue = vbNullString
    Else
        NeutralLineAngleValue = Atn2Local(-kappaY, kappaX)
    End If
End Function

' Возвращает знаковое расстояние c/sqrt(a^2+b^2) для линии
' kappaY*x + kappaX*y + epsilon0 = 0. Знак привязан к нормали (kappaY, kappaX).
Private Function NeutralLineSignedDistanceValue(ByVal epsilon0 As Double, _
        ByVal kappaX As Double, ByVal kappaY As Double) As Variant
    Dim normValue As Double
    normValue = NeutralCurvatureNorm(kappaX, kappaY)
    If normValue <= 0.000000000000001 Then
        NeutralLineSignedDistanceValue = vbNullString
    Else
        NeutralLineSignedDistanceValue = epsilon0 / normValue
    End If
End Function

' Норма вектора нормали нейтральной линии. Она же показывает, существует ли
' линия: при нулевых kappa все сечение имеет постоянную деформацию.
Private Function NeutralCurvatureNorm(ByVal kappaX As Double, ByVal kappaY As Double) As Double
    NeutralCurvatureNorm = Sqr(kappaX * kappaX + kappaY * kappaY)
End Function

' Переводит длину, которая может быть пустой Variant, в OUTPUT-единицы.
Private Function OutputLengthVariant(ByVal value As Variant, ByVal units As CUnitSystem) As Variant
    If Len(CStr(value)) = 0 Then
        OutputLengthVariant = vbNullString
    Else
        OutputLengthVariant = OutputLength(CDbl(value), units)
    End If
End Function

' Вычисляет atan2 для справочного угла нейтральной линии при записи на лист.
Private Function Atn2Local(ByVal y As Double, ByVal x As Double) As Double
    If Abs(x) <= 0.000000000000001 Then
        If y > 0# Then
            Atn2Local = GEOM_PI / 2#
        ElseIf y < 0# Then
            Atn2Local = -GEOM_PI / 2#
        Else
            Atn2Local = 0#
        End If
    Else
        Atn2Local = Atn(y / x)
        If x < 0# Then
            If y >= 0# Then
                Atn2Local = Atn2Local + GEOM_PI
            Else
                Atn2Local = Atn2Local - GEOM_PI
            End If
        End If
    End If
End Function

' Вычисляет точное число строк свойств сечения перед выделением массива.
' Количество сочетаний и сохраняемых НДС меняется; фиксированного запаса
' строк могло бы не хватить, что привело бы к выходу за границы массива.
Private Function SectionPropertyRowCount(ByVal batch As CBatchSectionCalculator) As Long
    Dim rowCount As Long
    rowCount = 1 + COMMON_SECTION_PROPERTY_ROWS + batch.Count * LC_SECTION_PROPERTY_ROWS

    Dim combinationIndex As Long
    For combinationIndex = 1 To batch.Count
        rowCount = rowCount + batch.ResultAt(combinationIndex).StateRepository.StateCount * STATE_SECTION_PROPERTY_ROWS
    Next combinationIndex

    SectionPropertyRowCount = rowCount
End Function

' Определяет точное число строк таблицы диаграмм по сохранённым состояниям
' и числу точек каждой используемой модели материалов.
Private Function MaterialDiagramRowCount(ByVal batch As CBatchSectionCalculator) As Long
    Dim rowCount As Long
    rowCount = 1

    Dim combinationIndex As Long
    For combinationIndex = 1 To batch.Count
        Dim stateIndex As Long
        For stateIndex = 1 To batch.ResultAt(combinationIndex).StateRepository.StateCount
            Dim stateResult As CSectionStateResult
            Set stateResult = batch.ResultAt(combinationIndex).StateRepository.StateAt(stateIndex)
            If Not stateResult Is Nothing Then
                If stateResult.HasMaterialSnapshot Then rowCount = rowCount + _
                    stateResult.ConcreteDiagramSnapshot.PointCount + stateResult.SteelDiagramSnapshot.PointCount
            End If
        Next stateIndex
    Next combinationIndex

    MaterialDiagramRowCount = rowCount
End Function

' Записывает минимальные свойства imported-preview: Bounds, количество
' элементов и единицы, достаточные для проверки геометрии на схеме.
Private Sub WriteSectionPropertiesPreview(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal snapshot As CSectionPropertiesCalculator, ByVal runID As Variant, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMSectionProperties").RefersToRange

    Dim data() As Variant
    ReDim data(1 To 46, 1 To SECTION_PROPERTY_COLUMNS)
    PutSectionPropertyHeaders data

    Dim rowIndex As Long
    rowIndex = 2

    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim bounds As Variant: bounds = snapshot.SnapshotBounds
    minX = bounds(1): maxX = bounds(2): minY = bounds(3): maxY = bounds(4)

    PutProperty data, rowIndex, runID, "ALL", "Bounds.MinX", OutputLength(minX, units), OutputLengthUnit(units), "Минимальный X бетонной расчетной модели в координатах модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MaxX", OutputLength(maxX, units), OutputLengthUnit(units), "Максимальный X бетонной расчетной модели в координатах модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MinY", OutputLength(minY, units), OutputLengthUnit(units), "Минимальный Y бетонной расчетной модели в координатах модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MaxY", OutputLength(maxY, units), OutputLengthUnit(units), "Максимальный Y бетонной расчетной модели в координатах модели."
    PutProperty data, rowIndex, runID, "ALL", "Geometry.Source", section.SourceType, "-", "Источник геометрии расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Geometry.ConcreteMeshSource", section.ConcreteMeshSource, "-", "Источник бетонной сетки: Generated, AutoCADRegions или AutoCADContours."
    PutProperty data, rowIndex, runID, "ALL", "Geometry.Revision", section.Revision, "-", "Версия геометрии сохраненного снимка."
    PutProperty data, rowIndex, runID, "ALL", "Geometry.RecipeDiagnostics", section.GeometryDiagnosticsText, "-", "Фактическая диагностика сохраненной сборки. StationStep и CenterStep в мм; для связанного ряда StationStep относится к исходному пути."

    Dim concreteProps As CSectionPropertiesCalculator
    Set concreteProps = snapshot
    PutProperty data, rowIndex, runID, "ALL", "ConcreteCount", section.ConcreteCount, "шт", "Количество бетонных элементов расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "RebarCount", section.RebarCount, "шт", "Количество элементов продольной арматуры расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Output.LengthUnit", OutputLengthUnit(units), "-", "Единица вывода длины на момент импорта."
    PutProperty data, rowIndex, runID, "ALL", "Output.AreaUnit", OutputAreaUnit(units), "-", "Единица вывода площади на момент импорта."
    PutProperty data, rowIndex, runID, "ALL", "Output.StressUnit", OutputStressUnit(units), "-", "Единица вывода напряжения на момент импорта."
    PutProperty data, rowIndex, runID, "ALL", "Output.CurvatureUnit", OutputCurvatureUnit(units), "-", "Единица вывода кривизны на момент импорта."
    PutConcreteSectionProperties data, rowIndex, runID, section, concreteProps, units

    target.Resize(UBound(data, 1), SECTION_PROPERTY_COLUMNS).Value2 = data
    FormatResultsBlock target, UBound(data, 1), SECTION_PROPERTY_COLUMNS, "Свойства всего сечения"
End Sub

' ------------------------------
' Размеры, подписи и расчётные области схемы
' ------------------------------
' Записывает аннотации сечения и готовые области трещин по сочетаниям.
' Исходные наружные контуры и отверстия сохраняются отдельно как геометрия.
Private Sub WriteSectionAnnotations(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal runID As Variant, ByVal units As CUnitSystem, Optional ByVal batch As CBatchSectionCalculator = Nothing)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMSectionAnnotations").RefersToRange

    Dim annotations As CSectionAnnotations
    Set annotations = section.Annotations

    Dim rowCount As Long
    rowCount = 1 + annotations.Count
    Dim regions As Collection, region As CConcreteRegion, combination As CCombinationResult, i As Long
    Set regions = New Collection
    If Not batch Is Nothing Then
        For i = 1 To batch.Count
            Set combination = batch.ResultAt(i)
            Set region = combination.CrackResult.Width.InteractionRegion
            If Not region Is Nothing Then
                regions.Add region
                rowCount = rowCount + region.SegmentCount + region.BoundaryProbeCount
            End If
        Next i
    End If
    If rowCount < 1 Then rowCount = 1

    Dim data() As Variant
    ReDim data(1 To rowCount, 1 To ANNOTATION_COLUMNS)
    PutAnnotationHeaders data

    Dim rowIndex As Long
    rowIndex = 2

    For i = 1 To annotations.Count
        PutAnnotationRow data, rowIndex, runID, annotations, i, units
    Next i

    For Each region In regions
        PutCrackRegionRows data, rowIndex, runID, region, units
        PutCrackBoundaryProbeRows data, rowIndex, runID, region, units
    Next region

    target.Resize(rowCount, ANNOTATION_COLUMNS).Value2 = data
    FormatResultsBlock target, rowCount, ANNOTATION_COLUMNS, "Аннотации схемы"
End Sub

' Оставляет таблицу результатов элементов с одной шапкой.
' Для preview нет LC, поэтому записывать Stress/Strain нельзя.
Private Sub WriteEmptyElementResults(ByVal workbook As Object, ByVal runID As Variant, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMElementResults").RefersToRange

    Dim data() As Variant
    ReDim data(1 To 1, 1 To ELEMENT_RESULT_COLUMNS)
    PutElementResultHeaders data, units

    target.Resize(1, ELEMENT_RESULT_COLUMNS).Value2 = data
    FormatResultsBlock target, 1, ELEMENT_RESULT_COLUMNS, "Результаты элементов по сочетаниям"
End Sub

' Очищает прежние аннотации, оставляя строку заголовков.
' Используется, когда в модели нет готовых размеров и подписей.
Private Sub WriteEmptySectionAnnotations(ByVal workbook As Object, ByVal runID As Variant)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMSectionAnnotations").RefersToRange

    Dim data() As Variant
    ReDim data(1 To 1, 1 To ANNOTATION_COLUMNS)
    PutAnnotationHeaders data

    target.Resize(1, ANNOTATION_COLUMNS).Value2 = data
    FormatResultsBlock target, 1, ANNOTATION_COLUMNS, "Аннотации схемы"
End Sub

' Оформляет диапазон как отдельную таблицу с заголовком над именованным началом.
' Дальнейшее чтение использует это имя и размеры таблицы,
' а не область CurrentRegion соседних заполненных ячеек.
Friend Sub FormatResultsBlock(ByVal target As Object, ByVal rowCount As Long, ByVal columnCount As Long, ByVal titleText As String)
    Dim tableRange As Object
    Set tableRange = target.Resize(rowCount, columnCount)

    tableRange.HorizontalAlignment = -4108
    tableRange.VerticalAlignment = -4108
    tableRange.Font.Name = "Arial"
    tableRange.Font.Size = 9
    If rowCount > 1 Then tableRange.Offset(1, 0).Resize(rowCount - 1, columnCount).WrapText = False
    tableRange.Rows.RowHeight = 15
    tableRange.Borders.LineStyle = 1
    tableRange.Borders.Color = RGB(200, 200, 200)
    tableRange.Rows.Item(1).Interior.Color = RGB(217, 217, 217)
    tableRange.Rows.Item(1).Font.Bold = True
    tableRange.Rows.Item(1).HorizontalAlignment = -4131
    tableRange.Columns.Item(1).NumberFormat = "dd.mm.yy"

    AlignCommentColumns tableRange
    ApplyResultNumberFormats tableRange

    If target.Row > 2 Then
        Dim titleRange As Object
        Set titleRange = target.Offset(-1, 0).Resize(1, columnCount)
        titleRange.UnMerge
        titleRange.Clear
        titleRange.Merge
        titleRange.Value2 = titleText
        titleRange.Font.Name = "Arial"
        titleRange.Font.Size = 9
        titleRange.Font.Bold = True
        titleRange.HorizontalAlignment = -4108
        titleRange.VerticalAlignment = -4108
        titleRange.Interior.Color = RGB(217, 217, 217)
        titleRange.Rows.RowHeight = 15
    End If
End Sub

' Задаёт одинаковую точность отображения общих численных столбцов во всех таблицах результатов.
Private Sub ApplyResultNumberFormats(ByVal tableRange As Object)
    Dim headers As Variant
    headers = tableRange.Rows.Item(1).Value2

    Dim colIndex As Long
    For colIndex = 1 To tableRange.Columns.Count
        If StrComp(ResultHeaderBase(CStr(headers(1, colIndex))), "Strain", vbTextCompare) = 0 Then
            tableRange.Columns.Item(colIndex).NumberFormat = "0.0000"
        End If
    Next colIndex
End Sub

' Ищет текстовые диагностические столбцы и выравнивает их вправо.
' Comment и PhysicalState читаются удобнее как текстовые признаки, поэтому
' они не центрируются вместе с числовыми данными таблиц Results.
Private Sub AlignCommentColumns(ByVal tableRange As Object)
    Dim headers As Variant
    headers = tableRange.Rows.Item(1).Value2

    Dim colIndex As Long
    For colIndex = 1 To tableRange.Columns.Count
        If ShouldRightAlignResultColumn(ResultHeaderBase(CStr(headers(1, colIndex)))) Then
            If tableRange.Rows.Count > 1 Then
                tableRange.Offset(1, colIndex - 1).Resize(tableRange.Rows.Count - 1, 1).HorizontalAlignment = -4152
            End If
        End If
    Next colIndex
End Sub

' Определяет текстовые столбцы Results, которые должны читаться как подпись,
' а не как числовой параметр. Пока это комментарии и физическое состояние
' элемента в rngNDMElementResults.
Private Function ShouldRightAlignResultColumn(ByVal headerBase As String) As Boolean
    ShouldRightAlignResultColumn = _
        (StrComp(headerBase, "Comment", vbTextCompare) = 0 Or _
         StrComp(headerBase, "PhysicalState", vbTextCompare) = 0)
End Function

' Возвращает имя столбца без единицы измерения.
' Например, заголовок "Stress, MPa" для служебных проверок считается просто
' "Stress"; это позволяет одинаково находить Comment и другие столбцы при
' разных пользовательских единицах вывода.
Private Function ResultHeaderBase(ByVal headerText As String) As String
    Dim commaPos As Long
    commaPos = InStr(1, headerText, ",", vbTextCompare)
    If commaPos > 0 Then
        ResultHeaderBase = Trim$(Left$(headerText, commaPos - 1))
    Else
        ResultHeaderBase = Trim$(headerText)
    End If
End Function

' Записывает готовые величины элементов сохранённого НДС.
' Данные уже подготовлены состоянием; здесь только переводятся напряжения.
Private Sub PutCombinationElementResults(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal section As CSectionModel, ByVal batch As CBatchSectionCalculator, ByVal combinationIndex As Long, _
        ByVal stateResult As CSectionStateResult, ByVal runID As Variant, ByVal units As CUnitSystem)
    If stateResult Is Nothing Then Exit Sub
    If Not stateResult.Converged Then Exit Sub
    Dim prepared As Variant, i As Long
    prepared = stateResult.PrepareElementSnapshot(section)
    For i = 1 To UBound(prepared, 1)
        PutElementRow data, rowIndex, runID, batch.CombinationID(combinationIndex), _
            batch.ResultAt(combinationIndex).ProfileId, stateResult.StateTypeText, CStr(prepared(i, 1)), CStr(prepared(i, 2)), _
            CDbl(prepared(i, 3)), CDbl(prepared(i, 4)), CStr(prepared(i, 5)), units
        rowIndex = rowIndex + 1
    Next i
End Sub

' Кладет одну строку rngNDMElementResults в общий массив перед блочной записью на лист.
Private Sub PutElementRow(ByRef data() As Variant, ByVal rowIndex As Long, ByVal runID As Variant, _
        ByVal combinationID As String, ByVal profileId As String, ByVal stateType As String, ByVal elementID As String, _
        ByVal materialType As String, ByVal strain As Double, ByVal stress As Double, ByVal state As String, ByVal units As CUnitSystem)
    data(rowIndex, 1) = runID
    data(rowIndex, 2) = combinationID
    data(rowIndex, 3) = profileId
    data(rowIndex, 4) = stateType
    data(rowIndex, 5) = elementID
    data(rowIndex, 6) = materialType
    data(rowIndex, 7) = strain
    data(rowIndex, 8) = OutputStress(stress, units)
    data(rowIndex, 9) = state
End Sub


' Формирует заголовки таблицы LC-зависимых результатов элементов.
Private Sub PutElementResultHeaders(ByRef data() As Variant, ByVal units As CUnitSystem)
    data(1, 1) = "RunID"
    data(1, 2) = "LoadCase"
    data(1, 3) = "ProfileId"
    data(1, 4) = "StateType"
    data(1, 5) = "ElementID"
    data(1, 6) = "MaterialType"
    data(1, 7) = "Strain"
    data(1, 8) = "Stress, " & OutputStressUnit(units)
    data(1, 9) = "PhysicalState"
End Sub

' Готовит заголовки таблицы геометрии, где рядом показываются
' фактические площадь и инерции и размеры графической границы элемента.
Private Sub PutGeometryHeaders(ByRef data() As Variant, ByVal units As CUnitSystem)
    data(1, 1) = "RunID"
    data(1, 2) = "ElementID"
    data(1, 3) = "MaterialType"
    data(1, 4) = "X, " & OutputLengthUnit(units)
    data(1, 5) = "Y, " & OutputLengthUnit(units)
    data(1, 6) = "Area, " & OutputAreaUnit(units)
    data(1, 7) = "GeometryInterpretationStatus"
    data(1, 8) = "Width, " & OutputLengthUnit(units)
    data(1, 9) = "Height, " & OutputLengthUnit(units)
    data(1, 10) = "Diameter, " & OutputLengthUnit(units)
    data(1, 11) = "Rotation, rad"
    data(1, 12) = "LocalIx, " & OutputLengthUnit(units) & "4"
    data(1, 13) = "LocalIy, " & OutputLengthUnit(units) & "4"
    data(1, 14) = "LocalIxy, " & OutputLengthUnit(units) & "4"
    data(1, 15) = "Comment"
End Sub

' Формирует заголовки таблицы общих и LC-зависимых свойств сечения.
Private Sub PutSectionPropertyHeaders(ByRef data() As Variant)
    data(1, 1) = "RunID"
    data(1, 2) = "LoadCase"
    data(1, 3) = "Parameter"
    data(1, 4) = "Value"
    data(1, 5) = "Unit"
    data(1, 6) = "Comment"
End Sub

' Формирует заголовки таблицы размеров, подписей и других аннотаций схемы и экспорта AutoCAD.
Private Sub PutAnnotationHeaders(ByRef data() As Variant)
    data(1, 1) = "RunID"
    data(1, 2) = "AnnotationType"
    data(1, 3) = "AnnotationID"
    data(1, 4) = "StartX"
    data(1, 5) = "StartY"
    data(1, 6) = "EndX"
    data(1, 7) = "EndY"
    data(1, 8) = "OutsideNormalX"
    data(1, 9) = "OutsideNormalY"
    data(1, 10) = "Text"
    data(1, 11) = "Value"
    data(1, 12) = "Unit"
    data(1, 13) = "Comment"
    data(1, 14) = "CombinationID"
    data(1, 15) = "Standard"
    data(1, 16) = "AnchorID"
    data(1, 17) = "LoopID"
    data(1, 18) = "LoopRole"
    data(1, 19) = "SectionRevision"
    data(1, 20) = "RegionArea"
    data(1, 21) = "AreaUnit"
    data(1, 22) = "GeometrySource"
End Sub

' Записывает готовую расчётную область без нового построения или отбора арматуры.
' Каждый замкнутый контур, включая отверстия, получает свой LoopID.
' Области CRACK_REGION_* не используются как исходные границы бетона.
Private Sub PutCrackRegionRows(ByRef data() As Variant, ByRef rowIndex As Long, ByVal runID As Variant, _
        ByVal region As CConcreteRegion, ByVal units As CUnitSystem)
    Dim i As Long, loopID As Long, x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim cx As Double, cy As Double, sweep As Double
    For i = 1 To region.SegmentCount
        region.GetSegment i, x1, y1, x2, y2, cx, cy, sweep, loopID
        data(rowIndex, 1) = runID
        data(rowIndex, 2) = "CRACK_REGION_LINE"
        data(rowIndex, 3) = "CRACK_" & region.OwnerID & "_" & region.Standard & "_" & CStr(loopID) & "_" & CStr(i)
        data(rowIndex, 4) = OutputLength(x1, units): data(rowIndex, 5) = OutputLength(y1, units)
        data(rowIndex, 6) = OutputLength(x2, units): data(rowIndex, 7) = OutputLength(y2, units)
        If Abs(sweep) > 0.0000000001 Then
            data(rowIndex, 2) = "CRACK_REGION_ARC"
            data(rowIndex, 8) = OutputLength(cx, units): data(rowIndex, 9) = OutputLength(cy, units)
            data(rowIndex, 10) = Replace$(Trim$(Str$(sweep)), ",", ".")
        End If
        data(rowIndex, 12) = OutputLengthUnit(units)
        data(rowIndex, 13) = "Граница сохраненной расчетной области нормальных трещин."
        data(rowIndex, 14) = region.OwnerID: data(rowIndex, 15) = region.Standard
        data(rowIndex, 16) = region.AnchorID: data(rowIndex, 17) = loopID
        data(rowIndex, 18) = "OUTER"
        If region.LoopIsOpening(loopID) Then data(rowIndex, 18) = "OPENING"
        data(rowIndex, 19) = region.SectionRevision
        data(rowIndex, 20) = OutputArea(region.Area, units): data(rowIndex, 21) = OutputAreaUnit(units)
        data(rowIndex, 22) = region.Source
        rowIndex = rowIndex + 1
    Next i
End Sub

' Записывает уже полученные P/Q отдельно от контурных сегментов. Эти строки
' не образуют замкнутые контуры и не используются для повторного вычисления площади;
' отсутствующий локальный интервал сохраняется явно, а не заменяется нулями.
Private Sub PutCrackBoundaryProbeRows(ByRef data() As Variant, ByRef rowIndex As Long, ByVal runID As Variant, _
        ByVal region As CConcreteRegion, ByVal units As CUnitSystem)
    Dim probes As Variant, i As Long
    probes = region.BoundaryProbes
    For i = 1 To region.BoundaryProbeCount
        data(rowIndex, 1) = runID: data(rowIndex, 2) = "CRACK_PROBE"
        data(rowIndex, 3) = "CRACK_PROBE_" & region.OwnerID & "_" & region.Standard & "_" & CStr(i)
        data(rowIndex, 10) = "P" & CStr(i) & "/Q" & CStr(i)
        data(rowIndex, 11) = i: data(rowIndex, 12) = OutputLengthUnit(units)
        data(rowIndex, 13) = "Локальный интервал материала не найден; переход через проем не выполняется."
        If CBool(probes(i, 1)) Then
            data(rowIndex, 4) = OutputLength(CDbl(probes(i, 4)), units): data(rowIndex, 5) = OutputLength(CDbl(probes(i, 5)), units)
            data(rowIndex, 6) = OutputLength(CDbl(probes(i, 2)), units): data(rowIndex, 7) = OutputLength(CDbl(probes(i, 3)), units)
            data(rowIndex, 13) = "Точки P и Q ограничивают участок бетона по нормали через выбранную арматуру."
        End If
        data(rowIndex, 14) = region.OwnerID: data(rowIndex, 15) = region.Standard: data(rowIndex, 16) = region.AnchorID
        data(rowIndex, 19) = region.SectionRevision
        data(rowIndex, 20) = OutputArea(region.Area, units): data(rowIndex, 21) = OutputAreaUnit(units)
        data(rowIndex, 22) = region.Source
        rowIndex = rowIndex + 1
    Next i
End Sub

' Добавляет одну строку свойства сечения в массив rngNDMSectionProperties.
Private Sub PutProperty(ByRef data() As Variant, ByRef rowIndex As Long, ByVal runID As Variant, _
        ByVal loadCase As String, ByVal parameter As String, ByVal value As Variant, _
        ByVal unitText As String, ByVal comment As String)
    data(rowIndex, 1) = runID
    data(rowIndex, 2) = loadCase
    data(rowIndex, 3) = parameter
    data(rowIndex, 4) = value
    data(rowIndex, 5) = unitText
    data(rowIndex, 6) = comment
    rowIndex = rowIndex + 1
End Sub

' Добавляет одну аннотацию в массив rngNDMSectionAnnotations, переводя координаты в единицы
' вывода.
Private Sub PutAnnotationRow(ByRef data() As Variant, ByRef rowIndex As Long, ByVal runID As Variant, _
        ByVal annotations As CSectionAnnotations, ByVal annotationIndex As Long, ByVal units As CUnitSystem)
    Dim annotationType As String
    annotationType = annotations.AnnotationType(annotationIndex)

    data(rowIndex, 1) = runID
    data(rowIndex, 2) = annotationType
    data(rowIndex, 3) = annotations.AnnotationID(annotationIndex)
    data(rowIndex, 4) = OutputLength(annotations.StartX(annotationIndex), units)
    data(rowIndex, 5) = OutputLength(annotations.StartY(annotationIndex), units)
    data(rowIndex, 6) = OutputLength(annotations.EndX(annotationIndex), units)
    data(rowIndex, 7) = OutputLength(annotations.EndY(annotationIndex), units)
    data(rowIndex, 8) = OutputAnnotationNormalX(annotations, annotationIndex, units)
    data(rowIndex, 9) = OutputAnnotationNormalY(annotations, annotationIndex, units)
    data(rowIndex, 10) = OutputAnnotationText(annotations, annotationIndex, units)
    data(rowIndex, 11) = OutputAnnotationValue(annotations, annotationIndex, units)
    data(rowIndex, 12) = OutputAnnotationUnit(annotations, annotationIndex, units)
    data(rowIndex, 13) = OutputAnnotationComment(annotations, annotationIndex)
    rowIndex = rowIndex + 1
End Sub

' Для обычных dimension/rebar-аннотаций NormalX/NormalY - безразмерная
' нормаль. Для CONTOUR_ARC в этих полях хранится центр дуги, поэтому его
' нужно выводить в тех же единицах длины, что Start/End.
Private Function OutputAnnotationNormalX(ByVal annotations As CSectionAnnotations, ByVal annotationIndex As Long, _
        ByVal units As CUnitSystem) As Variant
    If StrComp(annotations.AnnotationType(annotationIndex), "CONTOUR_ARC", vbTextCompare) = 0 Then
        OutputAnnotationNormalX = OutputLength(annotations.NormalX(annotationIndex), units)
    Else
        OutputAnnotationNormalX = annotations.NormalX(annotationIndex)
    End If
End Function

' Смотри OutputAnnotationNormalX: для дуги это координата центра, а для
' остальных annotation - безразмерная нормаль.
Private Function OutputAnnotationNormalY(ByVal annotations As CSectionAnnotations, ByVal annotationIndex As Long, _
        ByVal units As CUnitSystem) As Variant
    If StrComp(annotations.AnnotationType(annotationIndex), "CONTOUR_ARC", vbTextCompare) = 0 Then
        OutputAnnotationNormalY = OutputLength(annotations.NormalY(annotationIndex), units)
    Else
        OutputAnnotationNormalY = annotations.NormalY(annotationIndex)
    End If
End Function

' Готовит текст аннотации в пользовательских единицах. Размер на схеме
' округляется до целых миллиметров, а численное поле Value
' сохраняет исходную точность для проверки результата.
Private Function OutputAnnotationText(ByVal annotations As CSectionAnnotations, ByVal annotationIndex As Long, _
        ByVal units As CUnitSystem) As String
    Dim annotationType As String
    annotationType = annotations.AnnotationType(annotationIndex)
    If StrComp(annotationType, "DIMENSION", vbTextCompare) <> 0 And StrComp(annotationType, "MEASUREMENT_LABEL", vbTextCompare) <> 0 Then
        OutputAnnotationText = annotations.Text(annotationIndex)
        Exit Function
    End If

    Dim valueText As String
    valueText = FormatRoundedAnnotationLength(annotations.Value(annotationIndex), units)
    Dim unitText As String
    unitText = OutputLengthUnit(units)
    If StrComp(Trim$(annotations.Text(annotationIndex)), ApproxSign(), vbTextCompare) = 0 Then
        OutputAnnotationText = ApproxSign() & " " & valueText & " " & unitText
    ElseIf Len(Trim$(annotations.Text(annotationIndex))) > 0 Then
        OutputAnnotationText = annotations.Text(annotationIndex) & "=" & valueText & " " & unitText
    Else
        OutputAnnotationText = valueText & " " & unitText
    End If
End Function

' Округляет видимую подпись размерной линии до ближайшего целого миллиметра,
' а затем переводит округленное значение в выбранные пользователем единицы.
' Поэтому 500 мм при выводе в метрах пишется как 0.5 m, а не как 1 m или 0 m.
' VBA.Round использует банковское округление, поэтому здесь явно применено
' инженерно ожидаемое правило: половина миллиметра округляется от нуля.
Private Function FormatRoundedAnnotationLength(ByVal internalMmValue As Double, ByVal units As CUnitSystem) As String
    Dim roundedMmValue As Double
    If internalMmValue >= 0# Then
        roundedMmValue = Int(internalMmValue + 0.5)
    Else
        roundedMmValue = -Int(Abs(internalMmValue) + 0.5)
    End If

    FormatRoundedAnnotationLength = FormatNumberInvariant(OutputLength(roundedMmValue, units))
End Function

' Преобразует служебные сообщения построителей геометрии
' в понятные пользователю пояснения Results.
Private Function OutputAnnotationComment(ByVal annotations As CSectionAnnotations, ByVal annotationIndex As Long) As String
    Dim commentValue As String
    commentValue = Trim$(annotations.Comment(annotationIndex))

    If StrComp(commentValue, "Generated section contour segment.", vbTextCompare) = 0 Then
        OutputAnnotationComment = "Сегмент контура параметрического сечения."
    ElseIf StrComp(commentValue, "Generated circle contour.", vbTextCompare) = 0 Then
        OutputAnnotationComment = "Контур круглого параметрического сечения."
    ElseIf StrComp(commentValue, "Section dimension.", vbTextCompare) = 0 Then
        OutputAnnotationComment = "Размер сечения для отображения на схеме."
    ElseIf StrComp(commentValue, "Grouped rebar annotation.", vbTextCompare) = 0 Then
        OutputAnnotationComment = "Подпись общего количества и диаметров арматуры у выбранного участка контура."
    ElseIf StrComp(commentValue, "Overall concrete width.", vbTextCompare) = 0 Then
        OutputAnnotationComment = "Общая ширина бетонного сечения."
    ElseIf StrComp(commentValue, "Overall concrete height.", vbTextCompare) = 0 Then
        OutputAnnotationComment = "Общая высота бетонного сечения."
    ElseIf StrComp(commentValue, "Upper rectangle width.", vbTextCompare) = 0 Then
        OutputAnnotationComment = "Ширина верхнего прямоугольника Г-образного сечения."
    ElseIf StrComp(commentValue, "Lower rectangle height.", vbTextCompare) = 0 Then
        OutputAnnotationComment = "Высота нижнего прямоугольника Г-образного сечения."
    ElseIf Len(commentValue) > 0 Then
        OutputAnnotationComment = commentValue
    Else
        OutputAnnotationComment = "Описание элемента оформления схемы: размера или подписи."
    End If
End Function


' Возвращает знак приблизительного равенства без литерала в исходнике. Так VBA-модуль
' не зависит от кодировки файла, а на лист Results попадает настоящий знак приблизительного равенства.
Private Function ApproxSign() As String
    ApproxSign = ChrW$(&H2248)
End Function

' Переводит численное значение аннотации, если оно задано как длина.
Private Function OutputAnnotationValue(ByVal annotations As CSectionAnnotations, ByVal annotationIndex As Long, _
        ByVal units As CUnitSystem) As Variant
    If StrComp(annotations.Unit(annotationIndex), "Length", vbTextCompare) = 0 Then
        OutputAnnotationValue = OutputLength(annotations.Value(annotationIndex), units)
    ElseIf Len(annotations.Unit(annotationIndex)) = 0 Then
        OutputAnnotationValue = vbNullString
    Else
        OutputAnnotationValue = annotations.Value(annotationIndex)
    End If
End Function

' Возвращает пользовательскую единицу для длиновых аннотаций и исходную подпись для прочих типов.
Private Function OutputAnnotationUnit(ByVal annotations As CSectionAnnotations, ByVal annotationIndex As Long, _
        ByVal units As CUnitSystem) As String
    If StrComp(annotations.Unit(annotationIndex), "Length", vbTextCompare) = 0 Then
        OutputAnnotationUnit = OutputLengthUnit(units)
    Else
        OutputAnnotationUnit = annotations.Unit(annotationIndex)
    End If
End Function

' Ниже переводятся физические величины сохранённого результата в выбранные единицы вывода.
Private Function OutputLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputLength = value Else OutputLength = units.InternalLengthToOutput(value)
End Function

' Переводит площадь из внутренних мм2 в пользовательскую единицу.
Private Function OutputArea(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputArea = value Else OutputArea = units.InternalAreaToOutput(value)
End Function

' Переводит момент инерции из внутренних мм4 в пользовательскую единицу длины в четвертой степени.
Private Function OutputFourthPowerLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputFourthPowerLength = value Else OutputFourthPowerLength = units.InternalFourthPowerLengthToOutput(value)
End Function

' Переводит продольную силу из внутренних Н в пользовательскую систему вывода.
Private Function OutputForce(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputForce = value Else OutputForce = units.InternalForceToOutput(value)
End Function

' Переводит Mx с учетом пользовательской системы знаков.
Private Function OutputMx(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMx = value Else OutputMx = units.InternalMomentMxToOutput(value)
End Function

' Переводит My с учетом пользовательской системы знаков.
Private Function OutputMy(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMy = value Else OutputMy = units.InternalMomentMyToOutput(value)
End Function

' Переводит напряжение из внутренних МПа в пользовательскую единицу.
Private Function OutputStress(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputStress = value Else OutputStress = units.InternalStressToOutput(value)
End Function

' Переводит кривизну из внутренних 1/мм в пользовательскую единицу.
Private Function OutputCurvature(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputCurvature = value Else OutputCurvature = units.InternalCurvatureToOutput(value)
End Function

' Возвращает подпись выходной единицы длины для заголовков Results.
Private Function OutputLengthUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputLengthUnit = "mm" Else OutputLengthUnit = units.OutputLengthUnit
End Function

' Возвращает подпись выходной единицы площади для заголовков Results.
Private Function OutputAreaUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputAreaUnit = "mm2" Else OutputAreaUnit = units.OutputAreaUnit
End Function

' Возвращает подпись выходной единицы силы для заголовков Results.
Private Function OutputForceUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputForceUnit = "N" Else OutputForceUnit = units.OutputForceUnit
End Function

' Возвращает подпись выходной единицы момента для заголовков Results.
Private Function OutputMomentUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputMomentUnit = "N*mm" Else OutputMomentUnit = units.OutputMomentUnit
End Function

' Возвращает подпись выходной единицы напряжения для заголовков Results.
Private Function OutputStressUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputStressUnit = "MPa" Else OutputStressUnit = units.OutputStressUnit
End Function

' Возвращает подпись выходной единицы кривизны для заголовков Results.
Private Function OutputCurvatureUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputCurvatureUnit = "1/mm" Else OutputCurvatureUnit = units.OutputCurvatureUnit
End Function

' Возвращает пользовательское соглашение знаков для N справочной строкой в Results.
Private Function OutputSignN(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputSignN = "Tension" Else OutputSignN = units.OutputSignN
End Function

' Возвращает пользовательское соглашение знаков для Mx справочной строкой в Results.
Private Function OutputSignMx(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputSignMx = "+Y tension" Else OutputSignMx = units.OutputSignMx
End Function

' Возвращает пользовательское соглашение знаков для My справочной строкой в Results.
Private Function OutputSignMy(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputSignMy = "+X tension" Else OutputSignMy = units.OutputSignMy
End Function

' Форматирует числа для подписей размеров с точкой как десятичным разделителем.
Private Function FormatNumberInvariant(ByVal value As Double) As String
    Dim textValue As String
    textValue = Replace$(Format$(value, "0.###"), ",", ".")
    Do While InStr(1, textValue, ".", vbTextCompare) > 0 And Right$(textValue, 1) = "0"
        textValue = Left$(textValue, Len(textValue) - 1)
    Loop
    If Right$(textValue, 1) = "." Then textValue = Left$(textValue, Len(textValue) - 1)
    FormatNumberInvariant = textValue
End Function

' Возвращает максимум двух чисел для определения границ таблиц и значений результата.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

' Возвращает минимум двух чисел для определения границ таблиц и значений результата.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function
