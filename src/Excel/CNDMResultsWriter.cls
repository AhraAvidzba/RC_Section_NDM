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
' Запись расчетного снимка на Results
' ==========================================================================
' Класс получает уже посчитанную модель сечения и результаты batch-расчета,
' переводит их в выбранные OUTPUT-единицы и записывает на лист Results одним
' согласованным снимком последнего успешного расчета.
'
' Writer не запускает solver, не генерирует геометрию и не читает текущие
' Inputs. Он только сериализует то состояние, которое уже рассчитано.
'
' Этот writer не является подробной таблицей расчета по модели прочности:
' он пишет общий snapshot, включая строки named-state StrengthState,
' CapacityState, CrackedState и т.п. Отдельный блок "НДС и несущая
' способность по модели прочности" от rngStrengthSummaryAnchor оформляет
' CStrengthSummaryWriter.

Private Const ELEMENT_RESULT_COLUMNS As Long = 9 ' LC-зависимые результаты named-state: RunID, LoadCase, ProfileId, StateType, ElementID, MaterialType, Strain, Stress, PhysicalState.
Private Const GEOMETRY_COLUMNS As Long = 15 ' Постоянная геометрия элементов, записанная один раз на расчетный снимок.
Private Const SECTION_PROPERTY_COLUMNS As Long = 6 ' Общие свойства сечения и LC-свойства уровня всего сечения.
Private Const ANNOTATION_COLUMNS As Long = 13 ' Semantic-аннотации схемы: размеры, контур, групповые подписи арматуры.
Private Const MATERIAL_DIAGRAM_COLUMNS As Long = 11 ' Каталог уникальных точек диаграмм, переданных solver-ам.
Private Const PHYSICAL_TOLERANCE As Double = 0.000000001 ' Порог, ниже которого напряжение считается NearZero.
Private Const COMMON_SECTION_PROPERTY_ROWS As Long = 61 ' Строки общих свойств сечения до LC-зависимых блоков.
Private Const LC_SECTION_PROPERTY_ROWS As Long = 23 ' Строки свойств одного LC без metadata named-state.
Private Const STATE_SECTION_PROPERTY_ROWS As Long = 12 ' Строки metadata одного сохраненного named-state.
Private Const RESULTS_COLUMN_WIDTH As Double = 10# ' Единая ширина всех столбцов Results.
Private Const SNAPSHOT_TITLE_OFFSET As Long = 2 ' Общая синяя строка находится на две строки выше нижних snapshot-якорей.

' Возвращает фактическую ширину одного из блоков Results, которые пишет этот writer.
Public Function OutputColumnCount(ByVal rangeName As String) As Long
    Select Case LCase$(Trim$(rangeName))
        Case "rngndmelementresults"
            OutputColumnCount = ELEMENT_RESULT_COLUMNS
        Case "rngndmsectiongeometry"
            OutputColumnCount = GEOMETRY_COLUMNS
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

' Оценивает максимальную высоту rngNDMElementResults до запуска solver-а.
Public Function EstimatedElementResultRows(ByVal elementCount As Long, ByVal namedStateCount As Long) As Long
    If elementCount < 0 Then elementCount = 0
    If namedStateCount < 0 Then namedStateCount = 0
    EstimatedElementResultRows = 1 + elementCount * namedStateCount
End Function

' Оценивает высоту rngNDMSectionProperties до расчета по максимальному числу named-state.
Public Function EstimatedSectionPropertyRows(ByVal combinationCount As Long, ByVal namedStateCount As Long) As Long
    If combinationCount < 0 Then combinationCount = 0
    If namedStateCount < 0 Then namedStateCount = 0
    EstimatedSectionPropertyRows = 1 + COMMON_SECTION_PROPERTY_ROWS + _
        combinationCount * LC_SECTION_PROPERTY_ROWS + namedStateCount * STATE_SECTION_PROPERTY_ROWS
End Function

' ------------------------------
' Полная запись Results
' ------------------------------
' Очищает текущий снимок и заново пишет согласованные блоки:
' geometry, element results, section properties, material diagrams и annotations.
Public Sub WriteResults(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal batch As CBatchSectionCalculator, _
        Optional ByVal units As CUnitSystem = Nothing)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4450, "CNDMResultsWriter", "Книга Excel не передана."
    If section Is Nothing Then Err.Raise vbObjectError + 4451, "CNDMResultsWriter", "Модель сечения не передана."
    If materialProvider Is Nothing Then Err.Raise vbObjectError + 4452, "CNDMResultsWriter", "Поставщик материалов не передан."
    If batch Is Nothing Then Err.Raise vbObjectError + 4454, "CNDMResultsWriter", "Пакетный расчетчик не передан."

    ClearResults workbook

    Dim runID As Date
    runID = Now

    WriteSnapshotTitle workbook
    WriteSectionGeometry workbook, section, runID, units
    WriteElementResults workbook, section, materialProvider, batch, runID, units
    WriteSectionProperties workbook, section, materialProvider, batch, runID, units
    WriteMaterialDiagrams workbook, materialProvider, batch, runID, units
    EnsureAutoCADBoundsAnnotations section
    WriteSectionAnnotations workbook, section, runID, units
End Sub

' Записывает на Results только импортированную геометрию без расчетных LC.
' Такой снимок нужен после кнопки AutoCAD import: пользователь сразу видит
' элементы на схеме, а последующий расчет берет эту же геометрию из Results.
Public Sub WriteGeometryPreview(ByVal workbook As Object, ByVal section As CSectionModel, _
        Optional ByVal units As CUnitSystem = Nothing)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4450, "CNDMResultsWriter", "Книга Excel не передана."
    If section Is Nothing Then Err.Raise vbObjectError + 4451, "CNDMResultsWriter", "Модель сечения не передана."

    ClearResults workbook

    Dim runID As Date
    runID = Now

    WriteSnapshotTitle workbook
    WriteSectionGeometry workbook, section, runID, units
    WriteEmptyElementResults workbook, runID, units
    WriteSectionPropertiesPreview workbook, section, runID, units
    EnsureAutoCADBoundsAnnotations section
    WriteSectionAnnotations workbook, section, runID, units
End Sub

' Очищает только расчетные блоки Results перед новым снимком.
Public Sub ClearResults(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4455, "CNDMResultsWriter", "Книга Excel не передана."

    ClearSnapshotTitle workbook
    ClearNamedRegion workbook, "rngNDMElementResults"
    ClearNamedRegion workbook, "rngNDMSectionGeometry"
    ClearNamedRegion workbook, "rngNDMSectionProperties"
    ClearNamedRegion workbook, "rngNDMSectionAnnotations"
    ClearNamedRegion workbook, "rngNDMMaterialDiagrams"
End Sub

' Безопасно чистит только фактический прямоугольник одной snapshot-таблицы.
' Ширина берется из схемы writer-а, а высота - по заполненному первому
' столбцу старой таблицы. CurrentRegion здесь не используется: он может
' захватить соседние Results-блоки из-за старого форматирования листа.
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

' Определяет высоту уже записанной anchored-таблицы по первому столбцу.
' Все нижние Results-таблицы имеют непустой первый столбец в строке шапки
' и в каждой строке данных, поэтому первая пустая ячейка завершает блок.
Private Function ExistingAnchoredRowCount(ByVal anchor As Object) As Long
    Dim maxOffset As Long
    maxOffset = anchor.Worksheet.Rows.Count - anchor.Row

    Dim rowOffset As Long
    For rowOffset = 0 To maxOffset
        If Len(Trim$(CStr(anchor.Offset(rowOffset, 0).Value2))) = 0 Then Exit For
        ExistingAnchoredRowCount = ExistingAnchoredRowCount + 1
    Next rowOffset
End Function

' Записывает общий заголовок над нижними snapshot-таблицами Results.
Private Sub WriteSnapshotTitle(ByVal workbook As Object)
    On Error GoTo MissingRange
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngNDMElementResults").RefersToRange
    If anchor.Row <= SNAPSHOT_TITLE_OFFSET Then Exit Sub

    Dim titleRow As Long
    titleRow = anchor.Row - SNAPSHOT_TITLE_OFFSET
    anchor.Worksheet.Rows.Item(titleRow).Interior.Color = RGB(189, 215, 238)
    anchor.Worksheet.Rows.Item(titleRow).RowHeight = 15
    With anchor.Worksheet.Cells(titleRow, anchor.Column)
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

' Очищает общую строку-заголовок snapshot-блоков перед новой записью.
Private Sub ClearSnapshotTitle(ByVal workbook As Object)
    On Error GoTo MissingRange
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngNDMElementResults").RefersToRange
    If anchor.Row <= SNAPSHOT_TITLE_OFFSET Then Exit Sub
    anchor.Worksheet.Rows.Item(anchor.Row - SNAPSHOT_TITLE_OFFSET).Clear
    anchor.Worksheet.Rows.Item(anchor.Row - SNAPSHOT_TITLE_OFFSET).RowHeight = 15
MissingRange:
End Sub

' ------------------------------
' LC-зависимые результаты элементов
' ------------------------------
' В этот блок не пишутся координаты и площади: они уже есть в
' rngNDMSectionGeometry и связываются по ElementID.
Private Sub WriteElementResults(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal batch As CBatchSectionCalculator, _
        ByVal runID As Variant, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMElementResults").RefersToRange

    Dim elementCount As Long
    elementCount = section.ConcreteCount + section.RebarCount

    Dim stateCount As Long
    Dim i As Long
    ' Считаем только конечные named states. Один LC может дать несколько
    ' состояний: StrengthState, CapacityState, CrackedState, BeforeMcrcState и AfterMcrcState.
    For i = 1 To batch.Count
        stateCount = stateCount + batch.NamedStateCount(i)
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
        For stateIndex = 1 To batch.NamedStateCount(i)
            PutCombinationElementResults data, rowIndex, section, materialProvider, batch, i, _
                batch.NamedState(i, stateIndex), runID, units
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

    ' Арматура хранится в той же таблице геометрии, что и бетонные волокна, но с диаметром
    ' вместо width/height: downstream-слои различают тип элемента по MaterialType.
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
        rowIndex = rowIndex + 1
    Next i

    target.Resize(rowCount, GEOMETRY_COLUMNS).Value2 = data
    FormatResultsBlock target, rowCount, GEOMETRY_COLUMNS, "Расчетная геометрия элементов"
End Sub

' ------------------------------
' Свойства сечения и LC-состояния
' ------------------------------
' Здесь лежат свойства уровня всей модели: Bounds, бетонные и приведенные
' геометрические характеристики, сохраненные единицы вывода и плоскости
' деформаций named-state каждого LC. Нагрузки пользователя задаются от
' бетонного центра, а главные оси схемы и AutoCAD читаются из блока
' Transformed.* расчетного snapshot.
Private Sub WriteSectionProperties(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal batch As CBatchSectionCalculator, _
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
    GetConcreteBounds section, minX, maxX, minY, maxY

    ' Бетонные свойства нужны как база пользовательской точки нагрузки.
    ' Приведенные свойства сохраняются отдельно: они используются для справки,
    ' главных осей и совместимости существующего plot/export layer. Для
    ' геометрического приведения берем обычные модули Eb/Es, а не наклон
    ' выбранной диаграммы прочности: форма диаграммы не должна менять Ared/Ired.
    Dim concreteProps As CSectionPropertiesCalculator
    Set concreteProps = New CSectionPropertiesCalculator
    concreteProps.CalculateConcrete section

    Dim transformedProps As CSectionPropertiesCalculator
    Set transformedProps = New CSectionPropertiesCalculator
    transformedProps.CalculateTransformedByModuli section, _
        materialProvider.ConcreteElasticModulusFromValueSet(mvsULS), _
        materialProvider.SteelElasticModulusFromValueSet(mvsULS)

    PutProperty data, rowIndex, runID, "ALL", "Bounds.MinX", OutputLength(minX, units), OutputLengthUnit(units), "Минимальный X расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MaxX", OutputLength(maxX, units), OutputLengthUnit(units), "Максимальный X расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MinY", OutputLength(minY, units), OutputLengthUnit(units), "Минимальный Y расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MaxY", OutputLength(maxY, units), OutputLengthUnit(units), "Максимальный Y расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Geometry.Source", section.SourceType, "-", "Источник геометрии расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "LoadReferenceX", OutputLength(batch.LoadReferenceX, units), OutputLengthUnit(units), "Общая X точки приложения нагрузки относительно центра тяжести бетонного сечения."
    PutProperty data, rowIndex, runID, "ALL", "LoadReferenceY", OutputLength(batch.LoadReferenceY, units), OutputLengthUnit(units), "Общая Y точки приложения нагрузки относительно центра тяжести бетонного сечения."
    PutConcreteSectionProperties data, rowIndex, runID, section, concreteProps, units
    PutTransformedSectionProperties data, rowIndex, runID, section, transformedProps, units
    PutGeneralSectionCheckProperties data, rowIndex, runID, section, concreteProps, transformedProps, batch

    PutProperty data, rowIndex, runID, "ALL", "ConcreteCount", section.ConcreteCount, "шт", "Количество бетонных элементов."
    PutProperty data, rowIndex, runID, "ALL", "RebarCount", section.RebarCount, "шт", "Количество элементов арматуры."
    PutProperty data, rowIndex, runID, "ALL", "Output.LengthUnit", OutputLengthUnit(units), "-", "Единица вывода длины на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.AreaUnit", OutputAreaUnit(units), "-", "Единица вывода площади на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.ForceUnit", OutputForceUnit(units), "-", "Единица вывода силы на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.MomentUnit", OutputMomentUnit(units), "-", "Единица вывода момента на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.StressUnit", OutputStressUnit(units), "-", "Единица вывода напряжения на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.CurvatureUnit", OutputCurvatureUnit(units), "-", "Единица вывода кривизны на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.SignConvention.N", OutputSignN(units), "-", "Пользовательский знак N на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.SignConvention.Mx", OutputSignMx(units), "-", "Пользовательский знак Mx на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.SignConvention.My", OutputSignMy(units), "-", "Пользовательский знак My на момент расчета."
    Dim i As Long
    ' LC-зависимые свойства уровня сечения хранятся отдельно от поэлементных
    ' строк. Плоскости деформаций пишутся ниже только как State.<StateType>.*,
    ' чтобы схема и AutoCAD всегда выбирали конкретное named-state без
    ' неоднозначных legacy-полей.
    For i = 1 To batch.Count
        If batch.StateAvailable(i) Then
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "Nint", OutputForce(batch.Nint(i), units), OutputForceUnit(units), "Внутренняя продольная сила."
        Else
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "Nint", vbNullString, OutputForceUnit(units), "Внутренняя продольная сила."
        End If
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "N", OutputForce(batch.N(i), units), OutputForceUnit(units), "Заданная продольная сила."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "MxUser", OutputMx(batch.UserMx(i), units), OutputMomentUnit(units), "Пользовательский Mx."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "MyUser", OutputMy(batch.UserMy(i), units), OutputMomentUnit(units), "Пользовательский My."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "MxInternal", OutputMx(batch.Mx(i), units), OutputMomentUnit(units), "Mx после переноса к координатам элементов."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "MyInternal", OutputMy(batch.My(i), units), OutputMomentUnit(units), "My после переноса к координатам элементов."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "OverallStatus", batch.OverallStatus(i), "-", "Итоговый пользовательский статус сочетания."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "ProfileId", batch.ProfileId(i), "-", "Расчетный профиль, назначенный сочетанию в rngLoadCombinations."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "DirectStateStatus", batch.DirectStateStatus(i), "-", "Статус прямого решения НДС фактического сочетания."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "CapacityStatus", batch.CapacityStatus(i), "-", "Статус расчета несущей способности."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "CapacityLoadPath", batch.CapacityLoadPath(i), "-", "Какие компоненты нагрузки масштабировались при поиске несущей способности."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "CapacitySolutionMethod", batch.CapacitySolutionMethod(i), "-", "Фактический численный метод, которым найдена предельная несущая способность."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "LambdaCapacity", batch.LambdaCapacity(i), "-", "Предельный множитель для выбранной lambda-траектории."
        If batch.LambdaCapacity(i) > 0# Then
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "Nultimate", OutputForce(batch.NUltimate(i), units), OutputForceUnit(units), "Полная N при lambdaUltimate в точке приложения нагрузки."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "MxUltimate", OutputMx(batch.MxUltimate(i), units), OutputMomentUnit(units), "Полный Mx при lambdaUltimate в точке приложения нагрузки."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "MyUltimate", OutputMy(batch.MyUltimate(i), units), OutputMomentUnit(units), "Полный My при lambdaUltimate в точке приложения нагрузки."
        Else
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "Nultimate", vbNullString, OutputForceUnit(units), "Заполняется после успешного поиска несущей способности."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "MxUltimate", vbNullString, OutputMomentUnit(units), "Заполняется после успешного поиска несущей способности."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "MyUltimate", vbNullString, OutputMomentUnit(units), "Заполняется после успешного поиска несущей способности."
        End If
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "CrackStatus", batch.CrackStatus(i), "-", "Статус расчета ширины раскрытия трещин."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "LongitudinalCrackStatus", batch.LongitudinalCrackStatus(i), "-", "Статус SLS(II)-проверки продольных трещин по сжатому бетону; выполняется как часть Calculation.Crack.Width."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "LongitudinalCrack.MaxConcreteCompressionStress", OutputStress(batch.MaxConcreteCompressionStress(i), units), OutputStressUnit(units), "Максимальное по модулю сжимающее напряжение бетона в прямом НДС."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "LongitudinalCrack.RbMc2", OutputStress(batch.LongitudinalCrackRbMc2(i), units), OutputStressUnit(units), "Пользовательское допустимое Rb,mc2 для проверки продольных трещин."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "LongitudinalCrack.Utilization", batch.LongitudinalCrackUtilization(i), "-", "Использование по продольным трещинам: sigma_c,max / Rb,mc2."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "ExtensionUsed", BooleanText(batch.ExtensionUsed(i)), "-", "True, если прямой StateSolution вышел за физические пределы диаграммы в numerical extension."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "DirectState.DiagramExtensionEnabled", BooleanText(batch.DirectStateDiagramExtensionEnabled(i)), "-", "Сохраненное значение настройки Solver.DirectState.DiagramExtension на момент расчета."
        PutNamedStateProperties data, rowIndex, runID, batch, i, units
    Next i

    target.Resize(rowCount, SECTION_PROPERTY_COLUMNS).Value2 = data
    FormatResultsBlock target, rowCount, SECTION_PROPERTY_COLUMNS, "Свойства всего сечения"
End Sub

' Возвращает английскую запись логического флага для Results snapshot.
' Так файл не зависит от локали Excel, где Boolean может отображаться как
' ИСТИНА/ЛОЖЬ, TRUE/FALSE или в другом переводе интерфейса.
Private Function BooleanText(ByVal value As Boolean) As String
    If value Then
        BooleanText = "True"
    Else
        BooleanText = "False"
    End If
End Function

' ------------------------------
' Диагностический вывод фактических диаграмм
' ------------------------------
' Пишет уникальный каталог CMaterialDiagram, которые solver/provider
' использовали для конечных named-state текущего snapshot. Диаграммы зависят
' от профиля, StateType, MaterialModelSpec и режима extension, но не от номера
' LC, поэтому одинаковые сочетания не дублируются.
Private Sub WriteMaterialDiagrams(ByVal workbook As Object, ByVal materialProvider As CMaterialModelProvider, _
        ByVal batch As CBatchSectionCalculator, ByVal runID As Variant, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMMaterialDiagrams").RefersToRange

    Dim rowCapacity As Long
    rowCapacity = MaterialDiagramRowCount(materialProvider, batch)

    Dim data() As Variant
    ReDim data(1 To rowCapacity, 1 To MATERIAL_DIAGRAM_COLUMNS)
    PutMaterialDiagramHeaders data, units

    Dim rowIndex As Long
    rowIndex = 2
    Dim written As Object
    Set written = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To batch.Count
        PutCombinationMaterialDiagrams data, rowIndex, runID, materialProvider, batch, i, units, written
    Next i

    Dim rowCount As Long
    rowCount = rowIndex - 1
    If rowCount < 1 Then rowCount = 1

    target.Resize(rowCount, MATERIAL_DIAGRAM_COLUMNS).Value2 = data
    FormatResultsBlock target, rowCount, MATERIAL_DIAGRAM_COLUMNS, "Фактические диаграммы материалов"
End Sub

' Собирает материальные диаграммы одного сочетания по фактически сохраненным
' named states, но пишет только новые DiagramId. Это важно для профилей: один
' и тот же StateType у разных профилей может иметь разную MaterialModelSpec.
Private Sub PutCombinationMaterialDiagrams(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal runID As Variant, ByVal materialProvider As CMaterialModelProvider, _
        ByVal batch As CBatchSectionCalculator, ByVal combinationIndex As Long, ByVal units As CUnitSystem, _
        ByVal written As Object)
    Dim stateIndex As Long
    For stateIndex = 1 To batch.NamedStateCount(combinationIndex)
        Dim stateResult As CSectionStateResult
        Set stateResult = batch.NamedState(combinationIndex, stateIndex)
        If Not stateResult Is Nothing Then
            AppendMaterialDiagramPoints data, rowIndex, runID, batch.ProfileId(combinationIndex), _
                stateResult, "Concrete", _
                ConcreteDiagramForState(materialProvider, stateResult, batch.DirectStateDiagramExtensionEnabled(combinationIndex)), _
                batch.DirectStateDiagramExtensionEnabled(combinationIndex), units, written
            AppendMaterialDiagramPoints data, rowIndex, runID, batch.ProfileId(combinationIndex), _
                stateResult, "Rebar", _
                SteelDiagramForState(materialProvider, stateResult, batch.DirectStateDiagramExtensionEnabled(combinationIndex)), _
                batch.DirectStateDiagramExtensionEnabled(combinationIndex), units, written
        End If
    Next stateIndex
End Sub

' Возвращает бетонную диаграмму именно в том виде, в котором ее мог получить
' solver для named-state. Прямые StateSolution при включенном numerical
' extension используют state-диаграммы, а capacity/CrackInitiation-состояния
' остаются строго физическими.
Private Function ConcreteDiagramForState(ByVal materialProvider As CMaterialModelProvider, _
        ByVal stateResult As CSectionStateResult, ByVal directExtensionEnabled As Boolean) As CMaterialDiagram
    If IsDirectStateResult(stateResult) And directExtensionEnabled Then
        Set ConcreteDiagramForState = materialProvider.ConcreteStateMaterialFromSpec(stateResult.MaterialSpec)
    Else
        Set ConcreteDiagramForState = materialProvider.ConcreteMaterialFromSpec(stateResult.MaterialSpec)
    End If
End Function

' Возвращает диаграмму арматуры для диагностического snapshot-а по тем же
' правилам, по которым расчетный orchestrator отдавал материал solver-у.
Private Function SteelDiagramForState(ByVal materialProvider As CMaterialModelProvider, _
        ByVal stateResult As CSectionStateResult, ByVal directExtensionEnabled As Boolean) As CMaterialDiagram
    If IsDirectStateResult(stateResult) And directExtensionEnabled Then
        Set SteelDiagramForState = materialProvider.SteelStateMaterialFromSpec(stateResult.MaterialSpec)
    Else
        Set SteelDiagramForState = materialProvider.SteelMaterialFromSpec(stateResult.MaterialSpec)
    End If
End Function

' Прямые состояния LC строятся через CStateSolutionRunner; только они могут
' получать numerical extension. Остальные named-state рассчитываются по
' физическим диаграммам соответствующей MaterialModelSpec.
Private Function IsDirectStateResult(ByVal stateResult As CSectionStateResult) As Boolean
    If stateResult Is Nothing Then Exit Function
    IsDirectStateResult = (stateResult.StateType = sstStrengthState Or stateResult.StateType = sstCrackedState)
End Function

' Формирует стабильный идентификатор уникальной диаграммы текущего snapshot.
' LoadCase намеренно не входит в ключ: если разные сочетания используют один
' профиль и одно named-state, материал для solver-а у них одинаковый.
Private Function MaterialDiagramId(ByVal profileId As String, ByVal stateResult As CSectionStateResult, _
        ByVal materialKind As String, ByVal directExtensionEnabled As Boolean) As String
    MaterialDiagramId = profileId & "|" & stateResult.StateTypeText & "|" & _
        stateResult.MaterialModelRoleText & "|" & materialKind & "|" & _
        MaterialDiagramModeText(stateResult, directExtensionEnabled) & "|" & _
        stateResult.MaterialSpec.SpecKey
End Function

' Показывает, была ли в точках диаграммы записана физическая модель или
' техническое продолжение для прямого StateSolution.
Private Function MaterialDiagramModeText(ByVal stateResult As CSectionStateResult, _
        ByVal directExtensionEnabled As Boolean) As String
    If IsDirectStateResult(stateResult) And directExtensionEnabled Then
        MaterialDiagramModeText = "StateExtension"
    Else
        MaterialDiagramModeText = "Physical"
    End If
End Function

' Добавляет все точки одной диаграммы. Напряжения пишутся в OUTPUT stress unit,
' деформации остаются безразмерными, как их получает solver.
Private Sub AppendMaterialDiagramPoints(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal runID As Variant, ByVal profileId As String, ByVal stateResult As CSectionStateResult, _
        ByVal materialKind As String, ByVal diagram As CMaterialDiagram, _
        ByVal directExtensionEnabled As Boolean, ByVal units As CUnitSystem, ByVal written As Object)
    If diagram Is Nothing Then Exit Sub

    Dim diagramId As String
    diagramId = MaterialDiagramId(profileId, stateResult, materialKind, directExtensionEnabled)
    If written.Exists(diagramId) Then Exit Sub
    written.Add diagramId, True

    Dim pointIndex As Long
    For pointIndex = 1 To diagram.PointCount
        data(rowIndex, 1) = runID
        data(rowIndex, 2) = diagramId
        data(rowIndex, 3) = profileId
        data(rowIndex, 4) = stateResult.StateTypeText
        data(rowIndex, 5) = stateResult.MaterialModelRoleText
        data(rowIndex, 6) = stateResult.MaterialSpec.SpecKey
        data(rowIndex, 7) = MaterialDiagramModeText(stateResult, directExtensionEnabled)
        data(rowIndex, 8) = materialKind
        data(rowIndex, 9) = pointIndex
        data(rowIndex, 10) = diagram.PointStrain(pointIndex)
        data(rowIndex, 11) = OutputStress(diagram.PointStress(pointIndex), units)
        rowIndex = rowIndex + 1
    Next pointIndex
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

' Записывает справочные свойства бетонной части сечения.
' Эти строки нужны, чтобы на Results было видно, от какой точки реально
' задается Load.ReferenceOffset и какие геометрические характеристики имеет
' бетонное сечение без учета продольной арматуры.
Private Sub PutConcreteSectionProperties(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal runID As Variant, ByVal section As CSectionModel, _
        ByVal props As CSectionPropertiesCalculator, ByVal units As CUnitSystem)
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Area", OutputArea(props.Area, units), OutputAreaUnit(units), "Площадь бетонной части сечения по расчетной сетке."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.CentroidX", OutputLength(props.CentroidX, units), OutputLengthUnit(units), "X центра тяжести бетонного сечения; от этой точки задается Load.ReferenceOffsetX."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.CentroidY", OutputLength(props.CentroidY, units), OutputLengthUnit(units), "Y центра тяжести бетонного сечения; от этой точки задается Load.ReferenceOffsetY."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Ix", OutputFourthPowerLength(props.Ix, units), OutputLengthUnit(units) & "4", "Момент инерции бетонного сечения относительно исходной оси X."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Iy", OutputFourthPowerLength(props.Iy, units), OutputLengthUnit(units) & "4", "Момент инерции бетонного сечения относительно исходной оси Y."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Ixy", OutputFourthPowerLength(props.Ixy, units), OutputLengthUnit(units) & "4", "Центробежный момент бетонного сечения относительно исходных осей."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Ixc", OutputFourthPowerLength(props.Ixc, units), OutputLengthUnit(units) & "4", "Центральный момент инерции бетонного сечения относительно оси X."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Iyc", OutputFourthPowerLength(props.Iyc, units), OutputLengthUnit(units) & "4", "Центральный момент инерции бетонного сечения относительно оси Y."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Ixyc", OutputFourthPowerLength(props.Ixyc, units), OutputLengthUnit(units) & "4", "Центральный центробежный момент бетонного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.PrincipalAngle", props.PrincipalAngleRad, "rad", "Угол от +X к главной оси 1 бетонного сечения; положительный против часовой стрелки."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.PrincipalI1", OutputFourthPowerLength(props.PrincipalI1, units), OutputLengthUnit(units) & "4", "Больший главный момент инерции бетонного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.PrincipalI2", OutputFourthPowerLength(props.PrincipalI2, units), OutputLengthUnit(units) & "4", "Меньший главный момент инерции бетонного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.RadiusX", OutputLength(props.RadiusX, units), OutputLengthUnit(units), "Радиус инерции бетонного сечения относительно центральной оси X."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.RadiusY", OutputLength(props.RadiusY, units), OutputLengthUnit(units), "Радиус инерции бетонного сечения относительно центральной оси Y."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.PrincipalRadius1", OutputLength(props.PrincipalRadius1, units), OutputLengthUnit(units), "Радиус инерции бетонного сечения по главной оси I1."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.PrincipalRadius2", OutputLength(props.PrincipalRadius2, units), OutputLengthUnit(units), "Радиус инерции бетонного сечения по главной оси I2."
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
    PutProperty data, rowIndex, runID, "ALL", "Transformed.CentroidX", OutputLength(props.CentroidX, units), OutputLengthUnit(units), "X центра тяжести приведенного сечения бетон + арматура."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.CentroidY", OutputLength(props.CentroidY, units), OutputLengthUnit(units), "Y центра тяжести приведенного сечения бетон + арматура."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Ix", OutputFourthPowerLength(props.Ix, units), OutputLengthUnit(units) & "4", "Момент инерции приведенного сечения относительно исходной оси X."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Iy", OutputFourthPowerLength(props.Iy, units), OutputLengthUnit(units) & "4", "Момент инерции приведенного сечения относительно исходной оси Y."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Ixy", OutputFourthPowerLength(props.Ixy, units), OutputLengthUnit(units) & "4", "Центробежный момент приведенного сечения относительно исходных осей."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Ixc", OutputFourthPowerLength(props.Ixc, units), OutputLengthUnit(units) & "4", "Центральный момент инерции приведенного сечения относительно оси X."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Iyc", OutputFourthPowerLength(props.Iyc, units), OutputLengthUnit(units) & "4", "Центральный момент инерции приведенного сечения относительно оси Y."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Ixyc", OutputFourthPowerLength(props.Ixyc, units), OutputLengthUnit(units) & "4", "Центральный центробежный момент приведенного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.PrincipalAngle", props.PrincipalAngleRad, "rad", "Угол от +X к главной оси 1 приведенного сечения; положительный против часовой стрелки."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.PrincipalI1", OutputFourthPowerLength(props.PrincipalI1, units), OutputLengthUnit(units) & "4", "Главный момент инерции I1 приведенного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.PrincipalI2", OutputFourthPowerLength(props.PrincipalI2, units), OutputLengthUnit(units) & "4", "Главный момент инерции I2 приведенного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.RadiusX", OutputLength(props.RadiusX, units), OutputLengthUnit(units), "Радиус инерции приведенного сечения относительно центральной оси X."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.RadiusY", OutputLength(props.RadiusY, units), OutputLengthUnit(units), "Радиус инерции приведенного сечения относительно центральной оси Y."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.PrincipalRadius1", OutputLength(props.PrincipalRadius1, units), OutputLengthUnit(units), "Радиус инерции приведенного сечения по главной оси I1."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.PrincipalRadius2", OutputLength(props.PrincipalRadius2, units), OutputLengthUnit(units), "Радиус инерции приведенного сечения по главной оси I2."
    PutPrincipalCoreDistanceProperties data, rowIndex, runID, "Transformed", _
        "приведенного сечения", section, props, units
End Sub

' Записывает ядровые расстояния по двум главным направлениям. Для
' несимметричных сечений расстояния в сторону +оси и -оси могут отличаться,
' поэтому сохраняем все четыре величины. Контур c берется по бетонной оболочке:
' арматура влияет на A/I приведенного сечения, но не расширяет внешний бетон.
Private Sub PutPrincipalCoreDistanceProperties(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal runID As Variant, ByVal prefix As String, ByVal description As String, _
        ByVal section As CSectionModel, ByVal props As CSectionPropertiesCalculator, _
        ByVal units As CUnitSystem)
    PutProperty data, rowIndex, runID, "ALL", prefix & ".CoreDistance1Plus", _
        OutputLength(props.PrincipalPlaneCoreDistance(section, 1, True, False), units), _
        OutputLengthUnit(units), "Ядровое расстояние r " & description & " для изгиба вокруг главной оси 1, в положительную сторону эксцентриситета."
    PutProperty data, rowIndex, runID, "ALL", prefix & ".CoreDistance1Minus", _
        OutputLength(props.PrincipalPlaneCoreDistance(section, 1, False, False), units), _
        OutputLengthUnit(units), "Ядровое расстояние r " & description & " для изгиба вокруг главной оси 1, в отрицательную сторону эксцентриситета."
    PutProperty data, rowIndex, runID, "ALL", prefix & ".CoreDistance2Plus", _
        OutputLength(props.PrincipalPlaneCoreDistance(section, 2, True, False), units), _
        OutputLengthUnit(units), "Ядровое расстояние r " & description & " для изгиба вокруг главной оси 2, в положительную сторону эксцентриситета."
    PutProperty data, rowIndex, runID, "ALL", prefix & ".CoreDistance2Minus", _
        OutputLength(props.PrincipalPlaneCoreDistance(section, 2, False, False), units), _
        OutputLengthUnit(units), "Ядровое расстояние r " & description & " для изгиба вокруг главной оси 2, в отрицательную сторону эксцентриситета."
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
        "-", "Гибкость l0/i по СП 35 для главной плоскости 1 приведенного сечения."
    PutProperty data, rowIndex, runID, "ALL", "SP35.Slenderness2", _
        SlendernessValue(batch.StabilityElementLength, batch.StabilityMu2, transformedProps.PrincipalRadius2), _
        "-", "Гибкость l0/i по СП 35 для главной плоскости 2 приведенного сечения."
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

' Записывает компактную metadata по каждому конечному named-state LC.
' Поэлементная таблица хранит Stress/Strain, а здесь лежит то, что не нужно
' повторять для каждого волокна: плоскость деформаций и фактическая спецификация
' материальной модели.
Private Sub PutNamedStateProperties(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal runID As Variant, ByVal batch As CBatchSectionCalculator, ByVal combinationIndex As Long, _
        ByVal units As CUnitSystem)
    Dim stateIndex As Long
    For stateIndex = 1 To batch.NamedStateCount(combinationIndex)
        Dim stateResult As CSectionStateResult
        Set stateResult = batch.NamedState(combinationIndex, stateIndex)
        If Not stateResult Is Nothing Then
            Dim prefix As String
            prefix = "State." & stateResult.StateTypeText & "."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "MaterialModelRole", _
                stateResult.MaterialModelRoleText, "-", "Смысловая роль материальной модели для именованного расчетного состояния."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "MaterialModelSpec", _
                stateResult.MaterialSpec.SpecKey, "-", "Фактическая спецификация ValueSet/ConcreteDiagram/ConcreteTension/SteelDiagram."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "ConcreteDiagramId", _
                MaterialDiagramId(batch.ProfileId(combinationIndex), stateResult, "Concrete", batch.DirectStateDiagramExtensionEnabled(combinationIndex)), _
                "-", "Ссылка на бетонную диаграмму в rngNDMMaterialDiagrams."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "RebarDiagramId", _
                MaterialDiagramId(batch.ProfileId(combinationIndex), stateResult, "Rebar", batch.DirectStateDiagramExtensionEnabled(combinationIndex)), _
                "-", "Ссылка на диаграмму арматуры в rngNDMMaterialDiagrams."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "Epsilon0", _
                stateResult.Epsilon0, "-", "Деформация в начале координат для именованного расчетного состояния."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "KappaX", _
                OutputCurvature(stateResult.KappaX, units), OutputCurvatureUnit(units), "Кривизна KappaX для именованного расчетного состояния."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "KappaY", _
                OutputCurvature(stateResult.KappaY, units), OutputCurvatureUnit(units), "Кривизна KappaY для именованного расчетного состояния."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "NeutralLine.Angle", _
                NeutralLineAngleValue(stateResult.KappaX, stateResult.KappaY), "rad", _
                "Угол нейтральной линии от глобальной оси X против часовой стрелки."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "NeutralLine.SignedDistance", _
                OutputLengthVariant(NeutralLineSignedDistanceValue(stateResult.Epsilon0, stateResult.KappaX, stateResult.KappaY), units), _
                OutputLengthUnit(units), "Знаковое расстояние от начала координат до нейтральной линии по нормали (KappaY, KappaX)."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "ExtensionUsed", _
                BooleanText(stateResult.ExtensionUsed), "-", "True, если это расчетное состояние использовало numerical extension."
            PutProperty data, rowIndex, runID, batch.CombinationID(combinationIndex), prefix & "Status", _
                stateResult.Status, "-", "Короткий статус состояния, сохраненный orchestrator-ом."
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

' Локальный atan2 нужен writer-у для справочного угла нейтральной линии.
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

' Считает точную высоту rngNDMSectionProperties до выделения массива.
' Это важно для полного snapshot-а: при динамическом числе LC и нескольких named-state
' на LC фиксированного запаса может не хватить, и VBA падает с Subscript out of range.
Private Function SectionPropertyRowCount(ByVal batch As CBatchSectionCalculator) As Long
    Dim rowCount As Long
    rowCount = 1 + COMMON_SECTION_PROPERTY_ROWS + batch.Count * LC_SECTION_PROPERTY_ROWS

    Dim combinationIndex As Long
    For combinationIndex = 1 To batch.Count
        rowCount = rowCount + batch.NamedStateCount(combinationIndex) * STATE_SECTION_PROPERTY_ROWS
    Next combinationIndex

    SectionPropertyRowCount = rowCount
End Function

' Считает точную высоту диагностической таблицы фактических диаграмм.
' В расчетах с профилями одно сочетание может иметь несколько material spec,
' поэтому размер зависит от фактических states и количества точек диаграмм.
Private Function MaterialDiagramRowCount(ByVal materialProvider As CMaterialModelProvider, _
        ByVal batch As CBatchSectionCalculator) As Long
    Dim rowCount As Long
    rowCount = 1

    Dim combinationIndex As Long
    For combinationIndex = 1 To batch.Count
        Dim stateIndex As Long
        For stateIndex = 1 To batch.NamedStateCount(combinationIndex)
            Dim stateResult As CSectionStateResult
            Set stateResult = batch.NamedState(combinationIndex, stateIndex)
            If Not stateResult Is Nothing Then
                rowCount = rowCount + ConcreteDiagramForState(materialProvider, stateResult, _
                    batch.DirectStateDiagramExtensionEnabled(combinationIndex)).PointCount
                rowCount = rowCount + SteelDiagramForState(materialProvider, stateResult, _
                    batch.DirectStateDiagramExtensionEnabled(combinationIndex)).PointCount
            End If
        Next stateIndex
    Next combinationIndex

    MaterialDiagramRowCount = rowCount
End Function

' Записывает минимальные свойства imported-preview: Bounds, количество
' элементов и единицы, достаточные для проверки геометрии на схеме.
Private Sub WriteSectionPropertiesPreview(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal runID As Variant, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMSectionProperties").RefersToRange

    Dim data() As Variant
    ReDim data(1 To 35, 1 To SECTION_PROPERTY_COLUMNS)
    PutSectionPropertyHeaders data

    Dim rowIndex As Long
    rowIndex = 2

    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    GetConcreteBounds section, minX, maxX, minY, maxY

    PutProperty data, rowIndex, runID, "ALL", "Bounds.MinX", OutputLength(minX, units), OutputLengthUnit(units), "Минимальный X расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MaxX", OutputLength(maxX, units), OutputLengthUnit(units), "Максимальный X расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MinY", OutputLength(minY, units), OutputLengthUnit(units), "Минимальный Y расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MaxY", OutputLength(maxY, units), OutputLengthUnit(units), "Максимальный Y расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Geometry.Source", section.SourceType, "-", "Источник геометрии расчетной модели."

    Dim concreteProps As CSectionPropertiesCalculator
    Set concreteProps = New CSectionPropertiesCalculator
    concreteProps.CalculateConcrete section
    PutConcreteSectionProperties data, rowIndex, runID, section, concreteProps, units

    PutProperty data, rowIndex, runID, "ALL", "ConcreteCount", section.ConcreteCount, "шт", "Количество бетонных элементов."
    PutProperty data, rowIndex, runID, "ALL", "RebarCount", section.RebarCount, "шт", "Количество элементов арматуры."
    PutProperty data, rowIndex, runID, "ALL", "Output.LengthUnit", OutputLengthUnit(units), "-", "Единица вывода длины на момент импорта."
    PutProperty data, rowIndex, runID, "ALL", "Output.AreaUnit", OutputAreaUnit(units), "-", "Единица вывода площади на момент импорта."
    PutProperty data, rowIndex, runID, "ALL", "Output.StressUnit", OutputStressUnit(units), "-", "Единица вывода напряжения на момент импорта."
    PutProperty data, rowIndex, runID, "ALL", "Output.CurvatureUnit", OutputCurvatureUnit(units), "-", "Единица вывода кривизны на момент импорта."

    target.Resize(UBound(data, 1), SECTION_PROPERTY_COLUMNS).Value2 = data
    FormatResultsBlock target, UBound(data, 1), SECTION_PROPERTY_COLUMNS, "Свойства всего сечения"
End Sub

' ------------------------------
' Semantic-аннотации схемы
' ------------------------------
' Эти строки отвечают не за расчет, а за оформление последнего рассчитанного
' сечения: контуры, размеры и групповые подписи арматуры.
Private Sub WriteSectionAnnotations(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal runID As Variant, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMSectionAnnotations").RefersToRange

    Dim annotations As CSectionAnnotations
    Set annotations = section.Annotations

    Dim rowCount As Long
    rowCount = 1 + annotations.Count
    If rowCount < 1 Then rowCount = 1

    Dim data() As Variant
    ReDim data(1 To rowCount, 1 To ANNOTATION_COLUMNS)
    PutAnnotationHeaders data

    Dim rowIndex As Long
    rowIndex = 2

    Dim i As Long
    For i = 1 To annotations.Count
        PutAnnotationRow data, rowIndex, runID, annotations, i, units
    Next i

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

' Оставляет таблицу аннотаций с одной шапкой.
' У импортированной AutoCAD-сетки semantic-граней обычно нет, поэтому preview
' не должен наследовать старые размеры и подписи.
Private Sub WriteEmptySectionAnnotations(ByVal workbook As Object, ByVal runID As Variant)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMSectionAnnotations").RefersToRange

    Dim data() As Variant
    ReDim data(1 To 1, 1 To ANNOTATION_COLUMNS)
    PutAnnotationHeaders data

    target.Resize(1, ANNOTATION_COLUMNS).Value2 = data
    FormatResultsBlock target, 1, ANNOTATION_COLUMNS, "Аннотации схемы"
End Sub

' Оформляет один Results-диапазон как самостоятельную таблицу.
' Заголовок ставится прямо над якорем, а чтение Results ниже выполняется
' явным anchored-reader-ом, поэтому внешний вид больше не зависит от CurrentRegion.
Private Sub FormatResultsBlock(ByVal target As Object, ByVal rowCount As Long, ByVal columnCount As Long, ByVal titleText As String)
    Dim tableRange As Object
    Set tableRange = target.Resize(rowCount, columnCount)

    tableRange.HorizontalAlignment = -4108
    tableRange.VerticalAlignment = -4108
    tableRange.Font.Name = "Arial"
    tableRange.Font.Size = 9
    tableRange.Rows.RowHeight = 15
    tableRange.Borders.LineStyle = 1
    tableRange.Borders.Color = RGB(200, 200, 200)
    tableRange.Rows.Item(1).Interior.Color = RGB(217, 217, 217)
    tableRange.Rows.Item(1).Font.Bold = True
    tableRange.Rows.Item(1).HorizontalAlignment = -4131
    tableRange.Columns.Item(1).NumberFormat = "dd.mm.yy"

    Dim widthColumn As Long
    For widthColumn = 1 To columnCount
        target.Worksheet.Columns.Item(target.Column + widthColumn - 1).ColumnWidth = RESULTS_COLUMN_WIDTH
    Next widthColumn

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

' Применяет точность к общим числовым столбцам Results, которые должны читаться одинаково во всех snapshot-таблицах.
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

' Записывает LC-зависимые напряжения/деформации по всем бетонным и арматурным элементам
' для одного сохраненного состояния НДС. Используется только уже найденное состояние,
' solver здесь не запускается.
Private Sub PutCombinationElementResults(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        ByVal batch As CBatchSectionCalculator, ByVal combinationIndex As Long, _
        ByVal stateResult As CSectionStateResult, ByVal runID As Variant, ByVal units As CUnitSystem)
    Dim i As Long
    Dim strain As Double
    Dim stress As Double
    Dim concrete As CMaterialDiagram
    Dim steel As CMaterialDiagram

    If stateResult Is Nothing Then Exit Sub
    If Not stateResult.Converged Then Exit Sub

    Set concrete = ConcreteDiagramForState(materialProvider, stateResult, _
        batch.DirectStateDiagramExtensionEnabled(combinationIndex))
    Set steel = SteelDiagramForState(materialProvider, stateResult, _
        batch.DirectStateDiagramExtensionEnabled(combinationIndex))

    For i = 1 To section.ConcreteCount
        strain = stateResult.Epsilon0 + stateResult.KappaX * section.ConcreteY(i) + _
            stateResult.KappaY * section.ConcreteX(i)
        stress = concrete.GetStress(strain)
        PutElementRow data, rowIndex, runID, batch.CombinationID(combinationIndex), _
            batch.ProfileId(combinationIndex), stateResult.StateTypeText, section.ConcreteID(i), "Concrete", _
            strain, stress, PhysicalState("Concrete", strain, stress), units
        rowIndex = rowIndex + 1
    Next i

    For i = 1 To section.RebarCount
        strain = stateResult.Epsilon0 + stateResult.KappaX * section.RebarY(i) + _
            stateResult.KappaY * section.RebarX(i)
        stress = steel.GetStress(strain)
        PutElementRow data, rowIndex, runID, batch.CombinationID(combinationIndex), _
            batch.ProfileId(combinationIndex), stateResult.StateTypeText, section.RebarID(i), "Rebar", _
            strain, stress, PhysicalState("Rebar", strain, stress), units
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

' Классифицирует физическое состояние элемента для визуализации и диагностики Results.
Private Function PhysicalState(ByVal materialType As String, ByVal strain As Double, ByVal stress As Double) As String
    If Abs(stress) <= PHYSICAL_TOLERANCE Then
        If StrComp(materialType, "Concrete", vbTextCompare) = 0 And strain > PHYSICAL_TOLERANCE Then
            PhysicalState = "InactiveTensionConcrete"
        Else
            PhysicalState = "NearZero"
        End If
    ElseIf stress < 0# Then
        PhysicalState = "Compression"
    Else
        PhysicalState = "Tension"
    End If
End Function

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

' Формирует заголовки геометрического snapshot: реальные A/I и геометрическая оболочка идут рядом.
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

' Формирует заголовки semantic-аннотаций для схемы и AutoCAD export.
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

' Добавляет одну semantic-аннотацию в массив rngNDMSectionAnnotations с переводом координат в output units.
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

' Готовит видимый текст аннотации: для размеров пересобирает подпись уже в пользовательских единицах.
' На схему выводятся целые миллиметры, поэтому текст размера округляется, а
' числовое поле Value ниже остается точным для проверок и диагностического
' snapshot-а Results.
Private Function OutputAnnotationText(ByVal annotations As CSectionAnnotations, ByVal annotationIndex As Long, _
        ByVal units As CUnitSystem) As String
    Dim annotationType As String
    annotationType = annotations.AnnotationType(annotationIndex)
    If StrComp(annotationType, "DIMENSION", vbTextCompare) <> 0 Then
        OutputAnnotationText = annotations.Text(annotationIndex)
        Exit Function
    End If

    Dim valueText As String
    valueText = FormatRoundedAnnotationLength(OutputLength(annotations.Value(annotationIndex), units))
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

' Округляет видимую подпись размерной линии до ближайшего целого миллиметра.
' VBA.Round использует банковское округление, поэтому здесь явно применено
' инженерно ожидаемое правило: половина миллиметра округляется от нуля.
Private Function FormatRoundedAnnotationLength(ByVal value As Double) As String
    Dim roundedValue As Double
    If value >= 0# Then
        roundedValue = Int(value + 0.5)
    Else
        roundedValue = -Int(Abs(value) + 0.5)
    End If
    FormatRoundedAnnotationLength = FormatNumberInvariant(roundedValue)
End Function

' Переводит технические комментарии builders в понятные пользователю строки Results.
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
        OutputAnnotationComment = "Групповая подпись арматуры по смысловой грани."
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
        OutputAnnotationComment = "Semantic-аннотация схемы."
    End If
End Function

' Для AutoCAD-сетки у программы нет смысловых граней, но есть согласованная
' прямоугольная оболочка бетонных элементов. Ее строит CSectionModel:
' фактический прямоугольник, эквивалентный прямоугольник по A/I или
' последний эквивалентный квадрат по площади. Реальные A/I при этом остаются
' расчетными характеристиками элемента.
' Поэтому сохраняем только две приблизительные размерные аннотации: общую
' ширину и высоту по импортированной дискретной сетке.
Private Sub EnsureAutoCADBoundsAnnotations(ByVal section As CSectionModel)
    If section Is Nothing Then Exit Sub
    If StrComp(section.SourceType, "AutoCADImport", vbTextCompare) <> 0 Then Exit Sub
    If section.AnnotationCount > 0 Then Exit Sub
    If section.ConcreteCount <= 0 Then Exit Sub

    Dim minX As Double
    Dim maxX As Double
    Dim minY As Double
    Dim maxY As Double
    GetConcreteBounds section, minX, maxX, minY, maxY

    section.Annotations.AddDimension "DIM_AUTO_BOUNDS_B", minX, minY, maxX, minY, 0#, -1#, _
        ApproxSign(), maxX - minX, "Приблизительная общая ширина по расчетным границам импортированного бетона."
    section.Annotations.AddDimension "DIM_AUTO_BOUNDS_H", minX, minY, minX, maxY, -1#, 0#, _
        ApproxSign(), maxY - minY, "Приблизительная общая высота по расчетным границам импортированного бетона."
End Sub

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

' Получает внешние границы бетона через общую проекционную геометрию CSectionPropertiesCalculator.
Private Sub GetConcreteBounds(ByVal section As CSectionModel, ByRef minX As Double, ByRef maxX As Double, _
        ByRef minY As Double, ByRef maxY As Double)
    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateProjection section, 1#, 0#, False, minX, maxX
    props.CalculateProjection section, 0#, 1#, False, minY, maxY
End Sub

' Ниже собраны адаптеры output units для полного снимка Results.
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

' Локальный максимум нужен для компактного расчета границ и значений snapshot.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

' Локальный минимум нужен для компактного расчета границ и значений snapshot.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function
