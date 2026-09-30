VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStabilitySummaryWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Детальная таблица продольного изгиба и устойчивости на листе Results
' ==========================================================================
' Класс пишет отдельный блок от якоря rngStabilitySummaryAnchor. Расчет здесь
' не выполняется: writer только берет уже сохраненные в batch величины
' устойчивости, переводит их в пользовательские единицы и оформляет таблицу.
'
' Верхний rngBatchSummary остается компактной общей сводкой, а этот блок дает
' подробную расшифровку по главным плоскостям и выбранному нормативу СП 35/СП 63.

Private Const HEADER_ROWS As Long = 5 ' Количество строк шапки над якорем.
Private Const TOTAL_COLUMNS As Long = 84 ' Ширина таблицы с пустыми разделителями между блоками.

Private Const COL_SUMMARY_FIRST As Long = 1
Private Const COL_SUMMARY_LAST As Long = 14
Private Const COL_COMMON_FIRST As Long = 16
Private Const COL_COMMON_LAST As Long = 27
Private Const COL_SP35_FIRST As Long = 29
Private Const COL_SP35_LAST As Long = 59
Private Const COL_SP63_FIRST As Long = 61
Private Const COL_SP63_LAST As Long = 84

Private Const SP35_PLANE1_FIRST As Long = 30
Private Const SP35_PLANE2_FIRST As Long = 45
Private Const SP63_PLANE1_FIRST As Long = 61
Private Const SP63_PLANE2_FIRST As Long = 73

Private Const HEADER_COLOR As Long = &HD9D9D9 ' Серый фон шапки, как в остальных таблицах Results.
Private Const SMALL_VALUE As Double = 0.000000001 ' Малый порог для деления и пустых расчетных полей.
Private Const RESULTS_COLUMN_WIDTH As Double = 10# ' В Results все столбцы оставляем одинаковой ширины.

' Возвращает количество строк шапки, которые writer строит выше rngStabilitySummaryAnchor.
Public Function HeaderRowsAboveAnchor() As Long
    HeaderRowsAboveAnchor = HEADER_ROWS
End Function

' Возвращает полную высоту таблицы устойчивости с учетом шапки над якорем.
Public Function RequiredRows(ByVal combinationCount As Long) As Long
    If combinationCount < 1 Then combinationCount = 1
    RequiredRows = HEADER_ROWS + combinationCount
End Function

' Возвращает число строк данных, которое нужно оформить под доступные LC.
' Заполненные сочетания берутся из batch, а пустая сетка - из высоты rngLoadCombinations.
Public Function RequiredDataRowsForWorkbook(ByVal workbook As Object, ByVal combinationCount As Long) As Long
    RequiredDataRowsForWorkbook = combinationCount
    If RequiredDataRowsForWorkbook < 1 Then RequiredDataRowsForWorkbook = 1

    On Error GoTo SafeExit
    If workbook Is Nothing Then GoTo SafeExit
    Dim loads As Object
    Set loads = workbook.Names.Item("rngLoadCombinations").RefersToRange
    If loads.Rows.Count > 1 Then
        If loads.Rows.Count - 1 > RequiredDataRowsForWorkbook Then _
            RequiredDataRowsForWorkbook = loads.Rows.Count - 1
    End If
SafeExit:
    If RequiredDataRowsForWorkbook < 1 Then RequiredDataRowsForWorkbook = 1
End Function

' Возвращает полную высоту таблицы для фактической книги, включая пустые
' доступные строки сочетаний, которые тоже должны иметь рамки.
Public Function RequiredRowsForWorkbook(ByVal workbook As Object, ByVal combinationCount As Long) As Long
    RequiredRowsForWorkbook = RequiredRows(RequiredDataRowsForWorkbook(workbook, combinationCount))
End Function

' Возвращает полную ширину таблицы устойчивости, включая пустые столбцы-разделители.
Public Function RequiredColumns() As Long
    RequiredColumns = TOTAL_COLUMNS
End Function

' Записывает детальную таблицу устойчивости от якоря rngStabilitySummaryAnchor.
Public Sub WriteSummary(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator, _
        Optional ByVal units As CUnitSystem = Nothing)
    On Error GoTo Failed
    If workbook Is Nothing Then Err.Raise vbObjectError + 3980, "CStabilitySummaryWriter", "Книга Excel не передана."
    If batch Is Nothing Then Err.Raise vbObjectError + 3981, "CStabilitySummaryWriter", "Пакетный расчетчик не передан."

    Dim target As Object
    Set target = TargetRange(workbook, RequiredDataRowsForWorkbook(workbook, batch.Count))
    PrepareTarget target
    WriteBlockTitle target, "Расчет продольного изгиба и устойчивости"
    WriteHeaders target, units, batch
    WriteData target, batch, units
    FormatTable target, batch
    Exit Sub

Failed:
    Err.Raise Err.Number, "CStabilitySummaryWriter.WriteSummary", Err.Description
End Sub

' Полностью очищает старый блок устойчивости вместе с объединениями и форматами.
Public Sub ClearSummary(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3982, "CStabilitySummaryWriter", "Книга Excel не передана."
    PrepareTarget ClearTargetRange(workbook)
End Sub

' Возвращает полный диапазон таблицы: пять строк шапки над якорем и фактические строки LC.
Private Function TargetRange(ByVal workbook As Object, Optional ByVal dataRows As Long = 1) As Object
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange
    If anchor.Row <= HEADER_ROWS Then Err.Raise vbObjectError + 3983, "CStabilitySummaryWriter", _
        "rngStabilitySummaryAnchor должен иметь место для шапки над собой."
    Set TargetRange = anchor.Offset(-HEADER_ROWS, 0).Resize(RequiredRows(dataRows), TOTAL_COLUMNS)
End Function

' Возвращает область очистки старой таблицы до следующего нижнего якоря Results.
Private Function ClearTargetRange(ByVal workbook As Object) As Object
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange
    If anchor.Row <= HEADER_ROWS Then Err.Raise vbObjectError + 3983, "CStabilitySummaryWriter", _
        "rngStabilitySummaryAnchor должен иметь место для шапки над собой."

    Dim topRow As Long
    topRow = anchor.Row - HEADER_ROWS

    Dim rowsToClear As Long
    rowsToClear = RequiredRows(1)

    On Error GoTo UseDefaultHeight
    Dim nextAnchor As Object
    Set nextAnchor = workbook.Names.Item("rngNDMElementResults").RefersToRange
    If nextAnchor.Row > topRow Then rowsToClear = nextAnchor.Row - topRow

UseDefaultHeight:
    If topRow + rowsToClear - 1 > anchor.Worksheet.Rows.Count Then _
        rowsToClear = anchor.Worksheet.Rows.Count - topRow + 1
    Set ClearTargetRange = anchor.Worksheet.Cells(topRow, anchor.Column).Resize(rowsToClear, TOTAL_COLUMNS)
End Function

' Снимает старые объединения и чистит область, чтобы при переносе якоря не осталось старой верстки.
Private Sub PrepareTarget(ByVal target As Object)
    ClearBlockTitleRow target
    target.UnMerge
    target.Clear
End Sub

' Записывает отдельную строку-название над detailed-блоком устойчивости.
Private Sub WriteBlockTitle(ByVal target As Object, ByVal titleText As String)
    If target.Row <= 1 Then Exit Sub

    Dim titleRow As Long
    titleRow = target.Row - 1
    target.Worksheet.Rows.Item(titleRow).Interior.Color = RGB(189, 215, 238)
    target.Worksheet.Rows.Item(titleRow).RowHeight = 15
    With target.Worksheet.Cells(titleRow, target.Column)
        .Value2 = titleText
        .Font.Name = "Arial"
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(255, 0, 0)
        .HorizontalAlignment = -4131
        .VerticalAlignment = -4108
        .WrapText = False
    End With
End Sub

' Очищает строку-название вместе со старой заливкой, если якорь блока был перенесен.
Private Sub ClearBlockTitleRow(ByVal target As Object)
    If target.Row <= 1 Then Exit Sub
    target.Worksheet.Rows.Item(target.Row - 1).Clear
    target.Worksheet.Rows.Item(target.Row - 1).RowHeight = 15
End Sub

' Формирует многоуровневую шапку таблицы. Разделительные столбцы остаются пустыми.
Private Sub WriteHeaders(ByVal target As Object, ByVal units As CUnitSystem, _
        ByVal batch As CBatchSectionCalculator)
    MergeCells target, 1, COL_SUMMARY_FIRST, 1, BlockWidth(COL_SUMMARY_FIRST, COL_SUMMARY_LAST), _
        "Итог по расчету (с учетом " & ChrW$(&H3B7) & ")"
    MergeCells target, 2, 1, 4, 1, "Combination ID"
    MergeCells target, 2, 2, 4, 1, "ResultComment"
    MergeCells target, 2, 3, 4, 1, "общий статус"
    MergeCells target, 2, 4, 2, 5, "относительно главных центральных осей приведенного сечения"
    MergeCells target, 2, 9, 2, 3, "относительно центральных осей приведенного сечения"
    MergeCells target, 2, 12, 2, 3, "в точке приложения нагрузок"

    MergeCells target, 1, COL_COMMON_FIRST, 1, BlockWidth(COL_COMMON_FIRST, COL_COMMON_LAST), _
        "Общие данные для учета продольного изгиба и проверки устойчивости (по главным осям 1, 2)"
    MergeCells target, 2, 16, 1, 6, "Нагрузки до применения " & ChrW$(&H3B7) & " в ц.т. приведенного сечения"
    MergeCells target, 3, 16, 1, 3, "постоянные + временные"
    MergeCells target, 3, 19, 1, 3, "постоянные"
    MergeCells target, 2, 22, 2, 6, "характеристики сечения"

    MergeCells target, 1, COL_SP35_FIRST, 1, BlockWidth(COL_SP35_FIRST, COL_SP35_LAST), "Расчет по СП 35"
    MergeCells target, 2, 29, 3, 1, "общие для СП"
    WriteSP35PlaneHeaders target, SP35_PLANE1_FIRST, "Плоскость 1 (изгиб вокруг главной оси 1)"
    WriteSP35PlaneHeaders target, SP35_PLANE2_FIRST, "Плоскость 2 (изгиб вокруг главной оси 2)"

    MergeCells target, 1, COL_SP63_FIRST, 1, BlockWidth(COL_SP63_FIRST, COL_SP63_LAST), "Расчет по СП 63"
    WriteSP63PlaneHeaders target, SP63_PLANE1_FIRST, "Плоскость 1 (изгиб вокруг главной оси 1)"
    WriteSP63PlaneHeaders target, SP63_PLANE2_FIRST, "Плоскость 2 (изгиб вокруг главной оси 2)"

    PutHeaderNotes target
    PutHeaderRowLabels target, units, batch
End Sub

' Записывает верхние подписи для одной плоскости СП 35.
Private Sub WriteSP35PlaneHeaders(ByVal target As Object, ByVal firstColumn As Long, ByVal titleText As String)
    MergeCells target, 2, firstColumn, 1, 15, titleText
    MergeCells target, 3, firstColumn, 1, 3, "Общие для плоскости"
    MergeCells target, 3, firstColumn + 3, 1, 6, "при ec > r"
    MergeCells target, 3, firstColumn + 9, 1, 6, "при ec <= r"
End Sub

' Записывает верхние подписи для одной плоскости СП 63.
Private Sub WriteSP63PlaneHeaders(ByVal target As Object, ByVal firstColumn As Long, ByVal titleText As String)
    MergeCells target, 2, firstColumn, 2, 12, titleText
End Sub

' Заполняет строку примечаний над обозначениями. Это обычный текст без переноса:
' если он не помещается, Excel просто обрезает его по ширине ячейки.
Private Sub PutHeaderNotes(ByVal target As Object)
    Dim notes(1 To 1, 1 To TOTAL_COLUMNS) As Variant
    notes(1, 2) = "пояснение расчетного статуса блока устойчивости"
    notes(1, 4) = "Коэффициент продольного изгиба в плоскости 1: изгиб вокруг главной оси 1"
    notes(1, 5) = "Коэффициент продольного изгиба в плоскости 2: изгиб вокруг главной оси 2"
    notes(1, 6) = "Продольная сила после учета продольного изгиба"
    notes(1, 7) = "Момент после учета продольного изгиба в плоскости 1, вокруг главной оси 1"
    notes(1, 8) = "Момент после учета продольного изгиба в плоскости 2, вокруг главной оси 2"
    notes(1, 9) = "Продольная сила относительно центральных осей приведенного сечения"
    notes(1, 10) = "Момент Mx относительно центральных осей приведенного сечения"
    notes(1, 11) = "Момент My относительно центральных осей приведенного сечения"
    notes(1, 12) = "Продольная сила в пользовательской точке приложения нагрузок"
    notes(1, 13) = "Момент Mx в пользовательской точке приложения нагрузок"
    notes(1, 14) = "Момент My в пользовательской точке приложения нагрузок"

    notes(1, 16) = "Продольная сила до учета продольного изгиба"
    notes(1, 17) = "Момент в плоскости 1, вокруг главной оси 1, до учета продольного изгиба"
    notes(1, 18) = "Момент в плоскости 2, вокруг главной оси 2, до учета продольного изгиба"
    notes(1, 19) = "Продольная сила от постоянной или длительной части нагрузки"
    notes(1, 20) = "Момент в плоскости 1, вокруг главной оси 1, от постоянной или длительной части"
    notes(1, 21) = "Момент в плоскости 2, вокруг главной оси 2, от постоянной или длительной части"
    notes(1, 22) = "Размер сечения по нормали к плоскости 1, то есть вдоль главной оси 2"
    notes(1, 23) = "Размер сечения по нормали к плоскости 2, то есть вдоль главной оси 1"
    notes(1, 24) = "Радиус инерции для изгиба вокруг главной оси 1"
    notes(1, 25) = "Радиус инерции для изгиба вокруг главной оси 2"
    notes(1, 26) = "Расчетная длина для плоскости 1: изгиб вокруг главной оси 1"
    notes(1, 27) = "Расчетная длина для плоскости 2: изгиб вокруг главной оси 2"

    notes(1, 29) = "Коэффициент для ненапрягаемой арматуры"
    PutSP35PlaneNotes notes, SP35_PLANE1_FIRST
    PutSP35PlaneNotes notes, SP35_PLANE2_FIRST
    PutSP63PlaneNotes notes, SP63_PLANE1_FIRST
    PutSP63PlaneNotes notes, SP63_PLANE2_FIRST

    target.Offset(3, 0).Resize(1, TOTAL_COLUMNS).Value2 = notes
End Sub

' Заполняет примечания одной плоскости СП 35.
Private Sub PutSP35PlaneNotes(ByRef notes() As Variant, ByVal firstColumn As Long)
    notes(1, firstColumn) = "Случайный эксцентриситет для рассматриваемой плоскости"
    notes(1, firstColumn + 1) = "Расчетный эксцентриситет продольной силы"
    notes(1, firstColumn + 2) = "Ядровое расстояние для выбора ветви СП 35"
    notes(1, firstColumn + 3) = "Критическая продольная сила для ветви ec > r"
    notes(1, firstColumn + 4) = "Коэффициент формы эпюры для ветви ec > r"
    notes(1, firstColumn + 5) = "Коэффициент длительного действия нагрузки"
    notes(1, firstColumn + 6) = "Коэффициент увеличения момента"
    notes(1, firstColumn + 7) = "Отношение действующей сжимающей силы к критической"
    notes(1, firstColumn + 8) = "Статус проверки ветви ec > r"
    notes(1, firstColumn + 9) = "Табличный коэффициент продольного изгиба"
    notes(1, firstColumn + 10) = "Коэффициент влияния момента"
    notes(1, firstColumn + 11) = "Итоговый коэффициент ветви ec <= r"
    notes(1, firstColumn + 12) = "Предельная продольная сила по ветви ec <= r"
    notes(1, firstColumn + 13) = "Отношение предельной силы к действующей сжимающей"
    notes(1, firstColumn + 14) = "Статус проверки ветви ec <= r"
End Sub

' Заполняет примечания одной плоскости СП 63.
Private Sub PutSP63PlaneNotes(ByRef notes() As Variant, ByVal firstColumn As Long)
    notes(1, firstColumn) = "Гибкость элемента в рассматриваемой плоскости изгиба вокруг главной оси"
    notes(1, firstColumn + 1) = "Случайный эксцентриситет для рассматриваемой плоскости"
    notes(1, firstColumn + 2) = "Расчетный эксцентриситет продольной силы"
    notes(1, firstColumn + 3) = "Коэффициент длительного действия нагрузки"
    notes(1, firstColumn + 4) = "Коэффициент влияния эксцентриситета"
    notes(1, firstColumn + 5) = "Коэффициент жесткости бетонной части"
    notes(1, firstColumn + 6) = "Коэффициент жесткости арматуры"
    notes(1, firstColumn + 7) = "Условная жесткость элемента"
    notes(1, firstColumn + 8) = "Критическая продольная сила"
    notes(1, firstColumn + 9) = "Коэффициент увеличения момента"
    notes(1, firstColumn + 10) = "Отношение критической силы к действующей сжимающей"
    notes(1, firstColumn + 11) = "Статус проверки рассматриваемой плоскости"
End Sub

' Заполняет строку обозначений под многоуровневой шапкой.
Private Sub PutHeaderRowLabels(ByVal target As Object, ByVal units As CUnitSystem, _
        ByVal batch As CBatchSectionCalculator)
    Dim labels(1 To 1, 1 To TOTAL_COLUMNS) As Variant
    labels(1, 1) = "Combination ID"
    labels(1, 2) = "ResultComment"
    labels(1, 3) = "общий статус"
    labels(1, 4) = ChrW$(&H3B7) & "1"
    labels(1, 5) = ChrW$(&H3B7) & "2"
    labels(1, 6) = "N, " & OutputForceUnit(units)
    labels(1, 7) = "M1, " & OutputMomentUnit(units)
    labels(1, 8) = "M2, " & OutputMomentUnit(units)
    labels(1, 9) = "N, " & OutputForceUnit(units)
    labels(1, 10) = "Mx, " & OutputMomentUnit(units)
    labels(1, 11) = "My, " & OutputMomentUnit(units)
    labels(1, 12) = "N, " & OutputForceUnit(units)
    labels(1, 13) = "Mx, " & OutputMomentUnit(units)
    labels(1, 14) = "My, " & OutputMomentUnit(units)

    labels(1, 16) = "N, " & OutputForceUnit(units)
    labels(1, 17) = "M1, " & OutputMomentUnit(units)
    labels(1, 18) = "M2, " & OutputMomentUnit(units)
    labels(1, 19) = "N" & ChrW$(&H2097) & ", " & OutputForceUnit(units)
    labels(1, 20) = "M1" & ChrW$(&H2097) & ", " & OutputMomentUnit(units)
    labels(1, 21) = "M2" & ChrW$(&H2097) & ", " & OutputMomentUnit(units)
    labels(1, 22) = "h1, " & OutputLengthUnit(units)
    labels(1, 23) = "h2, " & OutputLengthUnit(units)
    labels(1, 24) = "i1, " & OutputLengthUnit(units)
    labels(1, 25) = "i2, " & OutputLengthUnit(units)
    labels(1, 26) = "l01, " & OutputLengthUnit(units)
    labels(1, 27) = "l02, " & OutputLengthUnit(units)

    labels(1, 29) = ChrW$(&H3C6) & "p"
    PutSP35PlaneLabels labels, SP35_PLANE1_FIRST, units, SP35EtaReserveLabel(batch)
    PutSP35PlaneLabels labels, SP35_PLANE2_FIRST, units, SP35EtaReserveLabel(batch)
    PutSP63PlaneLabels labels, SP63_PLANE1_FIRST, units
    PutSP63PlaneLabels labels, SP63_PLANE2_FIRST, units

    target.Offset(4, 0).Resize(1, TOTAL_COLUMNS).Value2 = labels
    ApplyHeaderSubscripts target
End Sub

' Заполняет обозначения одной плоскости СП 35.
' Формирует заголовок запаса для ветви СП 35 ec > r.
' В нем показывается не условное k, а фактический лимит N/Ncr,
' заданный пользователем в Stability.SP35.NOverNcrLimit.
Private Function SP35EtaReserveLabel(ByVal batch As CBatchSectionCalculator) As String
    Dim limitValue As Double
    limitValue = 0.7
    If Not batch Is Nothing Then
        If batch.StabilitySP35NOverNcrLimit > SMALL_VALUE Then _
            limitValue = batch.StabilitySP35NOverNcrLimit
    End If
    SP35EtaReserveLabel = FormatCoefficientInvariant(limitValue) & "*Ncr/N"
End Function

' Пишет коэффициент в заголовок с точкой в качестве десятичного разделителя.
Private Function FormatCoefficientInvariant(ByVal value As Double) As String
    FormatCoefficientInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

' Заполняет подписи одной плоскости СП 35, сохраняя раздельные колонки для
' ветвей ec > r и ec <= r.
Private Sub PutSP35PlaneLabels(ByRef labels() As Variant, ByVal firstColumn As Long, _
        ByVal units As CUnitSystem, ByVal etaReserveLabel As String)
    labels(1, firstColumn) = "ec,sl, " & OutputLengthUnit(units)
    labels(1, firstColumn + 1) = "ec, " & OutputLengthUnit(units)
    labels(1, firstColumn + 2) = "r, " & OutputLengthUnit(units)
    labels(1, firstColumn + 3) = "Ncr, " & OutputForceUnit(units)
    labels(1, firstColumn + 4) = ChrW$(&H3B4)
    labels(1, firstColumn + 5) = ChrW$(&H3C6) & "l (пост)"
    labels(1, firstColumn + 6) = ChrW$(&H3B7)
    labels(1, firstColumn + 7) = etaReserveLabel
    labels(1, firstColumn + 8) = "статус"
    labels(1, firstColumn + 9) = ChrW$(&H3C6) & "l (табл)"
    labels(1, firstColumn + 10) = ChrW$(&H3C6) & "m"
    labels(1, firstColumn + 11) = ChrW$(&H3C6)
    labels(1, firstColumn + 12) = "Nult, " & OutputForceUnit(units)
    labels(1, firstColumn + 13) = "Nult/N"
    labels(1, firstColumn + 14) = "статус"
End Sub

' Заполняет обозначения одной плоскости СП 63.
Private Sub PutSP63PlaneLabels(ByRef labels() As Variant, ByVal firstColumn As Long, ByVal units As CUnitSystem)
    labels(1, firstColumn) = "l0/i"
    labels(1, firstColumn + 1) = "ea, " & OutputLengthUnit(units)
    labels(1, firstColumn + 2) = "e0, " & OutputLengthUnit(units)
    labels(1, firstColumn + 3) = ChrW$(&H3C6) & "l (пост)"
    labels(1, firstColumn + 4) = ChrW$(&H3B4) & "e"
    labels(1, firstColumn + 5) = "kb"
    labels(1, firstColumn + 6) = "ks"
    labels(1, firstColumn + 7) = "D, " & OutputStiffnessUnit(units)
    labels(1, firstColumn + 8) = "Ncr, " & OutputForceUnit(units)
    labels(1, firstColumn + 9) = ChrW$(&H3B7)
    labels(1, firstColumn + 10) = "Ncr/N"
    labels(1, firstColumn + 11) = "статус"
End Sub

' Записывает строки LC одним массивом.
Private Sub WriteData(ByVal target As Object, ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem)
    Dim data() As Variant
    Dim dataRows As Long
    dataRows = target.Rows.Count - HEADER_ROWS
    If dataRows < 1 Then dataRows = 1
    ReDim data(1 To dataRows, 1 To TOTAL_COLUMNS)

    Dim i As Long
    For i = 1 To batch.Count
        PutCombination data, OutputDataOffset(batch, i, dataRows), batch, i, units
    Next i

    target.Offset(HEADER_ROWS, 0).Resize(dataRows, TOTAL_COLUMNS).Value2 = data
End Sub

' Заполняет одну строку LC: общий итог, общие исходные величины и выбранный блок СП.
' Возвращает позицию LC внутри data-части подробной таблицы.
' Основной поток читает offset из rngLoadCombinations; прямые тесты без offset продолжают выводиться компактно.
Private Function OutputDataOffset(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal dataRows As Long) As Long
    OutputDataOffset = batch.SourceDataOffset(index)
    If OutputDataOffset < 1 Then OutputDataOffset = index
    If OutputDataOffset > dataRows Then OutputDataOffset = index
    If OutputDataOffset > dataRows Then OutputDataOffset = dataRows
End Function

' Возвращает внешний статус устойчивости по ResultMeta. Старое строковое поле
' используется только как fallback для переходных тестовых результатов.
Private Function StabilityStatusText(ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal policy As CBatchStatusPolicy) As String
    Dim meta As CResultMeta
    Set meta = batch.StabilityMeta(index)
    If Not meta Is Nothing Then
        StabilityStatusText = policy.StatusFromMeta(meta)
    Else
        StabilityStatusText = policy.DisplayStatus(batch.StabilityStatus(index))
    End If
End Function

' Возвращает готовое пользовательское пояснение из ResultMeta блока.
' Writer не анализирует технические StopReason и не собирает комментарий заново.
Private Function MetaComment(ByVal meta As CResultMeta) As String
    If meta Is Nothing Then Exit Function
    MetaComment = meta.ResultComment
End Function

Private Sub PutCombination(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal units As CUnitSystem)
    Dim policy As CBatchStatusPolicy
    Set policy = StatusPolicy()

    Dim statusText As String
    statusText = StabilityStatusText(batch, index, policy)
    data(rowIndex, 1) = batch.CombinationID(index)
    data(rowIndex, 2) = MetaComment(batch.StabilityMeta(index))
    data(rowIndex, 3) = statusText
    PutDefaultStatusCells data, rowIndex, policy.NotApplicable
    If statusText = policy.NotApplicable Then Exit Sub
    If Not policy.IsFinished(statusText) Then Exit Sub

    data(rowIndex, 4) = OptionalValue(batch.StabilityEta1(index))
    data(rowIndex, 5) = OptionalValue(batch.StabilityEta2(index))
    data(rowIndex, 6) = OutputForce(batch.StabilityDesignN(index), units)
    data(rowIndex, 7) = OutputMomentMagnitude(batch.StabilityDesignMoment1(index), units)
    data(rowIndex, 8) = OutputMomentMagnitude(batch.StabilityDesignMoment2(index), units)
    data(rowIndex, 9) = OutputForce(batch.StabilityDesignN(index), units)
    data(rowIndex, 10) = OutputMx(batch.StabilityDesignMx(index), units)
    data(rowIndex, 11) = OutputMy(batch.StabilityDesignMy(index), units)
    data(rowIndex, 12) = OutputForce(batch.StabilityDesignN(index), units)
    data(rowIndex, 13) = OutputMx(batch.StabilityLoadPointMx(index), units)
    data(rowIndex, 14) = OutputMy(batch.StabilityLoadPointMy(index), units)

    data(rowIndex, 16) = OutputForce(batch.StabilityDesignN(index), units)
    data(rowIndex, 17) = OutputMomentMagnitude(batch.StabilityMoment1(index), units)
    data(rowIndex, 18) = OutputMomentMagnitude(batch.StabilityMoment2(index), units)
    data(rowIndex, 19) = OutputSignedCompressionForce(batch.StabilitySustainedN(index), units)
    data(rowIndex, 20) = OutputMomentMagnitude(batch.StabilitySustainedMoment1(index), units)
    data(rowIndex, 21) = OutputMomentMagnitude(batch.StabilitySustainedMoment2(index), units)
    data(rowIndex, 22) = OutputLength(batch.StabilityDepth1(index), units)
    data(rowIndex, 23) = OutputLength(batch.StabilityDepth2(index), units)
    data(rowIndex, 24) = OutputLength(batch.StabilityRadius1(index), units)
    data(rowIndex, 25) = OutputLength(batch.StabilityRadius2(index), units)
    data(rowIndex, 26) = OutputLength(batch.StabilityEffectiveLength1(index), units)
    data(rowIndex, 27) = OutputLength(batch.StabilityEffectiveLength2(index), units)

    If StrComp(batch.StabilityCode(index), "SP35", vbTextCompare) = 0 Then
        data(rowIndex, 29) = OptionalValue(MaxDouble(batch.StabilityPhiP1(index), batch.StabilityPhiP2(index)))
        PutSP35PlaneValues data, rowIndex, SP35_PLANE1_FIRST, batch, index, 1, units
        PutSP35PlaneValues data, rowIndex, SP35_PLANE2_FIRST, batch, index, 2, units
    ElseIf StrComp(batch.StabilityCode(index), "SP63", vbTextCompare) = 0 Then
        PutSP63PlaneValues data, rowIndex, SP63_PLANE1_FIRST, batch, index, 1, units
        PutSP63PlaneValues data, rowIndex, SP63_PLANE2_FIRST, batch, index, 2, units
    End If
End Sub

' Заполняет одну плоскость СП 35. В строке активна только ветвь ec > r или ec <= r.
Private Sub PutSP35PlaneValues(ByRef data() As Variant, ByVal rowIndex As Long, ByVal firstColumn As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long, _
        ByVal units As CUnitSystem)
    Dim branchName As String
    branchName = PlaneBranch(batch, index, planeIndex)

    data(rowIndex, firstColumn) = OutputLength(PlaneAccidentalEccentricity(batch, index, planeIndex), units)
    data(rowIndex, firstColumn + 1) = OutputLength(PlaneEccentricity(batch, index, planeIndex), units)
    data(rowIndex, firstColumn + 2) = OutputLength(PlaneCoreDistance(batch, index, planeIndex), units)
    data(rowIndex, firstColumn + 3) = OptionalCompressionForce(PlaneNcr(batch, index, planeIndex), units)

    If StrComp(branchName, "SP35-eta", vbTextCompare) = 0 Then
        data(rowIndex, firstColumn + 4) = OptionalValue(PlaneDelta(batch, index, planeIndex))
        data(rowIndex, firstColumn + 5) = OptionalValue(PlanePhiL(batch, index, planeIndex))
        data(rowIndex, firstColumn + 6) = OptionalValue(PlaneEta(batch, index, planeIndex))
        data(rowIndex, firstColumn + 7) = OptionalValue(PlaneReserve(batch, index, planeIndex))
        data(rowIndex, firstColumn + 8) = PlaneStatus(batch, index, planeIndex)
        data(rowIndex, firstColumn + 14) = StatusPolicy().NotApplicable
    ElseIf StrComp(branchName, "SP35-table", vbTextCompare) = 0 Then
        data(rowIndex, firstColumn + 8) = StatusPolicy().NotApplicable
        data(rowIndex, firstColumn + 9) = OptionalValue(PlanePhiLTable(batch, index, planeIndex))
        data(rowIndex, firstColumn + 10) = OptionalValue(PlanePhiM(batch, index, planeIndex))
        data(rowIndex, firstColumn + 11) = OptionalValue(PlanePhiValue(batch, index, planeIndex))
        data(rowIndex, firstColumn + 12) = OptionalCompressionForce(PlaneNultimate(batch, index, planeIndex), units)
        data(rowIndex, firstColumn + 13) = RatioLimitToCompression(PlaneNultimate(batch, index, planeIndex), batch.StabilityDesignN(index))
        data(rowIndex, firstColumn + 14) = PlaneStatus(batch, index, planeIndex)
    End If
End Sub

' Заполняет одну плоскость СП 63.
Private Sub PutSP63PlaneValues(ByRef data() As Variant, ByVal rowIndex As Long, ByVal firstColumn As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long, _
        ByVal units As CUnitSystem)
    data(rowIndex, firstColumn) = OptionalValue(PlaneSlenderness(batch, index, planeIndex))
    data(rowIndex, firstColumn + 1) = OutputLength(PlaneAccidentalEccentricity(batch, index, planeIndex), units)
    data(rowIndex, firstColumn + 2) = OutputLength(PlaneEccentricity(batch, index, planeIndex), units)
    data(rowIndex, firstColumn + 3) = OptionalValue(PlanePhiL(batch, index, planeIndex))
    data(rowIndex, firstColumn + 4) = OptionalValue(PlaneDelta(batch, index, planeIndex))
    data(rowIndex, firstColumn + 5) = OptionalValue(PlaneKb(batch, index, planeIndex))
    data(rowIndex, firstColumn + 6) = OptionalValue(PlaneKs(batch, index, planeIndex))
    data(rowIndex, firstColumn + 7) = OutputStiffness(PlaneStiffnessD(batch, index, planeIndex), units)
    data(rowIndex, firstColumn + 8) = OptionalCompressionForce(PlaneNcr(batch, index, planeIndex), units)
    data(rowIndex, firstColumn + 9) = OptionalValue(PlaneEta(batch, index, planeIndex))
    data(rowIndex, firstColumn + 10) = RatioLimitToCompression(PlaneNcr(batch, index, planeIndex), batch.StabilityDesignN(index))
    data(rowIndex, firstColumn + 11) = PlaneStatus(batch, index, planeIndex)
End Sub

' Применяет общее оформление, границы блоков и независимую красную подсветку.
Private Sub FormatTable(ByVal target As Object, ByVal batch As CBatchSectionCalculator)
    target.Font.Name = "Arial"
    target.Font.Size = 9
    target.HorizontalAlignment = -4108
    target.VerticalAlignment = -4108
    target.WrapText = False
    target.Rows.RowHeight = 15

    FormatBlock target, COL_SUMMARY_FIRST, COL_SUMMARY_LAST
    FormatBlock target, COL_COMMON_FIRST, COL_COMMON_LAST
    FormatBlock target, COL_SP35_FIRST, COL_SP35_LAST
    FormatBlock target, COL_SP63_FIRST, COL_SP63_LAST

    target.Offset(0, 0).Resize(HEADER_ROWS, TOTAL_COLUMNS).Font.Bold = True
    target.Offset(0, 0).Resize(HEADER_ROWS, TOTAL_COLUMNS).WrapText = True
    target.Offset(HEADER_ROWS - 1, 0).Resize(1, TOTAL_COLUMNS).WrapText = False
    FormatHeaderNotes target
    Dim dataRows As Long
    dataRows = target.Rows.Count - HEADER_ROWS
    If dataRows < 1 Then dataRows = 1
    target.Offset(HEADER_ROWS, 0).Resize(dataRows, TOTAL_COLUMNS).WrapText = False
    target.Offset(HEADER_ROWS, 0).Resize(dataRows, TOTAL_COLUMNS).NumberFormat = "0.000"
    target.Offset(HEADER_ROWS, 0).Resize(dataRows, 3).NumberFormat = "@"
    target.Offset(HEADER_ROWS, 1).Resize(dataRows, 1).HorizontalAlignment = -4131

    SetColumnWidths target

    FormatStatusCells target, batch
End Sub

' Оформляет строку примечаний как обычный текст, а не как дополнительный заголовок.
Private Sub FormatHeaderNotes(ByVal target As Object)
    With target.Offset(3, 0).Resize(1, TOTAL_COLUMNS)
        .Font.Bold = False
        .Font.Size = 8
        .HorizontalAlignment = -4131
        .VerticalAlignment = -4108
        .WrapText = False
    End With
    target.Worksheet.Rows.Item(target.Row + 3).RowHeight = 15
End Sub

' Задает рабочие ширины столбцов Results для детального блока устойчивости.
Private Sub SetColumnWidths(ByVal target As Object)
    Dim columnIndex As Long
    For columnIndex = 1 To TOTAL_COLUMNS
        target.Worksheet.Columns(target.Column + columnIndex - 1).ColumnWidth = RESULTS_COLUMN_WIDTH
    Next columnIndex
End Sub

' Оформляет один смысловой блок и не трогает пустой столбец-разделитель рядом.
Private Sub FormatBlock(ByVal target As Object, ByVal firstColumn As Long, ByVal lastColumn As Long)
    Dim block As Object
    Set block = target.Offset(0, firstColumn - 1).Resize(target.Rows.Count, lastColumn - firstColumn + 1)
    block.Borders.LineStyle = 1
    block.Borders.Color = RGB(0, 0, 0)
    block.Borders.Weight = 2
    block.HorizontalAlignment = -4108
    block.VerticalAlignment = -4108
    target.Offset(0, firstColumn - 1).Resize(HEADER_ROWS, lastColumn - firstColumn + 1).Interior.Color = HEADER_COLOR
End Sub

' Красит блоки строки LC по итоговому статусу расчета устойчивости.
Private Sub FormatStatusCells(ByVal target As Object, ByVal batch As CBatchSectionCalculator)
    If batch Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To batch.Count
        Dim rowOffset As Long
        rowOffset = OutputDataOffset(batch, i, target.Rows.Count - HEADER_ROWS)
        If HEADER_ROWS + rowOffset > target.Rows.Count Then Exit For

        ApplyStatusCell target, HEADER_ROWS + rowOffset, 3
        ApplyStatusCell target, HEADER_ROWS + rowOffset, SP35_PLANE1_FIRST + 8
        ApplyStatusCell target, HEADER_ROWS + rowOffset, SP35_PLANE1_FIRST + 14
        ApplyStatusCell target, HEADER_ROWS + rowOffset, SP35_PLANE2_FIRST + 8
        ApplyStatusCell target, HEADER_ROWS + rowOffset, SP35_PLANE2_FIRST + 14
        ApplyStatusCell target, HEADER_ROWS + rowOffset, SP63_PLANE1_FIRST + 11
        ApplyStatusCell target, HEADER_ROWS + rowOffset, SP63_PLANE2_FIRST + 11
    Next i
End Sub

' Заполняет все статусные колонки нормативных ветвей значением N/A до записи выбранной методики.
Private Sub PutDefaultStatusCells(ByRef data() As Variant, ByVal rowIndex As Long, ByVal statusText As String)
    data(rowIndex, SP35_PLANE1_FIRST + 8) = statusText
    data(rowIndex, SP35_PLANE1_FIRST + 14) = statusText
    data(rowIndex, SP35_PLANE2_FIRST + 8) = statusText
    data(rowIndex, SP35_PLANE2_FIRST + 14) = statusText
    data(rowIndex, SP63_PLANE1_FIRST + 11) = statusText
    data(rowIndex, SP63_PLANE2_FIRST + 11) = statusText
End Sub

' Красит одну статусную ячейку detailed-блока устойчивости по единой палитре Results.
Private Sub ApplyStatusCell(ByVal target As Object, ByVal rowIndex As Long, ByVal columnIndex As Long)
    Dim cell As Object
    Set cell = target.Cells(rowIndex, columnIndex)
    ApplyStatusFill cell, CStr(cell.Value2)
End Sub

' Объединяет ячейки шапки и записывает текст.
Private Sub MergeCells(ByVal target As Object, ByVal rowIndex As Long, ByVal columnIndex As Long, _
        ByVal rowCount As Long, ByVal columnCount As Long, ByVal valueText As String)
    Dim area As Object
    Set area = target.Offset(rowIndex - 1, columnIndex - 1).Resize(rowCount, columnCount)
    area.Merge
    area.Value2 = valueText
End Sub

' Возвращает ширину блока по его первой и последней колонке.
Private Function BlockWidth(ByVal firstColumn As Long, ByVal lastColumn As Long) As Long
    BlockWidth = lastColumn - firstColumn + 1
End Function

' Делает нижние индексы в обозначениях без изменения самого текста заголовка.
Private Sub ApplyHeaderSubscripts(ByVal target As Object)
    Dim labels As Object
    Set labels = target.Offset(4, 0).Resize(1, TOTAL_COLUMNS)

    Dim cell As Object
    For Each cell In labels.Cells
        ApplySubscriptToken cell, "ec,sl", 2, 4
        ApplySubscriptToken cell, "ec", 2, 1
        ApplySubscriptToken cell, "ea", 2, 1
        ApplySubscriptToken cell, "e0", 2, 1
        ApplySubscriptToken cell, "l01", 2, 2
        ApplySubscriptToken cell, "l02", 2, 2
        ApplySubscriptToken cell, "l0/i", 2, 1
        ApplySubscriptToken cell, "Ncr", 2, 2
        ApplySubscriptToken cell, "Nult", 2, 3
        ApplySubscriptToken cell, "M1", 2, 1
        ApplySubscriptToken cell, "M2", 2, 1
        ApplySubscriptToken cell, ChrW$(&H3C6) & "p", 2, 1
        ApplySubscriptToken cell, ChrW$(&H3C6) & "l", 2, 1
        ApplySubscriptToken cell, ChrW$(&H3C6) & "m", 2, 1
        ApplySubscriptToken cell, ChrW$(&H3B4) & "e", 2, 1
        ApplySubscriptToken cell, "kb", 2, 1
        ApplySubscriptToken cell, "ks", 2, 1
    Next cell
End Sub

' Применяет нижний индекс к найденному техническому обозначению.
Private Sub ApplySubscriptToken(ByVal cell As Object, ByVal token As String, ByVal subStart As Long, ByVal subLength As Long)
    Dim valueText As String
    valueText = CStr(cell.Value2)
    If Len(valueText) = 0 Then Exit Sub

    Dim position As Long
    position = InStr(1, valueText, token, vbBinaryCompare)
    If position <= 0 Then Exit Sub

    On Error Resume Next
    cell.Characters(position + subStart - 1, subLength).Font.Subscript = True
    On Error GoTo 0
End Sub

' Возвращает статус одной расчетной плоскости по общему пользовательскому словарю.
Private Function PlaneStatus(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As String
    Dim policy As CBatchStatusPolicy
    Set policy = StatusPolicy()
    If Not PlaneApplicable(batch, index, planeIndex) Then
        PlaneStatus = policy.NotApplicable
    ElseIf PlanePassed(batch, index, planeIndex) Then
        PlaneStatus = policy.OK
    Else
        PlaneStatus = policy.Fail
    End If
End Function

' Однотипные accessors ниже выбирают данные первой или второй главной
' плоскости из batch-результата. Они держат разметку таблицы независимой от
' конкретных имен свойств Stability...1/Stability...2.
Private Function PlaneApplicable(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As Boolean
    If planeIndex = 1 Then PlaneApplicable = batch.StabilityPlaneApplicable1(index) Else PlaneApplicable = batch.StabilityPlaneApplicable2(index)
End Function

' Возвращает итог PASS/FAIL одной плоскости устойчивости.
Private Function PlanePassed(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As Boolean
    If planeIndex = 1 Then PlanePassed = batch.StabilityPlanePassed1(index) Else PlanePassed = batch.StabilityPlanePassed2(index)
End Function

' Возвращает расчетную ветвь одной плоскости: SP35-table, SP35-eta или SP63.
Private Function PlaneBranch(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As String
    If planeIndex = 1 Then PlaneBranch = batch.StabilityPlaneBranch1(index) Else PlaneBranch = batch.StabilityPlaneBranch2(index)
End Function

' Возвращает критическую силу Ncr одной плоскости.
Private Function PlaneNcr(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As Double
    If planeIndex = 1 Then PlaneNcr = batch.StabilityNcr1(index) Else PlaneNcr = batch.StabilityNcr2(index)
End Function

' Возвращает табличную несущую силу Nult для ветви СП 35 ec <= r.
Private Function PlaneNultimate(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As Double
    If planeIndex = 1 Then PlaneNultimate = batch.StabilityNultimate1(index) Else PlaneNultimate = batch.StabilityNultimate2(index)
End Function

' Возвращает коэффициент eta одной плоскости.
Private Function PlaneEta(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As Double
    If planeIndex = 1 Then PlaneEta = batch.StabilityEta1(index) Else PlaneEta = batch.StabilityEta2(index)
End Function

' Возвращает коэффициент длительности phi_l одной плоскости.
Private Function PlanePhiL(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As Double
    If planeIndex = 1 Then PlanePhiL = batch.StabilityPhiL1(index) Else PlanePhiL = batch.StabilityPhiL2(index)
End Function

' Возвращает табличный phi_l для ветви СП 35 ec <= r.
Private Function PlanePhiLTable(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As Double
    If planeIndex = 1 Then PlanePhiLTable = batch.StabilityPhiLTable1(index) Else PlanePhiLTable = batch.StabilityPhiLTable2(index)
End Function

' Возвращает табличный phi_m для ветви СП 35 ec <= r.
Private Function PlanePhiM(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As Double
    If planeIndex = 1 Then PlanePhiM = batch.StabilityPhiM1(index) Else PlanePhiM = batch.StabilityPhiM2(index)
End Function

' Возвращает итоговый phi для табличной ветви СП 35.
Private Function PlanePhiValue(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As Double
    If planeIndex = 1 Then PlanePhiValue = batch.StabilityPhiValue1(index) Else PlanePhiValue = batch.StabilityPhiValue2(index)
End Function

' Возвращает полный эксцентриситет ec/e0 одной плоскости.
Private Function PlaneEccentricity(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As Double
    If planeIndex = 1 Then PlaneEccentricity = batch.StabilityEccentricity1(index) Else PlaneEccentricity = batch.StabilityEccentricity2(index)
End Function

' Возвращает случайный эксцентриситет ec,sl/ea одной плоскости.
Private Function PlaneAccidentalEccentricity(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As Double
    If planeIndex = 1 Then PlaneAccidentalEccentricity = batch.StabilityAccidentalEcc1(index) Else PlaneAccidentalEccentricity = batch.StabilityAccidentalEcc2(index)
End Function

' Возвращает гибкость l0/i одной плоскости.
Private Function PlaneSlenderness(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As Double
    If planeIndex = 1 Then PlaneSlenderness = batch.StabilitySlenderness1(index) Else PlaneSlenderness = batch.StabilitySlenderness2(index)
End Function

' Возвращает ядровое расстояние r в направлении эксцентриситета.
Private Function PlaneCoreDistance(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As Double
    If planeIndex = 1 Then PlaneCoreDistance = batch.StabilityCoreDistance1(index) Else PlaneCoreDistance = batch.StabilityCoreDistance2(index)
End Function

' Возвращает расчетную delta одной плоскости.
Private Function PlaneDelta(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As Double
    If planeIndex = 1 Then PlaneDelta = batch.StabilityDelta1(index) Else PlaneDelta = batch.StabilityDelta2(index)
End Function

' Возвращает коэффициент kb для ветви СП 63.
Private Function PlaneKb(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As Double
    If planeIndex = 1 Then PlaneKb = batch.StabilityKb1(index) Else PlaneKb = batch.StabilityKb2(index)
End Function

' Возвращает коэффициент ks для ветви СП 63.
Private Function PlaneKs(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As Double
    If planeIndex = 1 Then PlaneKs = batch.StabilityKs1(index) Else PlaneKs = batch.StabilityKs2(index)
End Function

' Возвращает жесткость D одной плоскости.
Private Function PlaneStiffnessD(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As Double
    If planeIndex = 1 Then PlaneStiffnessD = batch.StabilityStiffnessD1(index) Else PlaneStiffnessD = batch.StabilityStiffnessD2(index)
End Function

' Возвращает проверочный запас одной плоскости.
Private Function PlaneReserve(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal planeIndex As Long) As Double
    If planeIndex = 1 Then PlaneReserve = batch.StabilityReserve1(index) Else PlaneReserve = batch.StabilityReserve2(index)
End Function

' Возвращает пустую ячейку для численного нуля, чтобы summary не шумела
' несущественными расчетными величинами.
Private Function OptionalValue(ByVal value As Double) As Variant
    If Abs(value) > SMALL_VALUE Then OptionalValue = value Else OptionalValue = vbNullString
End Function

' Выводит сжимающую силу только когда она физически есть; иначе возвращает
' стандартный N/A из политики статусов.
Private Function OptionalCompressionForce(ByVal compressionMagnitude As Double, ByVal units As CUnitSystem) As Variant
    If compressionMagnitude > SMALL_VALUE Then OptionalCompressionForce = OutputCompressionForce(compressionMagnitude, units) Else OptionalCompressionForce = StatusPolicy().NotApplicable
End Function

' Возвращает коэффициент запаса limit/Ncompression для ограничения СП 35.
Private Function RatioLimitToCompression(ByVal limitCompression As Double, ByVal nValue As Double) As Variant
    Dim compression As Double
    compression = CompressionMagnitude(nValue)
    If compression <= SMALL_VALUE Or limitCompression <= SMALL_VALUE Then
        RatioLimitToCompression = vbNullString
    Else
        RatioLimitToCompression = limitCompression / compression
    End If
End Function

' Преобразует signed N во внутреннюю положительную величину сжатия.
Private Function CompressionMagnitude(ByVal nValue As Double) As Double
    Dim loadState As CSectionLoadState
    Set loadState = New CSectionLoadState
    loadState.Initialize nValue, 0#, 0#, 0#, 0#
    CompressionMagnitude = loadState.CompressionMagnitude(SMALL_VALUE)
End Function

' Переводит положительную величину сжатия в пользовательские единицы силы
' с учетом принятого знака Output.SignConvention.N.
Private Function OutputCompressionForce(ByVal compressionMagnitude As Double, ByVal units As CUnitSystem) As Double
    If compressionMagnitude <= SMALL_VALUE Then Exit Function
    OutputCompressionForce = OutputForce(-Abs(compressionMagnitude), units)
End Function

' Выводит подписанную сжимающую силу как пользовательское значение N.
Private Function OutputSignedCompressionForce(ByVal signedCompression As Double, ByVal units As CUnitSystem) As Double
    OutputSignedCompressionForce = OutputForce(-signedCompression, units)
End Function

' Группа функций Output... ниже только переводит внутренние единицы устойчивости
' в пользовательские выходные единицы для Results.
Private Function OutputLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputLength = value Else OutputLength = units.InternalLengthToOutput(value)
End Function

' Возвращает подпись выходной единицы длины.
Private Function OutputLengthUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputLengthUnit = "mm" Else OutputLengthUnit = units.OutputLengthUnit
End Function

' Переводит внутреннюю силу в пользовательские выходные единицы.
Private Function OutputForce(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputForce = value Else OutputForce = units.InternalForceToOutput(value)
End Function

' Возвращает подпись выходной единицы силы.
Private Function OutputForceUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputForceUnit = "N" Else OutputForceUnit = units.OutputForceUnit
End Function

' Переводит внутренний Mx в пользовательские выходные единицы и знак.
Private Function OutputMx(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMx = value Else OutputMx = units.InternalMomentMxToOutput(value)
End Function

' Переводит внутренний My в пользовательские выходные единицы и знак.
Private Function OutputMy(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMy = value Else OutputMy = units.InternalMomentMyToOutput(value)
End Function

' Переводит модуль момента без назначения знака конкретной оси.
Private Function OutputMomentMagnitude(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMomentMagnitude = value Else OutputMomentMagnitude = units.InternalMomentMagnitudeToOutput(value)
End Function

' Возвращает подпись выходной единицы момента.
Private Function OutputMomentUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputMomentUnit = "N*mm" Else OutputMomentUnit = units.OutputMomentUnit
End Function

' Переводит жесткость D из Н*мм2 в выбранные выходные единицы.
Private Function OutputStiffness(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputStiffness = value Else OutputStiffness = units.InternalForceLengthSquaredToOutput(value)
End Function

' Возвращает подпись выходной единицы жесткости.
Private Function OutputStiffnessUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputStiffnessUnit = "N*mm2" Else OutputStiffnessUnit = units.OutputForceLengthSquaredUnit
End Function

' Утилитарный максимум без WorksheetFunction для форматирования таблицы.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

' Возвращает общую политику отображения статусов PASS/FAIL/N/A.
Private Function StatusPolicy() As CBatchStatusPolicy
    Static policy As CBatchStatusPolicy
    If policy Is Nothing Then Set policy = New CBatchStatusPolicy
    Set StatusPolicy = policy
End Function
