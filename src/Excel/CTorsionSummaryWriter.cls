VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTorsionSummaryWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Вывод результатов кручения и суммы напряжений NMT на лист Results
' ==========================================================================
' Сохраняет готовые результаты в матрицы снимка, восстанавливает подробную
' и поэлементную таблицы, готовит шесть колонок общей сводки. НДМ не решает;
' изменяет только собственные области Results без изменения ширин и высот.

' Число строк данных включает доступные строки таблицы сочетаний, как у остальных сводок.
Public Function RequiredDataRowsForWorkbook(ByVal workbook As Object, ByVal actual As Long) As Long
    RequiredDataRowsForWorkbook = TorsionDataRows(workbook, actual)
End Function

' Высота подробной таблицы с четырьмя строками шапки; синий заголовок расположен отдельно над ней.
Public Function RequiredRowsForWorkbook(ByVal workbook As Object, ByVal actual As Long) As Long
    RequiredRowsForWorkbook = RequiredDataRowsForWorkbook(workbook, actual) + HeaderRowsAboveAnchor
End Function

' Число строк шапки над первой строкой результатов кручения.
Public Property Get HeaderRowsAboveAnchor() As Long
    HeaderRowsAboveAnchor = 4
End Property

' Полная ширина подробной таблицы, включая промежуток между её частями.
Public Function RequiredColumns() As Long
    RequiredColumns = TORSION_POST_DISPLAY_COLUMNS
End Function

' Собирает строку IO-контракта из смысловых свойств; неприменимые числа остаются Empty.
Public Function CreateSummaryData(ByVal result As CTorsionPostResult) As Variant
    Dim data(1 To 1, 1 To TORSION_POST_COLUMNS) As Variant
    data(1, tscCombinationID) = result.CombinationID
    data(1, tscLongitudinalStatus) = result.LongitudinalStatus
    data(1, tscHoopStatus) = result.HoopStatus
    data(1, tscConcreteStatus) = result.ConcreteStatus
    data(1, tscCombinedStatus) = result.CombinedStatus
    data(1, tscOverallStatus) = result.OverallStatus
    If result.HasResultComment Then data(1, tscResultComment) = result.ResultComment
    If result.HasTResults Then
        data(1, tscStateName) = result.StateName
        data(1, tscTorque) = result.Torque
        data(1, tscEnclosedArea) = result.EnclosedArea
        data(1, tscHoopPerimeter) = result.HoopPerimeter
        data(1, tscCotangent) = result.Cotangent
        data(1, tscShearFlow) = result.ShearFlow
        data(1, tscLongitudinalForce) = result.LongitudinalForce
        data(1, tscLongitudinalArea) = result.LongitudinalArea
        data(1, tscRequiredLongitudinalArea) = result.RequiredLongitudinalArea
        data(1, tscCriticalTRebarID) = result.CriticalTRebarID
        data(1, tscMaximumTStress) = result.MaximumTStress
        data(1, tscLongitudinalResistance) = result.LongitudinalResistance
        data(1, tscLongitudinalReserve) = result.LongitudinalReserve
        data(1, tscHoopArea) = result.HoopArea
        data(1, tscHoopSpacing) = result.HoopSpacing
        data(1, tscHoopStress) = result.HoopStress
        data(1, tscHoopResistance) = result.HoopResistance
        data(1, tscHoopReserve) = result.HoopReserve
        data(1, tscConcreteMethod) = result.ConcreteMethod
        data(1, tscBarBand) = result.BarBand
        data(1, tscCriticalTPlacement) = result.CriticalTPlacement
        If result.HasConcreteCapacity Then
            data(1, tscConcreteThickness) = result.ConcreteThickness
            data(1, tscConcreteTorqueLimit) = result.ConcreteTorqueLimit
            data(1, tscConcreteReserve) = result.ConcreteReserve
        End If
        If result.HasConcreteStress Then
            data(1, tscConcreteStress) = result.ConcreteStress
            data(1, tscConcreteResistance) = result.ConcreteResistance
            data(1, tscEffectiveConcreteArea) = result.EffectiveConcreteArea
            data(1, tscConcreteDesignStrength) = result.ConcreteDesignStrength
            data(1, tscConcreteEfficiency) = result.ConcreteEfficiency
        End If
        If result.HasCombinedResults Then
            data(1, tscCriticalNMTRebarID) = result.CriticalNMTRebarID
            data(1, tscNormalStress) = result.NormalStress
            data(1, tscTStressAtCriticalNMT) = result.TStressAtCriticalNMT
            data(1, tscCombinedForce) = result.CombinedForce
            data(1, tscCombinedStress) = result.CombinedStress
            data(1, tscCombinedResistance) = result.CombinedResistance
            data(1, tscCriticalNMTPlacement) = result.CriticalNMTPlacement
            If result.HasCombinedReserve Then data(1, tscCombinedReserve) = result.CombinedReserve
        End If
    End If
    CreateSummaryData = data
End Function

' Собирает подробные строки стержней для Excel и снимка без изменения объекта результата.
Public Function CreateBarData(ByVal result As CTorsionPostResult) As Variant
    If result.BarCount = 0 Then Exit Function
    Dim data() As Variant, i As Long
    ReDim data(1 To result.BarCount, 1 To TORSION_BAR_COLUMNS)
    For i = 1 To result.BarCount
        data(i, tbcCombinationID) = result.CombinationID
        data(i, tbcRebarID) = result.BarID(i)
        data(i, tbcPlacement) = result.BarPlacement(i)
        data(i, tbcForce) = result.BarForce(i)
        data(i, tbcAllocatedLength) = result.BarLength(i)
    Next i
    CreateBarData = data
End Function

' Сохраняет расчётные значения и поэлементные усилия отдельными матрицами.
Public Sub CaptureSummary(ByVal snapshot As CCalculationSnapshot, ByVal results As Collection, _
        ByVal units As CUnitSystem)
    Dim data() As Variant, bars() As Variant, summary As Variant, values As Variant
    Dim row As Long, column As Long, count As Long, barRows As Long, barRow As Long, i As Long
    Dim result As CTorsionPostResult
    count = results.Count: If count < 1 Then count = 1
    ReDim data(1 To count, 1 To TORSION_POST_COLUMNS)
    For Each result In results
        barRows = barRows + result.BarCount
    Next result
    ReDim bars(1 To barRows + 1, 1 To TORSION_BAR_COLUMNS)
    bars(1, 1) = "LoadCase": bars(1, 2) = "Rebar": bars(1, 3) = "Грань"
    bars(1, 4) = "F_T, " & units.OutputForceUnit: bars(1, 5) = "l_i, " & units.OutputLengthUnit
    barRow = 1
    For Each result In results
        row = row + 1: summary = CreateSummaryData(result)
        For column = 1 To TORSION_POST_COLUMNS
            data(row, column) = TorsionSummaryOutputValue(summary(1, column), column, units)
        Next column
        values = CreateBarData(result)
        If IsArray(values) Then
            For i = 1 To UBound(values, 1)
                barRow = barRow + 1
                bars(barRow, 1) = values(i, 1): bars(barRow, 2) = values(i, 2)
                bars(barRow, 3) = values(i, 3)
                bars(barRow, 4) = Abs(units.InternalForceToOutput(CDbl(values(i, 4))))
                bars(barRow, 5) = units.InternalLengthToOutput(CDbl(values(i, 5)))
            Next i
        End If
    Next result
    snapshot.AddTable TORSION_POST_HEADERS_KEY, SummaryHeaders(units)
    snapshot.AddTable TORSION_POST_DATA_KEY, data
    snapshot.AddTable TORSION_BAR_DATA_KEY, bars
End Sub

' Преобразует числа только через общий CUnitSystem; пустые поля и текст
' не превращаются в нули. Запасы и угловые коэффициенты безразмерны.
Private Function TorsionSummaryOutputValue(ByVal value As Variant, ByVal column As Long, ByVal units As CUnitSystem) As Variant
    TorsionSummaryOutputValue = value
    If IsEmpty(value) Then Exit Function
    If column = 39 Then
        TorsionSummaryOutputValue = Replace$(Replace$(CStr(value), vbCrLf, "; "), vbLf, "; ") ' Подробная причина остаётся компактным пояснением Results; цепочки адресов сохраняются.
        Exit Function
    End If
    If Not IsNumeric(value) Then Exit Function
    Select Case column
        Case 3, 26: TorsionSummaryOutputValue = units.InternalMomentMagnitudeToOutput(CDbl(value))
        Case 4, 9, 10, 16, 40: TorsionSummaryOutputValue = units.InternalAreaToOutput(CDbl(value))
        Case 5, 17, 23, 37: TorsionSummaryOutputValue = units.InternalLengthToOutput(CDbl(value))
        Case 7: TorsionSummaryOutputValue = Abs(units.InternalForceToOutput(CDbl(value))) / units.InternalLengthToOutput(1#)
        Case 8: TorsionSummaryOutputValue = Abs(units.InternalForceToOutput(CDbl(value)))
        Case 32: TorsionSummaryOutputValue = Sgn(CDbl(value)) * Abs(units.InternalForceToOutput(CDbl(value))) ' Положительное усилие блока NMT обозначает растяжение.
        Case 12, 13, 18, 19, 24, 25, 30, 31, 33, 34, 41: TorsionSummaryOutputValue = units.InternalStressToOutput(CDbl(value))
    End Select
End Function

' Возвращает порядок логических колонок в подробной таблице. Нули —
' отдельный промежуток и нагрузки из готовой общей сводки, а не новые расчёты.
Public Function DisplayColumns() As Variant
    DisplayColumns = Array(1, 39, 38, 0, 0, 0, 3, 0, 2, 37, 4, 5, 6, 7, 8, 9, 10, 11, 43, 12, 13, 14, 15, _
        16, 17, 18, 19, 20, 21, 22, 23, 40, 41, 42, 24, 25, 26, 27, 28, 29, 44, 30, 31, 32, 33, 34, 35, 36)
End Function

' Заполняет только самостоятельную таблицу после устойчивости и нижнюю
' техническую таблицу. Отказ вывода не изменяет готовое НДС и его статусы.
Public Function RestoreSummary(ByVal workbook As Object, ByVal source As Object) As String
    On Error GoTo Failed
    Dim anchor As Object: Set anchor = TorsionAnchor(workbook, "rngTorsionSummaryAnchor")
    If anchor Is Nothing Then
        If source.HasTable(TORSION_POST_DATA_KEY) Then Err.Raise 5, , "Не найден якорь rngTorsionSummaryAnchor."
        Exit Function
    End If
    If Not source.HasTable(TORSION_POST_DATA_KEY) Then
        If TypeOf source Is CResultReadSource Or source.IsGeometryPreview Then RestoreSummary = ClearSummary(workbook)
        Exit Function
    End If
    Dim data As Variant, headers As Variant, bars As Variant, summary As Variant, display() As Variant, map As Variant
    data = source.Table(TORSION_POST_DATA_KEY): bars = source.Table(TORSION_BAR_DATA_KEY)
    Dim restorePlacement As Boolean: restorePlacement = UBound(data, 2) = 39 Or UBound(bars, 2) < TORSION_BAR_COLUMNS
    headers = TorsionHeadersForSource(source)
    ValidateSnapshotTableShape TORSION_POST_DATA_KEY, data
    data = TorsionDataForDisplay(data)
    ValidateSnapshotTableShape TORSION_BAR_DATA_KEY, bars
    bars = TorsionBarsWithLengths(bars, data, headers)
    If restorePlacement Then FillSavedPlacementNames source, data, bars
    If source.HasTable("Summary") Then
        Dim writer As New CBatchResultWriter: summary = writer.SummaryValuesForDisplay(source)
    End If
    Dim rows As Long: rows = TorsionDataRows(workbook, UBound(data, 1))
    ReDim display(1 To rows, 1 To TORSION_POST_DISPLAY_COLUMNS)
    map = DisplayColumns()
    Dim row As Long, column As Long, other As Long
    For row = 1 To UBound(data, 1)
        For column = 1 To TORSION_POST_DISPLAY_COLUMNS
            If map(column - 1) > 0 Then display(row, column) = data(row, map(column - 1))
        Next column
        If IsArray(summary) And Len(CStr(data(row, 1))) > 0 Then
            For other = 13 To UBound(summary, 1)
                If StrComp(CStr(summary(other, 1)), CStr(data(row, 1)), vbTextCompare) = 0 Then
                    For column = 4 To 6: display(row, column) = summary(other, column + 31): Next column
                    Exit For
                End If
            Next other
        End If
    Next row
    Dim target As Object: Set target = anchor.Offset(-4, 0).Resize(rows + 4, TORSION_POST_DISPLAY_COLUMNS)
    ValidateTorsionTarget workbook, target.Offset(-1, 0).Resize(rows + 5, target.Columns.Count)
    Dim barAnchor As Object: Set barAnchor = TorsionAnchor(workbook, "rngTorsionBarForces")
    If barAnchor Is Nothing Then Err.Raise 5, , "Не найден якорь rngTorsionBarForces."
    Dim barCount As Long: barCount = UBound(bars, 1)
    If Not ResultRowsFitSheet(barAnchor.Row, barCount, barAnchor.Worksheet.Rows.Count) Then _
        Err.Raise 5, , ResultRowLimitDescription("rngTorsionBarForces", barAnchor.Row, barCount, barAnchor.Worksheet.Rows.Count)
    ValidateTorsionTarget workbook, barAnchor.Resize(barCount, TORSION_BAR_COLUMNS)
    If Not workbook.Application.Intersect(target, barAnchor.Resize(barCount, TORSION_BAR_COLUMNS)) Is Nothing Then _
        Err.Raise 5, , "Якоря подробной и поэлементной таблиц T пересекаются."
    target.UnMerge: target.ClearContents
    target.Resize(4, TORSION_POST_DISPLAY_COLUMNS).Value2 = headers
    target.Cells(5, 1).Resize(rows, TORSION_POST_DISPLAY_COLUMNS).Value2 = display
    FormatTorsionDetailed target, rows
    With anchor.Offset(-5, 0).Resize(1, TORSION_POST_DISPLAY_COLUMNS)
        .Interior.Color = RGB(189, 215, 238): .Borders.LineStyle = -4142
    End With
    With anchor.Offset(-5, 0)
        .Value2 = "Расчет на приближенный учет кручения в сечении NMT"
        .Font.Name = "Arial": .Font.Size = 12: .Font.Bold = True: .Font.Color = RGB(255, 0, 0)
        .HorizontalAlignment = -4131: .WrapText = False
    End With
    WriteTorsionBarForces barAnchor, bars, barCount
    Exit Function
Failed:
    RestoreSummary = WorkbookDialogText("Не удалось вывести T (кручение) / NMT. " & Err.Description & _
        vbCrLf & "Результаты НДМ не изменялись.", workbook)
End Function

' Очищает только данные своих якорей после предварительной проверки;
' шапки, имена, ширины столбцов и настройки пользователя сохраняются.
Public Function ClearSummary(ByVal workbook As Object) As String
    On Error GoTo Failed
    Dim anchor As Object, target As Object, count As Long
    Set anchor = TorsionAnchor(workbook, "rngTorsionSummaryAnchor")
    If Not anchor Is Nothing Then
        count = TorsionDataRows(workbook, 1)
        Set target = anchor.Resize(count, TORSION_POST_DISPLAY_COLUMNS)
        ValidateTorsionTarget workbook, target: target.ClearContents
    End If
    Set anchor = TorsionAnchor(workbook, "rngTorsionBarForces")
    If Not anchor Is Nothing Then
        count = anchor.Worksheet.Cells(anchor.Worksheet.Rows.Count, anchor.Column).End(-4162).Row - anchor.Row
        If count > 0 Then
            Set target = anchor.Offset(1, 0).Resize(count, TORSION_BAR_COLUMNS)
            ValidateTorsionTarget workbook, target: target.ClearContents: target.Borders.LineStyle = -4142
        End If
    End If
    Exit Function
Failed:
    ClearSummary = WorkbookDialogText("Данные T (кручение) не очищены: " & Err.Description, workbook)
End Function

' Дополняет трёхколоночную матрицу длиной по её собственным сохранённым
' усилиям, потоку и углу. Числа находятся в одних OUTPUT-единицах, поэтому
' F/(q*cot theta) сразу даёт OUTPUT Length без чтения листов или геометрии.
Private Function TorsionBarsWithLengths(ByRef bars As Variant, ByRef data As Variant, ByRef headers As Variant) As Variant
    If UBound(bars, 2) = TORSION_BAR_COLUMNS Then TorsionBarsWithLengths = bars: Exit Function
    Dim result() As Variant, row As Long, column As Long, key As String
    ReDim result(1 To UBound(bars, 1), 1 To TORSION_BAR_COLUMNS)
    For row = 1 To UBound(bars, 1)
        For column = 1 To 2: result(row, column) = bars(row, column): Next column
        result(row, 4) = bars(row, 3)
    Next row
    Dim label As String: label = CStr(headers(4, 12))
    result(1, 3) = "Грань": result(1, 5) = "l_i, " & Trim$(Mid$(label, InStr(1, label, ",", vbBinaryCompare) + 1))
    Dim factors As Object: Set factors = CreateObject("Scripting.Dictionary")
    factors.CompareMode = vbTextCompare
    For row = 1 To UBound(data, 1)
        key = CStr(data(row, 1))
        If Len(key) > 0 And IsNumeric(data(row, 7)) And IsNumeric(data(row, 6)) Then
            If CDbl(data(row, 7)) > 0# And CDbl(data(row, 6)) > 0# Then factors(key) = CDbl(data(row, 7)) * CDbl(data(row, 6))
        End If
    Next row
    For row = 2 To UBound(bars, 1)
        key = CStr(bars(row, 1))
        If UBound(bars, 2) = 4 Then
            result(row, 5) = bars(row, 4)
        Else
            If Not factors.Exists(key) Then Err.Raise 5, TypeName(Me), _
                "Не удалось определить длину участка арматуры: в сохранённом сочетании " & key & " отсутствуют положительные поток q или котангенс угла раскосов."
            result(row, 5) = CDbl(bars(row, 3)) / CDbl(factors(key))
        End If
    Next row
    TorsionBarsWithLengths = result
End Function

' В форматах без сохранённых подписей читает их из общей геометрии по точному ElementID.
' Для текущих матриц подписи уже принадлежат расчёту, поэтому геометрия повторно не читается.
Private Sub FillSavedPlacementNames(ByVal source As Object, ByRef data As Variant, ByRef bars As Variant)
    If Not source.HasTable("rngNDMSectionGeometry") Then Exit Sub
    Dim geometry As Variant: geometry = source.Table("rngNDMSectionGeometry")
    Dim column As Long, idColumn As Long, commentColumn As Long, row As Long
    For column = 1 To UBound(geometry, 2)
        If CStr(geometry(1, column)) = "ElementID" Then idColumn = column
        If CStr(geometry(1, column)) = "Comment" Then commentColumn = column
    Next column
    If idColumn = 0 Or commentColumn = 0 Then Exit Sub
    Dim names As Object: Set names = CreateObject("Scripting.Dictionary"): names.CompareMode = vbBinaryCompare
    For row = 2 To UBound(geometry, 1)
        names(CStr(geometry(row, idColumn))) = CStr(geometry(row, commentColumn))
    Next row
    Dim key As String
    For row = 1 To UBound(data, 1)
        key = CStr(data(row, 11)): If names.Exists(key) Then data(row, 43) = names(key)
        key = CStr(data(row, 29)): If names.Exists(key) Then data(row, 44) = names(key)
    Next row
    For row = 2 To UBound(bars, 1)
        key = CStr(bars(row, 2)): If names.Exists(key) Then bars(row, 3) = names(key)
    Next row
End Sub

' Извлекает единицы из сохранённой короткой подписи; текущие единицы Config не участвуют.
Private Function HeaderUnit(ByVal label As Variant) As String
    Dim comma As Long: comma = InStrRev(CStr(label), ",", -1, vbBinaryCompare) ' В индексе обозначения также может быть запятая, например sigma_T,max.
    If comma > 0 Then HeaderUnit = Trim$(Mid$(CStr(label), comma + 1))
End Function

' Читает конкретный контракт из 39 колонок без выдуманных параметров EC2.
' Сохранённые усилия и статусы не меняются; происхождение проверки Strut поясняется явно.
Private Function TorsionDataForDisplay(ByRef data As Variant) As Variant
    If UBound(data, 2) = TORSION_POST_COLUMNS Then TorsionDataForDisplay = data: Exit Function
    Dim output() As Variant, row As Long, column As Long
    ReDim output(1 To UBound(data, 1), 1 To TORSION_POST_COLUMNS)
    For row = 1 To UBound(data, 1)
        For column = 1 To UBound(data, 2): output(row, column) = data(row, column): Next column
        If CStr(output(row, 22)) = "Strut" Then
            output(row, 39) = CStr(output(row, 39)) & " Снимок TorsionPost.Data из 39 колонок: сопротивление раскоса задано как коэффициент * R_b; A_k, f_cd и nu по EC2 не сохранены. Для проверки по источнику выполните расчёт повторно."
        End If
    Next row
    TorsionDataForDisplay = output
End Function

' Требует одноячеечный якорь собственной книги на Results.
Private Function TorsionAnchor(ByVal workbook As Object, ByVal name As String) As Object
    Dim target As Object
    On Error Resume Next
    Set target = workbook.Names.Item(name).RefersToRange
    On Error GoTo 0
    If target Is Nothing Then Exit Function
    If Not target.Worksheet.Parent Is workbook Or target.Worksheet.Name <> "Results" Or target.Cells.CountLarge <> 1 Then _
        Err.Raise 5, , name & " должен ссылаться на одну ячейку листа Results текущей книги."
    Set TorsionAnchor = target
End Function

' Высота подробной таблицы соответствует доступным строкам сочетаний.
Private Function TorsionDataRows(ByVal workbook As Object, ByVal actual As Long) As Long
    Dim loads As Object: Set loads = workbook.Names.Item("rngLoadCombinations").RefersToRange
    TorsionDataRows = GeomMax(actual, loads.Rows.Count - 1)
    If TorsionDataRows < 1 Then TorsionDataRows = 1
End Function

' До записи исключает пересечения с основной сводкой, подробными таблицами
' НДМ и всеми техническими якорями. При ошибке диапазонов данные не очищаются.
Private Sub ValidateTorsionTarget(ByVal workbook As Object, ByVal target As Object)
    Dim item As Variant, other As Object, name As String, firstRow As Long, rows As Long, columns As Long
    Dim names As Variant: names = Array("rngBatchSummary", "rngStrengthSummaryAnchor", "rngCrackSummaryAnchor", "rngStabilitySummaryAnchor", _
        "rngNDMSectionGeometry", "rngNDMSectionProperties", "rngNDMMaterialDiagrams", "rngNDMSectionContours", _
        "rngNDMElementResults", "rngNDMElementGeometry", "rngNDMSectionAnnotations", "rngNDMElementResultIndex")
    Dim widths As Variant: widths = Array(43, 49, 72, STABILITY_SNAPSHOT_COLUMNS, 16, 6, 11, 17, 9, 18, 24, 9)
    Dim index As Long
    For index = 0 To UBound(names)
        name = CStr(names(index)): Set other = workbook.Names.Item(name).RefersToRange
        columns = widths(index): firstRow = other.Row
        If index = 0 Then
            rows = 42
        ElseIf index <= 3 Then
            rows = TorsionDataRows(workbook, 1)
            firstRow = firstRow - 5
            If index = 3 Then firstRow = firstRow - 1
            rows = rows + other.Row - firstRow
        Else
            rows = target.Worksheet.Rows.Count - firstRow + 1
        End If
        If target.Row <= firstRow + rows - 1 And target.Row + target.Rows.Count - 1 >= firstRow And _
                target.Column <= other.Column + columns - 1 And target.Column + target.Columns.Count - 1 >= other.Column Then _
            Err.Raise 5, , "Якорь вывода T пересекает " & name & ". Исправьте расположение именованных диапазонов."
    Next index
End Sub

' Оформляет шапку и данные по правилам подробных таблиц Results;
' пустой столбец-разделитель не имеет рамки, ровно rows строк данных.
Private Sub FormatTorsionDetailed(ByVal target As Object, ByVal rows As Long)
    With target
        .Font.Name = "Arial": .Font.Size = 9: .Font.Color = 0: .Font.Bold = False
        .HorizontalAlignment = -4108: .VerticalAlignment = -4108: .WrapText = False
        .NumberFormat = "0.000": .Interior.Pattern = -4142
        .Borders.LineStyle = 1: .Borders.Weight = 2: .Borders.Color = 0
    End With
    With target.Resize(4, target.Columns.Count)
        .Font.Bold = True: .Interior.Color = RGB(217, 217, 217): .NumberFormat = "General": .WrapText = True
    End With
    target.Columns(8).Clear
    target.Cells(1, 1).Resize(4, 1).Merge
    target.Cells(1, 2).Resize(4, 1).Merge
    target.Cells(1, 3).Resize(4, 1).Merge
    target.Cells(1, 4).Resize(2, 4).Merge
    target.Cells(1, 9).Resize(1, TORSION_POST_DISPLAY_COLUMNS - 8).Merge
    target.Cells(2, 9).Resize(1, 15).Merge
    target.Cells(2, 24).Resize(1, 6).Merge
    target.Cells(2, 30).Resize(1, 10).Merge
    target.Cells(2, 40).Resize(1, 9).Merge
    target.Cells(3, 4).Resize(1, TORSION_POST_DISPLAY_COLUMNS - 3).Font.Bold = False
    target.Cells(3, 4).Resize(1, TORSION_POST_DISPLAY_COLUMNS - 3).HorizontalAlignment = -4131
    target.Rows(3).WrapText = False: target.Rows(4).WrapText = False
    target.Cells(5, 1).Resize(rows, 3).NumberFormat = "@"
    target.Cells(5, 2).Resize(rows, 1).HorizontalAlignment = -4131
    target.Cells(5, 19).Resize(rows, 1).HorizontalAlignment = -4131
    target.Cells(5, 41).Resize(rows, 1).HorizontalAlignment = -4131
    Dim column As Variant, row As Long
    For Each column In Array(11, 16, 17, 24, 32): target.Cells(5, column).Resize(rows, 1).NumberFormat = "0.000000": Next column
    target.Cells(5, 14).Resize(rows, 1).NumberFormat = "0.00E+00" ' Три значащие цифры сохраняют порядок малых потоков и помещаются в заданную ширину.
    For row = 5 To rows + 4
        For Each column In Array(3, 23, 29, 39, 48)
            If Len(CStr(target.Cells(row, column).Value2)) > 0 Then ApplyStatusFill target.Cells(row, column), CStr(target.Cells(row, column).Value2)
        Next column
    Next row
End Sub

' Нижняя таблица оформляется как технические матрицы: без рамок и без
' запаса пустых строк. Якорь всегда остаётся одной ячейкой шапки.
Private Sub WriteTorsionBarForces(ByVal anchor As Object, ByRef bars As Variant, ByVal count As Long)
    Dim last As Long: last = anchor.Worksheet.Cells(anchor.Worksheet.Rows.Count, anchor.Column).End(-4162).Row
    If last < anchor.Row + count - 1 Then last = anchor.Row + count - 1
    With anchor.Resize(last - anchor.Row + 1, TORSION_BAR_COLUMNS)
        .ClearContents: .Borders.LineStyle = -4142: .Font.Name = "Arial": .Font.Size = 9: .Font.Bold = False
        .Font.Color = 0: .Interior.Pattern = -4142: .HorizontalAlignment = -4108: .VerticalAlignment = -4108: .WrapText = False
        .NumberFormat = "General"
    End With
    Dim values() As Variant, row As Long, column As Long: ReDim values(1 To count, 1 To TORSION_BAR_COLUMNS)
    For row = 1 To count: For column = 1 To TORSION_BAR_COLUMNS: values(row, column) = bars(row, column): Next column: Next row
    anchor.Resize(count, TORSION_BAR_COLUMNS).Value2 = values
    With anchor.Resize(1, TORSION_BAR_COLUMNS)
        .Font.Bold = True: .Interior.Color = RGB(217, 217, 217)
        .HorizontalAlignment = -4131
        .Borders.LineStyle = 1: .Borders.Weight = 2: .Borders.Color = RGB(200, 200, 200)
    End With
    With anchor.Offset(-1, 0).Resize(1, TORSION_BAR_COLUMNS)
        .UnMerge: .Merge: .Value2 = "Усилия от T (кручение)"
        .Font.Name = "Arial": .Font.Size = 9: .Font.Bold = True: .Font.Color = 0
        .Interior.Color = RGB(217, 217, 217): .Borders.LineStyle = -4142
        .Borders(9).LineStyle = 1: .Borders(9).Weight = 2: .Borders(9).Color = RGB(200, 200, 200)
        .HorizontalAlignment = -4108: .WrapText = False
    End With
    anchor.Offset(-2, -2).Resize(1, TORSION_BAR_COLUMNS + 2).Interior.Color = RGB(189, 215, 238)
    If count > 1 Then anchor.Offset(1, 3).Resize(count - 1, 1).NumberFormat = "0.000000"
    If count > 1 Then anchor.Offset(1, 2).Resize(count - 1, 1).HorizontalAlignment = -4131
    If count > 1 Then anchor.Offset(1, 4).Resize(count - 1, 1).NumberFormat = "0.000"
End Sub

' Строит шапку по единицам сохранённого снимка, не по текущему Config.
Private Function TorsionHeadersForSource(ByVal source As Object) As Variant
    Dim saved As Variant: saved = source.Table(TORSION_POST_HEADERS_KEY)
    If UBound(saved, 2) = TORSION_POST_DISPLAY_COLUMNS Then TorsionHeadersForSource = saved: Exit Function
    If UBound(saved, 2) = 43 Then
        Dim logical(1 To 4, 1 To TORSION_POST_COLUMNS) As Variant, oldMap As Variant, c As Long, r As Long
        oldMap = Array(1, 39, 38, 0, 0, 0, 3, 0, 2, 37, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, _
            16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36)
        For c = 9 To 43
            For r = 3 To 4: logical(r, oldMap(c - 1)) = saved(r, c): Next r
        Next c
        TorsionHeadersForSource = ReorderedTorsionHeaders(logical, HeaderUnit(saved(4, 12)), HeaderUnit(saved(4, 11)), _
            HeaderUnit(saved(4, 4)), HeaderUnit(saved(4, 7)), HeaderUnit(saved(4, 19)))
        Exit Function
    End If
    Dim props As Variant: props = source.Table("rngNDMSectionProperties")
    Dim i As Long, key As String, lengthUnit As String, areaUnit As String, forceUnit As String, momentUnit As String, stressUnit As String
    For i = 2 To UBound(props, 1)
        key = CStr(props(i, 3))
        Select Case key
            Case "Output.LengthUnit": lengthUnit = CStr(props(i, 4))
            Case "Output.AreaUnit": areaUnit = CStr(props(i, 4))
            Case "Output.ForceUnit": forceUnit = CStr(props(i, 4))
            Case "Output.MomentUnit": momentUnit = CStr(props(i, 4))
            Case "Output.StressUnit": stressUnit = CStr(props(i, 4))
        End Select
    Next i
    TorsionHeadersForSource = ReorderedTorsionHeaders(saved, lengthUnit, areaUnit, forceUnit, momentUnit, stressUnit)
End Function

' Использует логические подписи вычислителя, но размещает их в порядке
' самостоятельной таблицы; нагрузки и ResultComment находятся в её начале.
Public Function SummaryHeaders(ByVal units As CUnitSystem) As Variant
    Dim logical As Variant: logical = LogicalTorsionHeaders(units)
    SummaryHeaders = ReorderedTorsionHeaders(logical, units.OutputLengthUnit, units.OutputAreaUnit, units.OutputForceUnit, units.OutputMomentUnit, units.OutputStressUnit)
End Function

' Размещает логические подписи в самостоятельной подробной таблице вывода.
Private Function ReorderedTorsionHeaders(ByRef logical As Variant, ByVal lengthUnit As String, ByVal areaUnit As String, _
        ByVal forceUnit As String, ByVal momentUnit As String, ByVal stressUnit As String) As Variant
    Dim output(1 To 4, 1 To TORSION_POST_DISPLAY_COLUMNS) As Variant, map As Variant, column As Long
    map = DisplayColumns()
    For column = 9 To TORSION_POST_DISPLAY_COLUMNS
        If map(column - 1) <= UBound(logical, 2) Then
            output(3, column) = logical(3, map(column - 1))
            output(4, column) = logical(4, map(column - 1))
        End If
    Next column
    output(1, 1) = "Combination ID": output(1, 2) = "ResultComment": output(1, 3) = "общий статус"
    output(1, 4) = "нагрузки в точке приложения"
    output(1, 9) = "Приближённая постпроверка T (кручение) и продольной арматуры NMT"
    output(2, 9) = "T — продольная арматура": output(2, 24) = "T — хомут": output(2, 30) = "T — бетон"
    output(2, 40) = "NMT — сумма напряжений продольной арматуры"
    output(3, 4) = "Исходная продольная сила": output(3, 5) = "Исходный изгибающий момент Mx"
    output(3, 6) = "Исходный изгибающий момент My": output(3, 7) = "Заданный крутящий момент"
    output(4, 4) = "N, " & forceUnit: output(4, 5) = "Mx, " & momentUnit
    output(4, 6) = "My, " & momentUnit: output(4, 7) = "T, " & momentUnit
    output(4, 18) = "Rebar": output(4, 40) = "Rebar"
    output(3, 32) = "Площадь внутри средней линии бетонной стенки": output(4, 32) = "A_k, " & areaUnit
    output(3, 33) = "Расчётная цилиндрическая прочность бетона": output(4, 33) = "f_cd, " & stressUnit
    output(3, 34) = "Коэффициент снижения прочности раскоса": output(4, 34) = ChrW$(&H3BD)
    output(3, 19) = "Сохранённая принадлежность стержня к грани": output(4, 19) = "Грань"
    output(3, 41) = "Сохранённая принадлежность стержня к грани": output(4, 41) = "Грань"
    ReorderedTorsionHeaders = output
End Function

' Дополняет только шесть столбцов общей сводки: три статуса и три запаса.
' Итоговые статусы и причины НДМ не изменяются; все операции идут в массивах.
Public Sub ApplyBatchSummary(ByRef summary() As Variant, ByVal source As Object)
    PutBatchHeaders summary
    If Not source.HasTable(TORSION_POST_DATA_KEY) Then Exit Sub
    Dim data As Variant: data = source.Table(TORSION_POST_DATA_KEY)
    Dim mapStatus As Variant, mapReserve As Variant, i As Long, row As Long, part As Long
    mapStatus = Array(28, 36, 21): mapReserve = Array(27, 35, 20)
    For row = 12 To UBound(summary, 1)
        For part = 0 To 2
            summary(row, TORSION_BATCH_STATUS_COLUMN + part) = Empty
            summary(row, TORSION_BATCH_RESERVE_COLUMN + part) = Empty
        Next part
    Next row
    Dim minimum(0 To 2) As Double, worst(0 To 2) As String
    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, 1))) > 0 Then
            For row = 13 To UBound(summary, 1)
                If StrComp(CStr(summary(row, 1)), CStr(data(i, 1)), vbTextCompare) = 0 Then
                    For part = 0 To 2
                        summary(row, TORSION_BATCH_STATUS_COLUMN + part) = data(i, mapStatus(part))
                        summary(row, TORSION_BATCH_RESERVE_COLUMN + part) = data(i, mapReserve(part))
                        If IsNumeric(data(i, mapReserve(part))) And Not IsEmpty(data(i, mapReserve(part))) Then
                            If CDbl(data(i, mapReserve(part))) > 0# Then
                                If minimum(part) = 0# Or CDbl(data(i, mapReserve(part))) < minimum(part) Then
                                    minimum(part) = CDbl(data(i, mapReserve(part))): worst(part) = CStr(data(i, 1))
                                End If
                            End If
                        End If
                    Next part
                    Exit For
                End If
            Next row
        End If
    Next i
    For part = 0 To 2
        If minimum(part) > 0# Then
            summary(12, TORSION_BATCH_STATUS_COLUMN + part) = worst(part)
            summary(12, TORSION_BATCH_RESERVE_COLUMN + part) = minimum(part)
        End If
    Next part
End Sub

' Описывает три независимые проверки в принятых местах общей сводки.
Public Sub PutBatchHeaders(ByRef data() As Variant)
    Dim part As Long, captions As Variant, column As Long
    captions = Array("раскос бетона", "продольная арматура", "хомут")
    data(7, TORSION_BATCH_STATUS_COLUMN) = "кручение NMT"
    data(7, TORSION_BATCH_RESERVE_COLUMN) = "кручение NMT"
    For part = 0 To 2
        column = TORSION_BATCH_STATUS_COLUMN + part
        data(8, column) = captions(part): data(11, column) = "Статус самостоятельной проверки: " & captions(part)
        column = TORSION_BATCH_RESERVE_COLUMN + part
        data(8, column) = captions(part): data(11, column) = "Минимальный запас: " & captions(part)
    Next part
End Sub

' Даёт четыре яруса шапки: название, группы, русские пояснения и короткие
' обозначения с единицами. Размеры диапазона и порядок данных сохраняются.
Private Function LogicalTorsionHeaders(ByVal units As CUnitSystem) As Variant
    Dim result(1 To 4, 1 To TORSION_POST_COLUMNS) As Variant, labels As Variant, notes As Variant, column As Long
    Dim sigma As String, deltaSigma As String, unit As String
    sigma = ChrW$(&H3C3): deltaSigma = ChrW$(&H394) & sigma ' Греческие знаки не зависят от ANSI-кодировки сохранённого проекта VBA.
    labels = Array("ID", "State", "T", "A_o", "u_o", "cot " & ChrW$(&H3B8), "q", "F_T", "A_s", _
        "A_s,req", "ID", deltaSigma & "_T,max", "R_s", "k", "статус", "A_sw", "s", sigma & "_sw", "R_sw", "k", "статус", _
        "метод", "t_eff", sigma & "_diag", "R_diag", "T_lim", "k", "статус", _
        "ID", sigma & "_NM", deltaSigma & "_T", "F_NMT", sigma & "_NMT", "R_s", "k", "статус", "b_sel", "статус", "ResultComment", "A_k", "f_cd", ChrW$(&H3BD), "Грань", "Грань")
    notes = Array("Номер сочетания", "Сохранённое состояние НДМ", "Заданный крутящий момент", _
        "Площадь внутри оси наружного хомута", "Периметр оси наружного хомута", "Котангенс угла бетонных раскосов", _
        "Касательный поток", "Суммарная продольная тяга", "Площадь отобранной продольной арматуры", _
        "Требуемая площадь продольной арматуры", "Стержень с наибольшей добавкой напряжения", "Максимальная растягивающая добавка от T", _
        "Сопротивление продольной арматуры", "Запас продольной арматуры от T", "Статус проверки продольной арматуры", _
        "Площадь одной ветви хомута", "Шаг замкнутых хомутов", "Напряжение ветви хомута", "Сопротивление ветви хомута", _
        "Запас поперечной арматуры", "Статус проверки хомута", "Метод проверки бетона", "Эффективная толщина бетонной полосы", _
        "Напряжение бетонного раскоса", "Допустимое напряжение бетонного раскоса", "Предельный крутящий момент по бетону", _
        "Запас бетона", "Статус проверки бетона", "Определяющий стержень по сумме NMT", "Сохранённое напряжение НДМ", _
        "Растягивающая добавка от T", "Суммарное усилие стержня NMT", "Алгебраическая сумма напряжений NMT", _
        "Сопротивление продольной арматуры", "Запас по сумме напряжений NMT", "Статус проверки суммы NMT", _
        "Ширина полосы отбора продольных стержней", "Итоговый статус постпроверки", "Причины отказов и предупреждения T и NMT", _
        "Площадь внутри средней линии бетонной стенки", "Расчётная цилиндрическая прочность бетона", "Коэффициент снижения прочности раскоса", "Сохранённая принадлежность стержня к грани", "Сохранённая принадлежность стержня к грани")
    result(1, 1) = "Приближённая постпроверка T (кручение) и продольной арматуры NMT"
    For column = 1 To TORSION_POST_COLUMNS
        result(3, column) = notes(column - 1)
        Select Case column
            Case 1: result(2, column) = "Сочетание"
            Case 3: result(2, column) = "T — продольная арматура"
            Case 16: result(2, column) = "T — хомут"
            Case 22: result(2, column) = "T — бетон"
            Case 29: result(2, column) = "NMT — сумма напряжений продольной арматуры"
            Case 37: result(2, column) = "Отбор"
            Case 38: result(2, column) = "Постпроверка"
        End Select
        unit = vbNullString
        Select Case column
            Case 3, 26: unit = units.OutputMomentUnit
            Case 4, 9, 10, 16, 40: unit = units.OutputAreaUnit
            Case 5, 17, 23, 37: unit = units.OutputLengthUnit
            Case 7: unit = units.OutputForceUnit & "/" & units.OutputLengthUnit
            Case 8, 32: unit = units.OutputForceUnit
            Case 12, 13, 18, 19, 24, 25, 30, 31, 33, 34, 41: unit = units.OutputStressUnit
        End Select
        result(4, column) = labels(column - 1)
        If Len(unit) > 0 Then result(4, column) = result(4, column) & ", " & unit
    Next column
    LogicalTorsionHeaders = result
End Function
