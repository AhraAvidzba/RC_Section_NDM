Attribute VB_Name = "modTestPlotConfig"
Option Explicit

' ДЛЯ ТЕСТОВ: настройки схемы Excel и расположение размеров и подписей.
' Использует отдельные временные книги, сохранённые данные и независимо
' вычисленные координаты. Проверяет включение элементов, ошибки ввода,
' сохранность диаграммы и Results, повторную работу после ошибки. НДС не решает.

Private Type TUiTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

' ==================== ДЛЯ ТЕСТОВ: ОБЩИЕ НАСТРОЙКИ EXCEL-СХЕМЫ ====================

' Проверяет обычное построение схемы по заранее заданным сохранённым данным.
' Читает реальные фигуры диаграммы независимо от решателя; исходная книга
' и пользовательский экземпляр Excel в тесте не используются.
Public Function RunAudit03GeneralPlotTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TUiTestStats
    TestAudit03GeneralPlotContracts stats
    TestAudit03AnnotationLayoutContracts stats
    TestAudit03PlotEnableContracts stats
    TestAudit03SnapshotMetadataContracts stats
    TestAudit03AutoPlotContracts stats
    TestAudit03ContourArcContracts stats
    TestCrackAnnotationsDoNotRequirePlotNormals stats
    TestAudit03NamedAnnotationSettings stats
    AppendLine stats, "TOTAL_AUDIT03_GENERAL_PLOT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    passed = stats.Passed: failed = stats.Failed
    RunAudit03GeneralPlotTests = stats.Report
End Function

' Проверяет отдельно разделение схемы подписей и расчетных аннотаций трещин.
' Временная книга содержит пустые нормали области и координаты неуспешной
' пробы; для настоящей размерной подписи обязательность нормали сохраняется.
Public Function RunCrackAnnotationPlotTests() As String
    Dim stats As TUiTestStats
    TestCrackAnnotationsDoNotRequirePlotNormals stats
    AppendLine stats, "TOTAL_CRACK_ANNOTATION_PLOT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunCrackAnnotationPlotTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ: переносит named anchor и проверяет геометрическое preview.
' Расчетные аннотации остаются неизменными, solver не вызывается.
Private Sub TestCrackAnnotationsDoNotRequirePlotNormals(ByRef stats As TUiTestStats)
    Dim fixture As Object, config As Object, sheet As Object, source As Object, target As Object, anchor As Object
    Dim settings As CSystemSettingsReader, reader As CSectionPlotDataReader
    Dim addresses As Variant, name As Variant, headers As Variant, geometry As Variant, saved As Variant, actual As Variant
    Dim annotations(1 To 6, 1 To 9) As Variant, props(1 To 2, 1 To 4) As Variant
    Dim i As Long, j As Long, position As Long, code As Long, reason As String, solveCount As Long, prefix As String
    On Error GoTo Failed
    solveCount = SectionEquilibriumSolveCount()
    Set fixture = Application.Workbooks.Add(-4167): Set config = fixture.Worksheets(1): config.Name = "Config"
    Set sheet = fixture.Worksheets.Add: sheet.Name = "Results"
    addresses = Array("A5", "H5", "P5", "P25"): i = 0
    For Each name In Array("rngSystemSettings", "rngPlotAnnotationSettings", "rngUnitSettings", "rngSignConventionSettings")
        Set source = ThisWorkbook.Names.Item(CStr(name)).RefersToRange
        Set target = config.Range(CStr(addresses(i))).Resize(source.Rows.Count, source.Columns.Count)
        target.Value2 = source.Value2
        fixture.Names.Add Name:=CStr(name), RefersTo:="=Config!" & target.Address: i = i + 1
    Next name
    geometry = Audit03ContourGeometrySnapshot("CONTOUR_LINE")
    sheet.Range("A5").Resize(3, 15).Value2 = geometry
    fixture.Names.Add Name:="rngNDMSectionGeometry", RefersTo:="=Results!$A$5"
    Audit03PrepareContourFixture fixture, "CONTOUR_LINE"
    props(1, 1) = "LoadCase": props(1, 2) = "Parameter": props(1, 3) = "Value": props(1, 4) = "Unit"
    props(2, 1) = "ALL": props(2, 2) = "Output.LengthUnit": props(2, 3) = "mm": props(2, 4) = "-"
    sheet.Range("AB5").Resize(2, 4).Value2 = props
    fixture.Names.Add Name:="rngNDMSectionProperties", RefersTo:="=Results!$AB$5"
    headers = Array("AnnotationType", "StartX", "StartY", "EndX", "EndY", "OutsideNormalX", "OutsideNormalY", "Text", "Unit")
    For i = 0 To UBound(headers): annotations(1, i + 1) = headers(i): Next i
    For i = 2 To 3
        For j = 2 To 7: annotations(i, j) = 0#: Next j
        annotations(i, 9) = "mm"
    Next i
    annotations(2, 1) = "DIMENSION": annotations(2, 4) = 100#: annotations(2, 7) = 1#: annotations(2, 8) = "DIM"
    annotations(3, 1) = "REBAR_ANNOTATION": annotations(3, 4) = 100#: annotations(3, 7) = 1#: annotations(3, 8) = "REBAR"
    annotations(4, 1) = "CRACK_REGION_LINE": annotations(4, 2) = 10#: annotations(4, 3) = 20#
    annotations(4, 4) = 30#: annotations(4, 5) = 40#: annotations(4, 9) = "mm"
    annotations(5, 1) = "CRACK_REGION_ARC": annotations(5, 9) = "mm"
    annotations(6, 1) = "CRACK_PROBE": annotations(6, 8) = "P1/Q1": annotations(6, 9) = "mm"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook fixture
    Set reader = New CSectionPlotDataReader
    For position = 0 To 1
        If position = 0 Then Set anchor = sheet.Range("A20") Else Set anchor = sheet.Range("CH800")
        anchor.Resize(6, 9).Value2 = annotations
        fixture.Names.Add Name:="rngNDMSectionAnnotations", RefersTo:="=Results!" & anchor.Address
        saved = anchor.Resize(6, 9).Value2
            prefix = "crackAnnotation.p" & CStr(position)
            reader.LoadGeometryPreviewFromWorkbook fixture, settings
            AssertTrue stats, prefix & ".count", reader.AnnotationCount = 3
            AssertTrue stats, prefix & ".dimension", reader.AnnotationType(2) = "DIMENSION"
            AssertTrue stats, prefix & ".rebar", reader.AnnotationType(3) = "REBAR_ANNOTATION"
            actual = anchor.Resize(6, 9).Value2
            For i = 1 To 6
                For j = 1 To 9
                    AssertTrue stats, prefix & ".preserved." & CStr(i) & "." & CStr(j), CStr(actual(i, j)) = CStr(saved(i, j))
                Next j
            Next i
            anchor.Cells(2, 6).ClearContents
            On Error Resume Next
            Err.Clear: reader.LoadGeometryPreviewFromWorkbook fixture, settings
            code = Err.Number: reason = Err.Description
            On Error GoTo Failed
            AssertTrue stats, prefix & ".realNormalRejected", code <> 0
            AssertTrue stats, prefix & ".dynamicAddress", InStr(1, reason, "Results!" & anchor.Cells(2, 6).Address(False, False), vbTextCompare) > 0
            AssertTrue stats, prefix & ".reset", reader.AnnotationCount = 0 And reader.Count = 0
            anchor.Cells(2, 6).Value2 = 0#
            reader.LoadGeometryPreviewFromWorkbook fixture, settings
            AssertTrue stats, prefix & ".recovery", reader.AnnotationCount = 3
    Next position
    AssertTrue stats, "crackAnnotation.noSolve", SectionEquilibriumSolveCount() = solveCount
    fixture.Close False
    Exit Sub
Failed:
    AssertTrue stats, "crackAnnotation.runtime." & CStr(Err.Number) & "." & Err.Description, False
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
End Sub

' Проверяет углы дуг сохраненного контура через reader и настоящий preview.
' Рисунок и Results принадлежат отдельной книге; ошибочный текст должен
' отклоняться до очистки прежней схемы, а не превращаться в нулевую дугу.
Public Function RunAudit03ContourArcTests() As String
    Dim stats As TUiTestStats
    TestAudit03ContourArcContracts stats
    AppendLine stats, "TOTAL_AUDIT03_CONTOUR_ARC: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03ContourArcTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ: реальная толщина прямых, дуг и окружностей Outer/Opening;
' перенос таблицы, масштаб Chart и ошибки ввода не должны менять этот контракт.
Public Function RunContourLineWeightTests() As String
    Dim stats As TUiTestStats, fixture As Object, config As Object, sheet As Object
    Dim table As Object, source As Object, target As Object, cell As Object, chart As Object, shape As Object, marker As Object
    Dim settings As CSystemSettingsReader, reader As CSectionPlotDataReader, plotter As CSectionPlotter
    Dim addresses As Variant, name As Variant, kind As Variant, weight As Variant, bad As Variant
    Dim contours(1 To 7, 1 To 17) As Variant, segment As Variant, props(1 To 2, 1 To 4) As Variant, baseline As Variant
    Dim i As Long, j As Long, position As Long, role As Long, row As Long, count As Long, size As Long, code As Long, solves As Long
    Dim reason As String, prefix As String
    On Error GoTo Failed
    solves = SectionEquilibriumSolveCount()
    Set fixture = Application.Workbooks.Add(-4167): Set config = fixture.Worksheets(1): config.Name = "Config"
    Set sheet = fixture.Worksheets.Add: sheet.Name = "Results": fixture.Worksheets.Add.Name = "Расчет"
    addresses = Array("A5", "H5", "P5", "P25"): i = 0
    For Each name In Array("rngSystemSettings", "rngPlotAnnotationSettings", "rngUnitSettings", "rngSignConventionSettings")
        Set source = ThisWorkbook.Names.Item(CStr(name)).RefersToRange
        Set target = config.Range(CStr(addresses(i))).Resize(source.Rows.Count, source.Columns.Count)
        target.NumberFormat = "@": target.Value2 = source.Value2
        fixture.Names.Add Name:=CStr(name), RefersTo:="=Config!" & target.Address: i = i + 1
    Next name
    Set table = fixture.Names.Item("rngSystemSettings").RefersToRange
    Audit03SetPlotSetting table, "Plot.ContourEnabled", "Yes"
    Set table = fixture.Names.Item("rngPlotAnnotationSettings").RefersToRange
    Audit03NamedAnnotationCell(table, "Plot.Dimensions.Enabled").Value2 = "No"
    Audit03NamedAnnotationCell(table, "Plot.RebarLabels.Enabled").Value2 = "No"
    baseline = table.Value2
    sheet.Range("A5").Resize(3, 15).Value2 = Audit03ContourGeometrySnapshot("CONTOUR_LINE")
    fixture.Names.Add Name:="rngNDMSectionGeometry", RefersTo:="=Results!$A$5"
    props(1, 1) = "LoadCase": props(1, 2) = "Parameter": props(1, 3) = "Value": props(1, 4) = "Unit"
    props(2, 1) = "ALL": props(2, 2) = "Output.LengthUnit": props(2, 3) = "mm": props(2, 4) = "-"
    sheet.Range("R5").Resize(2, 4).Value2 = props
    fixture.Names.Add Name:="rngNDMSectionProperties", RefersTo:="=Results!$R$5"
    row = 1
    For role = 0 To 1
        For Each kind In Array("CONTOUR_LINE", "CONTOUR_ARC", "CONTOUR_CIRCLE")
            segment = Audit03ContourSnapshot(CStr(kind))
            For j = 1 To 17: contours(1, j) = segment(1, j): contours(row + 1, j) = segment(2, j): Next j
            contours(row + 1, 2) = "LOOP_" & CStr(row): contours(row + 1, 3) = "SEG_" & CStr(row)
            contours(row + 1, 4) = row
            If role = 1 Then contours(row + 1, 5) = "Opening"
            row = row + 1
        Next kind
    Next role
    sheet.Range("EZ5").Resize(7, 17).Value2 = contours
    fixture.Names.Add Name:="rngNDMSectionContours", RefersTo:="=Results!$EZ$5"
    Set chart = fixture.Worksheets("Расчет").ChartObjects.Add(10#, 10#, 640#, 430#): chart.Name = "chtNDMSectionPlot"
    Set settings = New CSystemSettingsReader: Set reader = New CSectionPlotDataReader: Set plotter = New CSectionPlotter
    For position = 0 To 1
        If position = 0 Then Set table = config.Range("H5").Resize(UBound(baseline, 1), 6) Else Set table = config.Range("CH800").Resize(UBound(baseline, 1), 6)
        table.NumberFormat = "@": table.Value2 = baseline
        fixture.Names.Item("rngPlotAnnotationSettings").RefersTo = "=Config!" & table.Address
        For i = 2 To table.Rows.Count
            If CStr(table.Cells(i, 1).Value2) = "ContourLineWeight" Then Set cell = table.Cells(i, 2)
        Next i
        settings.LoadFromWorkbook fixture
        AssertClose stats, "contourWeight.default.p" & CStr(position), settings.GetRequiredDouble("Plot.ContourLineWeight"), 2#, 0#
        AssertTrue stats, "contourWeight.singleKey.p" & CStr(position), Not settings.HasKey("Plot.RebarLabels.ContourLineWeight") And Not settings.HasKey("Plot.Dimensions.ContourLineWeight")
        reader.LoadGeometryPreviewFromWorkbook fixture, settings
        For Each weight In Array(2#, 0.7, 3.25)
            cell.Value2 = weight: settings.LoadFromWorkbook fixture
            For size = 0 To 1
                chart.Width = 640# + size * 260#: chart.Height = 430# + size * 170#
                plotter.Draw fixture, reader, settings: count = 0
                For Each shape In chart.Chart.Shapes
                    If InStr(shape.Name, "NDMPlot_Contour") = 1 Then
                        count = count + 1
                        AssertClose stats, "contourWeight.render." & CStr(position) & "." & CStr(weight) & "." & CStr(size) & "." & shape.Name, shape.Line.Weight, CDbl(weight), 0.000001
                    End If
                Next shape
                AssertTrue stats, "contourWeight.sixBoundaries", count = 6
            Next size
        Next weight
        For Each bad In Array("", "wrong", 0#, -1#, CVErr(2015))
            cell.Value2 = bad: settings.LoadFromWorkbook fixture
            Set marker = chart.Chart.Shapes.AddShape(1, 11#, 17#, 20#, 20#): marker.Name = "NDMPlot_ThicknessSentinel"
            On Error Resume Next
            Err.Clear: plotter.Draw fixture, reader, settings: code = Err.Number: reason = Err.Description
            On Error GoTo Failed
            prefix = "contourWeight.invalid.p" & CStr(position) & ".type" & CStr(VarType(bad))
            AssertTrue stats, prefix & ".rejected", code <> 0 And InStr(reason, "Plot.ContourLineWeight") > 0
            AssertTrue stats, prefix & ".address", InStr(reason, "Config") > 0 And InStr(reason, cell.Address(False, False)) > 0
            AssertTrue stats, prefix & ".preserved", Audit03PlotShapeExists(chart, "ThicknessSentinel")
            marker.Delete
        Next bad
        Set target = fixture.Names.Item("rngSystemSettings").RefersToRange
        Audit03SetPlotSetting target, "Plot.ContourEnabled", "No"
        settings.LoadFromWorkbook fixture: plotter.Draw fixture, reader, settings
        AssertTrue stats, "contourWeight.disabled.ignoresError.p" & CStr(position), Not Audit03PlotShapeExists(chart, "Contour")
        Audit03SetPlotSetting target, "Plot.ContourEnabled", "Yes"
        cell.Offset(0, -1).Value2 = "MissingContourWeight": settings.LoadFromWorkbook fixture
        On Error Resume Next
        Err.Clear: plotter.Draw fixture, reader, settings: code = Err.Number: reason = Err.Description
        On Error GoTo Failed
        AssertTrue stats, "contourWeight.missing.p" & CStr(position), code <> 0 And InStr(reason, "Plot.ContourLineWeight") > 0
        cell.Offset(0, -1).Value2 = "ContourLineWeight"
    Next position
    AssertTrue stats, "contourWeight.noSolve", SectionEquilibriumSolveCount() = solves
    GoTo Cleanup
Failed:
    AssertTrue stats, "contourWeight.runtime." & CStr(Err.Number) & "." & Err.Description, False
Cleanup:
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
    AppendLine stats, "TOTAL_CONTOUR_LINE_WEIGHT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunContourLineWeightTests = stats.Report
End Function

' Испытывает полное чтение Text, signed sweep и recovery после повреждения.
' Для quarter-circle число хорд задано независимо: шаг схемы 5 градусов,
' поэтому четверть окружности содержит 18 сегментов. НДС не решается.
' ДЛЯ ТЕСТОВ: создает новый geometry snapshot с двумя элементами и отдельной
' точной границей. Контур не пишется в annotation-таблицу даже в fixture.
Private Function Audit03ContourGeometrySnapshot(ByVal kind As String) As Variant
    Dim data(1 To 3, 1 To 15) As Variant, original As Variant, i As Long, j As Long
    original = Audit03GeometrySnapshotArray("mm", 1#)
    For i = 1 To 3
        data(i, 2) = original(i, 1): data(i, 3) = original(i, 2)
        data(i, 4) = original(i, 3): data(i, 5) = original(i, 4): data(i, 6) = original(i, 5)
        For j = 7 To 15: data(i, j) = original(i, j): Next j
        data(i, 1) = "TEST_RUN"
    Next i
    data(1, 1) = "RunID"
    Audit03ContourGeometrySnapshot = data
End Function

' ДЛЯ ТЕСТОВ: отдельный компактный contour-снимок, включая штатно пустой.
Private Function Audit03ContourSnapshot(ByVal kind As String) As Variant
    Dim data() As Variant, headers As Variant, j As Long, rows As Long
    rows = 1: If Len(kind) > 0 Then rows = 2
    ReDim data(1 To rows, 1 To 17)
    headers = Array("RunID", "LoopID", "SegmentID", "Sequence", "LoopRole", "SegmentType", _
        "StartX", "StartY", "EndX", "EndY", "CenterX", "CenterY", "Radius", "SweepAngle, rad", _
        "SourceID", "Comment", "LengthUnit (SectionXY)")
    For j = 0 To UBound(headers): data(1, j + 1) = headers(j): Next j
    If rows = 2 Then
        data(2, 1) = "TEST_RUN": data(2, 2) = "OUTER": data(2, 3) = "R1": data(2, 4) = 1
        data(2, 5) = "Outer": data(2, 6) = kind: data(2, 7) = 0#: data(2, 8) = 0#
        data(2, 15) = "TEST": data(2, 16) = "Точный тестовый контур.": data(2, 17) = "mm"
        Select Case kind
            Case "CONTOUR_LINE": data(2, 9) = 100#: data(2, 10) = 0#
            Case "CONTOUR_CIRCLE": data(2, 13) = 100#
            Case "CONTOUR_ARC"
                data(2, 7) = 100#: data(2, 11) = 0#: data(2, 12) = 0#: data(2, 14) = "1.570796326795"
        End Select
    End If
    Audit03ContourSnapshot = data
End Function

' ДЛЯ ТЕСТОВ: якорь вне остальных синтетических таблиц и их перенесенных копий.
Private Sub Audit03PrepareContourFixture(ByVal fixture As Object, Optional ByVal kind As String = "")
    Dim data As Variant
    data = Audit03ContourSnapshot(kind)
    fixture.Worksheets.Item("Results").Range("EZ5").Resize(UBound(data, 1), 17).Value2 = data
    fixture.Names.Add Name:="rngNDMSectionContours", RefersTo:="=Results!$EZ$5"
End Sub

' ДЛЯ ТЕСТОВ: проверяет углы, динамические адреса и сохранность новой geometry-таблицы.
Private Sub TestAudit03ContourArcContracts(ByRef stats As TUiTestStats)
    Dim fixture As Object, config As Object, sheet As Object, source As Object, target As Object, anchor As Object
    Dim settings As CSystemSettingsReader, reader As CSectionPlotDataReader, plotter As CSectionPlotter
    Dim addresses As Variant, name As Variant, bad As Variant, valid As Variant, shape As Object
    Dim index As Long, position As Long, variantIndex As Long, code As Long, reason As String, prefix As String, exportText As String
    Dim chart As Object, marker As Object, cell As Object, count As Long, cases As Long, solveCount As Long
    Dim annotations As Variant, props(1 To 2, 1 To 4) As Variant, keys As Variant
    Dim baseline As Variant, actual As Variant, expected As Variant, separator As Variant
    Dim savedSeparators As Boolean, savedDecimal As String, savedThousands As String
    On Error GoTo Failed
    solveCount = SectionEquilibriumSolveCount()
    savedSeparators = Application.UseSystemSeparators: savedDecimal = Application.DecimalSeparator: savedThousands = Application.ThousandsSeparator
    Set fixture = Application.Workbooks.Add(-4167): Set config = fixture.Worksheets(1): config.Name = "Config"
    Set sheet = fixture.Worksheets.Add: sheet.Name = "Results": fixture.Worksheets.Add.Name = "Расчет"
    addresses = Array("A5", "H5", "P5", "P25"): index = 0
    For Each name In Array("rngSystemSettings", "rngPlotAnnotationSettings", "rngUnitSettings", "rngSignConventionSettings")
        Set source = ThisWorkbook.Names.Item(CStr(name)).RefersToRange
        Set target = config.Range(CStr(addresses(index))).Resize(source.Rows.Count, source.Columns.Count)
        target.NumberFormat = "@": target.Value2 = source.Value2
        fixture.Names.Add Name:=CStr(name), RefersTo:="=Config!" & target.Address: index = index + 1
    Next name
    Audit03SetPlotSetting fixture.Names.Item("rngSystemSettings").RefersToRange, "Plot.Enabled", "Yes"
    Audit03SetPlotSetting fixture.Names.Item("rngSystemSettings").RefersToRange, "Plot.ContourEnabled", "Yes"
    annotations = Audit03ContourSnapshot("CONTOUR_ARC")
    sheet.Range("A5").Resize(3, 15).Value2 = Audit03ContourGeometrySnapshot("CONTOUR_ARC")
    fixture.Names.Add Name:="rngNDMSectionGeometry", RefersTo:="=Results!$A$5"
    Audit03PrepareContourFixture fixture, "CONTOUR_ARC"
    props(1, 1) = "LoadCase": props(1, 2) = "Parameter": props(1, 3) = "Value": props(1, 4) = "Unit"
    props(2, 1) = "ALL": props(2, 2) = "Output.LengthUnit": props(2, 3) = "mm": props(2, 4) = "-"
    sheet.Range("AB5").Resize(2, 4).Value2 = props
    fixture.Names.Add Name:="rngNDMSectionProperties", RefersTo:="=Results!$AB$5"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook fixture
    Set reader = New CSectionPlotDataReader: Set plotter = New CSectionPlotter
    Set chart = fixture.Worksheets.Item("Расчет").ChartObjects.Add(41#, 53#, 700#, 480#): chart.Name = "chtNDMSectionPlot"
    For position = 0 To 1
        If position = 0 Then Set anchor = sheet.Range("EZ5") Else Set anchor = sheet.Range("CH800")
        anchor.Resize(2, 17).NumberFormat = "@": anchor.Resize(2, 17).Value2 = annotations
        fixture.Names.Item("rngNDMSectionContours").RefersTo = "=Results!" & anchor.Address
        Set cell = anchor.Cells(2, 14): variantIndex = 0
        For Each bad In Array("", "TODO", "1.25garbage", "1,25garbage", "1.2.3", "1,2,3", "1E", "--1", "&H1", CVErr(2015), "1E309", "1E300", "7", "-7", "$1", "1 000", "+.", "1e+1E-2", "True", True)
            variantIndex = variantIndex + 1: cases = cases + 1: cell.Value2 = bad
            prefix = "audit03.contourArc.p" & CStr(position) & ".bad" & CStr(variantIndex)
            Audit03ReadLifecycleLoad reader, fixture, settings, 1, code, reason
            AssertTrue stats, prefix & ".rejected", code <> 0
            AssertTrue stats, prefix & ".address", InStr(1, reason, "Results!" & cell.Address(False, False), vbTextCompare) > 0
            AssertTrue stats, prefix & ".reason", InStr(1, reason, "угол", vbTextCompare) > 0 And InStr(1, reason, "радиан", vbTextCompare) > 0
            AssertTrue stats, prefix & ".empty", reader.Count = 0 And reader.AnnotationCount = 0
            On Error Resume Next
            Err.Clear: exportText = Audit03ReadContourArcSweepsForTests(fixture)
            code = Err.Number: reason = Err.Description
            On Error GoTo Failed
            AssertTrue stats, prefix & ".export", code <> 0 And InStr(1, reason, "Results!" & cell.Address(False, False), vbTextCompare) > 0
            Set marker = chart.Chart.Shapes.AddShape(1, 11#, 17#, 20#, 20#): marker.Name = "NDMPlot_Audit03Sentinel"
            On Error Resume Next
            Err.Clear: UpdateSectionGeometryPreviewForWorkbook fixture
            code = Err.Number: reason = Err.Description
            On Error GoTo Failed
            AssertTrue stats, prefix & ".entrypoint", code <> 0 And InStr(1, reason, cell.Address(False, False), vbTextCompare) > 0
            AssertTrue stats, prefix & ".chart", Audit03PlotShapeExists(chart, "Audit03Sentinel")
            AppendLine stats, "CONTOUR_ARC_ERROR: " & prefix & "|" & reason
            On Error Resume Next
            chart.Chart.Shapes.Item("NDMPlot_Audit03Sentinel").Delete
            On Error GoTo Failed
            cell.Value2 = annotations(2, 14): reader.LoadGeometryPreviewFromWorkbook fixture, settings
            AssertTrue stats, prefix & ".recovery", reader.Count = 2 And reader.AnnotationCount = 1
        Next bad
        expected = Array(1.570796326795, 1.570796326795, 1.570796326795, -1.570796326795, -0.5, 0.5, 0.001, 0#)
        index = 0
        For Each valid In Array("1.570796326795", "1,570796326795", "+1.570796326795E0", "-1.570796326795", "-.5", ".5", "1E-3", "0")
            cases = cases + 1: cell.Value2 = valid: baseline = anchor.Resize(2, 17).Value2
            reader.LoadGeometryPreviewFromWorkbook fixture, settings
            plotter.Draw fixture, reader, settings
            count = 0
            Dim arcShape As Object
            Set arcShape = Nothing
            For Each shape In chart.Chart.Shapes
                If InStr(1, shape.Name, "ContourArcLine", vbTextCompare) > 0 Then
                    count = count + 1: Set arcShape = shape
                End If
            Next shape
            prefix = "audit03.contourArc.valid.p" & CStr(position) & "." & CStr(valid)
            AssertClose stats, prefix & ".angle", reader.AnnotationSweepAngle(1), CDbl(expected(index)), 0#
            AssertTrue stats, prefix & ".export", Audit03ReadContourArcSweepsForTests(fixture) = "1|" & CStr(expected(index)) & vbLf
            AssertTrue stats, prefix & ".oneObject", count = 1
            If InStr(1, CStr(valid), "570796", vbBinaryCompare) > 0 Then
                AssertTrue stats, prefix & ".segments", arcShape.Nodes.Count = 19
            ElseIf CStr(valid) = "0" Then
                AssertTrue stats, prefix & ".line", arcShape.Type = 9
            Else
                AssertTrue stats, prefix & ".segments", arcShape.Nodes.Count = 7
            End If
            If CStr(valid) <> "0" Then AssertTrue stats, prefix & ".openStroke", arcShape.Type = 5 And arcShape.Fill.Visible = 0
            actual = anchor.Resize(2, 17).Value2: Audit03ComparePlainSnapshot stats, prefix, baseline, actual, 1
            index = index + 1
        Next valid
        For Each separator In Array(".", ",")
            Application.UseSystemSeparators = False
            If separator = "." Then Application.ThousandsSeparator = "," Else Application.ThousandsSeparator = "."
            Application.DecimalSeparator = CStr(separator)
            For Each valid In Array("1.570796326795", "1,570796326795")
                cell.Value2 = valid: reader.LoadGeometryPreviewFromWorkbook fixture, settings
                AssertClose stats, "audit03.contourArc.excelLocale.p" & CStr(position) & "." & CStr(separator) & "." & CStr(valid), reader.AnnotationSweepAngle(1), 1.570796326795, 0#
            Next valid
        Next separator
        Application.DecimalSeparator = savedDecimal: Application.ThousandsSeparator = savedThousands: Application.UseSystemSeparators = savedSeparators
    Next position
    AssertTrue stats, "audit03.contourArc.noSolve", SectionEquilibriumSolveCount() = solveCount
    AppendLine stats, "CONTOUR_ARC_CASES: variants=" & CStr(cases) & "; geometryFixtures=1; equilibriumCases=0"
    GoTo Cleanup
Failed:
    AssertTrue stats, "audit03.contourArc.runtime." & CStr(Err.Number) & "." & Err.Description, False
Cleanup:
    On Error Resume Next
    Application.DecimalSeparator = savedDecimal: Application.ThousandsSeparator = savedThousands: Application.UseSystemSeparators = savedSeparators
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
End Sub

' ==================== ДЛЯ ТЕСТОВ: ЯЧЕЙКИ КОМПАКТНОЙ ТАБЛИЦЫ АННОТАЦИЙ ====================

' Проверяет настоящий named-range adapter, а не только готовый словарь.
' Стандартные ячейки двух групп поступают через LoadFromWorkbook в layout
' и реальный Chart; исходная книга и расчетные результаты не изменяются.
Public Function RunAudit03NamedAnnotationTests() As String
    Dim stats As TUiTestStats
    TestAudit03NamedAnnotationSettings stats
    AppendLine stats, "TOTAL_AUDIT03_NAMED_ANNOTATION: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03NamedAnnotationTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ: поворот и рамки наклонных подписей, единицы, стороны и повторный layout.
Public Function RunArbitraryAnnotationLayoutTests() As String
    Dim stats As TUiTestStats, fixture As Object, sheet As Object, table As Object, settings As CSystemSettingsReader
    Dim layout As CPlotAnnotationLayout, degrees As Variant, value As Variant, unit As Variant, side As Variant, frame As Variant
    Dim angle As Double, dx As Double, dy As Double, nx As Double, ny As Double, sx As Double, sy As Double
    Dim px As Double, py As Double, length As Double, cx As Double, cy As Double, dot As Double, gap As Double, prefix As String
    Dim normalX As Double, normalY As Double, sign As Double, qx As Double, qy As Double, item As Variant, counter As Long
    Dim reverse As Variant, savedRotation As Double, savedLeft As Double
    On Error GoTo Failed
    Set fixture = Application.Workbooks.Add(-4167): Set sheet = fixture.Worksheets(1): sheet.Name = "AngleLayout"
    Set table = Audit03PlotSettingsTable(sheet.Range("A5")): Set settings = New CSystemSettingsReader
    For Each value In Array("Dimensions", "RebarLabels")
        Audit03SetPlotSetting table, "Plot." & CStr(value) & ".Enabled", "Yes"
        Audit03SetPlotSetting table, "Plot." & CStr(value) & ".TextHeight", 13#
        Audit03SetPlotSetting table, "Plot." & CStr(value) & ".TextGap", 9#
        Audit03SetPlotSetting table, "Plot." & CStr(value) & ".Offset", 30#
    Next value
    Audit03SetPlotSetting table, "Plot.RebarLabels.LineEnabled", "Yes"
    degrees = Array(0#, 30#, 45#, 90#, 135#, 225#, 315#)
    For Each frame In Array(430#, 640#)
        sx = 640# / 400#: sy = CDbl(frame) / 400#
        For Each unit In Array("pt", "mm")
            Audit03SetPlotSetting table, "Plot.Dimensions.TextUnits", CStr(unit)
            Audit03SetPlotSetting table, "Plot.RebarLabels.TextUnits", CStr(unit)
            For Each side In Array("Outside", "Inside")
                Audit03SetPlotSetting table, "Plot.Dimensions.Placement", CStr(side)
                Audit03SetPlotSetting table, "Plot.RebarLabels.Placement", CStr(side)
                settings.LoadFromRange table
                For Each value In degrees
                    For Each reverse In Array(1#, -1#)
                        counter = counter + 1: prefix = "arbitrary.layout." & CStr(counter)
                        angle = CDbl(value) * GEOM_PI / 180#: dx = Cos(angle) * CDbl(reverse): dy = Sin(angle) * CDbl(reverse)
                        nx = -dy: ny = dx
                        Set layout = New CPlotAnnotationLayout: layout.ConfigureFromSettings settings
                        layout.Initialize -200#, 200#, -200#, 200#, 0#, 0#, 640#, CDbl(frame), "mm"
                        layout.AddDimension -50# * dx, -50# * dy, 50# * dx, 50# * dy, nx, ny, "L=100"
                        layout.AddRebarLabel -50# * dx, -50# * dy, 50# * dx, 50# * dy, nx, ny, "8" & ChrW$(216) & "20"
                        AssertTrue stats, prefix & ".count", layout.Count = 6
                        px = dx * sx: py = -dy * sy: length = Sqr(px * px + py * py): px = px / length: py = py / length
                        normalX = nx * sx: normalY = -ny * sy: length = Sqr(normalX * normalX + normalY * normalY)
                        normalX = normalX / length: normalY = normalY / length
                        sign = 1#: If side = "Inside" Then sign = -1#
                        qx = -py: qy = px: If qx * normalX + qy * normalY < 0# Then qx = -qx: qy = -qy
                        For Each item In Array(4, 6)
                            If Abs(layout.Rotation(item)) = 90# Then
                                AssertClose stats, prefix & ".vertical." & CStr(item), Abs(px), 0#, 0.000001
                                AssertTrue stats, prefix & ".native.vertical." & CStr(item), layout.X2(item) < layout.Y2(item)
                                If item = 4 Then AssertClose stats, prefix & ".dimension.upward", layout.Rotation(item), 90#, 0#
                                If item = 6 Then
                                    If normalX * sign > 0# Then
                                        AssertClose stats, prefix & ".rebar.right.downward", layout.Rotation(item), -90#, 0#
                                    Else
                                        AssertClose stats, prefix & ".rebar.left.upward", layout.Rotation(item), 90#, 0#
                                    End If
                                End If
                            Else
                                angle = layout.Rotation(item) * GEOM_PI / 180#
                                AssertClose stats, prefix & ".parallel." & CStr(item), Cos(angle) * py - Sin(angle) * px, 0#, 0.000001
                                AssertTrue stats, prefix & ".readable." & CStr(item), Abs(layout.Rotation(item)) <= 90#
                            End If
                            ' Один отсчёт TextGap от центра текста для всех углов, включая 0/90.
                                cx = layout.X1(item) + layout.X2(item) / 2#: cy = layout.Y1(item) + layout.Y2(item) / 2#
                                If item = 4 Then cx = cx - 0.5 * (layout.X1(3) + layout.X2(3)): cy = cy - 0.5 * (layout.Y1(3) + layout.Y2(3))
                                If item = 6 Then cx = cx - 0.5 * (layout.X1(5) + layout.X2(5)): cy = cy - 0.5 * (layout.Y1(5) + layout.Y2(5))
                                gap = 9#: If unit = "mm" Then gap = 9# * GeomMin(sx, sy)
                                AssertClose stats, prefix & ".text.center.gap." & CStr(item), sign * (cx * qx + cy * qy), gap, 0.000001
                        Next item
                        savedRotation = layout.Rotation(6): savedLeft = layout.X1(6)
                        layout.Initialize -200#, 200#, -200#, 200#, 0#, 0#, 640#, CDbl(frame), "mm"
                        layout.AddDimension -50# * dx, -50# * dy, 50# * dx, 50# * dy, nx, ny, "L=100"
                        layout.AddRebarLabel -50# * dx, -50# * dy, 50# * dx, 50# * dy, nx, ny, "8" & ChrW$(216) & "20"
                        AssertTrue stats, prefix & ".repeat", layout.Count = 6 And layout.Rotation(6) = savedRotation And layout.X1(6) = savedLeft
                    Next reverse
                Next value
            Next side
        Next unit
    Next frame
    GoTo Finished
Failed:
    AssertTrue stats, "arbitrary.layout.runtime." & CStr(Err.Number) & "." & Err.Description, False
Finished:
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
    AppendLine stats, "TOTAL_ARBITRARY_ANNOTATION_LAYOUT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunArbitraryAnnotationLayoutTests = stats.Report
End Function

' Возвращает применимую ячейку исходной компактной таблицы по runtime-key.
' Названия RebarLabels/Dimensions переводятся только тестовой картой;
' production reader проверяется независимо через реальную ссылку имени.
Private Function Audit03NamedAnnotationCell(ByVal table As Object, ByVal settingKey As String) As Object
    Dim parts As Variant, row As Long, column As Long
    parts = Split(settingKey, ".")
    If CStr(parts(1)) = "RebarLabels" Then column = 2 Else column = 3
    For row = 2 To table.Rows.Count
        If CStr(table.Cells(row, 1).Value2) = CStr(parts(2)) Then
            Set Audit03NamedAnnotationCell = table.Cells(row, column): Exit Function
        End If
    Next row
    Err.Raise 5, "Audit03NamedAnnotationCell", "В исходной таблице нет параметра " & settingKey
End Function

' Все 21 применимое поле меняется в настоящей таблице Config. Две позиции
' именованного диапазона проверяют адрес ошибок и независимость от layout
' листа. Численные эффекты заданы независимо в points при фиксированном
' transform; стрелки дополнительно проверяются у настоящих Chart.Shape.
Private Sub TestAudit03NamedAnnotationSettings(ByRef stats As TUiTestStats)
    Dim fixture As Object, config As Object, sheet As Object, source As Object, table As Object, target As Object
    Dim settings As CSystemSettingsReader, layout As CPlotAnnotationLayout, reader As CSectionPlotDataReader, plotter As CSectionPlotter
    Dim keys As Variant, changed As Variant, addresses As Variant, name As Variant, key As Variant, groupName As Variant, bad As Variant
    Dim baseline As Variant, geometry As Variant, annotations(1 To 2, 1 To 10) As Variant, props(1 To 2, 1 To 4) As Variant
    Dim index As Long, position As Long, cases As Long, solveCount As Long, itemIndex As Long, fieldName As String, parts As Variant
    Dim cell As Object, chart As Object, marker As Object, shape As Object, code As Long, reason As String, prefix As String, measured As Double
    On Error GoTo Failed
    solveCount = SectionEquilibriumSolveCount()
    Set fixture = Application.Workbooks.Add(-4167): Set config = fixture.Worksheets(1): config.Name = "Config"
    Set sheet = fixture.Worksheets.Add: sheet.Name = "Results": fixture.Worksheets.Add.Name = "Расчет"
    addresses = Array("A5", "H5", "P5", "P25"): index = 0
    For Each name In Array("rngSystemSettings", "rngPlotAnnotationSettings", "rngUnitSettings", "rngSignConventionSettings")
        Set source = ThisWorkbook.Names.Item(CStr(name)).RefersToRange
        Set target = config.Range(CStr(addresses(index))).Resize(source.Rows.Count, source.Columns.Count)
        target.NumberFormat = "@": target.Value2 = source.Value2
        fixture.Names.Add Name:=CStr(name), RefersTo:="=Config!" & target.Address: index = index + 1
    Next name
    Set table = fixture.Names.Item("rngPlotAnnotationSettings").RefersToRange
    For Each groupName In Array("Dimensions", "RebarLabels")
        Audit03NamedAnnotationCell(table, "Plot." & CStr(groupName) & ".Enabled").Value2 = "Yes"
        Audit03NamedAnnotationCell(table, "Plot." & CStr(groupName) & ".TextHeight").Value2 = 6#
        Audit03NamedAnnotationCell(table, "Plot." & CStr(groupName) & ".TextGap").Value2 = 3#
        Audit03NamedAnnotationCell(table, "Plot." & CStr(groupName) & ".Offset").Value2 = 10#
        Audit03NamedAnnotationCell(table, "Plot." & CStr(groupName) & ".LineWeight").Value2 = 1.5
        Audit03NamedAnnotationCell(table, "Plot." & CStr(groupName) & ".Color").Value2 = "0,0,0"
        Audit03NamedAnnotationCell(table, "Plot." & CStr(groupName) & ".TextUnits").Value2 = "mm"
        Audit03NamedAnnotationCell(table, "Plot." & CStr(groupName) & ".Placement").Value2 = "Outside"
    Next groupName
    Audit03NamedAnnotationCell(table, "Plot.Dimensions.ExtensionLineWeight").Value2 = 0.75
    Audit03NamedAnnotationCell(table, "Plot.Dimensions.ExtensionLineColor").Value2 = "1,2,3"
    Audit03NamedAnnotationCell(table, "Plot.Dimensions.ArrowType").Value2 = "Triangle"
    Audit03NamedAnnotationCell(table, "Plot.Dimensions.ArrowSize").Value2 = "Medium"
    Audit03NamedAnnotationCell(table, "Plot.RebarLabels.LineEnabled").Value2 = "Yes"
    baseline = table.Value2
    Set settings = New CSystemSettingsReader: Set layout = New CPlotAnnotationLayout
    Set reader = New CSectionPlotDataReader: Set plotter = New CSectionPlotter
    geometry = Audit03GeometrySnapshotArray("mm", 1#): sheet.Range("A5").Resize(3, 15).Value2 = geometry
    fixture.Names.Add Name:="rngNDMSectionGeometry", RefersTo:="=Results!$A$5"
    Audit03PrepareContourFixture fixture
    props(1, 1) = "LoadCase": props(1, 2) = "Parameter": props(1, 3) = "Value": props(1, 4) = "Unit"
    props(2, 1) = "ALL": props(2, 2) = "Output.LengthUnit": props(2, 3) = "mm": props(2, 4) = "-"
    sheet.Range("R5").Resize(2, 4).Value2 = props
    fixture.Names.Add Name:="rngNDMSectionProperties", RefersTo:="=Results!$R$5"
    keys = Array("AnnotationType", "AnnotationID", "StartX", "StartY", "EndX", "EndY", "OutsideNormalX", "OutsideNormalY", "Text", "Unit")
    For index = 0 To UBound(keys): annotations(1, index + 1) = keys(index): Next index
    annotations(2, 1) = "DIMENSION": annotations(2, 2) = "BRIDGE_DIM": annotations(2, 3) = -50#: annotations(2, 4) = 0#
    annotations(2, 5) = 50#: annotations(2, 6) = 0#: annotations(2, 7) = 0#: annotations(2, 8) = 1#: annotations(2, 9) = "DIM": annotations(2, 10) = "mm"
    sheet.Range("A20").Resize(2, 10).Value2 = annotations
    fixture.Names.Add Name:="rngNDMSectionAnnotations", RefersTo:="=Results!$A$20"
    keys = Array("Plot.Dimensions.Enabled", "Plot.RebarLabels.Enabled", "Plot.Dimensions.Placement", "Plot.RebarLabels.Placement", _
        "Plot.Dimensions.Offset", "Plot.RebarLabels.Offset", "Plot.RebarLabels.LineEnabled", "Plot.Dimensions.LineWeight", _
        "Plot.RebarLabels.LineWeight", "Plot.Dimensions.ExtensionLineWeight", "Plot.Dimensions.TextUnits", "Plot.RebarLabels.TextUnits", _
        "Plot.Dimensions.TextHeight", "Plot.RebarLabels.TextHeight", "Plot.Dimensions.TextGap", "Plot.RebarLabels.TextGap", _
        "Plot.Dimensions.Color", "Plot.RebarLabels.Color", "Plot.Dimensions.ExtensionLineColor", "Plot.Dimensions.ArrowType", "Plot.Dimensions.ArrowSize")
    changed = Array("No", "No", "Inside", "Inside", 20#, 20#, "No", 2.5, 2.5, 1.25, "pt", "pt", _
        8#, 8#, 5#, 5#, "17,34,51", "17,34,51", "51,34,17", "Open", "Wide")
    For position = 0 To 1
        If position = 0 Then Set table = config.Range("H5").Resize(UBound(baseline, 1), UBound(baseline, 2)) Else Set table = config.Range("CH800").Resize(UBound(baseline, 1), UBound(baseline, 2))
        table.NumberFormat = "@": table.Value2 = baseline
        fixture.Names.Item("rngPlotAnnotationSettings").RefersTo = "=Config!" & table.Address
        For index = 0 To UBound(keys)
            cases = cases + 1: table.Value2 = baseline
            key = keys(index): Set cell = Audit03NamedAnnotationCell(table, CStr(key)): cell.Value2 = changed(index)
            settings.LoadFromWorkbook fixture
            prefix = "audit03.namedAnnotation.p" & CStr(position) & "." & CStr(key)
            AssertTrue stats, prefix & ".transmitted", settings.GetRequiredString(CStr(key)) = CStr(changed(index))
            layout.ConfigureFromSettings settings
            layout.Initialize -100#, 100#, -100#, 100#, 0#, 0#, 640#, 430#, "mm"
            layout.AddDimension -50#, 0#, 50#, 0#, 0#, 1#, "DIM"
            layout.AddRebarLabel -50#, 0#, 50#, 0#, 0#, 1#, "REBAR"
            parts = Split(CStr(key), "."): fieldName = CStr(parts(2))
            If CStr(parts(1)) = "Dimensions" Then itemIndex = 4 Else itemIndex = 6
            Select Case fieldName
                Case "Enabled"
                    If CStr(parts(1)) = "Dimensions" Then measured = 2# Else measured = 4#
                    AssertTrue stats, prefix & ".effect", layout.Count = CLng(measured)
                Case "LineEnabled": AssertTrue stats, prefix & ".effect", layout.Count = 5
                Case "Placement"
                    If itemIndex = 4 Then measured = 199.95 Else measured = 242.95
                    AssertClose stats, prefix & ".effect", layout.Y1(itemIndex) + layout.Y2(itemIndex) / 2#, measured, 0.000000001
                Case "Offset": AssertClose stats, prefix & ".effect", layout.Y1(itemIndex - 1), 172#, 0.000000001
                Case "LineWeight": AssertClose stats, prefix & ".effect", layout.Weight(itemIndex - 1), 2.5, 0#
                Case "ExtensionLineWeight": AssertClose stats, prefix & ".effect", layout.Weight(1), 1.25, 0#
                Case "TextUnits": AssertClose stats, prefix & ".effect", layout.FontSize(itemIndex), 6#, 0.000000001
                Case "TextHeight": AssertClose stats, prefix & ".effect", layout.FontSize(itemIndex), 17.2, 0.000000001
                Case "TextGap": AssertClose stats, prefix & ".effect", layout.Y1(itemIndex) + layout.Y2(itemIndex) / 2#, 182.75, 0.000000001
                Case "Color": AssertTrue stats, prefix & ".effect", layout.Color(itemIndex - 1) = RGB(17, 34, 51) And layout.Color(itemIndex) = RGB(17, 34, 51)
                Case "ExtensionLineColor": AssertTrue stats, prefix & ".effect", layout.Color(1) = RGB(51, 34, 17)
                Case "ArrowType", "ArrowSize"
                    reader.LoadGeometryPreviewFromWorkbook fixture, settings: plotter.Draw fixture, reader, settings
                    Set chart = fixture.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot"): Set marker = Nothing
                    For Each shape In chart.Chart.Shapes
                        If InStr(1, shape.Name, "AnnotationLine", vbTextCompare) > 0 Then
                            If shape.Line.BeginArrowheadStyle > 1 Then Set marker = shape: Exit For
                        End If
                    Next shape
                    AssertTrue stats, prefix & ".line", Not marker Is Nothing
                    If Not marker Is Nothing Then
                        If fieldName = "ArrowType" Then
                            AssertTrue stats, prefix & ".effect", marker.Line.BeginArrowheadStyle = 6 And marker.Line.EndArrowheadStyle = 6
                        Else
                            AssertTrue stats, prefix & ".effect", marker.Line.BeginArrowheadLength = 3 And marker.Line.BeginArrowheadWidth = 3
                        End If
                    End If
            End Select
            For Each bad In Array("", "TODO", "INVALID", CVErr(2015))
                cases = cases + 1: table.Value2 = baseline: cell.Value2 = bad
                settings.LoadFromWorkbook fixture
                On Error Resume Next
                Err.Clear
                If fieldName = "ArrowType" Or fieldName = "ArrowSize" Then
                    reader.LoadGeometryPreviewFromWorkbook fixture, settings: plotter.Draw fixture, reader, settings
                Else
                    layout.ConfigureFromSettings settings
                End If
                code = Err.Number: reason = Err.Description
                On Error GoTo Failed
                AssertTrue stats, prefix & ".bad" & CStr(VarType(bad)) & ".rejected", code <> 0
                AssertTrue stats, prefix & ".bad" & CStr(VarType(bad)) & ".address", InStr(1, reason, cell.Address(False, False), vbTextCompare) > 0 And InStr(1, reason, "Config", vbTextCompare) > 0
                AppendLine stats, "NAMED_ANNOTATION_ERROR: " & prefix & "|" & reason
            Next bad
        Next index
    Next position
    AssertTrue stats, "audit03.namedAnnotation.noSolve", SectionEquilibriumSolveCount() = solveCount
    AppendLine stats, "NAMED_ANNOTATION_CASES: fields=21; positions=2; variants=" & CStr(cases) & "; equilibriumCases=0"
    GoTo Cleanup
Failed:
    AssertTrue stats, "audit03.namedAnnotation.runtime." & CStr(Err.Number) & "." & Err.Description, False
Cleanup:
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
End Sub

' Проверяет автоматическое обновление через настоящий workbook-сценарий.
' Отдельная полная копия книги сохраняет Config/Results исходного fixture;
' НДС-состояния не запрашиваются, счетчик solve не растет.
Public Function RunAudit03AutoPlotTests() As String
    Dim stats As TUiTestStats
    TestAudit03AutoPlotContracts stats
    AppendLine stats, "TOTAL_AUDIT03_AUTO_PLOT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03AutoPlotTests = stats.Report
End Function

' Проверяет независимость расчёта от настроек автообновления схемы.
' Ошибки оформления возвращаются с адресом после сохранения результатов;
' отключённая схема не читает неактивный AutoUpdate и сохраняет прежний Chart.
Private Sub TestAudit03AutoPlotContracts(ByRef stats As TUiTestStats)
    Dim fixture As Object, config As Object, table As Object, cell As Object, profiles As Object, loads As Object, durationLoads As Object
    Dim chart As Object, marker As Object, baseline As Variant, bad As Variant, key As Variant
    Dim path As String, row As Long, column As Long, profileColumn As Long, position As Long, cases As Long, mode As Long
    Dim code As Long, reason As String, prefix As String, result As String
    On Error GoTo Failed
    path = ThisWorkbook.Path & "\Audit03_AutoPlot_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsm"
    ThisWorkbook.SaveCopyAs path
    Set fixture = Application.Workbooks.Open(path, 0, False)
    AssertTrue stats, "audit03.autoPlot.ownedCopy", StrComp(fixture.FullName, ThisWorkbook.FullName, vbTextCompare) <> 0
    Set config = fixture.Worksheets.Item("Config")
    Set table = fixture.Names.Item("rngSystemSettings").RefersToRange
    Audit03SetPlotSetting table, "General.ExecutionReportEnabled", "No"
    Audit03SetPlotSetting table, "Geometry.Source", "Generated"
    Audit03SetPlotSetting table, "Geometry.Type", "Circle"
    ' Выполняем действительный stability-only запуск без НДС-состояний.
    Audit03SetPlotSetting table, "Geometry.Source", "Generated": Audit03SetPlotSetting table, "Geometry.Type", "Circle"
    Audit03SetPlotSetting table, "Mesh.StepX", "100": Audit03SetPlotSetting table, "Mesh.StepY", "100"
    Set profiles = fixture.Names.Item("rngCalculationProfiles").RefersToRange
    For row = 1 To profiles.Rows.Count
        For column = 3 To profiles.Columns.Count
            If StrComp(Trim$(CStr(profiles.Cells(row, column).Value2)), "PR1", vbTextCompare) = 0 Then
                profileColumn = column: Exit For
            End If
        Next column
        If profileColumn > 0 Then Exit For
    Next row
    If profileColumn = 0 Then Err.Raise vbObjectError + 4250, "modTestPlotConfig", "В собственном fixture не найден ProfileId PR1."
    For row = 2 To profiles.Rows.Count
        Select Case CStr(profiles.Cells(row, 2).Value2)
            Case "Calculation.Strength.DirectState", "Calculation.Strength.Capacity", "Calculation.Crack.Width"
                profiles.Cells(row, profileColumn).Value2 = "No"
            Case "Calculation.Stability.Enabled"
                profiles.Cells(row, profileColumn).Value2 = "Yes"
        End Select
    Next row
    Set loads = fixture.Names.Item("rngLoadCombinations").RefersToRange
    loads.Offset(1, 0).Resize(loads.Rows.Count - 1, loads.Columns.Count).ClearContents
    loads.Cells(2, 1).Value2 = "NO_STATE": loads.Cells(2, 2).Value2 = 1#: loads.Cells(2, 3).Value2 = 0#: loads.Cells(2, 4).Value2 = 0#
    loads.Cells(2, 5).Value2 = "PR1": loads.Cells(2, 6).Value2 = "Auto"
    ' Устойчивость требует строку с тем же номером сочетания; нули задают отсутствие дополнительной нагрузки.
    Set durationLoads = fixture.Names.Item("rngStabilityDurationLoads").RefersToRange
    durationLoads.Offset(1, 0).Resize(durationLoads.Rows.Count - 1, durationLoads.Columns.Count).ClearContents
    durationLoads.Cells(2, 1).Value2 = loads.Cells(2, 1).Value2
    durationLoads.Cells(2, 2).Resize(1, 3).Value2 = 0#
    ' Задаем сжатие явно: исходная система знаков книги не должна превращать
    ' stability-only fixture в неприменимую проверку растянутого элемента.
    Audit03SetPlotSetting fixture.Names.Item("rngSignConventionSettings").RefersToRange, "+N", "Compression"
    result = FormSectionForWorkbook(fixture)
    Audit03SetPlotSetting table, "Geometry.Type", "INVALID_UNUSED_GEOMETRY"
    baseline = table.Value2
    On Error Resume Next
    Set chart = fixture.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")
    On Error GoTo Failed
    If chart Is Nothing Then Set chart = fixture.Worksheets.Item("Расчет").ChartObjects.Add(41#, 53#, 700#, 480#)
    chart.Name = "chtNDMSectionPlot"
    Set marker = chart.Chart.Shapes.AddShape(1, 11#, 17#, 20#, 20#): marker.Name = "NDMPlot_Audit03Sentinel"
    For position = 0 To 1
        If position = 1 Then Set table = config.Range("CH800").Resize(UBound(baseline, 1), UBound(baseline, 2))
        table.NumberFormat = "@": table.Value2 = baseline
        fixture.Names.Item("rngSystemSettings").RefersTo = "=Config!" & table.Address
        For Each key In Array("Plot.Enabled", "Plot.AutoUpdateAfterCalculation")
            For Each bad In Array("", "TODO", "INVALID", CVErr(2015))
                cases = cases + 1: table.Value2 = baseline
                Audit03SetPlotSetting table, "Plot.Enabled", "Yes"
                Set cell = Audit03PlotSettingCell(table, CStr(key)): cell.Value2 = bad
                Audit03CaptureCalculation fixture, code, reason, result
                prefix = "audit03.autoPlot.invalid.p" & CStr(position) & "." & CStr(key) & ".v" & CStr(cases)
                AssertTrue stats, prefix & ".warning", code = 0 And InStr(1, result, "Схема не обновлена", vbTextCompare) > 0 And InStr(1, result, CStr(key), vbTextCompare) > 0
                AssertTrue stats, prefix & ".address", InStr(1, result, "Config", vbTextCompare) > 0 And InStr(1, result, cell.Address(False, False), vbTextCompare) > 0
                AssertTrue stats, prefix & ".snapshot", ResultsGeometrySource(fixture) = "Generated" And SectionEquilibriumSolveCount() = 0
                AssertTrue stats, prefix & ".chart", Audit03PlotShapeExists(chart, "Audit03Sentinel")
                AppendLine stats, "AUTO_PLOT_WARNING: " & prefix & "|" & result
            Next bad
        Next key
        table.Value2 = baseline: Audit03SetPlotSetting table, "Plot.Enabled", "Yes"
        Set cell = Audit03PlotSettingCell(table, "Plot.AutoUpdateAfterCalculation"): cell.Value2 = "Yes": cell.Offset(0, -1).Value2 = "Audit03.RemovedAutoUpdate"
        Audit03CaptureCalculation fixture, code, reason, result
        AssertTrue stats, "audit03.autoPlot.missing.p" & CStr(position), code = 0 And InStr(1, result, "Plot.AutoUpdateAfterCalculation", vbTextCompare) > 0 And InStr(1, result, "Схема не обновлена", vbTextCompare) > 0
        table.Value2 = baseline: Audit03SetPlotSetting table, "Plot.Enabled", "No"
        Audit03SetPlotSetting table, "Plot.AutoUpdateAfterCalculation", CVErr(2015)
        Audit03CaptureCalculation fixture, code, reason, result
        AssertTrue stats, "audit03.autoPlot.inactive.p" & CStr(position), code = 0 And InStr(1, result, "Plot.AutoUpdateAfterCalculation", vbTextCompare) = 0 And Audit03PlotShapeExists(chart, "Audit03Sentinel")
    Next position
    For mode = 0 To 3
        Audit03SetPlotSetting table, "Plot.Enabled", "Yes": Audit03SetPlotSetting table, "Plot.AutoUpdateAfterCalculation", "Yes"
        Select Case mode
            Case 0: Audit03SetPlotSetting table, "Plot.AutoUpdateAfterCalculation", "No"
            Case 1: Audit03SetPlotSetting table, "Plot.Enabled", "No"
            Case 2: Audit03SetPlotSetting table, "Plot.Enabled", "No": Audit03SetPlotSetting table, "Plot.AutoUpdateAfterCalculation", CVErr(2015)
        End Select
        If Not Audit03PlotShapeExists(chart, "Audit03Sentinel") Then
            Set marker = chart.Chart.Shapes.AddShape(1, 11#, 17#, 20#, 20#): marker.Name = "NDMPlot_Audit03Sentinel"
        End If
        cases = cases + 1: Audit03CaptureCalculation fixture, code, reason, result
        prefix = "audit03.autoPlot.noState.mode" & CStr(mode)
        AssertTrue stats, prefix & ".completed", code = 0
        ' Batch сбрасывает счетчик в начале запуска; проверяем этот запуск,
        ' а не равенство счетчику предыдущих тестов общей suite.
        AssertTrue stats, prefix & ".noSolve", SectionEquilibriumSolveCount() = 0
        If code = 0 Then
            reason = CStr(fixture.Names.Item("rngBatchSummary").RefersToRange.Offset(12, 3).Value2)
            AssertTrue stats, prefix & ".validProfile", reason = "OK" Or reason = "FAIL"
        End If
        If mode < 3 Then
            AssertTrue stats, prefix & ".preserved", Audit03PlotShapeExists(chart, "Audit03Sentinel")
        Else
            AssertTrue stats, prefix & ".cleared", Not Audit03PlotShapeExists(chart, "Audit03Sentinel")
        End If
        AppendLine stats, "AUTO_PLOT_RUN: " & prefix & "|" & reason & "|" & result
    Next mode
    AssertTrue stats, "audit03.autoPlot.noSolve", SectionEquilibriumSolveCount() = 0
    AppendLine stats, "AUTO_PLOT_CASES: variants=" & CStr(cases) & "; workbookFixtures=1; equilibriumCases=0"
    GoTo Cleanup
Failed:
    AssertTrue stats, "audit03.autoPlot.runtime." & CStr(Err.Number) & "." & Err.Description, False
Cleanup:
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
End Sub

' Перехватывает workbook-ошибку без MsgBox; причины и реальные изменения
' проверяет вызывающий тест, штатный расчет не подменяется fake-объектом.
Private Sub Audit03CaptureCalculation(ByVal workbook As Object, ByRef code As Long, _
        ByRef reason As String, ByRef result As String)
    code = 0: reason = vbNullString: result = vbNullString
    On Error GoTo Failed
    result = RunSectionCalculationForWorkbook(workbook, False)
    Exit Sub
Failed:
    code = Err.Number: reason = Err.Description
    Err.Clear
End Sub

' Проверяет метаданные сохраненного snapshot отдельно от расчетного ядра.
' Порядок строк, текущие Config units и перенос якоря не меняют координаты;
' поврежденные активные единицы/флаги/статусы не подменяются defaults.
Public Function RunAudit03SnapshotMetadataTests() As String
    Dim stats As TUiTestStats
    TestAudit03SnapshotMetadataContracts stats
    AppendLine stats, "TOTAL_AUDIT03_SNAPSHOT_METADATA: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03SnapshotMetadataTests = stats.Report
End Function

' Создает собственный snapshot с единицами после численных свойств. Проверяет
' публичные reader-режимы, точные данные Results и восстановление того же
' reader-а после ошибки с фактическим, в том числе перенесенным адресом.
Private Sub TestAudit03SnapshotMetadataContracts(ByRef stats As TUiTestStats)
    Dim fixture As Object, config As Object, sheet As Object, source As Object, table As Object, anchor As Object, cell As Object
    Dim settings As CSystemSettingsReader, reader As CSectionPlotDataReader
    Dim geometry As Variant, props(1 To 18, 1 To 4) As Variant, elements(1 To 3, 1 To 7) As Variant
    Dim annotations(1 To 2, 1 To 9) As Variant, keys As Variant, values As Variant, units As Variant, factors As Variant
    Dim index As Long, row As Variant, position As Long, mode As Long, variantIndex As Long, code As Long, cases As Long
    Dim unit As Variant, bad As Variant, flag As Variant, state As Variant, baseline As Variant, actual As Variant
    Dim factor As Double, curvatureFactor As Double, reason As String, prefix As String, solveCount As Long
    On Error GoTo Failed
    solveCount = SectionEquilibriumSolveCount()
    Set fixture = Application.Workbooks.Add(-4167): Set config = fixture.Worksheets(1): config.Name = "MetadataConfig"
    Set sheet = fixture.Worksheets.Add: sheet.Name = "Results"
    Set table = Audit03PlotSettingsTable(config.Range("A5"))
    Set source = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    config.Range("AF5").Resize(source.Rows.Count, source.Columns.Count).Value2 = source.Value2
    For row = 2 To source.Rows.Count
        If CStr(source.Cells(row, 2).Value2) = "Visualization.State" Then config.Range("AF5").Cells(row, 3).Value2 = "StrengthState"
        If CStr(source.Cells(row, 2).Value2) = "Visualization.Quantity" Then config.Range("AF5").Cells(row, 3).Value2 = "Stress"
    Next row
    fixture.Names.Add Name:="rngCalculationProfiles", RefersTo:="=MetadataConfig!" & config.Range("AF5").Resize(source.Rows.Count, source.Columns.Count).Address
    fixture.Names.Add Name:="rngNDMSectionGeometry", RefersTo:="=Results!$A$5"
    fixture.Names.Add Name:="rngNDMElementResults", RefersTo:="=Results!$R$35"
    Audit03PrepareContourFixture fixture
    fixture.Names.Add Name:="rngNDMSectionAnnotations", RefersTo:="=Results!$A$20"
    fixture.Names.Add Name:="rngNDMSectionProperties", RefersTo:="=Results!$R$5"
    elements(1, 1) = "LoadCase": elements(1, 2) = "ProfileId": elements(1, 3) = "StateType": elements(1, 4) = "ElementID"
    elements(1, 5) = "Strain": elements(1, 6) = "Stress": elements(1, 7) = "PhysicalState"
    For row = 2 To 3
        elements(row, 1) = "PLOT": elements(row, 2) = "PR1": elements(row, 3) = "StrengthState"
        elements(row, 5) = -0.0001: elements(row, 6) = -1.25: elements(row, 7) = "Compression"
    Next row
    elements(2, 4) = "C1": elements(3, 4) = "R1"
    sheet.Range("R35").Resize(3, 7).Value2 = elements
    keys = Array("AnnotationType", "StartX", "StartY", "EndX", "EndY", "OutsideNormalX", "OutsideNormalY", "Text", "Unit")
    For index = 0 To UBound(keys): annotations(1, index + 1) = keys(index): Next index
    annotations(2, 1) = "DIMENSION": annotations(2, 2) = -5#: annotations(2, 3) = 0#
    annotations(2, 4) = 5#: annotations(2, 5) = 0#: annotations(2, 6) = 0#: annotations(2, 7) = 1#
    annotations(2, 8) = "DIM": annotations(2, 9) = "-"
    keys = Array("Bounds.MinX", "Bounds.MaxX", "Bounds.MinY", "Bounds.MaxY", "LoadReferenceX", "LoadReferenceY", _
        "State.StrengthState.Epsilon0", "State.StrengthState.KappaX", "State.StrengthState.KappaY", _
        "State.StrengthState.ExtensionUsed", "State.StrengthState.Status", "ProfileId", _
        "Output.LengthUnit", "Output.AreaUnit", "Output.StressUnit", "Output.CurvatureUnit", "Concrete.PrincipalAngle")
    values = Array(-100#, 100#, -100#, 100#, 25#, -50#, -0.0001, 0.000002, -0.000003, "False", "OK", "PR1", "mm", "mm2", "MPa", "1/mm", 0.25)
    props(1, 1) = "LoadCase": props(1, 2) = "Parameter": props(1, 3) = "Value": props(1, 4) = "Unit"
    For index = 0 To UBound(keys)
        row = index + 2: props(row, 1) = "ALL"
        If index >= 6 And index <= 11 Then props(row, 1) = "PLOT"
        props(row, 2) = keys(index): props(row, 3) = values(index): props(row, 4) = "-"
    Next index
    Set settings = New CSystemSettingsReader: settings.LoadFromRange table
    Set reader = New CSectionPlotDataReader
    units = Array("mm", "cm", "m"): factors = Array(1#, 10#, 1000#)
    For position = 0 To 1
        If position = 0 Then Set anchor = sheet.Range("R5") Else Set anchor = sheet.Range("CH800")
        sheet.Range("R5:U22").ClearContents
        fixture.Names.Item("rngNDMSectionProperties").RefersTo = "=Results!" & anchor.Address
        For index = 0 To UBound(units)
            unit = units(index): factor = CDbl(factors(index))
            geometry = Audit03GeometrySnapshotArray(CStr(unit), factor)
            geometry(1, 3) = "X": geometry(1, 4) = "Y": geometry(1, 5) = "Area"
            sheet.Range("A5").Resize(3, 15).Value2 = geometry
            For row = 2 To 7: props(row, 3) = CDbl(values(row - 2)) / factor: Next row
            props(14, 3) = unit: props(15, 3) = CStr(unit) & "2"
            If index = 0 Then curvatureFactor = 1# Else curvatureFactor = 1000#
            props(9, 3) = 0.000002 * curvatureFactor: props(10, 3) = -0.000003 * curvatureFactor
            If index = 0 Then props(17, 3) = "1/mm" Else props(17, 3) = "1/m"
            annotations(2, 2) = -50# / factor: annotations(2, 4) = 50# / factor
            anchor.Resize(18, 4).Value2 = props: sheet.Range("A20").Resize(2, 9).Value2 = annotations
            For mode = 0 To 2
                cases = cases + 1: Audit03ReadLifecycleLoad reader, fixture, settings, mode, code, reason
                prefix = "audit03.snapshotMetadata.units.p" & CStr(position) & "." & CStr(unit) & ".mode" & CStr(mode)
                AssertTrue stats, prefix & ".loaded", code = 0 And reader.Count = 2
                If code = 0 Then
                    AssertClose stats, prefix & ".minX", reader.MinX, -100#, 0.000000001
                    AssertClose stats, prefix & ".loadX", reader.LoadReferenceX, 25#, 0.000000001
                    AssertClose stats, prefix & ".elementX", reader.X(1), -23.125, 0.000000001
                    AssertClose stats, prefix & ".area", reader.Area(1), 10828.125, 0.000000001
                    AssertClose stats, prefix & ".annotation", reader.AnnotationStartX(1), -50#, 0.000000001
                    If mode = 0 Then AssertClose stats, prefix & ".curvature", reader.KappaX, 0.000002, 0.000000000001
                End If
            Next mode
        Next index
        ' Возвращаем mm baseline перед ошибками отдельных метаданных.
        For index = 0 To UBound(values): props(index + 2, 3) = values(index): Next index
        anchor.Resize(18, 4).Value2 = props
        sheet.Range("A5").Resize(3, 15).Value2 = Audit03GeometrySnapshotArray("mm", 1#)
        annotations(2, 2) = -50#: annotations(2, 4) = 50#: sheet.Range("A20").Resize(2, 9).Value2 = annotations
        baseline = sheet.Range("A5:X37").Value2
        For Each row In Array(11, 12, 14, 15, 16, 17)
            Set cell = anchor.Cells(row, 3)
            For Each bad In Array("", "TODO", "INVALID", CVErr(2015))
                variantIndex = variantIndex + 1: cases = cases + 1: cell.Value2 = bad
                Audit03ReadLifecycleLoad reader, fixture, settings, 0, code, reason
                prefix = "audit03.snapshotMetadata.invalid.p" & CStr(position) & ".row" & CStr(row) & ".v" & CStr(variantIndex)
                AssertTrue stats, prefix & ".rejected", code <> 0
                AssertTrue stats, prefix & ".field", InStr(1, reason, CStr(anchor.Cells(row, 2).Value2), vbTextCompare) > 0
                AssertTrue stats, prefix & ".address", InStr(1, reason, "Results!" & cell.Address(False, False), vbTextCompare) > 0
                AssertTrue stats, prefix & ".action", InStr(1, reason, "Повторите", vbTextCompare) > 0 Or InStr(1, reason, "Исправьте", vbTextCompare) > 0
                AssertTrue stats, prefix & ".empty", reader.Count = 0 And reader.AnnotationCount = 0
                AppendLine stats, "SNAPSHOT_METADATA_ERROR: " & prefix & "|" & reason
                cell.Value2 = values(row - 2): Audit03ReadLifecycleLoad reader, fixture, settings, 0, code, reason
                AssertTrue stats, prefix & ".recovery", code = 0 And reader.Count = 2
            Next bad
        Next row
        For Each row In Array(2, 9)
            Set cell = anchor.Cells(row, 4): cell.Value2 = CVErr(2015)
            Audit03ReadLifecycleLoad reader, fixture, settings, 0, code, reason
            prefix = "audit03.snapshotMetadata.fieldUnit.p" & CStr(position) & ".row" & CStr(row)
            AssertTrue stats, prefix & ".rejected", code <> 0
            AssertTrue stats, prefix & ".address", InStr(1, reason, "Results!" & cell.Address(False, False), vbTextCompare) > 0
            cell.Value2 = "-"
        Next row
        Set cell = sheet.Range("I21"): cell.Value2 = CVErr(2015)
        Audit03ReadLifecycleLoad reader, fixture, settings, 1, code, reason
        AssertTrue stats, "audit03.snapshotMetadata.annotationUnit.p" & CStr(position), code <> 0 And InStr(1, reason, "Results!I21", vbTextCompare) > 0
        cell.Value2 = "-"
        For Each flag In Array(True, False, "True", "False", "Yes", "No", "да", "нет", 1, 0)
            anchor.Cells(11, 3).Value2 = flag: reader.LoadFromWorkbook fixture, settings
            AssertTrue stats, "audit03.snapshotMetadata.flag.p" & CStr(position) & "." & CStr(flag), reader.ExtensionUsed = CBool(InStr(1, "|true|yes|да|1|", "|" & LCase$(CStr(flag)) & "|", vbBinaryCompare) > 0)
        Next flag
        anchor.Cells(11, 3).Value2 = "False"
        For Each state In Array("OK", "FAIL", "BaseFail", "NumFail", "InputErr", "CalcErr", "N/A")
            anchor.Cells(12, 3).Value2 = state: reader.LoadFromWorkbook fixture, settings
            AssertTrue stats, "audit03.snapshotMetadata.status.p" & CStr(position) & "." & CStr(state), reader.DirectStateStatus = CStr(state)
        Next state
        anchor.Cells(12, 3).Value2 = "OK"
        actual = sheet.Range("A5:X37").Value2: Audit03ComparePlainSnapshot stats, "audit03.snapshotMetadata.unchanged.p" & CStr(position), baseline, actual, 1
        ' Unit и Value распознаются по шапке, а не по взаимному смещению.
        values = anchor.Resize(18, 4).Value2
        actual = values
        For row = 1 To 18
            actual(row, 3) = values(row, 4): actual(row, 4) = values(row, 3)
        Next row
        anchor.Resize(18, 4).Value2 = actual
        reader.LoadFromWorkbook fixture, settings
        AssertClose stats, "audit03.snapshotMetadata.columns.p" & CStr(position), reader.MinX, -100#, 0#
        anchor.Cells(2, 3).Value2 = CVErr(2015)
        Audit03ReadLifecycleLoad reader, fixture, settings, 0, code, reason
        AssertTrue stats, "audit03.snapshotMetadata.columns.address.p" & CStr(position), code <> 0 And InStr(1, reason, "Results!" & anchor.Cells(2, 3).Address(False, False), vbTextCompare) > 0
        anchor.Resize(18, 4).Value2 = values
        ' Неиспользуемое состояние другого LC не влияет на выбранный снимок.
        anchor.Cells(19, 1).Value2 = "OTHER": anchor.Cells(19, 2).Value2 = "State.StrengthState.KappaX"
        anchor.Cells(19, 3).Value2 = CVErr(2015): anchor.Cells(19, 4).Value2 = CVErr(2015)
        reader.LoadFromWorkbook fixture, settings
        AssertTrue stats, "audit03.snapshotMetadata.otherLC.p" & CStr(position), reader.Count = 2 And reader.KappaX = 0.000002
        anchor.Cells(19, 1).Resize(1, 4).ClearContents
        ' Явная единица строки имеет приоритет над общим metadata default.
        anchor.Cells(14, 3).Value2 = "m": anchor.Cells(2, 3).Value2 = -10#: anchor.Cells(2, 4).Value2 = "cm"
        reader.LoadFromWorkbook fixture, settings
        AssertClose stats, "audit03.snapshotMetadata.fieldOverride.p" & CStr(position), reader.MinX, -100#, 0#
        anchor.Resize(18, 4).Value2 = values
        For Each unit In Array("Pa", "kPa", "MPa", "kgf/cm2", "tf/m2")
            anchor.Cells(16, 3).Value2 = unit: reader.LoadFromWorkbook fixture, settings
            AssertTrue stats, "audit03.snapshotMetadata.stress.p" & CStr(position) & "." & CStr(unit), reader.ResultUnit = CStr(unit) And reader.ResultValue(1) = -1.25
        Next unit
        anchor.Cells(16, 3).Value2 = "MPa"
        ' Import-preview не использует plane/status/extension и напряжения.
        ' Missing-state reader сохраняет проверку уже записанных данных LC.
        For Each row In Array(9, 11, 12, 16, 17): anchor.Cells(row, 3).Value2 = CVErr(2015): Next row
        For mode = 1 To 2
            Audit03ReadLifecycleLoad reader, fixture, settings, mode, code, reason
            If mode = 1 Then
                AssertTrue stats, "audit03.snapshotMetadata.geometry.inactive.p" & CStr(position), code = 0 And reader.Count = 2
            Else
                AssertTrue stats, "audit03.snapshotMetadata.missingState.validates.p" & CStr(position), code <> 0 And reader.Count = 0
            End If
        Next mode
        anchor.Resize(18, 4).Value2 = values
        values = Array(-100#, 100#, -100#, 100#, 25#, -50#, -0.0001, 0.000002, -0.000003, "False", "OK", "PR1", "mm", "mm2", "MPa", "1/mm", 0.25)
    Next position
    AssertTrue stats, "audit03.snapshotMetadata.noSolve", SectionEquilibriumSolveCount() = solveCount
    AppendLine stats, "SNAPSHOT_METADATA_CASES: variants=" & CStr(cases) & "; geometryFixtures=1; equilibriumCases=0"
    GoTo Cleanup
Failed:
    AssertTrue stats, "audit03.snapshotMetadata.runtime." & CStr(Err.Number) & "." & Err.Description, False
Cleanup:
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
End Sub

' Неверный активный ввод не должен очищать прежний рисунок. Проверяем
' восстановление, перенос таблицы, выключенные потребители и положительный
' мелкий шаг подписей, для которого номер ячейки сетки выходит за Long.
Private Sub TestAudit03GeneralPlotContracts(ByRef stats As TUiTestStats)
    Dim fixture As Object, config As Object, sheet As Object, table As Object, source As Object
    Dim settings As CSystemSettingsReader, reader As CSectionPlotDataReader, plotter As CSectionPlotter
    Dim values As Variant, keys As Variant, key As Variant, bad As Variant, cell As Object
    Dim props(1 To 26, 1 To 4) As Variant, elements(1 To 5, 1 To 7) As Variant, annotations(1 To 3, 1 To 9) As Variant
    Dim grid() As Variant, gridResults() As Variant, series As Object, colors As Object, channel As Long
    Dim arrow As Variant, size As Variant, shape As Object, style As Long, metric As Long, badIndex As Long, mode As Long
    Dim row As Long, position As Long, index As Long, code As Long, reason As String, prefix As String
    Dim baseline As Variant, annotationBaseline As Variant, annotationTable As Object
    Dim snapshot As Variant, actual As Variant, chart As Object, marker As Object
    Dim changedElements As Variant, expectedColor As Long, elementKind As String, compression As Boolean
    Dim cases As Long, solveCount As Long, geometry As Variant
    On Error GoTo Failed
    solveCount = SectionEquilibriumSolveCount()
    Set fixture = Application.Workbooks.Add(-4167)
    Set config = fixture.Worksheets(1): config.Name = "PlotConfig"
    Set sheet = fixture.Worksheets.Add: sheet.Name = "Results"
    fixture.Worksheets.Add.Name = "Расчет"
    Set table = Audit03NamedGeneralPlotSettingsTable(fixture, config)
    baseline = table.Value2
    Set annotationTable = fixture.Names.Item("rngPlotAnnotationSettings").RefersToRange
    annotationBaseline = annotationTable.Value2
    geometry = Audit03ContourGeometrySnapshot("CONTOUR_CIRCLE")
    sheet.Range("A5").Resize(3, 15).Value2 = geometry
    fixture.Names.Add Name:="rngNDMSectionGeometry", RefersTo:="=Results!$A$5"
    Audit03PrepareContourFixture fixture, "CONTOUR_CIRCLE"
    Set source = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    config.Range("AF5").Resize(source.Rows.Count, source.Columns.Count).Value2 = source.Value2
    For row = 2 To source.Rows.Count
        If CStr(source.Cells(row, 2).Value2) = "Visualization.State" Then config.Range("AF5").Cells(row, 3).Value2 = "StrengthState"
        If CStr(source.Cells(row, 2).Value2) = "Visualization.Quantity" Then config.Range("AF5").Cells(row, 3).Value2 = "Stress"
        If CStr(source.Cells(row, 2).Value2) = "Visualization.StressPrecision" Then config.Range("AF5").Cells(row, 3).Value2 = 2
    Next row
    fixture.Names.Add Name:="rngCalculationProfiles", RefersTo:="=PlotConfig!" & config.Range("AF5").Resize(source.Rows.Count, source.Columns.Count).Address
    keys = Array("Bounds.MinX", "Bounds.MaxX", "Bounds.MinY", "Bounds.MaxY", _
        "Concrete.CentroidX", "Concrete.CentroidY", "Concrete.PrincipalAngle", _
        "Transformed.CentroidX", "Transformed.CentroidY", "Transformed.PrincipalAngle", _
        "LoadReferenceX", "LoadReferenceY", "State.StrengthState.Epsilon0", _
        "State.StrengthState.KappaX", "State.StrengthState.KappaY", "ProfileId", _
        "State.StrengthState.ExtensionUsed", "State.StrengthState.Status", "Output.LengthUnit")
    values = Array(-100#, 100#, -100#, 100#, 0#, 0#, 0.125, 0#, 0#, 0.25, _
        0#, 0#, -0.0001, 0.000002, 0#, "PR1", "False", "OK", "mm")
    props(1, 1) = "LoadCase": props(1, 2) = "Parameter": props(1, 3) = "Value": props(1, 4) = "Unit"
    For index = 0 To UBound(keys)
        row = index + 2: props(row, 1) = "ALL"
        If index >= 12 And index <= 17 Then props(row, 1) = "PLOT"
        props(row, 2) = keys(index): props(row, 3) = values(index): props(row, 4) = "mm"
        If index = 6 Or index = 9 Or index = 12 Or index >= 15 Then props(row, 4) = "-"
        If index = 13 Or index = 14 Then props(row, 4) = "1/mm"
    Next index
    ' Второе сочетание содержит только собственные state-поля: общие свойства
    ' ALL не дублируются. Его числа специально отличаются от первого сочетания.
    For index = 0 To 5
        For row = 1 To 4: props(21 + index, row) = props(14 + index, row): Next row
        props(21 + index, 1) = "ALT"
    Next index
    props(21, 3) = 0.0002
    sheet.Range("AB5").Resize(26, 4).Value2 = props
    fixture.Names.Add Name:="rngNDMSectionProperties", RefersTo:="=Results!$AB$5"
    elements(1, 1) = "LoadCase": elements(1, 2) = "ProfileId": elements(1, 3) = "StateType": elements(1, 4) = "ElementID"
    elements(1, 5) = "Strain": elements(1, 6) = "Stress, MPa": elements(1, 7) = "PhysicalState"
    elements(2, 1) = "PLOT": elements(2, 2) = "PR1": elements(2, 3) = "StrengthState": elements(2, 4) = "C1"
    elements(2, 5) = -0.0001: elements(2, 6) = -1.25: elements(2, 7) = "Compression"
    elements(3, 1) = "PLOT": elements(3, 2) = "PR1": elements(3, 3) = "StrengthState": elements(3, 4) = "R1"
    elements(3, 5) = 0.000015: elements(3, 6) = 2.75: elements(3, 7) = "Tension"
    For row = 4 To 5
        For index = 1 To 7: elements(row, index) = elements(row - 2, index): Next index
        elements(row, 1) = "ALT"
    Next row
    elements(4, 5) = 0.0002: elements(4, 6) = 9.5: elements(4, 7) = "Tension"
    elements(5, 5) = -0.0003: elements(5, 6) = -7.25: elements(5, 7) = "Compression"
    sheet.Range("R35").Resize(5, 7).Value2 = elements
    fixture.Names.Add Name:="rngNDMElementResults", RefersTo:="=Results!$R$35"
    keys = Array("AnnotationType", "StartX", "StartY", "EndX", "EndY", "OutsideNormalX", "OutsideNormalY", "Text", "Unit")
    For index = 0 To UBound(keys): annotations(1, index + 1) = keys(index): Next index
    For row = 2 To 3
        For index = 2 To 7: annotations(row, index) = 0#: Next index
        annotations(row, 9) = "mm"
    Next row
    annotations(2, 1) = "DIMENSION": annotations(2, 2) = -50#: annotations(2, 4) = 50#: annotations(2, 7) = 1#: annotations(2, 8) = "DIM"
    annotations(3, 1) = "REBAR_ANNOTATION": annotations(3, 2) = -50#: annotations(3, 4) = 50#: annotations(3, 7) = 1#: annotations(3, 8) = "REBAR"
    sheet.Range("A20").Resize(3, 9).Value2 = annotations
    fixture.Names.Add Name:="rngNDMSectionAnnotations", RefersTo:="=Results!$A$20"
    snapshot = sheet.Range("A5:AI39").Value2
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook fixture
    Set reader = New CSectionPlotDataReader: reader.LoadFromWorkbook fixture, settings
    Set plotter = New CSectionPlotter
    plotter.Draw fixture, reader, settings
    Set chart = fixture.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")
    chart.Left = 41#: chart.Top = 53#: chart.Width = 700#: chart.Height = 480#
    keys = Array("Plot.ContourEnabled", "Plot.NeutralLineEnabled", "Plot.LoadApplicationPointEnabled", _
        "Plot.ResultLabelsEnabled", "Plot.LegendEnabled", "Plot.PrincipalAxesMode", "Plot.ResultLabelSpacing", _
        "Plot.ResultGradient", "Plot.AxisLabelsEnabled", "Plot.AxisLabelsFontSize", "Plot.LegendMode", _
        "Plot.Color.RebarCompression", "Plot.Color.RebarTension", "Plot.Color.ConcreteCompression", "Plot.Color.ConcreteTension")
    For position = 0 To 1
        If position = 0 Then Set table = config.Range("A5").Resize(UBound(baseline, 1), UBound(baseline, 2)) Else Set table = config.Range("CH800").Resize(UBound(baseline, 1), UBound(baseline, 2))
        table.NumberFormat = "@": table.Value2 = baseline: annotationTable.Value2 = annotationBaseline
        fixture.Names.Item("rngSystemSettings").RefersTo = "='" & config.Name & "'!" & table.Address
        Audit03SetPlotSetting table, "Plot.ResultLabelsEnabled", "Yes"
        For Each key In keys
            Set cell = Audit03PlotSettingCell(table, CStr(key))
            badIndex = 0
            For Each bad In Array("", "TODO", "INVALID", CVErr(2015))
                badIndex = badIndex + 1
                cases = cases + 1: values = cell.Value2: cell.Value2 = bad
                settings.LoadFromWorkbook fixture
                Set marker = chart.Chart.Shapes.AddShape(1, 11#, 17#, 20#, 20#)
                marker.Name = "NDMPlot_Audit03Sentinel"
                Audit03CapturePlotDraw plotter, fixture, reader, settings, code, reason
                prefix = "audit03.generalPlot.p" & CStr(position) & "." & CStr(key) & ".bad" & CStr(badIndex)
                AssertTrue stats, prefix & ".rejected", code <> 0
                AssertTrue stats, prefix & ".key", InStr(1, reason, CStr(key), vbTextCompare) > 0
                AssertTrue stats, prefix & ".address", InStr(1, reason, config.Name, vbTextCompare) > 0 And InStr(1, reason, cell.Address(False, False), vbTextCompare) > 0
                AssertTrue stats, prefix & ".action", InStr(1, reason, "Введите", vbTextCompare) > 0 Or InStr(1, reason, "Выберите", vbTextCompare) > 0 Or InStr(1, reason, "Исправьте", vbTextCompare) > 0 Or InStr(1, reason, "Заполните", vbTextCompare) > 0
                AssertTrue stats, prefix & ".oldChartPreserved", Audit03PlotShapeExists(chart, "Audit03Sentinel")
                AppendLine stats, "GENERAL_PLOT_ERROR: " & prefix & "|" & reason
                On Error Resume Next
                chart.Chart.Shapes.Item("NDMPlot_Audit03Sentinel").Delete
                On Error GoTo Failed
                cell.Value2 = values: settings.LoadFromWorkbook fixture
                Audit03CapturePlotDraw plotter, fixture, reader, settings, code, reason
                AssertTrue stats, prefix & ".recovery", code = 0
            Next bad
        Next key
        For Each bad In Array(0#, -1#)
            cases = cases + 1
            Audit03SetPlotSetting table, "Plot.ResultLabelSpacing", bad
            settings.LoadFromWorkbook fixture
            Audit03CapturePlotDraw plotter, fixture, reader, settings, code, reason
            AssertTrue stats, "audit03.generalPlot.spacing.sign.p" & CStr(position) & "." & CStr(bad), code <> 0 And InStr(1, reason, "Plot.ResultLabelSpacing", vbTextCompare) > 0
        Next bad
        Audit03SetPlotSetting table, "Plot.ResultLabelSpacing", 0.00000001
        settings.LoadFromWorkbook fixture
        Audit03CapturePlotDraw plotter, fixture, reader, settings, code, reason
        AssertTrue stats, "audit03.generalPlot.spacing.smallPositive.p" & CStr(position), code = 0
        AssertTrue stats, "audit03.generalPlot.spacing.label.p" & CStr(position), Audit03PlotTextExists(chart, "-1.25")
        Audit03SetPlotSetting table, "Plot.ResultLabelsEnabled", "No"
        Audit03SetPlotSetting table, "Plot.ResultLabelSpacing", CVErr(2015)
        settings.LoadFromWorkbook fixture
        Audit03CapturePlotDraw plotter, fixture, reader, settings, code, reason
        AssertTrue stats, "audit03.generalPlot.spacing.inactive.p" & CStr(position), code = 0
        Audit03SetPlotSetting table, "Plot.NeutralLineEnabled", CVErr(2015)
        Audit03SetPlotSetting table, "Plot.ResultLabelsEnabled", CVErr(2015)
        Audit03SetPlotSetting table, "Plot.LegendEnabled", CVErr(2015)
        settings.LoadFromWorkbook fixture: reader.LoadGeometryPreviewFromWorkbook fixture, settings
        Audit03CapturePlotDraw plotter, fixture, reader, settings, code, reason
        AssertTrue stats, "audit03.generalPlot.geometry.inactive.p" & CStr(position), code = 0
        table.Value2 = baseline: annotationTable.Value2 = annotationBaseline: settings.LoadFromWorkbook fixture: reader.LoadFromWorkbook fixture, settings
        For Each key In Array("Plot.NeutralLineEnabled", "Plot.LoadApplicationPointEnabled", "Plot.PrincipalAxesMode")
            Audit03SetPlotSetting table, CStr(key), IIf(CStr(key) = "Plot.PrincipalAxesMode", "None", "No")
            settings.LoadFromWorkbook fixture: plotter.Draw fixture, reader, settings
            Select Case CStr(key)
                Case "Plot.NeutralLineEnabled": AssertTrue stats, "audit03.generalPlot.neutral.off.p" & CStr(position), Not Audit03PlotShapeExists(chart, "NeutralAxis")
                Case "Plot.LoadApplicationPointEnabled": AssertTrue stats, "audit03.generalPlot.load.off.p" & CStr(position), Not Audit03PlotShapeExists(chart, "LoadPoint")
                Case Else: AssertTrue stats, "audit03.generalPlot.axes.off.p" & CStr(position), Not Audit03PlotShapeExists(chart, "Principal1")
            End Select
            Audit03SetPlotSetting table, CStr(key), IIf(CStr(key) = "Plot.PrincipalAxesMode", "Concrete", "Yes")
            settings.LoadFromWorkbook fixture: plotter.Draw fixture, reader, settings
            Select Case CStr(key)
                Case "Plot.NeutralLineEnabled": AssertTrue stats, "audit03.generalPlot.neutral.on.p" & CStr(position), Audit03PlotShapeExists(chart, "NeutralAxis")
                Case "Plot.LoadApplicationPointEnabled": AssertTrue stats, "audit03.generalPlot.load.on.p" & CStr(position), Audit03PlotShapeExists(chart, "LoadPoint")
                Case Else: AssertTrue stats, "audit03.generalPlot.axes.on.p" & CStr(position), Audit03PlotShapeExists(chart, "Principal1")
            End Select
        Next key
        Audit03SetPlotSetting table, "Plot.LegendEnabled", "No"
        Audit03SetPlotSetting table, "Plot.ResultLabelSpacing", 100#
        For Each key In Array("Plot.ContourEnabled", "Plot.ResultLabelsEnabled", "Plot.LegendEnabled")
            Audit03SetPlotSetting table, CStr(key), "No": settings.LoadFromWorkbook fixture: plotter.Draw fixture, reader, settings
            prefix = "audit03.generalPlot.toggle.p" & CStr(position) & "." & CStr(key)
            Select Case CStr(key)
                Case "Plot.ContourEnabled": AssertTrue stats, prefix & ".off", Not Audit03PlotShapeExists(chart, "ContourCircle")
                Case "Plot.ResultLabelsEnabled": AssertTrue stats, prefix & ".off", Not Audit03PlotTextExists(chart, "-1.25")
                Case Else: AssertTrue stats, prefix & ".off", Not Audit03PlotShapeExists(chart, "LegendGradient")
            End Select
            Audit03SetPlotSetting table, CStr(key), "Yes": settings.LoadFromWorkbook fixture: plotter.Draw fixture, reader, settings
            Select Case CStr(key)
                Case "Plot.ContourEnabled": AssertTrue stats, prefix & ".on", Audit03PlotShapeExists(chart, "ContourCircle")
                Case "Plot.ResultLabelsEnabled": AssertTrue stats, prefix & ".on", Audit03PlotTextExists(chart, "-1.25")
                Case Else: AssertTrue stats, prefix & ".on", Audit03PlotShapeExists(chart, "LegendGradient")
            End Select
        Next key
        ' Проверяем реальные подписи и шрифт, включая неактивную высоту текста.
        Audit03SetPlotSetting table, "Plot.AxisLabelsEnabled", "Yes"
        Audit03SetPlotSetting table, "Plot.AxisLabelsFontSize", 18#
        settings.LoadFromWorkbook fixture: plotter.Draw fixture, reader, settings
        Set marker = Audit03PlotShape(chart, "AxisLabelX")
        AssertTrue stats, "audit03.generalPlot.axisLabels.on.p" & CStr(position), Not marker Is Nothing
        If Not marker Is Nothing Then AssertClose stats, "audit03.generalPlot.axisLabels.font.p" & CStr(position), marker.TextFrame.Characters().Font.Size, 18#, 0.01
        Audit03SetPlotSetting table, "Plot.AxisLabelsEnabled", "No"
        Audit03SetPlotSetting table, "Plot.AxisLabelsFontSize", CVErr(2015)
        settings.LoadFromWorkbook fixture: plotter.Draw fixture, reader, settings
        AssertTrue stats, "audit03.generalPlot.axisLabels.off.p" & CStr(position), Not Audit03PlotShapeExists(chart, "AxisLabelX") And Not Audit03PlotShapeExists(chart, "AxisLabelY")
        table.Value2 = baseline: annotationTable.Value2 = annotationBaseline

        ' Все четыре цвета должны менять соответствующий элемент, а не только
        ' поле reader-а. Подменяем исключительно собственный синтетический снимок.
        For Each key In Array("Plot.Color.RebarCompression", "Plot.Color.RebarTension", "Plot.Color.ConcreteCompression", "Plot.Color.ConcreteTension")
            changedElements = elements
            compression = (Right$(CStr(key), 11) = "Compression")
            For row = 2 To 3
                changedElements(row, 5) = IIf(compression, -0.0001, 0.0001)
                changedElements(row, 6) = IIf(compression, -2.75, 2.75)
                changedElements(row, 7) = IIf(compression, "Compression", "Tension")
            Next row
            sheet.Range("R35").Resize(5, 7).Value2 = changedElements
            table.Value2 = baseline: annotationTable.Value2 = annotationBaseline
            Audit03SetPlotSetting table, "Plot.LegendMode", "Separate"
            Audit03SetPlotSetting table, "Plot.ResultGradient", "No"
            Audit03SetPlotSetting table, CStr(key), "17,34,51"
            settings.LoadFromWorkbook fixture: reader.LoadFromWorkbook fixture, settings
            plotter.Draw fixture, reader, settings
            elementKind = IIf(InStr(1, CStr(key), "Rebar", vbTextCompare) > 0, "ElementRebar", "ElementConcrete")
            Set marker = Audit03PlotShape(chart, elementKind)
            prefix = "audit03.generalPlot.color.p" & CStr(position) & "." & CStr(key)
            AssertTrue stats, prefix & ".shape", Not marker Is Nothing
            If Not marker Is Nothing Then AssertTrue stats, prefix & ".effect", marker.Fill.ForeColor.RGB = RGB(17, 34, 51)
            cases = cases + 1
        Next key

        ' Common использует арматурную шкалу и общий максимум по знаку.
        ' Для отношения 1.25/2.75 независимый RGB oracle равен 164,169,173.
        changedElements = elements
        changedElements(3, 5) = -0.0002: changedElements(3, 6) = -2.75: changedElements(3, 7) = "Compression"
        sheet.Range("R35").Resize(5, 7).Value2 = changedElements
        table.Value2 = baseline: annotationTable.Value2 = annotationBaseline
        Audit03SetPlotSetting table, "Plot.ResultGradient", "No"
        Audit03SetPlotSetting table, "Plot.Color.ConcreteCompression", "10,20,30"
        Audit03SetPlotSetting table, "Plot.Color.RebarCompression", "100,110,120"
        For Each key In Array("Separate", "Common")
            Audit03SetPlotSetting table, "Plot.LegendMode", CStr(key)
            settings.LoadFromWorkbook fixture: reader.LoadFromWorkbook fixture, settings
            plotter.Draw fixture, reader, settings
            Set marker = Audit03PlotShape(chart, "ElementConcrete")
            expectedColor = IIf(CStr(key) = "Separate", RGB(10, 20, 30), RGB(100, 110, 120))
            AssertTrue stats, "audit03.generalPlot.legendMode.p" & CStr(position) & "." & CStr(key) & ".shape", Not marker Is Nothing
            If Not marker Is Nothing Then AssertTrue stats, "audit03.generalPlot.legendMode.p" & CStr(position) & "." & CStr(key) & ".effect", marker.Fill.ForeColor.RGB = expectedColor
        Next key
        Audit03SetPlotSetting table, "Plot.ResultGradient", "Yes"
        settings.LoadFromWorkbook fixture: plotter.Draw fixture, reader, settings
        Set marker = Audit03PlotShape(chart, "ElementConcrete")
        AssertTrue stats, "audit03.generalPlot.gradient.p" & CStr(position) & ".shape", Not marker Is Nothing
        If Not marker Is Nothing Then AssertTrue stats, "audit03.generalPlot.gradient.p" & CStr(position) & ".effect", marker.Fill.ForeColor.RGB = RGB(164, 169, 173)
        sheet.Range("R35").Resize(5, 7).Value2 = elements
        table.Value2 = baseline: annotationTable.Value2 = annotationBaseline
        Audit03SetPlotSetting table, "Plot.LoadCase", "ALT"
        Audit03SetPlotSetting table, "Plot.ResultLabelsEnabled", "Yes"
        Audit03SetPlotSetting table, "Plot.ResultLabelSpacing", 100#
        settings.LoadFromWorkbook fixture: reader.LoadFromWorkbook fixture, settings
        plotter.Draw fixture, reader, settings
        AssertTrue stats, "audit03.generalPlot.loadCase.changed.p" & CStr(position), reader.LoadCase = "ALT"
        AssertClose stats, "audit03.generalPlot.loadCase.value.p" & CStr(position), reader.ResultValue(1), 9.5, 0#
        ' Точность Stress в fixture равна двум знакам, поэтому ноль сохраняется.
        AssertTrue stats, "audit03.generalPlot.loadCase.label.p" & CStr(position), Audit03PlotTextExists(chart, "9.50")
        table.Value2 = baseline: annotationTable.Value2 = annotationBaseline
        settings.LoadFromWorkbook fixture: reader.LoadFromWorkbook fixture, settings
        Audit03SetPlotSetting table, "Plot.ResultLabelsEnabled", "Yes"
        Audit03SetPlotSetting table, "Plot.ResultLabelSpacing", 100#
        For Each key In keys
            Set cell = Audit03PlotSettingCell(table, CStr(key))
            cases = cases + 1
            cell.Offset(0, -1).Value2 = "REMOVED_" & CStr(key): settings.LoadFromWorkbook fixture
            Audit03CapturePlotDraw plotter, fixture, reader, settings, code, reason
            AssertTrue stats, "audit03.generalPlot.missing.p" & CStr(position) & "." & CStr(key), code <> 0 And InStr(1, reason, CStr(key), vbTextCompare) > 0
            cell.Offset(0, -1).Value2 = key
        Next key
        table.Value2 = baseline: annotationTable.Value2 = annotationBaseline
        Audit03SetPlotSetting table, "Plot.Dimensions.Enabled", "Yes"
        For Each key In Array("Plot.Dimensions.ArrowType", "Plot.Dimensions.ArrowSize")
            Set cell = Audit03PlotSettingCell(table, CStr(key)): values = cell.Value2: badIndex = 0
            For Each bad In Array("", "TODO", "INVALID", CVErr(2015))
                badIndex = badIndex + 1: cases = cases + 1: cell.Value2 = bad: settings.LoadFromWorkbook fixture
                Set marker = chart.Chart.Shapes.AddShape(1, 11#, 17#, 20#, 20#): marker.Name = "NDMPlot_Audit03Sentinel"
                Audit03CapturePlotDraw plotter, fixture, reader, settings, code, reason
                prefix = "audit03.generalPlot.arrow.p" & CStr(position) & "." & CStr(key) & ".bad" & CStr(badIndex)
                AssertTrue stats, prefix & ".rejected", code <> 0
                AssertTrue stats, prefix & ".address", InStr(1, reason, config.Name, vbTextCompare) > 0 And InStr(1, reason, cell.Address(False, False), vbTextCompare) > 0
                AssertTrue stats, prefix & ".preserved", Audit03PlotShapeExists(chart, "Audit03Sentinel")
                chart.Chart.Shapes.Item("NDMPlot_Audit03Sentinel").Delete
                cell.Value2 = values
            Next bad
        Next key
        For mode = 0 To 2
            For Each bad In Array("", "TODO", CVErr(2015))
                cases = cases + 1: Audit03SetPlotSetting table, "Plot.LoadCase", bad: settings.LoadFromWorkbook fixture
                Audit03ReadLifecycleLoad reader, fixture, settings, mode, code, reason
                prefix = "audit03.generalPlot.loadCase.p" & CStr(position) & ".mode" & CStr(mode) & ".type" & CStr(VarType(bad))
                If mode = 1 Then
                    AssertTrue stats, prefix & ".inactive", code = 0 And reader.Count = 2
                Else
                    Set cell = Audit03PlotSettingCell(table, "Plot.LoadCase")
                    AssertTrue stats, prefix & ".rejected", code <> 0 And InStr(1, reason, "Plot.LoadCase", vbTextCompare) > 0
                    AssertTrue stats, prefix & ".address", InStr(1, reason, config.Name, vbTextCompare) > 0 And InStr(1, reason, cell.Address(False, False), vbTextCompare) > 0
                    AssertTrue stats, prefix & ".empty", reader.Count = 0
                End If
            Next bad
        Next mode
        table.Value2 = baseline: annotationTable.Value2 = annotationBaseline: settings.LoadFromWorkbook fixture: reader.LoadFromWorkbook fixture, settings
    Next position
    Audit03SetPlotSetting table, "Plot.Dimensions.Enabled", "Yes"
    For Each arrow In Array("Triangle", "Stealth", "Diamond", "Oval", "Open")
        Select Case CStr(arrow)
            Case "Triangle": style = 2
            Case "Stealth": style = 3
            Case "Diamond": style = 4
            Case "Oval": style = 5
            Case "Open": style = 6
        End Select
        For Each size In Array("Small", "Medium", "Wide")
            Select Case CStr(size)
                Case "Small": metric = 1
                Case "Medium": metric = 2
                Case "Wide": metric = 3
            End Select
            cases = cases + 1
            Audit03SetPlotSetting table, "Plot.Dimensions.ArrowType", arrow
            Audit03SetPlotSetting table, "Plot.Dimensions.ArrowSize", size
            settings.LoadFromWorkbook fixture: plotter.Draw fixture, reader, settings
            Set marker = Nothing
            For Each shape In chart.Chart.Shapes
                If InStr(1, shape.Name, "AnnotationLine", vbTextCompare) > 0 Then
                    If shape.Line.BeginArrowheadStyle > 1 Then Set marker = shape: Exit For
                End If
            Next shape
            prefix = "audit03.generalPlot.arrows." & CStr(arrow) & "." & CStr(size)
            AssertTrue stats, prefix & ".line", Not marker Is Nothing
            If Not marker Is Nothing Then
                AssertTrue stats, prefix & ".style", marker.Line.BeginArrowheadStyle = style And marker.Line.EndArrowheadStyle = style
                AssertTrue stats, prefix & ".length", marker.Line.BeginArrowheadLength = metric And marker.Line.EndArrowheadLength = metric
                AssertTrue stats, prefix & ".width", marker.Line.BeginArrowheadWidth = metric And marker.Line.EndArrowheadWidth = metric
            End If
        Next size
    Next arrow
    AssertClose stats, "audit03.generalPlot.chart.left", chart.Left, 41#, 0.01
    AssertClose stats, "audit03.generalPlot.chart.top", chart.Top, 53#, 0.01
    AssertClose stats, "audit03.generalPlot.chart.width", chart.Width, 700#, 0.01
    AssertClose stats, "audit03.generalPlot.chart.height", chart.Height, 480#, 0.01
    actual = sheet.Range("A5:AI39").Value2
    Audit03ComparePlainSnapshot stats, "audit03.generalPlot.ResultsUnchanged", snapshot, actual, 1
    ' На сетке >5000 точек проверяем все 17 уровней градиента. Округление
    ' номера bucket не должно пропускать нечетные целые значения.
    geometry = Audit03GeometrySnapshotArray("mm", 1#)
    ReDim grid(1 To 5002, 1 To 15): ReDim gridResults(1 To 5002, 1 To 7)
    For index = 1 To 15: grid(1, index) = geometry(1, index): Next index
    For index = 1 To 7: gridResults(1, index) = elements(1, index): Next index
    For row = 2 To 5002
        For index = 1 To 15: grid(row, index) = geometry(2, index): Next index
        grid(row, 1) = "C" & CStr(row - 1): grid(row, 3) = -90# + (row Mod 19) * 10#: grid(row, 4) = -90# + (row Mod 17) * 10#
        grid(row, 5) = 100#: grid(row, 8) = 10#: grid(row, 9) = 10#: grid(row, 11) = 0#
        For index = 1 To 7: gridResults(row, index) = elements(2, index): Next index
        gridResults(row, 4) = grid(row, 1): gridResults(row, 6) = -CDbl((row - 2) Mod 17)
    Next row
    sheet.Range("AF5").Resize(5002, 15).Value2 = grid
    sheet.Range("AV5").Resize(5002, 7).Value2 = gridResults
    fixture.Names.Item("rngNDMSectionGeometry").RefersTo = "=Results!$AF$5"
    fixture.Names.Item("rngNDMElementResults").RefersTo = "=Results!$AV$5"
    table.Value2 = baseline: annotationTable.Value2 = annotationBaseline
    Audit03SetPlotSetting table, "Plot.Color.ConcreteCompression", "0,0,0"
    Audit03SetPlotSetting table, "Plot.ResultGradient", "Yes"
    settings.LoadFromWorkbook fixture: reader.LoadFromWorkbook fixture, settings
    plotter.Draw fixture, reader, settings
    Set colors = CreateObject("Scripting.Dictionary")
    For Each series In chart.Chart.SeriesCollection
        colors(CStr(series.MarkerBackgroundColor)) = True
    Next series
    AssertTrue stats, "audit03.generalPlot.grouped.colorCount", colors.Count = 17
    For index = 0 To 16
        channel = CLng(218# * (1# - CDbl(index) / 16#))
        AssertTrue stats, "audit03.generalPlot.grouped.color" & CStr(index), colors.Exists(CStr(RGB(channel, channel, channel)))
    Next index
    AssertTrue stats, "audit03.generalPlot.noSolve", SectionEquilibriumSolveCount() = solveCount
    AppendLine stats, "GENERAL_PLOT_CASES: variants=" & CStr(cases) & "; geometryFixtures=2; equilibriumCases=0"
    GoTo Cleanup
Failed:
    AssertTrue stats, "audit03.generalPlot.runtime." & CStr(Err.Number) & "." & Err.Description, False
Cleanup:
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
End Sub

' Объединяет существующие System/Annotation строки в отдельный стандартный
' Range. Для layout используются явно заданные независимые контрольные
' числа; они не меняют defaults каталога или исходный Config пользователя.
Private Function Audit03PlotSettingsTable(ByVal anchor As Object) As Object
    Dim system As Variant, keys As Variant, key As Variant, data() As Variant, row As Long, target As Long
    Dim settings As CSystemSettingsReader
    system = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Value2
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    keys = Array("Plot.Dimensions.Enabled", "Plot.Dimensions.TextHeight", "Plot.Dimensions.TextUnits", _
        "Plot.Dimensions.Offset", "Plot.Dimensions.LineWeight", "Plot.Dimensions.ExtensionLineWeight", _
        "Plot.Dimensions.TextGap", "Plot.Dimensions.Placement", "Plot.Dimensions.Color", "Plot.Dimensions.ExtensionLineColor", _
        "Plot.Dimensions.ArrowType", "Plot.Dimensions.ArrowSize", "Plot.RebarLabels.Enabled", _
        "Plot.RebarLabels.TextHeight", "Plot.RebarLabels.TextUnits", "Plot.RebarLabels.Offset", _
        "Plot.RebarLabels.LineEnabled", "Plot.RebarLabels.LineWeight", "Plot.RebarLabels.TextGap", _
        "Plot.RebarLabels.Placement", "Plot.RebarLabels.Color")
    ReDim data(1 To UBound(system, 1) + UBound(keys) + 1, 1 To 3)
    data(1, 1) = "Параметр": data(1, 2) = "Значение": data(1, 3) = "Комментарий": target = 1
    For row = 2 To UBound(system, 1)
        target = target + 1: data(target, 1) = system(row, 1): data(target, 2) = system(row, 2)
    Next row
    For Each key In keys
        target = target + 1: data(target, 1) = key: data(target, 2) = settings.GetString(CStr(key))
    Next key
    Set Audit03PlotSettingsTable = anchor.Resize(target, 3)
    Audit03PlotSettingsTable.NumberFormat = "@"
    Audit03PlotSettingsTable.Value2 = data
    Audit03SetPlotSetting Audit03PlotSettingsTable, "Plot.LoadCase", "PLOT"
    Audit03SetPlotSetting Audit03PlotSettingsTable, "Plot.Dimensions.Enabled", "No"
    Audit03SetPlotSetting Audit03PlotSettingsTable, "Plot.RebarLabels.Enabled", "No"
    Audit03SetPlotSetting Audit03PlotSettingsTable, "Plot.ResultLabelsEnabled", "No"
    Audit03SetPlotSetting Audit03PlotSettingsTable, "Plot.LegendEnabled", "No"
End Function

' Копирует настоящие Config-таблицы в собственный presentation fixture.
' Системные настройки не сворачиваются в трехколоночный словарь, а группы
' аннотаций сохраняют исходный компактный формат и свои реальные имена.
' Начальные значения управляют только тестовой схемой, не defaults книги.
Private Function Audit03NamedGeneralPlotSettingsTable(ByVal fixture As Object, ByVal sheet As Object) As Object
    Dim name As Variant, addresses As Variant, source As Object, target As Object, index As Long
    addresses = Array("A5", "W5", "W20", "W30")
    For Each name In Array("rngSystemSettings", "rngUnitSettings", "rngSignConventionSettings", "rngPlotAnnotationSettings")
        Set source = ThisWorkbook.Names.Item(CStr(name)).RefersToRange
        Set target = sheet.Range(CStr(addresses(index))).Resize(source.Rows.Count, source.Columns.Count)
        target.NumberFormat = "@": target.Value2 = source.Value2
        fixture.Names.Add Name:=CStr(name), RefersTo:="='" & sheet.Name & "'!" & target.Address
        index = index + 1
    Next name
    Set Audit03NamedGeneralPlotSettingsTable = fixture.Names.Item("rngSystemSettings").RefersToRange
    Audit03SetPlotSetting Audit03NamedGeneralPlotSettingsTable, "Plot.LoadCase", "PLOT"
    Audit03SetPlotSetting Audit03NamedGeneralPlotSettingsTable, "Plot.Dimensions.Enabled", "No"
    Audit03SetPlotSetting Audit03NamedGeneralPlotSettingsTable, "Plot.RebarLabels.Enabled", "No"
    Audit03SetPlotSetting Audit03NamedGeneralPlotSettingsTable, "Plot.ResultLabelsEnabled", "No"
    Audit03SetPlotSetting Audit03NamedGeneralPlotSettingsTable, "Plot.LegendEnabled", "No"
    Audit03SetPlotSetting Audit03NamedGeneralPlotSettingsTable, "Plot.LegendMode", "Separate"
    Audit03SetPlotSetting Audit03NamedGeneralPlotSettingsTable, "Plot.AxisLabelsEnabled", "Yes"
    Audit03SetPlotSetting Audit03NamedGeneralPlotSettingsTable, "Plot.AxisLabelsFontSize", 11#
End Function

' Возвращает именно текущую ячейку значения, а не жестко заданный адрес.
' В полном Config группы аннотаций находятся в отдельной компактной таблице;
' у автономного scalar fixture сначала используется его собственный ключ.
' Отсутствующий test-key является ошибкой fixture, а не успешной проверкой.
Private Function Audit03PlotSettingCell(ByVal table As Object, ByVal key As String) As Object
    Dim row As Long
    For row = 2 To table.Rows.Count
        If CStr(table.Cells(row, 1).Value2) = key Then
            Set Audit03PlotSettingCell = table.Cells(row, 2): Exit Function
        End If
    Next row
    If Left$(key, 17) = "Plot.RebarLabels." Or Left$(key, 16) = "Plot.Dimensions." Then
        Set Audit03PlotSettingCell = Audit03NamedAnnotationCell(table.Worksheet.Parent.Names.Item("rngPlotAnnotationSettings").RefersToRange, key)
        Exit Function
    End If
    Err.Raise 5, "Audit03PlotSettingCell", "В fixture нет параметра " & key
End Function

' Меняет одну ячейку только внутри собственной временной таблицы теста.
Private Sub Audit03SetPlotSetting(ByVal table As Object, ByVal key As String, ByVal value As Variant)
    Audit03PlotSettingCell(table, key).Value2 = value
End Sub

' Сохраняет исходную VBA-ошибку Draw для проверки сообщения и восстановления.
Private Sub Audit03CapturePlotDraw(ByVal plotter As CSectionPlotter, ByVal workbook As Object, _
        ByVal reader As CSectionPlotDataReader, ByVal settings As CSystemSettingsReader, ByRef code As Long, ByRef reason As String)
    code = 0: reason = vbNullString
    On Error GoTo Failed
    plotter.Draw workbook, reader, settings
    Exit Sub
Failed:
    code = Err.Number: reason = Err.Description: Err.Clear
End Sub

' Ищет фактический Shape собственного Chart; не обращается к основной книге.
Private Function Audit03PlotShapeExists(ByVal chart As Object, ByVal fragment As String) As Boolean
    Dim shape As Object
    For Each shape In chart.Chart.Shapes
        If InStr(1, shape.Name, fragment, vbTextCompare) > 0 Then Audit03PlotShapeExists = True: Exit Function
    Next shape
End Function

' Возвращает фактическую фигуру для проверки ее цвета или шрифта. Отсутствие
' ожидаемой фигуры остается отдельным провалом assertion, без unsafe dereference.
Private Function Audit03PlotShape(ByVal chart As Object, ByVal fragment As String) As Object
    Dim shape As Object
    For Each shape In chart.Chart.Shapes
        If InStr(1, shape.Name, fragment, vbTextCompare) > 0 Then Set Audit03PlotShape = shape: Exit Function
    Next shape
End Function

' Сверяет фактический текст chart-подписи с независимым числом snapshot.
Private Function Audit03PlotTextExists(ByVal chart As Object, ByVal expected As String) As Boolean
    Dim shape As Object, text As String
    For Each shape In chart.Chart.Shapes
        text = vbNullString
        On Error Resume Next
        text = shape.TextFrame.Characters().Text
        On Error GoTo 0
        If Replace(text, ",", ".") = expected Then Audit03PlotTextExists = True: Exit Function
    Next shape
End Function

' Проверяет координаты, шрифты, цвета и толщины при мм и pt, внутри и снаружи
' сечения, с отключённой линией арматуры. Эталон в пунктах вычисляется отдельно
' при масштабах 3.2 по X и 2.15 по Y без вызова отрисовки.
Private Sub TestAudit03AnnotationLayoutContracts(ByRef stats As TUiTestStats)
    Dim fixture As Object, sheet As Object, table As Object, settings As CSystemSettingsReader
    Dim layout As CPlotAnnotationLayout, baseline As Variant, key As Variant, bad As Variant, cell As Object
    Dim group As Variant, keys As Variant, index As Long, row As Long, position As Long, code As Long, reason As String, prefix As String
    Dim cases As Long, units As Variant, placement As Variant, sign As Double, font As Double, gap As Double, solveCount As Long
    Dim mutationKeys As Variant, changed As Variant, expected As Variant, item As Variant, metric As Variant, measured As Double
    On Error GoTo Failed
    solveCount = SectionEquilibriumSolveCount()
    Set fixture = Application.Workbooks.Add(-4167): Set sheet = fixture.Worksheets(1): sheet.Name = "LayoutConfig"
    Set table = Audit03PlotSettingsTable(sheet.Range("A5"))
    For Each group In Array("Dimensions", "RebarLabels")
        Audit03SetPlotSetting table, "Plot." & CStr(group) & ".Enabled", "Yes"
        Audit03SetPlotSetting table, "Plot." & CStr(group) & ".TextHeight", 6#
        Audit03SetPlotSetting table, "Plot." & CStr(group) & ".TextGap", 3#
        Audit03SetPlotSetting table, "Plot." & CStr(group) & ".Offset", 10#
        Audit03SetPlotSetting table, "Plot." & CStr(group) & ".LineWeight", 1.5
        Audit03SetPlotSetting table, "Plot." & CStr(group) & ".Color", "0,0,0"
        Audit03SetPlotSetting table, "Plot." & CStr(group) & ".TextUnits", "mm"
        Audit03SetPlotSetting table, "Plot." & CStr(group) & ".Placement", "Outside"
    Next group
    Audit03SetPlotSetting table, "Plot.Dimensions.ExtensionLineWeight", 0.75
    Audit03SetPlotSetting table, "Plot.Dimensions.ExtensionLineColor", "1,2,3"
    Audit03SetPlotSetting table, "Plot.RebarLabels.LineEnabled", "Yes"
    baseline = table.Value2: Set settings = New CSystemSettingsReader: Set layout = New CPlotAnnotationLayout
    keys = Array("Plot.Dimensions.Enabled", "Plot.Dimensions.TextHeight", "Plot.Dimensions.TextUnits", _
        "Plot.Dimensions.Offset", "Plot.Dimensions.LineWeight", "Plot.Dimensions.ExtensionLineWeight", _
        "Plot.Dimensions.TextGap", "Plot.Dimensions.Placement", "Plot.Dimensions.Color", "Plot.Dimensions.ExtensionLineColor", _
        "Plot.RebarLabels.Enabled", "Plot.RebarLabels.TextHeight", "Plot.RebarLabels.TextUnits", _
        "Plot.RebarLabels.Offset", "Plot.RebarLabels.LineEnabled", "Plot.RebarLabels.LineWeight", _
        "Plot.RebarLabels.TextGap", "Plot.RebarLabels.Placement", "Plot.RebarLabels.Color")
    For position = 0 To 1
        If position = 0 Then Set table = sheet.Range("A5").Resize(UBound(baseline, 1), 3) Else Set table = sheet.Range("CH800").Resize(UBound(baseline, 1), 3)
        table.NumberFormat = "@": table.Value2 = baseline
        For Each key In keys
            Set cell = Audit03PlotSettingCell(table, CStr(key))
            For Each bad In Array("", "TODO", "INVALID", CVErr(2015))
                cases = cases + 1: cell.Value2 = bad: settings.LoadFromRange table
                On Error Resume Next
                Err.Clear: layout.ConfigureFromSettings settings
                code = Err.Number: reason = Err.Description
                On Error GoTo Failed
                prefix = "audit03.annotation.p" & CStr(position) & "." & CStr(key) & ".bad" & CStr(VarType(bad))
                AssertTrue stats, prefix & ".rejected", code <> 0
                AssertTrue stats, prefix & ".key", InStr(1, reason, CStr(key), vbTextCompare) > 0
                AssertTrue stats, prefix & ".address", InStr(1, reason, sheet.Name, vbTextCompare) > 0 And InStr(1, reason, cell.Address(False, False), vbTextCompare) > 0
                AppendLine stats, "ANNOTATION_ERROR: " & prefix & "|" & reason
                table.Value2 = baseline: settings.LoadFromRange table: layout.ConfigureFromSettings settings
                layout.Initialize -100#, 100#, -100#, 100#, 0#, 0#, 640#, 430#, "mm"
                layout.AddDimension -50#, 0#, 50#, 0#, 0#, 1#, "DIM"
                layout.AddRebarLabel -50#, 0#, 50#, 0#, 0#, 1#, "REBAR"
                AssertTrue stats, prefix & ".recovery", layout.Count = 6
            Next bad
        Next key
        mutationKeys = Array("Plot.Dimensions.TextHeight", "Plot.RebarLabels.TextHeight", _
            "Plot.Dimensions.Offset", "Plot.RebarLabels.Offset", "Plot.Dimensions.LineWeight", _
            "Plot.Dimensions.ExtensionLineWeight", "Plot.RebarLabels.LineWeight", "Plot.Dimensions.TextGap", "Plot.RebarLabels.TextGap")
        changed = Array(8#, 8#, 20#, 20#, 2.5, 1.25, 2.5, 5#, 5#)
        expected = Array(17.2, 17.2, 172#, 172#, 2.5, 1.25, 2.5, 182.75, 182.75)
        item = Array(4, 6, 3, 5, 3, 1, 5, 4, 6)
        metric = Array("Font", "Font", "Y", "Y", "Weight", "Weight", "Weight", "CenterY", "CenterY")
        For index = 0 To UBound(mutationKeys)
            cases = cases + 1
            table.Value2 = baseline
            Audit03SetPlotSetting table, CStr(mutationKeys(index)), changed(index)
            settings.LoadFromRange table: layout.ConfigureFromSettings settings
            layout.Initialize -100#, 100#, -100#, 100#, 0#, 0#, 640#, 430#, "mm"
            layout.AddDimension -50#, 0#, 50#, 0#, 0#, 1#, "DIM"
            layout.AddRebarLabel -50#, 0#, 50#, 0#, 0#, 1#, "REBAR"
            Select Case CStr(metric(index))
                Case "Font": measured = layout.FontSize(CLng(item(index)))
                Case "Y": measured = layout.Y1(CLng(item(index)))
                Case "Weight": measured = layout.Weight(CLng(item(index)))
                Case "CenterY": measured = layout.Y1(CLng(item(index))) + layout.Y2(CLng(item(index))) / 2#
            End Select
            AssertClose stats, "audit03.annotation.mutation.p" & CStr(position) & "." & CStr(mutationKeys(index)), measured, CDbl(expected(index)), 0.000000001
            For Each bad In Array(-1#, 0#)
                cases = cases + 1
                Audit03SetPlotSetting table, CStr(mutationKeys(index)), bad: settings.LoadFromRange table
                On Error Resume Next
                Err.Clear: layout.ConfigureFromSettings settings: code = Err.Number: reason = Err.Description
                On Error GoTo Failed
                prefix = "audit03.annotation.sign.p" & CStr(position) & "." & CStr(mutationKeys(index)) & "." & CStr(bad)
                If CDbl(bad) = 0# And (CStr(metric(index)) = "Y" Or CStr(metric(index)) = "CenterY") Then
                    AssertTrue stats, prefix & ".zeroAllowed", code = 0
                Else
                    Set cell = Audit03PlotSettingCell(table, CStr(mutationKeys(index)))
                    AssertTrue stats, prefix & ".rejected", code <> 0 And InStr(1, reason, CStr(mutationKeys(index)), vbTextCompare) > 0
                    AssertTrue stats, prefix & ".address", InStr(1, reason, cell.Address(False, False), vbTextCompare) > 0
                End If
            Next bad
        Next index
        table.Value2 = baseline
        For Each units In Array("mm", "pt")
            For Each placement In Array("Outside", "Inside")
                If CStr(placement) = "Outside" Then sign = -1# Else sign = 1#
                If CStr(units) = "mm" Then font = 12.9: gap = 6.45 Else font = 6#: gap = 3#
                For Each group In Array("Dimensions", "RebarLabels")
                    Audit03SetPlotSetting table, "Plot." & CStr(group) & ".TextUnits", units
                    Audit03SetPlotSetting table, "Plot." & CStr(group) & ".Placement", placement
                Next group
                settings.LoadFromRange table: layout.ConfigureFromSettings settings
                layout.Initialize -100#, 100#, -100#, 100#, 0#, 0#, 640#, 430#, "mm"
                layout.AddDimension -50#, 0#, 50#, 0#, 0#, 1#, "DIM"
                layout.AddRebarLabel -50#, 0#, 50#, 0#, 0#, 1#, "REBAR"
                prefix = "audit03.annotation.p" & CStr(position) & "." & CStr(units) & "." & CStr(placement)
                AssertTrue stats, prefix & ".count", layout.Count = 6
                AssertClose stats, prefix & ".dimension.x", layout.X1(3), 160#, 0.000000001
                AssertClose stats, prefix & ".dimension.y", layout.Y1(3), 193.5, 0.000000001
                AssertClose stats, prefix & ".dimension.end", layout.X2(3), 480#, 0.000000001
                AssertClose stats, prefix & ".font", layout.FontSize(4), font, 0.000000001
                AssertClose stats, prefix & ".textPlacement", layout.Y1(4) + layout.Y2(4) / 2#, 193.5 + sign * gap, 0.000000001
                AssertClose stats, prefix & ".rebarPlacement", layout.Y1(6) + layout.Y2(6) / 2#, 215# + sign * (21.5 + gap), 0.000000001
                AssertTrue stats, prefix & ".black", layout.Color(3) = 0 And layout.Color(4) = 0 And layout.Color(5) = 0
                AssertTrue stats, prefix & ".extensionColor", layout.Color(1) = RGB(1, 2, 3)
                AssertClose stats, prefix & ".lineWeight", layout.Weight(3), 1.5, 0#
                AssertClose stats, prefix & ".extensionWeight", layout.Weight(1), 0.75, 0#
                AssertClose stats, prefix & ".rebarWeight", layout.Weight(5), 1.5, 0#
                AssertClose stats, prefix & ".rebarFont", layout.FontSize(6), font, 0.000000001
                AssertTrue stats, prefix & ".arrows", layout.Arrowheads(3) And Not layout.Arrowheads(1) And Not layout.Arrowheads(5)
            Next placement
        Next units
        Audit03SetPlotSetting table, "Plot.RebarLabels.LineEnabled", "No"
        Audit03SetPlotSetting table, "Plot.RebarLabels.LineWeight", CVErr(2015)
        settings.LoadFromRange table: layout.ConfigureFromSettings settings
        layout.Initialize -100#, 100#, -100#, 100#, 0#, 0#, 640#, 430#, "mm"
        layout.AddRebarLabel -50#, 0#, 50#, 0#, 0#, 1#, "REBAR"
        AssertTrue stats, "audit03.annotation.noLeader.p" & CStr(position), layout.Count = 1 And layout.ItemKind(1) = "TEXT"
        For Each group In Array("Dimensions", "RebarLabels")
            Audit03SetPlotSetting table, "Plot." & CStr(group) & ".Enabled", "No"
        Next group
        For Each key In keys
            If Right$(CStr(key), 8) <> ".Enabled" Then Audit03SetPlotSetting table, CStr(key), CVErr(2015)
        Next key
        settings.LoadFromRange table: layout.ConfigureFromSettings settings
        layout.Initialize -100#, 100#, -100#, 100#, 0#, 0#, 640#, 430#, "mm"
        layout.AddDimension -50#, 0#, 50#, 0#, 0#, 1#, "DIM"
        layout.AddRebarLabel -50#, 0#, 50#, 0#, 0#, 1#, "REBAR"
        AssertTrue stats, "audit03.annotation.groupsOff.p" & CStr(position), layout.Count = 0
    Next position
    AssertTrue stats, "audit03.annotation.noSolve", SectionEquilibriumSolveCount() = solveCount
    AppendLine stats, "ANNOTATION_CASES: variants=" & CStr(cases) & "; layoutFixtures=1; equilibriumCases=0"
    GoTo Cleanup
Failed:
    AssertTrue stats, "audit03.annotation.runtime." & CStr(Err.Number) & "." & Err.Description, False
Cleanup:
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
End Sub

' Проверяет обычное обновление и предварительный просмотр импорта без решения НДС.
' Общее отключение схемы не требует исправления неиспользуемых параметров оформления.
' Ошибочное значение самого переключателя должно давать причину в Config.
Private Sub TestAudit03PlotEnableContracts(ByRef stats As TUiTestStats)
    Dim fixture As Object, config As Object, sheet As Object, source As Object, table As Object, target As Object
    Dim name As Variant, addresses As Variant, index As Long, position As Long, mode As Long, bad As Variant
    Dim baseline As Variant, chart As Object, marker As Object, settings As CSystemSettingsReader
    Dim code As Long, reason As String, prefix As String, solveCount As Long, props(1 To 2, 1 To 4) As Variant
    On Error GoTo Failed
    solveCount = SectionEquilibriumSolveCount()
    Set fixture = Application.Workbooks.Add(-4167): Set config = fixture.Worksheets(1): config.Name = "Config"
    Set sheet = fixture.Worksheets.Add: sheet.Name = "Results"
    fixture.Worksheets.Add.Name = "Расчет"
    addresses = Array("A5", "H5", "P5", "P25"): index = 0
    For Each name In Array("rngSystemSettings", "rngPlotAnnotationSettings", "rngUnitSettings", "rngSignConventionSettings")
        Set source = ThisWorkbook.Names.Item(CStr(name)).RefersToRange
        Set target = config.Range(CStr(addresses(index))).Resize(source.Rows.Count, source.Columns.Count)
        target.NumberFormat = "@": target.Value2 = source.Value2
        fixture.Names.Add Name:=CStr(name), RefersTo:="=Config!" & target.Address
        index = index + 1
    Next name
    Set table = fixture.Names.Item("rngSystemSettings").RefersToRange: baseline = table.Value2
    sheet.Range("A5").Resize(3, 15).Value2 = Audit03GeometrySnapshotArray("mm", 1#)
    fixture.Names.Add Name:="rngNDMSectionGeometry", RefersTo:="=Results!$A$5"
    Audit03PrepareContourFixture fixture
    props(1, 1) = "LoadCase": props(1, 2) = "Parameter": props(1, 3) = "Value": props(1, 4) = "Unit"
    props(2, 1) = "ALL": props(2, 2) = "Output.LengthUnit": props(2, 3) = "mm": props(2, 4) = "-"
    sheet.Range("R5").Resize(2, 4).Value2 = props
    fixture.Names.Add Name:="rngNDMSectionProperties", RefersTo:="=Results!$R$5"
    Set chart = fixture.Worksheets.Item("Расчет").ChartObjects.Add(41#, 53#, 700#, 480#): chart.Name = "chtNDMSectionPlot"
    Set marker = chart.Chart.Shapes.AddShape(1, 11#, 17#, 20#, 20#): marker.Name = "NDMPlot_Audit03Sentinel"
    For position = 0 To 1
        If position = 0 Then Set table = config.Range("A5").Resize(UBound(baseline, 1), UBound(baseline, 2)) Else Set table = config.Range("CH800").Resize(UBound(baseline, 1), UBound(baseline, 2))
        table.NumberFormat = "@": table.Value2 = baseline: fixture.Names.Item("rngSystemSettings").RefersTo = "=Config!" & table.Address
        Audit03SetPlotSetting table, "Plot.Enabled", "No"
        Audit03SetPlotSetting table, "Plot.ContourEnabled", CVErr(2015)
        For mode = 0 To 1
            On Error Resume Next
            Err.Clear
            If mode = 0 Then UpdateSectionPlotForWorkbook fixture Else UpdateSectionGeometryPreviewForWorkbook fixture
            code = Err.Number: reason = Err.Description
            On Error GoTo Failed
            prefix = "audit03.generalPlot.enable.p" & CStr(position) & ".mode" & CStr(mode)
            AssertTrue stats, prefix & ".disabled", code = 0
            AssertTrue stats, prefix & ".preserved", Audit03PlotShapeExists(chart, "Audit03Sentinel")
            AppendLine stats, "PLOT_DISABLED_ENTRYPOINT: " & prefix & "|" & reason
        Next mode
        table.Value2 = baseline
        For Each bad In Array("", "TODO", "INVALID", CVErr(2015))
            Audit03SetPlotSetting table, "Plot.Enabled", bad
            For mode = 0 To 1
                On Error Resume Next
                Err.Clear
                If mode = 0 Then UpdateSectionPlotForWorkbook fixture Else UpdateSectionGeometryPreviewForWorkbook fixture
                code = Err.Number: reason = Err.Description
                On Error GoTo Failed
                prefix = "audit03.generalPlot.enable.invalid.p" & CStr(position) & ".mode" & CStr(mode) & ".type" & CStr(VarType(bad))
                Set target = Audit03PlotSettingCell(table, "Plot.Enabled")
                AssertTrue stats, prefix & ".rejected", code <> 0 And InStr(1, reason, "Plot.Enabled", vbTextCompare) > 0
                AssertTrue stats, prefix & ".address", InStr(1, reason, "Config", vbTextCompare) > 0 And InStr(1, reason, target.Address(False, False), vbTextCompare) > 0
            Next mode
        Next bad
    Next position
    table.Value2 = baseline: Audit03SetPlotSetting table, "Plot.Enabled", "Yes"
    UpdateSectionGeometryPreviewForWorkbook fixture
    AssertTrue stats, "audit03.generalPlot.enable.recovery", Audit03PlotShapeExists(chart, "ElementConcrete") And Not Audit03PlotShapeExists(chart, "Audit03Sentinel")
    AssertTrue stats, "audit03.generalPlot.enable.noSolve", SectionEquilibriumSolveCount() = solveCount
    GoTo Cleanup
Failed:
    AssertTrue stats, "audit03.generalPlot.enable.runtime." & CStr(Err.Number) & "." & Err.Description, False
Cleanup:
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
End Sub

' Создаёт готовую тестовую геометрию с запретом изменений без сборщика и решателя.
' Дробные координаты и размеры изображений не зависят от текущих единиц ввода и вывода.
Private Function Audit03GeometrySnapshotArray(ByVal unit As String, ByVal factor As Double) As Variant
    Dim data(1 To 3, 1 To 15) As Variant, headers As Variant, column As Long
    headers = Array("ElementID", "MaterialType", "X, " & unit, "Y, " & unit, "Area, " & unit & "2", _
        "MaterialID", "GeometryInterpretationStatus", "Width, " & unit, "Height, " & unit, _
        "Diameter, " & unit, "Rotation, rad", "LocalIx, " & unit & "4", "LocalIy, " & unit & "4", "LocalIxy, " & unit & "4", "Comment")
    For column = 1 To 15: data(1, column) = headers(column - 1): Next column
    data(2, 1) = "C1": data(2, 2) = "Concrete": data(2, 3) = -23.125 / factor: data(2, 4) = 17.625 / factor
    data(2, 5) = 10828.125 / factor ^ 2: data(2, 6) = 1: data(2, 7) = "Rectangle"
    data(2, 8) = 123.75 / factor: data(2, 9) = 87.5 / factor: data(2, 10) = 0#: data(2, 11) = 0.3125
    data(2, 12) = 654321.125 / factor ^ 4: data(2, 13) = 987654.375 / factor ^ 4: data(2, 14) = -23456.625 / factor ^ 4
    data(2, 15) = "Дробные параметры Region"
    data(3, 1) = "R1": data(3, 2) = "Rebar": data(3, 3) = 12.375 / factor: data(3, 4) = -8.125 / factor
    data(3, 5) = 117.8581 / factor ^ 2: data(3, 6) = 1: data(3, 7) = "Circle": data(3, 10) = 12.25 / factor
    data(3, 15) = "Дробная арматура"
    Audit03GeometrySnapshotArray = data
End Function

' Сохраняет фактическую ошибку публичного чтения. Проверка восстановления
' повторяет чтение на том же объекте без подмены сохранённых данных.
Private Sub Audit03ReadLifecycleLoad(ByVal reader As CSectionPlotDataReader, ByVal workbook As Object, _
        ByVal settings As CSystemSettingsReader, ByVal mode As Long, ByRef code As Long, ByRef reason As String)
    code = 0: reason = vbNullString
    On Error GoTo Failed
    Select Case mode
        Case 0: reader.LoadFromWorkbook workbook, settings
        Case 1: reader.LoadGeometryPreviewFromWorkbook workbook, settings
        Case 2: reader.LoadGeometryOnlyForMissingState workbook, settings
    End Select
    Exit Sub
Failed:
    code = Err.Number: reason = Err.Description
    Err.Clear
End Sub

' Проверяет точную сохранность каждого Value2, включая тип. Оформление и
' повторное чтение не должны даже слегка изменять готовые числа Results.
Private Sub Audit03ComparePlainSnapshot(ByRef stats As TUiTestStats, ByVal prefix As String, _
        ByRef baseline As Variant, ByRef actual As Variant, ByVal firstColumn As Long)
    Dim row As Long, column As Long, same As Boolean
    same = (UBound(baseline, 1) = UBound(actual, 1) And UBound(baseline, 2) = UBound(actual, 2))
    If same Then
        For row = 1 To UBound(baseline, 1)
            For column = firstColumn To UBound(baseline, 2)
                same = (VarType(baseline(row, column)) = VarType(actual(row, column)))
                If same Then same = (baseline(row, column) = actual(row, column))
                If Not same Then Exit For
            Next column
            If Not same Then Exit For
        Next row
    End If
    AssertTrue stats, prefix & ".exactValue2", same
End Sub

' Фиксирует одну логическую проверку и не назначает статус расчетному результату.
Private Sub AssertTrue(ByRef stats As TUiTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1: AppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1: AppendLine stats, "FAIL: " & name
    End If
End Sub

' Сверяет независимый координатный oracle в points с явно заданным допуском.
Private Sub AssertClose(ByRef stats As TUiTestStats, ByVal name As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal tolerance As Double)
    AssertTrue stats, name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) <= tolerance
End Sub

' Добавляет строку только в протокол текущего направленного набора.
Private Sub AppendLine(ByRef stats As TUiTestStats, ByVal value As String)
    stats.Report = stats.Report & value & vbCrLf
End Sub
