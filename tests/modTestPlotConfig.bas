Attribute VB_Name = "modTestPlotConfig"
Option Explicit

' Проверяет общие настройки Excel-схемы и компоновку semantic-аннотаций.
' Использует собственные временные книги, готовые snapshots и независимые
' координатные oracle. Не решает НДС и не меняет Config исходной книги.
' Зона ответственности: активные/выключенные presentation-потребители,
' понятные ошибки ввода, сохранность Chart/Results и recovery.

Private Type TUiTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

' ==================== ДЛЯ ТЕСТОВ: ОБЩИЕ НАСТРОЙКИ EXCEL-СХЕМЫ ====================

' Проверяет только presentation pipeline на готовом синтетическом снимке.
' Реальные Chart.Shapes и layout читаются независимо от численного solver;
' исходная книга и пользовательский Excel не используются как test fixture.
Public Function RunAudit03GeneralPlotTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TUiTestStats
    TestAudit03GeneralPlotContracts stats
    TestAudit03AnnotationLayoutContracts stats
    TestAudit03PlotEnableContracts stats
    TestAudit03SnapshotMetadataContracts stats
    TestAudit03AutoPlotContracts stats
    TestAudit03ContourArcContracts stats
    AppendLine stats, "TOTAL_AUDIT03_GENERAL_PLOT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    passed = stats.Passed: failed = stats.Failed
    RunAudit03GeneralPlotTests = stats.Report
End Function

' Проверяет углы дуг сохраненного контура через reader и настоящий preview.
' Рисунок и Results принадлежат отдельной книге; ошибочный текст должен
' отклоняться до очистки прежней схемы, а не превращаться в нулевую дугу.
Public Function RunAudit03ContourArcTests() As String
    Dim stats As TUiTestStats
    TestAudit03ContourArcContracts stats
    AppendLine stats, "TOTAL_AUDIT03_CONTOUR_ARC: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03ContourArcTests = stats.Report
End Function

' Испытывает полное чтение Text, signed sweep и recovery после повреждения.
' Для quarter-circle число хорд задано независимо: шаг схемы 5 градусов,
' поэтому четверть окружности содержит 18 сегментов. НДС не решается.
Private Sub TestAudit03ContourArcContracts(ByRef stats As TUiTestStats)
    Dim fixture As Object, config As Object, sheet As Object, source As Object, target As Object, anchor As Object
    Dim settings As CSystemSettingsReader, reader As CSectionPlotDataReader, plotter As CSectionPlotter
    Dim addresses As Variant, name As Variant, bad As Variant, valid As Variant, shape As Object
    Dim index As Long, position As Long, variantIndex As Long, code As Long, reason As String, prefix As String, exportText As String
    Dim chart As Object, marker As Object, cell As Object, count As Long, cases As Long, solveCount As Long
    Dim annotations(1 To 2, 1 To 10) As Variant, props(1 To 2, 1 To 4) As Variant, keys As Variant
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
    sheet.Range("A5").Resize(3, 15).Value2 = Audit03GeometrySnapshotArray("mm", 1#)
    fixture.Names.Add Name:="rngNDMSectionGeometry", RefersTo:="=Results!$A$5"
    props(1, 1) = "LoadCase": props(1, 2) = "Parameter": props(1, 3) = "Value": props(1, 4) = "Unit"
    props(2, 1) = "ALL": props(2, 2) = "Output.LengthUnit": props(2, 3) = "mm": props(2, 4) = "-"
    sheet.Range("R5").Resize(2, 4).Value2 = props
    fixture.Names.Add Name:="rngNDMSectionProperties", RefersTo:="=Results!$R$5"
    keys = Array("AnnotationType", "AnnotationID", "StartX", "StartY", "EndX", "EndY", "OutsideNormalX", "OutsideNormalY", "Text", "Unit")
    For index = 0 To UBound(keys): annotations(1, index + 1) = keys(index): Next index
    annotations(2, 1) = "CONTOUR_ARC": annotations(2, 2) = "CONTOUR_ARC_1"
    annotations(2, 3) = 100#: annotations(2, 4) = 0#: annotations(2, 5) = 0#: annotations(2, 6) = 100#
    annotations(2, 7) = 0#: annotations(2, 8) = 0#: annotations(2, 9) = "1.570796326795": annotations(2, 10) = "mm"
    fixture.Names.Add Name:="rngNDMSectionAnnotations", RefersTo:="=Results!$A$20"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook fixture
    Set reader = New CSectionPlotDataReader: Set plotter = New CSectionPlotter
    Set chart = fixture.Worksheets.Item("Расчет").ChartObjects.Add(41#, 53#, 700#, 480#): chart.Name = "chtNDMSectionPlot"
    For position = 0 To 1
        If position = 0 Then Set anchor = sheet.Range("A20") Else Set anchor = sheet.Range("CH800")
        anchor.Resize(2, 10).NumberFormat = "@": anchor.Resize(2, 10).Value2 = annotations
        fixture.Names.Item("rngNDMSectionAnnotations").RefersTo = "=Results!" & anchor.Address
        Set cell = anchor.Cells(2, 9): variantIndex = 0
        For Each bad In Array("", "TODO", "1.25garbage", "1,25garbage", "1.2.3", "1,2,3", "1E", "--1", "&H1", CVErr(2015), "1E309", "1E300", "7", "-7", "$1", "1 000", "+.", "1e+1E-2", "True", True)
            variantIndex = variantIndex + 1: cases = cases + 1: cell.Value2 = bad
            prefix = "audit03.contourArc.p" & CStr(position) & ".bad" & CStr(variantIndex)
            Audit03ReadLifecycleLoad reader, fixture, settings, 1, code, reason
            AssertTrue stats, prefix & ".rejected", code <> 0
            AssertTrue stats, prefix & ".address", InStr(1, reason, "Results!" & cell.Address(False, False), vbTextCompare) > 0
            AssertTrue stats, prefix & ".reason", InStr(1, reason, "Text", vbTextCompare) > 0 And InStr(1, reason, "радиан", vbTextCompare) > 0
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
            cell.Value2 = annotations(2, 9): reader.LoadGeometryPreviewFromWorkbook fixture, settings
            AssertTrue stats, prefix & ".recovery", reader.Count = 2 And reader.AnnotationCount = 1
        Next bad
        expected = Array(1.570796326795, 1.570796326795, 1.570796326795, -1.570796326795, -0.5, 0.5, 0.001, 0#)
        index = 0
        For Each valid In Array("1.570796326795", "1,570796326795", "+1.570796326795E0", "-1.570796326795", "-.5", ".5", "1E-3", "0")
            cases = cases + 1: cell.Value2 = valid: baseline = anchor.Resize(2, 10).Value2
            reader.LoadGeometryPreviewFromWorkbook fixture, settings
            plotter.Draw fixture, reader, settings
            count = 0
            For Each shape In chart.Chart.Shapes
                If InStr(1, shape.Name, "ContourArcLine", vbTextCompare) > 0 Then count = count + 1
            Next shape
            prefix = "audit03.contourArc.valid.p" & CStr(position) & "." & CStr(valid)
            AssertClose stats, prefix & ".angle", reader.AnnotationSweepAngle(1), CDbl(expected(index)), 0#
            AssertTrue stats, prefix & ".export", Audit03ReadContourArcSweepsForTests(fixture) = "2|" & CStr(expected(index)) & vbLf
            If InStr(1, CStr(valid), "570796", vbBinaryCompare) > 0 Then
                AssertTrue stats, prefix & ".segments", count = 18
            ElseIf CStr(valid) = "0" Then
                AssertTrue stats, prefix & ".segments", count = 1
            Else
                AssertTrue stats, prefix & ".segments", count = 6
            End If
            actual = anchor.Resize(2, 10).Value2: Audit03ComparePlainSnapshot stats, prefix, baseline, actual, 1
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

' Проверяет автоматическое обновление через настоящий workbook-сценарий.
' Отдельная полная копия книги сохраняет Config/Results исходного fixture;
' НДС-состояния не запрашиваются, счетчик solve не растет.
Public Function RunAudit03AutoPlotTests() As String
    Dim stats As TUiTestStats
    TestAudit03AutoPlotContracts stats
    AppendLine stats, "TOTAL_AUDIT03_AUTO_PLOT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03AutoPlotTests = stats.Report
End Function

' Неверный активный выбор AutoUpdate отвергается до построения модели и
' очистки Results. Общий Plot.Enabled=No не потребляет AutoUpdate и сохраняет
' прежний Chart, в том числе если после запуска нет ни одного named-state.
Private Sub TestAudit03AutoPlotContracts(ByRef stats As TUiTestStats)
    Dim fixture As Object, config As Object, table As Object, cell As Object, profiles As Object, loads As Object
    Dim chart As Object, marker As Object, baseline As Variant, bad As Variant, key As Variant
    Dim path As String, row As Long, column As Long, profileColumn As Long, position As Long, cases As Long, mode As Long
    Dim code As Long, reason As String, prefix As String, solveCount As Long, result As String
    On Error GoTo Failed
    solveCount = SectionEquilibriumSolveCount()
    path = ThisWorkbook.Path & "\Audit03_AutoPlot_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsm"
    ThisWorkbook.SaveCopyAs path
    Set fixture = Application.Workbooks.Open(path, 0, False)
    AssertTrue stats, "audit03.autoPlot.ownedCopy", StrComp(fixture.FullName, ThisWorkbook.FullName, vbTextCompare) <> 0
    Set config = fixture.Worksheets.Item("Config")
    Set table = fixture.Names.Item("rngSystemSettings").RefersToRange
    Audit03SetPlotSetting table, "General.ExecutionReportEnabled", "No"
    Audit03SetPlotSetting table, "Geometry.Type", "INVALID"
    baseline = table.Value2
    On Error Resume Next
    Set chart = fixture.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")
    On Error GoTo Failed
    If chart Is Nothing Then Set chart = fixture.Worksheets.Item("Расчет").ChartObjects.Add(41#, 53#, 700#, 480#)
    chart.Name = "chtNDMSectionPlot"
    Set marker = chart.Chart.Shapes.AddShape(1, 11#, 17#, 20#, 20#): marker.Name = "NDMPlot_Audit03Sentinel"
    fixture.Names.Item("rngBatchSummary").RefersToRange.Cells(1, 1).Value2 = "AUDIT03_PRESERVE"
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
                AssertTrue stats, prefix & ".rejected", code <> 0 And InStr(1, reason, CStr(key), vbTextCompare) > 0
                AssertTrue stats, prefix & ".address", InStr(1, reason, "Config", vbTextCompare) > 0 And InStr(1, reason, cell.Address(False, False), vbTextCompare) > 0
                AssertTrue stats, prefix & ".snapshot", CStr(fixture.Names.Item("rngBatchSummary").RefersToRange.Cells(1, 1).Value2) = "AUDIT03_PRESERVE"
                AssertTrue stats, prefix & ".chart", Audit03PlotShapeExists(chart, "Audit03Sentinel")
                AppendLine stats, "AUTO_PLOT_ERROR: " & prefix & "|" & reason
            Next bad
        Next key
        table.Value2 = baseline: Audit03SetPlotSetting table, "Plot.Enabled", "Yes"
        Set cell = Audit03PlotSettingCell(table, "Plot.AutoUpdateAfterCalculation"): cell.Value2 = "Yes": cell.Offset(0, -1).Value2 = "Audit03.RemovedAutoUpdate"
        Audit03CaptureCalculation fixture, code, reason, result
        AssertTrue stats, "audit03.autoPlot.missing.p" & CStr(position), code <> 0 And InStr(1, reason, "Plot.AutoUpdateAfterCalculation", vbTextCompare) > 0
        table.Value2 = baseline: Audit03SetPlotSetting table, "Plot.Enabled", "No"
        Audit03SetPlotSetting table, "Plot.AutoUpdateAfterCalculation", CVErr(2015)
        Audit03CaptureCalculation fixture, code, reason, result
        AssertTrue stats, "audit03.autoPlot.inactive.p" & CStr(position), code <> 0 And InStr(1, reason, "Geometry.Type", vbTextCompare) > 0 And InStr(1, reason, "Plot.AutoUpdateAfterCalculation", vbTextCompare) = 0
    Next position
    ' Выполняем действительный stability-only запуск без НДС-состояний.
    table.Value2 = baseline
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
    AssertTrue stats, "audit03.autoPlot.noSolve", SectionEquilibriumSolveCount() = solveCount
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
    Dim props(1 To 20, 1 To 4) As Variant, elements(1 To 3, 1 To 7) As Variant, annotations(1 To 4, 1 To 9) As Variant
    Dim grid() As Variant, gridResults() As Variant, series As Object, colors As Object, channel As Long
    Dim arrow As Variant, size As Variant, shape As Object, style As Long, metric As Long, badIndex As Long, mode As Long
    Dim row As Long, position As Long, index As Long, code As Long, reason As String, prefix As String
    Dim baseline As Variant, snapshot As Variant, actual As Variant, chart As Object, marker As Object
    Dim cases As Long, solveCount As Long, geometry As Variant
    On Error GoTo Failed
    solveCount = SectionEquilibriumSolveCount()
    Set fixture = Application.Workbooks.Add(-4167)
    Set config = fixture.Worksheets(1): config.Name = "PlotConfig"
    Set sheet = fixture.Worksheets.Add: sheet.Name = "Results"
    fixture.Worksheets.Add.Name = "Расчет"
    Set table = Audit03PlotSettingsTable(config.Range("A5"))
    baseline = table.Value2
    geometry = Audit03GeometrySnapshotArray("mm", 1#)
    sheet.Range("A5").Resize(3, 15).Value2 = geometry
    fixture.Names.Add Name:="rngNDMSectionGeometry", RefersTo:="=Results!$A$5"
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
    sheet.Range("R5").Resize(20, 4).Value2 = props
    fixture.Names.Add Name:="rngNDMSectionProperties", RefersTo:="=Results!$R$5"
    elements(1, 1) = "LoadCase": elements(1, 2) = "ProfileId": elements(1, 3) = "StateType": elements(1, 4) = "ElementID"
    elements(1, 5) = "Strain": elements(1, 6) = "Stress, MPa": elements(1, 7) = "PhysicalState"
    elements(2, 1) = "PLOT": elements(2, 2) = "PR1": elements(2, 3) = "StrengthState": elements(2, 4) = "C1"
    elements(2, 5) = -0.0001: elements(2, 6) = -1.25: elements(2, 7) = "Compression"
    elements(3, 1) = "PLOT": elements(3, 2) = "PR1": elements(3, 3) = "StrengthState": elements(3, 4) = "R1"
    elements(3, 5) = 0.000015: elements(3, 6) = 2.75: elements(3, 7) = "Tension"
    sheet.Range("R35").Resize(3, 7).Value2 = elements
    fixture.Names.Add Name:="rngNDMElementResults", RefersTo:="=Results!$R$35"
    keys = Array("AnnotationType", "StartX", "StartY", "EndX", "EndY", "OutsideNormalX", "OutsideNormalY", "Text", "Unit")
    For index = 0 To UBound(keys): annotations(1, index + 1) = keys(index): Next index
    For row = 2 To 4
        For index = 2 To 7: annotations(row, index) = 0#: Next index
        annotations(row, 9) = "mm"
    Next row
    annotations(2, 1) = "CONTOUR_CIRCLE": annotations(2, 4) = 100#: annotations(2, 8) = ""
    annotations(3, 1) = "DIMENSION": annotations(3, 2) = -50#: annotations(3, 4) = 50#: annotations(3, 7) = 1#: annotations(3, 8) = "DIM"
    annotations(4, 1) = "REBAR_ANNOTATION": annotations(4, 2) = -50#: annotations(4, 4) = 50#: annotations(4, 7) = 1#: annotations(4, 8) = "REBAR"
    sheet.Range("A20").Resize(4, 9).Value2 = annotations
    fixture.Names.Add Name:="rngNDMSectionAnnotations", RefersTo:="=Results!$A$20"
    snapshot = sheet.Range("A5:X37").Value2
    Set settings = New CSystemSettingsReader: settings.LoadFromRange table
    Set reader = New CSectionPlotDataReader: reader.LoadFromWorkbook fixture, settings
    Set plotter = New CSectionPlotter
    plotter.Draw fixture, reader, settings
    Set chart = fixture.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")
    chart.Left = 41#: chart.Top = 53#: chart.Width = 700#: chart.Height = 480#
    keys = Array("Plot.ContourEnabled", "Plot.NeutralLineEnabled", "Plot.LoadApplicationPointEnabled", _
        "Plot.ResultLabelsEnabled", "Plot.LegendEnabled", "Plot.PrincipalAxesMode", "Plot.ResultLabelSpacing")
    For position = 0 To 1
        If position = 0 Then Set table = config.Range("A5").Resize(UBound(baseline, 1), 3) Else Set table = config.Range("CH800").Resize(UBound(baseline, 1), 3)
        table.NumberFormat = "@": table.Value2 = baseline
        Audit03SetPlotSetting table, "Plot.ResultLabelsEnabled", "Yes"
        For Each key In keys
            Set cell = Audit03PlotSettingCell(table, CStr(key))
            badIndex = 0
            For Each bad In Array("", "TODO", "INVALID", CVErr(2015))
                badIndex = badIndex + 1
                cases = cases + 1: values = cell.Value2: cell.Value2 = bad
                settings.LoadFromRange table
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
                cell.Value2 = values: settings.LoadFromRange table
                Audit03CapturePlotDraw plotter, fixture, reader, settings, code, reason
                AssertTrue stats, prefix & ".recovery", code = 0
            Next bad
        Next key
        For Each bad In Array(0#, -1#)
            cases = cases + 1
            Audit03SetPlotSetting table, "Plot.ResultLabelSpacing", bad
            settings.LoadFromRange table
            Audit03CapturePlotDraw plotter, fixture, reader, settings, code, reason
            AssertTrue stats, "audit03.generalPlot.spacing.sign.p" & CStr(position) & "." & CStr(bad), code <> 0 And InStr(1, reason, "Plot.ResultLabelSpacing", vbTextCompare) > 0
        Next bad
        Audit03SetPlotSetting table, "Plot.ResultLabelSpacing", 0.00000001
        settings.LoadFromRange table
        Audit03CapturePlotDraw plotter, fixture, reader, settings, code, reason
        AssertTrue stats, "audit03.generalPlot.spacing.smallPositive.p" & CStr(position), code = 0
        AssertTrue stats, "audit03.generalPlot.spacing.label.p" & CStr(position), Audit03PlotTextExists(chart, "-1.25")
        Audit03SetPlotSetting table, "Plot.ResultLabelsEnabled", "No"
        Audit03SetPlotSetting table, "Plot.ResultLabelSpacing", CVErr(2015)
        settings.LoadFromRange table
        Audit03CapturePlotDraw plotter, fixture, reader, settings, code, reason
        AssertTrue stats, "audit03.generalPlot.spacing.inactive.p" & CStr(position), code = 0
        Audit03SetPlotSetting table, "Plot.NeutralLineEnabled", CVErr(2015)
        Audit03SetPlotSetting table, "Plot.ResultLabelsEnabled", CVErr(2015)
        Audit03SetPlotSetting table, "Plot.LegendEnabled", CVErr(2015)
        settings.LoadFromRange table: reader.LoadGeometryPreviewFromWorkbook fixture, settings
        Audit03CapturePlotDraw plotter, fixture, reader, settings, code, reason
        AssertTrue stats, "audit03.generalPlot.geometry.inactive.p" & CStr(position), code = 0
        table.Value2 = baseline: settings.LoadFromRange table: reader.LoadFromWorkbook fixture, settings
        For Each key In Array("Plot.NeutralLineEnabled", "Plot.LoadApplicationPointEnabled", "Plot.PrincipalAxesMode")
            Audit03SetPlotSetting table, CStr(key), IIf(CStr(key) = "Plot.PrincipalAxesMode", "None", "No")
            settings.LoadFromRange table: plotter.Draw fixture, reader, settings
            Select Case CStr(key)
                Case "Plot.NeutralLineEnabled": AssertTrue stats, "audit03.generalPlot.neutral.off.p" & CStr(position), Not Audit03PlotShapeExists(chart, "NeutralAxis")
                Case "Plot.LoadApplicationPointEnabled": AssertTrue stats, "audit03.generalPlot.load.off.p" & CStr(position), Not Audit03PlotShapeExists(chart, "LoadPoint")
                Case Else: AssertTrue stats, "audit03.generalPlot.axes.off.p" & CStr(position), Not Audit03PlotShapeExists(chart, "Principal1")
            End Select
            Audit03SetPlotSetting table, CStr(key), IIf(CStr(key) = "Plot.PrincipalAxesMode", "Concrete", "Yes")
            settings.LoadFromRange table: plotter.Draw fixture, reader, settings
            Select Case CStr(key)
                Case "Plot.NeutralLineEnabled": AssertTrue stats, "audit03.generalPlot.neutral.on.p" & CStr(position), Audit03PlotShapeExists(chart, "NeutralAxis")
                Case "Plot.LoadApplicationPointEnabled": AssertTrue stats, "audit03.generalPlot.load.on.p" & CStr(position), Audit03PlotShapeExists(chart, "LoadPoint")
                Case Else: AssertTrue stats, "audit03.generalPlot.axes.on.p" & CStr(position), Audit03PlotShapeExists(chart, "Principal1")
            End Select
        Next key
        Audit03SetPlotSetting table, "Plot.LegendEnabled", "No"
        Audit03SetPlotSetting table, "Plot.ResultLabelSpacing", 100#
        For Each key In Array("Plot.ContourEnabled", "Plot.ResultLabelsEnabled", "Plot.LegendEnabled")
            Audit03SetPlotSetting table, CStr(key), "No": settings.LoadFromRange table: plotter.Draw fixture, reader, settings
            prefix = "audit03.generalPlot.toggle.p" & CStr(position) & "." & CStr(key)
            Select Case CStr(key)
                Case "Plot.ContourEnabled": AssertTrue stats, prefix & ".off", Not Audit03PlotShapeExists(chart, "ContourCircle")
                Case "Plot.ResultLabelsEnabled": AssertTrue stats, prefix & ".off", Not Audit03PlotTextExists(chart, "-1.25")
                Case Else: AssertTrue stats, prefix & ".off", Not Audit03PlotShapeExists(chart, "LegendGradient")
            End Select
            Audit03SetPlotSetting table, CStr(key), "Yes": settings.LoadFromRange table: plotter.Draw fixture, reader, settings
            Select Case CStr(key)
                Case "Plot.ContourEnabled": AssertTrue stats, prefix & ".on", Audit03PlotShapeExists(chart, "ContourCircle")
                Case "Plot.ResultLabelsEnabled": AssertTrue stats, prefix & ".on", Audit03PlotTextExists(chart, "-1.25")
                Case Else: AssertTrue stats, prefix & ".on", Audit03PlotShapeExists(chart, "LegendGradient")
            End Select
        Next key
        Audit03SetPlotSetting table, "Plot.ResultLabelsEnabled", "Yes"
        Audit03SetPlotSetting table, "Plot.ResultLabelSpacing", 100#
        For Each key In keys
            Set cell = Audit03PlotSettingCell(table, CStr(key))
            cases = cases + 1
            cell.Offset(0, -1).Value2 = "REMOVED_" & CStr(key): settings.LoadFromRange table
            Audit03CapturePlotDraw plotter, fixture, reader, settings, code, reason
            AssertTrue stats, "audit03.generalPlot.missing.p" & CStr(position) & "." & CStr(key), code <> 0 And InStr(1, reason, CStr(key), vbTextCompare) > 0
            cell.Offset(0, -1).Value2 = key
        Next key
        table.Value2 = baseline
        Audit03SetPlotSetting table, "Plot.Dimensions.Enabled", "Yes"
        For Each key In Array("Plot.Dimensions.ArrowType", "Plot.Dimensions.ArrowSize")
            Set cell = Audit03PlotSettingCell(table, CStr(key)): values = cell.Value2: badIndex = 0
            For Each bad In Array("", "TODO", "INVALID", CVErr(2015))
                badIndex = badIndex + 1: cases = cases + 1: cell.Value2 = bad: settings.LoadFromRange table
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
                cases = cases + 1: Audit03SetPlotSetting table, "Plot.LoadCase", bad: settings.LoadFromRange table
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
        table.Value2 = baseline: settings.LoadFromRange table: reader.LoadFromWorkbook fixture, settings
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
            settings.LoadFromRange table: plotter.Draw fixture, reader, settings
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
    actual = sheet.Range("A5:X37").Value2
    Audit03ComparePlainSnapshot stats, "audit03.generalPlot.ResultsUnchanged", snapshot, actual, 1
    ' На сетке >5000 точек проверяем все 17 уровней градиента. Округление
    ' номера bucket не должно пропускать нечетные целые значения.
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
    table.Value2 = baseline
    Audit03SetPlotSetting table, "Plot.Color.ConcreteCompression", "0,0,0"
    Audit03SetPlotSetting table, "Plot.ResultGradient", "Yes"
    settings.LoadFromRange table: reader.LoadFromWorkbook fixture, settings
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

' Возвращает именно текущую ячейку значения, а не жестко заданный адрес.
' Отсутствующий test-key является ошибкой fixture, а не успешной проверкой.
Private Function Audit03PlotSettingCell(ByVal table As Object, ByVal key As String) As Object
    Dim row As Long
    For row = 2 To table.Rows.Count
        If CStr(table.Cells(row, 1).Value2) = key Then
            Set Audit03PlotSettingCell = table.Cells(row, 2): Exit Function
        End If
    Next row
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

' Проверяет готовые координаты, шрифты, цвета и веса layout при mm/pt,
' Inside/Outside и выключенной линии арматуры. Oracle задан в points по
' независимому transform 3.2 по X и 2.15 по Y, без обращения к plotter.
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

' Проверяет два workbook-entrypoint-а без solver: общий No должен отключать
' и обычное обновление, и import-preview, не читая поврежденные параметры
' выключенной схемы. Ошибочный сам переключатель требует исправления Config.
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

' Создает замороженный геометрический fixture без builder/solver. Дробные
' координаты и оболочка не должны зависеть от текущей INPUT/OUTPUT системы.
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

' Перехватывает фактическую ошибку выбранного публичного reader-режима;
' recovery повторяет чтение тем же объектом и не подменяет данные snapshot.
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
