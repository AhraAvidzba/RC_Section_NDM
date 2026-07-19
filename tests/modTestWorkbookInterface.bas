Attribute VB_Name = "modTestWorkbookInterface"
Option Explicit

Private Type TUiTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

Public Function RunWorkbookInterfaceTests() As String
    On Error GoTo Failed

    Dim stats As TUiTestStats
    Dim t0 As Double
    t0 = Timer

    TestButtons stats
    TestSingleCombinationSkipsBlankRows stats
    TestPartialCombinationIsInvalid stats
    TestCircleWorkbookRunWritesResults stats
    TestLinearMatrixModeWritesResults stats
    TestClearResultsKeepsInputs stats

    AppendLine stats, "TOTAL_WORKBOOK_UI: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
    RunWorkbookInterfaceTests = stats.Report
    Exit Function

Failed:
    RunWorkbookInterfaceTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

Private Sub TestButtons(ByRef stats As TUiTestStats)
    Dim calc As Object
    Set calc = ThisWorkbook.Worksheets.Item("Расчет")

    Dim runButton As Object
    Dim clearButton As Object
    Dim acadButton As Object
    Set runButton = calc.Shapes.Item("btnRunSectionCalculation")
    Set clearButton = calc.Shapes.Item("btnClearSectionResults")
    Set acadButton = calc.Shapes.Item("btnExportStressToAutoCAD")

    AssertTrue stats, "ui.buttons.run.exists", Not runButton Is Nothing
    AssertTrue stats, "ui.buttons.clear.exists", Not clearButton Is Nothing
    AssertTrue stats, "ui.buttons.autocad.exists", Not acadButton Is Nothing
    AssertTrue stats, "ui.buttons.run.macro", InStr(1, runButton.OnAction, "RunSectionCalculation", vbTextCompare) > 0
    AssertTrue stats, "ui.buttons.clear.macro", InStr(1, clearButton.OnAction, "ClearSectionResults", vbTextCompare) > 0
    AssertTrue stats, "ui.buttons.autocad.macro", InStr(1, acadButton.OnAction, "ExportSectionStressToAutoCAD", vbTextCompare) > 0
    AssertTrue stats, "ui.buttons.outsidePrint", runButton.Left > calc.Range("AJ1").Left And clearButton.Left > calc.Range("AJ1").Left And acadButton.Left > calc.Range("AJ1").Left
End Sub

Private Sub TestSingleCombinationSkipsBlankRows(ByRef stats As TUiTestStats)
    PrepareCircleInput

    Dim batch As CBatchSectionCalculator
    Set batch = BuildUiBatch()

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch

    Dim t0 As Double
    t0 = Timer
    batch.Execute

    AssertTrue stats, "ui.loads.single.count", batch.Count = 1
    AssertTrue stats, "ui.loads.single.id", batch.CombinationID(1) = "LC1"
    AssertTrue stats, "ui.loads.single.elapsed", (Timer - t0) < 20#
    AssertTrue stats, "ui.loads.single.noBlankInvalid", InStr(1, batch.DiagnosticLog, "InvalidInput", vbTextCompare) = 0
End Sub

Private Sub TestPartialCombinationIsInvalid(ByRef stats As TUiTestStats)
    PrepareCircleInput
    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells.Item(2, 2).ClearContents

    Dim batch As CBatchSectionCalculator
    Set batch = BuildUiBatch()

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch
    batch.Execute

    AssertTrue stats, "ui.loads.partial.count", batch.Count = 1
    AssertTrue stats, "ui.loads.partial.invalid", InStr(1, batch.Status(1), "InvalidInput", vbTextCompare) > 0
End Sub

Private Sub TestCircleWorkbookRunWritesResults(ByRef stats As TUiTestStats)
    PrepareCircleInput
    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    AssertTrue stats, "ui.run.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    AssertTrue stats, "ui.run.capacity.lambda", Len(CStr(ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(5, 5).Value2)) > 0
    AssertTrue stats, "ui.run.fastMode.noLambdaSearch", CStr(ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(5, 5).Value2) = "NotCalculated"
    AssertTrue stats, "ui.run.deformations", Len(CStr(ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(12, 5).Value2)) > 0
    AssertTrue stats, "ui.run.crack", Len(CStr(ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(17, 5).Value2)) > 0
    AssertTrue stats, "ui.run.rebar.auto.firstX", Abs(CDbl(ThisWorkbook.Names.Item("rngRebarInput").RefersToRange.Cells.Item(2, 2).Value2) - 110#) < 0.000001
    AssertTrue stats, "ui.run.rebar.auto.firstY", Abs(CDbl(ThisWorkbook.Names.Item("rngRebarInput").RefersToRange.Cells.Item(2, 3).Value2)) < 0.000001
    Dim sys As Object
    Set sys = ThisWorkbook.Worksheets.Item("System")
    AssertTrue stats, "ui.run.crack.formulaBlock.title", Len(CStr(sys.Cells.Item(130, 9).Value2)) > 0
    AssertTrue stats, "ui.run.crack.formula.width", sys.Cells.Item(155, 10).HasFormula
    AssertTrue stats, "ui.run.crack.formula.ar", sys.Cells.Item(150, 10).HasFormula
    AssertTrue stats, "ui.run.crack.formula.rebar", sys.Cells.Item(160, 15).HasFormula
    AssertTrue stats, "ui.run.crack.result.link", ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(17, 5).HasFormula
End Sub

Private Sub TestLinearMatrixModeWritesResults(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "Calculation.Mode", "LinearMatrix"

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    AssertTrue stats, "ui.linear.message", InStr(1, message, "Расчет завершен", vbTextCompare) > 0
    AssertTrue stats, "ui.linear.mode", InStr(1, CStr(ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(2, 13).Value2), "Линейно", vbTextCompare) > 0
    AssertTrue stats, "ui.linear.noCrack", CStr(ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(17, 5).Value2) = "NotCalculated"

    SetSystemSetting "Calculation.Mode", "DirectState"
End Sub

Private Sub TestClearResultsKeepsInputs(ByRef stats As TUiTestStats)
    PrepareCircleInput
    ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(5, 5).Value2 = 123#
    ClearSectionResultsForWorkbook ThisWorkbook
    AssertTrue stats, "ui.clear.result", Len(CStr(ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(5, 5).Value2)) = 0
    AssertTrue stats, "ui.clear.input", CStr(ThisWorkbook.Names.Item("rngMainInput").RefersToRange.Cells.Item(2, 5).Value2) = "Circle"
End Sub

Private Sub SetSystemSetting(ByVal key As String, ByVal value As String)
    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange

    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), key, vbTextCompare) = 0 Then
            settings.Cells.Item(rowIndex, 2).Value2 = value
            Exit Sub
        End If
    Next rowIndex

    Err.Raise vbObjectError + 4210, "modTestWorkbookInterface", "System setting not found: " & key
End Sub

Private Function BuildUiBatch() As CBatchSectionCalculator
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByDiameter settings.GetDouble("Circle.Diameter", 300#), _
        settings.GetDouble("Circle.CenterX", 0#), _
        settings.GetDouble("Circle.CenterY", 0#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 20#, 20#, 1

    Dim builder As CCircleRebarLayoutBuilder
    Set builder = New CCircleRebarLayoutBuilder
    Dim rebars As CRebarLayout
    Set rebars = builder.Build(settings.GetDouble("Circle.Diameter", 300#), _
        settings.GetDouble("Circle.CenterX", 0#), _
        settings.GetDouble("Circle.CenterY", 0#), 40#, 8, 20#, "A400")
    Dim concrete As CConcreteDiagramMaterial
    Set concrete = New CConcreteDiagramMaterial
    concrete.Initialize -0.0015, -15.5, -0.0035, -15.5
    concrete.ApplySettings settings

    Dim steel As CSteelDiagramMaterial
    Set steel = New CSteelDiagramMaterial
    steel.Initialize 0.00175, 350#, 0.025

    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize mesh, rebars, concrete, steel
    batch.ApplySettings settings
    Set BuildUiBatch = batch
End Function

Private Sub PrepareCircleInput()
    SetSystemSetting "Geometry.Type", "Circle"
    SetSystemSetting "Calculation.Mode", "DirectState"

    Dim mainInput As Object
    Set mainInput = ThisWorkbook.Names.Item("rngMainInput").RefersToRange
    mainInput.Cells.Item(12, 5).Value2 = "B30"
    mainInput.Cells.Item(13, 5).Value2 = "A400"
    mainInput.Cells.Item(14, 5).Value2 = 40#
    mainInput.Cells.Item(15, 5).Value2 = 8
    mainInput.Cells.Item(16, 5).Value2 = 20#
    mainInput.Cells.Item(17, 5).Value2 = 0.3

    Dim rebars As Object
    Set rebars = ThisWorkbook.Names.Item("rngRebarInput").RefersToRange
    ClearDataRows rebars
    PutRebarRow rebars, 2, "MANUAL", 999#, 999#, 99#, 0#, "BAD", "must be overwritten"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    ClearDataRows loads
    loads.Cells.Item(2, 1).Value2 = "LC1"
    loads.Cells.Item(2, 2).Value2 = -100000#
    loads.Cells.Item(2, 3).Value2 = -4000000#
    loads.Cells.Item(2, 4).Value2 = -3000000#
    loads.Cells.Item(2, 5).Value2 = "StrengthAndCrack"
    loads.Cells.Item(2, 6).Value2 = "LongTerm"
    loads.Cells.Item(2, 7).Value2 = "ui test"
End Sub

Private Sub ClearDataRows(ByVal target As Object)
    Dim rowIndex As Long
    Dim colIndex As Long
    For rowIndex = 2 To target.Rows.Count
        For colIndex = 1 To target.Columns.Count
            target.Cells.Item(rowIndex, colIndex).ClearContents
        Next colIndex
    Next rowIndex
End Sub

Private Sub PutRebarRow(ByVal target As Object, ByVal rowIndex As Long, ByVal id As String, _
        ByVal xValue As Double, ByVal yValue As Double, ByVal diameter As Double, ByVal area As Double, _
        ByVal steelClass As String, ByVal comment As String)
    target.Cells.Item(rowIndex, 1).Value2 = id
    target.Cells.Item(rowIndex, 2).Value2 = xValue
    target.Cells.Item(rowIndex, 3).Value2 = yValue
    target.Cells.Item(rowIndex, 4).Value2 = diameter
    target.Cells.Item(rowIndex, 5).Value2 = area
    target.Cells.Item(rowIndex, 6).Value2 = steelClass
    target.Cells.Item(rowIndex, 7).Value2 = comment
End Sub

Private Sub AssertTrue(ByRef stats As TUiTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name
    End If
End Sub

Private Sub AppendLine(ByRef stats As TUiTestStats, ByVal text As String)
    stats.Report = stats.Report & text & vbCrLf
End Sub

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function









