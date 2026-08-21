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
    TestBlankMomentDefaultsToZeroAndZeroLoadsAreSkipped stats
    TestCircleWorkbookRunWritesResults stats
    TestLShapeWorkbookRunWritesResults stats
    TestAutoCADExportUsesSharedLoadReference stats
    TestGoverningCombinationWritesDetailedResults stats
    TestCapacitySearchMethodValidation stats
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
    Dim plotButton As Object
    Set runButton = calc.Shapes.Item("btnRunSectionCalculation")
    Set clearButton = calc.Shapes.Item("btnClearSectionResults")
    Set acadButton = calc.Shapes.Item("btnExportStressToAutoCAD")
    Set plotButton = calc.Shapes.Item("btnUpdateSectionPlot")

    AssertTrue stats, "ui.buttons.run.exists", Not runButton Is Nothing
    AssertTrue stats, "ui.buttons.clear.exists", Not clearButton Is Nothing
    AssertTrue stats, "ui.buttons.autocad.exists", Not acadButton Is Nothing
    AssertTrue stats, "ui.buttons.plot.exists", Not plotButton Is Nothing
    AssertTrue stats, "ui.buttons.run.macro", InStr(1, runButton.OnAction, "RunSectionCalculation", vbTextCompare) > 0
    AssertTrue stats, "ui.buttons.clear.macro", InStr(1, clearButton.OnAction, "ClearSectionResults", vbTextCompare) > 0
    AssertTrue stats, "ui.buttons.autocad.macro", InStr(1, acadButton.OnAction, "ExportSectionStressToAutoCAD", vbTextCompare) > 0
    AssertTrue stats, "ui.buttons.plot.macro", InStr(1, plotButton.OnAction, "UpdateSectionPlot", vbTextCompare) > 0
    AssertTrue stats, "ui.buttons.outsidePrint", runButton.Left > calc.Range("AJ1").Left And clearButton.Left > calc.Range("AJ1").Left And acadButton.Left > calc.Range("AJ1").Left And plotButton.Left > calc.Range("AJ1").Left
End Sub

Private Sub TestAutoCADExportUsesSharedLoadReference(ByRef stats As TUiTestStats)
    PrepareLShapeInput
    SetSystemSetting "Load.ReferenceOffsetX", "0"
    SetSystemSetting "Load.ReferenceOffsetY", "0"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells.Item(2, 2).Value2 = -1000000#
    loads.Cells.Item(2, 3).Value2 = 0#
    loads.Cells.Item(2, 4).Value2 = 0#

    Dim section As CSectionModel
    Dim concrete As CConcreteDiagramMaterial
    Dim steel As CSteelDiagramMaterial
    Dim solver As CSectionSolver
    Dim loadReferenceX As Double
    Dim loadReferenceY As Double
    PrepareAutoCADExportState ThisWorkbook, section, concrete, steel, solver, loadReferenceX, loadReferenceY

    AssertTrue stats, "ui.autocad.reference.converged", solver.Converged
    AssertClose stats, "ui.autocad.reference.kappaX", solver.KappaX, 0#, 0.00000001
    AssertClose stats, "ui.autocad.reference.kappaY", solver.KappaY, 0#, 0.00000001
    AssertTrue stats, "ui.autocad.reference.point", Abs(loadReferenceX) > 0.000001 Or Abs(loadReferenceY) > 0.000001

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateTransformed section, concrete, steel
    AssertClose stats, "ui.autocad.axes.centerX", props.CentroidX, loadReferenceX, 0.000001
    AssertClose stats, "ui.autocad.axes.centerY", props.CentroidY, loadReferenceY, 0.000001
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
    loads.Cells.Item(2, 1).ClearContents

    Dim batch As CBatchSectionCalculator
    Set batch = BuildUiBatch()

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch
    batch.Execute

    AssertTrue stats, "ui.loads.partial.count", batch.Count = 1
    AssertTrue stats, "ui.loads.partial.invalid", InStr(1, batch.Status(1), "InvalidInput", vbTextCompare) > 0

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    AssertTrue stats, "ui.loads.partial.message", InStr(1, message, "ошиб", vbTextCompare) > 0 And _
        InStr(1, message, "InvalidInput", vbTextCompare) > 0
End Sub

Private Sub TestBlankMomentDefaultsToZeroAndZeroLoadsAreSkipped(ByRef stats As TUiTestStats)
    PrepareCircleInput
    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells.Item(2, 4).ClearContents
    loads.Cells.Item(3, 1).Value2 = "ZERO"
    loads.Cells.Item(3, 2).Value2 = 0#
    loads.Cells.Item(3, 3).Value2 = 0#
    loads.Cells.Item(3, 4).Value2 = 0#
    loads.Cells.Item(3, 5).Value2 = "StrengthAndCrack"

    Dim batch As CBatchSectionCalculator
    Set batch = BuildUiBatch()

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch
    batch.Execute

    AssertTrue stats, "ui.loads.blankMoment.count", batch.Count = 1
    AssertClose stats, "ui.loads.blankMoment.myZero", batch.UserMy(1), 0#, 0.0000001
    AssertTrue stats, "ui.loads.blankMoment.valid", InStr(1, batch.Status(1), "InvalidInput", vbTextCompare) = 0
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
    Dim sys As Object
    Set sys = ThisWorkbook.Worksheets.Item("Settings")
    AssertTrue stats, "ui.run.system.noRebarTable", Len(CStr(sys.Cells.Item(130, 1).Value2)) = 0
    AssertTrue stats, "ui.run.system.noCrackFormulaBlock", Len(CStr(sys.Cells.Item(130, 9).Value2)) = 0
    AssertTrue stats, "ui.run.crack.result.value", Not ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(17, 5).HasFormula
    AssertTrue stats, "ui.results.elements.header", CStr(ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.Value2) = "RunID"
    AssertTrue stats, "ui.results.elements.rows", ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.CurrentRegion.Rows.Count > 1
    AssertTrue stats, "ui.results.elements.noCombinationIndex", ResultHeaderColumn(ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.CurrentRegion.Value2, "CombinationIndex") = 0
    AssertTrue stats, "ui.results.elements.units", ResultHeaderColumn(ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.CurrentRegion.Value2, "Stress, MPa") > 0
    AssertTrue stats, "ui.results.elements.noGeometryDup", ResultHeaderColumn(ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.CurrentRegion.Value2, "X, mm") = 0
    AssertTrue stats, "ui.results.elements.noPlaneDup", ResultHeaderColumn(ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.CurrentRegion.Value2, "Epsilon0") = 0
    AssertTrue stats, "ui.results.geometry.header", CStr(ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Value2) = "RunID"
    AssertTrue stats, "ui.results.geometry.rows", ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.CurrentRegion.Rows.Count > 1
    AssertTrue stats, "ui.results.geometry.position", ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Column = 10
    AssertTrue stats, "ui.results.geometry.noSource", ResultHeaderColumn(ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.CurrentRegion.Value2, "SourceName") = 0
    AssertTrue stats, "ui.results.properties.header", CStr(ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.Value2) = "RunID"
    AssertTrue stats, "ui.results.properties.hasEpsilon0", ResultsPropertyExists("LC1", "Epsilon0")
    AssertTrue stats, "ui.results.properties.hasBounds", ResultsPropertyExists("ALL", "Bounds.MinX")
    AssertTrue stats, "ui.results.annotations.header", CStr(ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.Value2) = "RunID"
    AssertTrue stats, "ui.results.annotations.rows", ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.CurrentRegion.Rows.Count > 1
    AssertTrue stats, "ui.plot.chart.created", PlotChartExists()
End Sub

Private Sub TestLShapeWorkbookRunWritesResults(ByRef stats As TUiTestStats)
    PrepareLShapeInput
    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    AssertTrue stats, "ui.lshape.message", InStr(1, message, "завершен", vbTextCompare) > 0
    AssertTrue stats, "ui.lshape.result.status", Len(CStr(ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(5, 5).Value2)) > 0

    Dim sys As Object
    Set sys = ThisWorkbook.Worksheets.Item("Settings")
    AssertTrue stats, "ui.lshape.system.noRebarTable", Len(CStr(sys.Cells.Item(130, 1).Value2)) = 0
    AssertTrue stats, "ui.lshape.system.noCrackFormulaBlock", Len(CStr(sys.Cells.Item(130, 9).Value2)) = 0
End Sub

Private Sub TestGoverningCombinationWritesDetailedResults(ByRef stats As TUiTestStats)
    PrepareCircleInput
    SetSystemSetting "Calculation.Mode", "FullCapacity"
    SetSystemSetting "Capacity.Method", "LoadMultiplier"
    SetSystemSetting "Capacity.ToleranceLambda", "0.05"
    SetSystemSetting "Capacity.MaxLambda", "10"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells.Item(2, 1).Value2 = "LC_SAFE"
    loads.Cells.Item(2, 2).Value2 = -100000#
    loads.Cells.Item(2, 3).Value2 = -1000000#
    loads.Cells.Item(2, 4).Value2 = 0#
    loads.Cells.Item(2, 5).Value2 = "StrengthAndCrack"
    loads.Cells.Item(2, 6).Value2 = "LongTerm"
    loads.Cells.Item(2, 7).Value2 = "less severe"

    loads.Cells.Item(3, 1).Value2 = "LC_GOV"
    loads.Cells.Item(3, 2).Value2 = -100000#
    loads.Cells.Item(3, 3).Value2 = -8000000#
    loads.Cells.Item(3, 4).Value2 = 0#
    loads.Cells.Item(3, 5).Value2 = "StrengthAndCrack"
    loads.Cells.Item(3, 6).Value2 = "LongTerm"
    loads.Cells.Item(3, 7).Value2 = "governing"

    Dim message As String
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)

    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim summaryRow As Long
    summaryRow = BatchSummaryStartRow()
    Dim governingID As String
    governingID = CStr(resultsSheet.Cells.Item(summaryRow + 1, 2).Value2)
    Dim expectedID As String
    expectedID = ExpectedGoverningByLowestLambda(resultsSheet, summaryRow)

    AssertTrue stats, "ui.governing.id", governingID = expectedID
    AssertTrue stats, "ui.governing.message", InStr(1, message, governingID, vbTextCompare) > 0
    AssertTrue stats, "ui.governing.details.lambda", Abs(CDbl(ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(5, 5).Value2) - LowestPositiveLambda(resultsSheet, summaryRow)) < 0.0000001
End Sub

Private Function ExpectedGoverningByLowestLambda(ByVal resultsSheet As Object, ByVal summaryRow As Long) As String
    Dim rowIndex As Long
    Dim bestLambda As Double
    For rowIndex = summaryRow + 6 To summaryRow + 25
        If Len(Trim$(CStr(resultsSheet.Cells.Item(rowIndex, 1).Value2))) > 0 Then
            If IsNumeric(resultsSheet.Cells.Item(rowIndex, 8).Value2) Then
                Dim lambdaValue As Double
                lambdaValue = CDbl(resultsSheet.Cells.Item(rowIndex, 8).Value2)
                If lambdaValue > 0# And (bestLambda = 0# Or lambdaValue < bestLambda) Then
                    bestLambda = lambdaValue
                    ExpectedGoverningByLowestLambda = CStr(resultsSheet.Cells.Item(rowIndex, 1).Value2)
                End If
            End If
        End If
    Next rowIndex
End Function

Private Function LowestPositiveLambda(ByVal resultsSheet As Object, ByVal summaryRow As Long) As Double
    Dim rowIndex As Long
    For rowIndex = summaryRow + 6 To summaryRow + 25
        If IsNumeric(resultsSheet.Cells.Item(rowIndex, 8).Value2) Then
            Dim lambdaValue As Double
            lambdaValue = CDbl(resultsSheet.Cells.Item(rowIndex, 8).Value2)
            If lambdaValue > 0# And (LowestPositiveLambda = 0# Or lambdaValue < LowestPositiveLambda) Then
                LowestPositiveLambda = lambdaValue
            End If
        End If
    Next rowIndex
End Function

Private Function BatchSummaryStartRow() As Long
    BatchSummaryStartRow = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Row
End Function

Private Function ResultHeaderColumn(ByRef data As Variant, ByVal headerText As String) As Long
    Dim colIndex As Long
    For colIndex = 1 To UBound(data, 2)
        If StrComp(CStr(data(1, colIndex)), headerText, vbTextCompare) = 0 Then
            ResultHeaderColumn = colIndex
            Exit Function
        End If
    Next colIndex
End Function

Private Sub TestCapacitySearchMethodValidation(ByRef stats As TUiTestStats)
    AssertTrue stats, "ui.validation.capacitySearchMethod", _
        SystemSettingValidationHasOptions("Capacity.SearchMethod", Array("Bisection", "Brent", "Secant"))
    AssertTrue stats, "ui.validation.autocadLabelMode", _
        SystemSettingValidationHasOptions("AutoCAD.Export.LabelMode", Array("ValuesOnly", "NamesAndValues"))
    AssertTrue stats, "ui.validation.autocadResultType", _
        SystemSettingValidationHasOptions("AutoCAD.Export.ResultType", Array("Stress", "Strain"))
    AssertTrue stats, "ui.validation.autocadNeutralLine", _
        SystemSettingValidationHasOptions("AutoCAD.Export.NeutralLineEnabled", Array("Yes", "No"))
    AssertTrue stats, "ui.validation.autocadPrincipalAxes", _
        SystemSettingValidationHasOptions("AutoCAD.Export.PrincipalAxesEnabled", Array("Yes", "No"))
    AssertTrue stats, "ui.validation.autocadLoadPoint", _
        SystemSettingValidationHasOptions("AutoCAD.Export.LoadPointEnabled", Array("Yes", "No"))
    AssertTrue stats, "ui.validation.autocadCombination", AutoCADCombinationValidationIsDynamic()
    AssertTrue stats, "ui.validation.plotLoadCase", PlotLoadCaseValidationIsDynamic()
    AssertTrue stats, "ui.validation.plotResultType", _
        SystemSettingValidationHasOptions("Plot.ResultType", Array("Stress", "Strain"))
    AssertTrue stats, "ui.validation.plotLabels", _
        SystemSettingValidationHasOptions("Plot.ResultLabelsEnabled", Array("Yes", "No"))
    AssertTrue stats, "ui.validation.plotRebarAnnotationEnabled", _
        PlotAnnotationValidationHasOptions("Enabled", 2, Array("Yes", "No"))
    AssertTrue stats, "ui.validation.plotDimensionEnabled", _
        PlotAnnotationValidationHasOptions("Enabled", 3, Array("Yes", "No"))
    AssertTrue stats, "ui.validation.plotRebarPlacement", _
        PlotAnnotationValidationHasOptions("Placement", 2, Array("Outside", "Inside"))
    AssertTrue stats, "ui.validation.plotDimensionPlacement", _
        PlotAnnotationValidationHasOptions("Placement", 3, Array("Outside", "Inside"))
    AssertTrue stats, "ui.validation.plotRebarLineEnabled", _
        PlotAnnotationValidationHasOptions("LineEnabled", 2, Array("Yes", "No"))
    AssertTrue stats, "ui.validation.plotDimensionArrowType", _
        PlotAnnotationValidationHasOptions("ArrowType", 3, Array("Triangle", "Stealth", "Diamond", "Oval", "Open"))
    AssertTrue stats, "ui.validation.plotDimensionArrowSize", _
        PlotAnnotationValidationHasOptions("ArrowSize", 3, Array("Small", "Medium", "Wide"))
    AssertTrue stats, "ui.validation.geometrySource", _
        SystemSettingValidationHasOptions("Geometry.Source", Array("Generated", "AutoCAD"))
    AssertTrue stats, "ui.validation.units.forceInput", _
        AnySettingValidationHasOptions("Force", Array("N", "kN", "tf"))
    AssertTrue stats, "ui.validation.units.momentOutput", _
        AnySettingValidationHasOptions("Moment", Array("N*mm", "kN*m", "tf*m"))
    AssertTrue stats, "ui.validation.sign.n", _
        AnySettingValidationHasOptions("+N", Array("Tension", "Compression"))
    SetSystemSetting "Geometry.Type", "Circle"
    AssertTrue stats, "ui.validation.circleLoc2row", _
        AnySettingValidationHasOptions("Rebar.Loc2row", Array("Stacked", "SideBySide"))
    AssertTrue stats, "ui.validation.circleLoc3row", _
        AnySettingValidationHasOptions("Rebar.Loc3row", Array("Stacked", "SideBySide"))
    Dim reader As CSystemSettingsReader
    Set reader = New CSystemSettingsReader
    reader.LoadFromWorkbook ThisWorkbook
    AssertTrue stats, "ui.circle.key.rebarDiameter2", reader.HasKey("Rebar.Diameter2")
    AssertTrue stats, "ui.circle.key.rebarDiameter3", reader.HasKey("Rebar.Diameter3")
    AssertTrue stats, "ui.circle.key.rebarLoc2row", reader.HasKey("Rebar.Loc2row")
    AssertTrue stats, "ui.circle.key.rebarLoc3row", reader.HasKey("Rebar.Loc3row")
End Sub

Private Function AutoCADCombinationValidationIsDynamic() As Boolean
    On Error GoTo Failed
    PrepareCircleInput

    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), "AutoCAD.Export.CombinationID", vbTextCompare) = 0 Then
            Dim formulaText As String
            formulaText = CStr(settings.Cells.Item(rowIndex, 2).Validation.Formula1)
            If Left$(formulaText, 1) <> "=" Then Exit Function

            Dim listRange As Object
            Set listRange = ThisWorkbook.Worksheets.Item("Settings").Range(Mid$(formulaText, 2))
            AutoCADCombinationValidationIsDynamic = (CStr(listRange.Cells.Item(1, 1).Value2) = "Worst" And _
                CStr(listRange.Cells.Item(2, 1).Value2) = "LC1" And listRange.Cells.Item(2, 1).HasFormula)
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

Private Function PlotLoadCaseValidationIsDynamic() As Boolean
    PlotLoadCaseValidationIsDynamic = LoadCaseValidationIsDynamicForSetting("Plot.LoadCase")
End Function

Private Function LoadCaseValidationIsDynamicForSetting(ByVal settingKey As String) As Boolean
    On Error GoTo Failed
    PrepareCircleInput

    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), settingKey, vbTextCompare) = 0 Then
            Dim formulaText As String
            formulaText = CStr(settings.Cells.Item(rowIndex, 2).Validation.Formula1)
            If Left$(formulaText, 1) <> "=" Then Exit Function

            Dim listRange As Object
            Set listRange = ThisWorkbook.Worksheets.Item("Settings").Range(Mid$(formulaText, 2))
            LoadCaseValidationIsDynamicForSetting = (CStr(listRange.Cells.Item(1, 1).Value2) = "Worst" And _
                CStr(listRange.Cells.Item(2, 1).Value2) = "LC1" And listRange.Cells.Item(2, 1).HasFormula)
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

Private Function ResultsPropertyExists(ByVal loadCase As String, ByVal parameter As String) As Boolean
    On Error GoTo Failed
    Dim data As Variant
    data = ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.CurrentRegion.Value2
    Dim colLC As Long: colLC = ResultHeaderColumn(data, "LoadCase")
    Dim colParam As Long: colParam = ResultHeaderColumn(data, "Parameter")
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(CStr(data(rowIndex, colLC)), loadCase, vbTextCompare) = 0 And _
                StrComp(CStr(data(rowIndex, colParam)), parameter, vbTextCompare) = 0 Then
            ResultsPropertyExists = True
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

Private Function PlotChartExists() As Boolean
    On Error GoTo Failed
    Dim chartObject As Object
    Set chartObject = ThisWorkbook.Worksheets.Item("Расчет").ChartObjects("chtNDMSectionPlot")
    PlotChartExists = Not chartObject Is Nothing
Failed:
End Function

Private Sub TestClearResultsKeepsInputs(ByRef stats As TUiTestStats)
    PrepareCircleInput
    ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(5, 5).Value2 = 123#
    ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.Value2 = "RunID"
    ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Value2 = "RunID"
    ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.Value2 = "RunID"
    ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.Value2 = "RunID"
    ClearSectionResultsForWorkbook ThisWorkbook
    AssertTrue stats, "ui.clear.result", Len(CStr(ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(5, 5).Value2)) = 0
    AssertTrue stats, "ui.clear.results.elements", Len(CStr(ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.Value2)) = 0
    AssertTrue stats, "ui.clear.results.geometry", Len(CStr(ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Value2)) = 0
    AssertTrue stats, "ui.clear.results.properties", Len(CStr(ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange.Value2)) = 0
    AssertTrue stats, "ui.clear.results.annotations", Len(CStr(ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.Value2)) = 0
    AssertTrue stats, "ui.clear.input", CStr(GetSystemSetting("Geometry.Type")) = "Circle"
End Sub

Private Sub SetSystemSetting(ByVal key As String, ByVal value As String)
    If TrySetUnitOrSignSetting(key, value) Then Exit Sub

    Dim settings As Object
    Dim ranges As Variant
    ranges = SettingsRangeSearchOrder()

    Dim rangeIndex As Long
    For rangeIndex = LBound(ranges) To UBound(ranges)
        Set settings = ThisWorkbook.Names.Item(CStr(ranges(rangeIndex))).RefersToRange

        Dim rowIndex As Long
        For rowIndex = 2 To settings.Rows.Count
            If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), key, vbTextCompare) = 0 Then
                settings.Cells.Item(rowIndex, 2).Value2 = value
                Exit Sub
            End If
        Next rowIndex
    Next rangeIndex

    If TrySetLShapeFaceSetting(key, value) Then Exit Sub

    Err.Raise vbObjectError + 4210, "modTestWorkbookInterface", "System setting not found: " & key
End Sub

Private Function TrySetUnitOrSignSetting(ByVal key As String, ByVal value As String) As Boolean
    Dim target As Object
    Dim rowIndex As Long
    Dim columnIndex As Long

    If UnitSettingAddress(key, rowIndex, columnIndex) Then
        Set target = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
        target.Cells.Item(rowIndex, columnIndex).Value2 = value
        TrySetUnitOrSignSetting = True
        Exit Function
    End If

    If SignSettingAddress(key, rowIndex, columnIndex) Then
        Set target = ThisWorkbook.Names.Item("rngSignConventionSettings").RefersToRange
        target.Cells.Item(rowIndex, columnIndex).Value2 = value
        TrySetUnitOrSignSetting = True
    End If
End Function

Private Function UnitSettingAddress(ByVal key As String, ByRef rowIndex As Long, ByRef columnIndex As Long) As Boolean
    If InStr(1, key, "Units.", vbTextCompare) <> 1 Then Exit Function

    Dim quantity As String
    Dim sideName As String
    Dim tail As String
    tail = Mid$(key, 7)
    Dim dotPos As Long
    dotPos = InStr(1, tail, ".", vbTextCompare)
    If dotPos <= 0 Then Exit Function
    quantity = Left$(tail, dotPos - 1)
    sideName = Mid$(tail, dotPos + 1)

    Select Case LCase$(quantity)
        Case "length": rowIndex = 2
        Case "area": rowIndex = 3
        Case "force": rowIndex = 4
        Case "moment": rowIndex = 5
        Case "stress": rowIndex = 6
        Case "curvature": rowIndex = 7
        Case Else: Exit Function
    End Select

    Select Case LCase$(sideName)
        Case "input": columnIndex = 2
        Case "internal": columnIndex = 3
        Case "output": columnIndex = 4
        Case Else: Exit Function
    End Select
    UnitSettingAddress = True
End Function

Private Function SignSettingAddress(ByVal key As String, ByRef rowIndex As Long, ByRef columnIndex As Long) As Boolean
    If InStr(1, key, "Sign.", vbTextCompare) <> 1 Then Exit Function

    Dim quantity As String
    Dim sideName As String
    Dim tail As String
    tail = Mid$(key, 6)
    Dim dotPos As Long
    dotPos = InStr(1, tail, ".", vbTextCompare)
    If dotPos <= 0 Then Exit Function
    quantity = Left$(tail, dotPos - 1)
    sideName = Mid$(tail, dotPos + 1)

    Select Case LCase$(quantity)
        Case "n": rowIndex = 2
        Case "mx": rowIndex = 3
        Case "my": rowIndex = 4
        Case Else: Exit Function
    End Select

    Select Case LCase$(sideName)
        Case "user": columnIndex = 2
        Case "internal": columnIndex = 3
        Case Else: Exit Function
    End Select
    SignSettingAddress = True
End Function

Private Function TrySetLShapeFaceSetting(ByVal key As String, ByVal value As String) As Boolean
    Dim target As Object
    Set target = ThisWorkbook.Names.Item("rngLShapeGeometry").RefersToRange

    Dim rowIndex As Long
    Dim columnIndex As Long
    If Not LShapeSettingAddress(key, rowIndex, columnIndex) Then Exit Function
    target.Cells.Item(rowIndex, columnIndex).Value2 = value
    TrySetLShapeFaceSetting = True
End Function

Private Function LShapeSettingAddress(ByVal key As String, ByRef rowIndex As Long, ByRef columnIndex As Long) As Boolean
    Dim faceName As String
    If InStr(1, key, "LShape.H1", vbTextCompare) = 1 Then
        faceName = "H1": columnIndex = 2
    ElseIf InStr(1, key, "LShape.B1", vbTextCompare) = 1 Then
        faceName = "B1": columnIndex = 3
    ElseIf InStr(1, key, "LShape.H2", vbTextCompare) = 1 Then
        faceName = "H2": columnIndex = 4
    ElseIf InStr(1, key, "LShape.B2", vbTextCompare) = 1 Then
        faceName = "B2": columnIndex = 5
    Else
        Exit Function
    End If

    If StrComp(key, "LShape." & faceName, vbTextCompare) = 0 Then
        rowIndex = 2
    ElseIf StrComp(key, "LShape." & faceName & ".as_1", vbTextCompare) = 0 Then
        rowIndex = 3
    ElseIf StrComp(key, "LShape." & faceName & ".as_2", vbTextCompare) = 0 Then
        rowIndex = 4
    ElseIf StrComp(key, "LShape." & faceName & ".d_1", vbTextCompare) = 0 Then
        rowIndex = 5
    ElseIf StrComp(key, "LShape." & faceName & ".d_2", vbTextCompare) = 0 Then
        rowIndex = 6
    ElseIf StrComp(key, "LShape." & faceName & ".n_1", vbTextCompare) = 0 Then
        rowIndex = 7
    ElseIf StrComp(key, "LShape." & faceName & ".n_2", vbTextCompare) = 0 Then
        rowIndex = 8
    ElseIf StrComp(key, "LShape." & faceName & ".StartOffset1", vbTextCompare) = 0 Then
        rowIndex = 9
    ElseIf StrComp(key, "LShape." & faceName & ".EndOffset1", vbTextCompare) = 0 Then
        rowIndex = 10
    ElseIf StrComp(key, "LShape." & faceName & ".StartOffset2", vbTextCompare) = 0 Then
        rowIndex = 11
    ElseIf StrComp(key, "LShape." & faceName & ".EndOffset2", vbTextCompare) = 0 Then
        rowIndex = 12
    ElseIf StrComp(key, "LShape." & faceName & ".d_2row_1", vbTextCompare) = 0 Then
        rowIndex = 14
    ElseIf StrComp(key, "LShape." & faceName & ".d_2row_2", vbTextCompare) = 0 Then
        rowIndex = 15
    ElseIf StrComp(key, "LShape." & faceName & ".d_3row_1", vbTextCompare) = 0 Then
        rowIndex = 16
    ElseIf StrComp(key, "LShape." & faceName & ".d_3row_2", vbTextCompare) = 0 Then
        rowIndex = 17
    ElseIf StrComp(key, "LShape." & faceName & ".loc_2row", vbTextCompare) = 0 Then
        rowIndex = 18
    ElseIf StrComp(key, "LShape." & faceName & ".loc_3row", vbTextCompare) = 0 Then
        rowIndex = 19
    Else
        Exit Function
    End If
    LShapeSettingAddress = True
End Function

Private Function GetSystemSetting(ByVal key As String) As String
    Dim settings As Object
    Dim ranges As Variant
    ranges = SettingsRangeSearchOrder()

    Dim rangeIndex As Long
    For rangeIndex = LBound(ranges) To UBound(ranges)
        Set settings = ThisWorkbook.Names.Item(CStr(ranges(rangeIndex))).RefersToRange

        Dim rowIndex As Long
        For rowIndex = 2 To settings.Rows.Count
            If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), key, vbTextCompare) = 0 Then
                GetSystemSetting = CStr(settings.Cells.Item(rowIndex, 2).Value2)
                Exit Function
            End If
        Next rowIndex
    Next rangeIndex

    Err.Raise vbObjectError + 4211, "modTestWorkbookInterface", "System setting not found: " & key
End Function

Private Function SettingsRangeSearchOrder() As Variant
    Dim geometryType As String
    geometryType = SystemGeometryType()
    If StrComp(geometryType, "LShape", vbTextCompare) = 0 Then
        SettingsRangeSearchOrder = Array("rngUnitSettings", "rngSignConventionSettings", "rngSystemSettings", "rngPlotAnnotationSettings", "rngLShapeGeometry", "rngCircleGeometry", "rngRoundedRectangleGeometry")
    ElseIf StrComp(geometryType, "RoundedRectangle", vbTextCompare) = 0 Then
        SettingsRangeSearchOrder = Array("rngUnitSettings", "rngSignConventionSettings", "rngSystemSettings", "rngPlotAnnotationSettings", "rngRoundedRectangleGeometry", "rngCircleGeometry", "rngLShapeGeometry")
    Else
        SettingsRangeSearchOrder = Array("rngUnitSettings", "rngSignConventionSettings", "rngSystemSettings", "rngPlotAnnotationSettings", "rngCircleGeometry", "rngRoundedRectangleGeometry", "rngLShapeGeometry")
    End If
End Function

Private Function SystemGeometryType() As String
    On Error GoTo Failed
    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange

    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), "Geometry.Type", vbTextCompare) = 0 Then
            SystemGeometryType = CStr(settings.Cells.Item(rowIndex, 2).Value2)
            Exit Function
        End If
    Next rowIndex
Failed:
    SystemGeometryType = "Circle"
End Function

Private Function SystemSettingValidationHasOptions(ByVal key As String, ByVal expectedOptions As Variant) As Boolean
    SystemSettingValidationHasOptions = SettingValidationHasOptionsInRange(ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange, key, expectedOptions)
End Function

Private Function AnySettingValidationHasOptions(ByVal key As String, ByVal expectedOptions As Variant) As Boolean
    Dim ranges As Variant
    ranges = SettingsRangeSearchOrder()

    Dim rangeIndex As Long
    For rangeIndex = LBound(ranges) To UBound(ranges)
        If SettingValidationHasOptionsInRange(ThisWorkbook.Names.Item(CStr(ranges(rangeIndex))).RefersToRange, key, expectedOptions) Then
            AnySettingValidationHasOptions = True
            Exit Function
        End If
    Next rangeIndex
End Function

Private Function SettingValidationHasOptionsInRange(ByVal settings As Object, ByVal key As String, ByVal expectedOptions As Variant) As Boolean
    On Error GoTo Failed

    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), key, vbTextCompare) = 0 Then
            Dim formulaText As String
            formulaText = CStr(settings.Cells.Item(rowIndex, 2).Validation.Formula1)
            If Left$(formulaText, 1) <> "=" Then Exit Function

            Dim listRange As Object
            Set listRange = ThisWorkbook.Worksheets.Item("Settings").Range(Mid$(formulaText, 2))
            Dim i As Long
            If listRange.Cells.Count <> (UBound(expectedOptions) - LBound(expectedOptions) + 1) Then Exit Function
            For i = LBound(expectedOptions) To UBound(expectedOptions)
                If CStr(listRange.Cells.Item(i - LBound(expectedOptions) + 1, 1).Value2) <> CStr(expectedOptions(i)) Then Exit Function
            Next i
            SettingValidationHasOptionsInRange = True
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

Private Function PlotAnnotationValidationHasOptions(ByVal rowName As String, ByVal valueColumn As Long, ByVal expectedOptions As Variant) As Boolean
    On Error GoTo Failed

    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngPlotAnnotationSettings").RefersToRange

    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), rowName, vbTextCompare) = 0 Then
            Dim formulaText As String
            formulaText = CStr(settings.Cells.Item(rowIndex, valueColumn).Validation.Formula1)
            If Left$(formulaText, 1) <> "=" Then Exit Function

            Dim listRange As Object
            Set listRange = ThisWorkbook.Worksheets.Item("Settings").Range(Mid$(formulaText, 2))
            Dim i As Long
            If listRange.Cells.Count <> (UBound(expectedOptions) - LBound(expectedOptions) + 1) Then Exit Function
            For i = LBound(expectedOptions) To UBound(expectedOptions)
                If CStr(listRange.Cells.Item(i - LBound(expectedOptions) + 1, 1).Value2) <> CStr(expectedOptions(i)) Then Exit Function
            Next i
            PlotAnnotationValidationHasOptions = True
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

Private Function BuildUiBatch() As CBatchSectionCalculator
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByDiameter settings.GetDouble("Circle.Diameter", 300#), _
        0#, _
        0#

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 20#, 20#, 1

    Dim builder As CCircleRebarLayoutBuilder
    Set builder = New CCircleRebarLayoutBuilder
    Dim rebars As CRebarLayout
    Set rebars = builder.Build(settings.GetDouble("Circle.Diameter", 300#), _
        0#, _
        0#, _
        settings.GetDouble("Rebar.AxisDistance", 40#), _
        settings.GetLong("Rebar.Count", 8), _
        settings.GetDouble("Rebar.Diameter", 20#), _
        settings.GetString("Steel.Class", "A400"))
    Dim concrete As CConcreteDiagramMaterial
    Set concrete = New CConcreteDiagramMaterial
    concrete.Initialize -0.0015, -15.5, -0.0035, -15.5
    concrete.ApplySettings settings

    Dim steel As CSteelDiagramMaterial
    Set steel = New CSteelDiagramMaterial
    steel.Initialize 0.00175, 350#, 0.025

    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, rebars, "WorkbookInterface")

    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize section, concrete, steel
    batch.ApplySettings settings
    Set BuildUiBatch = batch
End Function

Private Sub PrepareCircleInput()
    SetSystemSetting "Units.Force.Input", "N"
    SetSystemSetting "Units.Moment.Input", "N*mm"
    SetSystemSetting "Units.Length.Input", "mm"
    SetSystemSetting "Units.Area.Input", "mm2"
    SetSystemSetting "Units.Stress.Input", "MPa"
    SetSystemSetting "Units.Curvature.Input", "1/mm"
    SetSystemSetting "Sign.N.User", "Tension"
    SetSystemSetting "Sign.Mx.User", "+Y tension"
    SetSystemSetting "Sign.My.User", "+X tension"
    SetSystemSetting "Geometry.Source", "Generated"
    SetSystemSetting "Geometry.Type", "Circle"
    SetSystemSetting "Calculation.Mode", "DirectState"
    SetSystemSetting "Concrete.Class", "B30"
    SetSystemSetting "Steel.Class", "A400"
    SetSystemSetting "Rebar.AxisDistance", "40"
    SetSystemSetting "Rebar.Count", "8"
    SetSystemSetting "Rebar.Diameter", "20"
    SetSystemSetting "CrackWidth.Allowable", "0.3"

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

Private Sub PrepareLShapeInput()
    SetSystemSetting "Units.Force.Input", "N"
    SetSystemSetting "Units.Moment.Input", "N*mm"
    SetSystemSetting "Units.Length.Input", "mm"
    SetSystemSetting "Units.Area.Input", "mm2"
    SetSystemSetting "Units.Stress.Input", "MPa"
    SetSystemSetting "Units.Curvature.Input", "1/mm"
    SetSystemSetting "Sign.N.User", "Tension"
    SetSystemSetting "Sign.Mx.User", "+Y tension"
    SetSystemSetting "Sign.My.User", "+X tension"
    SetSystemSetting "Geometry.Source", "Generated"
    SetSystemSetting "Geometry.Type", "LShape"
    SetSystemSetting "Calculation.Mode", "DirectState"
    SetSystemSetting "Mesh.Step", "40"
    SetSystemSetting "Mesh.BoundarySubdivisions", "2"
    SetSystemSetting "LShape.B1", "160"
    SetSystemSetting "LShape.H1", "280"
    SetSystemSetting "LShape.B2", "360"
    SetSystemSetting "LShape.H2", "140"
    SetSystemSetting "LShape.H1.as_1", "40"
    SetSystemSetting "LShape.H1.as_2", "40"
    SetSystemSetting "LShape.H1.d_1", "16"
    SetSystemSetting "LShape.H1.d_2", "16"
    SetSystemSetting "LShape.H1.n_1", "3"
    SetSystemSetting "LShape.H1.n_2", "2"
    SetSystemSetting "LShape.H1.StartOffset1", "40"
    SetSystemSetting "LShape.H1.EndOffset1", "40"
    SetSystemSetting "LShape.H1.StartOffset2", "40"
    SetSystemSetting "LShape.H1.EndOffset2", "40"
    SetSystemSetting "LShape.H2.as_1", "40"
    SetSystemSetting "LShape.H2.as_2", "40"
    SetSystemSetting "LShape.H2.d_1", "16"
    SetSystemSetting "LShape.H2.d_2", "16"
    SetSystemSetting "LShape.H2.n_1", "0"
    SetSystemSetting "LShape.H2.n_2", "0"
    SetSystemSetting "LShape.H2.StartOffset1", "40"
    SetSystemSetting "LShape.H2.EndOffset1", "40"
    SetSystemSetting "LShape.H2.StartOffset2", "40"
    SetSystemSetting "LShape.H2.EndOffset2", "40"
    SetSystemSetting "LShape.B1.as_1", "40"
    SetSystemSetting "LShape.B1.as_2", "40"
    SetSystemSetting "LShape.B1.d_1", "16"
    SetSystemSetting "LShape.B1.d_2", "16"
    SetSystemSetting "LShape.B1.n_1", "2"
    SetSystemSetting "LShape.B1.n_2", "1"
    SetSystemSetting "LShape.B1.StartOffset1", "20"
    SetSystemSetting "LShape.B1.EndOffset1", "20"
    SetSystemSetting "LShape.B1.StartOffset2", "20"
    SetSystemSetting "LShape.B1.EndOffset2", "20"
    SetSystemSetting "LShape.B2.as_1", "40"
    SetSystemSetting "LShape.B2.as_2", "40"
    SetSystemSetting "LShape.B2.d_1", "16"
    SetSystemSetting "LShape.B2.d_2", "16"
    SetSystemSetting "LShape.B2.n_1", "2"
    SetSystemSetting "LShape.B2.n_2", "0"
    SetSystemSetting "LShape.B2.StartOffset1", "60"
    SetSystemSetting "LShape.B2.EndOffset1", "60"
    SetSystemSetting "LShape.B2.StartOffset2", "60"
    SetSystemSetting "LShape.B2.EndOffset2", "60"
    SetSystemSetting "CrackWidth.Allowable", "0.3"

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    ClearDataRows loads
    loads.Cells.Item(2, 1).Value2 = "LC_L"
    loads.Cells.Item(2, 2).Value2 = -80000#
    loads.Cells.Item(2, 3).Value2 = -1500000#
    loads.Cells.Item(2, 4).Value2 = -1000000#
    loads.Cells.Item(2, 5).Value2 = "StrengthAndCrack"
    loads.Cells.Item(2, 6).Value2 = "LongTerm"
    loads.Cells.Item(2, 7).Value2 = "lshape ui test"
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

Private Sub AssertTrue(ByRef stats As TUiTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name
    End If
End Sub

Private Sub AssertClose(ByRef stats As TUiTestStats, ByVal name As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal tolerance As Double)
    If Abs(actual - expected) <= tolerance Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name & "; actual=" & FormatNumberInvariant(actual) & _
            "; expected=" & FormatNumberInvariant(expected)
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name & "; actual=" & FormatNumberInvariant(actual) & _
            "; expected=" & FormatNumberInvariant(expected)
    End If
End Sub

Private Sub AppendLine(ByRef stats As TUiTestStats, ByVal text As String)
    stats.Report = stats.Report & text & vbCrLf
End Sub

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function










