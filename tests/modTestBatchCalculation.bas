Attribute VB_Name = "modTestBatchCalculation"
Option Explicit

Private Type TBatchTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

Public Function RunBatchCalculationTests() As String
    On Error GoTo Failed

    Dim stats As TBatchTestStats
    Dim t0 As Double
    t0 = Timer

    AppendLine stats, "RUN: TestBatchOneCombination"
    TestBatchOneCombination stats
    AppendLine stats, "RUN: TestBatchFiveCombinations"
    TestBatchFiveCombinations stats
    AppendLine stats, "RUN: TestBatchTwentyCombinations"
    TestBatchTwentyCombinations stats
    AppendLine stats, "RUN: TestInvalidCombinationFromNamedRange"
    TestInvalidCombinationFromNamedRange stats
    AppendLine stats, "RUN: TestBatchSummaryWriter"
    TestBatchSummaryWriter stats
    AppendLine stats, "RUN: TestBatchCapacityUsesSystemSettings"
    TestBatchCapacityUsesSystemSettings stats

    AppendLine stats, "TOTAL_BATCH: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
    RunBatchCalculationTests = stats.Report
    Exit Function

Failed:
    RunBatchCalculationTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

Private Sub TestBatchOneCombination(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "C1", -220000#, -7000000#, -5000000#, "StrengthAndCrack", "LongTerm", "single"
    batch.Execute

    AssertTrue stats, "batch.one.count", batch.Count = 1
    AssertTrue stats, "batch.one.governing", batch.GoverningCombinationID = "C1"
    AssertTrue stats, "batch.one.capacity.status", Len(batch.CapacityStatus(1)) > 0
    AssertTrue stats, "batch.one.crack.status", Len(batch.CrackStatus(1)) > 0
    AssertTrue stats, "batch.one.elapsed", batch.ElapsedSeconds >= 0#
End Sub

Private Sub TestBatchCapacityUsesSystemSettings(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldMaxLambda As String
    oldMode = GetSystemSetting("Calculation.Mode")
    oldMaxLambda = GetSystemSetting("Capacity.MaxLambda")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Calculation.Mode", "FullCapacity"
    SetSystemSetting "Capacity.MaxLambda", "0.5"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "LIMITED", -220000#, -7000000#, -5000000#, "StrengthAndCrack", "ShortTerm", "max-lambda"
    batch.Execute

    AssertTrue stats, "batch.settings.capacity.maxLambda", InStr(1, batch.CapacityStatus(1), "NumericalFailure", vbTextCompare) > 0

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "Capacity.MaxLambda", oldMaxLambda
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.settings.capacity.maxLambda; " & Err.Description
    Resume Restore
End Sub

Private Sub TestBatchFiveCombinations(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()

    Dim i As Long
    For i = 1 To 5
        batch.AddCombination "C" & CStr(i), -150000# - 10000# * i, -3000000# - 250000# * i, _
            -2000000# - 200000# * i, "StrengthAndCrack", "ShortTerm", "five-" & CStr(i)
    Next i
    batch.Execute

    AssertTrue stats, "batch.five.count", batch.Count = 5
    AssertTrue stats, "batch.five.governing.index", batch.GoverningCombinationIndex >= 1 And batch.GoverningCombinationIndex <= 5
    AssertTrue stats, "batch.five.diagnostics", InStr(1, batch.DiagnosticLog, "combination=C5", vbTextCompare) > 0
End Sub

Private Sub TestBatchTwentyCombinations(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()

    Dim i As Long
    For i = 1 To 20
        batch.AddCombination "LC" & CStr(i), -100000# - 2500# * i, -1800000# - 100000# * i, _
            -1200000# - 75000# * i, "StrengthAndCrack", "LongTerm", "twenty-" & CStr(i)
    Next i
    batch.Execute

    AssertTrue stats, "batch.twenty.count", batch.Count = 20
    AssertTrue stats, "batch.twenty.governing.index", batch.GoverningCombinationIndex >= 1 And batch.GoverningCombinationIndex <= 20
    AssertTrue stats, "batch.twenty.last.status", Len(batch.Status(20)) > 0
End Sub

Private Sub TestInvalidCombinationFromNamedRange(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.ClearContents
    loads.Cells.Item(1, 1).Value2 = "CombinationID"
    loads.Cells.Item(1, 2).Value2 = "N"
    loads.Cells.Item(1, 3).Value2 = "Mx"
    loads.Cells.Item(1, 4).Value2 = "My"
    loads.Cells.Item(1, 5).Value2 = "CalculationType"
    loads.Cells.Item(1, 6).Value2 = "DurationType"
    loads.Cells.Item(1, 7).Value2 = "Comment"
    loads.Cells.Item(2, 1).Value2 = "BAD"
    loads.Cells.Item(2, 2).Value2 = "not-a-number"
    loads.Cells.Item(2, 3).Value2 = -1000000#
    loads.Cells.Item(2, 4).Value2 = -500000#
    loads.Cells.Item(2, 5).Value2 = "StrengthAndCrack"
    loads.Cells.Item(2, 6).Value2 = "ShortTerm"
    loads.Cells.Item(2, 7).Value2 = "invalid source row"

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch
    batch.Execute

    AssertTrue stats, "batch.invalid.reader.count", batch.Count = 1
    AssertTrue stats, "batch.invalid.reader.status", InStr(1, batch.Status(1), "InvalidInput", vbTextCompare) > 0
End Sub

Private Sub TestBatchSummaryWriter(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "W1", -180000#, -3500000#, -2500000#, "StrengthAndCrack", "ShortTerm", "writer"
    batch.Execute

    Dim writer As CBatchResultWriter
    Set writer = New CBatchResultWriter
    writer.WriteSummary ThisWorkbook, batch

    Dim systemSheet As Object
    Set systemSheet = ThisWorkbook.Worksheets.Item("System")
    AssertTrue stats, "batch.writer.title", CStr(systemSheet.Cells.Item(92, 1).Value2) = "Batch summary"
    AssertTrue stats, "batch.writer.governing", CStr(systemSheet.Cells.Item(93, 2).Value2) = "W1"
End Sub

Private Function BuildBatchCalculator() As CBatchSectionCalculator
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 300#, 200#, 0#, 0#, 0#, 0#

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 30#, 20#, 1

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -90#, -60#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B2", 90#, -60#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B3", -90#, 60#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B4", 90#, 60#, 20#, 0#, "A400", "", geom

    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize mesh, rebars, ProvisionalConcrete(), ProvisionalSteel()
    Set BuildBatchCalculator = batch
End Function

Private Function GetSystemSetting(ByVal key As String) As String
    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), key, vbTextCompare) = 0 Then
            GetSystemSetting = CStr(settings.Cells.Item(rowIndex, 2).Value2)
            Exit Function
        End If
    Next rowIndex
    Err.Raise vbObjectError + 3930, "modTestBatchCalculation", "System setting not found: " & key
End Function

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
    Err.Raise vbObjectError + 3931, "modTestBatchCalculation", "System setting not found: " & key
End Sub

Private Function ProvisionalConcrete() As CConcreteDiagramMaterial
    Dim concrete As CConcreteDiagramMaterial
    Set concrete = New CConcreteDiagramMaterial
    concrete.Initialize -0.0015, -15.5, -0.0035, -15.5
    concrete.TensionMode = "Ignore"
    Set ProvisionalConcrete = concrete
End Function

Private Function ProvisionalSteel() As CSteelDiagramMaterial
    Dim steel As CSteelDiagramMaterial
    Set steel = New CSteelDiagramMaterial
    steel.Initialize 0.00175, 350#, 0.025
    Set ProvisionalSteel = steel
End Function

Private Sub AssertTrue(ByRef stats As TBatchTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name
    End If
End Sub

Private Sub AppendLine(ByRef stats As TBatchTestStats, ByVal text As String)
    stats.Report = stats.Report & text & vbCrLf
End Sub

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function


