' ДЛЯ ТЕСТОВ: один штатный batch из 15 разных сочетаний, включая N, M и N+Mx+My.
Public Function RunPlateauBatch15() As String
    On Error GoTo Failed
    Dim stats As TBatchTestStats, oldSystem As Variant, oldProfiles As Variant, oldLoads As Variant
    Dim systemRange As Object, profileRange As Object, loadRange As Object
    Dim settings As CSystemSettingsReader, units As CUnitSystem, section As CSectionModel
    Dim allCases As Collection, cases As New Collection, item As Variant, names As Variant
    Dim values() As Variant, i As Long, started As Double, elapsed As Double, message As String
    Dim result As CCombinationResult
    Set systemRange = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set profileRange = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Set loadRange = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    oldSystem = systemRange.Formula: oldProfiles = profileRange.Formula: oldLoads = loadRange.Formula
    FocusedSelectGeometry "SavedWorkbook"
    SetSystemSetting "Load.ReferenceOffsetX", "0"
    SetSystemSetting "Load.ReferenceOffsetY", "0"
    SetSystemSetting "Calculation.ZeroMomentPerDepth", "0"
    SetSystemSetting "General.ExecutionReportEnabled", "Yes"
    SetSystemSetting "SLS.Crack.Code", "SP35"
    SetSystemSetting "Plot.Enabled", "Yes"
    SetSystemSetting "Plot.AutoUpdateAfterCalculation", "Yes"
    SetSystemSetting "Plot.LoadCase", "Worst"
    SetProfileValue "Calculation.Strength.DirectState", "PR3", "Yes"
    SetProfileValue "Calculation.Strength.Capacity", "PR3", "Yes"
    SetProfileValue "Calculation.Crack.Width", "PR3", "Yes"
    SetProfileValue "Calculation.Stability.Enabled", "PR3", "No"
    SetProfileValue "Visualization.State", "PR3", "StrengthState"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    SetSystemSetting "Load.ReferenceOffsetX", FormatNumberInvariant(10# / units.InputLengthToInternal(1#))
    SetSystemSetting "Load.ReferenceOffsetY", FormatNumberInvariant(-7# / units.InputLengthToInternal(1#))
    settings.LoadFromWorkbook ThisWorkbook
    Set section = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    Set allCases = FocusedSelectedLoads(section, "Saved")
    names = Array("axialSmall-1", "axialSmall1", "Mx-1", "Mx1", "My-1", "My1", "mixedC1_-1", "mixedT-1_1", _
        "tinyMx1", "biaxial1_-1_10", "biaxial1_-1_" & CStr(0.000001), "N500", "N1000", "N500TinyMoment", "N2098Plateau")
    For Each item In allCases
        If FocusedNameIn(CStr(item(0)), names) Then cases.Add item
    Next item
    If cases.Count <> 15 Then Err.Raise 5, , "Batch must contain exactly 15 selected loads."
    ReDim values(1 To loadRange.Rows.Count - 1, 1 To loadRange.Columns.Count)
    i = 0
    For Each item In cases
        i = i + 1
        values(i, 1) = CStr(item(0))
        values(i, 2) = CDbl(item(1)) / units.InputForceToInternal(1#)
        values(i, 3) = CDbl(item(2)) / units.InputMomentMxToInternal(1#)
        values(i, 4) = CDbl(item(3)) / units.InputMomentMyToInternal(1#)
        values(i, 5) = "PR3": values(i, 6) = "Auto"
    Next item
    loadRange.Offset(1, 0).Resize(loadRange.Rows.Count - 1, loadRange.Columns.Count).Value2 = values
    Set mFocusedBatch = Nothing: Set mFocusedSection = Nothing
    modPlateauSlopeProbe.ResetCounters
    started = Timer
    message = FixtureRunSectionWorkflow(ThisWorkbook, False)
    elapsed = Timer - started: If elapsed < 0# Then elapsed = elapsed + 86400#
    AssertTrue stats, "plateau.batch15.capture", Not mFocusedBatch Is Nothing
    If Not mFocusedBatch Is Nothing Then
        AssertTrue stats, "plateau.batch15.count", mFocusedBatch.Count = 15
        AssertTrue stats, "plateau.batch15.preparationReleased", mFocusedBatch.RunPreparationReleased
        AssertTrue stats, "plateau.batch15.noRepeatedMaterialPreparation", mFocusedBatch.MechanicalPreparationBuildCount <= 2
        AssertTrue stats, "plateau.batch15.oneCrackDomain", mFocusedBatch.CrackDomainBuildCount <= 1
        modPlateauSlopeProbe.CaptureResult "__batch15_total__", elapsed, mFocusedBatch.ResultAt(1)
        modPlateauSlopeProbe.ResetCounters
        For i = 1 To mFocusedBatch.Count
            Set result = mFocusedBatch.ResultAt(i)
            item = cases(i)
            modPlateauSlopeProbe.CaptureResult "batch15." & CStr(item(0)), 0#, result
            FocusedCheckStates stats, result, settings, units, "plateau.batch15." & CStr(item(0))
            Audit03CheckResultComments stats, mFocusedBatch, i
        Next i
    End If
    AppendLine stats, "TOTAL_PLATEAU_BATCH15: seconds=" & FormatNumberInvariant(elapsed) & _
        "; passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    GoTo Restore
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: plateau.batch15.runtime; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    If IsArray(oldSystem) Then systemRange.Formula = oldSystem
    If IsArray(oldProfiles) Then profileRange.Formula = oldProfiles
    If IsArray(oldLoads) Then loadRange.Formula = oldLoads
    Set mFocusedBatch = Nothing: Set mFocusedSection = Nothing
    RunPlateauBatch15 = stats.Report
End Function
