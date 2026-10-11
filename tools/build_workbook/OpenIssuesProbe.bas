' ДЛЯ ТЕСТОВ: добавляется только к направленному probe на отдельной книге.
Public Function RunOpenIssuesReproduction() As String
    On Error GoTo Failed
    Dim stats As TBatchTestStats
    Dim oldStrategy As String, settings As CSystemSettingsReader, units As CUnitSystem
    Dim section As CSectionModel, item As Variant, cases As New Collection, output As String
    oldStrategy = GetSystemSetting("SLS.Crack.InitiationSolutionStrategy")
    FocusedSelectGeometry "SavedWorkbook"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set section = BuildWorkbookSectionModel(ThisWorkbook, settings, units)
    For Each item In Audit03NormalizedLoads(section, "Light")
        If CStr(item(0)) = "My1" Then cases.Add item
    Next item
    If cases.Count <> 1 Then Err.Raise 5, , "My fixture not found."
    Dim loads As Object, oldLoads As Variant, values() As Variant, factor As Variant
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    oldLoads = loads.Formula
    SetSystemSetting "Load.ReferenceOffsetX", FormatNumberInvariant(10# / units.InputLengthToInternal(1#))
    SetSystemSetting "Load.ReferenceOffsetY", FormatNumberInvariant(-7# / units.InputLengthToInternal(1#))
    SetSystemSetting "Calculation.ZeroMomentPerDepth", "0"
    SetSystemSetting "SLS.Crack.Code", "SP35"
    SetSystemSetting "SLS.Crack.InitiationSolutionStrategy", "Auto"
    SetProfileValue "Calculation.Strength.DirectState", "PR3", "Yes"
    SetProfileValue "Calculation.Strength.Capacity", "PR3", "Yes"
    SetProfileValue "Calculation.Crack.Width", "PR3", "Yes"
    SetProfileValue "Calculation.Stability.Enabled", "PR3", "No"
    item = cases(1)
    For Each factor In Array(12#, 16#)
        ReDim values(1 To loads.Rows.Count - 1, 1 To loads.Columns.Count)
        values(1, 1) = "opening_My_" & CStr(factor)
        values(1, 2) = CDbl(item(1)) * CDbl(factor) / units.InputForceToInternal(1#)
        values(1, 3) = CDbl(item(2)) * CDbl(factor) / units.InputMomentMxToInternal(1#)
        values(1, 4) = CDbl(item(3)) * CDbl(factor) / units.InputMomentMyToInternal(1#)
        values(1, 5) = "PR3": values(1, 6) = "Auto"
        loads.Offset(1, 0).Resize(loads.Rows.Count - 1, loads.Columns.Count).Value2 = values
        Set mFocusedBatch = Nothing
        Dim message As String
        message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
        If mFocusedBatch Is Nothing Then Err.Raise 5, , "Batch not captured."
        output = output & "MY_REPRO|factor=" & CStr(factor) & "|" & mFocusedBatch.ResultAt(1).CrackWidthMeta.ResultComment & vbCrLf
        AssertTrue stats, "my." & CStr(factor) & ".widthCalculated", mFocusedBatch.ResultAt(1).CrackWidthMeta.Calculated
        AssertTrue stats, "my." & CStr(factor) & ".noInputError", mFocusedBatch.ResultAt(1).CrackWidthMeta.InternalStatus <> rsInvalidInput
    Next factor
    SetSystemSetting "SLS.Crack.InitiationSolutionStrategy", "UltimateStrain"
    Dim pathReport As String
    pathReport = RunFocusedAcceptanceCases("SavedWorkbook", "Paths", "pathN")
    output = output & pathReport
    AssertTrue stats, "pathN.formation.physicalNotReached", InStr(1, pathReport, "formation|status=rsSuccessWithWarning|code=CRACK_NOT_FORMED", vbBinaryCompare) > 0
    AssertTrue stats, "pathN.formation.noNumFail", InStr(1, pathReport, "FOCUSED_UNRESOLVED", vbBinaryCompare) = 0
    GoTo Restore
Failed:
    output = output & "FAIL: reproduction; " & CStr(Err.Number) & "; " & Err.Description
Restore:
    If IsArray(oldLoads) Then loads.Formula = oldLoads
    SetSystemSetting "SLS.Crack.InitiationSolutionStrategy", oldStrategy
    RunOpenIssuesReproduction = output & stats.Report & "TOTAL_OPEN_ISSUES: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
End Function
