' ДЛЯ ТЕСТОВ: добавляется в test module изолированной книги.
' alpha, simple, shape, group, selected load, capacity strategy, formation strategy, repetitions.
Public Function RunPlateauSlopeCases(ByVal specification As String) As String
    On Error GoTo Failed
    Dim parts As Variant, oldSystem As Variant, output As String, one As String, repeat As Long, alphaValue As Double
    parts = Split(specification, "|")
    If UBound(parts) <> 7 Then Err.Raise 5, , "Expected eight experiment arguments."
    alphaValue = Val(CStr(parts(0)))
    oldSystem = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula
    modPlateauSlopeProbe.BeginRun alphaValue, CStr(parts(1)) = "1"
    SetSystemSetting "Capacity.SolutionStrategy", CStr(parts(5))
    SetSystemSetting "SLS.Crack.InitiationSolutionStrategy", CStr(parts(6))
    For repeat = 1 To CLng(parts(7))
        one = RunFocusedAcceptanceCases(CStr(parts(2)), CStr(parts(3)), CStr(parts(4)))
        output = output & "PLATEAU_REPEAT|repeat=" & CStr(repeat) & "|spec=" & specification & vbCrLf & one
        If InStr(1, one, "RUNTIME ERROR", vbBinaryCompare) > 0 Then Exit For
    Next repeat
    output = output & modPlateauSlopeProbe.Records()
    GoTo Restore
Failed:
    output = output & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Restore:
    If IsArray(oldSystem) Then ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula = oldSystem
    modPlateauSlopeProbe.Alpha = 0#: modPlateauSlopeProbe.SimpleInitialGuess = False
    RunPlateauSlopeCases = output
End Function
