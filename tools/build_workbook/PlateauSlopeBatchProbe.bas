' ==========================================================================
' ДЛЯ ТЕСТОВ — вспомогательный фрагмент частной приёмки; в выпуск не включается.
' ==========================================================================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторяет частный эксперимент с выбранными наклоном плато и стратегиями, затем сбрасывает экспериментальные параметры.
Public Function TestRunPlateauSlopeCases(ByVal specification As String) As String
    On Error GoTo Failed
    Dim parts As Variant, oldSystem As Variant, output As String, one As String, repeat As Long, alphaValue As Double
    parts = Split(specification, "|")
    If UBound(parts) <> 7 Then Err.Raise 5, , "Expected eight experiment arguments."
    alphaValue = Val(CStr(parts(0)))
    oldSystem = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula
    modTestPlateauSlopeProbe.TestPlateauBeginRun alphaValue, CStr(parts(1)) = "1"
    TestSetSystemSetting "Capacity.SolutionStrategy", CStr(parts(5))
    TestSetSystemSetting "SLS.Crack.InitiationSolutionStrategy", CStr(parts(6))
    For repeat = 1 To CLng(parts(7))
        one = TestRunFocusedAcceptanceCases(CStr(parts(2)), CStr(parts(3)), CStr(parts(4)))
        output = output & "PLATEAU_REPEAT|repeat=" & CStr(repeat) & "|spec=" & specification & vbCrLf & one
        If InStr(1, one, "RUNTIME ERROR", vbBinaryCompare) > 0 Then Exit For
    Next repeat
    output = output & modTestPlateauSlopeProbe.TestPlateauRecords()
    GoTo Restore
Failed:
    output = output & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Restore:
    If IsArray(oldSystem) Then ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula = oldSystem
    modTestPlateauSlopeProbe.Alpha = 0#: modTestPlateauSlopeProbe.SimpleInitialGuess = False
    TestRunPlateauSlopeCases = output
End Function
