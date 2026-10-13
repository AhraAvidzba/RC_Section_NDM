' ДЛЯ ТЕСТОВ: контроль коэффициента на отдельной книге со штатным запуском,
' записью Results, отчётом, схемой и независимым повторным интегрированием НДС.
Public Function RunPlateauFactorControlCases(ByVal factorText As String) As String
    On Error GoTo Failed
    Dim original As Variant, settings As CSystemSettingsReader, body As String
    Dim loadName As Variant, one As String, completed As Long
    original = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    If settings.HasKey("General.PlateauSlopeFactor") Then
        SetSystemSetting "General.PlateauSlopeFactor", factorText
    Else
        ' Исходная книга до изменения имеет согласованный фиксированный наклон.
        If factorText <> "0.0001" Then Err.Raise 5, , "Baseline supports only the former factor 0.0001."
        SetSystemSetting "Solver.PlateauSlope.Enabled", "Yes"
    End If
    For Each loadName In Array("axialSmall-1", "axialSmall1", "Mx-1", "Mx1", "My-1", "My1", _
            "mixedC1_-1", "mixedT-1_1", "N500", "N1000", "N2098Plateau", "N500TinyMoment")
        one = RunFocusedAcceptanceCases("SavedWorkbook", "Saved", CStr(loadName))
        body = body & "FACTOR_CONTROL|factor=" & factorText & "|id=" & CStr(loadName) & vbCrLf & one & vbCrLf
        completed = completed + 1
        If InStr(1, one, "RUNTIME ERROR", vbBinaryCompare) > 0 Then Exit For
    Next loadName
    body = body & "TOTAL_FACTOR_CONTROLS: factor=" & factorText & "; cases=" & CStr(completed) & vbCrLf
    GoTo Restore
Failed:
    body = body & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Restore:
    If IsArray(original) Then ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula = original
    RunPlateauFactorControlCases = body
End Function
