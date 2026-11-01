Attribute VB_Name = "modSnapshotTables"
Option Explicit
' ==========================================================================
' Общие контракты неизменяемых таблиц полного снимка и адресного чтения
' ==========================================================================
' Здесь нет файлового или Excel-ввода; проверки и объединения не выполняют расчёт.

' Проверяет размеры готовой таблицы, не меняя её содержимого.
Public Sub ValidateSnapshotTableShape(ByVal key As String, ByRef values As Variant)
    Dim columns As Long, rows As Long
    Select Case key
        Case "rngNDMElementResults": columns = 9
        Case "rngNDMSectionGeometry": columns = 16
        Case "rngNDMSectionContours": columns = 17
        Case "rngNDMElementGeometry": columns = 18
        Case "rngNDMSectionProperties", "LoadCase.Properties": columns = 6
        Case "rngNDMMaterialDiagrams": columns = 11
        Case "rngNDMSectionAnnotations", "LoadCase.Annotations": columns = 24
        Case "Summary"
            columns = 33
            If UBound(values, 2) = 26 Then columns = 26 ' Формат сводки из 26 колонок не содержит блока исходных нагрузок.
        Case "Strength.Headers": columns = 49: rows = 4
        Case "Strength.Data": columns = 49
        Case "Cracks.Headers": columns = 72: rows = 4
        Case "Cracks.Data": columns = 72
        Case "Stability.Labels": columns = 84: rows = 1
        Case "Stability.Data": columns = 84
        Case "Calculation.Profiles"
            If UBound(values, 2) < 3 Or UBound(values, 1) < 1 Then Err.Raise 5, "CCalculationSnapshot", "Повреждены сохранённые профили расчёта."
            Exit Sub
    End Select
    If UBound(values, 2) <> columns Or UBound(values, 1) < 1 Then Err.Raise 5, "CCalculationSnapshot", "Неверная структура таблицы снимка: " & key
    If rows > 0 Then If UBound(values, 1) <> rows Then Err.Raise 5, "CCalculationSnapshot", "Неверная шапка таблицы снимка: " & key
End Sub


' Восстанавливает только сохранённые настройки профиля просмотра.
Public Function SnapshotVisualizationProfile(ByRef values As Variant, ByVal profileID As String) As CCalculationProfile
    Dim column As Long, selected As Long, row As Long
    For row = 1 To UBound(values, 1)
        For column = 3 To UBound(values, 2)
            If StrComp(CStr(values(row, column)), profileID, vbTextCompare) = 0 Then selected = column: Exit For
        Next column
        If selected > 0 Then Exit For
    Next row
    If selected = 0 Then Err.Raise 5, , "В снимке отсутствует сохранённый профиль: " & profileID
    Dim fields As Object: Set fields = CreateObject("Scripting.Dictionary"): fields.CompareMode = vbTextCompare
    For row = 1 To UBound(values, 1)
        If Len(CStr(values(row, 2))) > 0 Then fields.Item(CStr(values(row, 2))) = CStr(values(row, selected))
    Next row
    Dim profile As CCalculationProfile: Set profile = New CCalculationProfile
    profile.Initialize profileID, CStr(fields.Item("Profile.DisplayName")), CStr(fields.Item("Profile.Description")), _
        CStr(fields.Item("Calculation.Strength.DirectState")) = "Yes", CStr(fields.Item("Calculation.Strength.Capacity")) = "Yes", _
        CStr(fields.Item("Calculation.Crack.Width")) = "Yes", CStr(fields.Item("Calculation.Stability.Enabled")) = "Yes"
    profile.SetVisualization CStr(fields.Item("Visualization.State")), CStr(fields.Item("Visualization.Quantity")), _
        CStr(fields.Item("Visualization.StressPrecision")), CStr(fields.Item("Visualization.StrainPrecision"))
    Set SnapshotVisualizationProfile = profile
End Function


' Сравнивает исходную геометрию перед повторным использованием проверенных границ.
Public Function SnapshotMatricesEqual(ByRef actual As Variant, ByRef expected As Variant) As Boolean
    On Error GoTo Different
    If Not IsArray(actual) Or Not IsArray(expected) Then Exit Function
    If LBound(actual, 1) <> 1 Or LBound(actual, 2) <> 1 Then Exit Function
    If UBound(actual, 1) <> UBound(expected, 1) Or UBound(actual, 2) <> UBound(expected, 2) Then Exit Function
    Dim row As Long, column As Long
    For row = 1 To UBound(actual, 1)
        For column = 1 To UBound(actual, 2)
            If VarType(actual(row, column)) <> VarType(expected(row, column)) Then Exit Function
            If IsError(actual(row, column)) Or IsNull(actual(row, column)) Or IsObject(actual(row, column)) Or IsArray(actual(row, column)) Then Exit Function
            If actual(row, column) <> expected(row, column) Then Exit Function
        Next column
    Next row
    SnapshotMatricesEqual = True
    Exit Function
Different:
    Err.Clear
End Function


' Объединяет общие и LC-зависимые строки в исходном порядке.
Public Function JoinSnapshotTables(ByRef common As Variant, ByRef cases As Variant) As Variant
    Dim joined() As Variant, i As Long, j As Long
    ReDim joined(1 To UBound(common, 1) + UBound(cases, 1) - 1, 1 To UBound(common, 2))
    For i = 1 To UBound(common, 1)
        For j = 1 To UBound(common, 2): joined(i, j) = common(i, j): Next j
    Next i
    For i = 2 To UBound(cases, 1)
        For j = 1 To UBound(cases, 2): joined(UBound(common, 1) + i - 1, j) = cases(i, j): Next j
    Next i
    JoinSnapshotTables = joined
End Function


' Проверяет принадлежность используемой таблицы текущему расчёту.
Public Sub ValidateSnapshotRunID(ByRef values As Variant, ByVal key As String, ByVal runID As Double)
    Dim row As Long
    If Left$(key, 6) <> "rngNDM" And Left$(key, 9) <> "LoadCase." Then Exit Sub
    For row = 2 To UBound(values, 1)
        If Len(CStr(values(row, 1))) > 0 Then
            If CDbl(values(row, 1)) <> runID Then Err.Raise 5, , "Таблицы относятся к разным запускам: " & key
        End If
    Next row
End Sub
