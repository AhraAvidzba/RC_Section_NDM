Attribute VB_Name = "modSectionRecipeSetup"
Option Explicit

' Обновляет только список Geometry.Type по реестру; структура готовой книги уже актуальна.
Public Sub UpdateSectionRecipeTypes()
    Dim table As Object, list As Object, target As Object, cell As Object, ids As Variant, values() As Variant
    Dim count As Long, row As Long, i As Long, guard As CExcelAppStateGuard
    Dim number As Long, source As String, description As String
    On Error GoTo Failed
    Set table = RequiredTable("rngSystemSettings")
    Set list = RequiredTable("rngSectionRecipeTypes")
    If list.Areas.Count <> 1 Or list.Columns.Count <> 1 Then _
        SectionRecipeError "modSectionRecipeSetup", "Именованный список rngSectionRecipeTypes должен занимать один непрерывный столбец."
    If Not list.Worksheet Is table.Worksheet Then _
        SectionRecipeError "modSectionRecipeSetup", "Список rngSectionRecipeTypes должен находиться на листе системных настроек."
    For i = 2 To table.Rows.Count
        If Not IsError(table.Cells(i, 1).Value2) Then
            If CStr(table.Cells(i, 1).Value2) = "Geometry.Type" Then
                If row > 0 Then SectionRecipeError "modSectionRecipeSetup", "В rngSystemSettings повторяется ключ Geometry.Type."
                row = i
            End If
        End If
    Next i
    If row = 0 Then SectionRecipeError "modSectionRecipeSetup", "В rngSystemSettings отсутствует ключ Geometry.Type."
    Set cell = table.Cells(row, 2)
    ids = RegisteredSectionRecipeTypes: count = UBound(ids) - LBound(ids) + 1
    Set target = list.Cells(1, 1).Resize(count, 1)
    ' Новый тип занимает только свободное продолжение служебного списка; чужие данные не заменяются.
    For i = list.Rows.Count + 1 To count
        If Len(CStr(target.Cells(i, 1).Formula)) > 0 Then _
            SectionRecipeError "modSectionRecipeSetup", "Продолжение списка типов занято: " & target.Worksheet.Name & "!" & target.Cells(i, 1).Address(False, False) & ". Освободите ячейку или перенесите именованный список."
    Next i
    ReDim values(1 To count, 1 To 1)
    For i = 1 To count: values(i, 1) = ids(LBound(ids) + i - 1): Next i
    Set guard = New CExcelAppStateGuard: guard.Enter ThisWorkbook.Application
    If list.Rows.Count > count Then list.Cells(count + 1, 1).Resize(list.Rows.Count - count, 1).ClearContents
    target.Value2 = values
    If list.Rows.Count <> count Then ThisWorkbook.Names("rngSectionRecipeTypes").RefersTo = _
        "='" & Replace$(target.Worksheet.Name, "'", "''") & "'!" & target.Address(True, True, xlA1)
    cell.Validation.Delete
    cell.Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "=rngSectionRecipeTypes"
    cell.Validation.IgnoreBlank = False: cell.Validation.InCellDropdown = True
    guard.Restore True
    Exit Sub
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    If Not guard Is Nothing Then guard.Restore
    Err.Raise number, source, description
End Sub

' Читает реестр без обращения к таблицам геометрии или запуска построения сечения.
Public Function RegisteredSectionRecipeTypes() As Variant
    Dim registry As CSectionTypeRegistry: Set registry = New CSectionTypeRegistry
    RegisteredSectionRecipeTypes = registry.RegisteredRecipeTypes
End Function

' Повреждённое имя требует восстановления; отсутствующие таблицы здесь не создаются.
Private Function RequiredTable(ByVal name As String) As Object
    On Error Resume Next
    Set RequiredTable = ThisWorkbook.Names(name).RefersToRange
    On Error GoTo 0
    If RequiredTable Is Nothing Then SectionRecipeError "modSectionRecipeSetup", "Не найден диапазон '" & name & "'. Восстановите именованный диапазон Config."
End Function
