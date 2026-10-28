Attribute VB_Name = "modSectionRecipeSetup"
Option Explicit

' ==========================================================================
' Обновление списка зарегистрированных типов параметрических сечений
' ==========================================================================
' Меняет только список выбора Geometry.Type по общему реестру; размеры,
' таблицы ввода и оформление существующей книги не перестраивает.

' Обновляет встроенный список Geometry.Type по реестру без служебных диапазонов.
' Значение настройки, размеры, таблицы ввода и оформление книги не перестраиваются.
Public Sub UpdateSectionRecipeTypes()
    Dim table As Object, cell As Object, ids As Variant, choices As String
    Dim row As Long, i As Long, guard As CExcelAppStateGuard
    Dim number As Long, source As String, description As String
    On Error GoTo Failed
    Set table = RequiredTable("rngSystemSettings")
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
    ids = RegisteredSectionRecipeTypes
    choices = Join(ids, ",") ' Validation.Add принимает инвариантный разделитель; локальная «;» склеивает варианты в один пункт.
    If Len(choices) = 0 Or Len(choices) > 255 Then SectionRecipeError "modSectionRecipeSetup", _
        "Список зарегистрированных типов не помещается в проверку данных Geometry.Type."
    Set guard = New CExcelAppStateGuard: guard.Enter ThisWorkbook.Application
    cell.Validation.Delete
    cell.Validation.Add xlValidateList, xlValidAlertStop, xlBetween, choices
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
