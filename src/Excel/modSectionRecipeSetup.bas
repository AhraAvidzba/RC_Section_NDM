Attribute VB_Name = "modSectionRecipeSetup"
Option Explicit

' Явное обновление интерфейса рецептов: отступы, привязки рядов и список типов.
' Не вызывается расчётом. Пользовательские размеры, арматура, LC и Results сохраняются.
Public Sub UpdateSectionRecipeConfig()
    Dim guard As CExcelAppStateGuard
    Dim roundedTable As Object, hollowTable As Object, systemTable As Object, circleTable As Object, rectSetTable As Object
    Dim number As Long, source As String, description As String
    On Error GoTo Failed
    Set guard = New CExcelAppStateGuard
    guard.Enter ThisWorkbook.Application
    ' Все имена и заголовки проверяются до изменения первого диапазона.
    Set roundedTable = RequiredTable("rngRoundedRectangleGeometry")
    Set hollowTable = RequiredTable("rngHollowRectangleGeometry")
    Set systemTable = RequiredTable("rngSystemSettings")
    Set circleTable = RequiredTable("rngCircleGeometry")
    Set rectSetTable = RequiredTable("rngRectSetGeometry")
    ValidateSetbackTable roundedTable, 4
    ValidateSetbackTable hollowTable, 8
    Dim checkedRow As Long
    checkedRow = HeaderRow(systemTable, "Geometry.Type", 1)
    checkedRow = HeaderRow(circleTable, "Rebar.Count", 1)
    ValidateCircleBindings circleTable
    checkedRow = HeaderRow(rectSetTable, "as", 2)
    UpgradeSetbackTable roundedTable, 4
    UpgradeSetbackTable hollowTable, 8
    CanonicalRectSetHeaders rectSetTable
    ExplainClosedCircle circleTable
    UpgradeCircleBindings circleTable, rectSetTable.Columns.Count
    UpdateRecipeTypeList systemTable
    guard.Restore True
    Exit Sub
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    If Not guard Is Nothing Then guard.Restore
    Err.Raise number, source, description
End Sub

' Проверяет структуру круга до миграции; новые строки не заменяют пользовательские данные.
Private Sub ValidateCircleBindings(ByVal table As Object)
    Dim last As Long
    last = HeaderRow(table, "Rebar.Loc3row", 1)
    If table.Columns.Count <> 4 Then SectionRecipeError "modSectionRecipeSetup", "Неверная ширина таблицы круглого сечения."
    If table.Rows.Count = last Then Exit Sub
    If table.Rows.Count <> last + 2 Then SectionRecipeError "modSectionRecipeSetup", "Неполная таблица привязок рядов круга."
    If CStr(table.Cells(last + 1, 1).Value2) <> "Rebar.Bind2row" Or CStr(table.Cells(last + 2, 1).Value2) <> "Rebar.Bind3row" Then _
        SectionRecipeError "modSectionRecipeSetup", "Восстановите строки Rebar.Bind2row и Rebar.Bind3row в таблице круга."
End Sub

' Добавляет два независимых выбора у каждого/через один; остальные таблицы сдвигаются целиком.
Private Sub UpgradeCircleBindings(ByVal table As Object, ByVal geometryWidth As Long)
    Dim last As Long, row As Long, column As Long, target As Object, source As Object, expanded As Object
    Dim keys As Variant, comments As Variant, listName As String, index As Long
    last = HeaderRow(table, "Rebar.Loc3row", 1)
    keys = Array("Rebar.Bind2row", "Rebar.Bind3row")
    comments = Array("Второй ряд: EachBar — у каждого стержня первого ряда; EverySecondBar — у первого, третьего и следующих нечётных.", _
        "Третий ряд: EachBar — у каждого стержня первого ряда; EverySecondBar — у первого, третьего и следующих нечётных. Выбирается независимо от второго ряда.")
    If table.Rows.Count = last Then
        ' Сдвигаем всю ширину составных геометрических таблиц, сохраняя их поля и именованные ссылки.
        table.Cells(last + 1, 1).Resize(2, geometryWidth).Insert xlShiftDown
        Set expanded = table.Resize(last + 2, 4)
        ThisWorkbook.Names("rngCircleGeometry").RefersTo = "='" & Replace$(table.Worksheet.Name, "'", "''") & "'!" & expanded.Address(True, True, xlA1)
        For index = 0 To 1
            row = last + 1 + index
            For column = 1 To 4
                Set source = table.Cells(last, column): Set target = expanded.Cells(row, column)
                target.NumberFormat = source.NumberFormat: target.Interior.Color = source.Interior.Color
                target.Font.Name = source.Font.Name: target.Font.Size = source.Font.Size: target.Font.Color = source.Font.Color
                target.Font.Bold = source.Font.Bold: target.Locked = source.Locked
                target.HorizontalAlignment = source.HorizontalAlignment: target.VerticalAlignment = source.VerticalAlignment
                target.WrapText = source.WrapText: target.Borders.LineStyle = xlContinuous: target.Borders.Weight = xlThin
            Next column
            expanded.Cells(row, 1).Value2 = keys(index): expanded.Cells(row, 2).Value2 = "EachBar"
            expanded.Cells(row, 3).Value2 = "-": expanded.Cells(row, 4).Value2 = comments(index)
        Next index
    Else
        Set expanded = table
    End If
    For index = 0 To 1
        Set target = expanded.Cells(last + 1 + index, 2)
        target.HorizontalAlignment = xlCenter: target.VerticalAlignment = xlCenter
        listName = "rngCircleBind" & CStr(index + 2) & "Choices"
        CircleBindingList target, listName
    Next index
End Sub

' У каждого ряда собственный именованный список; повтор сохраняет выбранное значение.
Private Sub CircleBindingList(ByVal cell As Object, ByVal listName As String)
    Dim list As Object, column As Long
    On Error Resume Next
    Set list = ThisWorkbook.Names(listName).RefersToRange
    On Error GoTo 0
    If list Is Nothing Then
        column = cell.Worksheet.UsedRange.Column + cell.Worksheet.UsedRange.Columns.Count
        If column > cell.Worksheet.Columns.Count Then SectionRecipeError "modSectionRecipeSetup", "Нет свободной колонки для списка привязки ряда."
        Set list = cell.Worksheet.Cells(1, column).Resize(2, 1)
        list.Cells(1, 1).Value2 = "EachBar": list.Cells(2, 1).Value2 = "EverySecondBar"
        list.EntireColumn.Hidden = True
        ThisWorkbook.Names.Add Name:=listName, RefersTo:="='" & Replace$(cell.Worksheet.Name, "'", "''") & "'!" & list.Address(True, True, xlA1)
    End If
    cell.Validation.Delete
    cell.Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "=" & listName
    cell.Validation.IgnoreBlank = False: cell.Validation.InCellDropdown = True
End Sub

' Реестр является единственным источником выбора: новый класс и одна
' регистрация появляются в списке после явного обновления интерфейса.
Public Function RegisteredSectionRecipeTypes() As Variant
    Dim registry As CSectionTypeRegistry: Set registry = New CSectionTypeRegistry
    RegisteredSectionRecipeTypes = registry.RegisteredRecipeTypes
End Function

Private Function RequiredTable(ByVal name As String) As Object
    On Error Resume Next
    Set RequiredTable = ThisWorkbook.Names(name).RefersToRange
    On Error GoTo 0
    If RequiredTable Is Nothing Then SectionRecipeError "modSectionRecipeSetup", "Не найден диапазон '" & name & "'. Восстановите именованную таблицу Config."
End Function

Private Function HeaderRow(ByVal table As Object, ByVal key As String, ByVal column As Long) As Long
    Dim values As Variant, row As Long
    values = table.Value2
    For row = 1 To UBound(values, 1)
        If Not IsError(values(row, column)) Then
            If CStr(values(row, column)) = key Then HeaderRow = row: Exit Function
        End If
    Next row
    SectionRecipeError "modSectionRecipeSetup", "В таблице " & table.Worksheet.Name & "!" & table.Address(False, False) & _
        " не найден заголовок '" & key & "'. Восстановите структуру таблицы."
End Function

Private Sub ValidateSetbackTable(ByVal table As Object, ByVal count As Long)
    Dim row As Long, start As String, finish As String
    row = HeaderRow(table, "as", 2)
    If table.Columns.Count < 8 Or table.Rows.Count < row + count Then _
        SectionRecipeError "modSectionRecipeSetup", "Неполная таблица " & table.Worksheet.Name & "!" & table.Address(False, False) & "."
    start = CStr(table.Cells(row, 5).Value2): finish = CStr(table.Cells(row, 6).Value2)
    If (start = "StartSetback") Xor (finish = "EndSetback") Then _
        SectionRecipeError "modSectionRecipeSetup", "В таблице " & table.Worksheet.Name & "!" & table.Address(False, False) & _
            " обновлён только один заголовок отступов. Восстановите согласованную копию таблицы."
End Sub

Private Sub UpgradeSetbackTable(ByVal table As Object, ByVal count As Long)
    Dim row As Long, i As Long, target As Object, headers As Variant, values As Variant
    row = HeaderRow(table, "as", 2)
    If CStr(table.Cells(row, 5).Value2) = "StartSetback" Then Exit Sub
    ' Переносятся только прежние единицы и пояснения. A1-формулы сохраняются
    ' буквально; числа as/d/n и дополнительные ряды остаются в своих ячейках.
    Set target = table.Cells(row + 1, 5).Resize(count, 2)
    values = target.Formula
    table.Cells(row + 1, 7).Resize(count, 2).Formula = values
    target.Value2 = 0#
    headers = Array("StartSetback", "EndSetback", "Ед.", "Комментарий")
    For i = 0 To 3
        table.Cells(row, i + 5).Value2 = headers(i)
        table.Cells(row, i + 5).Font.Bold = True
        table.Cells(row, i + 5).Interior.Color = table.Cells(row, 2).Interior.Color
    Next i
    target.HorizontalAlignment = xlCenter: target.VerticalAlignment = xlCenter
    target.NumberFormat = "0.########"
    table.Cells(row, 8).Resize(count + 1, 1).HorizontalAlignment = xlLeft
    table.Cells(row, 1).Resize(count + 1, 8).Borders.LineStyle = xlContinuous
    table.Cells(row, 1).Resize(count + 1, 8).Borders.Weight = xlThin
    table.Cells(row, 1).Resize(1, 8).WrapText = True
    table.Cells(row + 1, 1).Resize(count, 8).WrapText = False
End Sub

Private Sub CanonicalRectSetHeaders(ByVal table As Object)
    Dim row As Long
    row = HeaderRow(table, "as", 2)
    table.Cells(row, 5).Value2 = "StartSetback"
    table.Cells(row, 6).Value2 = "EndSetback"
End Sub

Private Sub ExplainClosedCircle(ByVal table As Object)
    Dim row As Long, explanation As String
    row = HeaderRow(table, "Rebar.Count", 1)
    explanation = "ClosedUniform: StartSetback/EndSetback неприменимы к полному кольцу. Для отступов задайте открытый участок дуги в рецепте."
    If InStr(1, CStr(table.Cells(row, 4).Value2), "ClosedUniform", vbBinaryCompare) = 0 Then _
        table.Cells(row, 4).Value2 = CStr(table.Cells(row, 4).Value2) & " " & explanation
End Sub

Private Sub UpdateRecipeTypeList(ByVal systemTable As Object)
    Dim row As Long, cell As Object, list As Object, name As Object, ids As Variant, values() As Variant
    Dim count As Long, column As Long, i As Long, formula As String
    row = HeaderRow(systemTable, "Geometry.Type", 1): Set cell = systemTable.Cells(row, 2)
    ids = RegisteredSectionRecipeTypes: count = UBound(ids) - LBound(ids) + 1
    ReDim values(1 To count, 1 To 1)
    For i = 1 To count: values(i, 1) = ids(LBound(ids) + i - 1): Next i
    On Error Resume Next
    Set name = ThisWorkbook.Names("rngSectionRecipeTypes")
    If Not name Is Nothing Then Set list = name.RefersToRange
    On Error GoTo 0
    If list Is Nothing Then
        On Error Resume Next
        formula = cell.Validation.Formula1
        If Left$(formula, 1) = "=" Then Set list = cell.Worksheet.Range(Mid$(formula, 2))
        On Error GoTo 0
        If list Is Nothing Then
            column = cell.Worksheet.UsedRange.Column + cell.Worksheet.UsedRange.Columns.Count
            If column > cell.Worksheet.Columns.Count Then SectionRecipeError "modSectionRecipeSetup", "Нет свободной колонки для списка зарегистрированных типов."
            Set list = cell.Worksheet.Cells(1, column).Resize(count, 1)
        End If
    End If
    If list.Columns.Count <> 1 Then SectionRecipeError "modSectionRecipeSetup", "Список типов должен быть вертикальным."
    If list.Rows.Count > count Then list.Offset(count, 0).Resize(list.Rows.Count - count, 1).ClearContents
    Set list = list.Cells(1, 1).Resize(count, 1): list.Value2 = values
    list.EntireColumn.Hidden = True
    ThisWorkbook.Names.Add Name:="rngSectionRecipeTypes", RefersTo:="=" & list.Address(True, True, xlA1, True)
    cell.Validation.Delete
    cell.Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "=rngSectionRecipeTypes"
    cell.Validation.IgnoreBlank = False: cell.Validation.InCellDropdown = True
End Sub
