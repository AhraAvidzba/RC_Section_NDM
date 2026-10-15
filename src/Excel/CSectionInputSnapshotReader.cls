VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionInputSnapshotReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Чтение входных таблиц рецепта из Excel
' ==========================================================================
' Читает каждую зарегистрированную таблицу одним обращением Value2
' и сохраняет значения с адресами ячеек для диагностики.
' Структуру таблиц разбирает рецепт; ссылки на Range и Workbook в него не передаются.

' Одним чтением Value2 получает значения каждой зарегистрированной таблицы
' и сохраняет адреса её ячеек для ошибок ввода, без передачи объектов Excel рецепту.
Public Function Capture(ByVal workbook As Object, ByVal tableNames As Variant, _
        ByVal units As CUnitSystem) As CSectionRecipeContext
    Dim result As CSectionRecipeContext, name As Variant, range As Object
    Dim values As Variant, addresses() As String, row As Long, column As Long
    Dim firstRow As Long, firstColumn As Long, sheetName As String, table As CSectionInputTable
    Set result = New CSectionRecipeContext: result.Initialize units
    For Each name In tableNames
        Set range = Nothing
        On Error Resume Next
        Set range = workbook.Names(CStr(name)).RefersToRange
        On Error GoTo 0
        If range Is Nothing Then SectionRecipeError "CSectionInputSnapshotReader", "Не найден диапазон '" & CStr(name) & "'."
        If range.Areas.Count <> 1 Or range.Cells.Count = 1 Then _
            SectionRecipeError "CSectionInputSnapshotReader", "Диапазон '" & CStr(name) & "' должен быть единой таблицей."
        ' В ручном режиме обновляем формулы выбранной таблицы до одного Value2.
        ' Неактивные таблицы и посторонние книги не пересчитываются.
        If workbook.Application.Calculation = xlCalculationManual Then range.Calculate
        values = range.Value2
        firstRow = range.Row: firstColumn = range.Column: sheetName = range.Worksheet.Name
        ReDim addresses(1 To UBound(values, 1), 1 To UBound(values, 2))
        For column = 1 To UBound(values, 2)
            For row = 1 To UBound(values, 1)
                addresses(row, column) = sheetName & "!" & ColumnName(firstColumn + column - 1) & CStr(firstRow + row - 1)
            Next row
        Next column
        Set table = New CSectionInputTable: table.Initialize CStr(name), values, addresses
        result.AddTable table
    Next name
    result.Seal
    Set Capture = result
End Function

' Переводит номер столбца Excel в буквенный адрес для привязки ошибки к исходной ячейке.
Private Function ColumnName(ByVal column As Long) As String
    Do While column > 0
        column = column - 1
        ColumnName = Chr$(65 + column Mod 26) & ColumnName
        column = column \ 26
    Loop
End Function
