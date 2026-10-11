VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionInputTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Неизменяемый табличный снимок для рецептов. Только значения и адреса;
' нет ссылок на Excel и знаний о Circle/RectSet или других типах сечений.
Private mName As String
Private mValues As Variant, mAddresses As Variant
Private mRows As Long, mColumns As Long

Public Sub Initialize(ByVal tableName As String, ByVal values As Variant, Optional ByVal addresses As Variant)
    If mRows > 0 Then InputError 0, 0, "Снимок таблицы уже создан; повторная инициализация запрещена."
    On Error GoTo Invalid
    If Len(Trim$(tableName)) = 0 Then GoTo Invalid
    Dim rows As Long, columns As Long, row As Long, column As Long, sourceRow As Long, sourceColumn As Long
    Dim copied() As Variant, locations() As String, value As Variant
    rows = UBound(values, 1) - LBound(values, 1) + 1
    columns = UBound(values, 2) - LBound(values, 2) + 1
    If rows <= 0 Or columns <= 0 Then GoTo Invalid
    If Not IsMissing(addresses) Then
        If UBound(addresses, 1) - LBound(addresses, 1) + 1 <> rows Or _
            UBound(addresses, 2) - LBound(addresses, 2) + 1 <> columns Then GoTo Invalid
    End If
    ReDim copied(1 To rows, 1 To columns): ReDim locations(1 To rows, 1 To columns)
    For row = 1 To rows
        sourceRow = LBound(values, 1) + row - 1
        For column = 1 To columns
            sourceColumn = LBound(values, 2) + column - 1
            If IsObject(values(sourceRow, sourceColumn)) Or IsArray(values(sourceRow, sourceColumn)) Then GoTo Invalid
            value = values(sourceRow, sourceColumn)
            copied(row, column) = value
            If IsMissing(addresses) Then
                locations(row, column) = tableName & "[" & CStr(row) & "," & CStr(column) & "]"
            Else
                locations(row, column) = CStr(addresses(LBound(addresses, 1) + row - 1, LBound(addresses, 2) + column - 1))
                If Len(Trim$(locations(row, column))) = 0 Then GoTo Invalid
            End If
        Next column
    Next row
    mValues = copied: mAddresses = locations
    mName = tableName: mRows = rows: mColumns = columns
    Exit Sub
Invalid:
    Err.Raise vbObjectError + 5525, "CSectionInputTable", "Таблица '" & tableName & "': требуется двумерный массив значений без объектов и согласованный массив адресов."
End Sub

Public Property Get Name() As String: Name = mName: End Property
Public Property Get RowCount() As Long: RowCount = mRows: End Property
Public Property Get ColumnCount() As Long: ColumnCount = mColumns: End Property

Public Function SourceReference(ByVal row As Long, ByVal column As Long) As String
    CheckCell row, column: SourceReference = mAddresses(row, column)
End Function

Public Function CellValue(ByVal row As Long, ByVal column As Long) As Variant
    CheckCell row, column
    If IsError(mValues(row, column)) Then InputError row, column, "Ячейка содержит ошибку Excel. Исправьте значение или формулу."
    CellValue = mValues(row, column)
End Function

Public Function IsBlank(ByVal row As Long, ByVal column As Long) As Boolean
    IsBlank = BlankValue(CellValue(row, column))
End Function

Public Function Text(ByVal row As Long, ByVal column As Long, Optional ByVal defaultValue As Variant) As String
    Dim value As Variant
    value = CellValue(row, column)
    If BlankValue(value) Then
        If IsMissing(defaultValue) Then InputError row, column, "Обязательное поле не заполнено."
        value = defaultValue
    End If
    If IsNull(value) Or IsError(value) Or IsObject(value) Or IsArray(value) Then InputError row, column, "Ожидается текстовое значение."
    Text = Trim$(CStr(value))
End Function

Public Function Number(ByVal row As Long, ByVal column As Long, Optional ByVal defaultValue As Variant) As Double
    Dim value As Variant
    value = CellValue(row, column)
    If BlankValue(value) Then
        If IsMissing(defaultValue) Then InputError row, column, "Обязательное числовое поле не заполнено."
        value = defaultValue
    End If
    If IsNull(value) Or IsError(value) Or IsObject(value) Or IsArray(value) Then InputError row, column, "Ожидается число."
    If VarType(value) = vbBoolean Or Not IsNumeric(value) Then InputError row, column, "Ожидается число."
    On Error GoTo Invalid
    Number = CDbl(value)
    Exit Function
Invalid:
    InputError row, column, "Число выходит за допустимый диапазон."
End Function

' Число позиций проверяется до выделения массивов и преобразования в Long.
Public Function Count(ByVal row As Long, ByVal column As Long, Optional ByVal defaultValue As Long = 0) As Long
    Dim value As Double
    value = Number(row, column, defaultValue)
    If value < 0# Or value > GEOM_MAX_REBAR_COUNT Or Fix(value) <> value Then _
        InputError row, column, "Число стержней должно быть целым неотрицательным и представимым счетчиком Long."
    Count = CLng(value)
End Function

Public Function Choice(ByVal row As Long, ByVal column As Long, ByVal allowed As Variant, _
        Optional ByVal defaultValue As String = vbNullString) As String
    Dim value As String, item As Variant
    value = Text(row, column, defaultValue)
    For Each item In allowed
        If StrComp(value, CStr(item), vbTextCompare) = 0 Then Choice = CStr(item): Exit Function
    Next item
    InputError row, column, "Недопустимый вариант '" & value & "'; допустимы: " & Join(allowed, ", ") & "."
End Function

Public Function Length(ByVal row As Long, ByVal column As Long, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then InputError row, column, "Не передана система единиц."
    Length = units.InputLengthToInternal(Number(row, column))
End Function

' Прежний пользовательский смысл: пусто/TODO/0 выключают ряд, отрицательное
' значение является ошибкой. Конвертация длины выполняется только CUnitSystem.
Public Function OptionalDiameter(ByVal row As Long, ByVal column As Long, ByVal units As CUnitSystem) As Double
    Dim raw As String, value As Double
    raw = Text(row, column, vbNullString)
    If Len(raw) = 0 Or UCase$(raw) = "TODO" Then Exit Function
    value = Number(row, column)
    If value < 0# Then InputError row, column, "Диаметр не может быть отрицательным; 0 или пустая ячейка выключают ряд."
    If units Is Nothing Then InputError row, column, "Не передана система единиц."
    OptionalDiameter = units.InputLengthToInternal(value)
End Function

Public Function ColumnByCaption(ByVal headerRow As Long, ByVal caption As String) As Long
    Dim column As Long, value As Variant
    CheckCell headerRow, 1
    For column = 1 To mColumns
        value = CellValue(headerRow, column)
        If Not BlankValue(value) Then
            If StrComp(Trim$(CStr(value)), caption, vbTextCompare) = 0 Then
                If ColumnByCaption > 0 Then InputError headerRow, column, "Повторный заголовок '" & caption & "'."
                ColumnByCaption = column
            End If
        End If
    Next column
    If ColumnByCaption = 0 Then InputError headerRow, 1, "Не найден столбец '" & caption & "'."
End Function

Public Function RowByKey(ByVal key As String, Optional ByVal keyColumn As Long = 1) As Long
    Dim row As Long, value As Variant
    CheckCell 1, keyColumn
    For row = 1 To mRows
        value = CellValue(row, keyColumn)
        If Not BlankValue(value) Then
            If StrComp(Trim$(CStr(value)), key, vbTextCompare) = 0 Then
                If RowByKey > 0 Then InputError row, keyColumn, "Повторный ключ '" & key & "'."
                RowByKey = row
            End If
        End If
    Next row
    If RowByKey = 0 Then InputError 1, keyColumn, "Не найден параметр '" & key & "'."
End Function

Private Function BlankValue(ByVal value As Variant) As Boolean
    If IsEmpty(value) Or IsNull(value) Then
        BlankValue = True
    ElseIf VarType(value) = vbString Then
        BlankValue = (Len(Trim$(CStr(value))) = 0)
    End If
End Function

Private Sub CheckCell(ByVal row As Long, ByVal column As Long)
    If row < 1 Or row > mRows Or column < 1 Or column > mColumns Then _
        InputError 0, 0, "Адрес поля вне границ снимка: строка " & CStr(row) & ", столбец " & CStr(column) & "."
End Sub

Private Sub InputError(ByVal row As Long, ByVal column As Long, ByVal reason As String)
    Dim location As String
    location = mName
    If row >= 1 And row <= mRows And column >= 1 And column <= mColumns Then location = mAddresses(row, column) & " (" & mName & ")"
    Err.Raise vbObjectError + 5525, "CSectionInputTable", location & ": " & reason
End Sub
