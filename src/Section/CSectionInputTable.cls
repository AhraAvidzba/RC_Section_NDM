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

' Копирует двумерные scalar values и согласованные addresses один раз; Excel-объекты не
' сохраняет.
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

' Имя корневого блока snapshot.
Public Property Get Name() As String: Name = mName: End Property
' Число строк неизменяемого input snapshot.
Public Property Get RowCount() As Long: RowCount = mRows: End Property
' Число столбцов input snapshot.
Public Property Get ColumnCount() As Long: ColumnCount = mColumns: End Property

' Возвращает исходный адрес ячейки или программный tableName[row,col].
Public Function SourceReference(ByVal row As Long, ByVal column As Long) As String
    CheckCell row, column: SourceReference = mAddresses(row, column)
End Function

' Читает scalar snapshot; ошибка Excel в ячейке даёт адресное исключение.
Public Function CellValue(ByVal row As Long, ByVal column As Long) As Variant
    CheckCell row, column
    If IsError(mValues(row, column)) Then InputError row, column, "Ячейка содержит ошибку Excel. Исправьте значение или формулу."
    CellValue = mValues(row, column)
End Function

' Проверяет пустое значение снимка; ошибка Excel не подменяется пустотой.
Public Function IsBlank(ByVal row As Long, ByVal column As Long) As Boolean
    IsBlank = BlankValue(CellValue(row, column))
End Function

' Читает непустой текст либо явно переданный fallback, обрезая крайние пробелы.
Public Function Text(ByVal row As Long, ByVal column As Long, Optional ByVal defaultValue As Variant) As String
    Dim value As Variant
    value = CellValue(row, column)
    If BlankValue(value) Then
        If IsMissing(defaultValue) Then InputError row, column, "Обязательное поле не заполнено. Введите значение."
        value = defaultValue
    End If
    If IsNull(value) Or IsError(value) Or IsObject(value) Or IsArray(value) Then InputError row, column, "Ожидается текстовое значение. Введите текст."
    Text = Trim$(CStr(value))
End Function

' Читает число, исключая Boolean/ошибки/объекты; обязательное пустое поле не заменяется нулём.
Public Function Number(ByVal row As Long, ByVal column As Long, Optional ByVal defaultValue As Variant) As Double
    Dim value As Variant
    value = CellValue(row, column)
    If BlankValue(value) Then
        If IsMissing(defaultValue) Then InputError row, column, "Обязательное числовое поле не заполнено. Введите число."
        value = defaultValue
    End If
    If IsNull(value) Or IsError(value) Or IsObject(value) Or IsArray(value) Then InputError row, column, "Ожидается число. Введите числовое значение."
    If VarType(value) = vbBoolean Or Not IsNumeric(value) Then InputError row, column, "Ожидается число. Введите числовое значение."
    On Error GoTo Invalid
    Number = CDbl(value)
    Exit Function
Invalid:
    InputError row, column, "Число выходит за допустимый диапазон. Введите число меньшего модуля."
End Function

' Читает весь неотрицательный диапазон Long. Допустимость общего числа
' активных позиций проверяет executor до выделения массива или цикла.
Public Function Count(ByVal row As Long, ByVal column As Long, Optional ByVal defaultValue As Long = 0) As Long
    Dim value As Double
    value = Number(row, column, defaultValue)
    If value < 0# Or value > 2147483647# Or Fix(value) <> value Then _
        InputError row, column, "Число стержней должно быть целым неотрицательным и представимым счетчиком Long. Введите целое число от 0 до 2147483647."
    Count = CLng(value)
End Function

' Проверяет вариант по allowed без учёта регистра и возвращает каноническое значение allowed.
Public Function Choice(ByVal row As Long, ByVal column As Long, ByVal allowed As Variant, _
        Optional ByVal defaultValue As String = vbNullString) As String
    Dim value As String, item As Variant
    value = Text(row, column, defaultValue)
    For Each item In allowed
        If StrComp(value, CStr(item), vbTextCompare) = 0 Then Choice = CStr(item): Exit Function
    Next item
    InputError row, column, "Недопустимый вариант '" & value & "'. Выберите: " & Join(allowed, ", ") & "."
End Function

' Читает обязательное число и один раз переводит input length в мм через CUnitSystem.
Public Function Length(ByVal row As Long, ByVal column As Long, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then InputError row, column, "Не передана система единиц."
    Length = units.InputLengthToInternal(Number(row, column))
End Function

' Читает и переводит длину единственным Length, затем проверяет строго положительный размер.
Public Function PositiveLength(ByVal row As Long, ByVal column As Long, ByVal units As CUnitSystem) As Double
    PositiveLength = Length(row, column, units)
    If PositiveLength <= 0# Then InputError row, column, "Размер должен быть положительным. Введите значение больше нуля."
End Function

' Нулевой радиус или отступ допустим; пустота и недопустимый текст не становятся нулём.
Public Function NonNegativeLength(ByVal row As Long, ByVal column As Long, ByVal units As CUnitSystem) As Double
    NonNegativeLength = Length(row, column, units)
    If NonNegativeLength < 0# Then InputError row, column, "Длина не может быть отрицательной. Введите значение не меньше нуля."
End Function

' Пустая ячейка или ноль выключают ряд. Текст и отрицательный диаметр
' отклоняются; положительный диаметр переводится в мм через CUnitSystem.
Public Function OptionalDiameter(ByVal row As Long, ByVal column As Long, ByVal units As CUnitSystem) As Double
    Dim raw As String, value As Double
    raw = Text(row, column, vbNullString)
    If Len(raw) = 0 Then Exit Function
    value = Number(row, column)
    If value < 0# Then InputError row, column, "Диаметр не может быть отрицательным; 0 или пустая ячейка выключают ряд. Введите неотрицательное значение."
    If units Is Nothing Then InputError row, column, "Не передана система единиц."
    OptionalDiameter = units.InputLengthToInternal(value)
End Function

' Находит уникальный столбец заголовка, без фиксированного адреса.
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
    If ColumnByCaption = 0 Then InputError headerRow, 1, "Не найден столбец '" & caption & "'. Восстановите заголовок столбца."
End Function

' Находит уникальную строку смыслового ключа в указанном подблоке.
Public Function RowByKey(ByVal key As String, Optional ByVal keyColumn As Long = 1, _
        Optional ByVal firstRow As Long = 1, Optional ByVal lastRow As Long = 0) As Long
    RowByKey = RowByAnyKey(Array(key), keyColumn, firstRow, lastRow)
End Function

' Подблок и допустимые подписи задает рецепт. Снимок знает лишь границы
' поиска; совпадения в другом подблоке не делают локальный ключ неоднозначным.
Public Function RowByAnyKey(ByVal keys As Variant, Optional ByVal keyColumn As Long = 1, _
        Optional ByVal firstRow As Long = 1, Optional ByVal lastRow As Long = 0) As Long
    Dim row As Long, value As Variant, key As Variant, caption As String
    If lastRow = 0 Then lastRow = mRows
    CheckCell firstRow, keyColumn: CheckCell lastRow, keyColumn
    If lastRow < firstRow Then InputError firstRow, keyColumn, "Конец подблока находится раньше начала."
    If Not IsArray(keys) Then InputError firstRow, keyColumn, "Допустимые подписи должны быть массивом строк."
    For Each key In keys
        If VarType(key) <> vbString Then InputError firstRow, keyColumn, "Подпись поиска должна быть строкой."
        If Len(Trim$(CStr(key))) = 0 Then InputError firstRow, keyColumn, "Подпись поиска должна быть непустой строкой."
    Next key
    caption = Join(keys, " / ")
    For row = firstRow To lastRow
        ' Поиск подписи не читает значение каждого параметра в этом столбце.
        ' Ошибка Excel в неактивном поле не является искомой подписью;
        ' настоящий потребитель поля по-прежнему получает её через CellValue.
        value = mValues(row, keyColumn)
        If Not IsError(value) Then
            If Not BlankValue(value) Then
                For Each key In keys
                    If StrComp(Trim$(CStr(value)), CStr(key), vbTextCompare) = 0 Then
                        If RowByAnyKey > 0 Then InputError row, keyColumn, "Повторный ключ '" & caption & "'."
                        RowByAnyKey = row: Exit For
                    End If
                Next key
            End If
        End If
    Next row
    If RowByAnyKey = 0 Then InputError firstRow, keyColumn, "Не найден параметр '" & caption & "'. Восстановите строку параметра."
End Function

' Определяет пустое значение табличного снимка; ошибочный или произвольный текст не превращается
' в число.
Private Function BlankValue(ByVal value As Variant) As Boolean
    If IsEmpty(value) Or IsNull(value) Then
        BlankValue = True
    ElseIf VarType(value) = vbString Then
        BlankValue = (Len(Trim$(CStr(value))) = 0)
    End If
End Function

' Проверяет координаты внутри табличного снимка до чтения значения и адреса.
Private Sub CheckCell(ByVal row As Long, ByVal column As Long)
    If row < 1 Or row > mRows Or column < 1 Or column > mColumns Then _
        InputError 0, 0, "Адрес поля вне границ снимка: строка " & CStr(row) & ", столбец " & CStr(column) & "."
End Sub

' Формирует ошибку входной таблицы с фактическим адресом ячейки и причиной.
Private Sub InputError(ByVal row As Long, ByVal column As Long, ByVal reason As String)
    Dim location As String
    location = mName
    If row >= 1 And row <= mRows And column >= 1 And column <= mColumns Then
        location = mAddresses(row, column) & " (" & mName & ")"
        ' Человекочитаемая подпись строки берется из самого снимка, без знания формы.
        If column > 1 And VarType(mValues(row, 1)) = vbString Then
            If Len(Trim$(CStr(mValues(row, 1)))) > 0 Then location = location & "; " & CStr(mValues(row, 1))
        End If
    End If
    Err.Raise vbObjectError + 5525, "CSectionInputTable", location & ": " & reason
End Sub
