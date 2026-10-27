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

' ==========================================================================
' Сохранённые значения входной таблицы рецепта
' ==========================================================================
' Хранит неизменяемую копию значений и адресов ячеек, без ссылок на Excel.
' Проверяет типы значений и предоставляет общий способ чтения параметров;
' назначение строк и столбцов определяет конкретный рецепт сечения.

Private mName As String ' Имя входной таблицы рецепта.
Private mValues As Variant, mAddresses As Variant ' Копии значений ячеек и их адресов для пояснений об ошибках.
Private mRows As Long, mColumns As Long ' Число строк и столбцов входной таблицы.

' Копирует двумерные значения и соответствующие адреса ячеек один раз.
' Объекты Excel в таблице не сохраняются.
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

' Возвращает имя сохранённой входной таблицы.
Public Property Get Name() As String: Name = mName: End Property
' Возвращает число строк сохранённой входной таблицы.
Public Property Get RowCount() As Long: RowCount = mRows: End Property
' Возвращает число столбцов сохранённой входной таблицы.
Public Property Get ColumnCount() As Long: ColumnCount = mColumns: End Property

' Возвращает адрес исходной ячейки; для программно заданной таблицы — запись tableName[row,col].
Public Function SourceReference(ByVal row As Long, ByVal column As Long) As String
    CheckCell row, column: SourceReference = mAddresses(row, column)
End Function

' Возвращает сохранённое значение ячейки; ошибка Excel вызывает исключение с её адресом.
Public Function CellValue(ByVal row As Long, ByVal column As Long) As Variant
    CheckCell row, column
    If IsError(mValues(row, column)) Then InputError row, column, "Ячейка содержит ошибку Excel. Исправьте значение или формулу."
    CellValue = mValues(row, column)
End Function

' Проверяет служебную подпись при поиске структуры, не читая данные остальных блоков.
' Ошибка формулы не считается подписью; обращение к самому полю через CellValue её отклонит.
Friend Function HasTextAt(ByVal row As Long, ByVal column As Long, ByVal text As String) As Boolean
    CheckCell row, column
    If VarType(mValues(row, column)) = vbString Then HasTextAt = (StrComp(Trim$(CStr(mValues(row, column))), text, vbTextCompare) = 0)
End Function

' Проверяет, пусто ли сохранённое значение ячейки. Ошибка Excel не считается пустой ячейкой.
Public Function IsBlank(ByVal row As Long, ByVal column As Long) As Boolean
    IsBlank = BlankValue(CellValue(row, column))
End Function

' Читает текст с удалением пробелов по краям; для пустого значения использует
' явно переданную замену, иначе сообщает об ошибке.
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
    Dim parsed As Double
    If Not TryReadStrictDouble(value, parsed) Then InputError row, column, "Ожидается число без разделителей разрядов. Введите числовое значение."
    On Error GoTo Invalid
    Number = parsed
    Exit Function
Invalid:
    InputError row, column, "Число выходит за допустимый диапазон. Введите число меньшего модуля."
End Function

' Читает неотрицательное целое число в диапазоне Long. Общее число активных
' позиций проверяет CRebarRuleExecutor до выделения массивов и расстановки.
Public Function Count(ByVal row As Long, ByVal column As Long, Optional ByVal defaultValue As Long = 0) As Long
    Dim value As Double
    value = Number(row, column, defaultValue)
    If value < 0# Or value > 2147483647# Or Fix(value) <> value Then _
        InputError row, column, "Число стержней должно быть целым неотрицательным и представимым счетчиком Long. Введите целое число от 0 до 2147483647."
    Count = CLng(value)
End Function

' Проверяет значение по списку allowed без учёта регистра
' и возвращает написание, указанное в этом списке.
Public Function Choice(ByVal row As Long, ByVal column As Long, ByVal allowed As Variant, _
        Optional ByVal defaultValue As String = vbNullString) As String
    Dim value As String, item As Variant
    value = Text(row, column, defaultValue)
    For Each item In allowed
        If StrComp(value, CStr(item), vbTextCompare) = 0 Then Choice = CStr(item): Exit Function
    Next item
    InputError row, column, "Недопустимый вариант '" & value & "'. Выберите: " & Join(allowed, ", ") & "."
End Function

' Читает обязательную длину и переводит её из входных единиц в мм через CUnitSystem.
Public Function Length(ByVal row As Long, ByVal column As Long, ByVal units As CUnitSystem) As Double
    On Error GoTo InvalidConversion
    If units Is Nothing Then InputError row, column, "Не передана система единиц."
    Length = units.InputLengthToInternal(Number(row, column))
    Exit Function
InvalidConversion:
    If Err.Number = 6 Then InputError row, column, "Размер после перевода в миллиметры выходит за числовой диапазон. Уменьшите значение и проверьте выбранную единицу длины."
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Читает длину через общий перевод Length в мм, затем проверяет, что она больше нуля.
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
    OptionalDiameter = Length(row, column, units)
End Function

' Читает необязательный отступ ряда: пустая ячейка равна нулю; остальные размеры не ослабляет.
Public Function OptionalRowOffset(ByVal row As Long, ByVal column As Long, ByVal units As CUnitSystem) As Double
    If IsBlank(row, column) Then Exit Function
    On Error GoTo Invalid
    OptionalRowOffset = NonNegativeLength(row, column, units)
    Exit Function
Invalid:
    If Err.Number = 6 Then InputError row, column, "Отступ выходит за числовой диапазон. Введите меньшее неотрицательное расстояние."
    Err.Raise Err.Number, Err.Source, Err.Description
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

' Находит единственную строку с заданным названием параметра в указанной части таблицы.
Public Function RowByKey(ByVal key As String, Optional ByVal keyColumn As Long = 1, _
        Optional ByVal firstRow As Long = 1, Optional ByVal lastRow As Long = 0) As Long
    RowByKey = RowByAnyKey(Array(key), keyColumn, firstRow, lastRow)
End Function

' Ищет подпись только в указанной части таблицы по вариантам, заданным рецептом.
' Одноимённая строка в другом разделе на результат этого поиска не влияет.
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
        value = mValues(row, keyColumn) ' Поиск подписи не читает значение каждого параметра в этом столбце. Ошибка Excel в неактивном поле не является искомой подписью; потребитель значения получает её через CellValue.
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

' Проверяет, пусто ли сохранённое значение. Ошибочный или произвольный текст
' не подменяется нулевым числом.
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
    Dim location As String, value As Variant, valueText As String
    location = mName
    If row >= 1 And row <= mRows And column >= 1 And column <= mColumns Then
        location = mAddresses(row, column) & " (" & mName & ")"
        If column > 1 And VarType(mValues(row, 1)) = vbString Then ' Получает понятное название строки из сохранённой таблицы без привязки к конкретному типу сечения.
            If Len(Trim$(CStr(mValues(row, 1)))) > 0 Then location = location & "; " & CStr(mValues(row, 1))
        End If
        value = mValues(row, column)
        If IsError(value) Then
            valueText = "ошибка формулы Excel"
        ElseIf IsNull(value) Then
            valueText = "неопределённое значение"
        ElseIf IsEmpty(value) Then
            valueText = "пустая ячейка"
        ElseIf Not IsObject(value) And Not IsArray(value) Then
            valueText = """" & CStr(value) & """"
        End If
        If Len(valueText) > 0 Then reason = "Прочитано: " & valueText & "." & vbCrLf & reason
    End If
    Err.Raise vbObjectError + 5525, "CSectionInputTable", SectionRecipeErrorText(reason, location)
End Sub
