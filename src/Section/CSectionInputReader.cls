VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionInputReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Чтение блоков входной таблицы по коротким ключам
' ==========================================================================
' Разбирает неизменяемый снимок CSectionInputTable один раз. Находит поле по имени
' блока, ключу строки и ключу столбца; подписи для пользователя и порядок блоков,
' строк и столбцов на результат не влияют. Числа проверяет CSectionInputTable,
' а длины переводит в миллиметры CUnitSystem. Ссылок на объекты Excel нет.
'
' Ридер подходит любому рецепту с такой структурой таблицы. Он не выбирает форму
' сечения, не расставляет арматуру и не проверяет инженерные ограничения рецепта.

Private mTable As CSectionInputTable ' Исходные значения и фактические адреса ячеек.
Private mUnits As CUnitSystem ' Система единиц для чтения длин; для поиска ключей не требуется.
Private mIndexed As Boolean ' Показывает, что структура блоков уже проверена и проиндексирована.
Private mRowPrefix As String ' Общая часть ключей строк выбранного набора, например Circle.; пустая строка означает полные ключи.
Private mBlocks As Object ' Индексы строк и столбцов каждого блока; ключи сравниваются без учёта регистра.

' Сохраняет таблицу и независимые единицы ввода один раз.
' rowPrefix задаёт общую начальную часть ключей выбранного набора строк.
' Структура блоков проверяется только при первом обращении по ключам.
Public Sub Initialize(ByVal table As CSectionInputTable, Optional ByVal units As CUnitSystem = Nothing, Optional ByVal rowPrefix As String = vbNullString)
    If Not mTable Is Nothing Then Err.Raise vbObjectError + 5525, "CSectionInputReader", "Объект чтения входной таблицы уже создан; повторная инициализация запрещена."
    If table Is Nothing Then Err.Raise vbObjectError + 5525, "CSectionInputReader", "Не передан снимок входной таблицы."
    Set mTable = table
    mRowPrefix = rowPrefix
    If Not units Is Nothing Then Set mUnits = units.Clone: mUnits.Freeze

End Sub

' Читает исходное значение по номерам строки и столбца внутри всего диапазона, начиная с 1.
' Подходит собственной структуре таблицы без блоков и ключей; единицы не преобразует.
' Выход за диапазон и ошибки формул проверяет снимок с фактическим адресом поля.
Public Function ValueAt(ByVal row As Long, ByVal column As Long) As Variant
    If mTable Is Nothing Then Err.Raise vbObjectError + 5525, "CSectionInputReader", "Перед чтением данных передайте входную таблицу методом Initialize."
    ValueAt = mTable.CellValue(row, column)
End Function

' Возвращает строки выбранного блока в порядке таблицы, без шапок и строки единиц Units.
Public Function RowKeys(ByVal blockKey As String) As Variant
    Dim block As Object: Set block = RequiredBlock(blockKey)
    Dim result As Object, key As Variant: Set result = NewIndex()
    For Each key In block("Rows").Keys
        If StrComp(Left$(CStr(key), Len(mRowPrefix)), mRowPrefix, vbTextCompare) = 0 Then result.Add Mid$(CStr(key), Len(mRowPrefix) + 1), True
    Next key
    RowKeys = result.Keys
End Function

' Находит строку только внутри указанного блока; одинаковые ключи разных блоков допустимы.
Public Function RowIndex(ByVal blockKey As String, ByVal rowKey As String) As Long
    Dim block As Object: Set block = RequiredBlock(blockKey)
    rowKey = mRowPrefix & rowKey
    If Not block("Rows").Exists(rowKey) Then InvalidCell CLng(block("TitleRow")), 1, "В блоке '" & blockKey & "' не найдена строка '" & rowKey & "'. Восстановите ключ строки."
    RowIndex = CLng(block("Rows")(rowKey))
End Function

' Находит столбец по машинному ключу независимо от его русской подписи и положения.
Public Function ColumnIndex(ByVal blockKey As String, ByVal columnKey As String) As Long
    Dim block As Object: Set block = RequiredBlock(blockKey)
    If Not block("Columns").Exists(columnKey) Then InvalidCell CLng(block("HeaderRow")), 1, "В блоке '" & blockKey & "' не найден столбец '" & columnKey & "'. Восстановите ключ столбца."
    ColumnIndex = CLng(block("Columns")(columnKey))
End Function

' Возвращает исходное значение поля, сохраняя проверку ошибок Excel и фактический адрес.
Public Function CellValue(ByVal blockKey As String, ByVal rowKey As String, Optional ByVal columnKey As String = "val") As Variant
    CellValue = mTable.CellValue(RowIndex(blockKey, rowKey), ColumnIndex(blockKey, columnKey))
End Function

' Читает обязательный текст; явно переданное значение используется только для пустого поля.
Public Function Text(ByVal blockKey As String, ByVal rowKey As String, Optional ByVal columnKey As String = "val", Optional ByVal defaultValue As Variant) As String
    If IsMissing(defaultValue) Then
        Text = mTable.Text(RowIndex(blockKey, rowKey), ColumnIndex(blockKey, columnKey))
    Else
        Text = mTable.Text(RowIndex(blockKey, rowKey), ColumnIndex(blockKey, columnKey), defaultValue)
    End If
End Function

' Читает число с проверками типов снимка таблицы; обязательная пустая ячейка не становится нулём.
' Значение defaultValue используется только при явно переданной замене пустого поля.
Public Function Number(ByVal blockKey As String, ByVal rowKey As String, Optional ByVal columnKey As String = "val", Optional ByVal defaultValue As Variant) As Double
    If IsMissing(defaultValue) Then
        Number = mTable.Number(RowIndex(blockKey, rowKey), ColumnIndex(blockKey, columnKey))
    Else
        Number = mTable.Number(RowIndex(blockKey, rowKey), ColumnIndex(blockKey, columnKey), defaultValue)
    End If
End Function

' Читает целое неотрицательное число; пустое значение равно явно заданной замене, обычно нулю.
Public Function Count(ByVal blockKey As String, ByVal rowKey As String, Optional ByVal columnKey As String = "val", Optional ByVal defaultValue As Long = 0) As Long
    Count = mTable.Count(RowIndex(blockKey, rowKey), ColumnIndex(blockKey, columnKey), defaultValue)
End Function

' Проверяет вариант по переданному списку и возвращает его согласованное написание.
Public Function Choice(ByVal blockKey As String, ByVal rowKey As String, ByVal allowed As Variant, Optional ByVal columnKey As String = "val", Optional ByVal defaultValue As String = vbNullString) As String
    Choice = mTable.Choice(RowIndex(blockKey, rowKey), ColumnIndex(blockKey, columnKey), allowed, defaultValue)
End Function

' Читает явный способ расстановки основного ряда без автоматической подмены режима.
' Совместимость с открытым или замкнутым путём проверяет рецепт либо общий исполнитель.
Public Function RebarEndMode(ByVal blockKey As String, ByVal rowKey As String) As ERebarEndMode
    Select Case Choice(blockKey, rowKey, Array("IncludeBoth", "ExcludeBoth", "IncludeStart", "IncludeEnd"), "endMode")
        Case "IncludeBoth": RebarEndMode = IncludeBoth
        Case "ExcludeBoth": RebarEndMode = ExcludeBoth
        Case "IncludeStart": RebarEndMode = IncludeStart
        Case "IncludeEnd": RebarEndMode = IncludeEnd
    End Select
End Function

' Читает длину и переводит её в мм через единую систему единиц.
Public Function Length(ByVal blockKey As String, ByVal rowKey As String, Optional ByVal columnKey As String = "val") As Double
    Length = mTable.Length(RowIndex(blockKey, rowKey), ColumnIndex(blockKey, columnKey), mUnits)
End Function

' Читает положительный размер в мм; нулевое или отрицательное значение вызывает ошибку ввода.
Public Function PositiveLength(ByVal blockKey As String, ByVal rowKey As String, Optional ByVal columnKey As String = "val") As Double
    PositiveLength = mTable.PositiveLength(RowIndex(blockKey, rowKey), ColumnIndex(blockKey, columnKey), mUnits)
End Function

' Читает неотрицательный размер в мм; пустая обязательная ячейка не заменяется нулём.
Public Function NonNegativeLength(ByVal blockKey As String, ByVal rowKey As String, Optional ByVal columnKey As String = "val") As Double
    NonNegativeLength = mTable.NonNegativeLength(RowIndex(blockKey, rowKey), ColumnIndex(blockKey, columnKey), mUnits)
End Function

' Читает диаметр в мм. Пустота и ноль выключают ряд, отрицательный размер и текст недопустимы.
Public Function OptionalDiameter(ByVal blockKey As String, ByVal rowKey As String, Optional ByVal columnKey As String = "val") As Double
    OptionalDiameter = mTable.OptionalDiameter(RowIndex(blockKey, rowKey), ColumnIndex(blockKey, columnKey), mUnits)
End Function

' Читает необязательное неотрицательное смещение в мм; пустая ячейка означает ноль.
Public Function OptionalRowOffset(ByVal blockKey As String, ByVal rowKey As String, Optional ByVal columnKey As String = "val") As Double
    OptionalRowOffset = mTable.OptionalRowOffset(RowIndex(blockKey, rowKey), ColumnIndex(blockKey, columnKey), mUnits)
End Function

' Возвращает фактический адрес найденного поля, в том числе после перестановки строк и столбцов.
Public Function SourceReference(ByVal blockKey As String, ByVal rowKey As String, Optional ByVal columnKey As String = "val") As String
    SourceReference = mTable.SourceReference(RowIndex(blockKey, rowKey), ColumnIndex(blockKey, columnKey))
End Function

' Собирает адреса связанных параметров для сообщения об ошибке общего геометрического слоя.
Public Function SourceReferences(ByVal blockKey As String, ByVal rowKey As String, ByVal columnKeys As Variant) As String
    Dim key As Variant
    For Each key In columnKeys
        If Len(SourceReferences) > 0 Then SourceReferences = SourceReferences & "; "
        SourceReferences = SourceReferences & SourceReference(blockKey, rowKey, CStr(key))
    Next key
End Function

' Привязывает проверку рецепта к найденному полю; само инженерное условие задаёт рецепт.
Public Sub Require(ByVal blockKey As String, ByVal rowKey As String, ByVal columnKey As String, ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then SectionRecipeError "CSectionInputReader", reason & " Исправьте указанное значение.", SourceReference(blockKey, rowKey, columnKey)
End Sub

' Строит индексы при первом запросе по ключам; прямое чтение ValueAt не требует блоков.
' Над строкой key расположены пользовательские заголовки, ещё выше — ключ и название блока.
Private Sub EnsureIndex()
    If mIndexed Then Exit Sub
    Set mBlocks = NewIndex()
    Dim row As Long, column As Long, key As String, block As Object, columns As Object, rows As Object
    Dim names As Variant, index As Long, first As Long, last As Long
    For row = 3 To mTable.RowCount
        If mTable.HasTextAt(row, 1, "key") Then
            key = KeyText(row - 2, 1)
            If Len(key) = 0 Then InvalidCell row - 2, 1, "Не задан ключ блока над строкой ключей столбцов."
            If mBlocks.Exists(key) Then InvalidCell row - 2, 1, "Повторный ключ блока '" & key & "'."
            Set block = NewIndex(): Set columns = NewIndex(): Set rows = NewIndex()
            block.Add "HeaderRow", row: block.Add "TitleRow", row - 2 ' Служебная строка key находится на две строки ниже заголовка блока.
            block.Add "Columns", columns: block.Add "Rows", rows
            mBlocks.Add key, block
            For column = 1 To mTable.ColumnCount
                key = KeyText(row, column)
                If Len(key) > 0 Then
                    If columns.Exists(key) Then InvalidCell row, column, "Повторный ключ столбца '" & key & "' в одном блоке."
                    columns.Add key, column
                End If
            Next column
        End If
    Next row
    If mBlocks.Count = 0 Then InvalidCell 1, 1, "Не найдена строка key с ключами столбцов. Восстановите структуру входной таблицы."
    names = mBlocks.Keys ' Dictionary сохраняет порядок обнаружения блоков; он нужен для границ строк, не для поиска полей.
    For index = 0 To mBlocks.Count - 1
        Set block = mBlocks(CStr(names(index))): Set rows = block("Rows")
        first = CLng(block("HeaderRow")) + 1: last = mTable.RowCount
        If index < mBlocks.Count - 1 Then last = CLng(mBlocks(CStr(names(index + 1)))("TitleRow")) - 1
        block.Add "FirstRow", first: block.Add "LastRow", last: block.Add "RowsIndexed", False
    Next index
    mIndexed = True
End Sub

' Возвращает индекс блока либо объясняет, какой ключ отсутствует в исходной таблице.
Private Function RequiredBlock(ByVal key As String) As Object
    If mTable Is Nothing Then Err.Raise vbObjectError + 5525, "CSectionInputReader", "Перед чтением данных передайте входную таблицу методом Initialize."
    EnsureIndex
    If Not mBlocks.Exists(key) Then InvalidCell 1, 1, "Не найден блок '" & key & "'. Восстановите ключ блока."
    Set RequiredBlock = mBlocks(key)
    IndexBlockRows RequiredBlock, key
End Function

' Индексирует строки только запрошенного блока. Ошибка исходного размера сначала
' указывает на поле ввода, а не на зависимую формулу ключа в ещё не запрошенном блоке.
Private Sub IndexBlockRows(ByVal block As Object, ByVal blockKey As String)
    If CBool(block("RowsIndexed")) Then Exit Sub
    Dim rows As Object, row As Long, key As String: Set rows = NewIndex()
    For row = CLng(block("FirstRow")) To CLng(block("LastRow"))
        key = KeyText(row, 1)
        If Len(key) > 0 And StrComp(key, "Units", vbTextCompare) <> 0 Then
            If rows.Exists(key) Then InvalidCell row, 1, "Повторный ключ строки '" & key & "' в блоке '" & blockKey & "'."
            rows.Add key, row
        End If
    Next row
    Set block("Rows") = rows: block("RowsIndexed") = True
End Sub

' Создаёт независимый индекс без различения регистра, не хранящий объекты Excel.
Private Function NewIndex() As Object
    Set NewIndex = CreateObject("Scripting.Dictionary"): NewIndex.CompareMode = vbTextCompare
End Function

' Читает только структурный ключ; числовые поля не проверяются до обращения потребителя.
Private Function KeyText(ByVal row As Long, ByVal column As Long) As String
    Dim value As Variant: value = mTable.CellValue(row, column)
    If IsEmpty(value) Or IsNull(value) Then Exit Function
    If VarType(value) <> vbString Then InvalidCell row, column, "Ключ таблицы должен быть текстом."
    KeyText = Trim$(CStr(value))
End Function

' Сообщает о нарушении структуры с именем таблицы и адресом проблемной ячейки.
Private Sub InvalidCell(ByVal row As Long, ByVal column As Long, ByVal reason As String)
    Err.Raise vbObjectError + 5525, "CSectionInputReader", mTable.SourceReference(row, column) & " (" & mTable.Name & "): " & reason
End Sub
