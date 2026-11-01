VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CIndexedSnapshotFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Индексированный NDMSNP06: последовательная запись и независимое чтение блоков
' ==========================================================================
' Читатель не является полным CCalculationSnapshot. Полный набор создаёт только LoadAll.
' SHA256 проверяет каталог и каждый прочитанный блок; непрочитанные блоки не проверяются.
Private Const SIGNATURE As String = "NDMSNP06"
Private Const FOOTER As String = "NDMIDX06"
Private Const HEADER_BYTES As Long = 96
Private Const FOOTER_BYTES As Long = 88
Private mFile As CBinarySnapshotFile, mCodec As CCalculationSnapshotStore
Private mTables As Object, mCache As Object, mEntries As Collection, mStateEntries As Collection
Private mIndex As CResultBlockIndex
Private mRunID As Double, mElapsed As Double, mDirectoryHash As String
Private mIdentityEntry As Variant, mHeaderEntry As Variant, mIdentities As Variant
Private mLoadedStateBlocks As Long

' Создаёт небольшой каталог и независимый кэш только общих таблиц текущей операции.
Private Sub Class_Initialize()
    ResetSource
End Sub

' Новый путь всегда получает пустые каталоги и счётчики; кэш другого файла не переиспользуется.
Private Sub ResetSource()
    If Not mFile Is Nothing Then mFile.CloseFile
    Set mFile = Nothing: Set mCodec = Nothing: Set mIndex = Nothing
    Set mTables = CreateObject("Scripting.Dictionary"): mTables.CompareMode = vbBinaryCompare
    Set mCache = CreateObject("Scripting.Dictionary"): mCache.CompareMode = vbBinaryCompare
    Set mEntries = New Collection: Set mStateEntries = New Collection
    mIdentityEntry = Empty: mHeaderEntry = Empty: mIdentities = Empty
    mLoadedStateBlocks = 0: mDirectoryHash = vbNullString: mRunID = 0#: mElapsed = 0#
End Sub

' Записывает блоки подряд, запоминая смещения при записи; повторного прохода по файлу нет.
Public Sub WriteFile(ByVal path As String, ByVal snapshot As CCalculationSnapshot, ByVal codec As CCalculationSnapshotStore)
    ResetSource
    Set mCodec = codec: Set mFile = New CBinarySnapshotFile
    mRunID = snapshot.RunID: mElapsed = snapshot.ElapsedSeconds
    Dim key As Variant, count As Long, digest As String, values As Variant, first As Long
    For Each key In snapshot.Keys: count = count + 1: Next key
    On Error GoTo Failed
    mFile.OpenWrite path: mFile.BeginWriteDigest
    mFile.WriteASCII SIGNATURE: mFile.WriteLong 6: mFile.WriteDouble mRunID: mFile.WriteDouble mElapsed: mFile.WriteLong count
    digest = mFile.EndWriteDigest: mFile.WriteASCII digest
    For Each key In snapshot.Keys
        If CStr(key) = "rngNDMElementResults" Then
            snapshot.WriteIndexedElementBlocks Me
        Else
            values = snapshot.Table(CStr(key))
            WriteTable CStr(key), "Table", values
        End If
    Next key
    values = DirectoryMatrix(): first = mFile.Position: mFile.BeginWriteDigest
    mCodec.WriteIndexedValues mFile, values
    digest = mFile.EndWriteDigest
    Dim bytes As Long: bytes = mFile.Position - first
    mFile.WriteLong first: mFile.WriteLong bytes: mFile.WriteDouble mRunID: mFile.WriteASCII digest: mFile.WriteASCII FOOTER
    mFile.CloseFile
    Exit Sub
Failed:
    Dim reason As String: reason = Err.Description
    On Error Resume Next
    mFile.CloseFile
    On Error GoTo 0
    Err.Raise 5, "CIndexedSnapshotFile.WriteFile", reason
End Sub

' Пишет общую матрицу один раз и сохраняет её физический адрес и сумму.
Private Sub WriteTable(ByVal key As String, ByVal kind As String, ByRef values As Variant)
    Dim first As Long, digest As String
    first = mFile.Position: mFile.BeginWriteDigest
    mCodec.WriteIndexedValues mFile, values
    digest = mFile.EndWriteDigest
    mEntries.Add Array(key, kind, first, mFile.Position - first, digest, UBound(values, 1), UBound(values, 2), 0, Empty)
End Sub

' Собирает общий словарь ElementID/MaterialType один раз, а не для каждого состояния.
' Числовые результаты остаются исходными Double; ключи состояния записываются один раз на участок.
Friend Sub WriteElements(ByRef values As Variant, ByVal index As CResultBlockIndex)
    Dim header As Variant: header = index.Header: WriteTable "rngNDMElementResults", "ElementHeader", header
    Dim dictionary As Object: Set dictionary = CreateObject("Scripting.Dictionary"): dictionary.CompareMode = vbBinaryCompare
    Dim pairs As New Collection, baseCount As Long, row As Long, id As String, material As String, identity As String
    If index.Count > 0 Then baseCount = index.RowCount(1)
    For row = 2 To baseCount + 1
        If VarType(values(row, 5)) <> vbString Or VarType(values(row, 6)) <> vbString Then Err.Raise 5, , "ElementID и MaterialType должны быть строками."
        id = CStr(values(row, 5)): material = CStr(values(row, 6)): identity = CStr(Len(id)) & ":" & id & material
        pairs.Add Array(id, material)
        If Not dictionary.Exists(identity) Then dictionary.Add identity, pairs.Count + 1
    Next row
    Dim block As Long, first As Long, last As Long, column As Long, offset As Long, sample As Variant, localCodes() As Long, digest As String, same As Boolean, relative As Long
    For block = 1 To index.Count
        first = index.FirstRow(block): last = first + index.RowCount(block) - 1: sample = index.Representative(block)
        offset = mFile.Position: mFile.BeginWriteDigest: mFile.WriteLong index.RowCount(block)
        For column = 1 To 4: mCodec.WriteIndexedScalar mFile, sample(1, column): Next column
        same = (index.RowCount(block) = baseCount)
        If same Then
            For row = first To last
                relative = row - first + 2
                If VarType(values(row, 5)) <> vbString Or VarType(values(row, 6)) <> vbString Then same = False: Exit For
                If StrComp(values(row, 5), values(relative, 5), vbBinaryCompare) <> 0 Or StrComp(values(row, 6), values(relative, 6), vbBinaryCompare) <> 0 Then same = False: Exit For
            Next row
        End If
        If same Then
            mFile.WriteByte 0 ' Общая последовательность ID/Material хранится один раз; значения каждой строки проверены.
        Else
            mFile.WriteByte 1: ReDim localCodes(1 To index.RowCount(block))
            For row = first To last
                If VarType(values(row, 5)) <> vbString Or VarType(values(row, 6)) <> vbString Then Err.Raise 5, , "ElementID и MaterialType должны быть строками."
                id = CStr(values(row, 5)): material = CStr(values(row, 6)): identity = CStr(Len(id)) & ":" & id & material
                If Not dictionary.Exists(identity) Then pairs.Add Array(id, material): dictionary.Add identity, pairs.Count + 1
                localCodes(row - first + 1) = CLng(dictionary.Item(identity))
            Next row
            mFile.WriteLongs localCodes
        End If
        WriteNumericColumn values, first, last, 7
        WriteNumericColumn values, first, last, 8
        mCodec.WriteIndexedValues mFile, values, first, last, 9, 9
        digest = mFile.EndWriteDigest
        mEntries.Add Array("rngNDMElementResults", "State", offset, mFile.Position - offset, digest, index.RowCount(block), 9, first, sample)
    Next block
    Dim identityValues() As Variant, pair As Variant, sharedValues As Variant
    ReDim identityValues(1 To pairs.Count + 1, 1 To 2): identityValues(1, 1) = "ElementID": identityValues(1, 2) = "MaterialType"
    For row = 1 To pairs.Count
        pair = pairs.Item(row): identityValues(row + 1, 1) = pair(0): identityValues(row + 1, 2) = pair(1)
    Next row
    sharedValues = identityValues: WriteTable "#ElementIdentities", "ElementIDs", sharedValues
End Sub

' Числовая колонка Double пишется плотным массивом; редкие смешанные типы используют общий кодек.
Private Sub WriteNumericColumn(ByRef values As Variant, ByVal first As Long, ByVal last As Long, ByVal column As Long)
    Dim numbers() As Double, row As Long: ReDim numbers(1 To last - first + 1)
    For row = first To last
        If VarType(values(row, column)) <> vbDouble Then
            mFile.WriteByte 1: mCodec.WriteIndexedValues mFile, values, first, last, column, column
            Exit Sub
        End If
        numbers(row - first + 1) = values(row, column)
    Next row
    mFile.WriteByte 0: mFile.WriteDoubles numbers
End Sub

' Сериализует только небольшой каталог с представительскими строками для indexOnly.
Private Function DirectoryMatrix() As Variant
    Dim values() As Variant, row As Long, column As Long, entry As Variant, sample As Variant
    ReDim values(1 To mEntries.Count, 1 To 17)
    For row = 1 To mEntries.Count
        entry = mEntries.Item(row)
        For column = 1 To 8: values(row, column) = entry(column - 1): Next column
        values(row, 9) = mRunID
        If CStr(entry(1)) = "State" Then
            sample = entry(8)
            For column = 1 To 9: values(row, column + 8) = sample(1, column): Next column
        End If
    Next row
    DirectoryMatrix = values
End Function

' Открывает только заголовок и каталог. False означает конкретную совместимость с v4/v5.
' Неизвестный или повреждённый NDMSNP06 никогда не подменяется другим источником.
Public Function OpenRead(ByVal path As String, ByVal codec As CCalculationSnapshotStore) As Boolean
    ResetSource
    Set mCodec = codec: Set mFile = New CBinarySnapshotFile
    On Error GoTo Failed
    mFile.OpenRead path
    If mFile.Size < 104 Then Err.Raise 5, , "Файл снимка оборван."
    mFile.BeginReadDigest 0, 12
    Dim signature As String, version As Long, digest As String, savedHash As String, tableCount As Long
    signature = mFile.ReadASCII(8): version = mFile.ReadLong: digest = mFile.EndReadDigest
    If (signature = "NDMSNP04" And version = 4) Or (signature = "NDMSNP05" And version = 5) Then mFile.CloseFile: Exit Function
    If signature <> SIGNATURE Or version <> 6 Or mFile.Size < HEADER_BYTES + FOOTER_BYTES Then Err.Raise 5, , "Неизвестная версия формата снимка."
    mFile.BeginReadDigest 0, 32
    signature = mFile.ReadASCII(8): version = mFile.ReadLong
    mRunID = mFile.ReadDouble: mElapsed = mFile.ReadDouble: tableCount = mFile.ReadLong: digest = mFile.EndReadDigest
    mFile.BeginReadDigest 32, 64: savedHash = mFile.ReadASCII(64): signature = mFile.EndReadDigest
    If savedHash <> digest Or tableCount < 1 Or tableCount > 64 Then Err.Raise 5, , "Повреждён заголовок снимка."
    mFile.BeginReadDigest mFile.Size - FOOTER_BYTES, FOOTER_BYTES
    Dim offset As Long, bytes As Long, footerRun As Double
    offset = mFile.ReadLong: bytes = mFile.ReadLong: footerRun = mFile.ReadDouble
    mDirectoryHash = mFile.ReadASCII(64): signature = mFile.ReadASCII(8): digest = mFile.EndReadDigest
    If signature <> FOOTER Or footerRun <> mRunID Or offset < HEADER_BYTES Or bytes < 8 Then Err.Raise 5, , "Повреждён адрес каталога снимка."
    If CDbl(offset) + bytes <> mFile.Size - FOOTER_BYTES Then Err.Raise 5, , "Каталог выходит за границы снимка."
    mFile.BeginReadDigest offset, bytes
    Dim directory As Variant: directory = mCodec.ReadIndexedValues(mFile, , 17)
    digest = mFile.EndReadDigest
    If digest <> mDirectoryHash Then Err.Raise 5, , "Контрольная сумма каталога не совпадает."
    ReadDirectory directory, offset, tableCount
    OpenRead = True
    Exit Function
Failed:
    Dim reason As String: reason = Err.Description
    On Error Resume Next
    mFile.CloseFile
    On Error GoTo 0
    Err.Raise 5, "CIndexedSnapshotFile.OpenRead", "Не удалось прочитать каталог " & path & ": " & reason
End Function

' Проверяет непрерывность адресов, число таблиц, относительные строки и принадлежность RunID.
Private Sub ReadDirectory(ByRef values As Variant, ByVal directoryOffset As Long, ByVal tableCount As Long)
    Dim row As Long, column As Long, offset As Long, nextOffset As Long, length As Long, rows As Long, columns As Long
    Dim entry As Variant, key As String, kind As String, sample As Variant, cells(1 To 1, 1 To 9) As Variant
    Dim stateRecords As New Collection
    nextOffset = HEADER_BYTES
    For row = 1 To UBound(values, 1)
        If CDbl(values(row, 9)) <> mRunID Then Err.Raise 5, , "Каталог относится к другому RunID."
        key = CStr(values(row, 1)): kind = CStr(values(row, 2)): offset = ExactLong(values(row, 3)): length = ExactLong(values(row, 4))
        rows = ExactLong(values(row, 6)): columns = ExactLong(values(row, 7))
        If offset <> nextOffset Or length < 1 Or CDbl(offset) + length > directoryOffset Then Err.Raise 5, , "Наложение, пропуск или выход блока за границы каталога."
        If Len(CStr(values(row, 5))) <> 64 Or rows < 1 Or rows > 10000000 Or columns < 1 Or columns > 256 Or CDbl(rows) * columns > 100000000# Then Err.Raise 5, , "Повреждены размеры или сумма блока."
        For column = 1 To 9: cells(1, column) = values(row, column + 8): Next column
        sample = cells
        entry = Array(key, kind, offset, length, CStr(values(row, 5)), rows, columns, ExactLong(values(row, 8)), sample)
        Select Case kind
            Case "Table"
                If mTables.Exists(key) Or key = "rngNDMElementResults" Then Err.Raise 5, , "Повторная таблица каталога."
                mTables.Add key, entry
            Case "ElementHeader"
                If IsArray(mHeaderEntry) Or key <> "rngNDMElementResults" Or rows <> 1 Or columns <> 9 Then Err.Raise 5, , "Повреждена запись шапки результатов."
                mHeaderEntry = entry: mTables.Add key, entry
            Case "ElementIDs"
                If IsArray(mIdentityEntry) Or key <> "#ElementIdentities" Or columns <> 2 Then Err.Raise 5, , "Повреждён общий список элементов."
                mIdentityEntry = entry
            Case "State"
                If key <> "rngNDMElementResults" Or columns <> 9 Or CDbl(sample(1, 1)) <> mRunID Then Err.Raise 5, , "Неверная принадлежность состояния."
                mStateEntries.Add entry
            Case Else: Err.Raise 5, , "Неизвестный вид блока каталога."
        End Select
        mEntries.Add entry
        nextOffset = offset + length
    Next row
    If nextOffset <> directoryOffset Or mTables.Count <> tableCount Or Not IsArray(mHeaderEntry) Or Not IsArray(mIdentityEntry) Then Err.Raise 5, , "Неполный каталог снимка."
    Dim header As Variant: header = ReadEntry(mHeaderEntry)
    Set mIndex = New CResultBlockIndex: mIndex.Initialize header, mRunID
    For row = 1 To mStateEntries.Count
        entry = mStateEntries.Item(row): sample = entry(8)
        mIndex.Append CLng(entry(7)), CLng(entry(5)), sample
    Next row
End Sub

' Принимает только целые Long: дробь и округление повреждённого адреса запрещены.
Private Function ExactLong(ByVal value As Variant) As Long
    If Not IsNumeric(value) Then Err.Raise 5, , "В каталоге требуется целое число."
    If CDbl(value) < 0# Or CDbl(value) > 2147483647# Then Err.Raise 5, , "Адрес каталога вне Long."
    ExactLong = CLng(value)
    If CDbl(ExactLong) <> CDbl(value) Then Err.Raise 5, , "Дробный адрес в каталоге."
End Function

' Читает одну общую матрицу в ограниченном окне с проверкой суммы и RunID.
Private Function ReadEntry(ByVal entry As Variant) As Variant
    mFile.BeginReadDigest CLng(entry(2)), CLng(entry(3))
    Dim values As Variant: values = mCodec.ReadIndexedValues(mFile, CLng(entry(5)), CLng(entry(6)))
    Dim digest As String: digest = mFile.EndReadDigest
    If digest <> CStr(entry(4)) Then Err.Raise 5, , "Контрольная сумма блока не совпадает: " & CStr(entry(0))
    Dim key As String, row As Long: key = CStr(entry(0))
    If Left$(key, 6) = "rngNDM" And key <> "rngNDMElementResults" Or Left$(key, 9) = "LoadCase." Then
        For row = 2 To UBound(values, 1)
            If CDbl(values(row, 1)) <> mRunID Then Err.Raise 5, , "Общая таблица относится к другому расчёту: " & key
        Next row
    End If
    ReadEntry = values
End Function

' Возвращает одну таблицу; полные подробные результаты запрашиваются только явно.
Public Function Table(ByVal key As String) As Variant
    If key = "rngNDMElementResults" Then Table = SelectElementResults("", "", False): Exit Function
    If Not mTables.Exists(key) Then Err.Raise 5, , "В снимке отсутствует таблица: " & key
    If Not mCache.Exists(key) Then mCache.Add key, ReadEntry(mTables.Item(key))
    Dim values As Variant: values = mCache.Item(key)
    If key = "Summary" And mElapsed > 0# Then values(4, 1) = mElapsed
    Table = values
End Function

' Отбирает все подходящие физические блоки; indexOnly не читает ни одного блока состояний.
Public Function SelectElementResults(ByVal combinationID As String, ByVal stateType As String, ByVal indexOnly As Boolean, Optional ByVal profileID As String = "") As Variant
    If indexOnly Then SelectElementResults = mIndex.StateRows(combinationID, stateType, profileID): Exit Function
    Dim selected As Collection: Set selected = mIndex.SelectBlocks(combinationID, stateType, profileID)
    Dim block As Variant, rows As Long, result() As Variant, row As Long, column As Long, used As Long, values As Variant, header As Variant
    For Each block In selected: rows = rows + mIndex.RowCount(CLng(block)): Next block
    ReDim result(1 To rows + 1, 1 To 9): header = mIndex.Header
    For column = 1 To 9: result(1, column) = header(1, column): Next column
    used = 1
    For Each block In selected
        values = ReadState(CLng(block))
        For row = 1 To UBound(values, 1)
            used = used + 1
            For column = 1 To 9: result(used, column) = values(row, column): Next column
        Next row
    Next block
    SelectElementResults = result
End Function

' Восстанавливает один компактный блок с общим словарём неизменяемых ID/MaterialType.
Private Function ReadState(ByVal block As Long) As Variant
    If Not IsArray(mIdentities) Then
        mIdentities = ReadEntry(mIdentityEntry)
        If CStr(mIdentities(1, 1)) <> "ElementID" Or CStr(mIdentities(1, 2)) <> "MaterialType" Then Err.Raise 5, , "Повреждена шапка списка элементов."
    End If
    Dim entry As Variant: entry = mStateEntries.Item(block)
    mFile.BeginReadDigest CLng(entry(2)), CLng(entry(3))
    Dim rows As Long: rows = mFile.ReadLong
    If rows <> mIndex.RowCount(block) Then Err.Raise 5, , "Число строк состояния не совпадает с каталогом."
    Dim keys(1 To 4) As Variant, column As Long, codes() As Long, data As Variant, strain As Variant, stress As Variant, sample As Variant, digest As String, idMode As Byte
    sample = mIndex.Representative(block)
    For column = 1 To 4
        keys(column) = mCodec.ReadIndexedScalar(mFile)
        If VarType(keys(column)) <> VarType(sample(1, column)) Or keys(column) <> sample(1, column) Then Err.Raise 5, , "Ключ состояния не совпадает с каталогом."
    Next column
    idMode = mFile.ReadByte
    If idMode = 1 Then
        codes = mFile.ReadLongs(rows)
    ElseIf idMode = 0 Then
        If rows + 1 > UBound(mIdentities, 1) Then Err.Raise 5, , "Общая последовательность ID слишком короткая."
        ReDim codes(1 To rows)
        Dim element As Long: For element = 1 To rows: codes(element) = element + 1: Next element
    Else
        Err.Raise 5, , "Неизвестный способ адресации ID."
    End If
    strain = ReadNumericColumn(rows): stress = ReadNumericColumn(rows)
    data = mCodec.ReadIndexedValues(mFile, rows, 1): digest = mFile.EndReadDigest
    If digest <> CStr(entry(4)) Then Err.Raise 5, , "Контрольная сумма состояния не совпадает."
    Dim values() As Variant, row As Long: ReDim values(1 To rows, 1 To 9)
    For row = 1 To rows
        If codes(row) < 2 Or codes(row) > UBound(mIdentities, 1) Then Err.Raise 5, , "Неверный индекс ElementID состояния."
        For column = 1 To 4: values(row, column) = keys(column): Next column
        values(row, 5) = mIdentities(codes(row), 1): values(row, 6) = mIdentities(codes(row), 2)
        values(row, 7) = strain(row, 1): values(row, 8) = stress(row, 1): values(row, 9) = data(row, 1)
    Next row
    For column = 1 To 9
        If VarType(values(1, column)) <> VarType(sample(1, column)) Or values(1, column) <> sample(1, column) Then Err.Raise 5, , "Представительская строка каталога не совпадает с состоянием."
    Next column
    mLoadedStateBlocks = mLoadedStateBlocks + 1: ReadState = values
End Function

' Полное восстановление читает и проверяет все блоки, затем проверяет полный контракт снимка.
Public Function LoadAll() As CCalculationSnapshot
    Dim snapshot As New CCalculationSnapshot, key As Variant, values As Variant
    snapshot.RunID = mRunID: snapshot.ElapsedSeconds = mElapsed
    Dim elements() As Variant
    For Each key In mTables.Keys
        If CStr(key) = "rngNDMElementResults" Then
            elements = SelectElementResults("", "", False)
            snapshot.AddPreparedElementResults elements, mIndex.TotalRows, mIndex
        Else
            values = Table(CStr(key)): snapshot.AddTable CStr(key), values
        End If
    Next key
    snapshot.Validate
    Set LoadAll = snapshot
End Function

' Проверяет наличие таблицы по проверенному каталогу без чтения её содержимого.
Public Function HasTable(ByVal key As String) As Boolean
    HasTable = mTables.Exists(key)
End Function

' Перечисляет только исходные таблицы, без внутреннего списка ID и блоков состояний.
Public Function Keys() As Variant
    Keys = mTables.Keys
End Function

' Возвращает высоту из каталога без загрузки большой матрицы.
Public Function TableRowCount(ByVal key As String) As Long
    If key = "rngNDMElementResults" Then TableRowCount = mIndex.TotalRows: Exit Function
    Dim entry As Variant: entry = mTables.Item(key): TableRowCount = CLng(entry(5))
End Function

' Исходная шапка результатов хранится отдельным общим блоком.
Public Function ElementHeader() As Variant
    ElementHeader = mIndex.Header
End Function

' RunID относится к проверенному заголовку и каталогу текущего открытого дескриптора.
Public Property Get RunID() As Double
    RunID = mRunID
End Property

' Время исходного запуска не меняет численные результаты блоков.
Public Property Get ElapsedSeconds() As Double
    ElapsedSeconds = mElapsed
End Property

' Привязка v6 использует проверенный каталог, а не повторный полный SHA256 файла.
Public Property Get Identity() As String
    Identity = "IDX6:" & mDirectoryHash
End Property

' Физические байты включают заголовок, каталог и все реально проверенные блоки.
Public Property Get BytesRead() As Double
    BytesRead = mFile.BytesRead
End Property

' Отдельный счётчик подтверждает отсутствие чтения остальных состояний.
Public Property Get LoadedStateBlocks() As Long
    LoadedStateBlocks = mLoadedStateBlocks
End Property

' Обновляет фиксированный заголовок и его SHA256, не перечитывая состояния.
' Вызывается только для собственной временной копии; атомарная публикация принадлежит Store.
Public Sub UpdateElapsed(ByVal path As String, ByVal seconds As Double)
    mFile.CloseFile
    mFile.OpenUpdate path: mFile.Position = 0: mFile.BeginWriteDigest
    mFile.WriteASCII SIGNATURE: mFile.WriteLong 6: mFile.WriteDouble mRunID: mFile.WriteDouble seconds: mFile.WriteLong mTables.Count
    Dim digest As String: digest = mFile.EndWriteDigest: mFile.WriteASCII digest: mFile.CloseFile
End Sub

' Восстанавливает плотный Double без округления; общий кодек сохраняет смешанные значения.
Private Function ReadNumericColumn(ByVal rows As Long) As Variant
    Dim mode As Byte: mode = mFile.ReadByte
    If mode = 1 Then ReadNumericColumn = mCodec.ReadIndexedValues(mFile, rows, 1): Exit Function
    If mode <> 0 Then Err.Raise 5, , "Неизвестный способ хранения числовой колонки."
    Dim numbers() As Double, values() As Variant, row As Long
    numbers = mFile.ReadDoubles(rows): ReDim values(1 To rows, 1 To 1)
    For row = 1 To rows: values(row, 1) = numbers(row): Next row
    ReadNumericColumn = values
End Function

' Проверяет SHA256 всех блоков по каталогу без восстановления больших матриц.
' Полный LoadAll дополнительно проверяет семантический контракт восстановленных таблиц.
Public Sub VerifyAllBlocks()
    Dim entry As Variant, bytes As Variant, take As Long, digest As String
    For Each entry In mEntries
        mFile.BeginReadDigest CLng(entry(2)), CLng(entry(3))
        Do While mFile.RemainingBytes > 0
            take = mFile.RemainingBytes: If take > 65536 Then take = 65536
            bytes = mFile.ReadBytes(take)
        Loop
        digest = mFile.EndReadDigest
        If digest <> CStr(entry(4)) Then Err.Raise 5, , "Контрольная сумма блока не совпадает: " & CStr(entry(0))
    Next entry
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает адреса первого и последнего состояния для контролируемого повреждения частной копии.
Friend Function TestStateBlockLocation(ByVal block As Long) As String
    If block = 0 Then block = mStateEntries.Count
    Dim entry As Variant, sample As Variant: entry = mStateEntries.Item(block): sample = entry(8)
    TestStateBlockLocation = CStr(entry(2)) & "|" & CStr(entry(3)) & "|" & CStr(sample(1, 2)) & "|" & CStr(sample(1, 4))
End Function

' Явно освобождает собственный дескриптор перед удалением или заменой временного файла.
Public Sub CloseFile()
    If Not mFile Is Nothing Then mFile.CloseFile
End Sub
