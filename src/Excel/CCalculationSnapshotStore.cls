VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CCalculationSnapshotStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Файловый снимок расчёта или явно сформированной геометрии с проверкой целостности
' ==========================================================================
' Сохраняет полный набор в одном .ndm; Excel-вывод не использует этот архив.
' Одинаковые строки колонки хранятся словарём, числа — исходными Double.
' Подмена файла выполняется атомарно после записи и контрольной суммы.
Private Declare PtrSafe Function MoveFileExW Lib "kernel32" (ByVal source As LongPtr, ByVal target As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Sub CopySnapshotBits Lib "kernel32" Alias "RtlMoveMemory" (ByRef target As Currency, ByRef source As Double, ByVal count As LongPtr)
Private Declare PtrSafe Sub CopyPackedBytes Lib "kernel32" Alias "RtlMoveMemory" (ByVal target As LongPtr, ByVal source As LongPtr, ByVal count As LongPtr)
Private Const MAGIC As String = "NDMSNP04" ' Контракт плотных колонок и точных повторяющихся серий версии 4.
Private Const BOUNDARY_MAGIC As String = "NDMSNP05" ' NDMSNP05 явно отклоняется читателем NDMSNP04, который не поддерживает обязательные границы.
Private Const FILE_EXTENSION As String = ".ndm" ' Имя в Config не требует расширения; новые снимки получают один общий суффикс.
Private Const LEGACY_FILE_EXTENSION As String = ".ndms" ' Расширение .ndms поддерживается для чтения без изменения содержимого или имени.
Private Const LINK_PREFIX As String = "NDM.CalculationSnapshot." ' Привязка показанных результатов к уже загруженному файлу.

' Проверяет источник уже показанного расчёта независимо от нового выбора Config.
Public Function HasFileReference(ByVal workbook As Object) As Boolean
    HasFileReference = Len(LinkText(workbook, "File")) > 0
End Function

' Связывает Results с точным файлом после успешного сохранения или загрузки.
Public Sub LinkFile(ByVal workbook As Object)
    Dim path As String: path = SelectedReadPath(workbook)
    Dim fileSystem As Object: Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Dim digest As CFileChecksum: Set digest = New CFileChecksum
    SetLink workbook, "File", fileSystem.GetFileName(path)
    SetLink workbook, "Folder", fileSystem.GetParentFolderName(path)
    SetLink workbook, "Hash", digest.FileSHA256(path)
End Sub

' Читает файл ранее показанного снимка и проверяет, что его не заменили другим расчётом.
' Выбор другого имени в Config применяется явно кнопкой обновления.
Public Function LoadLinkedFile(ByVal workbook As Object) As CCalculationSnapshot
    Dim path As String: path = CStr(workbook.Path) & Application.PathSeparator & LinkText(workbook, "File")
    If Not FileExists(path) Then path = LinkText(workbook, "Folder") & Application.PathSeparator & LinkText(workbook, "File")
    Dim digest As CFileChecksum: Set digest = New CFileChecksum
    If digest.FileSHA256(path) <> LinkText(workbook, "Hash") Then Err.Raise vbObjectError + 4955, , _
        "Файл показанного расчёта изменился: " & path & ". Нажмите «Обновить схему и результаты», чтобы загрузить выбранный снимок."
    Dim snapshot As CCalculationSnapshot: Set snapshot = ReadBinary(path): snapshot.Validate
    Set LoadLinkedFile = snapshot
End Function

' Удаляет только привязку; сохранённый файл предыдущего расчёта не удаляется.
Public Sub ClearReference(ByVal workbook As Object)
    Dim key As Variant
    For Each key In Array("File", "Folder", "Hash")
        On Error Resume Next
        workbook.CustomDocumentProperties.Item(LINK_PREFIX & CStr(key)).Delete
        On Error GoTo 0
    Next key
End Sub

' Записывает одно собственное служебное свойство без изменения чужих метаданных.
Private Sub SetLink(ByVal workbook As Object, ByVal key As String, ByVal value As String)
    On Error Resume Next
    workbook.CustomDocumentProperties.Item(LINK_PREFIX & key).Delete
    On Error GoTo 0
    workbook.CustomDocumentProperties.Add LINK_PREFIX & key, False, 4, value
End Sub

' Возвращает пустоту при отсутствии собственной ссылки в ещё не рассчитанной книге.
Private Function LinkText(ByVal workbook As Object, ByVal key As String) As String
    On Error Resume Next
    LinkText = CStr(workbook.CustomDocumentProperties.Item(LINK_PREFIX & key).Value)
    On Error GoTo 0
End Function

' Читает текущий выбор через общий reader; неподдерживаемое значение не заменяет.
Public Function StorageMode(ByVal workbook As Object) As String
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader
    settings.LoadFromRange workbook.Names.Item("rngSystemSettings").RefersToRange
    StorageMode = settings.GetRequiredString("Results.SnapshotStorage")
    If StorageMode <> "Excel" And StorageMode <> "File" Then Err.Raise 5, "CCalculationSnapshotStore", _
        settings.InputErrorMessage("Results.SnapshotStorage", "Допустимы Excel или File.", "Выберите место сохранения снимка.")
End Function

' Разрешает только имя файла рядом с книгой; путь к произвольному каталогу не принимается.
Public Function SnapshotPath(ByVal workbook As Object) As String
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader
    settings.LoadFromRange workbook.Names.Item("rngSystemSettings").RefersToRange
    Dim name As String: name = Trim$(settings.GetRequiredString("Results.SnapshotName"))
    If LCase$(Right$(name, Len(LEGACY_FILE_EXTENSION))) = LEGACY_FILE_EXTENSION Then
        name = Left$(name, Len(name) - Len(LEGACY_FILE_EXTENSION))
    ElseIf LCase$(Right$(name, Len(FILE_EXTENSION))) = FILE_EXTENSION Then
        name = Left$(name, Len(name) - Len(FILE_EXTENSION))
    End If
    Dim token As Variant, invalid As Boolean
    For Each token In Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
        If InStr(name, CStr(token)) > 0 Then invalid = True
    Next token
    Dim i As Long, stem As String
    For i = 1 To Len(name)
        If AscW(Mid$(name, i, 1)) >= 0 And AscW(Mid$(name, i, 1)) < 32 Then invalid = True
    Next i
    stem = UCase$(Split(name & ".", ".")(0))
    If stem = "CON" Or stem = "PRN" Or stem = "AUX" Or stem = "NUL" Then invalid = True
    If Len(stem) = 4 Then
        If (Left$(stem, 3) = "COM" Or Left$(stem, 3) = "LPT") And Mid$(stem, 4, 1) >= "1" And Mid$(stem, 4, 1) <= "9" Then invalid = True
    End If
    If Len(name) = 0 Or Len(name) > 180 Or Right$(name, 1) = "." Or Right$(name, 1) = " " Then invalid = True
    If invalid Then Err.Raise 5, "CCalculationSnapshotStore", settings.InputErrorMessage( _
        "Results.SnapshotName", "Требуется допустимое имя снимка без пути.", "Введите, например, NDM_Snapshot; расширение .ndm добавляется автоматически.")
    name = name & FILE_EXTENSION
    If Len(CStr(workbook.Path)) = 0 Or InStr(CStr(workbook.Path), "://") > 0 Then Err.Raise 5, , "Сначала сохраните книгу в локальную папку."
    SnapshotPath = CStr(workbook.Path) & Application.PathSeparator & name
    If Len(SnapshotPath) > 245 Then Err.Raise 5, , settings.InputErrorMessage("Results.SnapshotName", _
        "Полный путь снимка слишком длинный для файловых операций VBA.", "Сократите имя или путь к папке книги.")
End Function

' Выбирает файл текущего имени Config; .ndm имеет приоритет перед .ndms.
' Повреждённый .ndm не подменяется файлом .ndms: после выбора проверяется SHA256.
Private Function SelectedReadPath(ByVal workbook As Object) As String
    SelectedReadPath = SnapshotPath(workbook)
    If FileExists(SelectedReadPath) Then Exit Function
    Dim previous As String: previous = Left$(SelectedReadPath, Len(SelectedReadPath) - Len(FILE_EXTENSION)) & LEGACY_FILE_EXTENSION
    If FileExists(previous) Then SelectedReadPath = previous
End Function

' Сохраняет полный файл; существующий расчёт заменяет только после успешной записи нового.
Public Sub SaveFile(ByVal workbook As Object, ByVal snapshot As CCalculationSnapshot)
    snapshot.Validate
    Dim target As String: target = SnapshotPath(workbook)
    Dim temporary As String: temporary = target & ".part"
    Dim ownsTemporary As Boolean
    On Error GoTo Failed
    If FileExists(temporary) Then Err.Raise 5, , "Временный файл занят другим сохранением: " & temporary
    ownsTemporary = True
    WriteBinary temporary, snapshot
    If MoveFileExW(StrPtr(temporary), StrPtr(target), 9) = 0 Then Err.Raise 5, , _
        "Не удалось заменить файл снимка. Проверьте права на папку и закройте файл в другой программе. Код Windows: " & CStr(Err.LastDllError)
    Exit Sub
Failed:
    Dim reason As String: reason = Err.Description
    On Error Resume Next
    If ownsTemporary Then If FileExists(temporary) Then DeleteFile temporary
    On Error GoTo 0
    Err.Raise vbObjectError + 4950, "CCalculationSnapshotStore", "Снимок не сохранён (" & target & "): " & reason
End Sub

' Загружает файл, выбранный сейчас в Config, независимо от привязки показанных Results.
Public Function LoadFile(ByVal workbook As Object) As CCalculationSnapshot
    Dim path As String: path = SelectedReadPath(workbook)
    If Not FileExists(path) Then Err.Raise vbObjectError + 4956, "CCalculationSnapshotStore", "Файл снимка не найден: " & path
    Dim snapshot As CCalculationSnapshot: Set snapshot = ReadBinary(path)
    snapshot.Validate
    Set LoadFile = snapshot
End Function

' Записывает матрицы плотным колоночным кодеком. Обязательная фактическая
' граница и явный GeometryPreview требуют NDMSNP05; полный компактный расчёт сохраняет NDMSNP04.
Private Sub WriteBinary(ByVal path As String, ByVal snapshot As CCalculationSnapshot)
    Dim file As CBinarySnapshotFile: Set file = New CBinarySnapshotFile
    Dim count As Long, key As Variant, hash As String, signature As String, version As Long
    On Error GoTo Failed
    For Each key In snapshot.Keys()
        count = count + 1
    Next key
    signature = MAGIC: version = 4
    If HasConcreteBoundaries(snapshot) Or snapshot.IsGeometryPreview Then signature = BOUNDARY_MAGIC: version = 5
    file.OpenWrite path
    file.WriteASCII signature: file.WriteLong version: file.WriteDouble snapshot.RunID: file.WriteDouble snapshot.ElapsedSeconds: file.WriteLong count
    For Each key In snapshot.Keys()
        PutText file, CStr(key): snapshot.WritePreparedTable CStr(key), Me, file
    Next key
    file.CloseFile
    Dim digest As CFileChecksum: Set digest = New CFileChecksum: hash = digest.FileSHA256(path)
    file.OpenUpdate path: file.Position = file.Size
    file.WriteASCII hash: file.WriteASCII signature: file.CloseFile
    Exit Sub
Failed:
    Dim reason As String: reason = Err.Description
    file.CloseFile
    Err.Raise vbObjectError + 4953, "CCalculationSnapshotStore.WriteBinary", reason
End Sub

' Отличает новую обязательную геометрию от пустой дополнительной шапки.
' Contour без обязательных границ отклоняется snapshot.Validate до начала записи.
Private Function HasConcreteBoundaries(ByVal snapshot As CCalculationSnapshot) As Boolean
    If snapshot.HasTable("rngNDMElementGeometry") Then HasConcreteBoundaries = snapshot.TableRowCount("rngNDMElementGeometry") > 1
End Function

' Записывает переданный снимком массив без копирования большой матрицы элементов.
' Таблица остаётся собственностью снимка; ни значения, ни порядок строк не меняются.
Friend Sub WritePreparedMatrix(ByVal file As CBinarySnapshotFile, ByRef values As Variant, Optional ByVal elementCount As Long = 0)
    WriteMatrix file, values, elementCount
End Sub

' Проверяет контрольную сумму до выделения матриц и чтения значений.
Private Function ReadBinary(ByVal path As String) As CCalculationSnapshot
    Dim file As CBinarySnapshotFile: Set file = New CBinarySnapshotFile
    Dim signature As String, footer As String, hash As String, version As Long, count As Long, i As Long, key As String, values As Variant
    On Error GoTo Failed
    file.OpenRead path
    If file.Size < 104 Then Err.Raise 5, , "Файл снимка оборван."
    file.Position = file.Size - 72: hash = file.ReadASCII(64): footer = file.ReadASCII(8)
    Dim digest As CFileChecksum: Set digest = New CFileChecksum
    If (footer <> MAGIC And footer <> BOUNDARY_MAGIC) Or digest.FileSHA256(path, file.Size - 72) <> hash Then Err.Raise 5, , "Контрольная сумма снимка не совпадает."
    file.Position = 0
    signature = file.ReadASCII(8): version = file.ReadLong
    Dim snapshot As CCalculationSnapshot: Set snapshot = New CCalculationSnapshot
    snapshot.RunID = file.ReadDouble: snapshot.ElapsedSeconds = file.ReadDouble: count = file.ReadLong
    If signature <> footer Or count < 1 Or count > 64 Then Err.Raise 5, , "Неизвестный формат снимка."
    If Not ((signature = MAGIC And version = 4) Or (signature = BOUNDARY_MAGIC And version = 5)) Then Err.Raise 5, , "Неизвестная версия формата снимка."
    For i = 1 To count
        key = GetText(file): values = ReadMatrix(file): snapshot.AddTable key, values
    Next i
    If version = 4 And (HasConcreteBoundaries(snapshot) Or snapshot.HasTable("Snapshot.Kind")) Then Err.Raise 5, , "Обязательные границы элементов требуют версии снимка 5."
    If file.Position <> file.Size - 72 Then Err.Raise 5, , "Неверная длина снимка."
    file.CloseFile
    Set ReadBinary = snapshot
    Exit Function
Failed:
    Dim reason As String: reason = Err.Description
    file.CloseFile
    Err.Raise vbObjectError + 4954, "CCalculationSnapshotStore.ReadBinary", "Не удалось прочитать снимок " & path & ": " & reason
End Function

' Сохраняет колонку плотными массивами типов, чисел и индексов Unicode-словаря.
' Повторяющиеся имена сочетаний и материалов не записываются заново для каждого элемента.
Private Sub WriteMatrix(ByVal file As CBinarySnapshotFile, ByRef values As Variant, Optional ByVal elementCount As Long = 0)
    Dim rows As Long, columns As Long, i As Long, j As Long, flags() As Byte, numbers() As Double, codes() As Long
    Dim dictionary As Object, item As Variant, mask As Byte, count As Long, kind As Long
    Dim text As String, previousText As String, previousCode As Long
    rows = UBound(values, 1): columns = UBound(values, 2)
    file.WriteLong rows: file.WriteLong columns
    For j = 1 To columns
        If rows <= 128 Then
            file.WriteByte 5
            For i = 1 To rows: WriteScalar file, values(i, j): Next i
            GoTo NextColumn
        End If
        If WriteRepeatedColumn(file, values, j, rows) Then GoTo NextColumn
        If elementCount > 0 And j = 5 Then
            If WriteRepeatedElementIDs(file, values, rows, elementCount) Then GoTo NextColumn
        End If
        ReDim flags(1 To rows): ReDim numbers(1 To rows): ReDim codes(1 To rows)
        Set dictionary = CreateObject("Scripting.Dictionary"): dictionary.CompareMode = vbBinaryCompare: mask = 0: previousCode = 0
        For i = 1 To rows
            kind = VarType(values(i, j))
            Select Case kind
                Case vbEmpty
                Case vbString
                    flags(i) = 1: mask = mask Or 2
                    text = CStr(values(i, j))
                    If previousCode = 0 Or StrComp(text, previousText, vbBinaryCompare) <> 0 Then
                        If Not dictionary.Exists(text) Then dictionary.Add text, dictionary.Count + 1
                        previousCode = CLng(dictionary.Item(text)): previousText = text
                    End If
                    codes(i) = previousCode
                Case vbBoolean
                    flags(i) = 3: mask = mask Or 1
                    If CBool(values(i, j)) Then numbers(i) = 1#
                Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate
                    flags(i) = 2: numbers(i) = CDbl(values(i, j)): mask = mask Or 1
                Case Else: Err.Raise 5, , "Неподдерживаемое значение снимка, тип " & CStr(kind)
            End Select
        Next i
        file.WriteByte mask: file.WriteBytes flags
        If (mask And 1) <> 0 Then file.WriteDoubles numbers
        If (mask And 2) <> 0 Then
            file.WriteLongs codes: count = dictionary.Count: file.WriteLong count
            If Not WritePackedDictionary(file, dictionary) Then
        For Each item In dictionary.Keys: PutText file, CStr(item): Next item
    End If
        End If
NextColumn:
    Next j
End Sub

' Проверяет повторение точного списка ElementID во всех сохранённых состояниях.
' Создаёт словарь первого блока; несовпадение выбирает обычный кодек до записи байтов.
Private Function WriteRepeatedElementIDs(ByVal file As CBinarySnapshotFile, ByRef values As Variant, ByVal rows As Long, ByVal elementCount As Long) As Boolean
    If elementCount < 1 Or rows < 4096 Then Exit Function
    If (rows - 1) Mod elementCount <> 0 Or (rows - 1) \ elementCount < 2 Then Exit Function
    Dim dictionary As Object: Set dictionary = CreateObject("Scripting.Dictionary"): dictionary.CompareMode = vbBinaryCompare
    Dim flags() As Byte, codes() As Long, row As Long, first As Long, text As String, key As Variant
    ReDim flags(1 To rows): ReDim codes(1 To rows)
    For row = 1 To elementCount + 1
        If VarType(values(row, 5)) <> vbString Then Exit Function
        text = CStr(values(row, 5))
        If dictionary.Exists(text) Then Exit Function
        dictionary.Add text, row: flags(row) = 1: codes(row) = row
    Next row
    first = 2
    For row = elementCount + 2 To rows
        If VarType(values(row, 5)) <> vbString Then Exit Function
        If StrComp(CStr(values(row, 5)), CStr(values(first, 5)), vbBinaryCompare) <> 0 Then Exit Function
        flags(row) = 1: codes(row) = first
        first = first + 1: If first > elementCount + 1 Then first = 2
    Next row
    file.WriteByte 2: file.WriteBytes flags: file.WriteLongs codes: file.WriteLong dictionary.Count
    If Not WritePackedDictionary(file, dictionary) Then
        For Each key In dictionary.Keys: PutText file, CStr(key): Next key
    End If
    WriteRepeatedElementIDs = True
End Function

' Готовит крупный словарь в один непрерывный блок с Long-длинами и текстом UTF-16LE.
' Предел 8 МиБ ограничивает дополнительную память; маленькие и слишком большие словари идут потоком.
Private Function WritePackedDictionary(ByVal file As CBinarySnapshotFile, ByVal dictionary As Object) As Boolean
    If dictionary.Count < 512 Then Exit Function
    Dim keys As Variant: keys = dictionary.Keys
    Dim key As Variant, total As Double, text As String
    For Each key In keys
        total = total + 4# + LenB(CStr(key))
        If total > 8388608# Then Exit Function
    Next key
    Dim data() As Byte: ReDim data(0 To CLng(total) - 1)
    Dim offset As Long, count As Long
    For Each key In keys
        text = CStr(key): count = LenB(text)
        CopyPackedBytes VarPtr(data(offset)), VarPtr(count), 4
        offset = offset + 4
        If count > 0 Then CopyPackedBytes VarPtr(data(offset)), StrPtr(text), count
        offset = offset + count
    Next key
    file.WriteBytes data
    WritePackedDictionary = True
End Function

' Небольшие сводки и шапки не создают сотни отдельных COM-словарей.
' Буфер потока объединяет короткие значения в крупные физические записи.
Private Sub WriteScalar(ByVal file As CBinarySnapshotFile, ByRef value As Variant)
    Select Case VarType(value)
        Case vbEmpty: file.WriteByte 0
        Case vbString: file.WriteByte 1: PutText file, CStr(value)
        Case vbBoolean: file.WriteByte 3: file.WriteDouble CDbl(value)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate
            file.WriteByte 2: file.WriteDouble CDbl(value)
        Case Else: Err.Raise 5, , "Неподдерживаемый тип значения снимка: " & CStr(VarType(value))
    End Select
End Sub

' Восстанавливает явный тип малой ячейки; строки никогда не становятся формулами.
Private Function ReadScalar(ByVal file As CBinarySnapshotFile) As Variant
    Select Case file.ReadByte
        Case 0: ReadScalar = Empty
        Case 1: ReadScalar = GetText(file)
        Case 2: ReadScalar = file.ReadDouble
        Case 3: ReadScalar = (file.ReadDouble <> 0#)
        Case Else: Err.Raise 5, , "Повреждён тип малого значения снимка."
    End Select
End Function

' Служебные поля состояния хранятся сериями без словарного COM-вызова на каждый элемент.
' На плотной колонке проба прекращается рано; числа сравниваются по всем восьми байтам.
Private Function WriteRepeatedColumn(ByVal file As CBinarySnapshotFile, ByRef values As Variant, ByVal column As Long, ByVal rows As Long) As Boolean
    If rows < 256 Then Exit Function
    Dim limit As Long: limit = rows \ 128
    If limit < 8 Then limit = 8
    If limit > 512 Then limit = 512
    Dim lengths() As Long, items() As Variant, runs As Long, row As Long, value As Variant, previous As Variant
    ReDim lengths(1 To limit): ReDim items(1 To limit)
    Dim kind As Long, previousKind As Long, number As Double, bits As Currency, previousBits As Currency, same As Boolean
    For row = 1 To rows
        value = values(row, column): kind = VarType(value): same = False
        If kind = previousKind And runs > 0 Then
            Select Case kind
                Case vbEmpty: same = True
                Case vbString: same = (StrComp(CStr(value), CStr(previous), vbBinaryCompare) = 0)
                Case vbBoolean: same = (CBool(value) = CBool(previous))
                Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate
                    number = CDbl(value): CopySnapshotBits bits, number, 8
                    same = (bits = previousBits)
            End Select
        End If
        If Not same Then
            If runs = limit Then Exit Function
            runs = runs + 1: items(runs) = value: previous = value: previousKind = kind
            If IsNumeric(value) And kind <> vbString And kind <> vbBoolean Then
                number = CDbl(value): CopySnapshotBits previousBits, number, 8
            End If
        End If
        lengths(runs) = lengths(runs) + 1
    Next row
    file.WriteByte 4: file.WriteLong runs
    For row = 1 To runs
        file.WriteLong lengths(row): kind = VarType(items(row))
        Select Case kind
            Case vbEmpty: file.WriteByte 0
            Case vbString: file.WriteByte 1: PutText file, CStr(items(row))
            Case vbBoolean
                file.WriteByte 3: file.WriteDouble CDbl(items(row))
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate
                file.WriteByte 2: file.WriteDouble CDbl(items(row))
            Case Else: Err.Raise 5, , "Неподдерживаемое значение серии снимка, тип " & CStr(kind)
        End Select
    Next row
    WriteRepeatedColumn = True
End Function

' Восстанавливает значения без десятичных преобразований и без исполнения формул.
Private Function ReadMatrix(ByVal file As CBinarySnapshotFile) As Variant
    Dim rows As Long, columns As Long, i As Long, j As Long, count As Long, mask As Byte
    rows = file.ReadLong: columns = file.ReadLong
    If rows < 1 Or rows > 10000000 Or columns < 1 Or columns > 256 Or CDbl(rows) * columns > 100000000# Then Err.Raise 5, , "Неверный размер матрицы снимка."
    Dim values() As Variant, flags() As Byte, numbers() As Double, codes() As Long, strings() As String
    ReDim values(1 To rows, 1 To columns)
    For j = 1 To columns
        mask = file.ReadByte
        If mask = 5 Then
            For i = 1 To rows: values(i, j) = ReadScalar(file): Next i
            GoTo NextColumn
        End If
        If mask = 4 Then
            ReadRepeatedColumn file, values, j, rows
            GoTo NextColumn
        End If
        If mask > 3 Then Err.Raise 5, , "Неизвестный тип колонки снимка."
        ReDim flags(1 To rows): ReDim numbers(1 To rows): ReDim codes(1 To rows)
        flags = file.ReadBytes(rows)
        If (mask And 1) <> 0 Then numbers = file.ReadDoubles(rows)
        If (mask And 2) <> 0 Then
            codes = file.ReadLongs(rows): count = file.ReadLong
            If count < 1 Or count > rows Then Err.Raise 5, , "Повреждён словарь снимка."
            ReDim strings(1 To count)
            For i = 1 To count: strings(i) = GetText(file): Next i
        End If
        For i = 1 To rows
            Select Case flags(i - 1)
                Case 0
                Case 1
                    If (mask And 2) = 0 Or codes(i) < 1 Or codes(i) > count Then Err.Raise 5, , "Повреждён индекс строки снимка."
                    values(i, j) = strings(codes(i))
                Case 2: values(i, j) = numbers(i)
                Case 3: values(i, j) = (numbers(i) <> 0#)
                Case Else: Err.Raise 5, , "Неизвестное значение снимка."
            End Select
        Next i
NextColumn:
    Next j
    ReadMatrix = values
End Function

' Восстанавливает серию с проверкой вместимости до заполнения массива.
Private Sub ReadRepeatedColumn(ByVal file As CBinarySnapshotFile, ByRef values() As Variant, ByVal column As Long, ByVal rows As Long)
    Dim count As Long, length As Long, run As Long, row As Long, first As Long, kind As Byte, value As Variant
    count = file.ReadLong: first = 1
    If count < 1 Or count > rows Then Err.Raise 5, , "Неверное число серий колонки снимка."
    For run = 1 To count
        length = file.ReadLong: kind = file.ReadByte
        If length < 1 Or length > rows - first + 1 Then Err.Raise 5, , "Повреждена длина серии снимка."
        Select Case kind
            Case 0: value = Empty
            Case 1: value = GetText(file)
            Case 2: value = file.ReadDouble
            Case 3: value = (file.ReadDouble <> 0#)
            Case Else: Err.Raise 5, , "Неизвестный тип серии снимка."
        End Select
        For row = first To first + length - 1: values(row, column) = value: Next row
        first = first + length
    Next run
    If first <> rows + 1 Then Err.Raise 5, , "Оборвана колонка серий снимка."
End Sub

' Сохраняет UTF-16LE с явной длиной, не завися от системной ANSI-кодировки.
Private Sub PutText(ByVal file As CBinarySnapshotFile, ByVal value As String)
    file.WriteSizedUTF16 value
End Sub

' Проверяет размер строки до выделения памяти и чтения точного Unicode.
Private Function GetText(ByVal file As CBinarySnapshotFile) As String
    GetText = file.ReadSizedUTF16(131072)
End Function

' После переноса книги копирует связанный снимок, если рядом его ещё нет.
' Одноимённый иной расчёт не заменяет; обычное сохранение не переписывает файл.
Public Sub EnsureBesideWorkbook(ByVal workbook As Object, Optional ByVal previousFolder As String = "")
    If Not HasFileReference(workbook) Then Exit Sub
    If StorageMode(workbook) <> "File" Then Exit Sub
    Dim fileSystem As Object: Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Dim target As String: target = CStr(workbook.Path) & Application.PathSeparator & LinkText(workbook, "File")
    Dim source As String: source = LinkText(workbook, "Folder") & Application.PathSeparator & LinkText(workbook, "File")
    Dim digest As CFileChecksum: Set digest = New CFileChecksum
    If Not fileSystem.FileExists(target) Then
        If Len(previousFolder) > 0 Then
            Dim previous As String: previous = previousFolder & Application.PathSeparator & LinkText(workbook, "File")
            If fileSystem.FileExists(previous) Then
                If digest.FileSHA256(previous) = LinkText(workbook, "Hash") Then source = previous
            End If
        End If
        If digest.FileSHA256(source) <> LinkText(workbook, "Hash") Then Err.Raise 5, , "Исходный снимок изменился: " & source
        fileSystem.CopyFile source, target, False
    End If
    If digest.FileSHA256(target) <> LinkText(workbook, "Hash") Then Err.Raise 5, , _
        "Рядом с книгой находится другой одноимённый снимок: " & target & ". Исходный файл сохранён в " & source
End Sub

' Обновляет только фиксированное поле времени и SHA256, не собирая большие таблицы заново.
' Excel-режим хранит время только в верхней сводке, без сериализованного дополнения.
Public Sub UpdateSavedElapsed(ByVal workbook As Object, ByVal snapshot As CCalculationSnapshot)
    If StorageMode(workbook) = "Excel" Then Exit Sub
    Dim target As String: target = SnapshotPath(workbook)
    Dim temporary As String: temporary = target & ".part"
    Dim ownsTemporary As Boolean, elapsed As Double, hash As String, signature As String, version As Long, footer As String
    Dim file As CBinarySnapshotFile: Set file = New CBinarySnapshotFile
    On Error GoTo Failed
    If FileExists(temporary) Then Err.Raise 5, , "Временный файл снимка занят: " & temporary
    ownsTemporary = True: CreateObject("Scripting.FileSystemObject").CopyFile target, temporary, False
    file.OpenRead temporary: signature = file.ReadASCII(8): version = file.ReadLong
    file.Position = file.Size - 8: footer = file.ReadASCII(8): file.CloseFile
    If signature <> footer Or Not ((signature = MAGIC And version = 4) Or (signature = BOUNDARY_MAGIC And version = 5)) Then _
        Err.Raise 5, , "Неизвестная версия формата снимка при обновлении времени."
    file.OpenUpdate temporary: file.Position = 20: file.WriteDouble snapshot.ElapsedSeconds
    Dim size As Long: size = file.Size: file.CloseFile
    Dim digest As CFileChecksum: Set digest = New CFileChecksum: hash = digest.FileSHA256(temporary, size - 72)
    file.OpenUpdate temporary: file.Position = size - 72
    file.WriteASCII hash: file.WriteASCII signature: file.CloseFile
    If MoveFileExW(StrPtr(temporary), StrPtr(target), 9) = 0 Then Err.Raise 5, , "Не удалось обновить время снимка: " & target
    LinkFile workbook
    Exit Sub
Failed:
    Dim reason As String: reason = Err.Description
    On Error Resume Next
    file.CloseFile
    If ownsTemporary Then If FileExists(temporary) Then DeleteFile temporary
    On Error GoTo 0
    Err.Raise 5, "CCalculationSnapshotStore.UpdateSavedElapsed", reason
End Sub

' Проверяет существование Unicode-имени без преобразования в системную ANSI-кодировку.
Private Function FileExists(ByVal path As String) As Boolean
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(path)
End Function

' Удаляет только явно созданный собственный временный файл.
Private Sub DeleteFile(ByVal path As String)
    CreateObject("Scripting.FileSystemObject").DeleteFile path, True
End Sub
