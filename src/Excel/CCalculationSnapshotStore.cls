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
    SetLink workbook, "Hash", FileIdentity(path)
End Sub

' Читает прежнюю служебную привязку для проверки переноса/диагностики.
' Пользовательские действия выбирают текущий файл через LoadFile, а не эту привязку.
Public Function LoadLinkedFile(ByVal workbook As Object) As CCalculationSnapshot
    Dim path As String: path = CStr(workbook.Path) & Application.PathSeparator & LinkText(workbook, "File")
    If Not FileExists(path) Then path = LinkText(workbook, "Folder") & Application.PathSeparator & LinkText(workbook, "File")
    Dim digest As CFileChecksum: Set digest = New CFileChecksum
    If FileIdentity(path) <> LinkText(workbook, "Hash") Then Err.Raise vbObjectError + 4955, , _
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

' Называет фактически выбранный .ndm/.ndms без пути для короткого сообщения об источнике данных.
Public Function SelectedFileName(ByVal workbook As Object) As String
    SelectedFileName = CreateObject("Scripting.FileSystemObject").GetFileName(SelectedReadPath(workbook))
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
Friend Function SelectedReadPath(ByVal workbook As Object) As String
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
    If Not FileExists(path) Then
        Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader
        settings.LoadFromRange workbook.Names.Item("rngSystemSettings").RefersToRange
        Err.Raise vbObjectError + 4956, "CCalculationSnapshotStore", _
            "Файл снимка «" & CreateObject("Scripting.FileSystemObject").GetFileName(path) & "» не найден." & vbCrLf & vbCrLf & _
            "Проверьте имя Results.SnapshotName и поместите файл в папку книги." & vbCrLf & _
            "Если снимка ещё нет, нажмите «Сформировать сечение»." & vbCrLf & vbCrLf & _
            "Настройка: " & settings.InputLocation("Results.SnapshotName") & vbCrLf & "Ожидаемый файл:" & vbCrLf & path
    End If
    Dim snapshot As CCalculationSnapshot: Set snapshot = ReadBinary(path)
    snapshot.Validate
    Set LoadFile = snapshot
End Function

' Отличает новую обязательную геометрию от пустой дополнительной шапки.
' Contour без обязательных границ отклоняется snapshot.Validate до начала записи.
Private Function HasConcreteBoundaries(ByVal snapshot As CCalculationSnapshot) As Boolean
    If snapshot.HasTable("rngNDMElementGeometry") Then HasConcreteBoundaries = snapshot.TableRowCount("rngNDMElementGeometry") > 1
End Function

' Полное восстановление v6 проверяет каждый блок; v4/v5 проверяет общий SHA до выделения матриц.
Private Function ReadBinary(ByVal path As String) As CCalculationSnapshot
    Dim indexed As New CIndexedSnapshotFile
    If indexed.OpenRead(path, Me) Then Set ReadBinary = indexed.LoadAll: Exit Function
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
        key = GetText(file): values = ReadIndexedValues(file): snapshot.AddTable key, values
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

' Восстанавливает значения без десятичных преобразований и без исполнения формул.
Friend Function ReadIndexedValues(ByVal file As CBinarySnapshotFile, Optional ByVal expectedRows As Long = 0, Optional ByVal expectedColumns As Long = 0) As Variant
    Dim rows As Long, columns As Long, i As Long, j As Long, count As Long, mask As Byte
    rows = file.ReadLong: columns = file.ReadLong
    If rows < 1 Or rows > 10000000 Or columns < 1 Or columns > 256 Or CDbl(rows) * columns > 100000000# Then Err.Raise 5, , "Неверный размер матрицы снимка."
    If expectedRows > 0 Then If rows <> expectedRows Then Err.Raise 5, , "Число строк блока не совпадает с каталогом."
    If expectedColumns > 0 Then If columns <> expectedColumns Then Err.Raise 5, , "Число колонок блока не совпадает с каталогом."
    If expectedRows = 0 And expectedColumns = 17 Then If rows > 100000 Then Err.Raise 5, , "Превышено число записей каталога."
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
    ReadIndexedValues = values
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

' После переноса книги копирует подтверждённый снимок только текущего имени.
' Привязка другого имени не участвует; одноимённый иной расчёт не заменяется.
Public Sub EnsureBesideWorkbook(ByVal workbook As Object, Optional ByVal previousFolder As String = "")
    If Not HasFileReference(workbook) Then Exit Sub
    If StorageMode(workbook) <> "File" Then Exit Sub
    Dim fileSystem As Object: Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Dim name As String: name = LinkText(workbook, "File")
    If name <> fileSystem.GetFileName(name) Then Err.Raise 5, , "В привязке снимка требуется имя без пути."
    If StrComp(fileSystem.GetBaseName(SnapshotPath(workbook)), fileSystem.GetBaseName(name), vbTextCompare) <> 0 Then Exit Sub
    Dim target As String: target = CStr(workbook.Path) & Application.PathSeparator & name
    Dim source As String: source = LinkText(workbook, "Folder") & Application.PathSeparator & name
    If fileSystem.FileExists(target) Then
        If FileIdentity(target) <> LinkText(workbook, "Hash") Then Err.Raise 5, , "Рядом с книгой находится другой одноимённый снимок: " & target
        Exit Sub ' Обычное сохранение книги проверяет каталог v6; подробные блоки остаются адресными.
    End If
    If Len(previousFolder) > 0 Then
        Dim previous As String: previous = previousFolder & Application.PathSeparator & name
        If fileSystem.FileExists(previous) Then
            If FileIdentity(previous) = LinkText(workbook, "Hash") Then source = previous
        End If
    End If
    If FileIdentity(source) <> LinkText(workbook, "Hash") Then Err.Raise 5, , "Исходный снимок изменился: " & source
    Dim temporary As String: temporary = target & ".part"
    Dim ownsTemporary As Boolean
    On Error GoTo Failed
    If FileExists(temporary) Then Err.Raise 5, , "Временный файл переноса занят: " & temporary
    ownsTemporary = True: fileSystem.CopyFile source, temporary, False
    If FileIdentity(temporary) <> LinkText(workbook, "Hash") Then Err.Raise 5, , "Копия не соответствует привязанному снимку."
    VerifyPath temporary ' Проверка всех блоков выполняется только при фактическом переносе, до публикации копии.
    If MoveFileExW(StrPtr(temporary), StrPtr(target), 8) = 0 Then Err.Raise 5, , "Не удалось опубликовать снимок рядом с новой книгой. Код Windows: " & CStr(Err.LastDllError)
    Exit Sub
Failed:
    Dim number As Long, reason As String: number = Err.Number: reason = Err.Description
    On Error Resume Next
    If ownsTemporary Then If FileExists(temporary) Then DeleteFile temporary
    On Error GoTo 0
    Err.Raise number, "CCalculationSnapshotStore.EnsureBesideWorkbook", reason
End Sub

' Явно проверяет все блоки текущего выбранного файла, не выдавая локальную проверку за полную.
Public Sub VerifyFile(ByVal workbook As Object)
    VerifyPath SelectedReadPath(workbook)
End Sub

' Для v6 проверяет весь поток по блоковым SHA; v4/v5 сохраняют полный совместимый reader.
Private Sub VerifyPath(ByVal path As String)
    Dim indexed As New CIndexedSnapshotFile
    If indexed.OpenRead(path, Me) Then
        indexed.VerifyAllBlocks
    Else
        Dim snapshot As CCalculationSnapshot: Set snapshot = ReadBinary(path): snapshot.Validate
    End If
End Sub

' Обновляет только фиксированное поле времени и SHA256, не собирая большие таблицы заново.
' Excel-режим хранит время только в верхней сводке, без сериализованного дополнения.
Public Sub UpdateSavedElapsed(ByVal workbook As Object, ByVal snapshot As CCalculationSnapshot)
    UpdateElapsedValue workbook, snapshot.ElapsedSeconds, snapshot.RunID
End Sub

' Обновляет только длительность подтверждённого запуска; не принимает частичный объект полного снимка.
Public Sub UpdateElapsedValue(ByVal workbook As Object, ByVal seconds As Double, Optional ByVal expectedRunID As Double = 0#)
    If StorageMode(workbook) = "Excel" Then Exit Sub
    Dim target As String: target = SnapshotPath(workbook)
    Dim temporary As String: temporary = target & ".part"
    Dim ownsTemporary As Boolean, elapsed As Double, hash As String, signature As String, version As Long, footer As String
    Dim file As CBinarySnapshotFile: Set file = New CBinarySnapshotFile
    On Error GoTo Failed
    If FileExists(temporary) Then Err.Raise 5, , "Временный файл снимка занят: " & temporary
    ownsTemporary = True: CreateObject("Scripting.FileSystemObject").CopyFile target, temporary, False
    Dim indexed As New CIndexedSnapshotFile
    If indexed.OpenRead(temporary, Me) Then
        If expectedRunID <> 0# Then If indexed.RunID <> expectedRunID Then Err.Raise 5, , "Обновляемое время относится к другому расчёту."
        indexed.UpdateElapsed temporary, seconds
        Set indexed = Nothing
        If MoveFileExW(StrPtr(temporary), StrPtr(target), 9) = 0 Then Err.Raise 5, , "Не удалось обновить время снимка: " & target
        LinkFile workbook
        Exit Sub
    End If
    file.OpenRead temporary: signature = file.ReadASCII(8): version = file.ReadLong
    file.Position = file.Size - 8: footer = file.ReadASCII(8): file.CloseFile
    If signature <> footer Or Not ((signature = MAGIC And version = 4) Or (signature = BOUNDARY_MAGIC And version = 5)) Then _
        Err.Raise 5, , "Неизвестная версия формата снимка при обновлении времени."
    file.OpenUpdate temporary: file.Position = 20: file.WriteDouble seconds
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
    indexed.CloseFile
    Set indexed = Nothing
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

' Пишет только переданный непрерывный участок исходного массива, не копируя всю таблицу.
' Штатный колоночный кодек и повторяющиеся серии сохранены; индексы колонок не зависят от Excel.
Friend Sub WriteIndexedValues(ByVal file As CBinarySnapshotFile, ByRef values As Variant, Optional ByVal first As Long = 1, Optional ByVal last As Long = 0, Optional ByVal firstColumn As Long = 1, Optional ByVal lastColumn As Long = 0)
    Dim rows As Long, columns As Long, i As Long, j As Long, flags() As Byte, numbers() As Double, codes() As Long
    Dim dictionary As Object, item As Variant, mask As Byte, count As Long, kind As Long
    Dim text As String, previousText As String, previousCode As Long
    If last = 0 Then last = UBound(values, 1)
    If lastColumn = 0 Then lastColumn = UBound(values, 2)
    If first < 1 Or first > last Or last > UBound(values, 1) Or firstColumn < 1 Or firstColumn > lastColumn Or lastColumn > UBound(values, 2) Then Err.Raise 5, , "Неверные границы записываемого участка."
    rows = last - first + 1: columns = lastColumn - firstColumn + 1
    file.WriteLong rows: file.WriteLong columns
    For j = firstColumn To lastColumn
        If rows <= 128 Then
            file.WriteByte 5
            For i = 1 To rows: WriteScalar file, values(i + first - 1, j): Next i
            GoTo NextColumn
        End If
        If WriteIndexedRepeatedColumn(file, values, j, rows, first) Then GoTo NextColumn
        ReDim flags(1 To rows): ReDim numbers(1 To rows): ReDim codes(1 To rows)
        Set dictionary = CreateObject("Scripting.Dictionary"): dictionary.CompareMode = vbBinaryCompare: mask = 0: previousCode = 0
        For i = 1 To rows
            kind = VarType(values(i + first - 1, j))
            Select Case kind
                Case vbEmpty
                Case vbString
                    flags(i) = 1: mask = mask Or 2
                    text = CStr(values(i + first - 1, j))
                    If previousCode = 0 Or StrComp(text, previousText, vbBinaryCompare) <> 0 Then
                        If Not dictionary.Exists(text) Then dictionary.Add text, dictionary.Count + 1
                        previousCode = CLng(dictionary.Item(text)): previousText = text
                    End If
                    codes(i) = previousCode
                Case vbBoolean
                    flags(i) = 3: mask = mask Or 1
                    If CBool(values(i + first - 1, j)) Then numbers(i) = 1#
                Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate
                    flags(i) = 2: numbers(i) = CDbl(values(i + first - 1, j)): mask = mask Or 1
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


' Сжимает серии одного участка с точным сравнением Double и Unicode.
Private Function WriteIndexedRepeatedColumn(ByVal file As CBinarySnapshotFile, ByRef values As Variant, ByVal column As Long, ByVal rows As Long, ByVal first As Long) As Boolean
    If rows < 256 Then Exit Function
    Dim limit As Long: limit = rows \ 128
    If limit < 8 Then limit = 8
    If limit > 512 Then limit = 512
    Dim lengths() As Long, items() As Variant, runs As Long, row As Long, value As Variant, previous As Variant
    ReDim lengths(1 To limit): ReDim items(1 To limit)
    Dim kind As Long, previousKind As Long, number As Double, bits As Currency, previousBits As Currency, same As Boolean
    For row = 1 To rows
        value = values(row + first - 1, column): kind = VarType(value): same = False
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
    WriteIndexedRepeatedColumn = True
End Function


' Скалярный кодек общий для плотных таблиц и четырёх ключей независимого блока.
Friend Sub WriteIndexedScalar(ByVal file As CBinarySnapshotFile, ByRef value As Variant)
    WriteScalar file, value
End Sub

' Восстанавливает тип скалярного ключа без преобразования строки в число.
Friend Function ReadIndexedScalar(ByVal file As CBinarySnapshotFile) As Variant
    ReadIndexedScalar = ReadScalar(file)
End Function

' Новые сохранения используют индексированный v6; v4/v5 остаются читаемыми без конвертации.
Private Sub WriteBinary(ByVal path As String, ByVal snapshot As CCalculationSnapshot)
    Dim indexed As New CIndexedSnapshotFile
    indexed.WriteFile path, snapshot, Me
End Sub

' Проверяет каталог v6 для привязки; совместимость v4/v5 использует их полный файловый SHA256.
Private Function FileIdentity(ByVal path As String) As String
    Dim indexed As New CIndexedSnapshotFile
    If indexed.OpenRead(path, Me) Then
        FileIdentity = indexed.Identity
    Else
        Dim digest As New CFileChecksum: FileIdentity = digest.FileSHA256(path)
    End If
End Function
