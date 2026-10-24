VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CBinarySnapshotFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Двоичный поток снимка с Unicode-именами штатными средствами Windows
' ==========================================================================
' Массивы записываются блоками в исходном представлении Windows little-endian.
' Файл ограничен 2 ГиБ для проверяемых Long-смещений; иных файлов класс не открывает.
Private Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal name As LongPtr, ByVal access As Long, ByVal sharing As Long, ByVal security As LongPtr, ByVal creation As Long, ByVal attributes As Long, ByVal template As LongPtr) As LongPtr
Private Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal file As LongPtr, ByVal target As LongPtr, ByVal count As Long, ByRef readCount As Long, ByVal overlapped As LongPtr) As Long
Private Declare PtrSafe Function WriteFile Lib "kernel32" (ByVal file As LongPtr, ByVal source As LongPtr, ByVal count As Long, ByRef writtenCount As Long, ByVal overlapped As LongPtr) As Long
Private Declare PtrSafe Function GetFileSize Lib "kernel32" (ByVal file As LongPtr, ByVal high As LongPtr) As Long
Private Declare PtrSafe Function SetFilePointer Lib "kernel32" (ByVal file As LongPtr, ByVal distance As Long, ByVal high As LongPtr, ByVal origin As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal handle As LongPtr) As Long
Private Declare PtrSafe Sub CopyStreamBytes Lib "kernel32" Alias "RtlMoveMemory" (ByVal target As LongPtr, ByVal source As LongPtr, ByVal count As LongPtr)
Private mHandle As LongPtr ' Принадлежит только этому объекту; INVALID_HANDLE_VALUE не хранится.
Private mPosition As Long ' Смещение следующего байта от начала файла, начиная с нуля.
Private mSize As Long ' Проверенная длина доступного файла.
Private mPath As String ' Имя для сообщения о фактической файловой ошибке.
Private mWriteBuffer(0 To 65535) As Byte ' Объединяет короткие записи служебных полей в блоки по 64 КиБ.
Private mBuffered As Long ' Логическая позиция включает эти ещё не отправленные Windows байты.
Private mReadBuffer(0 To 65535) As Byte ' Небольшие поля и строки читаются блоками, без отдельного ReadFile для каждого значения.
Private mReadCount As Long, mReadOffset As Long ' Физический указатель опережает Position только на непрочитанный остаток этого буфера.

' Открывает существующий файл на чтение, не создавая пустой файл при опечатке.
Public Sub OpenRead(ByVal path As String)
    OpenFile path, &H80000000, 1, 3
End Sub

' Создаёт или полностью заменяет только явно переданный временный файл.
Public Sub OpenWrite(ByVal path As String)
    OpenFile path, &H40000000, 0, 2
End Sub

' Открывает существующий собственный временный файл для точечного обновления.
Public Sub OpenUpdate(ByVal path As String)
    OpenFile path, &HC0000000, 0, 3
End Sub

' Проверяет дескриптор и размер до первой операции с данными.
Private Sub OpenFile(ByVal path As String, ByVal access As Long, ByVal sharing As Long, ByVal creation As Long)
    CloseFile
    mPath = path
    mHandle = CreateFileW(StrPtr(path), access, sharing, 0, creation, 128, 0)
    If mHandle = -1 Then mHandle = 0: Err.Raise 5, "CBinarySnapshotFile", "Не удалось открыть файл: " & path & ". Код Windows: " & CStr(Err.LastDllError)
    Dim high As Long: mSize = GetFileSize(mHandle, VarPtr(high))
    If high <> 0 Or mSize < 0 Then CloseFile: Err.Raise 5, "CBinarySnapshotFile", "Снимок превышает поддерживаемый размер 2 ГиБ: " & path
    mPosition = 0
End Sub

' Закрывает только собственный дескриптор, в том числе после ошибки вызывающего кода.
Public Sub CloseFile()
    On Error GoTo Failed
    If mHandle <> 0 Then FlushWrites: CloseHandle mHandle: mHandle = 0
    mReadCount = 0: mReadOffset = 0
    Exit Sub
Failed:
    Dim number As Long, reason As String: number = Err.Number: reason = Err.Description
    If mHandle <> 0 Then CloseHandle mHandle: mHandle = 0
    mBuffered = 0: mReadCount = 0: mReadOffset = 0
    Err.Raise number, "CBinarySnapshotFile.CloseFile", reason
End Sub

' Гарантирует закрытие потока при освобождении класса.
Private Sub Class_Terminate()
    On Error Resume Next ' Явное CloseFile сообщает ошибку записи; освобождение объекта обязано закрыть дескриптор.
    CloseFile
End Sub

' Возвращает длину файла в байтах.
Public Property Get Size() As Long
    Size = mSize
End Property

' Возвращает нулевое смещение следующего чтения или записи.
Public Property Get Position() As Long
    Position = mPosition
End Property

' Перемещает позицию только внутри существующего файла; случайные пропуски запрещены.
Public Property Let Position(ByVal value As Long)
    If value < 0 Or value > mSize Or mHandle = 0 Then Err.Raise 5, "CBinarySnapshotFile", "Неверное смещение снимка: " & mPath
    FlushWrites
    If SetFilePointer(mHandle, value, 0, 0) <> value Then Err.Raise 5, "CBinarySnapshotFile", "Не удалось изменить позицию: " & mPath
    mPosition = value: mReadCount = 0: mReadOffset = 0
End Property

' Читает строго заданное число байтов, не допускает незаметного усечения массива.
Private Sub ReadExact(ByVal target As LongPtr, ByVal count As Long)
    Dim done As Long, copied As Long, take As Long, remaining As Long
    If count < 0 Or count > mSize - mPosition Or mHandle = 0 Then Err.Raise 5, "CBinarySnapshotFile", "Файл снимка оборван: " & mPath
    If count = 0 Then Exit Sub
    FlushWrites
    Do While copied < count
        remaining = count - copied
        If mReadOffset < mReadCount Then
            take = mReadCount - mReadOffset: If take > remaining Then take = remaining
            CopyStreamBytes target + copied, VarPtr(mReadBuffer(mReadOffset)), take
            mReadOffset = mReadOffset + take: mPosition = mPosition + take: copied = copied + take
        ElseIf remaining >= 65536 Then
            If ReadFile(mHandle, target + copied, remaining, done, 0) = 0 Or done <> remaining Then Err.Raise 5, "CBinarySnapshotFile", "Ошибка чтения файла: " & mPath
            mPosition = mPosition + remaining: copied = count
            mReadCount = 0: mReadOffset = 0
        Else
            take = mSize - mPosition: If take > 65536 Then take = 65536
            If ReadFile(mHandle, VarPtr(mReadBuffer(0)), take, done, 0) = 0 Or done <> take Then Err.Raise 5, "CBinarySnapshotFile", "Ошибка чтения файла: " & mPath
            mReadCount = take: mReadOffset = 0
        End If
    Loop
End Sub

' Пишет блок целиком; ограничение размера проверяет до переполнения Long.
Private Sub WriteExact(ByVal source As LongPtr, ByVal count As Long)
    Dim done As Long
    If count < 0 Or CDbl(mPosition) + count > 2147483647# Or mHandle = 0 Then Err.Raise 5, "CBinarySnapshotFile", "Превышен размер снимка: " & mPath
    If count = 0 Then Exit Sub
    If mReadOffset < mReadCount Then
        If SetFilePointer(mHandle, mPosition, 0, 0) <> mPosition Then Err.Raise 5, "CBinarySnapshotFile", "Не удалось изменить позицию: " & mPath
    End If
    mReadCount = 0: mReadOffset = 0
    If count <= 65536 Then
        If mBuffered + count > 65536 Then FlushWrites
        CopyStreamBytes VarPtr(mWriteBuffer(mBuffered)), source, count
        mBuffered = mBuffered + count
    Else
        FlushWrites
        If WriteFile(mHandle, source, count, done, 0) = 0 Or done <> count Then Err.Raise 5, "CBinarySnapshotFile", "Ошибка записи файла: " & mPath
    End If
    mPosition = mPosition + count
    If mPosition > mSize Then mSize = mPosition
End Sub

' Перед чтением, перемещением и закрытием сохраняет все накопленные байты.
' Большие числовые массивы проходят напрямую и не копируются через этот буфер.
Private Sub FlushWrites()
    If mBuffered = 0 Then Exit Sub
    Dim done As Long
    If WriteFile(mHandle, VarPtr(mWriteBuffer(0)), mBuffered, done, 0) = 0 Or done <> mBuffered Then Err.Raise 5, "CBinarySnapshotFile", "Ошибка блочной записи файла: " & mPath
    mBuffered = 0
End Sub

' Записывает одно 32-битное целое без десятичного преобразования.
Public Sub WriteLong(ByVal value As Long)
    WriteExact VarPtr(value), 4
End Sub

' Читает одно 32-битное целое.
Public Function ReadLong() As Long
    Dim value As Long: ReadExact VarPtr(value), 4: ReadLong = value
End Function

' Записывает длину и UTF-16LE прямо из BSTR, без временного массива байтов.
Public Sub WriteSizedUTF16(ByVal value As String)
    Dim count As Long: count = LenB(value)
    WriteExact VarPtr(count), 4
    WriteExact StrPtr(value), count
End Sub

' Читает проверенную UTF-16LE строку прямо в отдельный BSTR нужной длины.
' Пустые строки, встроенный NUL и любые кодовые единицы Unicode сохраняются побайтно.
Public Function ReadSizedUTF16(ByVal maximumBytes As Long) As String
    Dim count As Long: count = ReadLong()
    If count < 0 Or count > maximumBytes Or count Mod 2 <> 0 Or count > mSize - mPosition Then Err.Raise 5, , "Неверная длина строки снимка."
    If count = 0 Then Exit Function
    ReadSizedUTF16 = String$(count \ 2, vbNullChar)
    ReadExact StrPtr(ReadSizedUTF16), count
End Function

' Записывает исходные восемь байтов числа IEEE 754.
Public Sub WriteDouble(ByVal value As Double)
    WriteExact VarPtr(value), 8
End Sub

' Читает исходное число IEEE 754 без округления.
Public Function ReadDouble() As Double
    Dim value As Double: ReadExact VarPtr(value), 8: ReadDouble = value
End Function

' Записывает код типа одной колонки или значения.
Public Sub WriteByte(ByVal value As Byte)
    WriteExact VarPtr(value), 1
End Sub

' Читает один код типа.
Public Function ReadByte() As Byte
    Dim value As Byte: ReadExact VarPtr(value), 1: ReadByte = value
End Function

' Записывает непрерывный массив байтов без служебного SAFEARRAY-дескриптора.
Public Sub WriteBytes(ByRef values() As Byte)
    WriteExact VarPtr(values(LBound(values))), UBound(values) - LBound(values) + 1
End Sub

' Получает плотный массив байтов с проверкой длины до выделения памяти.
Public Function ReadBytes(ByVal count As Long) As Variant
    If count < 1 Or count > mSize - mPosition Then Err.Raise 5, "CBinarySnapshotFile", "Неверная длина байтов снимка."
    Dim values() As Byte: ReDim values(0 To count - 1)
    ReadExact VarPtr(values(0)), count: ReadBytes = values
End Function

' Записывает числовую колонку одним вызовом Windows.
Public Sub WriteDoubles(ByRef values() As Double)
    WriteExact VarPtr(values(LBound(values))), (UBound(values) - LBound(values) + 1) * 8
End Sub

' Читает числовую колонку без цикла по файловым операциям.
Public Function ReadDoubles(ByVal count As Long) As Variant
    If count < 1 Or CDbl(count) * 8 > mSize - mPosition Then Err.Raise 5, "CBinarySnapshotFile", "Неверная длина числовой колонки."
    Dim values() As Double: ReDim values(1 To count)
    ReadExact VarPtr(values(1)), count * 8: ReadDoubles = values
End Function

' Записывает индексы строкового словаря плотным 32-битным массивом.
Public Sub WriteLongs(ByRef values() As Long)
    WriteExact VarPtr(values(LBound(values))), (UBound(values) - LBound(values) + 1) * 4
End Sub

' Читает индексы словаря с проверкой размера до выделения массива.
Public Function ReadLongs(ByVal count As Long) As Variant
    If count < 1 Or CDbl(count) * 4 > mSize - mPosition Then Err.Raise 5, "CBinarySnapshotFile", "Неверная длина индексов словаря."
    Dim values() As Long: ReDim values(1 To count)
    ReadExact VarPtr(values(1)), count * 4: ReadLongs = values
End Function

' Сохраняет только ASCII-сигнатуры и контрольную сумму фиксированной длины.
Public Sub WriteASCII(ByVal value As String)
    Dim values() As Byte: values = StrConv(value, vbFromUnicode): WriteBytes values
End Sub

' Читает ASCII-сигнатуру заданной длины.
Public Function ReadASCII(ByVal count As Long) As String
    Dim values() As Byte: values = ReadBytes(count): ReadASCII = StrConv(values, vbUnicode)
End Function
