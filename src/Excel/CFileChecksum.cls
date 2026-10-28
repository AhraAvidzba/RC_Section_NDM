VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CFileChecksum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Контрольные суммы файла и числовой геометрической записи
' ==========================================================================
' Читает файл блоками; не требует программ, не меняет содержимое.
Private Declare PtrSafe Function CryptAcquireContextW Lib "advapi32" (ByRef provider As LongPtr, ByVal container As LongPtr, ByVal providerName As LongPtr, ByVal providerType As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptCreateHash Lib "advapi32" (ByVal provider As LongPtr, ByVal algorithm As Long, ByVal key As LongPtr, ByVal flags As Long, ByRef hash As LongPtr) As Long
Private Declare PtrSafe Function CryptHashData Lib "advapi32" (ByVal hash As LongPtr, ByRef data As Byte, ByVal length As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32" (ByVal hash As LongPtr, ByVal parameter As Long, ByRef data As Byte, ByRef length As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32" (ByVal hash As LongPtr) As Long
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32" (ByVal provider As LongPtr, ByVal flags As Long) As Long

' Две независимые числовые свёртки; каждый промежуточный целый меньше 2^49 и точно представим Double.
Private Const GEOMETRY_HASH_PRIME_A As Double = 4294967291#
Private Const GEOMETRY_HASH_PRIME_B As Double = 4294967279#
Private Type TGeometryHashNumber
    Value As Double
End Type
Private Type TGeometryHashBits
    Low As Long
    High As Long
End Type
Private mGeometrySeedA As Double, mGeometrySeedB As Double ' Подпись общей шапки одной неизменяемой пары матриц.
Private mGeometrySeedReady As Boolean ' Разрешает PreparedGeometryHash только после явной подготовки текущих заголовков.
Private mGeometryProvider As LongPtr ' Один штатный провайдер Windows на операцию, без открытия для каждого элемента.
Private Declare PtrSafe Function HashGeometryMemory Lib "advapi32" Alias "CryptHashData" (ByVal hash As LongPtr, ByVal data As LongPtr, ByVal count As Long, ByVal flags As Long) As Long

' SHA256 средствами Windows; блоки ограничены 64 КиБ, внешние программы не запускаются.
Public Function FileSHA256(ByVal path As String, Optional ByVal byteCount As Long = 0) As String
    Dim provider As LongPtr, hash As LongPtr
    Dim bytes() As Byte, digest(0 To 31) As Byte, count As Long, remaining As Long, i As Long, result As String
    Dim file As CBinarySnapshotFile: Set file = New CBinarySnapshotFile
    On Error GoTo Failed
    If CryptAcquireContextW(provider, 0, 0, 24, &HF0000000) = 0 Then Err.Raise 5, , "Не удалось открыть SHA256 Windows."
    If CryptCreateHash(provider, &H800C&, 0, 0, hash) = 0 Then Err.Raise 5, , "Не удалось создать SHA256 Windows."
    file.OpenRead path: remaining = file.Size
    If byteCount > 0 Then
        If byteCount > remaining Then Err.Raise 5, , "Размер контрольной суммы больше файла."
        remaining = byteCount
    End If
    Do While remaining > 0
        count = remaining: If count > 65536 Then count = 65536
        bytes = file.ReadBytes(count)
        If CryptHashData(hash, bytes(0), count, 0) = 0 Then Err.Raise 5, , "Ошибка вычисления SHA256."
        remaining = remaining - count
    Loop
    file.CloseFile: count = 32
    If CryptGetHashParam(hash, 2, digest(0), count, 0) = 0 Or count <> 32 Then Err.Raise 5, , "Не удалось прочитать SHA256."
    For i = 0 To 31: result = result & Right$("0" & Hex$(digest(i)), 2): Next i
    CryptDestroyHash hash: CryptReleaseContext provider, 0
    FileSHA256 = result
    Exit Function
Failed:
    Dim reason As String: reason = Err.Description
    file.CloseFile
    If hash <> 0 Then CryptDestroyHash hash
    If provider <> 0 Then CryptReleaseContext provider, 0
    Err.Raise vbObjectError + 4943, "CFileChecksum.FileSHA256", reason
End Function

' Вычисляет подпись полной числовой записи без геометрических операций и округления.
' RunID и необязательные подписи исключены; роли, ID, единицы и все A/c/I включены.
Public Function GeometryHash(ByRef geometry As Variant, ByVal geometryRow As Long, _
        ByRef boundaries As Variant, ByVal firstRow As Long, ByVal lastRow As Long) As String
    PrepareGeometryHashHeaders geometry, boundaries
    GeometryHash = PreparedGeometryHash(geometry, geometryRow, boundaries, firstRow, lastRow)
End Function

' Один раз сворачивает общие заголовки в пределах чтения либо записи неизменяемых матриц.
' Изменение шапки требует нового вызова; обычный GeometryHash всегда подготавливает её сам.
Friend Sub PrepareGeometryHashHeaders(ByRef geometry As Variant, ByRef boundaries As Variant)
    mGeometrySeedA = 2166136261#: mGeometrySeedB = 1315423911#
    GeometryMixValue mGeometrySeedA, mGeometrySeedB, "NDM.ConcreteGeometry"
    Dim column As Long
    For column = 2 To 14: GeometryMixValue mGeometrySeedA, mGeometrySeedB, geometry(1, column): Next column
    For column = 2 To 18
        If column <> 15 And column <> 16 Then GeometryMixValue mGeometrySeedA, mGeometrySeedB, boundaries(1, column)
    Next column
    mGeometrySeedReady = True
End Sub

' Использует явно подготовленную общую шапку, обходя только собственные числа и сегменты элемента.
' Это короткий контекст одной операции; между изменяемыми снимками экземпляр не переиспользуется.
Friend Function PreparedGeometryHash(ByRef geometry As Variant, ByVal geometryRow As Long, _
        ByRef boundaries As Variant, ByVal firstRow As Long, ByVal lastRow As Long) As String
    If Not mGeometrySeedReady Then Err.Raise 5, "CFileChecksum.GeometryHash", "Не подготовлена шапка геометрической подписи."
    Dim values() As Double, index As Long, row As Long, column As Long, value As Variant
    ReDim values(1 To 3 + 2 * (13 + 15 * (lastRow - firstRow + 1)))
    values(1) = mGeometrySeedA: values(2) = mGeometrySeedB: values(3) = lastRow - firstRow + 1: index = 3
    Dim text As String, payload As String, pieces As Collection
    For column = 2 To 14
        value = geometry(geometryRow, column): index = index + 2
        Select Case VarType(value)
            Case vbEmpty
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, 14
                values(index - 1) = 1#: values(index) = CDbl(value)
            Case vbString
                text = CStr(value)
                If Len(text) > 0 Then
                    values(index - 1) = 2#: values(index) = Len(text)
                    If Len(payload) + Len(text) > 32768 Then
                        If pieces Is Nothing Then Set pieces = New Collection
                        pieces.Add payload: payload = text
                    Else
                        payload = payload & text
                    End If
                End If
            Case Else: Err.Raise 5, "CFileChecksum.GeometryHash", "Неподдерживаемое значение геометрической записи."
        End Select
    Next column
    For row = firstRow To lastRow
        For column = 2 To 18
            If column <> 15 And column <> 16 Then
                value = boundaries(row, column): index = index + 2
                Select Case VarType(value)
                    Case vbEmpty
                    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, 14
                        values(index - 1) = 1#: values(index) = CDbl(value)
                    Case vbString
                        text = CStr(value)
                        If Len(text) > 0 Then
                            values(index - 1) = 2#: values(index) = Len(text)
                            If Len(payload) + Len(text) > 32768 Then
                                If pieces Is Nothing Then Set pieces = New Collection
                                pieces.Add payload: payload = text
                            Else
                                payload = payload & text
                            End If
                        End If
                    Case Else: Err.Raise 5, "CFileChecksum.GeometryHash", "Неподдерживаемое значение геометрической записи."
                End Select
            End If
        Next column
    Next row
    PreparedGeometryHash = NativeGeometryHash(values, payload, pieces)
End Function

' Передаёт весь массив чисел одним вызовом Windows; строки хешируются отдельными крупными блоками.
' Теги типов и длины в числовом массиве исключают неоднозначное объединение строк и чисел; берутся первые 64 бита SHA256.
Private Function NativeGeometryHash(ByRef values() As Double, ByVal text As String, ByVal pieces As Collection) As String
    Dim hash As LongPtr, digest(0 To 31) As Byte, count As Long, i As Long, result As String, piece As Variant, block As String
    On Error GoTo Failed
    If mGeometryProvider = 0 Then
        If CryptAcquireContextW(mGeometryProvider, 0, 0, 24, &HF0000000) = 0 Then Err.Raise 5, , "Не удалось открыть контрольную сумму Windows."
    End If
    If CryptCreateHash(mGeometryProvider, &H800C&, 0, 0, hash) = 0 Then Err.Raise 5, , "Не удалось создать контрольную сумму геометрии."
    If HashGeometryMemory(hash, VarPtr(values(1)), UBound(values) * 8, 0) = 0 Then Err.Raise 5, , "Не удалось проверить числа геометрии."
    If Not pieces Is Nothing Then
        For Each piece In pieces
            block = CStr(piece)
            If LenB(block) > 0 Then
                If HashGeometryMemory(hash, StrPtr(block), LenB(block), 0) = 0 Then Err.Raise 5, , "Не удалось проверить обозначения геометрии."
            End If
        Next piece
    End If
    If LenB(text) > 0 Then
        If HashGeometryMemory(hash, StrPtr(text), LenB(text), 0) = 0 Then Err.Raise 5, , "Не удалось проверить обозначения геометрии."
    End If
    count = 32
    If CryptGetHashParam(hash, 2, digest(0), count, 0) = 0 Or count <> 32 Then Err.Raise 5, , "Не удалось прочитать контрольную сумму геометрии."
    For i = 0 To 7: result = result & Right$("0" & Hex$(digest(i)), 2): Next i
    CryptDestroyHash hash
    NativeGeometryHash = result
    Exit Function
Failed:
    Dim reason As String: reason = Err.Description
    If hash <> 0 Then CryptDestroyHash hash
    Err.Raise vbObjectError + 4943, "CFileChecksum.GeometryHash", reason
End Function

' Освобождает единственный провайдер операции; файловый SHA256 использует собственные локальные дескрипторы.
Private Sub Class_Terminate()
    If mGeometryProvider <> 0 Then CryptReleaseContext mGeometryProvider, 0
End Sub

' Сворачивает точные биты чисел либо UTF-16 строк с длиной и меткой типа; числовая строка отличается от Double.
' Empty и пустая строка означают одинаковое незаполненное поле; дату в физических данных принимать нельзя.
Private Sub GeometryMixValue(ByRef a As Double, ByRef b As Double, ByVal value As Variant)
    Dim number As TGeometryHashNumber, bits As TGeometryHashBits, word As Double
    Dim text As String, i As Long, length As Long
    Select Case VarType(value)
        Case vbEmpty: GeometryMixWord a, b, 0#
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, 14
            GeometryMixWord a, b, 1#
            number.Value = CDbl(value): LSet bits = number
            word = bits.Low: If word < 0# Then word = word + 4294967296#
            GeometryMixWord a, b, word
            word = bits.High: If word < 0# Then word = word + 4294967296#
            GeometryMixWord a, b, word
        Case vbString
            text = CStr(value): length = Len(text)
            If length = 0 Then GeometryMixWord a, b, 0#: Exit Sub
            GeometryMixWord a, b, 2#: GeometryMixWord a, b, CDbl(length)
            For i = 1 To length Step 2
                word = AscW(Mid$(text, i, 1)) And &HFFFF&
                If i < length Then word = word + CDbl(AscW(Mid$(text, i + 1, 1)) And &HFFFF&) * 65536#
                GeometryMixWord a, b, word
            Next i
        Case Else: Err.Raise 5, "CFileChecksum.GeometryHash", "Неподдерживаемое значение геометрической записи."
    End Select
End Sub

' Смешивает одно 32-битное слово двумя полиномами по разным простым модулям.
' Все операции целочисленно точны в Double; это контроль случайных изменений, а не криптографическая защита.
Private Sub GeometryMixWord(ByRef a As Double, ByRef b As Double, ByVal word As Double)
    a = a * 33# + word: a = a - Fix(a / GEOMETRY_HASH_PRIME_A) * GEOMETRY_HASH_PRIME_A
    b = b * 65599# + word: b = b - Fix(b / GEOMETRY_HASH_PRIME_B) * GEOMETRY_HASH_PRIME_B
End Sub

' Записывает беззнаковое 32-битное целое восемью шестнадцатеричными знаками без локали.
Private Function GeometryHashHex(ByVal value As Double) As String
    If value > 2147483647# Then value = value - 4294967296#
    GeometryHashHex = Right$("00000000" & Hex$(CLng(value)), 8)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Даёт независимый короткий вектор той же числовой свёртки для сравнения с Python.
Friend Function TestGeometryHashText(ByVal value As String) As String
    Dim a As Double, b As Double: a = 2166136261#: b = 1315423911#
    GeometryMixValue a, b, value
    TestGeometryHashText = GeometryHashHex(a) & GeometryHashHex(b)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет побайтовую передачу Double в штатную свёртку независимо от матриц и сериализатора.
Friend Function TestGeometryHashNumber(ByVal value As Double) As String
    Dim values(1 To 1) As Double, pieces As Collection: values(1) = value
    TestGeometryHashNumber = NativeGeometryHash(values, vbNullString, pieces)
End Function
