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
' Контрольная сумма файла штатными средствами Windows
' ==========================================================================
' Читает файл блоками; не требует программ, не меняет содержимое.
Private Declare PtrSafe Function CryptAcquireContextW Lib "advapi32" (ByRef provider As LongPtr, ByVal container As LongPtr, ByVal providerName As LongPtr, ByVal providerType As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptCreateHash Lib "advapi32" (ByVal provider As LongPtr, ByVal algorithm As Long, ByVal key As LongPtr, ByVal flags As Long, ByRef hash As LongPtr) As Long
Private Declare PtrSafe Function CryptHashData Lib "advapi32" (ByVal hash As LongPtr, ByRef data As Byte, ByVal length As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32" (ByVal hash As LongPtr, ByVal parameter As Long, ByRef data As Byte, ByRef length As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32" (ByVal hash As LongPtr) As Long
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32" (ByVal provider As LongPtr, ByVal flags As Long) As Long

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
