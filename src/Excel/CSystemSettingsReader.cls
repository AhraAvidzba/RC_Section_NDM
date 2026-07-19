VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSystemSettingsReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mKeys() As String
Private mValues() As String
Private mDefaults() As String
Private mSources() As String
Private mCount As Long

Public Sub LoadFromWorkbook(ByVal workbook As Object)
    Dim settingsRange As Object
    Set settingsRange = workbook.Names.Item("rngSystemSettings").RefersToRange
    LoadFromRange settingsRange
End Sub

Public Sub LoadFromRange(ByVal settingsRange As Object)
    Clear

    Dim rowIndex As Long
    For rowIndex = 2 To settingsRange.Rows.Count
        Dim keyValue As String
        keyValue = Trim$(CStr(settingsRange.Cells.Item(rowIndex, 1).Value2))
        If Len(keyValue) > 0 And Left$(keyValue, 1) <> "[" Then
            mCount = mCount + 1
            ReDim Preserve mKeys(1 To mCount)
            ReDim Preserve mValues(1 To mCount)
            ReDim Preserve mDefaults(1 To mCount)
            ReDim Preserve mSources(1 To mCount)
            mKeys(mCount) = keyValue
            mValues(mCount) = Trim$(CStr(settingsRange.Cells.Item(rowIndex, 2).Value2))
            mDefaults(mCount) = Trim$(CStr(settingsRange.Cells.Item(rowIndex, 3).Value2))
            If settingsRange.Columns.Count >= 6 Then
                mSources(mCount) = Trim$(CStr(settingsRange.Cells.Item(rowIndex, 6).Value2))
            ElseIf settingsRange.Columns.Count >= 4 Then
                mSources(mCount) = Trim$(CStr(settingsRange.Cells.Item(rowIndex, 4).Value2))
            Else
                mSources(mCount) = vbNullString
            End If
        End If
    Next rowIndex
End Sub

Public Sub Clear()
    Erase mKeys
    Erase mValues
    Erase mDefaults
    Erase mSources
    mCount = 0
End Sub

Public Function GetString(ByVal key As String, Optional ByVal defaultValue As String = vbNullString) As String
    Dim index As Long
    index = FindKey(key)
    If index = 0 Then
        GetString = defaultValue
    ElseIf Len(mValues(index)) = 0 Or UCase$(mValues(index)) = "TODO" Then
        GetString = defaultValue
    Else
        GetString = mValues(index)
    End If
End Function

Public Function GetDouble(ByVal key As String, Optional ByVal defaultValue As Double = 0#) As Double
    Dim textValue As String
    textValue = GetString(key, vbNullString)
    If Len(textValue) = 0 Then
        GetDouble = defaultValue
    Else
        GetDouble = CDbl(Replace$(textValue, ".", ApplicationDecimalSeparator()))
    End If
End Function

Public Function GetLong(ByVal key As String, Optional ByVal defaultValue As Long = 0) As Long
    GetLong = CLng(GetDouble(key, CDbl(defaultValue)))
End Function

Public Function GetBoolean(ByVal key As String, Optional ByVal defaultValue As Boolean = False) As Boolean
    Dim textValue As String
    textValue = UCase$(GetString(key, IIf(defaultValue, "YES", "NO")))
    GetBoolean = (textValue = "YES" Or textValue = "TRUE" Or textValue = "1" Or textValue = "ДА")
End Function

Public Function GetSource(ByVal key As String) As String
    Dim index As Long
    index = FindKey(key)
    If index = 0 Then
        GetSource = vbNullString
    Else
        GetSource = mSources(index)
    End If
End Function

Public Function HasKey(ByVal key As String) As Boolean
    HasKey = (FindKey(key) > 0)
End Function

Public Function KeyCount() As Long
    KeyCount = mCount
End Function

Public Function KeyAt(ByVal index As Long) As String
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 4300, "CSystemSettingsReader", "Setting index is out of range."
    KeyAt = mKeys(index)
End Function

Public Function DuplicateCount() As Long
    Dim duplicates As Long
    Dim i As Long
    Dim j As Long
    For i = 1 To mCount
        For j = i + 1 To mCount
            If StrComp(mKeys(i), mKeys(j), vbTextCompare) = 0 Then
                duplicates = duplicates + 1
                Exit For
            End If
        Next j
    Next i
    DuplicateCount = duplicates
End Function

Private Function FindKey(ByVal key As String) As Long
    Dim i As Long
    For i = 1 To mCount
        If StrComp(mKeys(i), key, vbTextCompare) = 0 Then
            FindKey = i
            Exit Function
        End If
    Next i
    FindKey = 0
End Function

Private Function ApplicationDecimalSeparator() As String
    ApplicationDecimalSeparator = Mid$(CStr(1.1), 2, 1)
End Function





