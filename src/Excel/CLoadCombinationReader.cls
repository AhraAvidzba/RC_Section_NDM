VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLoadCombinationReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub LoadFromWorkbook(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3950, "CLoadCombinationReader", "Workbook is missing."
    If batch Is Nothing Then Err.Raise vbObjectError + 3951, "CLoadCombinationReader", "Batch calculator is missing."
    LoadFromRange workbook.Names.Item("rngLoadCombinations").RefersToRange, batch
End Sub

Public Sub LoadFromRange(ByVal source As Object, ByVal batch As CBatchSectionCalculator)
    If source Is Nothing Then Err.Raise vbObjectError + 3952, "CLoadCombinationReader", "Load combination range is missing."
    If batch Is Nothing Then Err.Raise vbObjectError + 3953, "CLoadCombinationReader", "Batch calculator is missing."

    Dim rowIndex As Long
    For rowIndex = 2 To source.Rows.Count
        If IsEmptyCombinationRow(source, rowIndex) Then
            ' Keep blank rows available for user input.
        Else
            AddRow source, batch, rowIndex
        End If
    Next rowIndex
End Sub

Private Sub AddRow(ByVal source As Object, ByVal batch As CBatchSectionCalculator, ByVal rowIndex As Long)
    Dim combinationID As String
    Dim calculationType As String
    Dim durationType As String
    Dim comment As String
    combinationID = Trim$(CStr(source.Cells.Item(rowIndex, 1).Value2))
    calculationType = Trim$(CStr(source.Cells.Item(rowIndex, 5).Value2))
    durationType = Trim$(CStr(source.Cells.Item(rowIndex, 6).Value2))
    comment = Trim$(CStr(source.Cells.Item(rowIndex, 7).Value2))

    Dim errorText As String
    Dim nValue As Double
    Dim mxValue As Double
    Dim myValue As Double
    nValue = ReadRequiredDouble(source.Cells.Item(rowIndex, 2).Value2, "N", errorText)
    mxValue = ReadRequiredDouble(source.Cells.Item(rowIndex, 3).Value2, "Mx", errorText)
    myValue = ReadRequiredDouble(source.Cells.Item(rowIndex, 4).Value2, "My", errorText)

    If Len(Trim$(combinationID)) = 0 Then AppendError errorText, "CombinationID is empty"

    If Len(errorText) > 0 Then
        batch.AddInvalidCombination combinationID, calculationType, durationType, comment, errorText
    Else
        batch.AddCombination combinationID, nValue, mxValue, myValue, calculationType, durationType, comment
    End If
End Sub

Private Function IsEmptyCombinationRow(ByVal source As Object, ByVal rowIndex As Long) As Boolean
    Dim col As Long
    For col = 1 To 7
        If Len(Trim$(CStr(source.Cells.Item(rowIndex, col).Value2))) > 0 Then Exit Function
    Next col
    IsEmptyCombinationRow = True
End Function

Private Function ReadRequiredDouble(ByVal value As Variant, ByVal fieldName As String, ByRef errorText As String) As Double
    If Len(Trim$(CStr(value))) = 0 Then
        AppendError errorText, fieldName & " is empty"
    ElseIf Not IsNumeric(value) Then
        AppendError errorText, fieldName & " is not numeric"
    Else
        ReadRequiredDouble = CDbl(value)
    End If
End Function

Private Sub AppendError(ByRef errorText As String, ByVal item As String)
    If Len(errorText) > 0 Then errorText = errorText & "; "
    errorText = errorText & item
End Sub

