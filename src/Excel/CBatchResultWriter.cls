VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBatchResultWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const START_ROW As Long = 92
Private Const START_COLUMN As Long = 1

Public Sub WriteSummary(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3970, "CBatchResultWriter", "Workbook is missing."
    If batch Is Nothing Then Err.Raise vbObjectError + 3971, "CBatchResultWriter", "Batch calculator is missing."

    Dim sheet As Object
    Set sheet = workbook.Worksheets.Item("System")
    Dim target As Object
    Set target = sheet.Range(sheet.Cells.Item(START_ROW, START_COLUMN), sheet.Cells.Item(START_ROW + 28, START_COLUMN + 10))
    target.ClearContents

    sheet.Cells.Item(START_ROW, START_COLUMN).Value2 = "Batch summary"
    sheet.Cells.Item(START_ROW + 1, START_COLUMN).Value2 = "GoverningCombinationID"
    sheet.Cells.Item(START_ROW + 1, START_COLUMN + 1).Value2 = batch.GoverningCombinationID
    sheet.Cells.Item(START_ROW + 2, START_COLUMN).Value2 = "GoverningCombinationName"
    sheet.Cells.Item(START_ROW + 2, START_COLUMN + 1).Value2 = batch.GoverningCombinationName
    sheet.Cells.Item(START_ROW + 3, START_COLUMN).Value2 = "ElapsedSeconds"
    sheet.Cells.Item(START_ROW + 3, START_COLUMN + 1).Value2 = batch.ElapsedSeconds

    PutHeaders sheet, START_ROW + 5

    Dim i As Long
    For i = 1 To batch.Count
        PutCombination sheet, START_ROW + 5 + i, batch, i
    Next i

    Dim diagRow As Long
    diagRow = START_ROW + 7 + batch.Count
    sheet.Cells.Item(diagRow, START_COLUMN).Value2 = "Batch diagnostics"
    sheet.Cells.Item(diagRow + 1, START_COLUMN).Value2 = batch.DiagnosticLog
End Sub

Private Sub PutHeaders(ByVal sheet As Object, ByVal rowIndex As Long)
    sheet.Cells.Item(rowIndex, START_COLUMN).Value2 = "CombinationID"
    sheet.Cells.Item(rowIndex, START_COLUMN + 1).Value2 = "Name"
    sheet.Cells.Item(rowIndex, START_COLUMN + 2).Value2 = "Status"
    sheet.Cells.Item(rowIndex, START_COLUMN + 3).Value2 = "CapacityStatus"
    sheet.Cells.Item(rowIndex, START_COLUMN + 4).Value2 = "CrackStatus"
    sheet.Cells.Item(rowIndex, START_COLUMN + 5).Value2 = "LambdaCapacity"
    sheet.Cells.Item(rowIndex, START_COLUMN + 6).Value2 = "CapacityUtilization"
    sheet.Cells.Item(rowIndex, START_COLUMN + 7).Value2 = "CrackWidth"
    sheet.Cells.Item(rowIndex, START_COLUMN + 8).Value2 = "CrackUtilization"
    sheet.Cells.Item(rowIndex, START_COLUMN + 9).Value2 = "WorstUtilization"
    sheet.Cells.Item(rowIndex, START_COLUMN + 10).Value2 = "Governing"
End Sub

Private Sub PutCombination(ByVal sheet As Object, ByVal rowIndex As Long, ByVal batch As CBatchSectionCalculator, ByVal index As Long)
    sheet.Cells.Item(rowIndex, START_COLUMN).Value2 = batch.CombinationID(index)
    sheet.Cells.Item(rowIndex, START_COLUMN + 1).Value2 = batch.CombinationName(index)
    sheet.Cells.Item(rowIndex, START_COLUMN + 2).Value2 = batch.Status(index)
    sheet.Cells.Item(rowIndex, START_COLUMN + 3).Value2 = batch.CapacityStatus(index)
    sheet.Cells.Item(rowIndex, START_COLUMN + 4).Value2 = batch.CrackStatus(index)
    sheet.Cells.Item(rowIndex, START_COLUMN + 5).Value2 = batch.LambdaCapacity(index)
    sheet.Cells.Item(rowIndex, START_COLUMN + 6).Value2 = batch.UtilizationCapacity(index)
    sheet.Cells.Item(rowIndex, START_COLUMN + 7).Value2 = batch.CrackWidth(index)
    sheet.Cells.Item(rowIndex, START_COLUMN + 8).Value2 = batch.CrackUtilization(index)
    sheet.Cells.Item(rowIndex, START_COLUMN + 9).Value2 = batch.GoverningUtilization(index)
    sheet.Cells.Item(rowIndex, START_COLUMN + 10).Value2 = IIf(index = batch.GoverningCombinationIndex, "Yes", "No")
End Sub

