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

Private Const TARGET_ROWS As Long = 29
Private Const TARGET_COLUMNS As Long = 14

Public Sub WriteSummary(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator, Optional ByVal units As CUnitSystem = Nothing)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3970, "CBatchResultWriter", "Книга Excel не передана."
    If batch Is Nothing Then Err.Raise vbObjectError + 3971, "CBatchResultWriter", "Пакетный расчетчик не передан."

    Dim target As Object
    Set target = SummaryTargetRange(workbook)

    Dim data() As Variant
    ReDim data(1 To TARGET_ROWS, 1 To TARGET_COLUMNS)

    data(1, 1) = "Сводка пакетного расчета"
    data(2, 1) = "Определяющее сочетание"
    data(2, 2) = batch.GoverningCombinationID
    data(3, 1) = "Наименование определяющего сочетания"
    data(3, 2) = batch.GoverningCombinationName
    data(4, 1) = "Время расчета, с"
    data(4, 2) = batch.ElapsedSeconds
    data(5, 1) = "Точка приложения нагрузки"
    data(5, 2) = "X=" & CStr(OutputLength(batch.LoadReferenceX, units)) & " " & OutputLengthUnit(units)
    data(5, 3) = "Y=" & CStr(OutputLength(batch.LoadReferenceY, units)) & " " & OutputLengthUnit(units)

    PutHeaders data, 6

    Dim i As Long
    For i = 1 To batch.Count
        PutCombination data, 6 + i, batch, i
    Next i

    target.Value2 = data
End Sub

Private Function OutputLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputLength = value Else OutputLength = units.InternalLengthToOutput(value)
End Function

Private Function OutputLengthUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputLengthUnit = "mm" Else OutputLengthUnit = units.OutputLengthUnit
End Function

Public Sub ClearSummary(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3972, "CBatchResultWriter", "Книга Excel не передана."
    SummaryTargetRange(workbook).ClearContents
End Sub

Private Function SummaryTargetRange(ByVal workbook As Object) As Object
    Dim target As Object
    Set target = workbook.Names.Item("rngBatchSummary").RefersToRange
    If target.Rows.Count < TARGET_ROWS Or target.Columns.Count < TARGET_COLUMNS Then
        Err.Raise vbObjectError + 3973, "CBatchResultWriter", "rngBatchSummary меньше требуемого размера сводки."
    End If
    Set SummaryTargetRange = target.Resize(TARGET_ROWS, TARGET_COLUMNS)
End Function

Private Sub PutHeaders(ByRef data() As Variant, ByVal rowIndex As Long)
    data(rowIndex, 1) = "CombinationID"
    data(rowIndex, 2) = "Наименование"
    data(rowIndex, 3) = "Status"
    data(rowIndex, 4) = "CapacityStatus"
    data(rowIndex, 5) = "CapacityLimitState"
    data(rowIndex, 6) = "StrengthCheckStatus"
    data(rowIndex, 7) = "CrackStatus"
    data(rowIndex, 8) = "LambdaCapacity"
    data(rowIndex, 9) = "CapacityUtilization"
    data(rowIndex, 10) = "StrainSafetyFactor"
    data(rowIndex, 11) = "CrackWidth"
    data(rowIndex, 12) = "CrackUtilization"
    data(rowIndex, 13) = "WorstUtilization"
    data(rowIndex, 14) = "Определяющее"
End Sub

Private Sub PutCombination(ByRef data() As Variant, ByVal rowIndex As Long, ByVal batch As CBatchSectionCalculator, ByVal index As Long)
    If rowIndex > UBound(data, 1) Then Exit Sub

    data(rowIndex, 1) = batch.CombinationID(index)
    data(rowIndex, 2) = batch.CombinationName(index)
    data(rowIndex, 3) = batch.Status(index)
    data(rowIndex, 4) = batch.CapacityStatus(index)
    data(rowIndex, 5) = batch.CapacityLimitState(index)
    data(rowIndex, 6) = batch.StrengthCheckStatus(index)
    data(rowIndex, 7) = batch.CrackStatus(index)
    data(rowIndex, 8) = batch.LambdaCapacity(index)
    data(rowIndex, 9) = batch.UtilizationCapacity(index)
    data(rowIndex, 10) = batch.StrainSafetyFactor(index)
    data(rowIndex, 11) = batch.CrackWidth(index)
    data(rowIndex, 12) = batch.CrackUtilization(index)
    data(rowIndex, 13) = batch.GoverningUtilization(index)
    data(rowIndex, 14) = IIf(index = batch.GoverningCombinationIndex, "Да", "Нет")
End Sub
