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

' ==========================================================================
' Краткая сводка пакетного расчета на лист Results
' ==========================================================================
' Writer выводит верхний блок batch summary: определяющее сочетание, время
' расчета и компактную таблицу статусов по всем LC. Данные уже рассчитаны в
' CBatchSectionCalculator; здесь нет повторного solve и нет изменения выбора
' худшего сочетания.

Private Const TARGET_ROWS As Long = 29
Private Const TARGET_COLUMNS As Long = 32

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

    PutHeaders data, 6, units

    Dim i As Long
    For i = 1 To batch.Count
        PutCombination data, 6 + i, batch, i, units
    Next i

    target.Value2 = data
End Sub

Private Function OutputLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputLength = value Else OutputLength = units.InternalLengthToOutput(value)
End Function

Private Function OutputLengthUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputLengthUnit = "mm" Else OutputLengthUnit = units.OutputLengthUnit
End Function

Private Function OutputArea(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputArea = value Else OutputArea = units.InternalAreaToOutput(value)
End Function

Private Function OutputAreaUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputAreaUnit = "mm2" Else OutputAreaUnit = units.OutputAreaUnit
End Function

Private Function OutputStress(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputStress = value Else OutputStress = units.InternalStressToOutput(value)
End Function

Private Function OutputStressUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputStressUnit = "MPa" Else OutputStressUnit = units.OutputStressUnit
End Function

Private Function OutputForce(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputForce = value Else OutputForce = units.InternalForceToOutput(value)
End Function

Private Function OutputForceUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputForceUnit = "N" Else OutputForceUnit = units.OutputForceUnit
End Function

' Очищает накопленное состояние перед новым расчетом или повторным формированием вывода.
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

Private Sub PutHeaders(ByRef data() As Variant, ByVal rowIndex As Long, ByVal units As CUnitSystem)
    data(rowIndex, 1) = "CombinationID"
    data(rowIndex, 2) = "Наименование"
    data(rowIndex, 3) = "CalculationType"
    data(rowIndex, 4) = "Status"

    data(rowIndex, 5) = "CapacityStatus"
    data(rowIndex, 6) = "StrengthCheckStatus"
    data(rowIndex, 7) = "CapacityLimitState"
    data(rowIndex, 8) = "LambdaCapacity"
    data(rowIndex, 9) = "CapacityUtilization"
    data(rowIndex, 10) = "StrainSafetyFactor"

    data(rowIndex, 11) = "CrackStatus"
    data(rowIndex, 12) = "CrackFormed"
    data(rowIndex, 13) = "CrackWidth, " & OutputLengthUnit(units)
    data(rowIndex, 14) = "CrackAllowable, " & OutputLengthUnit(units)
    data(rowIndex, 15) = "CrackUtilization"
    data(rowIndex, 16) = "phi1"
    data(rowIndex, 17) = "phi2"
    data(rowIndex, 18) = "phi3"
    data(rowIndex, 19) = "psi_s"
    data(rowIndex, 20) = "sigma_s, " & OutputStressUnit(units)
    data(rowIndex, 21) = "sigma_s_crc, " & OutputStressUnit(units)
    data(rowIndex, 22) = "As, " & OutputAreaUnit(units)
    data(rowIndex, 23) = "Abt, " & OutputAreaUnit(units)
    data(rowIndex, 24) = "ds_eq, " & OutputLengthUnit(units)
    data(rowIndex, 25) = "ls_raw, " & OutputLengthUnit(units)
    data(rowIndex, 26) = "ls, " & OutputLengthUnit(units)
    data(rowIndex, 27) = "lambda_crc"
    data(rowIndex, 28) = "Ncrc, " & OutputForceUnit(units)
    data(rowIndex, 29) = "TensionRebarCount"
    data(rowIndex, 30) = "EffectiveZoneDepth, " & OutputLengthUnit(units)
    data(rowIndex, 31) = "WorstUtilization"
    data(rowIndex, 32) = "Определяющее"
End Sub

Private Sub PutCombination(ByRef data() As Variant, ByVal rowIndex As Long, ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal units As CUnitSystem)
    If rowIndex > UBound(data, 1) Then Exit Sub

    data(rowIndex, 1) = batch.CombinationID(index)
    data(rowIndex, 2) = batch.CombinationName(index)
    data(rowIndex, 3) = batch.CalculationType(index)
    data(rowIndex, 4) = batch.Status(index)

    data(rowIndex, 5) = batch.CapacityStatus(index)
    data(rowIndex, 6) = batch.StrengthCheckStatus(index)
    data(rowIndex, 7) = batch.CapacityLimitState(index)
    data(rowIndex, 8) = batch.LambdaCapacity(index)
    data(rowIndex, 9) = batch.UtilizationCapacity(index)
    data(rowIndex, 10) = batch.StrainSafetyFactor(index)

    data(rowIndex, 11) = batch.CrackStatus(index)
    If StrComp(batch.CrackStatus(index), "OK", vbTextCompare) = 0 Then
        data(rowIndex, 12) = IIf(batch.CrackFormed(index), "Да", "Нет")
        data(rowIndex, 13) = OutputLength(batch.CrackWidth(index), units)
        data(rowIndex, 14) = OutputLength(batch.CrackAllowable(index), units)
        data(rowIndex, 15) = batch.CrackUtilization(index)
        data(rowIndex, 16) = batch.CrackPhi1(index)
        data(rowIndex, 17) = batch.CrackPhi2(index)
        data(rowIndex, 18) = batch.CrackPhi3(index)
        data(rowIndex, 19) = batch.CrackPsiS(index)
        data(rowIndex, 20) = OutputStress(batch.CrackSigmaS(index), units)
        data(rowIndex, 21) = OutputStress(batch.CrackSigmaSCrc(index), units)
        data(rowIndex, 22) = OutputArea(batch.CrackAs(index), units)
        data(rowIndex, 23) = OutputArea(batch.CrackAbt(index), units)
        data(rowIndex, 24) = OutputLength(batch.CrackDsEquivalent(index), units)
        data(rowIndex, 25) = OutputLength(batch.CrackSpacingRaw(index), units)
        data(rowIndex, 26) = OutputLength(batch.CrackSpacing(index), units)
        data(rowIndex, 27) = batch.CrackLambdaCrc(index)
        data(rowIndex, 28) = OutputForce(batch.CrackNcrc(index), units)
        data(rowIndex, 29) = batch.CrackTensionRebarCount(index)
        data(rowIndex, 30) = OutputLength(batch.CrackEffectiveZoneDepth(index), units)
    End If
    data(rowIndex, 31) = batch.GoverningUtilization(index)
    data(rowIndex, 32) = IIf(index = batch.GoverningCombinationIndex, "Да", "Нет")
End Sub

