VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCapacityResultWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub WriteCapacityResult(ByVal workbook As Object, ByVal capacity As CCapacitySolver)
    WriteResult workbook, "rngResultSection", capacity
End Sub

Public Sub WriteDirectSectionResult(ByVal workbook As Object, ByVal solver As CSectionSolver)
    WriteDirectResult workbook, "rngResultSection", solver
End Sub

Public Sub WriteLinearSectionResult(ByVal workbook As Object, ByVal solver As CLinearSectionSolver)
    WriteLinearResult workbook, "rngResultSection", solver
End Sub
Private Sub WriteResult(ByVal workbook As Object, ByVal rangeName As String, ByVal capacity As CCapacitySolver)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3600, "CCapacityResultWriter", "Workbook is missing."
    If capacity Is Nothing Then Err.Raise vbObjectError + 3601, "CCapacityResultWriter", "Capacity result is missing."

    Dim target As Object
    Set target = workbook.Names.Item(rangeName).RefersToRange
    ClearResultRange target

    PutRow target, 1, "Indicator", "Value", "Unit", "Comment"
    PutRow target, 2, "Mode", "N+Mx+My", "", "Единый расчет по направлению вектора моментов"
    PutRow target, 3, "Converged", IIf(capacity.Converged, "Yes", "No"), "", capacity.StopReason
    PutRow target, 4, "LimitState", capacity.LimitState, "", "Физические и численные причины остановки разделяются"
    PutRow target, 5, "LambdaUltimate", capacity.LambdaUltimate, "", "SafetyFactor = LambdaUltimate"
    PutRow target, 6, "MxUltimate", capacity.MxUltimate / 1000000#, "kN*m", "Внутренние единицы: N*mm"
    PutRow target, 7, "MyUltimate", capacity.MyUltimate / 1000000#, "kN*m", "Внутренние единицы: N*mm"
    PutRow target, 8, "MomentResultant", capacity.MomentUltimate / 1000000#, "kN*m", "Результирующий момент вектора Mx/My"
    PutRow target, 9, "SafetyFactor", capacity.SafetyFactor, "", "Utilization=" & CStr(capacity.Utilization)
    PutRow target, 10, "MomentResidual", capacity.MomentEquilibriumResidual, "N*mm", "Независимая проверка равновесия"
    If Not capacity.LastSolver Is Nothing Then
        PutRow target, 11, "Epsilon0", capacity.LastSolver.Epsilon0, "", "Параметр плоскости деформаций"
        PutRow target, 12, "KappaX", capacity.LastSolver.KappaX, "1/mm", "Кривизна, связанная с координатой y"
        PutRow target, 13, "KappaY", capacity.LastSolver.KappaY, "1/mm", "Кривизна, связанная с координатой x"
        PutRow target, 14, "StrainLimits", capacity.LastSolver.MinConcreteStrain, "", "MaxAbsSteelStrain=" & CStr(capacity.LastSolver.MaxAbsSteelStrain)
    End If
End Sub

Private Sub WriteDirectResult(ByVal workbook As Object, ByVal rangeName As String, ByVal solver As CSectionSolver)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3604, "CCapacityResultWriter", "Workbook is missing."
    If solver Is Nothing Then Err.Raise vbObjectError + 3605, "CCapacityResultWriter", "Section solver result is missing."

    Dim target As Object
    Set target = workbook.Names.Item(rangeName).RefersToRange
    ClearResultRange target

    PutRow target, 1, "Indicator", "Value", "Unit", "Comment"
    PutRow target, 2, "Mode", "N+Mx+My", "", "Прямой расчет НДС через CSectionSolver"
    PutRow target, 3, "Converged", IIf(solver.Converged, "Yes", "No"), "", solver.StopReason
    PutRow target, 4, "LimitState", "NotCalculated", "", "Поиск несущей способности отключен выбранным режимом Calculation.Mode"
    PutRow target, 5, "LambdaUltimate", "NotCalculated", "", "Для расчета задайте Calculation.Mode = FullCapacity"
    PutRow target, 6, "MxUltimate", "NotCalculated", "kN*m", "Поиск несущей способности отключен"
    PutRow target, 7, "MyUltimate", "NotCalculated", "kN*m", "Поиск несущей способности отключен"
    PutRow target, 8, "MomentResultant", "NotCalculated", "kN*m", "Поиск несущей способности отключен"
    PutRow target, 9, "SafetyFactor", "NotCalculated", "", "Выполнен только прямой расчет заданного состояния"
    PutRow target, 10, "MomentResidual", Sqr(solver.ResidualMx ^ 2 + solver.ResidualMy ^ 2), "N*mm", "Невязка равновесия прямого расчета"
    PutRow target, 11, "Epsilon0", solver.Epsilon0, "", "Параметр плоскости деформаций"
    PutRow target, 12, "KappaX", solver.KappaX, "1/mm", "Кривизна, связанная с координатой y"
    PutRow target, 13, "KappaY", solver.KappaY, "1/mm", "Кривизна, связанная с координатой x"
    PutRow target, 14, "StrainLimits", solver.MinConcreteStrain, "", "MaxAbsSteelStrain=" & CStr(solver.MaxAbsSteelStrain)
End Sub

Private Sub WriteLinearResult(ByVal workbook As Object, ByVal rangeName As String, ByVal solver As CLinearSectionSolver)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3606, "CCapacityResultWriter", "Workbook is missing."
    If solver Is Nothing Then Err.Raise vbObjectError + 3607, "CCapacityResultWriter", "Linear solver result is missing."

    Dim target As Object
    Set target = workbook.Names.Item(rangeName).RefersToRange
    ClearResultRange target

    PutRow target, 1, "Indicator", "Value", "Unit", "Comment"
    PutRow target, 2, "Mode", "N+Mx+My", "", "Линейно-упругий матричный расчет через CLinearSectionSolver"
    PutRow target, 3, "Converged", IIf(solver.Converged, "Yes", "No"), "", solver.StopReason
    PutRow target, 4, "LimitState", "NotCalculated", "", "Линейный режим не определяет предельное состояние"
    PutRow target, 5, "LambdaUltimate", "NotCalculated", "", "Для расчета задайте Calculation.Mode = FullCapacity"
    PutRow target, 6, "MxUltimate", "NotCalculated", "kN*m", "Линейный матричный режим"
    PutRow target, 7, "MyUltimate", "NotCalculated", "kN*m", "Линейный матричный режим"
    PutRow target, 8, "MomentResultant", "NotCalculated", "kN*m", "Линейный матричный режим"
    PutRow target, 9, "SafetyFactor", "NotCalculated", "", "Выполнен только линейно-упругий расчет состояния"
    PutRow target, 10, "MomentResidual", Sqr(solver.ResidualMx ^ 2 + solver.ResidualMy ^ 2), "N*mm", "Невязка равновесия линейного расчета"
    PutRow target, 11, "Epsilon0", solver.Epsilon0, "", "Параметр линейной плоскости деформаций"
    PutRow target, 12, "KappaX", solver.KappaX, "1/mm", "Линейная кривизна, связанная с координатой y"
    PutRow target, 13, "KappaY", solver.KappaY, "1/mm", "Линейная кривизна, связанная с координатой x"
    PutRow target, 14, "StrainLimits", solver.MinConcreteStrain, "", "MaxAbsSteelStrain=" & CStr(MaxLinearSteelStrain(solver))
    PutRow target, 16, "CrackConverged", "NotCalculated", "", "Ширина трещин не рассчитывается в режиме LinearMatrix"
    PutRow target, 17, "CrackWidth", "NotCalculated", "mm", "Используйте Calculation.Mode = DirectState или FullCapacity"
End Sub

Private Function MaxLinearSteelStrain(ByVal solver As CLinearSectionSolver) As Double
    Dim result As Double
    Dim i As Long
    For i = 1 To solver.RebarCount
        If Abs(solver.RebarStrain(i)) > result Then result = Abs(solver.RebarStrain(i))
    Next i
    MaxLinearSteelStrain = result
End Function

Private Sub ClearResultRange(ByVal target As Object)
    Dim rowIndex As Long
    Dim colIndex As Variant
    For rowIndex = 1 To target.Rows.Count
        For Each colIndex In Array(1, 5, 9, 13)
            If Not target.Cells.Item(rowIndex, CLng(colIndex)).MergeCells Then
                target.Cells.Item(rowIndex, CLng(colIndex)).ClearContents
            End If
        Next colIndex
    Next rowIndex
End Sub

Public Sub WriteCrackResult(ByVal workbook As Object, ByVal crack As CCrackWidthCalculator)
    WriteCrackResultToRange workbook, "rngResultSection", crack
End Sub

Private Sub WriteCrackResultToRange(ByVal workbook As Object, ByVal rangeName As String, ByVal crack As CCrackWidthCalculator)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3602, "CCapacityResultWriter", "Workbook is missing."
    If crack Is Nothing Then Err.Raise vbObjectError + 3603, "CCapacityResultWriter", "Crack result is missing."

    Dim target As Object
    Set target = workbook.Names.Item(rangeName).RefersToRange
    PutRow target, 16, "CrackConverged", IIf(crack.Converged, "Yes", "No"), "", crack.StopReason
    PutRow target, 17, "CrackWidth", crack.CrackWidth, "mm", "Уже образовавшаяся нормальная трещина"
    PutRow target, 18, "CrackAllowable", crack.AllowableCrackWidth, "mm", "Utilization=" & CStr(crack.Utilization)
    PutRow target, 19, "SteelMax", crack.MaxSteelStrain, "", "StressMPa=" & CStr(crack.MaxSteelStress) & "; tensionBars=" & CStr(crack.TensionRebarCount)
End Sub
Private Sub PutRow(ByVal target As Object, ByVal rowIndex As Long, _
        ByVal indicator As Variant, ByVal value As Variant, ByVal unitName As Variant, ByVal comment As Variant)
    target.Cells.Item(rowIndex, 1).Value2 = indicator
    target.Cells.Item(rowIndex, 5).Value2 = value
    target.Cells.Item(rowIndex, 9).Value2 = unitName
    target.Cells.Item(rowIndex, 13).Value2 = comment
End Sub





















