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

' ==========================================================================
' Вывод результатов несущей способности в основной отчет
' ==========================================================================
' Класс отвечает за человекочитаемый блок результатов capacity-проверки на
' листе расчета. Он получает готовые значения lambda, статусы и предельные
' моменты, форматирует их для Excel и не участвует в самом поиске lambda.

Public Sub WriteCapacityResult(ByVal workbook As Object, ByVal capacity As CCapacitySolver, Optional ByVal units As CUnitSystem = Nothing)
    WriteResult workbook, "rngResultSection", capacity, units
End Sub

Public Sub WriteDirectSectionResult(ByVal workbook As Object, ByVal solver As CSectionSolver, Optional ByVal units As CUnitSystem = Nothing)
    WriteDirectResult workbook, "rngResultSection", solver, units
End Sub

Public Sub WriteCrackResult(ByVal workbook As Object, ByVal crack As CCrackWidthCalculator, Optional ByVal units As CUnitSystem = Nothing)
    WriteCrackResultToRange workbook, "rngResultSection", crack, units
End Sub

Private Sub WriteResult(ByVal workbook As Object, ByVal rangeName As String, ByVal capacity As CCapacitySolver, ByVal units As CUnitSystem)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3600, "CCapacityResultWriter", "Книга Excel не передана."
    If capacity Is Nothing Then Err.Raise vbObjectError + 3601, "CCapacityResultWriter", "Результат поиска несущей способности не передан."

    Dim target As Object
    Set target = workbook.Names.Item(rangeName).RefersToRange

    Dim rows As Variant
    rows = EmptyResultRows(target.Rows.Count)
    PutRow rows, 1, "Indicator", "Value", "Unit", "Comment"
    PutRow rows, 2, "Mode", "N+Mx+My", "", "Единый расчет по направлению вектора моментов"
    PutRow rows, 3, "Converged", IIf(capacity.Converged, "Yes", "No"), "", capacity.StopReason
    PutRow rows, 4, "LimitState", capacity.LimitState, "", "Физические и численные причины остановки разделяются"
    PutRow rows, 5, "LambdaUltimate", capacity.LambdaUltimate, "", "SafetyFactor = LambdaUltimate"
    PutRow rows, 6, "MxUltimate", OutputMx(capacity.MxUltimate, units), OutputMomentUnit(units), "Внутренние единицы: N*mm"
    PutRow rows, 7, "MyUltimate", OutputMy(capacity.MyUltimate, units), OutputMomentUnit(units), "Внутренние единицы: N*mm"
    PutRow rows, 8, "MomentResultant", OutputMomentMagnitude(capacity.MomentUltimate, units), OutputMomentUnit(units), "Результирующий момент вектора Mx/My"
    PutRow rows, 9, "SafetyFactor", capacity.SafetyFactor, "", "Utilization=" & CStr(capacity.Utilization)
    PutRow rows, 10, "MomentResidual", OutputMomentMagnitude(capacity.MomentEquilibriumResidual, units), OutputMomentUnit(units), "Независимая проверка равновесия"
    If Not capacity.LastSolver Is Nothing Then
        PutRow rows, 11, "Epsilon0", capacity.LastSolver.Epsilon0, "", "Параметр плоскости деформаций"
        PutRow rows, 12, "KappaX", OutputCurvature(capacity.LastSolver.KappaX, units), OutputCurvatureUnit(units), "Кривизна, связанная с координатой y"
        PutRow rows, 13, "KappaY", OutputCurvature(capacity.LastSolver.KappaY, units), OutputCurvatureUnit(units), "Кривизна, связанная с координатой x"
        PutRow rows, 14, "StrainLimits", capacity.LastSolver.MinConcreteStrain, "", "MaxAbsSteelStrain=" & CStr(capacity.LastSolver.MaxAbsSteelStrain)
    End If
    WriteResultRows target, rows
End Sub

Private Sub WriteDirectResult(ByVal workbook As Object, ByVal rangeName As String, ByVal solver As CSectionSolver, ByVal units As CUnitSystem)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3604, "CCapacityResultWriter", "Книга Excel не передана."
    If solver Is Nothing Then Err.Raise vbObjectError + 3605, "CCapacityResultWriter", "Результат решателя сечения не передан."

    Dim target As Object
    Set target = workbook.Names.Item(rangeName).RefersToRange

    Dim rows As Variant
    rows = EmptyResultRows(target.Rows.Count)
    PutRow rows, 1, "Indicator", "Value", "Unit", "Comment"
    PutRow rows, 2, "Mode", "N+Mx+My", "", "Прямой расчет НДС через CSectionSolver"
    PutRow rows, 3, "Converged", IIf(solver.Converged, "Yes", "No"), "", solver.StopReason
    PutRow rows, 4, "LimitState", "NotCalculated", "", "Поиск несущей способности отключен выбранным режимом Calculation.Mode"
    PutRow rows, 5, "LambdaUltimate", "NotCalculated", "", "Для расчета задайте Calculation.Mode = FullCapacity"
    PutRow rows, 6, "MxUltimate", "NotCalculated", OutputMomentUnit(units), "Поиск несущей способности отключен"
    PutRow rows, 7, "MyUltimate", "NotCalculated", OutputMomentUnit(units), "Поиск несущей способности отключен"
    PutRow rows, 8, "MomentResultant", "NotCalculated", OutputMomentUnit(units), "Поиск несущей способности отключен"
    PutRow rows, 9, "SafetyFactor", "NotCalculated", "", "Выполнен только прямой расчет заданного состояния"
    PutRow rows, 10, "MomentResidual", OutputMomentMagnitude(Sqr(solver.ResidualMx ^ 2 + solver.ResidualMy ^ 2), units), OutputMomentUnit(units), "Невязка равновесия прямого расчета"
    PutRow rows, 11, "Epsilon0", solver.Epsilon0, "", "Параметр плоскости деформаций"
    PutRow rows, 12, "KappaX", OutputCurvature(solver.KappaX, units), OutputCurvatureUnit(units), "Кривизна, связанная с координатой y"
    PutRow rows, 13, "KappaY", OutputCurvature(solver.KappaY, units), OutputCurvatureUnit(units), "Кривизна, связанная с координатой x"
    PutRow rows, 14, "StrainLimits", solver.MinConcreteStrain, "", "MaxAbsSteelStrain=" & CStr(solver.MaxAbsSteelStrain)
    WriteResultRows target, rows
End Sub

Private Sub WriteCrackResultToRange(ByVal workbook As Object, ByVal rangeName As String, ByVal crack As CCrackWidthCalculator, ByVal units As CUnitSystem)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3602, "CCapacityResultWriter", "Книга Excel не передана."
    If crack Is Nothing Then Err.Raise vbObjectError + 3603, "CCapacityResultWriter", "Результат расчета трещин не передан."

    Dim target As Object
    Set target = workbook.Names.Item(rangeName).RefersToRange

    Dim rows As Variant
    rows = EmptyResultRows(12)
    PutRow rows, 1, "CrackConverged", IIf(crack.Converged, "Yes", "No"), "", crack.StopReason
    PutRow rows, 2, "CrackFormed", IIf(crack.CrackFormed, "Yes", "No"), "", "Auto/central branch can return No when lambda_crc > 1"
    PutRow rows, 3, "CrackWidth", OutputLength(crack.CrackWidth, units), OutputLengthUnit(units), "a_crc for long-term SLS crack opening"
    PutRow rows, 4, "CrackAllowable", OutputLength(crack.AllowableCrackWidth, units), OutputLengthUnit(units), "Utilization=" & CStr(crack.Utilization)
    PutRow rows, 5, "CrackModes", crack.PsiMode & "; " & crack.TensionZoneMode, "", "psi_s mode and Abt zone mode"
    PutRow rows, 6, "SigmaS", OutputStress(crack.SigmaS, units), OutputStressUnit(units), "sigma_s averaged by selected tensile As"
    If crack.SigmaSCrc > 0# Then
        PutRow rows, 7, "SigmaS_crc", OutputStress(crack.SigmaSCrc, units), OutputStressUnit(units), "Заполняется только если SLS.Crack.PsiMode = Auto запустил расчет lambda_crc"
    Else
        PutRow rows, 7, "SigmaS_crc", "NotCalcul.", OutputStressUnit(units), "Не считалось: Fixed1 или Auto-проверка с psi_s=1 прошла"
    End If
    PutRow rows, 8, "PsiS", crack.PsiS, "", "psi_s = 1 или Auto-уточнение 1 - 0.8*sigma_s,crc/sigma_s"
    PutRow rows, 9, "CrackSpacing", OutputLength(crack.CrackSpacing, units), OutputLengthUnit(units), "ls after limits 10ds/100/40ds/400"
    PutRow rows, 10, "Abt_As_ds", OutputArea(crack.Abt, units) & "; " & OutputArea(crack.AsTension, units) & "; " & OutputLength(crack.DsEquivalent, units), OutputAreaUnit(units) & "; " & OutputLengthUnit(units), "Abt; As; ds,eq"
    If crack.SigmaSCrc > 0# Then
        PutRow rows, 11, "LambdaCrc", crack.LambdaCrc, "", "Множитель образования трещины; считается только при Auto-уточнении"
    Else
        PutRow rows, 11, "LambdaCrc", "NotCalcul.", "", "Не считалось: Fixed1 или Auto-проверка с psi_s=1 прошла"
    End If
    PutRow rows, 12, "SteelMax", crack.MaxSteelStrain, "", "Stress=" & CStr(OutputStress(crack.MaxSteelStress, units)) & " " & OutputStressUnit(units) & "; tensionBars=" & CStr(crack.TensionRebarCount)
    WriteResultRows target.Offset(15, 0).Resize(12, target.Columns.Count), rows
End Sub

Private Function OutputMx(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMx = value / 1000000#: Exit Function
    OutputMx = units.InternalMomentMxToOutput(value)
End Function

Private Function OutputMy(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMy = value / 1000000#: Exit Function
    OutputMy = units.InternalMomentMyToOutput(value)
End Function

Private Function OutputMomentMagnitude(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMomentMagnitude = value / 1000000#: Exit Function
    OutputMomentMagnitude = units.InternalMomentMagnitudeToOutput(value)
End Function

Private Function OutputLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputLength = value: Exit Function
    OutputLength = units.InternalLengthToOutput(value)
End Function

Private Function OutputArea(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputArea = value: Exit Function
    OutputArea = units.InternalAreaToOutput(value)
End Function

Private Function OutputStress(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputStress = value: Exit Function
    OutputStress = units.InternalStressToOutput(value)
End Function

Private Function OutputCurvature(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputCurvature = value: Exit Function
    OutputCurvature = units.InternalCurvatureToOutput(value)
End Function

Private Function OutputMomentUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputMomentUnit = "kN*m" Else OutputMomentUnit = units.OutputMomentUnit
End Function

Private Function OutputLengthUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputLengthUnit = "mm" Else OutputLengthUnit = units.OutputLengthUnit
End Function

Private Function OutputAreaUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputAreaUnit = "mm2" Else OutputAreaUnit = units.OutputAreaUnit
End Function

Private Function OutputStressUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputStressUnit = "MPa" Else OutputStressUnit = units.OutputStressUnit
End Function

Private Function OutputCurvatureUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputCurvatureUnit = "1/mm" Else OutputCurvatureUnit = units.OutputCurvatureUnit
End Function

Private Function EmptyResultRows(ByVal rowCount As Long) As Variant
    Dim rows() As Variant
    ReDim rows(1 To rowCount, 1 To 4)
    EmptyResultRows = rows
End Function

Private Sub PutRow(ByRef rows As Variant, ByVal rowIndex As Long, _
        ByVal indicator As Variant, ByVal value As Variant, ByVal unitName As Variant, ByVal comment As Variant)
    rows(rowIndex, 1) = indicator
    rows(rowIndex, 2) = value
    rows(rowIndex, 3) = unitName
    rows(rowIndex, 4) = comment
End Sub

Private Sub WriteResultRows(ByVal target As Object, ByRef rows As Variant)
    Dim rowCount As Long
    rowCount = UBound(rows, 1)
    WriteResultColumn target, 1, rows, 1, rowCount
    WriteResultColumn target, 5, rows, 2, rowCount
    WriteResultColumn target, 9, rows, 3, rowCount
    WriteResultColumn target, 13, rows, 4, rowCount
End Sub

Private Sub WriteResultColumn(ByVal target As Object, ByVal targetColumn As Long, _
        ByRef rows As Variant, ByVal sourceColumn As Long, ByVal rowCount As Long)
    Dim columnData() As Variant
    ReDim columnData(1 To rowCount, 1 To 1)

    Dim rowIndex As Long
    For rowIndex = 1 To rowCount
        columnData(rowIndex, 1) = rows(rowIndex, sourceColumn)
    Next rowIndex

    target.Offset(0, targetColumn - 1).Resize(rowCount, 1).Value2 = columnData
End Sub

