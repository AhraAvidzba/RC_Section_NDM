Attribute VB_Name = "modTestBatchCalculation"
Option Explicit

' ==========================================================================
' Тесты пакетного расчета и интерфейсных статусов
' ==========================================================================
' Модуль проверяет, что batch-слой правильно обрабатывает несколько LC,
' выбирает худшее сочетание и не смешивает статусы capacity/crack/direct state.

Private Type TBatchTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

' Запускает связанный набор операций и возвращает пользователю итоговый статус выполнения.
Public Function RunBatchCalculationTests() As String
    On Error GoTo Failed

    Dim stats As TBatchTestStats
    Dim t0 As Double
    t0 = Timer

    AppendLine stats, "RUN: TestBatchOneCombination"
    TestBatchOneCombination stats
    AppendLine stats, "RUN: TestCalculationTypeControlsLimitStateGroup"
    TestCalculationTypeControlsLimitStateGroup stats
    AppendLine stats, "RUN: TestBatchFiveCombinations"
    TestBatchFiveCombinations stats
    AppendLine stats, "RUN: TestBatchGoverningUsesLowestSafetyFactor"
    TestBatchGoverningUsesLowestSafetyFactor stats
    AppendLine stats, "RUN: TestBatchGoverningCanUseGroup2StrainSafety"
    TestBatchGoverningCanUseGroup2StrainSafety stats
    AppendLine stats, "RUN: TestLoadReferenceTransformsUserMoments"
    TestLoadReferenceTransformsUserMoments stats
    AppendLine stats, "RUN: TestAxialReferenceRemovesPureCompressionEccentricity"
    TestAxialReferenceRemovesPureCompressionEccentricity stats
    AppendLine stats, "RUN: TestAxialTensionReferenceAndEccentricity"
    TestAxialTensionReferenceAndEccentricity stats
    AppendLine stats, "RUN: TestDirectStateReportsStrainSafety"
    TestDirectStateReportsStrainSafety stats
    AppendLine stats, "RUN: TestDirectStateKeepsStrainSafetyOnFailure"
    TestDirectStateKeepsStrainSafetyOnFailure stats
    AppendLine stats, "RUN: TestBatchTwentyCombinations"
    TestBatchTwentyCombinations stats
    AppendLine stats, "RUN: TestInvalidCombinationFromNamedRange"
    TestInvalidCombinationFromNamedRange stats
    AppendLine stats, "RUN: TestBatchSummaryWriter"
    TestBatchSummaryWriter stats
    AppendLine stats, "RUN: TestBatchCapacityUsesSystemSettings"
    TestBatchCapacityUsesSystemSettings stats
    AppendLine stats, "RUN: TestInvalidModeSettingsAreNotFallbacks"
    TestInvalidModeSettingsAreNotFallbacks stats

    AppendLine stats, "TOTAL_BATCH: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
    RunBatchCalculationTests = stats.Report
    Exit Function

Failed:
    RunBatchCalculationTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestBatchOneCombination(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "C1", -220000#, -7000000#, -5000000#, "Group1", "single"
    batch.Execute

    AssertTrue stats, "batch.one.count", batch.Count = 1
    AssertTrue stats, "batch.one.governing", batch.GoverningCombinationID = "C1"
    AssertTrue stats, "batch.one.capacity.status", Len(batch.CapacityStatus(1)) > 0
    AssertTrue stats, "batch.one.crack.status", Len(batch.CrackStatus(1)) > 0
    AssertTrue stats, "batch.one.elapsed", batch.ElapsedSeconds >= 0#
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestCalculationTypeControlsLimitStateGroup(ByRef stats As TBatchTestStats)
    Dim group1 As CBatchSectionCalculator
    Set group1 = BuildBatchCalculator()
    group1.AddCombination "G1", -220000#, -7000000#, -5000000#, "Group1", "strength"
    group1.Execute

    AssertTrue stats, "batch.calculationType.group1.capacity", group1.LambdaCapacity(1) > 0#
    AssertTrue stats, "batch.calculationType.group1.noCrack", group1.CrackStatus(1) = "NotCalculated"

    Dim group2 As CBatchSectionCalculator
    Set group2 = BuildBatchCalculator()
    group2.AddCombination "G2", -220000#, -7000000#, -5000000#, "Group2", "crack"
    group2.Execute

    AssertTrue stats, "batch.calculationType.group2.noCapacity", group2.CapacityStatus(1) = "NotCalculated"
    AssertTrue stats, "batch.calculationType.group2.crack", group2.CrackStatus(1) <> "NotCalculated"
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestBatchCapacityUsesSystemSettings(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldMaxLambda As String
    oldMode = GetSystemSetting("Calculation.Mode")
    oldMaxLambda = GetSystemSetting("Capacity.MaxLambda")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Calculation.Mode", "FullCapacity"
    SetSystemSetting "Capacity.MaxLambda", "0.5"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "LIMITED", -220000#, -7000000#, -5000000#, "Group1", "max-lambda"
    batch.Execute

    AssertTrue stats, "batch.settings.capacity.maxLambda", InStr(1, batch.CapacityStatus(1), "NumericalFailure", vbTextCompare) > 0

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "Capacity.MaxLambda", oldMaxLambda
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.settings.capacity.maxLambda; " & Err.Description
    Resume Restore
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestInvalidModeSettingsAreNotFallbacks(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCapacityMethod As String
    oldMode = GetSystemSetting("Calculation.Mode")
    oldCapacityMethod = GetSystemSetting("Capacity.Method")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Calculation.Mode", "WrongMode"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "BAD_MODE", -220000#, -7000000#, -5000000#, "Group1", "wrong mode"
    batch.Execute
    AssertTrue stats, "batch.invalid.calculationMode.status", InStr(1, batch.Status(1), "InvalidInput", vbTextCompare) > 0
    AssertTrue stats, "batch.invalid.calculationMode.noCapacity", batch.LambdaCapacity(1) = 0#

    SetSystemSetting "Calculation.Mode", "FullCapacity"
    SetSystemSetting "Capacity.Method", "WrongCapacity"
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "BAD_CAP", -220000#, -7000000#, -5000000#, "Group1", "wrong capacity"
    batch.Execute
    AssertTrue stats, "batch.invalid.capacityMethod.status", InStr(1, batch.Status(1), "InvalidInput", vbTextCompare) > 0
    AssertTrue stats, "batch.invalid.capacityMethod.noLambda", batch.LambdaCapacity(1) = 0#

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "Capacity.Method", oldCapacityMethod
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.invalid.modeSettings; " & Err.Description
    Resume Restore
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestBatchFiveCombinations(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()

    Dim i As Long
    For i = 1 To 5
        batch.AddCombination "C" & CStr(i), -150000# - 10000# * i, -3000000# - 250000# * i, _
            -2000000# - 200000# * i, "Group1", "five-" & CStr(i)
    Next i
    batch.Execute

    AssertTrue stats, "batch.five.count", batch.Count = 5
    AssertTrue stats, "batch.five.governing.index", batch.GoverningCombinationIndex >= 1 And batch.GoverningCombinationIndex <= 5
    AssertTrue stats, "batch.five.diagnostics", InStr(1, batch.DiagnosticLog, "combination=C5", vbTextCompare) > 0
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestBatchGoverningUsesLowestSafetyFactor(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "SAFE", -150000#, -1000000#, -500000#, "Group1", "larger safety"
    batch.AddCombination "GOV", -150000#, -7000000#, -3500000#, "Group1", "smaller safety"
    batch.Execute

    AssertTrue stats, "batch.governing.lambda.order", batch.LambdaCapacity(2) > 0# And batch.LambdaCapacity(2) < batch.LambdaCapacity(1)
    AssertTrue stats, "batch.governing.lowestSafety", batch.GoverningCombinationID = "GOV"
    AssertTrue stats, "batch.governing.limitState", Len(batch.CapacityLimitState(2)) > 0
    AssertTrue stats, "batch.governing.strength.status", _
        (batch.CapacityStatus(2) = "OK" And batch.LambdaCapacity(2) < 1# And batch.StrengthCheckStatus(2) = "StrengthFailed" And batch.Status(2) = "StrengthFailed") Or _
        (batch.CapacityStatus(2) = "OK" And batch.LambdaCapacity(2) >= 1# And batch.StrengthCheckStatus(2) = "StrengthPassed") Or _
        (batch.CapacityStatus(2) <> "OK" And batch.StrengthCheckStatus(2) = "NotCalculated")
End Sub

' Проверяет, что определяющее сочетание по прочности выбирается по минимальному
' запасу среди всех LC, включая строки второй группы. Это важно для режима
' Plot/AutoCAD = Worst: эксплуатационное сочетание может иметь более опасные
' деформации, даже если capacity-поиск для него не запускается.
Private Sub TestBatchGoverningCanUseGroup2StrainSafety(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    oldMode = GetSystemSetting("Calculation.Mode")
    oldCrackEnabled = GetSystemSetting("SLS.Crack.Enabled")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Calculation.Mode", "DirectState"
    SetSystemSetting "SLS.Crack.Enabled", "No"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "G1_SAFE", -120000#, -1200000#, -600000#, "Group1", "first group"
    batch.AddCombination "G2_GOV", -120000#, -5200000#, -2600000#, "Group2", "second group controls strain"
    batch.Execute

    AssertTrue stats, "batch.governing.group2.strain.order", _
        batch.StrainSafetyFactor(2) > 0# And batch.StrainSafetyFactor(2) < batch.StrainSafetyFactor(1)
    AssertTrue stats, "batch.governing.group2.id", batch.GoverningCombinationID = "G2_GOV"

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "SLS.Crack.Enabled", oldCrackEnabled
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.governing.group2; " & Err.Description
    Resume Restore
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestLoadReferenceTransformsUserMoments(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "REF", -1000#, 20000#, -30000#, "Group1", "reference"
    batch.ApplyLoadReference 40#, -25#

    AssertClose stats, "batch.reference.userMx", batch.UserMx(1), 20000#, 0.000001
    AssertClose stats, "batch.reference.userMy", batch.UserMy(1), -30000#, 0.000001
    AssertClose stats, "batch.reference.internalMx", batch.Mx(1), 45000#, 0.000001
    AssertClose stats, "batch.reference.internalMy", batch.My(1), -70000#, 0.000001
    AssertClose stats, "batch.reference.x", batch.LoadReferenceX, 40#, 0.000001
    AssertClose stats, "batch.reference.y", batch.LoadReferenceY, -25#, 0.000001
    AssertClose stats, "batch.reference.offsetX.defaultBase", batch.LoadReferenceOffsetX, 40#, 0.000001
    AssertClose stats, "batch.reference.offsetY.defaultBase", batch.LoadReferenceOffsetY, -25#, 0.000001

    Dim shiftedBatch As CBatchSectionCalculator
    Set shiftedBatch = BuildBatchCalculator()
    shiftedBatch.AddCombination "REF2", -1000#, 20000#, -30000#, "Group1", "reference shifted"
    shiftedBatch.ApplyLoadReference 140#, 75#, 100#, 100#
    AssertClose stats, "batch.reference.offsetX.centroidBase", shiftedBatch.LoadReferenceOffsetX, 40#, 0.000001
    AssertClose stats, "batch.reference.offsetY.centroidBase", shiftedBatch.LoadReferenceOffsetY, -25#, 0.000001
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestAxialReferenceRemovesPureCompressionEccentricity(ByRef stats As TBatchTestStats)
    CheckPureCompressionReference stats, "circle", CircleGeometry(300#, 125#, -75#), _
        CircleRebars(300#, 125#, -75#, 40#, 12, 20#), 25#, 0.00000001
    CheckPureCompressionReference stats, "rounded", RoundedRectangleGeometry(360#, 240#), _
        RectangleRebars(RoundedRectangleGeometry(360#, 240#)), 30#, 0.00000001
    CheckPureCompressionReference stats, "lshape", LShapeGeometry(250#, 550#, 600#, 250#), _
        LShapeRebars(250#, 550#, 600#, 250#, 40#, 50, 32#), 50#, 0.00000001
End Sub

' Проверяет осевое растяжение как отдельный физический сценарий:
' при приложении N через центр тяжести кривизны должны быть нулевыми, а при
' смещении той же силы появляется изгиб. Этот контроль остается в тестах,
' чтобы рабочий алгоритм трещин не использовал кривизны как fallback-критерий.
Private Sub TestAxialTensionReferenceAndEccentricity(ByRef stats As TBatchTestStats)
    CheckPureTensionReference stats, "circle", CircleGeometry(300#, 125#, -75#), _
        CircleRebars(300#, 125#, -75#, 40#, 12, 20#), 25#, 0.00000001
    CheckPureTensionReference stats, "rounded", RoundedRectangleGeometry(360#, 240#), _
        RectangleRebars(RoundedRectangleGeometry(360#, 240#)), 30#, 0.00000001
    CheckPureTensionReference stats, "lshape", LShapeGeometry(250#, 550#, 600#, 250#), _
        LShapeRebars(250#, 550#, 600#, 250#, 40#, 50, 32#), 50#, 0.00000001, -180#, 250#
End Sub

Private Sub CheckPureCompressionReference(ByRef stats As TBatchTestStats, ByVal caseName As String, _
        ByVal geom As ISectionGeometry, ByVal rebars As CRebarLayout, ByVal meshStep As Double, ByVal tolerance As Double)
    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, meshStep, meshStep, 1, 1

    Dim concrete As CConcreteDiagramMaterial
    Set concrete = ProvisionalConcrete()
    Dim steel As CSteelDiagramMaterial
    Set steel = ProvisionalSteel()

    Dim refX As Double
    Dim refY As Double
    CalculateTransformedSectionCentroid BuildGeneratedSectionModel(mesh, rebars), concrete, steel, refX, refY

    Dim nValue As Double
    nValue = -100000#
    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    solver.LoadSteps = 1
    solver.MaxIterations = 40
    solver.Solve BuildGeneratedSectionModel(mesh, rebars), concrete, steel, nValue, nValue * refY, nValue * refX

    AssertTrue stats, "batch.reference." & caseName & ".converged", solver.Converged
    AssertClose stats, "batch.reference." & caseName & ".kappaX", solver.KappaX, 0#, tolerance
    AssertClose stats, "batch.reference." & caseName & ".kappaY", solver.KappaY, 0#, tolerance
End Sub

Private Sub CheckPureTensionReference(ByRef stats As TBatchTestStats, ByVal caseName As String, _
        ByVal geom As ISectionGeometry, ByVal rebars As CRebarLayout, ByVal meshStep As Double, ByVal tolerance As Double, _
        Optional ByVal eccentricOffsetX As Double = -25#, Optional ByVal eccentricOffsetY As Double = 40#)
    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, meshStep, meshStep, 1, 1

    Dim concrete As CConcreteDiagramMaterial
    Set concrete = ProvisionalConcrete()
    Dim steel As CSteelDiagramMaterial
    Set steel = ProvisionalSteel()
    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, rebars)

    Dim refX As Double
    Dim refY As Double
    CalculateTransformedSectionCentroid section, concrete, steel, refX, refY

    Dim nValue As Double
    nValue = 100000#

    Dim axialSolver As CSectionSolver
    Set axialSolver = New CSectionSolver
    axialSolver.LoadSteps = 1
    axialSolver.MaxIterations = 60
    axialSolver.Solve section, concrete, steel, nValue, nValue * refY, nValue * refX

    AssertTrue stats, "batch.tension." & caseName & ".central.converged", axialSolver.Converged
    AssertClose stats, "batch.tension." & caseName & ".central.kappaX", axialSolver.KappaX, 0#, tolerance
    AssertClose stats, "batch.tension." & caseName & ".central.kappaY", axialSolver.KappaY, 0#, tolerance

    Dim eccentricSolver As CSectionSolver
    Set eccentricSolver = New CSectionSolver
    eccentricSolver.LoadSteps = 1
    eccentricSolver.MaxIterations = 60
    eccentricSolver.Solve section, concrete, steel, nValue, nValue * (refY + eccentricOffsetY), _
        nValue * (refX + eccentricOffsetX)

    AssertTrue stats, "batch.tension." & caseName & ".eccentric.converged", eccentricSolver.Converged
    AssertTrue stats, "batch.tension." & caseName & ".eccentric.kappa", _
        Abs(eccentricSolver.KappaX) > tolerance Or Abs(eccentricSolver.KappaY) > tolerance
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestDirectStateReportsStrainSafety(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    oldMode = GetSystemSetting("Calculation.Mode")
    oldCrackEnabled = GetSystemSetting("SLS.Crack.Enabled")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Calculation.Mode", "DirectState"
    SetSystemSetting "SLS.Crack.Enabled", "No"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "DS_SAFE", -120000#, -1500000#, -800000#, "Group1", "smaller strain"
    batch.AddCombination "DS_GOV", -120000#, -4500000#, -2400000#, "Group1", "larger strain"
    batch.Execute

    AssertTrue stats, "batch.direct.lambda.zero", batch.LambdaCapacity(1) = 0# And batch.LambdaCapacity(2) = 0#
    AssertTrue stats, "batch.direct.crack.notCalculated", StrComp(batch.CrackStatus(1), "NotCalculated", vbTextCompare) = 0
    AssertTrue stats, "batch.direct.strainSafety.positive", batch.StrainSafetyFactor(1) > 0# And batch.StrainSafetyFactor(2) > 0#
    AssertTrue stats, "batch.direct.strainSafety.order", batch.StrainSafetyFactor(2) < batch.StrainSafetyFactor(1)
    AssertTrue stats, "batch.direct.governing", batch.GoverningCombinationID = "DS_GOV"

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "SLS.Crack.Enabled", oldCrackEnabled
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.direct.strainSafety; " & Err.Description
    Resume Restore
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestDirectStateKeepsStrainSafetyOnFailure(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldMaxIterations As String
    oldMode = GetSystemSetting("Calculation.Mode")
    oldCrackEnabled = GetSystemSetting("SLS.Crack.Enabled")
    oldMaxIterations = GetSystemSetting("Solver.MaxIterations")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Calculation.Mode", "DirectState"
    SetSystemSetting "SLS.Crack.Enabled", "No"
    SetSystemSetting "Solver.MaxIterations", "1"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "DS_FAIL", -120000#, -12000000#, -7000000#, "Group1", "forced non-convergence"
    batch.Execute

    AssertTrue stats, "batch.direct.failure.status", InStr(1, batch.Status(1), "NumericalFailure", vbTextCompare) > 0 Or batch.Status(1) = "StrainLimitExceeded"
    AssertTrue stats, "batch.direct.failure.strainSafety", batch.StrainSafetyFactor(1) > 0#

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "SLS.Crack.Enabled", oldCrackEnabled
    SetSystemSetting "Solver.MaxIterations", oldMaxIterations
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.direct.failure.strainSafety; " & Err.Description
    Resume Restore
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestBatchTwentyCombinations(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()

    Dim i As Long
    For i = 1 To 20
        batch.AddCombination "LC" & CStr(i), -100000# - 2500# * i, -1800000# - 100000# * i, _
            -1200000# - 75000# * i, "Group1", "twenty-" & CStr(i)
    Next i
    batch.Execute

    AssertTrue stats, "batch.twenty.count", batch.Count = 20
    AssertTrue stats, "batch.twenty.governing.index", batch.GoverningCombinationIndex >= 1 And batch.GoverningCombinationIndex <= 20
    AssertTrue stats, "batch.twenty.last.status", Len(batch.Status(20)) > 0
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestInvalidCombinationFromNamedRange(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()

    Dim loads As Object
    Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.ClearContents
    loads.Cells.Item(1, 1).Value2 = "CombinationID"
    loads.Cells.Item(1, 2).Value2 = "N"
    loads.Cells.Item(1, 3).Value2 = "Mx"
    loads.Cells.Item(1, 4).Value2 = "My"
    loads.Cells.Item(1, 5).Value2 = "CalculationType"
    loads.Cells.Item(1, 6).Value2 = "Comment"
    loads.Cells.Item(2, 1).Value2 = "BAD"
    loads.Cells.Item(2, 2).Value2 = "not-a-number"
    loads.Cells.Item(2, 3).Value2 = -1000000#
    loads.Cells.Item(2, 4).Value2 = -500000#
    loads.Cells.Item(2, 5).Value2 = "Group1"
    loads.Cells.Item(2, 6).Value2 = "invalid source row"

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch
    batch.Execute

    AssertTrue stats, "batch.invalid.reader.count", batch.Count = 1
    AssertTrue stats, "batch.invalid.reader.status", InStr(1, batch.Status(1), "InvalidInput", vbTextCompare) > 0
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestBatchSummaryWriter(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "W1", -180000#, -3500000#, -2500000#, "Group1", "writer"
    batch.Execute

    Dim writer As CBatchResultWriter
    Set writer = New CBatchResultWriter
    writer.WriteSummary ThisWorkbook, batch

    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim summaryRow As Long
    summaryRow = BatchSummaryStartRow()
    AssertTrue stats, "batch.writer.fixedRow", summaryRow = 1
    AssertTrue stats, "batch.writer.noResultOverlap", summaryRow + ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Rows.Count - 1 < ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.Row
    AssertTrue stats, "batch.writer.rangeSize", ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Rows.Count >= 31 And ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Columns.Count >= 46
    AssertTrue stats, "batch.writer.title", CStr(resultsSheet.Cells.Item(summaryRow, 1).Value2) = "Сводка пакетного расчета (Подробнее)"
    AssertTrue stats, "batch.writer.titleNotMerged", Not resultsSheet.Cells.Item(summaryRow, 1).MergeCells
    AssertTrue stats, "batch.writer.titleHyperlink", resultsSheet.Cells.Item(summaryRow, 1).Hyperlinks.Count > 0
    AssertTrue stats, "batch.writer.governing", CStr(resultsSheet.Cells.Item(summaryRow + 1, 5).Value2) = "W1"
    AssertTrue stats, "batch.writer.crackGoverning.row", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 3, 1).Value2), "трещинам", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.loadPoint.relativeLabel", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 5, 1).Value2), "относительно центра тяжести", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.loadPoint.zeroX", CStr(resultsSheet.Cells.Item(summaryRow + 5, 5).Value2) = "X=0 mm"
    AssertTrue stats, "batch.writer.subheader.psi", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 7, 37).Value2), "psi", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.strainFormula.simple", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 9, 15).Formula), "AGGREGATE", vbTextCompare) = 0 And _
        InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 9, 15).Formula), "CHOOSE", vbTextCompare) = 0
    AssertTrue stats, "batch.writer.momentFormula.simple", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 9, 21).Formula), "IFERROR", vbTextCompare) > 0 And _
        InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 9, 21).Formula), "IF(", vbTextCompare) = 0
    AssertTrue stats, "batch.writer.header.strength", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 8, 15).Value2), "StrainSafetyFactor", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.header.limitState", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 8, 16).Value2), "CapacityLimitState", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.header.moment", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 8, 21).Value2), "MomentSafetyFactor", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.header.overall", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 8, 46).Value2), "MinSafetyFactor", vbTextCompare) > 0
End Sub

Private Function BatchSummaryStartRow() As Long
    BatchSummaryStartRow = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Row
End Function

' Создает расчетный или интерфейсный объект из нормализованных исходных данных и локальных настроек.
Private Function BuildBatchCalculator() As CBatchSectionCalculator
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 300#, 200#, 0#, 0#, 0#, 0#

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 30#, 20#, 1

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -90#, -60#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B2", 90#, -60#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B3", -90#, 60#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B4", 90#, 60#, 20#, 0#, "A400", "", geom

    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, rebars, "TestBatch")

    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize section, ProvisionalConcrete(), ProvisionalSteel()
    Set BuildBatchCalculator = batch
End Function

Private Function GetSystemSetting(ByVal key As String) As String
    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), key, vbTextCompare) = 0 Then
            GetSystemSetting = CStr(settings.Cells.Item(rowIndex, 2).Value2)
            Exit Function
        End If
    Next rowIndex
    Err.Raise vbObjectError + 3930, "modTestBatchCalculation", "System setting not found: " & key
End Function

Private Sub SetSystemSetting(ByVal key As String, ByVal value As String)
    Dim settings As Object
    Set settings = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Dim rowIndex As Long
    For rowIndex = 2 To settings.Rows.Count
        If StrComp(CStr(settings.Cells.Item(rowIndex, 1).Value2), key, vbTextCompare) = 0 Then
            settings.Cells.Item(rowIndex, 2).Value2 = value
            Exit Sub
        End If
    Next rowIndex
    Err.Raise vbObjectError + 3931, "modTestBatchCalculation", "System setting not found: " & key
End Sub

Private Function CircleGeometry(ByVal diameter As Double, ByVal centerX As Double, ByVal centerY As Double) As ISectionGeometry
    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByDiameter diameter, centerX, centerY
    Set CircleGeometry = geom
End Function

Private Function RoundedRectangleGeometry(ByVal width As Double, ByVal height As Double) As ISectionGeometry
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize width, height, 0#, 0#, 0#, 0#
    Set RoundedRectangleGeometry = geom
End Function

Private Function LShapeGeometry(ByVal b1 As Double, ByVal h1 As Double, ByVal b2 As Double, ByVal h2 As Double) As ISectionGeometry
    Dim geom As CGeometryLShape
    Set geom = New CGeometryLShape
    geom.Initialize b1, h1, b2, h2, 0#, 0#
    Set LShapeGeometry = geom
End Function

Private Function CircleRebars(ByVal diameter As Double, ByVal centerX As Double, ByVal centerY As Double, _
        ByVal axisDistance As Double, ByVal barCount As Long, ByVal barDiameter As Double) As CRebarLayout
    Dim builder As CCircleRebarLayoutBuilder
    Set builder = New CCircleRebarLayoutBuilder
    Set CircleRebars = builder.Build(diameter, centerX, centerY, axisDistance, barCount, barDiameter, "A400")
End Function

Private Function LShapeRebars(ByVal b1 As Double, ByVal h1 As Double, ByVal b2 As Double, ByVal h2 As Double, _
        ByVal axisDistance As Double, ByVal barCount As Long, ByVal barDiameter As Double) As CRebarLayout
    Dim builder As CLShapeRebarLayoutBuilder
    Set builder = New CLShapeRebarLayoutBuilder
    Set LShapeRebars = builder.Build(b1, h1, b2, h2, 0#, 0#, _
        Array(axisDistance, axisDistance, barDiameter, barDiameter, 13, 12, axisDistance, axisDistance, axisDistance, axisDistance), _
        Array(axisDistance, axisDistance, barDiameter, barDiameter, 6, 6, axisDistance, axisDistance, axisDistance, axisDistance), _
        Array(axisDistance, axisDistance, barDiameter, barDiameter, 6, 6, axisDistance, axisDistance, axisDistance, axisDistance), _
        Array(axisDistance, axisDistance, barDiameter, barDiameter, 13, 12, axisDistance, axisDistance, axisDistance, axisDistance), _
        "A400")
End Function

Private Function RectangleRebars(ByVal geom As ISectionGeometry) As CRebarLayout
    Dim layout As CRebarLayout
    Set layout = New CRebarLayout
    layout.AddBar "R1", -140#, -80#, 20#, 0#, "A400", "test", geom
    layout.AddBar "R2", 140#, -80#, 20#, 0#, "A400", "test", geom
    layout.AddBar "R3", 140#, 80#, 20#, 0#, "A400", "test", geom
    layout.AddBar "R4", -140#, 80#, 20#, 0#, "A400", "test", geom
    Set RectangleRebars = layout
End Function

Private Function ProvisionalConcrete() As CConcreteDiagramMaterial
    Dim concrete As CConcreteDiagramMaterial
    Set concrete = New CConcreteDiagramMaterial
    concrete.Initialize -0.0015, -15.5, -0.0035, -15.5
    concrete.TensionMode = "Ignore"
    Set ProvisionalConcrete = concrete
End Function

Private Function ProvisionalSteel() As CSteelDiagramMaterial
    Dim steel As CSteelDiagramMaterial
    Set steel = New CSteelDiagramMaterial
    steel.Initialize 0.00175, 350#, 0.025
    Set ProvisionalSteel = steel
End Function

Private Sub AssertTrue(ByRef stats As TBatchTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name
    End If
End Sub

Private Sub AssertClose(ByRef stats As TBatchTestStats, ByVal name As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal tolerance As Double)
    If Abs(actual - expected) <= tolerance Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name & "; actual=" & FormatNumberInvariant(actual) & _
            "; expected=" & FormatNumberInvariant(expected)
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name & "; actual=" & FormatNumberInvariant(actual) & _
            "; expected=" & FormatNumberInvariant(expected)
    End If
End Sub

Private Sub AppendLine(ByRef stats As TBatchTestStats, ByVal text As String)
    stats.Report = stats.Report & text & vbCrLf
End Sub

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function




