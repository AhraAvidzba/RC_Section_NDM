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
    AppendLine stats, "RUN: TestBatchGoverningCanUseGroup2CapacitySafety"
    TestBatchGoverningCanUseGroup2CapacitySafety stats
    AppendLine stats, "RUN: TestBatchPureAxialCapacityUsesNult"
    TestBatchPureAxialCapacityUsesNult stats
    AppendLine stats, "RUN: TestBatchLShapeN200CapacityPathNDoesNotNumFail"
    TestBatchLShapeN200CapacityPathNDoesNotNumFail stats
    AppendLine stats, "RUN: TestBatchExplicitCapacityLoadPathScalesMxy"
    TestBatchExplicitCapacityLoadPathScalesMxy stats
    AppendLine stats, "RUN: TestBatchExplicitCapacityLoadPathScalesNWithMoments"
    TestBatchExplicitCapacityLoadPathScalesNWithMoments stats
    AppendLine stats, "RUN: TestBatchCapacityLoadPathVariants"
    TestBatchCapacityLoadPathVariants stats
    AppendLine stats, "RUN: TestBatchCapacityLoadPathAllowsZeroInactiveComponents"
    TestBatchCapacityLoadPathAllowsZeroInactiveComponents stats
    AppendLine stats, "RUN: TestBatchNMxyWithoutMomentsUsesStableForcePath"
    TestBatchNMxyWithoutMomentsUsesStableForcePath stats
    AppendLine stats, "RUN: TestBatchInvalidCapacityLoadPathReportsInputErr"
    TestBatchInvalidCapacityLoadPathReportsInputErr stats
    AppendLine stats, "RUN: TestCapacityScopeGroup1OnlySkipsGroup2"
    TestCapacityScopeGroup1OnlySkipsGroup2 stats
    AppendLine stats, "RUN: TestLoadReferenceTransformsUserMoments"
    TestLoadReferenceTransformsUserMoments stats
    AppendLine stats, "RUN: TestAxialReferenceRemovesPureCompressionEccentricity"
    TestAxialReferenceRemovesPureCompressionEccentricity stats
    AppendLine stats, "RUN: TestAxialTensionReferenceAndEccentricity"
    TestAxialTensionReferenceAndEccentricity stats
    AppendLine stats, "RUN: TestDirectStateReportsSectionStatus"
    TestDirectStateReportsSectionStatus stats
    AppendLine stats, "RUN: TestLongitudinalCrackCheckUsesDirectStateStress"
    TestLongitudinalCrackCheckUsesDirectStateStress stats
    AppendLine stats, "RUN: TestLongitudinalCrackSkippedForGroup1"
    TestLongitudinalCrackSkippedForGroup1 stats
    AppendLine stats, "RUN: TestCapacityOnlySkipsDirectStateAndCrack"
    TestCapacityOnlySkipsDirectStateAndCrack stats
    AppendLine stats, "RUN: TestDirectStateReportsNumericalFailure"
    TestDirectStateReportsNumericalFailure stats
    AppendLine stats, "RUN: TestGroup2PhysicalStateRunsCrackWithExtensionEnabled"
    TestGroup2PhysicalStateRunsCrackWithExtensionEnabled stats
    AppendLine stats, "RUN: TestGroup1AxialTensionBeyondPhysicalLimitUsesExtension"
    TestGroup1AxialTensionBeyondPhysicalLimitUsesExtension stats
    AppendLine stats, "RUN: TestGroup1AxialTensionNearLimitDoesNotJumpToNumFail"
    TestGroup1AxialTensionNearLimitDoesNotJumpToNumFail stats
    AppendLine stats, "RUN: TestGroup1AxialCompressionNearLimitDoesNotJumpToNumFail"
    TestGroup1AxialCompressionNearLimitDoesNotJumpToNumFail stats
    AppendLine stats, "RUN: TestGroup1AxialTensionProgressionAfterLimitIsStableFail"
    TestGroup1AxialTensionProgressionAfterLimitIsStableFail stats
    AppendLine stats, "RUN: TestGroup1AxialCompressionProgressionAfterLimitIsStableFail"
    TestGroup1AxialCompressionProgressionAfterLimitIsStableFail stats
    AppendLine stats, "RUN: TestGroup2AxialTensionBeyondPhysicalLimitUsesExtension"
    TestGroup2AxialTensionBeyondPhysicalLimitUsesExtension stats
    AppendLine stats, "RUN: TestGroup2AxialCompressionBeyondPhysicalLimitUsesExtension"
    TestGroup2AxialCompressionBeyondPhysicalLimitUsesExtension stats
    AppendLine stats, "RUN: TestGroup2BendingBeyondPhysicalLimitUsesExtension"
    TestGroup2BendingBeyondPhysicalLimitUsesExtension stats
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
    AssertTrue stats, "batch.calculationType.group1.noCrack", group1.CrackStatus(1) = "N/A"

    Dim group2 As CBatchSectionCalculator
    Set group2 = BuildBatchCalculator()
    group2.AddCombination "G2", -220000#, -7000000#, -5000000#, "Group2", "crack"
    group2.Execute

    AssertTrue stats, "batch.calculationType.group2.capacity", group2.CapacityStatus(1) <> "N/A"
    AssertTrue stats, "batch.calculationType.group2.crack", group2.CrackStatus(1) <> "N/A"
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

    AssertTrue stats, "batch.settings.capacity.maxLambda", batch.CapacityStatus(1) = "NumFail"

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
    Dim oldCapacitySolutionStrategy As String
    oldMode = GetSystemSetting("Calculation.Mode")
    oldCapacitySolutionStrategy = GetSystemSetting("Capacity.SolutionStrategy")

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
    AssertTrue stats, "batch.invalid.calculationMode.status", batch.Status(1) = "InputErr"
    AssertTrue stats, "batch.invalid.calculationMode.noCapacity", batch.LambdaCapacity(1) = 0#

    SetSystemSetting "Calculation.Mode", "FullCapacity"
    SetSystemSetting "Capacity.SolutionStrategy", "WrongCapacity"
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "BAD_CAP", -220000#, -7000000#, -5000000#, "Group1", "wrong capacity"
    batch.Execute
    AssertTrue stats, "batch.invalid.CapacitySolutionStrategy.status", batch.Status(1) = "InputErr"
    AssertTrue stats, "batch.invalid.CapacitySolutionStrategy.noLambda", batch.LambdaCapacity(1) = 0#

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "Capacity.SolutionStrategy", oldCapacitySolutionStrategy
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
        batch.CapacityStatus(2) = "OK" Or batch.CapacityStatus(2) = "FAIL" Or batch.CapacityStatus(2) = "NumFail"
End Sub

' Проверяет, что batch для чистой продольной силы автоматически выбирает
' траекторию lambda*N. Нулевые пользовательские моменты в этом режиме не
' являются ошибкой: до предела масштабируется именно продольная сила.
Private Sub TestBatchPureAxialCapacityUsesNult(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "N_ONLY", 50000#, 0#, 0#, "Group1", "pure axial"
    batch.Execute

    AssertTrue stats, "batch.nult.path", batch.CapacityLoadPathKey(1) = "LambdaN"
    AssertTrue stats, "batch.nult.solutionMethod", batch.CapacitySolutionMethod(1) = "LoadMultiplier"
    AppendLine stats, "INFO: batch.nult.status=" & batch.CapacityStatus(1) & _
        "; limitState=" & batch.CapacityLimitState(1) & _
        "; solutionMethod=" & batch.CapacitySolutionMethod(1) & _
        "; lambda=" & FormatNumberInvariant(batch.LambdaCapacity(1)) & _
        "; Nult=" & FormatNumberInvariant(batch.NUltimate(1))
    AssertTrue stats, "batch.nult.status", batch.CapacityStatus(1) = "OK" Or batch.CapacityStatus(1) = "FAIL"
    AssertTrue stats, "batch.nult.axialUltimate", Abs(batch.NUltimate(1)) > Abs(batch.N(1))
    AssertTrue stats, "batch.nult.noMomentUltimate", Abs(batch.MomentUltimate(1)) < 0.000001
    AssertTrue stats, "batch.nult.governing", batch.GoverningCombinationID = "N_ONLY"
End Sub

' Проверяет пользовательский сценарий из книги: Г-сечение, нагрузка
' N=-200 тс при принятом знаке +N=Compression, то есть внутреннее растяжение,
' и путь CapacityLoadPath = lambda*N. Точка приложения проходит через бетонный
' центр тяжести, поэтому внутри solver-а вместе с N масштабируются и моменты
' переноса, но пользовательская постановка остается чистым Nult.
Private Sub TestBatchLShapeN200CapacityPathNDoesNotNumFail(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldMethod As String
    Dim oldScope As String
    Dim oldBaseLoadSteps As String
    Dim oldMaxRetries As String
    oldMode = GetSystemSetting("Calculation.Mode")
    oldMethod = GetSystemSetting("Capacity.SolutionStrategy")
    oldScope = GetSystemSetting("Capacity.CalculationScope")
    oldBaseLoadSteps = GetSystemSetting("Capacity.BaseLoadSteps")
    oldMaxRetries = GetSystemSetting("Capacity.MaxRetries")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Calculation.Mode", "FullCapacity"
    SetSystemSetting "Capacity.CalculationScope", "Group1Only"
    SetSystemSetting "Capacity.BaseLoadSteps", "1"
    SetSystemSetting "Capacity.MaxRetries", "0"

    CheckBatchLShapeN200CapacitySolutionStrategy stats, "Auto"
    CheckBatchLShapeN200CapacitySolutionStrategy stats, "UltimateStrain"
    CheckBatchLShapeN200CapacitySolutionStrategy stats, "LoadMultiplier"

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "Capacity.SolutionStrategy", oldMethod
    SetSystemSetting "Capacity.CalculationScope", oldScope
    SetSystemSetting "Capacity.BaseLoadSteps", oldBaseLoadSteps
    SetSystemSetting "Capacity.MaxRetries", oldMaxRetries
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.lshape.n200.capacityPathN; " & Err.Description
    Resume Restore
End Sub

Private Sub CheckBatchLShapeN200CapacitySolutionStrategy(ByRef stats As TBatchTestStats, ByVal methodName As String)
    SetSystemSetting "Capacity.SolutionStrategy", methodName

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.Initialize settings, units

    Dim referenceX As Double
    Dim referenceY As Double
    Dim batch As CBatchSectionCalculator
    Set batch = BuildUserLShapeTensionBatch(referenceX, referenceY, provider)
    batch.ApplySettings settings, units
    batch.AddCombination "G1_N200_" & methodName, 200# * 9806.65, 0#, 0#, _
        "Group1", "user N=-200 tf, lambda*N", ChrW$(&H3BB) & "*N"
    batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
    batch.Execute

    AppendLine stats, "INFO: batch.lshape.n200." & methodName & _
        "; capacityStatus=" & batch.CapacityStatus(1) & _
        "; limitState=" & batch.CapacityLimitState(1) & _
        "; solutionMethod=" & batch.CapacitySolutionMethod(1) & _
        "; lambda=" & FormatNumberInvariant(batch.LambdaCapacity(1)) & _
        "; Nult=" & FormatNumberInvariant(batch.NUltimate(1))
    AssertTrue stats, "batch.lshape.n200." & methodName & ".notNumFail", batch.CapacityStatus(1) <> "NumFail"
    AssertTrue stats, "batch.lshape.n200." & methodName & ".capacityStatus", _
        batch.CapacityStatus(1) = "OK" Or batch.CapacityStatus(1) = "FAIL"
    AssertTrue stats, "batch.lshape.n200." & methodName & ".nult", Abs(batch.NUltimate(1)) > Abs(batch.N(1))
    AssertTrue stats, "batch.lshape.n200." & methodName & ".solutionMethod", _
        batch.CapacitySolutionMethod(1) = "LoadMultiplier"
End Sub

' Проверяет, что явный выбор lambda*Mxy масштабирует оба пользовательских
' момента при постоянной продольной силе. Это основной вариант для общего
' изгиба N + Mx + My, когда нужно найти предельный момент при заданной N.
Private Sub TestBatchExplicitCapacityLoadPathScalesMxy(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "M_BRANCH", -150000#, -3000000#, -1000000#, "Group1", "moment branch", ChrW$(&H3BB) & "*Mxy"
    batch.Execute

    AssertTrue stats, "batch.capacityPath.mxy.value", batch.CapacityLoadPath(1) = ChrW$(&H3BB) & "*Mxy"
    AssertTrue stats, "batch.capacityPath.mxy.key", batch.CapacityLoadPathKey(1) = "LambdaMxy"
    AssertTrue stats, "batch.capacityPath.mxy.mx", Abs(batch.MxUltimate(1)) > 0#
    AssertTrue stats, "batch.capacityPath.mxy.my", Abs(batch.MyUltimate(1)) > 0#
End Sub

' Проверяет, что явный выбор lambda*N масштабирует продольную силу даже при
' наличии пользовательских моментов. Моменты остаются постоянной частью
' траектории, а момент от эксцентриситета N масштабируется вместе с N.
Private Sub TestBatchExplicitCapacityLoadPathScalesNWithMoments(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "N_BRANCH", -150000#, -3000000#, -1000000#, "Group1", "axial branch", ChrW$(&H3BB) & "*N"
    batch.Execute

    AssertTrue stats, "batch.capacityPath.n.value", batch.CapacityLoadPath(1) = ChrW$(&H3BB) & "*N"
    AssertTrue stats, "batch.capacityPath.n.key", batch.CapacityLoadPathKey(1) = "LambdaN"
    AssertTrue stats, "batch.capacityPath.n.nult", Abs(batch.NUltimate(1)) > 0#
    AssertTrue stats, "batch.capacityPath.n.status", batch.CapacityStatus(1) = "OK" Or batch.CapacityStatus(1) = "FAIL" Or batch.CapacityStatus(1) = "NumFail"
End Sub

' Проверяет все пользовательские варианты CapacityLoadPath. Тест не
' привязывается к конкретной величине запаса: здесь важно, что batch
' корректно распознает путь и не подменяет выбранную пользователем траекторию.
Private Sub TestBatchCapacityLoadPathVariants(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "PATH_MX", -150000#, -3000000#, 0#, "Group1", "lambda mx", ChrW$(&H3BB) & "*Mx"
    batch.AddCombination "PATH_MY", -150000#, 0#, -3000000#, "Group1", "lambda my", ChrW$(&H3BB) & "*My"
    batch.AddCombination "PATH_MXY", -150000#, -3000000#, -1000000#, "Group1", "lambda mxy", ChrW$(&H3BB) & "*Mxy"
    batch.AddCombination "PATH_N", -150000#, -3000000#, -1000000#, "Group1", "lambda n", ChrW$(&H3BB) & "*N"
    batch.AddCombination "PATH_ALL", -150000#, -3000000#, -1000000#, "Group1", "lambda all", ChrW$(&H3BB) & "*NMxy"
    batch.Execute

    AssertTrue stats, "batch.capacityPath.mx.key", batch.CapacityLoadPathKey(1) = "LambdaMx"
    AssertTrue stats, "batch.capacityPath.my.key", batch.CapacityLoadPathKey(2) = "LambdaMy"
    AssertTrue stats, "batch.capacityPath.mxy.key", batch.CapacityLoadPathKey(3) = "LambdaMxy"
    AssertTrue stats, "batch.capacityPath.n.key", batch.CapacityLoadPathKey(4) = "LambdaN"
    AssertTrue stats, "batch.capacityPath.all.key", batch.CapacityLoadPathKey(5) = "LambdaNMxy"
    AssertTrue stats, "batch.capacityPath.noInputErr", _
        batch.CapacityStatus(1) <> "InputErr" And batch.CapacityStatus(2) <> "InputErr" And _
        batch.CapacityStatus(3) <> "InputErr" And batch.CapacityStatus(4) <> "InputErr" And _
        batch.CapacityStatus(5) <> "InputErr"
End Sub

' Проверяет пользовательский контракт таблицы сочетаний: неучаствующие
' компоненты нагрузки могут быть пустыми/нулевыми. Ошибкой является только
' полностью нулевой масштабируемый Base-вектор, потому что тогда lambda не
' имеет расчетного смысла.
Private Sub TestBatchCapacityLoadPathAllowsZeroInactiveComponents(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "MX_NO_N", 0#, -3000000#, 0#, "Group1", "lambda mx without N", ChrW$(&H3BB) & "*Mx"
    batch.AddCombination "MY_NO_N", 0#, 0#, -3000000#, "Group1", "lambda my without N", ChrW$(&H3BB) & "*My"
    batch.AddCombination "MXY_ONLY_MX", 0#, -3000000#, 0#, "Group1", "lambda mxy with only Mx", ChrW$(&H3BB) & "*Mxy"
    batch.AddCombination "MXY_ONLY_MY", 0#, 0#, -3000000#, "Group1", "lambda mxy with only My", ChrW$(&H3BB) & "*Mxy"
    batch.AddCombination "NMXY_ONLY_N", -150000#, 0#, 0#, "Group1", "lambda nmxy with only N", ChrW$(&H3BB) & "*NMxy"
    batch.AddCombination "NMXY_ONLY_MX", 0#, -3000000#, 0#, "Group1", "lambda nmxy with only Mx", ChrW$(&H3BB) & "*NMxy"
    batch.AddCombination "NMXY_ONLY_MY", 0#, 0#, -3000000#, "Group1", "lambda nmxy with only My", ChrW$(&H3BB) & "*NMxy"
    batch.AddCombination "NMXY_EMPTY", 0#, 0#, 0#, "Group1", "lambda nmxy empty", ChrW$(&H3BB) & "*NMxy"
    batch.Execute

    AssertTrue stats, "batch.capacityPath.zero.mx.noInputErr", batch.CapacityStatus(1) <> "InputErr"
    AssertTrue stats, "batch.capacityPath.zero.my.noInputErr", batch.CapacityStatus(2) <> "InputErr"
    AssertTrue stats, "batch.capacityPath.zero.mxyMx.noInputErr", batch.CapacityStatus(3) <> "InputErr"
    AssertTrue stats, "batch.capacityPath.zero.mxyMy.noInputErr", batch.CapacityStatus(4) <> "InputErr"
    AssertTrue stats, "batch.capacityPath.zero.nmxyN.noInputErr", batch.CapacityStatus(5) <> "InputErr"
    AssertTrue stats, "batch.capacityPath.zero.nmxyMx.noInputErr", batch.CapacityStatus(6) <> "InputErr"
    AssertTrue stats, "batch.capacityPath.zero.nmxyMy.noInputErr", batch.CapacityStatus(7) <> "InputErr"
    AssertTrue stats, "batch.capacityPath.zero.nmxyEmpty.inputErr", batch.CapacityStatus(8) = "InputErr"
    AssertTrue stats, "batch.capacityPath.zero.nmxyN.pathKept", batch.CapacityLoadPathKey(5) = "LambdaNMxy"
    AssertTrue stats, "batch.capacityPath.zero.nmxyMx.pathKept", batch.CapacityLoadPathKey(6) = "LambdaNMxy"
    AssertTrue stats, "batch.capacityPath.zero.nmxyMy.pathKept", batch.CapacityLoadPathKey(7) = "LambdaNMxy"
End Sub

' Проверяет вырожденный пользовательский случай: выбран lambda*NMxy, но в строке
' сочетания Mx=0 и My=0. Это не ошибка ввода и не особая геометрия; по
' фактическим нагрузкам пользователь масштабирует только продольную силу, а
' значит capacity должен идти устойчивым силовым путем LoadMultiplier.
Private Sub TestBatchNMxyWithoutMomentsUsesStableForcePath(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldMethod As String
    oldMode = GetSystemSetting("Calculation.Mode")
    oldMethod = GetSystemSetting("Capacity.SolutionStrategy")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Calculation.Mode", "FullCapacity"
    SetSystemSetting "Capacity.SolutionStrategy", "UltimateStrain"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.Initialize settings, units

    Dim referenceX As Double
    Dim referenceY As Double
    Dim batch As CBatchSectionCalculator
    Set batch = BuildUserLShapeTensionBatch(referenceX, referenceY, provider)
    batch.ApplySettings settings, units
    batch.AddCombination "NMXY_ZERO_M", 100# * 9806.65, 0#, 0#, _
        "Group1", "lambda NMxy with zero moments", ChrW$(&H3BB) & "*NMxy"
    batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
    batch.Execute

    AssertTrue stats, "batch.capacityPath.nmxyZeroM.path", batch.CapacityLoadPathKey(1) = "LambdaNMxy"
    AssertTrue stats, "batch.capacityPath.nmxyZeroM.solutionMethod", batch.CapacitySolutionMethod(1) = "LoadMultiplier"
    AssertTrue stats, "batch.capacityPath.nmxyZeroM.notNumFail", batch.CapacityStatus(1) <> "NumFail"
    AssertTrue stats, "batch.capacityPath.nmxyZeroM.nult", Abs(batch.NUltimate(1)) > Abs(batch.N(1))

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "Capacity.SolutionStrategy", oldMethod
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.capacityPath.nmxyZeroM; " & Err.Description
    Resume Restore
End Sub

' Проверяет, что ошибочный текст CapacityLoadPath не заменяется молча
' авто-выбором. Пользователь должен сразу увидеть ошибку в строке LC.
Private Sub TestBatchInvalidCapacityLoadPathReportsInputErr(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "BAD_CONST", -150000#, -3000000#, 0#, "Group1", "bad branch", "WrongPath"
    batch.Execute

    AssertTrue stats, "batch.capacityPath.invalid.status", batch.CapacityStatus(1) = "InputErr"
    AssertTrue stats, "batch.capacityPath.invalid.overall", batch.OverallStatus(1) = "InputErr"
    AssertTrue stats, "batch.capacityPath.invalid.noPath", Len(batch.CapacityLoadPathKey(1)) = 0
End Sub

' Проверяет, что физическое плато нормативной диаграммы не считается
' numerical extension. Extension начинается только после eps_ult; иначе
' нормальное состояние с запасом по capacity ошибочно получит FAIL.
Private Sub TestGroup2PhysicalStateRunsCrackWithExtensionEnabled(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldExtension As String
    oldMode = GetSystemSetting("Calculation.Mode")
    oldCrackEnabled = GetSystemSetting("SLS.Crack.Enabled")
    oldExtension = GetSystemSetting("Solver.DirectState.DiagramExtension")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Calculation.Mode", "DirectState"
    SetSystemSetting "SLS.Crack.Enabled", "Yes"
    SetSystemSetting "Solver.DirectState.DiagramExtension", "Yes"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "G2_PHYS", -220000#, -7000000#, -5000000#, "Group2", "physical state"
    batch.Execute

    AssertTrue stats, "batch.group2.physical.noExtension", Not batch.ExtensionUsed(1)
    AssertTrue stats, "batch.group2.physical.directOK", batch.DirectStateStatus(1) = "OK"
    AssertTrue stats, "batch.group2.physical.crackRuns", batch.CrackStatus(1) <> "N/A"

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "SLS.Crack.Enabled", oldCrackEnabled
    SetSystemSetting "Solver.DirectState.DiagramExtension", oldExtension
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.group2.physical; " & Err.Description
    Resume Restore
End Sub

' Проверяет сценарий из пользовательского расчета: большое осевое растяжение
' второй группы должно доходить до технического продолжения диаграммы и давать
' инженерный FAIL, а не теряться как численная несходимость NumFail.
Private Sub TestGroup2AxialTensionBeyondPhysicalLimitUsesExtension(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldMaxIterations As String
    Dim oldLoadSteps As String
    Dim oldExtension As String
    oldMode = GetSystemSetting("Calculation.Mode")
    oldCrackEnabled = GetSystemSetting("SLS.Crack.Enabled")
    oldMaxIterations = GetSystemSetting("Solver.MaxIterations")
    oldLoadSteps = GetSystemSetting("Solver.LoadSteps")
    oldExtension = GetSystemSetting("Solver.DirectState.DiagramExtension")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Calculation.Mode", "DirectState"
    SetSystemSetting "SLS.Crack.Enabled", "Yes"
    SetSystemSetting "Solver.DirectState.DiagramExtension", "Yes"
    SetSystemSetting "Solver.MaxIterations", "80"
    SetSystemSetting "Solver.LoadSteps", "1"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.Initialize settings, units

    Dim referenceX As Double
    Dim referenceY As Double
    Dim batch As CBatchSectionCalculator
    Set batch = BuildUserLShapeTensionBatch(referenceX, referenceY, provider)
    batch.ApplySettings settings, units
    batch.AddCombination "G2_TENSION_EXT", 900# * 9806.65, 0#, 0#, "Group2", "tension over SLS yield"
    batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
    batch.Execute

    AssertTrue stats, "batch.group2.extension.directFail", batch.DirectStateStatus(1) = "FAIL"
    AssertTrue stats, "batch.group2.extension.used", batch.ExtensionUsed(1)
    AssertTrue stats, "batch.group2.extension.noCrack", batch.CrackStatus(1) = "N/A"
    AssertTrue stats, "batch.group2.extension.overall", batch.OverallStatus(1) = "FAIL"
    AssertTrue stats, "batch.group2.extension.strain", batch.MaxSteelStrain(1) > 0.025

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "SLS.Crack.Enabled", oldCrackEnabled
    SetSystemSetting "Solver.MaxIterations", oldMaxIterations
    SetSystemSetting "Solver.LoadSteps", oldLoadSteps
    SetSystemSetting "Solver.DirectState.DiagramExtension", oldExtension
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.group2.extension.tension; " & Err.Description
    Resume Restore
End Sub

' Проверяет пользовательский сценарий Group1: растянутый бетон в прочности не
' работает, перегрузка должна распознаваться через extension арматуры, а не
' превращаться в рассинхрон DirectStateStatus=FAIL при ExtensionUsed=False.
Private Sub TestGroup1AxialTensionBeyondPhysicalLimitUsesExtension(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldMaxIterations As String
    Dim oldLoadSteps As String
    Dim oldExtension As String
    oldMode = GetSystemSetting("Calculation.Mode")
    oldCrackEnabled = GetSystemSetting("SLS.Crack.Enabled")
    oldMaxIterations = GetSystemSetting("Solver.MaxIterations")
    oldLoadSteps = GetSystemSetting("Solver.LoadSteps")
    oldExtension = GetSystemSetting("Solver.DirectState.DiagramExtension")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Calculation.Mode", "DirectState"
    SetSystemSetting "SLS.Crack.Enabled", "Yes"
    SetSystemSetting "Solver.DirectState.DiagramExtension", "Yes"
    SetSystemSetting "Solver.MaxIterations", "100"
    SetSystemSetting "Solver.LoadSteps", "1"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.Initialize settings, units

    Dim referenceX As Double
    Dim referenceY As Double
    Dim batch As CBatchSectionCalculator
    Set batch = BuildUserLShapeTensionBatch(referenceX, referenceY, provider)
    batch.ApplySettings settings, units
    batch.AddCombination "G1_TENSION_EXT", 900# * 9806.65, 0#, 0#, "Group1", "tension over ULS diagram"
    batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
    batch.Execute

    AssertTrue stats, "batch.group1.extension.directFail", batch.DirectStateStatus(1) = "FAIL"
    AssertTrue stats, "batch.group1.extension.used", batch.ExtensionUsed(1)
    AssertTrue stats, "batch.group1.extension.noCrack", batch.CrackStatus(1) = "N/A"
    AssertTrue stats, "batch.group1.extension.overall", batch.OverallStatus(1) = "FAIL"

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "SLS.Crack.Enabled", oldCrackEnabled
    SetSystemSetting "Solver.MaxIterations", oldMaxIterations
    SetSystemSetting "Solver.LoadSteps", oldLoadSteps
    SetSystemSetting "Solver.DirectState.DiagramExtension", oldExtension
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.group1.extension.tension; " & Err.Description
    Resume Restore
End Sub

' Проверяет ряд почти соседних осевых растягивающих нагрузок возле физического
' предела. Первый найденный warm-start может сорваться на одной точке
' ряда и давал NumFail между двумя корректными FAIL; теперь batch пробует
' несколько стартов и не должен терять равновесие из-за неудачной подсказки.
Private Sub TestGroup1AxialTensionNearLimitDoesNotJumpToNumFail(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldMaxIterations As String
    Dim oldLoadSteps As String
    Dim oldExtension As String
    oldMode = GetSystemSetting("Calculation.Mode")
    oldCrackEnabled = GetSystemSetting("SLS.Crack.Enabled")
    oldMaxIterations = GetSystemSetting("Solver.MaxIterations")
    oldLoadSteps = GetSystemSetting("Solver.LoadSteps")
    oldExtension = GetSystemSetting("Solver.DirectState.DiagramExtension")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Calculation.Mode", "DirectState"
    SetSystemSetting "SLS.Crack.Enabled", "Yes"
    SetSystemSetting "Solver.DirectState.DiagramExtension", "Yes"
    SetSystemSetting "Solver.MaxIterations", "100"
    SetSystemSetting "Solver.LoadSteps", "1"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.Initialize settings, units

    Dim referenceX As Double
    Dim referenceY As Double
    Dim batch As CBatchSectionCalculator
    Set batch = BuildUserLShapeTensionBatch(referenceX, referenceY, provider)
    batch.ApplySettings settings, units

    Dim loads As Variant
    loads = Array(796#, 797#, 800#, 801#, 809#, 810#, 811#, 812#)

    Dim i As Long
    For i = LBound(loads) To UBound(loads)
        batch.AddCombination "G1_T" & CStr(CLng(loads(i))), CDbl(loads(i)) * 9806.65, 0#, 0#, _
            "Group1", "near physical tension limit"
    Next i
    batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
    batch.Execute

    For i = 1 To UBound(loads) - LBound(loads) + 1
        AppendLine stats, "INFO: batch.group1.nearLimit." & batch.CombinationID(i) & _
            " direct=" & batch.DirectStateStatus(i) & _
            "; extensionUsed=" & CStr(batch.ExtensionUsed(i)) & _
            "; epsSmax=" & FormatNumberInvariant(batch.MaxSteelStrain(i))
        If batch.DirectStateStatus(i) = "NumFail" Then
            AppendLine stats, "DIAG: batch.group1.nearLimit." & batch.CombinationID(i) & vbCrLf & batch.DiagnosticLog
        End If
        AssertTrue stats, "batch.group1.nearLimit.noNumFail." & batch.CombinationID(i), _
            batch.DirectStateStatus(i) = "OK" Or batch.DirectStateStatus(i) = "FAIL"
        If batch.DirectStateStatus(i) = "FAIL" Then
            AssertTrue stats, "batch.group1.nearLimit.failUsesExtension." & batch.CombinationID(i), batch.ExtensionUsed(i)
        End If
    Next i

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "SLS.Crack.Enabled", oldCrackEnabled
    SetSystemSetting "Solver.MaxIterations", oldMaxIterations
    SetSystemSetting "Solver.LoadSteps", oldLoadSteps
    SetSystemSetting "Solver.DirectState.DiagramExtension", oldExtension
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.group1.nearLimit; " & Err.Description
    Resume Restore
End Sub

' Проверяет соседние значения осевого сжатия около предела сечения.
' Регрессия защищает от численной дырки, когда одно значение N между двумя
' корректными FAIL ошибочно превращается в NumFail из-за неудачного старта
' Newton на technical extension.
Private Sub TestGroup1AxialCompressionNearLimitDoesNotJumpToNumFail(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldMaxIterations As String
    Dim oldLoadSteps As String
    Dim oldExtension As String
    oldMode = GetSystemSetting("Calculation.Mode")
    oldCrackEnabled = GetSystemSetting("SLS.Crack.Enabled")
    oldMaxIterations = GetSystemSetting("Solver.MaxIterations")
    oldLoadSteps = GetSystemSetting("Solver.LoadSteps")
    oldExtension = GetSystemSetting("Solver.DirectState.DiagramExtension")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Calculation.Mode", "DirectState"
    SetSystemSetting "SLS.Crack.Enabled", "Yes"
    SetSystemSetting "Solver.DirectState.DiagramExtension", "Yes"
    SetSystemSetting "Solver.MaxIterations", "100"
    SetSystemSetting "Solver.LoadSteps", "1"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.Initialize settings, units

    Dim referenceX As Double
    Dim referenceY As Double
    Dim batch As CBatchSectionCalculator
    Set batch = BuildUserLShapeTensionBatch(referenceX, referenceY, provider)
    batch.ApplySettings settings, units

    Dim loads As Variant
    loads = Array(1222#, 1223#, 1224#, 1225#)

    Dim i As Long
    For i = LBound(loads) To UBound(loads)
        batch.AddCombination "G1_C" & CStr(CLng(loads(i))), -CDbl(loads(i)) * 9806.65, 0#, 0#, _
            "Group1", "near physical compression limit"
    Next i
    batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
    batch.Execute

    For i = 1 To UBound(loads) - LBound(loads) + 1
        AppendLine stats, "INFO: batch.group1.compressionNearLimit." & batch.CombinationID(i) & _
            " direct=" & batch.DirectStateStatus(i) & _
            "; extensionUsed=" & CStr(batch.ExtensionUsed(i)) & _
            "; epsCmin=" & FormatNumberInvariant(batch.MinConcreteStrain(i))
        If batch.DirectStateStatus(i) = "NumFail" Then
            AppendLine stats, "DIAG: batch.group1.compressionNearLimit." & batch.CombinationID(i) & vbCrLf & batch.DiagnosticLog
        End If
        AssertTrue stats, "batch.group1.compressionNearLimit.noNumFail." & batch.CombinationID(i), _
            batch.DirectStateStatus(i) = "OK" Or batch.DirectStateStatus(i) = "FAIL"
        If batch.DirectStateStatus(i) = "FAIL" Then
            AssertTrue stats, "batch.group1.compressionNearLimit.failUsesExtension." & batch.CombinationID(i), batch.ExtensionUsed(i)
        End If
    Next i

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "SLS.Crack.Enabled", oldCrackEnabled
    SetSystemSetting "Solver.MaxIterations", oldMaxIterations
    SetSystemSetting "Solver.LoadSteps", oldLoadSteps
    SetSystemSetting "Solver.DirectState.DiagramExtension", oldExtension
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.group1.compressionNearLimit; " & Err.Description
    Resume Restore
End Sub

' Проверяет длинный растянутый ряд после последнего физически проходящего
' состояния. Набор нагрузок намеренно неровный: сначала шаг 1 т, затем скачок,
' затем снова малый шаг на более высокой нагрузке. Так тест ловит не только
' соседний порог, но и зависимость warm-start от величины предыдущего скачка.
Private Sub TestGroup1AxialTensionProgressionAfterLimitIsStableFail(ByRef stats As TBatchTestStats)
    Dim loads As Variant
    loads = Array(750#, 796#, 797#, 798#, 799#, 800#, 801#, 805#, 810#, 811#, 812#, _
        820#, 830#, 850#, 900#, 925#, 930#, 950#, 1000#, 1200#)
    RunAxialProgressionAfterLimit stats, "batch.group1.tensionProgression", loads, True, 1
End Sub

' Проверяет аналогичный ряд для сжатия. Это защищает от ситуации, когда при
' росте N одно сочетание внутри перегруженной области внезапно получает NumFail,
' хотя соседние нагрузки уже корректно сходятся в extension и дают FAIL.
Private Sub TestGroup1AxialCompressionProgressionAfterLimitIsStableFail(ByRef stats As TBatchTestStats)
    Dim loads As Variant
    loads = Array(1200#, 1221#, 1222#, 1223#, 1224#, 1225#, 1226#, 1230#, 1231#, 1240#, 1250#, _
        1275#, 1300#, 1350#, 1400#, 1500#, 1600#, 1700#, 1800#, 2000#)
    RunAxialProgressionAfterLimit stats, "batch.group1.compressionProgression", loads, False, 1
End Sub

' Общая проверка длинного осевого ряда. firstOkIndex указывает строку, которая
' должна еще оставаться физически допустимой; все последующие строки обязаны
' сходиться как FAIL с ExtensionUsed=True, а не как NumFail.
Private Sub RunAxialProgressionAfterLimit(ByRef stats As TBatchTestStats, ByVal testPrefix As String, _
        ByVal loads As Variant, ByVal isTension As Boolean, ByVal firstOkIndex As Long)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldMaxIterations As String
    Dim oldLoadSteps As String
    Dim oldExtension As String
    oldMode = GetSystemSetting("Calculation.Mode")
    oldCrackEnabled = GetSystemSetting("SLS.Crack.Enabled")
    oldMaxIterations = GetSystemSetting("Solver.MaxIterations")
    oldLoadSteps = GetSystemSetting("Solver.LoadSteps")
    oldExtension = GetSystemSetting("Solver.DirectState.DiagramExtension")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Calculation.Mode", "DirectState"
    SetSystemSetting "SLS.Crack.Enabled", "Yes"
    SetSystemSetting "Solver.DirectState.DiagramExtension", "Yes"
    SetSystemSetting "Solver.MaxIterations", "100"
    SetSystemSetting "Solver.LoadSteps", "1"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.Initialize settings, units

    Dim referenceX As Double
    Dim referenceY As Double
    Dim batch As CBatchSectionCalculator
    Set batch = BuildUserLShapeTensionBatch(referenceX, referenceY, provider)
    batch.ApplySettings settings, units

    Dim i As Long
    Dim signedN As Double
    For i = LBound(loads) To UBound(loads)
        If isTension Then
            signedN = CDbl(loads(i)) * 9806.65
        Else
            signedN = -CDbl(loads(i)) * 9806.65
        End If
        batch.AddCombination ProgressionCombinationID(isTension, CDbl(loads(i))), signedN, 0#, 0#, _
            "Group1", "axial progression after physical limit"
    Next i
    batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
    batch.Execute

    For i = 1 To UBound(loads) - LBound(loads) + 1
        AppendLine stats, "INFO: " & testPrefix & "." & batch.CombinationID(i) & _
            " direct=" & batch.DirectStateStatus(i) & _
            "; extensionUsed=" & CStr(batch.ExtensionUsed(i)) & _
            "; epsCmin=" & FormatNumberInvariant(batch.MinConcreteStrain(i)) & _
            "; epsSmax=" & FormatNumberInvariant(batch.MaxSteelStrain(i))
        If i = firstOkIndex Then
            AssertTrue stats, testPrefix & ".lastPhysicalOK." & batch.CombinationID(i), _
                batch.DirectStateStatus(i) = "OK" And Not batch.ExtensionUsed(i)
        Else
            If batch.DirectStateStatus(i) = "NumFail" Then
                AppendLine stats, "DIAG: " & testPrefix & "." & batch.CombinationID(i) & vbCrLf & batch.DiagnosticLog
            End If
            AssertTrue stats, testPrefix & ".stableFail." & batch.CombinationID(i), _
                batch.DirectStateStatus(i) = "FAIL"
            AssertTrue stats, testPrefix & ".usesExtension." & batch.CombinationID(i), batch.ExtensionUsed(i)
        End If
    Next i

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "SLS.Crack.Enabled", oldCrackEnabled
    SetSystemSetting "Solver.MaxIterations", oldMaxIterations
    SetSystemSetting "Solver.LoadSteps", oldLoadSteps
    SetSystemSetting "Solver.DirectState.DiagramExtension", oldExtension
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: " & testPrefix & "; " & Err.Description
    Resume Restore
End Sub

Private Function ProgressionCombinationID(ByVal isTension As Boolean, ByVal loadTf As Double) As String
    If isTension Then
        ProgressionCombinationID = "G1_TP" & CStr(CLng(loadTf))
    Else
        ProgressionCombinationID = "G1_CP" & CStr(CLng(loadTf))
    End If
End Function

' Проверяет симметричный для сжатия сценарий: если заданное N больше
' физической сжатой области диаграммы, повторный StateSolution должен найти
' формальное равновесие в compression-extension и вернуть FAIL, а не NumFail.
Private Sub TestGroup2AxialCompressionBeyondPhysicalLimitUsesExtension(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldMaxIterations As String
    Dim oldLoadSteps As String
    Dim oldExtension As String
    oldMode = GetSystemSetting("Calculation.Mode")
    oldCrackEnabled = GetSystemSetting("SLS.Crack.Enabled")
    oldMaxIterations = GetSystemSetting("Solver.MaxIterations")
    oldLoadSteps = GetSystemSetting("Solver.LoadSteps")
    oldExtension = GetSystemSetting("Solver.DirectState.DiagramExtension")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Calculation.Mode", "DirectState"
    SetSystemSetting "SLS.Crack.Enabled", "Yes"
    SetSystemSetting "Solver.DirectState.DiagramExtension", "Yes"
    SetSystemSetting "Solver.MaxIterations", "80"
    SetSystemSetting "Solver.LoadSteps", "1"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.Initialize settings, units

    Dim referenceX As Double
    Dim referenceY As Double
    Dim batch As CBatchSectionCalculator
    Set batch = BuildUserLShapeTensionBatch(referenceX, referenceY, provider)
    batch.ApplySettings settings, units
    batch.AddCombination "G2_COMPRESSION_EXT", -2000# * 9806.65, 0#, 0#, "Group2", "compression over SLS diagram"
    batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
    batch.Execute

    AssertTrue stats, "batch.group2.extension.compression.directFail", batch.DirectStateStatus(1) = "FAIL"
    AssertTrue stats, "batch.group2.extension.compression.used", batch.ExtensionUsed(1)
    AssertTrue stats, "batch.group2.extension.compression.noCrack", batch.CrackStatus(1) = "N/A"
    AssertTrue stats, "batch.group2.extension.compression.strain", _
        batch.MinConcreteStrain(1) < provider.ConcreteCompressionLimit(cpCrackedNDS)

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "SLS.Crack.Enabled", oldCrackEnabled
    SetSystemSetting "Solver.MaxIterations", oldMaxIterations
    SetSystemSetting "Solver.LoadSteps", oldLoadSteps
    SetSystemSetting "Solver.DirectState.DiagramExtension", oldExtension
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.group2.extension.compression; " & Err.Description
    Resume Restore
End Sub

' Проверяет неосевую перегрузку. Builder должен использовать форму
' деформаций первой несошедшейся попытки, поэтому warm-start остается
' применимым и при одновременных N + Mx + My, а не только при чистом N.
Private Sub TestGroup2BendingBeyondPhysicalLimitUsesExtension(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldMaxIterations As String
    Dim oldLoadSteps As String
    Dim oldExtension As String
    oldMode = GetSystemSetting("Calculation.Mode")
    oldCrackEnabled = GetSystemSetting("SLS.Crack.Enabled")
    oldMaxIterations = GetSystemSetting("Solver.MaxIterations")
    oldLoadSteps = GetSystemSetting("Solver.LoadSteps")
    oldExtension = GetSystemSetting("Solver.DirectState.DiagramExtension")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Calculation.Mode", "DirectState"
    SetSystemSetting "SLS.Crack.Enabled", "Yes"
    SetSystemSetting "Solver.DirectState.DiagramExtension", "Yes"
    SetSystemSetting "Solver.MaxIterations", "100"
    SetSystemSetting "Solver.LoadSteps", "1"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.Initialize settings, units

    Dim referenceX As Double
    Dim referenceY As Double
    Dim batch As CBatchSectionCalculator
    Set batch = BuildUserLShapeTensionBatch(referenceX, referenceY, provider)
    batch.ApplySettings settings, units
    batch.AddCombination "G2_BENDING_EXT", -120# * 9806.65, 420# * 9806.65 * 1000#, _
        -180# * 9806.65 * 1000#, "Group2", "bending over SLS diagram"
    batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
    batch.Execute

    AssertTrue stats, "batch.group2.extension.bending.directFail", batch.DirectStateStatus(1) = "FAIL"
    AssertTrue stats, "batch.group2.extension.bending.used", batch.ExtensionUsed(1)
    AssertTrue stats, "batch.group2.extension.bending.noCrack", batch.CrackStatus(1) = "N/A"
    AssertTrue stats, "batch.group2.extension.bending.curvature", _
        Abs(batch.KappaX(1)) > 0.000000001 Or Abs(batch.KappaY(1)) > 0.000000001

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "SLS.Crack.Enabled", oldCrackEnabled
    SetSystemSetting "Solver.MaxIterations", oldMaxIterations
    SetSystemSetting "Solver.LoadSteps", oldLoadSteps
    SetSystemSetting "Solver.DirectState.DiagramExtension", oldExtension
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.group2.extension.bending; " & Err.Description
    Resume Restore
End Sub

' Проверяет, что определяющее сочетание по прочности выбирается по минимальному
' запасу Capacity среди всех LC, включая строки второй группы. Это нужно для
' Plot/AutoCAD = Worst после удаления пользовательского StrainSafetyFactor.
Private Sub TestBatchGoverningCanUseGroup2CapacitySafety(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "G1_SAFE", -120000#, -1200000#, -600000#, "Group1", "first group"
    batch.AddCombination "G2_GOV", -120000#, -5200000#, -2600000#, "Group2", "second group controls capacity"
    batch.Execute

    AssertTrue stats, "batch.governing.group2.capacity.order", _
        batch.LambdaCapacity(2) > 0# And batch.LambdaCapacity(2) < batch.LambdaCapacity(1)
    AssertTrue stats, "batch.governing.group2.id", batch.GoverningCombinationID = "G2_GOV"
End Sub

' Проверяет новую настройку Capacity.CalculationScope. При Group1Only
' предельный момент считается только для первой группы; Group2 остается
' доступной для прямого НДС и трещин, но CapacityStatus получает N/A.
Private Sub TestCapacityScopeGroup1OnlySkipsGroup2(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldCapacityScope As String
    oldMode = GetSystemSetting("Calculation.Mode")
    oldCrackEnabled = GetSystemSetting("SLS.Crack.Enabled")
    oldCapacityScope = GetSystemSetting("Capacity.CalculationScope")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Calculation.Mode", "FullCapacity"
    SetSystemSetting "SLS.Crack.Enabled", "No"
    SetSystemSetting "Capacity.CalculationScope", "Group1Only"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "G1", -120000#, -1200000#, -600000#, "Group1", "first group"
    batch.AddCombination "G2", -120000#, -5200000#, -2600000#, "Group2", "second group"
    batch.Execute

    AssertTrue stats, "batch.capacity.scope.group1.runs", batch.CapacityStatus(1) <> "N/A"
    AssertTrue stats, "batch.capacity.scope.group2.skipped", batch.CapacityStatus(2) = "N/A"

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "SLS.Crack.Enabled", oldCrackEnabled
    SetSystemSetting "Capacity.CalculationScope", oldCapacityScope
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.capacity.scope; " & Err.Description
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

' Проверяет осевое сжатие через бетонный центр тяжести.
' Для симметричных сечений это состояние не должно создавать кривизну. Для
' Г-сечения с несимметричной арматурой отдельная проверка ниже фиксирует именно
' выбор бетонного центра, а не нулевую кривизну как физическое требование.
Private Sub TestAxialReferenceRemovesPureCompressionEccentricity(ByRef stats As TBatchTestStats)
    CheckPureCompressionReference stats, "circle", CircleGeometry(300#, 125#, -75#), _
        CircleRebars(300#, 125#, -75#, 40#, 12, 20#), 25#, 0.00000001
    CheckPureCompressionReference stats, "rounded", RoundedRectangleGeometry(360#, 240#), _
        RectangleRebars(RoundedRectangleGeometry(360#, 240#)), 30#, 0.00000001
    CheckLShapeReferenceUsesConcreteCentroid stats
End Sub

' Проверяет осевое растяжение как отдельный физический сценарий:
' при приложении N через бетонный центр симметричного сечения кривизны должны
' быть нулевыми, а при смещении той же силы появляется изгиб. Этот контроль
' остается в тестах, чтобы рабочий алгоритм трещин не использовал кривизны как
' fallback-критерий.
Private Sub TestAxialTensionReferenceAndEccentricity(ByRef stats As TBatchTestStats)
    CheckPureTensionReference stats, "circle", CircleGeometry(300#, 125#, -75#), _
        CircleRebars(300#, 125#, -75#, 40#, 12, 20#), 25#, 0.00000001
    CheckPureTensionReference stats, "rounded", RoundedRectangleGeometry(360#, 240#), _
        RectangleRebars(RoundedRectangleGeometry(360#, 240#)), 30#, 0.00000001
End Sub

Private Sub CheckPureCompressionReference(ByRef stats As TBatchTestStats, ByVal caseName As String, _
        ByVal geom As ISectionGeometry, ByVal rebars As CRebarLayout, ByVal meshStep As Double, ByVal tolerance As Double)
    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, meshStep, meshStep, 1, 1

    Dim concrete As CMaterialDiagram
    Set concrete = ProvisionalConcrete()
    Dim steel As CMaterialDiagram
    Set steel = ProvisionalSteel()
    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, rebars)

    Dim refX As Double
    Dim refY As Double
    CalculateConcreteSectionCentroid section, refX, refY

    Dim nValue As Double
    nValue = -100000#
    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    solver.LoadSteps = 1
    solver.MaxIterations = 40
    solver.Solve section, concrete, steel, nValue, nValue * refY, nValue * refX

    AssertTrue stats, "batch.reference." & caseName & ".converged", solver.Converged
    AssertClose stats, "batch.reference." & caseName & ".kappaX", solver.KappaX, 0#, tolerance
    AssertClose stats, "batch.reference." & caseName & ".kappaY", solver.KappaY, 0#, tolerance
End Sub

' Проверяет, что для несимметричного Г-сечения новый reference point берется
' от бетонной части. Это важнее, чем требовать нулевую кривизну: при
' несимметричной арматуре бетонный и приведенный центры могут не совпадать.
Private Sub CheckLShapeReferenceUsesConcreteCentroid(ByRef stats As TBatchTestStats)
    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh LShapeGeometry(250#, 550#, 600#, 250#), 50#, 50#, 1, 1

    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, _
        LShapeRebars(250#, 550#, 600#, 250#, 40#, 50, 32#))

    Dim concreteX As Double
    Dim concreteY As Double
    CalculateConcreteSectionCentroid section, concreteX, concreteY

    Dim transformedX As Double
    Dim transformedY As Double
    CalculateTransformedSectionCentroid section, ProvisionalConcrete(), ProvisionalSteel(), transformedX, transformedY

    AssertTrue stats, "batch.reference.lshape.concreteCenter.exists", Abs(concreteX) + Abs(concreteY) > 0.000001
    AssertTrue stats, "batch.reference.lshape.centerDifference", _
        Abs(concreteX - transformedX) > 0.000001 Or Abs(concreteY - transformedY) > 0.000001
End Sub

Private Sub CheckPureTensionReference(ByRef stats As TBatchTestStats, ByVal caseName As String, _
        ByVal geom As ISectionGeometry, ByVal rebars As CRebarLayout, ByVal meshStep As Double, ByVal tolerance As Double, _
        Optional ByVal eccentricOffsetX As Double = -25#, Optional ByVal eccentricOffsetY As Double = 40#)
    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, meshStep, meshStep, 1, 1

    Dim concrete As CMaterialDiagram
    Set concrete = ProvisionalConcrete()
    Dim steel As CMaterialDiagram
    Set steel = ProvisionalSteel()
    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, rebars)

    Dim refX As Double
    Dim refY As Double
    CalculateConcreteSectionCentroid section, refX, refY

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

' Проверяет DirectState после отказа от пользовательского запаса по деформациям.
' В этом режиме должен быть статус фактического НДС, а capacity остается N/A.
Private Sub TestDirectStateReportsSectionStatus(ByRef stats As TBatchTestStats)
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
    AssertTrue stats, "batch.direct.capacity.na", batch.CapacityStatus(1) = "N/A" And batch.CapacityStatus(2) = "N/A"
    AssertTrue stats, "batch.direct.crack.na", StrComp(batch.CrackStatus(1), "N/A", vbTextCompare) = 0
    AssertTrue stats, "batch.direct.state.status", Len(batch.DirectStateStatus(1)) > 0 And Len(batch.DirectStateStatus(2)) > 0
    AssertTrue stats, "batch.direct.governing.none", Len(batch.GoverningCombinationID) = 0 Or batch.GoverningCombinationID = "DS_SAFE"

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "SLS.Crack.Enabled", oldCrackEnabled
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.direct.sectionStatus; " & Err.Description
    Resume Restore
End Sub

' Проверяет режим CapacityOnly: batch ищет только несущую способность и не
' создает прямое НДС. Это ускоренный сценарий для оценки запаса, поэтому
' поэлементные Stress/Strain и расчет трещин должны быть недоступны.
Private Sub TestCapacityOnlySkipsDirectStateAndCrack(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    oldMode = GetSystemSetting("Calculation.Mode")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Calculation.Mode", "CapacityOnly"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "CAP_ONLY", -150000#, -3000000#, 0#, "Group1", "capacity only"
    batch.Execute

    AssertTrue stats, "batch.capacityOnly.mode", batch.CalculationMode = "CapacityOnly"
    AssertTrue stats, "batch.capacityOnly.capacityRuns", batch.CapacityStatus(1) <> "N/A"
    AssertTrue stats, "batch.capacityOnly.directNA", batch.DirectStateStatus(1) = "N/A"
    AssertTrue stats, "batch.capacityOnly.crackNA", batch.CrackStatus(1) = "N/A"
    AssertTrue stats, "batch.capacityOnly.noState", Not batch.StateAvailable(1) And batch.StateAvailableCount = 0
    AssertTrue stats, "batch.capacityOnly.noExtension", Not batch.ExtensionUsed(1)

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.capacityOnly; " & Err.Description
    Resume Restore
End Sub

' Проверяет, что численная несходимость прямого НДС превращается в короткий
' пользовательский статус NumFail и не порождает расчет трещин.
Private Sub TestDirectStateReportsNumericalFailure(ByRef stats As TBatchTestStats)
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

    AssertTrue stats, "batch.direct.failure.status", batch.Status(1) = "NumFail" Or batch.Status(1) = "FAIL"
    AssertTrue stats, "batch.direct.failure.directStatus", batch.DirectStateStatus(1) = "NumFail" Or batch.DirectStateStatus(1) = "FAIL"
    AssertTrue stats, "batch.direct.failure.crack", batch.CrackStatus(1) = "N/A"

Restore:
    SetSystemSetting "Calculation.Mode", oldMode
    SetSystemSetting "SLS.Crack.Enabled", oldCrackEnabled
    SetSystemSetting "Solver.MaxIterations", oldMaxIterations
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.direct.failure.status; " & Err.Description
    Resume Restore
End Sub

' Проверяет проверку продольных трещин: она должна брать максимальное
' сжимающее напряжение бетона из уже найденного прямого НДС Group2 и не
' требовать отдельного расчетного purpose или повторного запуска solver-а.
Private Sub TestLongitudinalCrackCheckUsesDirectStateStress(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "LONG", -90000#, 0#, 0#, "Group2", "longitudinal"
    batch.Execute

    AssertTrue stats, "batch.longCrack.status.finished", _
        batch.LongitudinalCrackStatus(1) = "OK" Or batch.LongitudinalCrackStatus(1) = "FAIL"
    AssertTrue stats, "batch.longCrack.sigma", batch.MaxConcreteCompressionStress(1) > 0#
    AssertClose stats, "batch.longCrack.rbMc2", batch.LongitudinalCrackRbMc2(1), 14.6, 0.000000001
    AssertClose stats, "batch.longCrack.util", batch.LongitudinalCrackUtilization(1), _
        batch.MaxConcreteCompressionStress(1) / batch.LongitudinalCrackRbMc2(1), 0.000000001
End Sub

' Проверяет нормативную область применения: продольные трещины являются
' проверкой трещиностойкости по Group2, поэтому для Group1 этот блок остается
' N/A и не влияет на прочностной результат.
Private Sub TestLongitudinalCrackSkippedForGroup1(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "LONG_G1", -90000#, 0#, 0#, "Group1", "longitudinal group1"
    batch.Execute

    AssertTrue stats, "batch.longCrack.group1.na", batch.LongitudinalCrackStatus(1) = "N/A"
    AssertClose stats, "batch.longCrack.group1.noRbMc2", batch.LongitudinalCrackRbMc2(1), 0#, 0.000000001
    AssertClose stats, "batch.longCrack.group1.noUtil", batch.LongitudinalCrackUtilization(1), 0#, 0.000000001
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
    AssertTrue stats, "batch.invalid.reader.status", batch.Status(1) = "InputErr"
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestBatchSummaryWriter(ByRef stats As TBatchTestStats)
    On Error GoTo Failed

    Dim stage As String
    stage = "BuildBatchCalculator"
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    stage = "AddCombination"
    batch.AddCombination "W1", -180000#, -3500000#, -2500000#, "Group2", "writer"
    stage = "Execute"
    batch.Execute

    stage = "WriteSummary"
    Dim writer As CBatchResultWriter
    Set writer = New CBatchResultWriter
    writer.WriteSummary ThisWorkbook, batch

    stage = "Asserts"
    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim summaryRow As Long
    summaryRow = BatchSummaryStartRow()
    AssertTrue stats, "batch.writer.fixedRow", summaryRow = 1
    AssertTrue stats, "batch.writer.noResultOverlap", summaryRow + ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Rows.Count - 1 < ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.Row
    AssertTrue stats, "batch.writer.rangeSize", ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Rows.Count = 29 And ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Columns.Count >= 55
    AssertTrue stats, "batch.writer.title", CStr(resultsSheet.Cells.Item(summaryRow, 1).Value2) = "Сводка пакетного расчета (Подробнее)"
    AssertTrue stats, "batch.writer.titleNotMerged", Not resultsSheet.Cells.Item(summaryRow, 1).MergeCells
    AssertTrue stats, "batch.writer.titleHyperlink", resultsSheet.Cells.Item(summaryRow, 1).Hyperlinks.Count > 0
    AssertTrue stats, "batch.writer.governing", CStr(resultsSheet.Cells.Item(summaryRow + 1, 5).Value2) = "W1"
    AssertTrue stats, "batch.writer.crackGoverning.row", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 3, 1).Value2), "трещинам", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.loadPoint.relativeLabel", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 5, 1).Value2), "бетонного сечения", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.loadPoint.zeroX", CStr(resultsSheet.Cells.Item(summaryRow + 5, 5).Value2) = "X=0 mm"
    AssertTrue stats, "batch.writer.subheader.psi", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 7, 42).Value2), "psi", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.capacityFormula.simple", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 9, 26).Formula), "IFERROR", vbTextCompare) > 0 And _
        InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 9, 26).Formula), "IF(", vbTextCompare) = 0
    AssertTrue stats, "batch.writer.header.directStatus", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 8, 5).Value2), "DirectStateStatus", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.header.capacityStatus", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 8, 16).Value2), "CapacityStatus", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.header.capacityPath", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 8, 18).Value2), "CapacityLoadPath", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.header.capacitySolutionMethod", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 8, 19).Value2), "CapacitySolutionMethod", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.header.capacitySafety", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 8, 26).Value2), "CapacitySafetyFactor", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.header.longitudinalCrackStatus", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 8, 51).Value2), "LongitudinalCrackStatus", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.header.longitudinalCrackSafety", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 8, 54).Value2), "LongitudinalCrackSafetyFactor", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.header.overall", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 8, 55).Value2), "MinSafetyFactor", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.longitudinalFormula", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 9, 54).Formula), "IFERROR", vbTextCompare) > 0 And _
        InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 9, 54).Formula), "BA", vbTextCompare) > 0 And _
        InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 9, 54).Formula), "AZ", vbTextCompare) > 0
    Exit Sub

Failed:
    AppendLine stats, "FAIL-TRACE: TestBatchSummaryWriter." & stage & _
        "; err=" & CStr(Err.Number) & "; " & Err.Description
    Err.Raise Err.Number, Err.Source, "TestBatchSummaryWriter." & stage & ": " & Err.Description
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
    batch.Initialize section, TestMaterialProvider()
    Set BuildBatchCalculator = batch
End Function

' Собирает Г-сечение из пользовательского примера: H1/B1/H2/B2 = 550/250/250/600,
' арматура Ø32 по всем внешним и внутренним граням. Этот сценарий нужен именно
' для проверки StateSolution при почти предельном осевом растяжении Group2.
Private Function BuildUserLShapeTensionBatch(ByRef referenceX As Double, ByRef referenceY As Double, _
        Optional ByVal providerOverride As CMaterialModelProvider = Nothing) As CBatchSectionCalculator
    Dim geom As ISectionGeometry
    Set geom = LShapeGeometry(250#, 550#, 600#, 250#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 50#, 50#, 1

    Dim rebars As CRebarLayout
    Set rebars = UserLShapeTensionRebars()

    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, rebars, "UserLShapeTension")

    Dim provider As CMaterialModelProvider
    If providerOverride Is Nothing Then
        Set provider = TestMaterialProvider()
    Else
        Set provider = providerOverride
    End If
    CalculateConcreteSectionCentroid section, referenceX, referenceY

    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize section, provider
    Set BuildUserLShapeTensionBatch = batch
End Function

Private Function UserLShapeTensionRebars() As CRebarLayout
    Dim builder As CLShapeRebarLayoutBuilder
    Set builder = New CLShapeRebarLayoutBuilder
    Set UserLShapeTensionRebars = builder.Build(250#, 550#, 600#, 250#, 0#, 0#, _
        LShapeFaceSettingsForTest(5, 5), _
        LShapeFaceSettingsForTest(2, 2), _
        LShapeFaceSettingsForTest(2, 2), _
        LShapeFaceSettingsForTest(5, 5), _
        "A400")
End Function

Private Function LShapeFaceSettingsForTest(ByVal count1 As Long, ByVal count2 As Long) As Variant
    LShapeFaceSettingsForTest = Array(40#, 40#, 32#, 32#, count1, count2, 80#, 80#, 80#, 80#, _
        0#, 0#, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")
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

Private Function ProvisionalConcrete() As CMaterialDiagram
    Dim concrete As CMaterialDiagram
    Set concrete = New CMaterialDiagram
    concrete.Initialize -0.0015, -15.5, -0.0035, -15.5
    Set ProvisionalConcrete = concrete
End Function

Private Function TestMaterialProvider() As CMaterialModelProvider
    Dim concreteParameters As CConcreteMaterialParameters
    Set concreteParameters = New CConcreteMaterialParameters
    concreteParameters.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6

    Dim steelParameters As CSteelMaterialParameters
    Set steelParameters = New CSteelMaterialParameters
    steelParameters.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#, "Ribbed"

    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.InitializeFromParameters concreteParameters, steelParameters
    Set TestMaterialProvider = provider
End Function

Private Function ProvisionalSteel() As CMaterialDiagram
    Dim steel As CMaterialDiagram
    Set steel = New CMaterialDiagram
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





