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

Private Const TEST_TF_M_IN_NMM As Double = 9806650# ' 1 tf*m во внутренних Н*мм.
Private Const TEST_DEFAULT_ZERO_MOMENT_PER_DEPTH As Double = 4903.325 ' 0.5 tf*m/m = 4903.325 Н*мм/мм.

' Запускает связанный набор операций и возвращает пользователю итоговый статус выполнения.
Public Function RunBatchCalculationTests() As String
    On Error GoTo Failed

    Dim stats As TBatchTestStats
    Dim t0 As Double
    t0 = Timer

    AppendLine stats, "RUN: TestBatchOneCombination"
    TestBatchOneCombination stats
    AppendLine stats, "RUN: TestProfileIdControlsLimitStateGroup"
    TestProfileIdControlsLimitStateGroup stats
    AppendLine stats, "RUN: TestBatchFiveCombinations"
    TestBatchFiveCombinations stats
    AppendLine stats, "RUN: TestBatchGoverningUsesLowestSafetyFactor"
    TestBatchGoverningUsesLowestSafetyFactor stats
    AppendLine stats, "RUN: TestBatchGoverningUsesStrengthProfilesOnly"
    TestBatchGoverningUsesStrengthProfilesOnly stats
    AppendLine stats, "RUN: TestBatchPureAxialCapacityUsesNult"
    TestBatchPureAxialCapacityUsesNult stats
    AppendLine stats, "RUN: TestBatchEccentricAxialCapacityTriesUltimateStrain"
    TestBatchEccentricAxialCapacityTriesUltimateStrain stats
    AppendLine stats, "RUN: TestBatchRectSetN200CapacityPathNDoesNotNumFail"
    TestBatchRectSetN200CapacityPathNDoesNotNumFail stats
    AppendLine stats, "RUN: TestBatchExplicitCapacityLoadPathScalesMxy"
    TestBatchExplicitCapacityLoadPathScalesMxy stats
    AppendLine stats, "RUN: TestBatchExplicitCapacityLoadPathScalesNWithMoments"
    TestBatchExplicitCapacityLoadPathScalesNWithMoments stats
    AppendLine stats, "RUN: TestCapacityLoadPathNReportsUltimateAtLoadPoint"
    TestCapacityLoadPathNReportsUltimateAtLoadPoint stats
    AppendLine stats, "RUN: TestBatchCapacityLoadPathVariants"
    TestBatchCapacityLoadPathVariants stats
    AppendLine stats, "RUN: TestBatchCapacityLoadPathAllowsZeroInactiveComponents"
    TestBatchCapacityLoadPathAllowsZeroInactiveComponents stats
    AppendLine stats, "RUN: TestBatchNMxyWithoutMomentsUsesStableForcePath"
    TestBatchNMxyWithoutMomentsUsesStableForcePath stats
    AppendLine stats, "RUN: TestMomentZeroFilterDefaultThresholds"
    TestMomentZeroFilterDefaultThresholds stats
    AppendLine stats, "RUN: TestCapacityLoadPathFiltersEngineeringSmallMoments"
    TestCapacityLoadPathFiltersEngineeringSmallMoments stats
    AppendLine stats, "RUN: TestBatchZeroMomentFilterNormalizesCapacityPath"
    TestBatchZeroMomentFilterNormalizesCapacityPath stats
    AppendLine stats, "RUN: TestStabilityZeroMomentFilterUsesZeroMomentSigns"
    TestStabilityZeroMomentFilterUsesZeroMomentSigns stats
    AppendLine stats, "RUN: TestBatchInvalidCapacityLoadPathReportsInputErr"
    TestBatchInvalidCapacityLoadPathReportsInputErr stats
    AppendLine stats, "RUN: TestPR2SkipsCapacityByProfile"
    TestPR2SkipsCapacityByProfile stats
    AppendLine stats, "RUN: TestLoadReferenceTransformsUserMoments"
    TestLoadReferenceTransformsUserMoments stats
    AppendLine stats, "RUN: TestSectionLoadStateTransfersMoments"
    TestSectionLoadStateTransfersMoments stats
    AppendLine stats, "RUN: TestAxialReferenceRemovesPureCompressionEccentricity"
    TestAxialReferenceRemovesPureCompressionEccentricity stats
    AppendLine stats, "RUN: TestAxialTensionReferenceAndEccentricity"
    TestAxialTensionReferenceAndEccentricity stats
    AppendLine stats, "RUN: TestDirectStateReportsSectionStatus"
    TestDirectStateReportsSectionStatus stats
    AppendLine stats, "RUN: TestPR1RectSetSmallTensionMomentDirectStateDoesNotNumFail"
    TestPR1RectSetSmallTensionMomentDirectStateDoesNotNumFail stats
    AppendLine stats, "RUN: TestLongitudinalCrackCheckUsesDirectStateStress"
    TestLongitudinalCrackCheckUsesDirectStateStress stats
    AppendLine stats, "RUN: TestLongitudinalCrackSkippedForPR1"
    TestLongitudinalCrackSkippedForPR1 stats
    AppendLine stats, "RUN: TestStabilityProfileEnablesSP63ForCompression"
    TestStabilityProfileEnablesSP63ForCompression stats
    AppendLine stats, "RUN: TestStabilitySkippedForTension"
    TestStabilitySkippedForTension stats
    AppendLine stats, "RUN: TestStabilitySP35UsesConfigTableAndProfileValueSet"
    TestStabilitySP35UsesConfigTableAndProfileValueSet stats
    AppendLine stats, "RUN: TestStabilitySP35TableSeparatesNcrAndNult"
    TestStabilitySP35TableSeparatesNcrAndNult stats
    AppendLine stats, "RUN: TestStabilitySP35TableIgnoresPhiL2"
    TestStabilitySP35TableIgnoresPhiL2 stats
    AppendLine stats, "RUN: TestStabilitySP35TableRebarAreaCorrection"
    TestStabilitySP35TableRebarAreaCorrection stats
    AppendLine stats, "RUN: TestStabilitySP35TableInterpolationIntermediate"
    TestStabilitySP35TableInterpolationIntermediate stats
    AppendLine stats, "RUN: TestStabilitySP35TableBoundaryReservePasses"
    TestStabilitySP35TableBoundaryReservePasses stats
    AppendLine stats, "RUN: TestStabilityFailStopsDownstream"
    TestStabilityFailStopsDownstream stats
    AppendLine stats, "RUN: TestStabilitySP63ShortSlendernessEtaIsOne"
    TestStabilitySP63ShortSlendernessEtaIsOne stats
    AppendLine stats, "RUN: TestStabilityInvalidMuReportsInputErr"
    TestStabilityInvalidMuReportsInputErr stats
    AppendLine stats, "RUN: TestStabilityInvalidSettingsReportInputErr"
    TestStabilityInvalidSettingsReportInputErr stats
    AppendLine stats, "RUN: TestStabilitySP35MixedBranchReportsMixed"
    TestStabilitySP35MixedBranchReportsMixed stats
    AppendLine stats, "RUN: TestStabilityCircleMxDoesNotCreateMy"
    TestStabilityCircleMxDoesNotCreateMy stats
    AppendLine stats, "RUN: TestStabilityAccidentalBothPlanes"
    TestStabilityAccidentalBothPlanes stats
    AppendLine stats, "RUN: TestStabilityUsesTransformedCentroidForEccentricity"
    TestStabilityUsesTransformedCentroidForEccentricity stats
    AppendLine stats, "RUN: TestStabilityAccidentalUsesGeometricLength"
    TestStabilityAccidentalUsesGeometricLength stats
    AppendLine stats, "RUN: TestStabilityAccidentalUserMode"
    TestStabilityAccidentalUserMode stats
    AppendLine stats, "RUN: TestStabilitySP35UsesCoreDistanceNotRadius"
    TestStabilitySP35UsesCoreDistanceNotRadius stats
    AppendLine stats, "RUN: TestStabilityDepthUsesConcreteContourOnly"
    TestStabilityDepthUsesConcreteContourOnly stats
    AppendLine stats, "RUN: TestStabilityPhiLUsesSignedSustainedMoment"
    TestStabilityPhiLUsesSignedSustainedMoment stats
    AppendLine stats, "RUN: TestStabilitySustainedNNotClamped"
    TestStabilitySustainedNNotClamped stats
    AppendLine stats, "RUN: TestStabilitySP35OppositeMomentSigns"
    TestStabilitySP35OppositeMomentSigns stats
    AppendLine stats, "RUN: TestBatchSummaryWritesOnlySelectedStabilityCode"
    TestBatchSummaryWritesOnlySelectedStabilityCode stats
    AppendLine stats, "RUN: TestPR1RunsStrengthWithoutCrackWidth"
    TestPR1RunsStrengthWithoutCrackWidth stats
    AppendLine stats, "RUN: TestDirectStateReportsNumericalFailure"
    TestDirectStateReportsNumericalFailure stats
    AppendLine stats, "RUN: TestPR2PhysicalStateRunsCrackWithExtensionEnabled"
    TestPR2PhysicalStateRunsCrackWithExtensionEnabled stats
    AppendLine stats, "RUN: TestBatchCrackCoverDistanceModeChangesAs"
    TestBatchCrackCoverDistanceModeChangesAs stats
    AppendLine stats, "RUN: TestPR2AutoCrackStoresBeforeAndAfterMcrcStates"
    TestPR2AutoCrackStoresBeforeAndAfterMcrcStates stats
    AppendLine stats, "RUN: TestPR2AutoCrackPureBendingStoresMcrcStates"
    TestPR2AutoCrackPureBendingStoresMcrcStates stats
    AppendLine stats, "RUN: TestPR2RectSetCompressionSmallMomentCrackDoesNotNumFail"
    TestPR2RectSetCompressionSmallMomentCrackDoesNotNumFail stats
    AppendLine stats, "RUN: TestCrackInitiationLoadPathsWriteFormationSummary"
    TestCrackInitiationLoadPathsWriteFormationSummary stats
    AppendLine stats, "RUN: TestCrackAutoFormationPathSwitchesForRectSet"
    TestCrackAutoFormationPathSwitchesForRectSet stats
    AppendLine stats, "RUN: TestPR1AxialTensionBeyondPhysicalLimitUsesExtension"
    TestPR1AxialTensionBeyondPhysicalLimitUsesExtension stats
    AppendLine stats, "RUN: TestPR1AxialTensionNearLimitDoesNotJumpToNumFail"
    TestPR1AxialTensionNearLimitDoesNotJumpToNumFail stats
    AppendLine stats, "RUN: TestPR1AxialCompressionNearLimitDoesNotJumpToNumFail"
    TestPR1AxialCompressionNearLimitDoesNotJumpToNumFail stats
    AppendLine stats, "RUN: TestPR1AxialTensionProgressionAfterLimitIsStableFail"
    TestPR1AxialTensionProgressionAfterLimitIsStableFail stats
    AppendLine stats, "RUN: TestPR1AxialCompressionProgressionAfterLimitIsStableFail"
    TestPR1AxialCompressionProgressionAfterLimitIsStableFail stats
    AppendLine stats, "RUN: TestPR2AxialTensionBeyondPhysicalLimitUsesExtension"
    TestPR2AxialTensionBeyondPhysicalLimitUsesExtension stats
    AppendLine stats, "RUN: TestPR2AxialCompressionBeyondPhysicalLimitUsesExtension"
    TestPR2AxialCompressionBeyondPhysicalLimitUsesExtension stats
    AppendLine stats, "RUN: TestPR2BendingBeyondPhysicalLimitUsesExtension"
    TestPR2BendingBeyondPhysicalLimitUsesExtension stats
    AppendLine stats, "RUN: TestBatchMoreThanTwentyCombinations"
    TestBatchMoreThanTwentyCombinations stats
    AppendLine stats, "RUN: TestInvalidCombinationFromNamedRange"
    TestInvalidCombinationFromNamedRange stats
    AppendLine stats, "RUN: TestBatchSummaryWriter"
    TestBatchSummaryWriter stats
    AppendLine stats, "RUN: TestBatchSummaryRowsUseAvailableLoadRange"
    TestBatchSummaryRowsUseAvailableLoadRange stats
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

' ДЛЯ ТЕСТОВ: запускает только сверку верхнего batch summary с подробными
' блоками Results. Нужен как быстрый диагностический вход, когда общий набор
' batch-тестов слишком долгий и надо проверить именно новый пользовательский
' контракт сводки запасов.
Public Function RunBatchSummaryReserveConsistencyTest() As String
    On Error GoTo Failed

    Dim stats As TBatchTestStats
    AppendLine stats, "RUN: TestBatchSummaryWriter"
    TestBatchSummaryWriter stats
    AppendLine stats, "RUN: TestStabilitySP35TableBoundaryReservePasses"
    TestStabilitySP35TableBoundaryReservePasses stats
    AppendLine stats, "TOTAL_BATCH_SUMMARY_RESERVE: passed=" & CStr(stats.Passed) & _
        "; failed=" & CStr(stats.Failed)
    RunBatchSummaryReserveConsistencyTest = stats.Report
    Exit Function

Failed:
    RunBatchSummaryReserveConsistencyTest = stats.Report & "RUNTIME ERROR: " & _
        CStr(Err.Number) & "; source=" & Err.Source & "; description=" & Err.Description
End Function

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestBatchOneCombination(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "C1", -220000#, -7000000#, -5000000#, "PR1", "single"
    batch.Execute

    AssertTrue stats, "batch.one.count", batch.Count = 1
    AssertTrue stats, "batch.one.governing", batch.GoverningCombinationID = "C1"
    AssertTrue stats, "batch.one.capacity.status", Len(batch.CapacityStatus(1)) > 0
    AssertTrue stats, "batch.one.crack.status", Len(batch.CrackStatus(1)) > 0
    AssertTrue stats, "batch.one.elapsed", batch.ElapsedSeconds >= 0#
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestProfileIdControlsLimitStateGroup(ByRef stats As TBatchTestStats)
    Dim group1 As CBatchSectionCalculator
    Set group1 = BuildBatchCalculator()
    group1.AddCombination "G1", -220000#, -7000000#, -5000000#, "PR1", "strength"
    group1.Execute

    AssertTrue stats, "batch.profileId.group1.capacity", group1.LambdaCapacity(1) > 0#
    AssertTrue stats, "batch.profileId.group1.noCrack", group1.CrackStatus(1) = "N/A"

    Dim group2 As CBatchSectionCalculator
    Set group2 = BuildBatchCalculator()
    group2.AddCombination "G2", -220000#, -7000000#, -5000000#, "PR2", "crack"
    group2.Execute

    AssertTrue stats, "batch.profileId.group2.noCapacity", group2.CapacityStatus(1) = "N/A"
    AssertTrue stats, "batch.profileId.group2.crack", group2.CrackStatus(1) <> "N/A"
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestBatchCapacityUsesSystemSettings(ByRef stats As TBatchTestStats)
    Dim oldStrategy As String
    Dim oldMaxLambda As String
    oldStrategy = GetSystemSetting("Capacity.SolutionStrategy")
    oldMaxLambda = GetSystemSetting("Capacity.MaxLambda")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Capacity.SolutionStrategy", "LoadMultiplier"
    SetSystemSetting "Capacity.MaxLambda", "0.5"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "LIMITED", -220000#, -7000000#, -5000000#, "PR1", "max-lambda"
    batch.Execute

    AssertTrue stats, "batch.settings.capacity.maxLambda", batch.CapacityStatus(1) = "NumFail"

Restore:
    SetSystemSetting "Capacity.SolutionStrategy", oldStrategy
    SetSystemSetting "Capacity.MaxLambda", oldMaxLambda
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.settings.capacity.maxLambda; " & Err.Description
    Resume Restore
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestInvalidModeSettingsAreNotFallbacks(ByRef stats As TBatchTestStats)
    Dim oldCapacitySolutionStrategy As String
    oldCapacitySolutionStrategy = GetSystemSetting("Capacity.SolutionStrategy")

    On Error GoTo RestoreAndFail

    SetSystemSetting "Capacity.SolutionStrategy", "WrongCapacity"
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "BAD_CAP", -220000#, -7000000#, -5000000#, "PR1", "wrong capacity"
    batch.Execute
    AssertTrue stats, "batch.invalid.CapacitySolutionStrategy.status", batch.Status(1) = "InputErr"
    AssertTrue stats, "batch.invalid.CapacitySolutionStrategy.noLambda", batch.LambdaCapacity(1) = 0#

Restore:
    SetSystemSetting "Capacity.SolutionStrategy", oldCapacitySolutionStrategy
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.invalid.solutionStrategy; " & Err.Description
    Resume Restore
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestBatchFiveCombinations(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()

    Dim i As Long
    For i = 1 To 5
        batch.AddCombination "C" & CStr(i), -150000# - 10000# * i, -3000000# - 250000# * i, _
            -2000000# - 200000# * i, "PR1", "five-" & CStr(i)
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
    batch.AddCombination "SAFE", -150000#, -1000000#, -500000#, "PR1", "larger safety"
    batch.AddCombination "GOV", -150000#, -7000000#, -3500000#, "PR1", "smaller safety"
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
    batch.AddCombination "N_ONLY", 50000#, 0#, 0#, "PR1", "pure axial"
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

' Проверяет, что "только N" не считается чистой осевой траекторией, если
' после переноса в центр приведенного сечения остается изгибающий момент.
' Путь lambda*N сохраняется, но стратегия Auto может пробовать UltimateStrain.
Private Sub TestBatchEccentricAxialCapacityTriesUltimateStrain(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplyLoadReference 40#, -25#
    batch.AddCombination "N_ECC", -150000#, 0#, 0#, "PR1", "eccentric axial", ChrW$(&H3BB) & "*N"
    batch.Execute

    AssertTrue stats, "batch.nEcc.path", batch.CapacityLoadPathKey(1) = "LambdaN"
    AssertTrue stats, "batch.nEcc.status", batch.CapacityStatus(1) = "OK" Or batch.CapacityStatus(1) = "FAIL"
    AssertTrue stats, "batch.nEcc.method", batch.CapacitySolutionMethod(1) = "UltimateStrain"
    AssertTrue stats, "batch.nEcc.loadPointMomentsRemainZero", _
        Abs(batch.MxUltimate(1)) < 100000# And Abs(batch.MyUltimate(1)) < 100000#
End Sub

' Проверяет пользовательский сценарий из книги: Г-сечение, нагрузка
' N=-200 тс при принятом знаке +N=Compression, то есть внутреннее растяжение,
' и путь CapacityLoadPath = lambda*N. Точка приложения проходит через бетонный
' центр тяжести, поэтому внутри solver-а вместе с N масштабируются и моменты
' переноса, но пользовательская постановка остается чистым Nult.
Private Sub TestBatchRectSetN200CapacityPathNDoesNotNumFail(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldMethod As String
    Dim oldScope As String
    Dim oldBaseLoadSteps As String
    Dim oldMaxRetries As String
    oldMethod = GetSystemSetting("Capacity.SolutionStrategy")
    oldBaseLoadSteps = GetSystemSetting("Capacity.BaseLoadSteps")
    oldMaxRetries = GetSystemSetting("Capacity.MaxRetries")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Capacity.BaseLoadSteps", "1"
    SetSystemSetting "Capacity.MaxRetries", "0"

    CheckBatchRectSetN200CapacitySolutionStrategy stats, "Auto"
    CheckBatchRectSetN200CapacitySolutionStrategy stats, "UltimateStrain"
    CheckBatchRectSetN200CapacitySolutionStrategy stats, "LoadMultiplier"

Restore:
    SetSystemSetting "Capacity.SolutionStrategy", oldMethod
    SetSystemSetting "Capacity.BaseLoadSteps", oldBaseLoadSteps
    SetSystemSetting "Capacity.MaxRetries", oldMaxRetries
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.rectset.n200.capacityPathN; " & Err.Description
    Resume Restore
End Sub

Private Sub CheckBatchRectSetN200CapacitySolutionStrategy(ByRef stats As TBatchTestStats, ByVal methodName As String)
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
    Set batch = BuildUserRectSetTensionBatch(referenceX, referenceY, provider)
    batch.ApplySettings settings, units
    batch.AddCombination "G1_N200_" & methodName, 200# * 9806.65, 0#, 0#, _
        "PR1", "user N=-200 tf, lambda*N", ChrW$(&H3BB) & "*N"
    batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
    batch.Execute

    AppendLine stats, "INFO: batch.rectset.n200." & methodName & _
        "; capacityStatus=" & batch.CapacityStatus(1) & _
        "; limitState=" & batch.CapacityLimitState(1) & _
        "; solutionMethod=" & batch.CapacitySolutionMethod(1) & _
        "; lambda=" & FormatNumberInvariant(batch.LambdaCapacity(1)) & _
        "; Nult=" & FormatNumberInvariant(batch.NUltimate(1))
    AssertTrue stats, "batch.rectset.n200." & methodName & ".notNumFail", batch.CapacityStatus(1) <> "NumFail"
    AssertTrue stats, "batch.rectset.n200." & methodName & ".capacityStatus", _
        batch.CapacityStatus(1) = "OK" Or batch.CapacityStatus(1) = "FAIL"
    AssertTrue stats, "batch.rectset.n200." & methodName & ".nult", Abs(batch.NUltimate(1)) > Abs(batch.N(1))
    AssertTrue stats, "batch.rectset.n200." & methodName & ".solutionMethod", _
        batch.CapacitySolutionMethod(1) = "LoadMultiplier"
End Sub

' Проверяет, что явный выбор lambda*Mxy масштабирует оба пользовательских
' момента при постоянной продольной силе. Это основной вариант для общего
' изгиба N + Mx + My, когда нужно найти предельный момент при заданной N.
Private Sub TestBatchExplicitCapacityLoadPathScalesMxy(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "M_BRANCH", -150000#, -3000000#, -3000000#, "PR1", "moment branch", ChrW$(&H3BB) & "*Mxy"
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
    batch.AddCombination "N_BRANCH", -150000#, -3000000#, -1000000#, "PR1", "axial branch", ChrW$(&H3BB) & "*N"
    batch.Execute

    AssertTrue stats, "batch.capacityPath.n.value", batch.CapacityLoadPath(1) = ChrW$(&H3BB) & "*N"
    AssertTrue stats, "batch.capacityPath.n.key", batch.CapacityLoadPathKey(1) = "LambdaN"
    AssertTrue stats, "batch.capacityPath.n.nult", Abs(batch.NUltimate(1)) > 0#
    AssertTrue stats, "batch.capacityPath.n.status", batch.CapacityStatus(1) = "OK" Or batch.CapacityStatus(1) = "FAIL" Or batch.CapacityStatus(1) = "NumFail"
End Sub

' Проверяет пользовательский вывод для lambda*N: solver внутри масштабирует
' момент от эксцентриситета N, но предельные Mx/My в batch должны быть
' перенесены обратно в точку приложения нагрузки через CSectionLoadState.
Private Sub TestCapacityLoadPathNReportsUltimateAtLoadPoint(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "N_REF", -150000#, 3000000#, 0#, "PR1", _
        "lambda n at shifted load point", ChrW$(&H3BB) & "*N"
    batch.ApplyLoadReference 40#, -25#
    batch.Execute

    AssertTrue stats, "batch.capacityPath.n.ref.status", _
        batch.CapacityStatus(1) = "OK" Or batch.CapacityStatus(1) = "FAIL"
    AssertClose stats, "batch.capacityPath.n.ref.mxAtLoadPoint", _
        batch.MxUltimate(1), batch.UserMx(1), 100000#
    AssertClose stats, "batch.capacityPath.n.ref.myAtLoadPoint", _
        batch.MyUltimate(1), batch.UserMy(1), 100000#
    AssertTrue stats, "batch.capacityPath.n.ref.internalDiffers", _
        Abs(batch.CapacityStateMy(1) - batch.MyUltimate(1)) > 100000#
End Sub

' Проверяет все пользовательские варианты CapacityLoadPath. Тест не
' привязывается к конкретной величине запаса: здесь важно, что batch
' корректно распознает путь и не подменяет выбранную пользователем траекторию.
Private Sub TestBatchCapacityLoadPathVariants(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "PATH_MX", -150000#, -3000000#, 0#, "PR1", "lambda mx", ChrW$(&H3BB) & "*Mx"
    batch.AddCombination "PATH_MY", -150000#, 0#, -3000000#, "PR1", "lambda my", ChrW$(&H3BB) & "*My"
    batch.AddCombination "PATH_MXY", -150000#, -3000000#, -1000000#, "PR1", "lambda mxy", ChrW$(&H3BB) & "*Mxy"
    batch.AddCombination "PATH_N", -150000#, -3000000#, -1000000#, "PR1", "lambda n", ChrW$(&H3BB) & "*N"
    batch.AddCombination "PATH_ALL", -150000#, -3000000#, -1000000#, "PR1", "lambda all", ChrW$(&H3BB) & "*NMxy"
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
    batch.AddCombination "MX_NO_N", 0#, -3000000#, 0#, "PR1", "lambda mx without N", ChrW$(&H3BB) & "*Mx"
    batch.AddCombination "MY_NO_N", 0#, 0#, -3000000#, "PR1", "lambda my without N", ChrW$(&H3BB) & "*My"
    batch.AddCombination "MXY_ONLY_MX", 0#, -3000000#, 0#, "PR1", "lambda mxy with only Mx", ChrW$(&H3BB) & "*Mxy"
    batch.AddCombination "MXY_ONLY_MY", 0#, 0#, -3000000#, "PR1", "lambda mxy with only My", ChrW$(&H3BB) & "*Mxy"
    batch.AddCombination "NMXY_ONLY_N", -150000#, 0#, 0#, "PR1", "lambda nmxy with only N", ChrW$(&H3BB) & "*NMxy"
    batch.AddCombination "NMXY_ONLY_MX", 0#, -3000000#, 0#, "PR1", "lambda nmxy with only Mx", ChrW$(&H3BB) & "*NMxy"
    batch.AddCombination "NMXY_ONLY_MY", 0#, 0#, -3000000#, "PR1", "lambda nmxy with only My", ChrW$(&H3BB) & "*NMxy"
    batch.AddCombination "NMXY_EMPTY", 0#, 0#, 0#, "PR1", "lambda nmxy empty", ChrW$(&H3BB) & "*NMxy"
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
    oldMethod = GetSystemSetting("Capacity.SolutionStrategy")

    On Error GoTo RestoreAndFail
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
    Set batch = BuildUserRectSetTensionBatch(referenceX, referenceY, provider)
    batch.ApplySettings settings, units
    batch.AddCombination "NMXY_ZERO_M", 100# * 9806.65, 0#, 0#, _
        "PR1", "lambda NMxy with zero moments", ChrW$(&H3BB) & "*NMxy"
    batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
    batch.Execute

    AssertTrue stats, "batch.capacityPath.nmxyZeroM.path", batch.CapacityLoadPathKey(1) = "LambdaNMxy"
    AssertTrue stats, "batch.capacityPath.nmxyZeroM.solutionMethod", batch.CapacitySolutionMethod(1) = "LoadMultiplier"
    AssertTrue stats, "batch.capacityPath.nmxyZeroM.notNumFail", batch.CapacityStatus(1) <> "NumFail"
    AssertTrue stats, "batch.capacityPath.nmxyZeroM.nult", Abs(batch.NUltimate(1)) > Abs(batch.N(1))

Restore:
    SetSystemSetting "Capacity.SolutionStrategy", oldMethod
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.capacityPath.nmxyZeroM; " & Err.Description
    Resume Restore
End Sub

' Проверяет сам инженерный фильтр практически нулевого момента без solver-а.
' Порог растет линейно с бетонным габаритом h: Mtol = value*h.
Private Sub TestMomentZeroFilterDefaultThresholds(ByRef stats As TBatchTestStats)
    Dim filter As CMomentZeroFilter
    Set filter = New CMomentZeroFilter
    filter.Initialize TEST_DEFAULT_ZERO_MOMENT_PER_DEPTH

    AssertClose stats, "batch.zeroMoment.tinyResidual", filter.NormalizeMoment(0.000000000000000242, 1000#), 0#, 0#
    AssertClose stats, "batch.zeroMoment.h1m.threshold", filter.NormalizeMoment(0.5 * TEST_TF_M_IN_NMM, 1000#), 0#, 0#
    AssertTrue stats, "batch.zeroMoment.h1m.above", _
        Abs(filter.NormalizeMoment(0.5002 * TEST_TF_M_IN_NMM, 1000#)) > 0#
    AssertClose stats, "batch.zeroMoment.h02m.threshold", filter.NormalizeMoment(0.1 * TEST_TF_M_IN_NMM, 200#), 0#, 0#
    AssertClose stats, "batch.zeroMoment.h3m.threshold", filter.NormalizeMoment(1.5 * TEST_TF_M_IN_NMM, 3000#), 0#, 0#
End Sub

' Проверяет, что CapacityLoadPath не пропускает в solver момент, который уже
' меньше инженерного порога Calculation.ZeroMomentPerDepth * h.
Private Sub TestCapacityLoadPathFiltersEngineeringSmallMoments(ByRef stats As TBatchTestStats)
    Dim filter As CMomentZeroFilter
    Set filter = New CMomentZeroFilter
    filter.Initialize 0.001 * TEST_TF_M_IN_NMM

    Dim pathMxy As CCapacityLoadPath
    Dim loadMxy As CSectionLoadState
    Set loadMxy = New CSectionLoadState
    loadMxy.Initialize -100# * 9806.65, 0#, -0.158 * TEST_TF_M_IN_NMM, _
        0#, 0#, filter, 870#, 870#
    Set pathMxy = New CCapacityLoadPath
    pathMxy.InitializeFromLoadState ChrW$(&H3BB) & "*Mxy", loadMxy

    AssertTrue stats, "batch.zeroMoment.pathMxy.noScaledMoment", Not pathMxy.HasScaledLoad
    AssertClose stats, "batch.zeroMoment.pathMxy.myBase", pathMxy.MyBase, 0#, 0#

    Dim pathNMxy As CCapacityLoadPath
    Dim loadNMxy As CSectionLoadState
    Set loadNMxy = New CSectionLoadState
    loadNMxy.Initialize -100# * 9806.65, 0#, -0.158 * TEST_TF_M_IN_NMM, _
        0#, 0#, filter, 870#, 870#
    Set pathNMxy = New CCapacityLoadPath
    pathNMxy.InitializeFromLoadState ChrW$(&H3BB) & "*NMxy", loadNMxy

    AssertTrue stats, "batch.zeroMoment.pathNMxy.forceOnly", pathNMxy.ForceOnly
    AssertClose stats, "batch.zeroMoment.pathNMxy.myBase", pathNMxy.MyBase, 0#, 0#
End Sub

' Проверяет, что малый момент после чтения настроек и единиц не заставляет
' capacity выбирать моментную lambda-траекторию вместо осевой.
Private Sub TestBatchZeroMomentFilterNormalizesCapacityPath(ByRef stats As TBatchTestStats)
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings, units
    batch.AddCombination "M_TINY", -200000#, 0.000000000000000242, 0#, "PR1", "tiny residual moment"
    batch.Execute

    AssertTrue stats, "batch.zeroMoment.capacityPath.axialDefault", batch.CapacityLoadPathKey(1) = "LambdaN"
    AssertTrue stats, "batch.zeroMoment.capacityPath.notInputErr", batch.CapacityStatus(1) <> "InputErr"
    AssertClose stats, "batch.zeroMoment.capacityPath.mx", batch.UserMx(1), 0#, 0#
End Sub

' Проверяет, что Stability в режиме OnlyMomentPlane считает отфильтрованный
' микромомент нулевым и берет направление из ZeroMomentEccentricitySign1/2.
Private Sub TestStabilityZeroMomentFilterUsesZeroMomentSigns(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    Dim oldAccMode As String
    Dim oldAccPlanes As String
    Dim oldSign1 As String
    Dim oldSign2 As String
    Dim oldAccUser1 As String
    Dim oldAccUser2 As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")
    oldAccMode = GetSystemSetting("Stability.AccidentalEccentricityMode")
    oldAccPlanes = GetSystemSetting("Stability.AccidentalEccentricityPlanes")
    oldSign1 = GetSystemSetting("Stability.ZeroMomentEccentricitySign1")
    oldSign2 = GetSystemSetting("Stability.ZeroMomentEccentricitySign2")
    oldAccUser1 = GetSystemSetting("Stability.AccidentalEccentricityUser1")
    oldAccUser2 = GetSystemSetting("Stability.AccidentalEccentricityUser2")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP63"
    SetSystemSetting "Stability.AccidentalEccentricityMode", "User"
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", "OnlyMomentPlane"
    SetSystemSetting "Stability.ZeroMomentEccentricitySign1", "-1"
    SetSystemSetting "Stability.ZeroMomentEccentricitySign2", "-1"
    SetSystemSetting "Stability.AccidentalEccentricityUser1", "1"
    SetSystemSetting "Stability.AccidentalEccentricityUser2", "1"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings, units
    batch.SetSP35Table721 ThisWorkbook.Names.Item("rngSP35Table721").RefersToRange.Value2
    batch.AddCombination "STAB_TINY_M", -120000#, 0.000000000000000242, 0#, "PR1", "tiny stability moment"
    batch.Execute

    AssertTrue stats, "batch.zeroMoment.stability.status", batch.StabilityStatus(1) = "OK" Or batch.StabilityStatus(1) = "FAIL"
    AssertTrue stats, "batch.zeroMoment.stability.sign1", batch.StabilityAccidentalEcc1(1) < 0#
    AssertTrue stats, "batch.zeroMoment.stability.sign2", batch.StabilityAccidentalEcc2(1) < 0#
    AssertClose stats, "batch.zeroMoment.stability.finalUserMx", batch.UserMx(1), 0#, 0#
    AssertClose stats, "batch.zeroMoment.stability.finalUserMy", batch.UserMy(1), 0#, 0#
    AssertClose stats, "batch.zeroMoment.stability.finalSummaryMx", batch.StabilityDesignMx(1), 0#, 0#
    AssertClose stats, "batch.zeroMoment.stability.finalSummaryMy", batch.StabilityDesignMy(1), 0#, 0#
    AssertTrue stats, "batch.zeroMoment.stability.capacityPath", batch.CapacityLoadPathKey(1) = "LambdaN"
    AssertTrue stats, "batch.zeroMoment.stability.capacityNotNumFail", batch.CapacityStatus(1) <> "NumFail"

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    SetSystemSetting "Stability.AccidentalEccentricityMode", oldAccMode
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", oldAccPlanes
    SetSystemSetting "Stability.ZeroMomentEccentricitySign1", oldSign1
    SetSystemSetting "Stability.ZeroMomentEccentricitySign2", oldSign2
    SetSystemSetting "Stability.AccidentalEccentricityUser1", oldAccUser1
    SetSystemSetting "Stability.AccidentalEccentricityUser2", oldAccUser2
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.zeroMoment.stability; " & Err.Description
    Resume Restore
End Sub

' Проверяет, что ошибочный текст CapacityLoadPath не заменяется молча
' авто-выбором. Пользователь должен сразу увидеть ошибку в строке LC.
Private Sub TestBatchInvalidCapacityLoadPathReportsInputErr(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "BAD_CONST", -150000#, -3000000#, 0#, "PR1", "bad branch", "WrongPath"
    batch.Execute

    AssertTrue stats, "batch.capacityPath.invalid.status", batch.CapacityStatus(1) = "InputErr"
    AssertTrue stats, "batch.capacityPath.invalid.overall", batch.OverallStatus(1) = "InputErr"
    AssertTrue stats, "batch.capacityPath.invalid.noPath", Len(batch.CapacityLoadPathKey(1)) = 0
End Sub

' Проверяет, что физическое плато нормативной диаграммы не считается
' numerical extension. Extension начинается только после eps_ult; иначе
' нормальное состояние с запасом по capacity ошибочно получит FAIL.
Private Sub TestPR2PhysicalStateRunsCrackWithExtensionEnabled(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldExtension As String
    oldExtension = GetSystemSetting("Solver.DirectState.DiagramExtension")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Solver.DirectState.DiagramExtension", "Yes"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "G2_PHYS", -220000#, -7000000#, -5000000#, "PR2", "physical state"
    batch.Execute

    AssertTrue stats, "batch.group2.physical.noExtension", Not batch.ExtensionUsed(1)
    AssertTrue stats, "batch.group2.physical.directOK", batch.DirectStateStatus(1) = "OK"
    AssertTrue stats, "batch.group2.physical.crackRuns", batch.CrackStatus(1) <> "N/A"

Restore:
    SetSystemSetting "Solver.DirectState.DiagramExtension", oldExtension
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.group2.physical; " & Err.Description
    Resume Restore
End Sub

' Проверяет, что настройка SLS.Crack.CoverDistanceMode проходит через полный
' batch-конвейер до CCrackWidthCalculator. Ступенчатый бетонный контур выбран
' специально: расстояние до ближайшего локального выхода и до дальней общей
' растянутой линии у него заметно отличаются.
Private Sub TestBatchCrackCoverDistanceModeChangesAs(ByRef stats As TBatchTestStats)
    Dim oldCoverMode As String
    Dim oldPsiMode As String
    Dim oldZoneMode As String
    oldCoverMode = GetSystemSetting("SLS.Crack.CoverDistanceMode")
    oldPsiMode = GetSystemSetting("SLS.Crack.PsiMode")
    oldZoneMode = GetSystemSetting("SLS.Crack.TensionZoneMode")

    On Error GoTo RestoreAndFail
    SetSystemSetting "SLS.Crack.PsiMode", "User"
    SetSystemSetting "SLS.Crack.TensionZoneMode", "Effective"

    Dim localCover As Double
    localCover = BatchSteppedCrackCoverA("NearestContour")

    Dim globalCover As Double
    globalCover = BatchSteppedCrackCoverA("GlobalExtreme")

    AppendLine stats, "INFO: batch.crack.coverDistanceMode local=" & FormatNumberInvariant(localCover) & _
        "; global=" & FormatNumberInvariant(globalCover)
    AssertTrue stats, "batch.crack.coverMode.localPositive", localCover > 0#
    AssertTrue stats, "batch.crack.coverMode.globalLarger", globalCover > localCover * 2#

Restore:
    SetSystemSetting "SLS.Crack.CoverDistanceMode", oldCoverMode
    SetSystemSetting "SLS.Crack.PsiMode", oldPsiMode
    SetSystemSetting "SLS.Crack.TensionZoneMode", oldZoneMode
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.crack.coverDistanceMode; " & Err.Description
    Resume Restore
End Sub

' Проверяет, что Auto-ветка трещин сохраняет оба состояния около Mcrc.
' BeforeMcrcState нужен для контроля состояния с работающим растянутым бетоном,
' AfterMcrcState - для ручной проверки sigma_s,crc после раскрытия трещины.
Private Sub TestPR2AutoCrackStoresBeforeAndAfterMcrcStates(ByRef stats As TBatchTestStats)
    Dim oldPsiMode As String
    Dim oldAllowable As String
    oldPsiMode = GetSystemSetting("SLS.Crack.PsiMode")
    oldAllowable = GetSystemSetting("SLS.Crack.Allowable")

    On Error GoTo RestoreAndFail
    SetSystemSetting "SLS.Crack.PsiMode", "Auto"
    SetSystemSetting "SLS.Crack.Allowable", "0.0001"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "G2_AUTO_MCRC", -20000#, -15000000#, 0#, "PR2", "auto mcrc states"
    batch.Execute

    AssertTrue stats, "batch.group2.autoMcrc.crackCalculated", _
        batch.CrackStatus(1) = "OK" Or batch.CrackStatus(1) = "FAIL"
    AssertTrue stats, "batch.group2.autoMcrc.beforeState", _
        Not batch.FindNamedState(1, sstBeforeMcrcState) Is Nothing
    AssertTrue stats, "batch.group2.autoMcrc.afterState", _
        Not batch.FindNamedState(1, sstAfterMcrcState) Is Nothing
    AssertTrue stats, "batch.group2.autoMcrc.afterRole", _
        batch.FindNamedState(1, sstAfterMcrcState).MaterialModelRoleText = "CrackedState"

Restore:
    SetSystemSetting "SLS.Crack.PsiMode", oldPsiMode
    SetSystemSetting "SLS.Crack.Allowable", oldAllowable
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.group2.autoMcrc.states; " & Err.Description
    Resume Restore
End Sub

' Проверяет чистый изгиб N=0 в полном batch-конвейере: исходное CrackedState
' находит CStateSolutionRunner, а Auto-ветка трещин затем сохраняет состояния
' до и после Mcrc без отдельной подстановки продольной силы.
Private Sub TestPR2AutoCrackPureBendingStoresMcrcStates(ByRef stats As TBatchTestStats)
    Dim oldPsiMode As String
    Dim oldAllowable As String
    oldPsiMode = GetSystemSetting("SLS.Crack.PsiMode")
    oldAllowable = GetSystemSetting("SLS.Crack.Allowable")

    On Error GoTo RestoreAndFail
    SetSystemSetting "SLS.Crack.PsiMode", "Auto"
    SetSystemSetting "SLS.Crack.Allowable", "0.0001"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "G2_AUTO_PURE_MX", 0#, -15000000#, 0#, "PR2", "auto pure bending"
    batch.Execute

    AssertTrue stats, "batch.group2.autoMcrcPure.directOK", batch.DirectStateStatus(1) = "OK"
    AssertTrue stats, "batch.group2.autoMcrcPure.crackCalculated", _
        batch.CrackStatus(1) = "OK" Or batch.CrackStatus(1) = "FAIL"
    AssertTrue stats, "batch.group2.autoMcrcPure.beforeState", _
        Not batch.FindNamedState(1, sstBeforeMcrcState) Is Nothing
    AssertTrue stats, "batch.group2.autoMcrcPure.afterState", _
        Not batch.FindNamedState(1, sstAfterMcrcState) Is Nothing

Restore:
    SetSystemSetting "SLS.Crack.PsiMode", oldPsiMode
    SetSystemSetting "SLS.Crack.Allowable", oldAllowable
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.group2.autoMcrcPure.states; " & Err.Description
    Resume Restore
End Sub

' Проверяет пользовательский случай Г-сечения: при сжатии 100 тс и сравнительно
' небольшом моменте Mx поиск стадии образования трещины должен завершаться
' расчетным статусом, а не падать в NumFail из-за неудачного стартового НДС.
Private Sub TestPR2RectSetCompressionSmallMomentCrackDoesNotNumFail(ByRef stats As TBatchTestStats)
    Dim oldPsiMode As String
    oldPsiMode = GetSystemSetting("SLS.Crack.PsiMode")

    On Error GoTo RestoreAndFail
    SetSystemSetting "SLS.Crack.PsiMode", "Auto"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim referenceX As Double
    Dim referenceY As Double
    Dim batch As CBatchSectionCalculator
    Set batch = BuildUserRectSetTensionBatch(referenceX, referenceY)
    batch.ApplySettings settings
    batch.AddCombination "G2_L_SMALL_M", -100# * 9806.65, 30# * TEST_TF_M_IN_NMM, 0#, _
        "PR2", "rectset compression small moment crack"
    batch.Execute

    AppendLine stats, "INFO: batch.group2.rectsetSmallMoment crack=" & batch.CrackStatus(1) & _
        "; formed=" & CStr(batch.CrackFormed(1)) & "; lambda=" & FormatNumberInvariant(batch.CrackLambdaCrc(1))
    AssertTrue stats, "batch.group2.rectsetSmallMoment.notNumFail", batch.CrackStatus(1) <> "NumFail"
    AssertTrue stats, "batch.group2.rectsetSmallMoment.finished", _
        batch.CrackStatus(1) = "OK" Or batch.CrackStatus(1) = "FAIL"
    AssertTrue stats, "batch.group2.rectsetSmallMoment.solverCallsIncludeCrack", batch.SolverCallCount > 1

Restore:
    SetSystemSetting "SLS.Crack.PsiMode", oldPsiMode
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.group2.rectsetSmallMoment; " & Err.Description
    Resume Restore
End Sub

' Проверяет полный путь batch + Results writer для новых траекторий
' трещинообразования lambda*N и lambda*NMxy. Раньше найденные Ncrc/Mxy,crc
' могли сохраниться в результате LC, но не попасть в таблицу трещин из-за
' позднего N/A после проверки CrackedState.
Private Sub TestCrackInitiationLoadPathsWriteFormationSummary(ByRef stats As TBatchTestStats)
    Dim oldPath As String
    Dim oldStrategy As String
    Dim oldPsiMode As String
    Dim oldAllowable As String
    oldPath = GetSystemSetting("SLS.Crack.InitiationLoadPath")
    oldStrategy = GetSystemSetting("SLS.Crack.InitiationSolutionStrategy")
    oldPsiMode = GetSystemSetting("SLS.Crack.PsiMode")
    oldAllowable = GetSystemSetting("SLS.Crack.Allowable")

    On Error GoTo RestoreAndFail
    SetSystemSetting "SLS.Crack.InitiationSolutionStrategy", "Auto"
    SetSystemSetting "SLS.Crack.PsiMode", "Auto"
    SetSystemSetting "SLS.Crack.Allowable", "0.0001"

    CheckCrackFormationSummaryForPath stats, "Auto", _
        "batch.crack.pathAuto.summary", 20# * 9806.65, _
        0#, 0#, True, False
    CheckCrackFormationSummaryForPath stats, "lambda*N", _
        "batch.crack.pathN.summary", 20# * 9806.65, _
        0#, 0#, True, False
    CheckCrackFormationSummaryForPath stats, "lambda*NMxy", _
        "batch.crack.pathNMxy.summary", 100# * 9806.65, _
        50# * TEST_TF_M_IN_NMM, 0#, True, True

Restore:
    SetSystemSetting "SLS.Crack.InitiationLoadPath", oldPath
    SetSystemSetting "SLS.Crack.InitiationSolutionStrategy", oldStrategy
    SetSystemSetting "SLS.Crack.PsiMode", oldPsiMode
    SetSystemSetting "SLS.Crack.Allowable", oldAllowable
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.crack.path.summary; " & Err.Description
    Resume Restore
End Sub

' Жестко проверяет Auto-переключение пути образования трещины на пользовательском
' Г-сечении: чистая N должна выбрать lambda*N, чистый момент - lambda*Mxy,
' а сочетание, где N и M по отдельности уже дают трещину, должно дойти до
' пропорционального пути lambda*NMxy вместо fallback psi_s=1.
Private Sub TestCrackAutoFormationPathSwitchesForRectSet(ByRef stats As TBatchTestStats)
    Dim oldPath As String
    Dim oldStrategy As String
    Dim oldPsiMode As String
    Dim oldAllowable As String
    oldPath = GetSystemSetting("SLS.Crack.InitiationLoadPath")
    oldStrategy = GetSystemSetting("SLS.Crack.InitiationSolutionStrategy")
    oldPsiMode = GetSystemSetting("SLS.Crack.PsiMode")
    oldAllowable = GetSystemSetting("SLS.Crack.Allowable")

    On Error GoTo RestoreAndFail
    SetSystemSetting "SLS.Crack.InitiationLoadPath", "Auto"
    SetSystemSetting "SLS.Crack.InitiationSolutionStrategy", "Auto"
    SetSystemSetting "SLS.Crack.PsiMode", "Auto"
    SetSystemSetting "SLS.Crack.Allowable", "0.0001"

    CheckCrackAutoFormationCase stats, "batch.crack.auto.rectset.nOnly", _
        100# * 9806.65, 0#, 0#, ChrW$(&H3BB) & "*N", True, False
    CheckCrackAutoFormationCase stats, "batch.crack.auto.rectset.mOnly", _
        0#, 50# * TEST_TF_M_IN_NMM, 0#, ChrW$(&H3BB) & "*Mxy", False, True
    CheckCrackAutoFormationCase stats, "batch.crack.auto.rectset.nAndM", _
        100# * 9806.65, 50# * TEST_TF_M_IN_NMM, 0#, ChrW$(&H3BB) & "*NMxy", True, True

Restore:
    SetSystemSetting "SLS.Crack.InitiationLoadPath", oldPath
    SetSystemSetting "SLS.Crack.InitiationSolutionStrategy", oldStrategy
    SetSystemSetting "SLS.Crack.PsiMode", oldPsiMode
    SetSystemSetting "SLS.Crack.Allowable", oldAllowable
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.crack.auto.rectset; " & Err.Description
    Resume Restore
End Sub

' Выполняет один Auto-сценарий Mcrc/Ncrc на Г-сечении и проверяет, что
' фактически принятый путь совпал с ожидаемым физическим случаем.
Private Sub CheckCrackAutoFormationCase(ByRef stats As TBatchTestStats, _
        ByVal prefix As String, ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double, ByVal expectedMethod As String, _
        ByVal expectNcrc As Boolean, ByVal expectMcrc As Boolean)
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim referenceX As Double
    Dim referenceY As Double
    Dim batch As CBatchSectionCalculator
    Set batch = BuildUserRectSetTensionBatch(referenceX, referenceY)
    batch.ApplySettings settings
    batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
    batch.AddCombination UCase$(Replace$(Replace$(expectedMethod, ChrW$(&H3BB) & "*", vbNullString), "*", vbNullString)), _
        nValue, mxValue, myValue, "PR2", "auto crack formation"
    batch.Execute

    AppendLine stats, "INFO: " & prefix & "; status=" & batch.CrackStatus(1) & _
        "; method=" & batch.CrackFormationMethod(1) & _
        "; lambda=" & FormatNumberInvariant(batch.CrackLambdaCrc(1)) & _
        "; Ncrc=" & FormatNumberInvariant(batch.CrackFormationNcrc(1)) & _
        "; MxyCrc=" & FormatNumberInvariant(batch.CrackMcrc(1))
    If batch.CrackFormationMethod(1) <> expectedMethod Or batch.CrackStatus(1) = "NumFail" Then _
        AppendLine stats, "DIAG: " & prefix & vbCrLf & batch.DiagnosticLog

    AssertTrue stats, prefix & ".notNumFail", batch.CrackStatus(1) <> "NumFail"
    AssertTrue stats, prefix & ".method", batch.CrackFormationMethod(1) = expectedMethod
    AssertTrue stats, prefix & ".lambda", batch.CrackLambdaCrc(1) > 0# And batch.CrackLambdaCrc(1) <= 1#
    If expectNcrc Then _
        AssertTrue stats, prefix & ".ncrc", Abs(batch.CrackFormationNcrc(1)) > 0#
    If expectMcrc Then _
        AssertTrue stats, prefix & ".mcrc", batch.CrackMcrc(1) > 0#
End Sub

' Выполняет один сценарий трещинообразования и проверяет, что найденная точка
' записана как в объект batch, так и в отдельный блок rngCrackSummaryAnchor.
Private Sub CheckCrackFormationSummaryForPath(ByRef stats As TBatchTestStats, _
        ByVal pathText As String, ByVal prefix As String, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, _
        ByVal expectNcrc As Boolean, ByVal expectMcrc As Boolean)
    SetSystemSetting "SLS.Crack.InitiationLoadPath", pathText

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    AssertTrue stats, prefix & ".pathSetting", _
        StrComp(settings.GetString("SLS.Crack.InitiationLoadPath", vbNullString), pathText, vbTextCompare) = 0

    Dim referenceX As Double
    Dim referenceY As Double
    Dim batch As CBatchSectionCalculator
    Set batch = BuildUserRectSetTensionBatch(referenceX, referenceY)
    batch.ApplySettings settings
    batch.AddCombination UCase$(Replace$(Replace$(pathText, "lambda*", vbNullString), "*", vbNullString)), _
        nValue, mxValue, myValue, "PR2", "crack formation path"
    batch.Execute

    AppendLine stats, "INFO: " & prefix & "; status=" & batch.CrackStatus(1) & _
        "; lambda=" & FormatNumberInvariant(batch.CrackLambdaCrc(1)) & _
        "; Ncrc=" & FormatNumberInvariant(batch.CrackFormationNcrc(1)) & _
        "; MxyCrc=" & FormatNumberInvariant(batch.CrackMcrc(1))
    If batch.CrackStatus(1) = "NumFail" Or _
            (expectNcrc And batch.CrackFormationNcrc(1) = 0#) Or _
            (expectMcrc And batch.CrackMcrc(1) = 0#) Then
        AppendLine stats, "DIAG: " & prefix & vbCrLf & batch.DiagnosticLog
    End If

    Dim writer As CCrackSummaryWriter
    Set writer = New CCrackSummaryWriter
    writer.WriteSummary ThisWorkbook, batch

    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim anchor As Object
    Set anchor = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange
    AssertTrue stats, prefix & ".sheetMethod", _
        CellHasDisplayedResult(resultsSheet.Cells.Item(anchor.Row, 14).Value2)
    If expectNcrc Then
        AssertTrue stats, prefix & ".sheetNcrc", _
            CellHasDisplayedResult(resultsSheet.Cells.Item(anchor.Row, 15).Value2)
    End If
    If expectMcrc Then
        AssertTrue stats, prefix & ".sheetMcrc", _
            CellHasDisplayedResult(resultsSheet.Cells.Item(anchor.Row, 16).Value2)
    End If
End Sub

' Отличает реально выведенную величину от пустой ячейки и статуса N/A.
Private Function CellHasDisplayedResult(ByVal value As Variant) As Boolean
    Dim text As String
    text = Trim$(CStr(value))
    CellHasDisplayedResult = Len(text) > 0 And StrComp(text, "N/A", vbTextCompare) <> 0
End Function

' Проверяет сценарий из пользовательского расчета: большое осевое растяжение
' второй группы должно доходить до технического продолжения диаграммы и давать
' инженерный FAIL, а не теряться как численная несходимость NumFail.
Private Sub TestPR2AxialTensionBeyondPhysicalLimitUsesExtension(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldMaxIterations As String
    Dim oldLoadSteps As String
    Dim oldExtension As String
    oldMaxIterations = GetSystemSetting("Solver.MaxIterations")
    oldLoadSteps = GetSystemSetting("Solver.LoadSteps")
    oldExtension = GetSystemSetting("Solver.DirectState.DiagramExtension")

    On Error GoTo RestoreAndFail
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
    Set batch = BuildUserRectSetTensionBatch(referenceX, referenceY, provider)
    batch.ApplySettings settings, units
    batch.AddCombination "G2_TENSION_EXT", 900# * 9806.65, 0#, 0#, "PR2", "tension over SLS yield"
    batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
    batch.Execute

    AssertTrue stats, "batch.group2.extension.directFail", batch.DirectStateStatus(1) = "FAIL"
    AssertTrue stats, "batch.group2.extension.used", batch.ExtensionUsed(1)
    AssertTrue stats, "batch.group2.extension.noCrack", batch.CrackStatus(1) = "N/A"
    AssertTrue stats, "batch.group2.extension.overall", batch.OverallStatus(1) = "FAIL"
    AssertTrue stats, "batch.group2.extension.strain", batch.MaxSteelStrain(1) > 0.025

Restore:
    SetSystemSetting "Solver.MaxIterations", oldMaxIterations
    SetSystemSetting "Solver.LoadSteps", oldLoadSteps
    SetSystemSetting "Solver.DirectState.DiagramExtension", oldExtension
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.group2.extension.tension; " & Err.Description
    Resume Restore
End Sub

' Проверяет пользовательский сценарий PR1: растянутый бетон в прочности не
' работает, перегрузка должна распознаваться через extension арматуры, а не
' превращаться в рассинхрон DirectStateStatus=FAIL при ExtensionUsed=False.
Private Sub TestPR1AxialTensionBeyondPhysicalLimitUsesExtension(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldMaxIterations As String
    Dim oldLoadSteps As String
    Dim oldExtension As String
    oldMaxIterations = GetSystemSetting("Solver.MaxIterations")
    oldLoadSteps = GetSystemSetting("Solver.LoadSteps")
    oldExtension = GetSystemSetting("Solver.DirectState.DiagramExtension")

    On Error GoTo RestoreAndFail
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
    Set batch = BuildUserRectSetTensionBatch(referenceX, referenceY, provider)
    batch.ApplySettings settings, units
    batch.AddCombination "G1_TENSION_EXT", 900# * 9806.65, 0#, 0#, "PR1", "tension over ULS diagram"
    batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
    batch.Execute

    AssertTrue stats, "batch.group1.extension.directFail", batch.DirectStateStatus(1) = "FAIL"
    AssertTrue stats, "batch.group1.extension.used", batch.ExtensionUsed(1)
    AssertTrue stats, "batch.group1.extension.noCrack", batch.CrackStatus(1) = "N/A"
    AssertTrue stats, "batch.group1.extension.overall", batch.OverallStatus(1) = "FAIL"

Restore:
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
Private Sub TestPR1AxialTensionNearLimitDoesNotJumpToNumFail(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldMaxIterations As String
    Dim oldLoadSteps As String
    Dim oldExtension As String
    oldMaxIterations = GetSystemSetting("Solver.MaxIterations")
    oldLoadSteps = GetSystemSetting("Solver.LoadSteps")
    oldExtension = GetSystemSetting("Solver.DirectState.DiagramExtension")

    On Error GoTo RestoreAndFail
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
    Set batch = BuildUserRectSetTensionBatch(referenceX, referenceY, provider)
    batch.ApplySettings settings, units

    Dim loads As Variant
    loads = Array(796#, 797#, 800#, 801#, 809#, 810#, 811#, 812#)

    Dim i As Long
    For i = LBound(loads) To UBound(loads)
        batch.AddCombination "G1_T" & CStr(CLng(loads(i))), CDbl(loads(i)) * 9806.65, 0#, 0#, _
            "PR1", "near physical tension limit"
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
Private Sub TestPR1AxialCompressionNearLimitDoesNotJumpToNumFail(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldMaxIterations As String
    Dim oldLoadSteps As String
    Dim oldExtension As String
    oldMaxIterations = GetSystemSetting("Solver.MaxIterations")
    oldLoadSteps = GetSystemSetting("Solver.LoadSteps")
    oldExtension = GetSystemSetting("Solver.DirectState.DiagramExtension")

    On Error GoTo RestoreAndFail
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
    Set batch = BuildUserRectSetTensionBatch(referenceX, referenceY, provider)
    batch.ApplySettings settings, units

    Dim loads As Variant
    loads = Array(1222#, 1223#, 1224#, 1225#)

    Dim i As Long
    For i = LBound(loads) To UBound(loads)
        batch.AddCombination "G1_C" & CStr(CLng(loads(i))), -CDbl(loads(i)) * 9806.65, 0#, 0#, _
            "PR1", "near physical compression limit"
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
Private Sub TestPR1AxialTensionProgressionAfterLimitIsStableFail(ByRef stats As TBatchTestStats)
    Dim loads As Variant
    loads = Array(750#, 796#, 797#, 798#, 799#, 800#, 801#, 805#, 810#, 811#, 812#, _
        820#, 830#, 850#, 900#, 925#, 930#, 950#, 1000#, 1200#)
    RunAxialProgressionAfterLimit stats, "batch.group1.tensionProgression", loads, True, 1
End Sub

' Проверяет аналогичный ряд для сжатия. Это защищает от ситуации, когда при
' росте N одно сочетание внутри перегруженной области внезапно получает NumFail,
' хотя соседние нагрузки уже корректно сходятся в extension и дают FAIL.
Private Sub TestPR1AxialCompressionProgressionAfterLimitIsStableFail(ByRef stats As TBatchTestStats)
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
    oldMaxIterations = GetSystemSetting("Solver.MaxIterations")
    oldLoadSteps = GetSystemSetting("Solver.LoadSteps")
    oldExtension = GetSystemSetting("Solver.DirectState.DiagramExtension")

    On Error GoTo RestoreAndFail
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
    Set batch = BuildUserRectSetTensionBatch(referenceX, referenceY, provider)
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
            "PR1", "axial progression after physical limit"
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
Private Sub TestPR2AxialCompressionBeyondPhysicalLimitUsesExtension(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldMaxIterations As String
    Dim oldLoadSteps As String
    Dim oldExtension As String
    oldMaxIterations = GetSystemSetting("Solver.MaxIterations")
    oldLoadSteps = GetSystemSetting("Solver.LoadSteps")
    oldExtension = GetSystemSetting("Solver.DirectState.DiagramExtension")

    On Error GoTo RestoreAndFail
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
    Set batch = BuildUserRectSetTensionBatch(referenceX, referenceY, provider)
    batch.ApplySettings settings, units
    batch.AddCombination "G2_COMPRESSION_EXT", -2000# * 9806.65, 0#, 0#, "PR2", "compression over SLS diagram"
    batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
    batch.Execute

    AssertTrue stats, "batch.group2.extension.compression.directFail", batch.DirectStateStatus(1) = "FAIL"
    AssertTrue stats, "batch.group2.extension.compression.used", batch.ExtensionUsed(1)
    AssertTrue stats, "batch.group2.extension.compression.noCrack", batch.CrackStatus(1) = "N/A"
    AssertTrue stats, "batch.group2.extension.compression.strain", _
        batch.MinConcreteStrain(1) < provider.ConcreteCompressionLimit(cpCrackedNDS)

Restore:
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
Private Sub TestPR2BendingBeyondPhysicalLimitUsesExtension(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldMaxIterations As String
    Dim oldLoadSteps As String
    Dim oldExtension As String
    oldMaxIterations = GetSystemSetting("Solver.MaxIterations")
    oldLoadSteps = GetSystemSetting("Solver.LoadSteps")
    oldExtension = GetSystemSetting("Solver.DirectState.DiagramExtension")

    On Error GoTo RestoreAndFail
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
    Set batch = BuildUserRectSetTensionBatch(referenceX, referenceY, provider)
    batch.ApplySettings settings, units
    batch.AddCombination "G2_BENDING_EXT", -120# * 9806.65, 420# * 9806.65 * 1000#, _
        -180# * 9806.65 * 1000#, "PR2", "bending over SLS diagram"
    batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
    batch.Execute

    AssertTrue stats, "batch.group2.extension.bending.directFail", batch.DirectStateStatus(1) = "FAIL"
    AssertTrue stats, "batch.group2.extension.bending.used", batch.ExtensionUsed(1)
    AssertTrue stats, "batch.group2.extension.bending.noCrack", batch.CrackStatus(1) = "N/A"
    AssertTrue stats, "batch.group2.extension.bending.curvature", _
        Abs(batch.KappaX(1)) > 0.000000001 Or Abs(batch.KappaY(1)) > 0.000000001

Restore:
    SetSystemSetting "Solver.MaxIterations", oldMaxIterations
    SetSystemSetting "Solver.LoadSteps", oldLoadSteps
    SetSystemSetting "Solver.DirectState.DiagramExtension", oldExtension
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.group2.extension.bending; " & Err.Description
    Resume Restore
End Sub

' Проверяет, что governing по прочности выбирается среди профилей, где capacity
' действительно запрошен. PR2 в базовом шаблоне трещин не участвует в этой
' выборке, потому что его CapacityStatus получает N/A.
Private Sub TestBatchGoverningUsesStrengthProfilesOnly(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "SAFE", -120000#, -1200000#, -600000#, "PR1", "larger safety"
    batch.AddCombination "GOV", -120000#, -5200000#, -2600000#, "PR1", "smaller safety"
    batch.AddCombination "CRACK", -120000#, -9000000#, -4500000#, "PR2", "crack-only profile"
    batch.Execute

    AssertTrue stats, "batch.governing.profiles.capacity.order", _
        batch.LambdaCapacity(2) > 0# And batch.LambdaCapacity(2) < batch.LambdaCapacity(1)
    AssertTrue stats, "batch.governing.profiles.pr2Skipped", batch.CapacityStatus(3) = "N/A"
    AssertTrue stats, "batch.governing.profiles.id", batch.GoverningCombinationID = "GOV"
End Sub

' Проверяет, что отключение capacity задается самим профилем, а не старой
' глобальной настройкой области расчета.
Private Sub TestPR2SkipsCapacityByProfile(ByRef stats As TBatchTestStats)
    On Error GoTo RestoreAndFail

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "STRENGTH", -120000#, -1200000#, -600000#, "PR1", "strength profile"
    batch.AddCombination "CRACK", -120000#, -5200000#, -2600000#, "PR2", "crack profile"
    batch.Execute

    AssertTrue stats, "batch.profile.capacity.pr1.runs", batch.CapacityStatus(1) <> "N/A"
    AssertTrue stats, "batch.profile.capacity.pr2.skipped", batch.CapacityStatus(2) = "N/A"

Restore:
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.profile.capacity; " & Err.Description
    Resume Restore
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestLoadReferenceTransformsUserMoments(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "REF", -1000#, 3000000#, -3000000#, "PR1", "reference"
    batch.ApplyLoadReference 40#, -25#

    AssertClose stats, "batch.reference.userMx", batch.UserMx(1), 3000000#, 0.000001
    AssertClose stats, "batch.reference.userMy", batch.UserMy(1), -3000000#, 0.000001
    AssertClose stats, "batch.reference.internalMx", batch.Mx(1), 3025000#, 0.000001
    AssertClose stats, "batch.reference.internalMy", batch.My(1), -3040000#, 0.000001
    AssertClose stats, "batch.reference.x", batch.LoadReferenceX, 40#, 0.000001
    AssertClose stats, "batch.reference.y", batch.LoadReferenceY, -25#, 0.000001
    AssertClose stats, "batch.reference.offsetX.defaultBase", batch.LoadReferenceOffsetX, 40#, 0.000001
    AssertClose stats, "batch.reference.offsetY.defaultBase", batch.LoadReferenceOffsetY, -25#, 0.000001

    Dim shiftedBatch As CBatchSectionCalculator
    Set shiftedBatch = BuildBatchCalculator()
    shiftedBatch.AddCombination "REF2", -1000#, 20000#, -30000#, "PR1", "reference shifted"
    shiftedBatch.ApplyLoadReference 140#, 75#, 100#, 100#
    AssertClose stats, "batch.reference.offsetX.centroidBase", shiftedBatch.LoadReferenceOffsetX, 40#, 0.000001
    AssertClose stats, "batch.reference.offsetY.centroidBase", shiftedBatch.LoadReferenceOffsetY, -25#, 0.000001
End Sub

' Проверяет единый объект нагрузки: он хранит Mx/My в пользовательской точке,
' умеет переносить их к произвольному центру и обратно восстанавливает
' пользовательские моменты из усилий, возвращенных stability-фильтром.
Private Sub TestSectionLoadStateTransfersMoments(ByRef stats As TBatchTestStats)
    Dim loadState As CSectionLoadState
    Set loadState = New CSectionLoadState
    loadState.Initialize -1000#, 3000000#, -3000000#, 40#, -25#

    AssertClose stats, "batch.loadState.loadMx", loadState.LoadPointMx, 3000000#, 0.000001
    AssertClose stats, "batch.loadState.loadMy", loadState.LoadPointMy, -3000000#, 0.000001
    AssertClose stats, "batch.loadState.internalMx", loadState.InternalMx, 3025000#, 0.000001
    AssertClose stats, "batch.loadState.internalMy", loadState.InternalMy, -3040000#, 0.000001
    AssertClose stats, "batch.loadState.mxAboutShifted", loadState.MxAboutPoint(75#), 3100000#, 0.000001
    AssertClose stats, "batch.loadState.myAboutShifted", loadState.MyAboutPoint(10#), -3030000#, 0.000001
    AssertClose stats, "batch.loadState.axialMx", loadState.AxialMxAboutPoint(0#), 25000#, 0.000001
    AssertClose stats, "batch.loadState.axialMy", loadState.AxialMyAboutPoint(0#), -40000#, 0.000001
    AssertTrue stats, "batch.loadState.compression", loadState.IsCompression(0.000001)
    AssertClose stats, "batch.loadState.compressionMagnitude", loadState.CompressionMagnitude(0.000001), 1000#, 0.000001
    AssertClose stats, "batch.loadState.signedCompression", loadState.SignedCompression, 1000#, 0.000001

    Dim fromPoint As CSectionLoadState
    Set fromPoint = New CSectionLoadState
    fromPoint.InitializeFromPointMoments -1000#, -12000#, 5000#, 10#, 20#, 40#, -25#
    AssertClose stats, "batch.loadState.restoreMx", fromPoint.MxAboutPoint(20#), -12000#, 0.000001
    AssertClose stats, "batch.loadState.restoreMy", fromPoint.MyAboutPoint(10#), 5000#, 0.000001

    Dim central As CSectionLoadState
    Set central = New CSectionLoadState
    central.Initialize 200000#, 0#, 0#, 40#, -25#
    AssertTrue stats, "batch.loadState.tension", central.IsTension(0.000001)
    AssertClose stats, "batch.loadState.tensionSignedCompression", central.SignedCompression, -200000#, 0.000001
    AssertTrue stats, "batch.loadState.centralAtLoadPoint", _
        central.IsCentralTensionAbout(40#, -25#, 0.000001, 1#, 1#)
    AssertTrue stats, "batch.loadState.eccentricAtShiftedPoint", _
        Not central.IsCentralTensionAbout(50#, -25#, 0.000001, 1#, 1#)
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
    CheckRectSetReferenceUsesConcreteCentroid stats
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
    Dim loadState As CSectionLoadState
    Set loadState = New CSectionLoadState
    loadState.Initialize nValue, 0#, 0#, refX, refY
    solver.Solve section, concrete, steel, loadState.N, loadState.InternalMx, loadState.InternalMy

    AssertTrue stats, "batch.reference." & caseName & ".converged", solver.Converged
    AssertClose stats, "batch.reference." & caseName & ".kappaX", solver.KappaX, 0#, tolerance
    AssertClose stats, "batch.reference." & caseName & ".kappaY", solver.KappaY, 0#, tolerance
End Sub

' Проверяет, что для несимметричного Г-сечения новый reference point берется
' от бетонной части. Это важнее, чем требовать нулевую кривизну: при
' несимметричной арматуре бетонный и приведенный центры могут не совпадать.
Private Sub CheckRectSetReferenceUsesConcreteCentroid(ByRef stats As TBatchTestStats)
    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh RectSetGeometry(250#, 550#, 600#, 250#), 50#, 50#, 1, 1

    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, _
        RectSetRebars(250#, 550#, 600#, 250#, 40#, 50, 32#))

    Dim concreteX As Double
    Dim concreteY As Double
    CalculateConcreteSectionCentroid section, concreteX, concreteY

    Dim transformedX As Double
    Dim transformedY As Double
    CalculateTransformedSectionCentroid section, ProvisionalConcrete(), ProvisionalSteel(), transformedX, transformedY

    AssertTrue stats, "batch.reference.rectset.concreteCenter.exists", Abs(concreteX) + Abs(concreteY) > 0.000001
    AssertTrue stats, "batch.reference.rectset.centerDifference", _
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
    Dim axialLoad As CSectionLoadState
    Set axialLoad = New CSectionLoadState
    axialLoad.Initialize nValue, 0#, 0#, refX, refY
    axialSolver.Solve section, concrete, steel, axialLoad.N, axialLoad.InternalMx, axialLoad.InternalMy

    AssertTrue stats, "batch.tension." & caseName & ".central.converged", axialSolver.Converged
    AssertClose stats, "batch.tension." & caseName & ".central.kappaX", axialSolver.KappaX, 0#, tolerance
    AssertClose stats, "batch.tension." & caseName & ".central.kappaY", axialSolver.KappaY, 0#, tolerance

    Dim eccentricSolver As CSectionSolver
    Set eccentricSolver = New CSectionSolver
    eccentricSolver.LoadSteps = 1
    eccentricSolver.MaxIterations = 60
    Dim eccentricLoad As CSectionLoadState
    Set eccentricLoad = New CSectionLoadState
    eccentricLoad.Initialize nValue, 0#, 0#, refX + eccentricOffsetX, refY + eccentricOffsetY
    eccentricSolver.Solve section, concrete, steel, eccentricLoad.N, eccentricLoad.InternalMx, eccentricLoad.InternalMy

    AssertTrue stats, "batch.tension." & caseName & ".eccentric.converged", eccentricSolver.Converged
    AssertTrue stats, "batch.tension." & caseName & ".eccentric.kappa", _
        Abs(eccentricSolver.KappaX) > tolerance Or Abs(eccentricSolver.KappaY) > tolerance
End Sub

' Проверяет профиль PR1 после отказа от пользовательского запаса по деформациям.
' В этом профиле выполняются и прямое НДС по прочности, и поиск capacity.
Private Sub TestDirectStateReportsSectionStatus(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String

    On Error GoTo RestoreAndFail

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "DS_SAFE", -120000#, -1500000#, -800000#, "PR1", "smaller strain"
    batch.AddCombination "DS_GOV", -120000#, -4500000#, -2400000#, "PR1", "larger strain"
    batch.Execute

    AssertTrue stats, "batch.direct.lambda.positive", batch.LambdaCapacity(1) > 0# And batch.LambdaCapacity(2) > 0#
    AssertTrue stats, "batch.direct.capacity.ok", batch.CapacityStatus(1) = "OK" And batch.CapacityStatus(2) = "OK"
    AssertTrue stats, "batch.direct.crack.na", StrComp(batch.CrackStatus(1), "N/A", vbTextCompare) = 0
    AssertTrue stats, "batch.direct.state.status", Len(batch.DirectStateStatus(1)) > 0 And Len(batch.DirectStateStatus(2)) > 0
    AssertTrue stats, "batch.direct.governing.present", Len(batch.GoverningCombinationID) > 0

Restore:
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.direct.sectionStatus; " & Err.Description
    Resume Restore
End Sub

' Проверяет внецентренное растяжение стандартного Г-сечения. Этот случай
' регрессирует численную дырку прямого StrengthState: при N=-9 тс и Mx=50 тс*м
' capacity находил предельное состояние, а прямой solve от заданных усилий мог
' остановиться как NumFail из-за неудачного стартового приближения.
Private Sub TestPR1RectSetSmallTensionMomentDirectStateDoesNotNumFail(ByRef stats As TBatchTestStats)
    Dim oldDirect As String
    Dim oldCapacity As String
    Dim oldCrack As String
    Dim oldStability As String
    oldDirect = GetProfileValue("Calculation.Strength.DirectState", "PR1")
    oldCapacity = GetProfileValue("Calculation.Strength.Capacity", "PR1")
    oldCrack = GetProfileValue("Calculation.Crack.Width", "PR1")
    oldStability = GetProfileValue("Calculation.Stability.Enabled", "PR1")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Strength.DirectState", "PR1", "Yes"
    SetProfileValue "Calculation.Strength.Capacity", "PR1", "Yes"
    SetProfileValue "Calculation.Crack.Width", "PR1", "No"
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "No"

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
    Set batch = BuildUserRectSetTensionBatch(referenceX, referenceY, provider)
    batch.ApplySettings settings, units
    batch.AddCombination "G1_T9_MX50", 9# * 9806.65, 50# * TEST_TF_M_IN_NMM, 0#, _
        "PR1", "user N=-9 tf and Mx=50 tf*m"
    batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
    batch.Execute

    AppendLine stats, "INFO: batch.group1.tensionMoment.direct=" & batch.DirectStateStatus(1) & _
        "; capacity=" & batch.CapacityStatus(1) & _
        "; epsCmin=" & FormatNumberInvariant(batch.MinConcreteStrain(1)) & _
        "; epsSmax=" & FormatNumberInvariant(batch.MaxSteelStrain(1))
    If batch.DirectStateStatus(1) = "NumFail" Then
        AppendLine stats, "DIAG: batch.group1.tensionMoment" & vbCrLf & batch.DiagnosticLog
    End If
    AssertTrue stats, "batch.group1.tensionMoment.directNotNumFail", _
        batch.DirectStateStatus(1) = "OK" Or batch.DirectStateStatus(1) = "FAIL"
    AssertTrue stats, "batch.group1.tensionMoment.capacityNotNumFail", _
        batch.CapacityStatus(1) = "OK" Or batch.CapacityStatus(1) = "FAIL"

Restore:
    SetProfileValue "Calculation.Strength.DirectState", "PR1", oldDirect
    SetProfileValue "Calculation.Strength.Capacity", "PR1", oldCapacity
    SetProfileValue "Calculation.Crack.Width", "PR1", oldCrack
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldStability
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.group1.tensionMoment; " & Err.Description
    Resume Restore
End Sub

' Проверяет базовый профиль PR1: прочностное НДС и capacity выполняются, а
' расчет раскрытия трещин не запускается.
Private Sub TestPR1RunsStrengthWithoutCrackWidth(ByRef stats As TBatchTestStats)
    On Error GoTo RestoreAndFail

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "PR1_STRENGTH", -150000#, -3000000#, 0#, "PR1", "strength profile"
    batch.Execute

    AssertTrue stats, "batch.pr1.capacityRuns", batch.CapacityStatus(1) <> "N/A"
    AssertTrue stats, "batch.pr1.directRuns", batch.DirectStateStatus(1) <> "N/A"
    AssertTrue stats, "batch.pr1.crackNA", batch.CrackStatus(1) = "N/A"
    AssertTrue stats, "batch.pr1.hasState", batch.StateAvailable(1) And batch.StateAvailableCount > 0

Restore:
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.pr1.strengthWithoutCrack; " & Err.Description
    Resume Restore
End Sub

' Проверяет, что численная несходимость прямого НДС превращается в короткий
' пользовательский статус NumFail и не порождает расчет трещин.
Private Sub TestDirectStateReportsNumericalFailure(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldMaxIterations As String
    oldMaxIterations = GetSystemSetting("Solver.MaxIterations")

    On Error GoTo RestoreAndFail
    SetSystemSetting "Solver.MaxIterations", "1"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "DS_FAIL", -120000#, -12000000#, -7000000#, "PR1", "forced non-convergence"
    batch.Execute

    AssertTrue stats, "batch.direct.failure.status", batch.Status(1) = "NumFail" Or batch.Status(1) = "FAIL"
    AssertTrue stats, "batch.direct.failure.directStatus", batch.DirectStateStatus(1) = "NumFail" Or batch.DirectStateStatus(1) = "FAIL"
    AssertTrue stats, "batch.direct.failure.crack", batch.CrackStatus(1) = "N/A"

Restore:
    SetSystemSetting "Solver.MaxIterations", oldMaxIterations
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.direct.failure.status; " & Err.Description
    Resume Restore
End Sub

' Проверяет проверку продольных трещин: она должна брать максимальное
' сжимающее напряжение бетона из уже найденного прямого НДС PR2 и не
' требовать отдельного расчетного purpose или повторного запуска solver-а.
Private Sub TestLongitudinalCrackCheckUsesDirectStateStress(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "LONG", -90000#, 0#, 0#, "PR2", "longitudinal"
    batch.Execute

    AssertTrue stats, "batch.longCrack.status.finished", _
        batch.LongitudinalCrackStatus(1) = "OK" Or batch.LongitudinalCrackStatus(1) = "FAIL"
    AssertTrue stats, "batch.longCrack.sigma", batch.MaxConcreteCompressionStress(1) > 0#
    AssertClose stats, "batch.longCrack.rbMc2", batch.LongitudinalCrackRbMc2(1), 14.6, 0.000000001
    AssertClose stats, "batch.longCrack.util", batch.LongitudinalCrackUtilization(1), _
        batch.MaxConcreteCompressionStress(1) / batch.LongitudinalCrackRbMc2(1), 0.000000001
End Sub

' Проверяет нормативную область применения: продольные трещины являются
' проверкой трещиностойкости по PR2, поэтому для PR1 этот блок остается
' N/A и не влияет на прочностной результат.
Private Sub TestLongitudinalCrackSkippedForPR1(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "LONG_G1", -90000#, 0#, 0#, "PR1", "longitudinal group1"
    batch.Execute

    AssertTrue stats, "batch.longCrack.group1.na", batch.LongitudinalCrackStatus(1) = "N/A"
    AssertClose stats, "batch.longCrack.group1.noRbMc2", batch.LongitudinalCrackRbMc2(1), 0#, 0.000000001
    AssertClose stats, "batch.longCrack.group1.noUtil", batch.LongitudinalCrackUtilization(1), 0#, 0.000000001
End Sub

' Проверяет, что профильный флаг включает расчет устойчивости как отдельный
' фильтр перед НДМ: при сжатии SP63 должен дать применимый статус и расчетные
' величины Ncr/eta, а не оставаться в N/A.
Private Sub TestStabilityProfileEnablesSP63ForCompression(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    Dim oldAccPlanes As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")
    oldAccPlanes = GetSystemSetting("Stability.AccidentalEccentricityPlanes")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP63"
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", "OnlyMomentPlane"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "STAB_SP63", -120000#, 1000000#, 0#, "PR1", "stability sp63"
    batch.Execute

    AssertTrue stats, "batch.stability.sp63.status", batch.StabilityStatus(1) = "OK" Or batch.StabilityStatus(1) = "FAIL"
    AssertTrue stats, "batch.stability.sp63.ncr", batch.StabilityCriticalForce(1) > 0#
    AssertTrue stats, "batch.stability.sp63.eta", batch.StabilityEta1(1) > 0# Or batch.StabilityEta2(1) > 0#

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", oldAccPlanes
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.sp63; " & Err.Description
    Resume Restore
End Sub

' Проверяет область применения фильтра устойчивости: при растяжении сжатой
' продольной силы нет, поэтому проверка становится N/A и не ухудшает OverallStatus.
Private Sub TestStabilitySkippedForTension(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP63"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "STAB_TENSION", 50000#, 0#, 0#, "PR1", "stability tension"
    batch.Execute

    AssertTrue stats, "batch.stability.tension.na", batch.StabilityStatus(1) = "N/A"
    AssertClose stats, "batch.stability.tension.noNcr", batch.StabilityCriticalForce(1), 0#, 0.000000001
    AssertTrue stats, "batch.stability.tension.overall", batch.Status(1) <> "InputErr"

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.tension; " & Err.Description
    Resume Restore
End Sub

' Проверяет две договоренности по СП 35: табличные коэффициенты приходят из
' rngSP35Table721, а материал устойчивости выбирается отдельной строкой профиля.
' При SLS(II) в тестовом наборе Rb/Rsc больше, поэтому расчетная сила должна
' получиться выше, чем при нормативном для устойчивости ULS(I).
Private Sub TestStabilitySP35UsesConfigTableAndProfileValueSet(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetSystemSetting "Stability.Code", "SP35"

    Dim ncrULS As Double
    Dim ncrSLS As Double
    ncrULS = StabilitySP35CriticalForceForValueSet("ULS(I)")
    ncrSLS = StabilitySP35CriticalForceForValueSet("SLS(II)")

    AssertTrue stats, "batch.stability.sp35.uls.ncr", ncrULS > 0#
    AssertTrue stats, "batch.stability.sp35.sls.ncr", ncrSLS > 0#
    AssertTrue stats, "batch.stability.sp35.valueSet.affectsResult", ncrSLS > ncrULS * 1.05

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.sp35; " & Err.Description
    Resume Restore
End Sub

' Проверяет, что табличная ветвь СП 35 хранит предельную силу отдельно от Ncr.
' В этой ветви критическая сила по формуле Ncr не считается, поэтому Ncr остается 0,
' а запас берется по Nult,stab.
Private Sub TestStabilitySP35TableSeparatesNcrAndNult(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    Dim oldLength As String
    Dim oldMu1 As String
    Dim oldMu2 As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")
    oldLength = GetSystemSetting("Stability.ElementLength")
    oldMu1 = GetSystemSetting("Stability.Mu1")
    oldMu2 = GetSystemSetting("Stability.Mu2")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP35"
    SetSystemSetting "Stability.ElementLength", "1000"
    SetSystemSetting "Stability.Mu1", "1"
    SetSystemSetting "Stability.Mu2", "1"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.SetSP35Table721 ThisWorkbook.Names.Item("rngSP35Table721").RefersToRange.Value2
    batch.AddCombination "STAB_SP35_TABLE", -120000#, 0#, 0#, "PR1", "sp35 table"
    batch.Execute

    AssertTrue stats, "batch.stability.sp35.table.status", _
        batch.StabilityStatus(1) = "OK" Or batch.StabilityStatus(1) = "FAIL"
    AssertClose stats, "batch.stability.sp35.table.ncr1.na", batch.StabilityNcr1(1), 0#, 0.000000001
    AssertClose stats, "batch.stability.sp35.table.ncr2.na", batch.StabilityNcr2(1), 0#, 0.000000001
    AssertTrue stats, "batch.stability.sp35.table.nult1", batch.StabilityNultimate1(1) > 0#
    AssertTrue stats, "batch.stability.sp35.table.nult2", batch.StabilityNultimate2(1) > 0#

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    SetSystemSetting "Stability.ElementLength", oldLength
    SetSystemSetting "Stability.Mu1", oldMu1
    SetSystemSetting "Stability.Mu2", oldMu2
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.sp35.tableNult; " & Err.Description
    Resume Restore
End Sub

' Проверяет СП 35: в табличной ветви ec/r <= 1 коэффициент phi_l всегда
' берется из таблицы 7.21. Режим PhiL2 не должен подменять это табличное
' значение на 2, потому что он относится к ветвям с расчетом длительности.
Private Sub TestStabilitySP35TableIgnoresPhiL2(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    Dim oldLength As String
    Dim oldMu1 As String
    Dim oldMu2 As String
    Dim oldPhiMode As String
    Dim oldAccMode As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")
    oldLength = GetSystemSetting("Stability.ElementLength")
    oldMu1 = GetSystemSetting("Stability.Mu1")
    oldMu2 = GetSystemSetting("Stability.Mu2")
    oldPhiMode = GetSystemSetting("Stability.PhiLMode")
    oldAccMode = GetSystemSetting("Stability.AccidentalEccentricityMode")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP35"
    SetSystemSetting "Stability.ElementLength", "1000"
    SetSystemSetting "Stability.Mu1", "1"
    SetSystemSetting "Stability.Mu2", "1"
    SetSystemSetting "Stability.PhiLMode", "PhiL2"
    SetSystemSetting "Stability.AccidentalEccentricityMode", "AutoWithL"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.SetSP35Table721 ThisWorkbook.Names.Item("rngSP35Table721").RefersToRange.Value2
    batch.AddCombination "SP35_TABLE_PHIL", -120000#, 0#, 0#, "PR1", "sp35 table phi_l"
    batch.Execute

    AssertTrue stats, "batch.stability.sp35.tablePhi.branch", _
        batch.StabilityPlaneBranch1(1) = "SP35-table" Or batch.StabilityPlaneBranch2(1) = "SP35-table"
    AssertClose stats, "batch.stability.sp35.tablePhi.phiL1", _
        batch.StabilityPhiL1(1), batch.StabilityPhiLTable1(1), 0.000000001
    AssertTrue stats, "batch.stability.sp35.tablePhi.notTwo", batch.StabilityPhiL1(1) < 1.99

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    SetSystemSetting "Stability.ElementLength", oldLength
    SetSystemSetting "Stability.Mu1", oldMu1
    SetSystemSetting "Stability.Mu2", oldMu2
    SetSystemSetting "Stability.PhiLMode", oldPhiMode
    SetSystemSetting "Stability.AccidentalEccentricityMode", oldAccMode
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.sp35.tablePhiL; " & Err.Description
    Resume Restore
End Sub

' Проверяет табличную ветвь СП 35: если As/Ab больше 3%, бетонный вклад в
' Nult уменьшается на As. Это не заменяет приведенное сечение (n - 1)As,
' которое отвечает только за центр тяжести, оси и ядровое расстояние.
Private Sub TestStabilitySP35TableRebarAreaCorrection(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldDirect As String
    Dim oldCapacity As String
    Dim oldValueSet As String
    Dim oldCode As String
    Dim oldLength As String
    Dim oldMu1 As String
    Dim oldMu2 As String
    Dim oldAccMode As String
    Dim oldAccPlanes As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldDirect = GetProfileValue("Calculation.Strength.DirectState", "PR1")
    oldCapacity = GetProfileValue("Calculation.Strength.Capacity", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")
    oldLength = GetSystemSetting("Stability.ElementLength")
    oldMu1 = GetSystemSetting("Stability.Mu1")
    oldMu2 = GetSystemSetting("Stability.Mu2")
    oldAccMode = GetSystemSetting("Stability.AccidentalEccentricityMode")
    oldAccPlanes = GetSystemSetting("Stability.AccidentalEccentricityPlanes")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "Calculation.Strength.DirectState", "PR1", "No"
    SetProfileValue "Calculation.Strength.Capacity", "PR1", "No"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP35"
    SetSystemSetting "Stability.ElementLength", "1000"
    SetSystemSetting "Stability.Mu1", "1"
    SetSystemSetting "Stability.Mu2", "1"
    SetSystemSetting "Stability.AccidentalEccentricityMode", "AutoWithL"
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", "BothPlanes"

    Dim concreteArea As Double
    Dim steelArea As Double
    Dim phiValue As Double
    Dim nUltimate As Double

    nUltimate = SP35TableNultForCircle(500#, 4, 12#, concreteArea, steelArea, phiValue)
    AssertTrue stats, "batch.stability.sp35.rebarRatio.low", steelArea / concreteArea < 0.03
    AssertClose stats, "batch.stability.sp35.rebarRatio.lowNult", nUltimate, _
        phiValue * (15.5 * concreteArea + 350# * steelArea), nUltimate * 0.0000001

    nUltimate = SP35TableNultForCircle(500#, 18, 25#, concreteArea, steelArea, phiValue)
    AssertTrue stats, "batch.stability.sp35.rebarRatio.high", steelArea / concreteArea > 0.03
    AssertClose stats, "batch.stability.sp35.rebarRatio.highNult", nUltimate, _
        phiValue * (15.5 * (concreteArea - steelArea) + 350# * steelArea), nUltimate * 0.0000001

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "Calculation.Strength.DirectState", "PR1", oldDirect
    SetProfileValue "Calculation.Strength.Capacity", "PR1", oldCapacity
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    SetSystemSetting "Stability.ElementLength", oldLength
    SetSystemSetting "Stability.Mu1", oldMu1
    SetSystemSetting "Stability.Mu2", oldMu2
    SetSystemSetting "Stability.AccidentalEccentricityMode", oldAccMode
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", oldAccPlanes
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.sp35.rebarAreaCorrection; " & Err.Description
    Resume Restore
End Sub

' Проверяет, что таблица 7.21 СП 35 интерполируется в два шага: сначала по
' ec/r между столбцами q=0, 0.25, 0.50, 1.00, затем между строками по l0/i.
Private Sub TestStabilitySP35TableInterpolationIntermediate(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldDirect As String
    Dim oldCapacity As String
    Dim oldValueSet As String
    Dim oldCode As String
    Dim oldLength As String
    Dim oldMu1 As String
    Dim oldMu2 As String
    Dim oldAccMode As String
    Dim oldAccPlanes As String
    Dim oldAccUser1 As String
    Dim oldAccUser2 As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldDirect = GetProfileValue("Calculation.Strength.DirectState", "PR1")
    oldCapacity = GetProfileValue("Calculation.Strength.Capacity", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")
    oldLength = GetSystemSetting("Stability.ElementLength")
    oldMu1 = GetSystemSetting("Stability.Mu1")
    oldMu2 = GetSystemSetting("Stability.Mu2")
    oldAccMode = GetSystemSetting("Stability.AccidentalEccentricityMode")
    oldAccPlanes = GetSystemSetting("Stability.AccidentalEccentricityPlanes")
    oldAccUser1 = GetSystemSetting("Stability.AccidentalEccentricityUser1")
    oldAccUser2 = GetSystemSetting("Stability.AccidentalEccentricityUser2")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "Calculation.Strength.DirectState", "PR1", "No"
    SetProfileValue "Calculation.Strength.Capacity", "PR1", "No"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP35"
    SetSystemSetting "Stability.Mu1", "1"
    SetSystemSetting "Stability.Mu2", "1"
    SetSystemSetting "Stability.AccidentalEccentricityMode", "User"
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", "BothPlanes"

    AssertClose stats, "batch.stability.sp35.interp.q0125", _
        SP35InterpolatedPhiM(0.125), 6.5, 0.000001
    AssertClose stats, "batch.stability.sp35.interp.q0375", _
        SP35InterpolatedPhiM(0.375), 7.5, 0.000001
    AssertClose stats, "batch.stability.sp35.interp.q0750", _
        SP35InterpolatedPhiM(0.75), 9#, 0.000001

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "Calculation.Strength.DirectState", "PR1", oldDirect
    SetProfileValue "Calculation.Strength.Capacity", "PR1", oldCapacity
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    SetSystemSetting "Stability.ElementLength", oldLength
    SetSystemSetting "Stability.Mu1", oldMu1
    SetSystemSetting "Stability.Mu2", oldMu2
    SetSystemSetting "Stability.AccidentalEccentricityMode", oldAccMode
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", oldAccPlanes
    SetSystemSetting "Stability.AccidentalEccentricityUser1", oldAccUser1
    SetSystemSetting "Stability.AccidentalEccentricityUser2", oldAccUser2
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.sp35.tableInterpolation; " & Err.Description
    Resume Restore
End Sub

' Проверяет граничное условие табличной ветви СП 35: N <= Nult, поэтому
' коэффициент запаса ровно 1.0 считается прохождением проверки.
Private Sub TestStabilitySP35TableBoundaryReservePasses(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldDirect As String
    Dim oldCapacity As String
    Dim oldValueSet As String
    Dim oldCode As String
    Dim oldLength As String
    Dim oldMu1 As String
    Dim oldMu2 As String
    Dim oldAccMode As String
    Dim oldAccPlanes As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldDirect = GetProfileValue("Calculation.Strength.DirectState", "PR1")
    oldCapacity = GetProfileValue("Calculation.Strength.Capacity", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")
    oldLength = GetSystemSetting("Stability.ElementLength")
    oldMu1 = GetSystemSetting("Stability.Mu1")
    oldMu2 = GetSystemSetting("Stability.Mu2")
    oldAccMode = GetSystemSetting("Stability.AccidentalEccentricityMode")
    oldAccPlanes = GetSystemSetting("Stability.AccidentalEccentricityPlanes")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "Calculation.Strength.DirectState", "PR1", "No"
    SetProfileValue "Calculation.Strength.Capacity", "PR1", "No"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP35"
    SetSystemSetting "Stability.ElementLength", "1000"
    SetSystemSetting "Stability.Mu1", "1"
    SetSystemSetting "Stability.Mu2", "1"
    SetSystemSetting "Stability.AccidentalEccentricityMode", "AutoWithL"
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", "BothPlanes"

    Dim concreteArea As Double
    Dim steelArea As Double
    Dim phiValue As Double
    Dim nUltimate As Double
    nUltimate = SP35TableNultForCircle(500#, 4, 12#, concreteArea, steelArea, phiValue)

    Dim section As CSectionModel
    Set section = BuildCircleStabilitySection(500#, 4, 12#)

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize section, TestMaterialProvider()
    Set batch.ProfileCatalog = TestProfileCatalog()
    batch.ApplySettings settings
    batch.SetSP35Table721 ThisWorkbook.Names.Item("rngSP35Table721").RefersToRange.Value2
    batch.AddCombination "SP35_BOUNDARY", -nUltimate, 0#, 0#, "PR1", "sp35 boundary"
    batch.Execute

    AssertClose stats, "batch.stability.sp35.boundary.reserve", batch.StabilityReserveFactor(1), 1#, 0.0000001
    AssertTrue stats, "batch.stability.sp35.boundary.status", batch.StabilityStatus(1) = "OK"
    Dim writer As CBatchResultWriter
    Set writer = New CBatchResultWriter
    writer.WriteSummary ThisWorkbook, batch
    AssertBatchSummaryReservesMatchDetailed stats, ThisWorkbook.Worksheets.Item("Results"), _
        "SP35_BOUNDARY", False, False, True

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "Calculation.Strength.DirectState", "PR1", oldDirect
    SetProfileValue "Calculation.Strength.Capacity", "PR1", oldCapacity
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    SetSystemSetting "Stability.ElementLength", oldLength
    SetSystemSetting "Stability.Mu1", oldMu1
    SetSystemSetting "Stability.Mu2", oldMu2
    SetSystemSetting "Stability.AccidentalEccentricityMode", oldAccMode
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", oldAccPlanes
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.sp35.boundaryReserve; " & Err.Description
    Resume Restore
End Sub

' Проверяет orchestrator: если устойчивость уже дала FAIL, последующие
' DirectState/Capacity/Crack не запускаются и не превращают итог в NumFail.
Private Sub TestStabilityFailStopsDownstream(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldCrack As String
    Dim oldValueSet As String
    Dim oldCode As String
    Dim oldLength As String
    Dim oldMu1 As String
    Dim oldMu2 As String
    Dim oldAccMode As String
    Dim oldAccPlanes As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldCrack = GetProfileValue("Calculation.Crack.Width", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")
    oldLength = GetSystemSetting("Stability.ElementLength")
    oldMu1 = GetSystemSetting("Stability.Mu1")
    oldMu2 = GetSystemSetting("Stability.Mu2")
    oldAccMode = GetSystemSetting("Stability.AccidentalEccentricityMode")
    oldAccPlanes = GetSystemSetting("Stability.AccidentalEccentricityPlanes")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "Calculation.Crack.Width", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP35"
    SetSystemSetting "Stability.ElementLength", "1000"
    SetSystemSetting "Stability.Mu1", "1"
    SetSystemSetting "Stability.Mu2", "1"
    SetSystemSetting "Stability.AccidentalEccentricityMode", "AutoWithL"
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", "BothPlanes"

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    batch.ApplySettings settings
    batch.SetSP35Table721 ThisWorkbook.Names.Item("rngSP35Table721").RefersToRange.Value2
    batch.AddCombination "SP35_STOP", -50000000#, 0#, 0#, "PR1", "stability stop"
    batch.Execute

    AssertTrue stats, "batch.stability.stop.stabilityFail", batch.StabilityStatus(1) = "FAIL"
    AssertTrue stats, "batch.stability.stop.overallFail", batch.Status(1) = "FAIL"
    AssertTrue stats, "batch.stability.stop.directNA", batch.DirectStateStatus(1) = "N/A"
    AssertTrue stats, "batch.stability.stop.capacityNA", batch.CapacityStatus(1) = "N/A"
    AssertTrue stats, "batch.stability.stop.crackNA", batch.CrackStatus(1) = "N/A"

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "Calculation.Crack.Width", "PR1", oldCrack
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    SetSystemSetting "Stability.ElementLength", oldLength
    SetSystemSetting "Stability.Mu1", oldMu1
    SetSystemSetting "Stability.Mu2", oldMu2
    SetSystemSetting "Stability.AccidentalEccentricityMode", oldAccMode
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", oldAccPlanes
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.stopDownstream; " & Err.Description
    Resume Restore
End Sub

' Проверяет СП 63: при гибкости l0/i <= 14 коэффициент eta не применяется и
' принимается равным 1, но критическая сила Ncr все равно считается для справки.
Private Sub TestStabilitySP63ShortSlendernessEtaIsOne(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    Dim oldLength As String
    Dim oldMu1 As String
    Dim oldMu2 As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")
    oldLength = GetSystemSetting("Stability.ElementLength")
    oldMu1 = GetSystemSetting("Stability.Mu1")
    oldMu2 = GetSystemSetting("Stability.Mu2")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP63"
    SetSystemSetting "Stability.ElementLength", "500"
    SetSystemSetting "Stability.Mu1", "1"
    SetSystemSetting "Stability.Mu2", "1"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "SP63_SHORT", -120000#, 1000000#, 0#, "PR1", "short slenderness"
    batch.Execute

    AssertTrue stats, "batch.stability.sp63.short.status", batch.StabilityStatus(1) = "OK" Or batch.StabilityStatus(1) = "FAIL"
    AssertTrue stats, "batch.stability.sp63.short.slenderness1", batch.StabilitySlenderness1(1) > 0# And batch.StabilitySlenderness1(1) <= 14#
    AssertTrue stats, "batch.stability.sp63.short.slenderness2", batch.StabilitySlenderness2(1) > 0# And batch.StabilitySlenderness2(1) <= 14#
    AssertClose stats, "batch.stability.sp63.short.eta1", batch.StabilityEta1(1), 1#, 0.000000001
    AssertClose stats, "batch.stability.sp63.short.eta2", batch.StabilityEta2(1), 1#, 0.000000001
    AssertTrue stats, "batch.stability.sp63.short.ncr", batch.StabilityNcr1(1) > 0# And batch.StabilityNcr2(1) > 0#

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    SetSystemSetting "Stability.ElementLength", oldLength
    SetSystemSetting "Stability.Mu1", oldMu1
    SetSystemSetting "Stability.Mu2", oldMu2
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.sp63.shortSlenderness; " & Err.Description
    Resume Restore
End Sub

' Проверяет, что расчет устойчивости не заменяет некорректные mu1/mu2 на 1.
' При включенной устойчивости оба коэффициента расчетной длины должны быть
' заданы положительными числами, иначе пользователь получает InputErr.
Private Sub TestStabilityInvalidMuReportsInputErr(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    Dim oldMu1 As String
    Dim oldMu2 As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")
    oldMu1 = GetSystemSetting("Stability.Mu1")
    oldMu2 = GetSystemSetting("Stability.Mu2")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP63"
    SetSystemSetting "Stability.Mu1", "0"
    SetSystemSetting "Stability.Mu2", "1"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "BAD_MU", -120000#, 1000000#, 0#, "PR1", "bad mu"
    batch.Execute

    AssertTrue stats, "batch.stability.mu.inputErr", batch.StabilityStatus(1) = "InputErr"

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    SetSystemSetting "Stability.Mu1", oldMu1
    SetSystemSetting "Stability.Mu2", oldMu2
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.invalidMu; " & Err.Description
    Resume Restore
End Sub

' Проверяет, что критичные настройки устойчивости не получают скрытые fallback
' значения. Если пользователь очистил или неверно задал параметр, включенная
' устойчивость должна вернуть InputErr через общую status policy.
Private Sub TestStabilityInvalidSettingsReportInputErr(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    Dim oldKs As String
    Dim oldPhiP As String
    Dim oldLimit As String
    Dim oldPhiMode As String
    Dim oldAccMode As String
    Dim oldAccPlanes As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")
    oldKs = GetSystemSetting("Stability.SP63.Ks")
    oldPhiP = GetSystemSetting("Stability.SP35.PhiP")
    oldLimit = GetSystemSetting("Stability.SP35.NOverNcrLimit")
    oldPhiMode = GetSystemSetting("Stability.PhiLMode")
    oldAccMode = GetSystemSetting("Stability.AccidentalEccentricityMode")
    oldAccPlanes = GetSystemSetting("Stability.AccidentalEccentricityPlanes")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"

    SetSystemSetting "Stability.Code", "SP63"
    SetSystemSetting "Stability.SP63.Ks", ""
    AssertTrue stats, "batch.stability.invalidSettings.ks", _
        StabilityStatusForCurrentSettings() = "InputErr"
    AssertTrue stats, "batch.stability.invalidSettings.ksMessage", _
        InStr(1, StabilityFirstInvalidMessageForCurrentSettings(), "Stability.SP63.Ks", vbTextCompare) > 0
    SetSystemSetting "Stability.SP63.Ks", oldKs

    SetSystemSetting "Stability.PhiLMode", "WrongPhi"
    AssertTrue stats, "batch.stability.invalidSettings.phiMode", _
        StabilityStatusForCurrentSettings() = "InputErr"
    SetSystemSetting "Stability.PhiLMode", oldPhiMode

    SetSystemSetting "Stability.AccidentalEccentricityMode", "WrongAccidental"
    AssertTrue stats, "batch.stability.invalidSettings.accMode", _
        StabilityStatusForCurrentSettings() = "InputErr"
    SetSystemSetting "Stability.AccidentalEccentricityMode", oldAccMode

    SetSystemSetting "Stability.AccidentalEccentricityPlanes", "WrongPlanes"
    AssertTrue stats, "batch.stability.invalidSettings.accPlanes", _
        StabilityStatusForCurrentSettings() = "InputErr"
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", oldAccPlanes

    SetSystemSetting "Stability.Code", "SP35"
    SetSystemSetting "Stability.SP35.PhiP", "0"
    AssertTrue stats, "batch.stability.invalidSettings.phiP", _
        StabilityStatusForCurrentSettings() = "InputErr"
    SetSystemSetting "Stability.SP35.PhiP", oldPhiP

    SetSystemSetting "Stability.SP35.NOverNcrLimit", "0"
    AssertTrue stats, "batch.stability.invalidSettings.nOverNcr", _
        StabilityStatusForCurrentSettings() = "InputErr"

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    SetSystemSetting "Stability.SP63.Ks", oldKs
    SetSystemSetting "Stability.SP35.PhiP", oldPhiP
    SetSystemSetting "Stability.SP35.NOverNcrLimit", oldLimit
    SetSystemSetting "Stability.PhiLMode", oldPhiMode
    SetSystemSetting "Stability.AccidentalEccentricityMode", oldAccMode
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", oldAccPlanes
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.invalidSettings; " & Err.Description
    Resume Restore
End Sub

' Проверяет общий диагностический Branch для СП 35. Если первая и вторая
' главные плоскости попали в разные ветви, общий Branch должен быть mixed,
' а точная детализация остается в PlaneBranch1/2.
Private Sub TestStabilitySP35MixedBranchReportsMixed(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    Dim oldLength As String
    Dim oldMu1 As String
    Dim oldMu2 As String
    Dim oldAccMode As String
    Dim oldAccPlanes As String
    Dim oldAccUser1 As String
    Dim oldAccUser2 As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")
    oldLength = GetSystemSetting("Stability.ElementLength")
    oldMu1 = GetSystemSetting("Stability.Mu1")
    oldMu2 = GetSystemSetting("Stability.Mu2")
    oldAccMode = GetSystemSetting("Stability.AccidentalEccentricityMode")
    oldAccPlanes = GetSystemSetting("Stability.AccidentalEccentricityPlanes")
    oldAccUser1 = GetSystemSetting("Stability.AccidentalEccentricityUser1")
    oldAccUser2 = GetSystemSetting("Stability.AccidentalEccentricityUser2")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP35"
    SetSystemSetting "Stability.ElementLength", "3000"
    SetSystemSetting "Stability.Mu1", "1"
    SetSystemSetting "Stability.Mu2", "1"
    SetSystemSetting "Stability.AccidentalEccentricityMode", "User"
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", "BothPlanes"
    SetSystemSetting "Stability.AccidentalEccentricityUser1", "10"
    SetSystemSetting "Stability.AccidentalEccentricityUser2", "200"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.SetSP35Table721 ThisWorkbook.Names.Item("rngSP35Table721").RefersToRange.Value2
    batch.AddCombination "SP35_MIXED", -120000#, 0#, 0#, "PR1", "sp35 mixed branch"
    batch.Execute

    AssertTrue stats, "batch.stability.sp35Mixed.status", _
        batch.StabilityStatus(1) = "OK" Or batch.StabilityStatus(1) = "FAIL"
    AssertTrue stats, "batch.stability.sp35Mixed.planeBranches", _
        StrComp(batch.StabilityPlaneBranch1(1), batch.StabilityPlaneBranch2(1), vbTextCompare) <> 0
    AssertTrue stats, "batch.stability.sp35Mixed.branch", batch.StabilityBranch(1) = "SP35-mixed"

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    SetSystemSetting "Stability.ElementLength", oldLength
    SetSystemSetting "Stability.Mu1", oldMu1
    SetSystemSetting "Stability.Mu2", oldMu2
    SetSystemSetting "Stability.AccidentalEccentricityMode", oldAccMode
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", oldAccPlanes
    SetSystemSetting "Stability.AccidentalEccentricityUser1", oldAccUser1
    SetSystemSetting "Stability.AccidentalEccentricityUser2", oldAccUser2
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.sp35MixedBranch; " & Err.Description
    Resume Restore
End Sub

' Проверяет круговое сечение с моментом только в глобальной плоскости X.
' Для почти изотропной геометрии главные оси фиксируются по X/Y, а случайный
' эксцентриситет не должен создавать расчетный момент в пустой второй плоскости.
Private Sub TestStabilityCircleMxDoesNotCreateMy(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    Dim oldAccPlanes As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")
    oldAccPlanes = GetSystemSetting("Stability.AccidentalEccentricityPlanes")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP63"
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", "OnlyMomentPlane"

    Dim geom As ISectionGeometry
    Set geom = CircleGeometry(500#, 0#, 0#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 20#, 20#, 1

    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, CircleRebars(500#, 0#, 0#, 40#, 12, 20#), "CircleStability")

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize section, TestMaterialProvider()
    Set batch.ProfileCatalog = TestProfileCatalog()
    batch.ApplySettings settings
    batch.AddCombination "CIRCLE_MX", -120000#, 1000000#, 0#, "PR1", "stability circle mx"
    batch.Execute

    AssertTrue stats, "batch.stability.circleMx.status", batch.StabilityStatus(1) = "OK" Or batch.StabilityStatus(1) = "FAIL"
    AssertTrue stats, "batch.stability.circleMx.designMx", Abs(batch.StabilityDesignMx(1)) > 0#
    AssertClose stats, "batch.stability.circleMx.noDesignMy", batch.StabilityDesignMy(1), 0#, 1#

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", oldAccPlanes
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.circleMx; " & Err.Description
    Resume Restore
End Sub

' Проверяет новую настройку плоскостей случайного эксцентриситета. В режиме
' BothPlanes случайный эксцентриситет добавляется и во вторую главную плоскость,
' даже если исходный момент был задан только в одной плоскости.
Private Sub TestStabilityAccidentalBothPlanes(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    Dim oldLength As String
    Dim oldMu1 As String
    Dim oldMu2 As String
    Dim oldAccMode As String
    Dim oldAccPlanes As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")
    oldLength = GetSystemSetting("Stability.ElementLength")
    oldMu1 = GetSystemSetting("Stability.Mu1")
    oldMu2 = GetSystemSetting("Stability.Mu2")
    oldAccMode = GetSystemSetting("Stability.AccidentalEccentricityMode")
    oldAccPlanes = GetSystemSetting("Stability.AccidentalEccentricityPlanes")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP63"
    SetSystemSetting "Stability.ElementLength", "12000"
    SetSystemSetting "Stability.Mu1", "1"
    SetSystemSetting "Stability.Mu2", "1"
    SetSystemSetting "Stability.AccidentalEccentricityMode", "AutoWithL"

    SetSystemSetting "Stability.AccidentalEccentricityPlanes", "OnlyMomentPlane"
    Dim onlyMomentPlane As Double
    onlyMomentPlane = StabilityAccidentalForSingleMoment()

    SetSystemSetting "Stability.AccidentalEccentricityPlanes", "BothPlanes"
    Dim bothPlanes As Double
    bothPlanes = StabilityAccidentalForSingleMoment()

    AssertTrue stats, "batch.stability.accidentalPlanes.onlyPositive", onlyMomentPlane > 0#
    AssertTrue stats, "batch.stability.accidentalPlanes.bothLarger", bothPlanes > onlyMomentPlane * 1.5

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    SetSystemSetting "Stability.ElementLength", oldLength
    SetSystemSetting "Stability.Mu1", oldMu1
    SetSystemSetting "Stability.Mu2", oldMu2
    SetSystemSetting "Stability.AccidentalEccentricityMode", oldAccMode
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", oldAccPlanes
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.accidentalBothPlanes; " & Err.Description
    Resume Restore
End Sub

' Проверяет, что эксцентриситет устойчивости считается относительно центра
' тяжести приведенного сечения. Пользовательская точка нагрузки остается в
' бетонном центре, поэтому несимметричная арматура должна дать реальный
' статический эксцентриситет даже при пользовательском Mx=My=0.
Private Sub TestStabilityUsesTransformedCentroidForEccentricity(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP63"

    Dim geom As ISectionGeometry
    Set geom = RectSetGeometry(250#, 550#, 600#, 250#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 50#, 20#, 1

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "RSHIFT1", 560#, 60#, 120#, 0#, "A400", "shift transformed centroid", geom
    rebars.AddBar "RSHIFT2", 560#, 150#, 120#, 0#, "A400", "shift transformed centroid", geom

    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, rebars, "RectSetStabilityCentroid")

    Dim concreteProps As CSectionPropertiesCalculator
    Set concreteProps = New CSectionPropertiesCalculator
    concreteProps.CalculateConcrete section

    Dim transformedProps As CSectionPropertiesCalculator
    Set transformedProps = New CSectionPropertiesCalculator
    transformedProps.CalculateTransformedByModuli section, 32500#, 200000#

    Dim centroidGap As Double
    centroidGap = Sqr((transformedProps.CentroidX - concreteProps.CentroidX) ^ 2 + _
        (transformedProps.CentroidY - concreteProps.CentroidY) ^ 2)
    AssertTrue stats, "batch.stability.centroid.transformedShift", centroidGap > 25#

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize section, TestMaterialProvider()
    Set batch.ProfileCatalog = TestProfileCatalog()
    batch.ApplySettings settings
    batch.AddCombination "STAB_CENTROID", -100000#, 0#, 0#, "PR1", "stability centroid"
    batch.ApplyLoadReference concreteProps.CentroidX, concreteProps.CentroidY, concreteProps.CentroidX, concreteProps.CentroidY
    batch.Execute

    AssertTrue stats, "batch.stability.centroid.status", batch.StabilityStatus(1) = "OK" Or batch.StabilityStatus(1) = "FAIL"
    AssertTrue stats, "batch.stability.centroid.eFromTransformedCenter", _
        Sqr(batch.StabilityEccentricity1(1) ^ 2 + batch.StabilityEccentricity2(1) ^ 2) > centroidGap * 0.8
    AssertClose stats, "batch.stability.centroid.loadPointMxIsUserMx", _
        batch.StabilityLoadPointMx(1), batch.UserMx(1), 0.000001
    AssertClose stats, "batch.stability.centroid.loadPointMyIsUserMy", _
        batch.StabilityLoadPointMy(1), batch.UserMy(1), 0.000001

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.centroid; " & Err.Description
    Resume Restore
End Sub

' Проверяет, что нормативный AutoWithL для случайного эксцентриситета берет
' геометрическую длину L, а не расчетную длину l0 = mu*L. Режим AutoWithMuL
' остается явной проверочной опцией и должен давать другое значение.
Private Sub TestStabilityAccidentalUsesGeometricLength(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    Dim oldLength As String
    Dim oldMu1 As String
    Dim oldMu2 As String
    Dim oldAccMode As String
    Dim oldAccPlanes As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")
    oldLength = GetSystemSetting("Stability.ElementLength")
    oldMu1 = GetSystemSetting("Stability.Mu1")
    oldMu2 = GetSystemSetting("Stability.Mu2")
    oldAccMode = GetSystemSetting("Stability.AccidentalEccentricityMode")
    oldAccPlanes = GetSystemSetting("Stability.AccidentalEccentricityPlanes")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP63"
    SetSystemSetting "Stability.ElementLength", "12000"
    SetSystemSetting "Stability.Mu1", "2"
    SetSystemSetting "Stability.Mu2", "2"
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", "OnlyMomentPlane"

    SetSystemSetting "Stability.AccidentalEccentricityMode", "AutoWithL"
    Dim byGeometricLength As Double
    byGeometricLength = StabilityAccidentalForSingleMoment()

    SetSystemSetting "Stability.AccidentalEccentricityMode", "AutoWithMuL"
    Dim byEffectiveLength As Double
    byEffectiveLength = StabilityAccidentalForSingleMoment()

    AssertClose stats, "batch.stability.accidental.L", byGeometricLength, 20#, 0.000001
    AssertClose stats, "batch.stability.accidental.muL", byEffectiveLength, 40#, 0.000001

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    SetSystemSetting "Stability.ElementLength", oldLength
    SetSystemSetting "Stability.Mu1", oldMu1
    SetSystemSetting "Stability.Mu2", oldMu2
    SetSystemSetting "Stability.AccidentalEccentricityMode", oldAccMode
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", oldAccPlanes
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.accidentalLength; " & Err.Description
    Resume Restore
End Sub

' Проверяет пользовательский режим случайного эксцентриситета: пользователь
' задает только модуль, а знак по-прежнему выбирается по направлению момента
' или настройкам для нулевого момента.
Private Sub TestStabilityAccidentalUserMode(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    Dim oldAccMode As String
    Dim oldAccUser1 As String
    Dim oldAccUser2 As String
    Dim oldAccPlanes As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")
    oldAccMode = GetSystemSetting("Stability.AccidentalEccentricityMode")
    oldAccUser1 = GetSystemSetting("Stability.AccidentalEccentricityUser1")
    oldAccUser2 = GetSystemSetting("Stability.AccidentalEccentricityUser2")
    oldAccPlanes = GetSystemSetting("Stability.AccidentalEccentricityPlanes")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP35"
    SetSystemSetting "Stability.AccidentalEccentricityMode", "User"
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", "OnlyMomentPlane"
    SetSystemSetting "Stability.AccidentalEccentricityUser1", "25"
    SetSystemSetting "Stability.AccidentalEccentricityUser2", "40"

    AssertClose stats, "batch.stability.accidental.user", StabilityAccidentalForSingleMoment(), 40#, 0.000001

    SetSystemSetting "Stability.AccidentalEccentricityPlanes", "BothPlanes"
    AssertClose stats, "batch.stability.accidental.userBoth", StabilityAccidentalForSingleMoment(), 65#, 0.000001

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    SetSystemSetting "Stability.AccidentalEccentricityMode", oldAccMode
    SetSystemSetting "Stability.AccidentalEccentricityUser1", oldAccUser1
    SetSystemSetting "Stability.AccidentalEccentricityUser2", oldAccUser2
    SetSystemSetting "Stability.AccidentalEccentricityPlanes", oldAccPlanes
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.accidentalUser; " & Err.Description
    Resume Restore
End Sub

' Проверяет исправление СП 35: выбор ветви идет по ядровому расстоянию r, а
' не по радиусу инерции i. Для круга D=500 мм ядровое расстояние примерно
' равно половине радиуса инерции, поэтому перепутать эти величины легко.
Private Sub TestStabilitySP35UsesCoreDistanceNotRadius(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    Dim oldLength As String
    Dim oldMu1 As String
    Dim oldMu2 As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")
    oldLength = GetSystemSetting("Stability.ElementLength")
    oldMu1 = GetSystemSetting("Stability.Mu1")
    oldMu2 = GetSystemSetting("Stability.Mu2")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP35"
    SetSystemSetting "Stability.ElementLength", "3000"
    SetSystemSetting "Stability.Mu1", "1"
    SetSystemSetting "Stability.Mu2", "1"

    Dim geom As ISectionGeometry
    Set geom = CircleGeometry(500#, 0#, 0#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 20#, 20#, 1

    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, CircleRebars(500#, 0#, 0#, 40#, 12, 20#), "CircleCore")

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize section, TestMaterialProvider()
    Set batch.ProfileCatalog = TestProfileCatalog()
    batch.ApplySettings settings
    batch.SetSP35Table721 ThisWorkbook.Names.Item("rngSP35Table721").RefersToRange.Value2
    batch.AddCombination "SP35_CORE", -120000#, 10000000#, 0#, "PR1", "sp35 core"
    batch.Execute

    Dim coreDistance As Double
    Dim radiusValue As Double
    coreDistance = ActiveStabilityValue(batch.StabilityCoreDistance1(1), batch.StabilityCoreDistance2(1))
    radiusValue = ActiveStabilityValue(batch.StabilityRadius1(1), batch.StabilityRadius2(1))
    AssertTrue stats, "batch.stability.sp35.core.positive", coreDistance > 0# And radiusValue > 0#
    AssertTrue stats, "batch.stability.sp35.core.notRadius", coreDistance < radiusValue * 0.75

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    SetSystemSetting "Stability.ElementLength", oldLength
    SetSystemSetting "Stability.Mu1", oldMu1
    SetSystemSetting "Stability.Mu2", oldMu2
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.sp35.coreDistance; " & Err.Description
    Resume Restore
End Sub

' Проверяет, что h для случайного эксцентриситета и delta_e берется по
' бетонному контуру. Даже если в модели есть очень крупный стержень, он
' участвует в приведенных характеристиках, но не расширяет внешний габарит h.
Private Sub TestStabilityDepthUsesConcreteContourOnly(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP63"

    Dim geom As ISectionGeometry
    Set geom = RoundedRectangleGeometry(300#, 200#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 25#, 25#, 1

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "RBIG", 0#, 90#, 160#, 0#, "A400", "large test bar", geom

    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, rebars, "ConcreteDepth")

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize section, TestMaterialProvider()
    Set batch.ProfileCatalog = TestProfileCatalog()
    batch.ApplySettings settings
    batch.AddCombination "DEPTH_CONCRETE", -120000#, 1000000#, 0#, "PR1", "depth concrete"
    batch.Execute

    AssertTrue stats, "batch.stability.depth.concreteOnly1", _
        batch.StabilityDepth1(1) > 190# And batch.StabilityDepth1(1) < 310#
    AssertTrue stats, "batch.stability.depth.concreteOnly2", _
        batch.StabilityDepth2(1) > 190# And batch.StabilityDepth2(1) < 310#

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.depthConcrete; " & Err.Description
    Resume Restore
End Sub

' Проверяет СП 63: phi_l считается как 1 + Ml1/M1 со знаком, а не по
' модулю отношения. Длительный момент противоположного знака не должен
' увеличивать phi_l сверх 1.
Private Sub TestStabilityPhiLUsesSignedSustainedMoment(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    Dim oldPhiMode As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")
    oldPhiMode = GetSystemSetting("Stability.PhiLMode")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP63"
    SetSystemSetting "Stability.PhiLMode", "Auto"

    Dim phiSameSign As Double
    phiSameSign = StabilityPhiLForDurationLoad("PHIL_SAME", -120000#, 1000000#, 0#, -60000#, 500000#, 0#)

    Dim phiOppositeSign As Double
    phiOppositeSign = StabilityPhiLForDurationLoad("PHIL_OPP", -120000#, 1000000#, 0#, -60000#, -10000000#, 0#)

    AssertTrue stats, "batch.stability.phil.sameSign", phiSameSign > 1.05
    AssertClose stats, "batch.stability.phil.oppositeSign", phiOppositeSign, 1#, 0.000001

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    SetSystemSetting "Stability.PhiLMode", oldPhiMode
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.phiLSp63; " & Err.Description
    Resume Restore
End Sub

' Проверяет, что постоянная/длительная продольная сила не обрезается к диапазону
' 0...N. Пользовательское значение передается в расчет как есть: сжатие
' хранится положительным, растяжение отрицательным в формулах устойчивости.
Private Sub TestStabilitySustainedNNotClamped(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP63"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddStabilityDurationLoad "SUSTAINED_FREE", -240000#, 500000#, 0#
    batch.AddCombination "SUSTAINED_FREE", -120000#, 1000000#, 0#, "PR1", "sustained not clamped"
    batch.Execute

    AssertClose stats, "batch.stability.sustainedN.notClamped", batch.StabilitySustainedN(1), 240000#, 0.000001

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.sustainedNNotClamped; " & Err.Description
    Resume Restore
End Sub

' Проверяет специальную ветвь СП 35 для противоположных знаков полного и
' длительного моментов: Auto не должен превращать ее в Abs-отношение.
Private Sub TestStabilitySP35OppositeMomentSigns(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    Dim oldPhiMode As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")
    oldPhiMode = GetSystemSetting("Stability.PhiLMode")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP35"
    SetSystemSetting "Stability.PhiLMode", "Auto"

    Dim phiValue As Double
    phiValue = StabilityPhiLForDurationLoad("SP35_OPP", -120000#, 10000000#, 0#, -60000#, -10000000#, 0#)

    AssertTrue stats, "batch.stability.sp35.opposite.phi", phiValue = 1# Or phiValue = 1.05

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    SetSystemSetting "Stability.PhiLMode", oldPhiMode
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.stability.sp35.oppositePhiL; " & Err.Description
    Resume Restore
End Sub

' Проверяет, что Summary заполняет значения только для выбранной методики
' устойчивости. Заголовки СП 35 и СП 63 остаются всегда, но строка LC не должна
' одновременно содержать расчетные значения двух нормативных блоков.
Private Sub TestBatchSummaryWritesOnlySelectedStabilityCode(ByRef stats As TBatchTestStats)
    Dim oldEnabled As String
    Dim oldValueSet As String
    Dim oldCode As String
    oldEnabled = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    oldValueSet = GetProfileValue("MaterialModel.Stability.ValueSet", "PR1")
    oldCode = GetSystemSetting("Stability.Code")

    On Error GoTo RestoreAndFail
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", "ULS(I)"
    SetSystemSetting "Stability.Code", "SP63"

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    batch.ApplySettings settings
    batch.AddCombination "STAB_WRITER", -120000#, 1000000#, 0#, "PR1", "stability writer"
    batch.Execute

    Dim writer As CBatchResultWriter
    Set writer = New CBatchResultWriter
    writer.WriteSummary ThisWorkbook, batch

    Dim stabilityAnchor As Object
    Set stabilityAnchor = ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange
    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    AssertTrue stats, "batch.writer.stability.anchor", stabilityAnchor.Row = 122 And stabilityAnchor.Column = 1
    AssertTrue stats, "batch.writer.stability.sp35.empty", Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row, 28).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row, 29).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row, 44).Value2)) = 0
    AssertTrue stats, "batch.writer.stability.sp35.naStatus", _
        CStr(resultsSheet.Cells.Item(stabilityAnchor.Row, 37).Value2) = "N/A" And _
        CStr(resultsSheet.Cells.Item(stabilityAnchor.Row, 43).Value2) = "N/A" And _
        CStr(resultsSheet.Cells.Item(stabilityAnchor.Row, 52).Value2) = "N/A" And _
        CStr(resultsSheet.Cells.Item(stabilityAnchor.Row, 58).Value2) = "N/A" And _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 37) And _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 43)
    AssertTrue stats, "batch.writer.stability.sp63.filled", Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row, 60).Value2)) > 0 Or _
        Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row, 71).Value2)) > 0
    AssertTrue stats, "batch.writer.stability.sp63.statusColor", _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 71) And _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 83) And _
        CellHasNoFill(resultsSheet, stabilityAnchor.Row, 60)

Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldEnabled
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", oldValueSet
    SetSystemSetting "Stability.Code", oldCode
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.writer.stability.selectedCode; " & Err.Description
    Resume Restore
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestBatchMoreThanTwentyCombinations(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()

    Dim i As Long
    For i = 1 To 24
        batch.AddCombination "LC" & CStr(i), -100000# - 2500# * i, -1800000# - 100000# * i, _
            -1200000# - 75000# * i, "PR1", "dynamic-" & CStr(i)
    Next i
    batch.Execute

    AssertTrue stats, "batch.dynamic.count", batch.Count = 24
    AssertTrue stats, "batch.dynamic.governing.index", batch.GoverningCombinationIndex >= 1 And batch.GoverningCombinationIndex <= 24
    AssertTrue stats, "batch.dynamic.last.status", Len(batch.Status(24)) > 0
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
    loads.Cells.Item(1, 5).Value2 = "ProfileId"
    loads.Cells.Item(1, 6).Value2 = "Comment"
    loads.Cells.Item(2, 1).Value2 = "BAD"
    loads.Cells.Item(2, 2).Value2 = "not-a-number"
    loads.Cells.Item(2, 3).Value2 = -1000000#
    loads.Cells.Item(2, 4).Value2 = -500000#
    loads.Cells.Item(2, 5).Value2 = "PR1"
    loads.Cells.Item(2, 6).Value2 = "invalid source row"

    Dim reader As CLoadCombinationReader
    Set reader = New CLoadCombinationReader
    reader.LoadFromWorkbook ThisWorkbook, batch
    batch.Execute

    AssertTrue stats, "batch.invalid.reader.count", batch.Count = 1
    AssertTrue stats, "batch.invalid.reader.status", batch.Status(1) = "InputErr"
End Sub

' Проверяет, что высота верхней сводки берется не только по заполненным LC,
' но и по доступным строкам rngLoadCombinations. Иначе пользователь видит
' неполную сетку при ручном расширении таблицы сочетаний.
Private Sub TestBatchSummaryRowsUseAvailableLoadRange(ByRef stats As TBatchTestStats)
    On Error GoTo Failed

    Dim originalRefersTo As String
    originalRefersTo = ThisWorkbook.Names.Item("rngLoadCombinations").RefersTo

    Dim app As Object
    Set app = ThisWorkbook.Application
    Dim oldDisplayAlerts As Boolean
    oldDisplayAlerts = app.DisplayAlerts
    app.DisplayAlerts = False

    Dim tempSheet As Object
    On Error Resume Next
    ThisWorkbook.Worksheets.Item("__tmpBatchSummaryRows").Delete
    On Error GoTo Failed
    Set tempSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets.Item(ThisWorkbook.Worksheets.Count))
    tempSheet.Name = "__tmpBatchSummaryRows"

    Dim tempRange As Object
    Set tempRange = tempSheet.Range("A1:G32")
    ThisWorkbook.Names.Item("rngLoadCombinations").RefersTo = "=" & tempRange.Address(True, True, 1, True)

    Dim writer As CBatchResultWriter
    Set writer = New CBatchResultWriter
    AssertTrue stats, "batch.writer.availableRows.dataRows31", _
        writer.RequiredDataRowsForWorkbook(ThisWorkbook, 1) = 31
    AssertTrue stats, "batch.writer.availableRows.summaryRows43", _
        writer.RequiredSummaryOutputRowsForWorkbook(ThisWorkbook, 1) = 43

CleanUp:
    On Error Resume Next
    ThisWorkbook.Names.Item("rngLoadCombinations").RefersTo = originalRefersTo
    If Not tempSheet Is Nothing Then tempSheet.Delete
    app.DisplayAlerts = oldDisplayAlerts
    On Error GoTo 0
    Exit Sub

Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.writer.availableRows; " & Err.Description
    Resume CleanUp
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestBatchSummaryWriter(ByRef stats As TBatchTestStats)
    On Error GoTo Failed

    Dim stage As String
    stage = "BuildBatchCalculator"
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    stage = "AddCombination"
    batch.AddCombination "W1", -180000#, -3500000#, -2500000#, "PR1", "writer"
    batch.AddCombination "W2", -90000#, 0#, 0#, "PR2", "crack writer"
    batch.AddCombination "WT", 10000#, 0#, 0#, "PR4", "pure tension strength writer"
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
    AssertTrue stats, "batch.writer.noResultOverlap", summaryRow + writer.RequiredSummaryOutputRowsForWorkbook(ThisWorkbook, batch.Count) - 1 < ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.Row
    AssertTrue stats, "batch.writer.rangeSize", ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Rows.Count = 1 And ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Columns.Count = 1
    AssertTrue stats, "batch.writer.strengthBlockPosition", ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange.Row = 49
    AssertTrue stats, "batch.writer.crackBlockPosition", ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange.Row = 85
    AssertTrue stats, "batch.writer.stabilityBlockPosition", ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange.Row = 122 And _
        ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.Row = 156
    AssertTrue stats, "batch.writer.meta.worstLabel", CStr(resultsSheet.Cells.Item(summaryRow, 2).Value2) = "Определяющее сочетание"
    AssertTrue stats, "batch.writer.meta.worstValue", CStr(resultsSheet.Cells.Item(summaryRow, 1).Value2) = batch.WorstCombinationID
    AssertTrue stats, "batch.writer.meta.solverCalls", CStr(resultsSheet.Cells.Item(summaryRow + 2, 2).Value2) = "Количество решений НДС"
    AssertTrue stats, "batch.writer.meta.valuesCentered", resultsSheet.Cells.Item(summaryRow, 1).HorizontalAlignment = -4108 And _
        resultsSheet.Cells.Item(summaryRow + 3, 1).HorizontalAlignment = -4108
    AssertTrue stats, "batch.writer.header.statusGroup", CStr(resultsSheet.Cells.Item(summaryRow + 5, 5).Value2) = "статус проверки"
    AssertTrue stats, "batch.writer.header.reserveGroup", CStr(resultsSheet.Cells.Item(summaryRow + 5, 16).Value2) = "минимальные коэффициенты запаса"
    AssertTrue stats, "batch.writer.header.statusStrength", CStr(resultsSheet.Cells.Item(summaryRow + 6, 5).Value2) = "прочность"
    AssertTrue stats, "batch.writer.header.reserveStrength", CStr(resultsSheet.Cells.Item(summaryRow + 6, 16).Value2) = "прочность"
    AssertTrue stats, "batch.writer.header.centered", resultsSheet.Cells.Item(summaryRow + 5, 5).HorizontalAlignment = -4108 And _
        resultsSheet.Cells.Item(summaryRow + 7, 16).HorizontalAlignment = -4108
    AssertTrue stats, "batch.writer.header.commentLeft", resultsSheet.Cells.Item(summaryRow + 10, 16).HorizontalAlignment = -4131
    AssertTrue stats, "batch.writer.header.epsilon", CStr(resultsSheet.Cells.Item(summaryRow + 7, 5).Value2) = _
        "по деформациям " & ChrW$(&H3B5)
    AssertTrue stats, "batch.writer.header.sp63StatusPlaneMerge", _
        CStr(resultsSheet.Cells.Item(summaryRow + 7, 13).MergeArea.Cells.Item(1, 1).Value2) = "Плоскость 1" And _
        resultsSheet.Cells.Item(summaryRow + 7, 13).MergeArea.Rows.Count = 2 And _
        resultsSheet.Cells.Item(summaryRow + 7, 13).MergeArea.Columns.Count = 1 And _
        CStr(resultsSheet.Cells.Item(summaryRow + 9, 13).Value2) = "Ncr/N" And _
        Not resultsSheet.Cells.Item(summaryRow + 9, 13).MergeCells
    AssertTrue stats, "batch.writer.header.sp63ReservePlaneMerge", _
        CStr(resultsSheet.Cells.Item(summaryRow + 7, 24).MergeArea.Cells.Item(1, 1).Value2) = "Плоскость 1" And _
        resultsSheet.Cells.Item(summaryRow + 7, 24).MergeArea.Rows.Count = 2 And _
        resultsSheet.Cells.Item(summaryRow + 7, 24).MergeArea.Columns.Count = 1 And _
        CStr(resultsSheet.Cells.Item(summaryRow + 9, 24).Value2) = "Ncr/N" And _
        Not resultsSheet.Cells.Item(summaryRow + 9, 24).MergeCells
    AssertTrue stats, "batch.writer.header.id", _
        CStr(resultsSheet.Cells.Item(summaryRow + 5, 1).MergeArea.Cells.Item(1, 1).Value2) = "Combination ID"
    AssertTrue stats, "batch.writer.header.idMergeRows", resultsSheet.Cells.Item(summaryRow + 5, 1).MergeArea.Rows.Count = 5
    AssertTrue stats, "batch.writer.header.commentRow", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 10, 16).Value2), "деформациям", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.worst.label", CStr(resultsSheet.Cells.Item(summaryRow + 11, 1).Value2) = "worst LC"
    AssertTrue stats, "batch.writer.worst.commentDash", CStr(resultsSheet.Cells.Item(summaryRow + 11, 2).Value2) = "-"
    AssertTrue stats, "batch.writer.worst.overallDash", CStr(resultsSheet.Cells.Item(summaryRow + 11, 3).Value2) = "-"
    AssertTrue stats, "batch.writer.worst.noNa", Not BatchSummaryWorstRowContainsText(resultsSheet, "N/A")
    AssertTrue stats, "batch.writer.worst.bold", resultsSheet.Cells.Item(summaryRow + 11, 1).Font.Bold And _
        resultsSheet.Cells.Item(summaryRow + 11, 16).Font.Bold
    AssertTrue stats, "batch.writer.statusLegend.title", BatchSummaryCellText(resultsSheet, summaryRow + 5, 27) = "Расшифровка статусов"
    AssertTrue stats, "batch.writer.statusLegend.header", _
        BatchSummaryCellText(resultsSheet, summaryRow + 6, 27) = "Статус" And _
        BatchSummaryCellText(resultsSheet, summaryRow + 6, 28) = "Описание"
    AssertTrue stats, "batch.writer.statusLegend.values", _
        BatchSummaryCellText(resultsSheet, summaryRow + 7, 27) = "OK" And _
        BatchSummaryCellText(resultsSheet, summaryRow + 10, 27) = "InputErr"
    AssertTrue stats, "batch.writer.statusLegend.mergeOnlyTitle", _
        resultsSheet.Cells.Item(summaryRow + 5, 27).MergeArea.Columns.Count = 2 And _
        Not resultsSheet.Cells.Item(summaryRow + 6, 27).MergeCells And _
        Not resultsSheet.Cells.Item(summaryRow + 7, 28).MergeCells
    AssertTrue stats, "batch.writer.statusLegend.noWrap", _
        Not resultsSheet.Cells.Item(summaryRow + 7, 28).WrapText
    AssertTrue stats, "batch.writer.statusLegend.italicValues", _
        resultsSheet.Cells.Item(summaryRow + 7, 27).Font.Italic And _
        resultsSheet.Cells.Item(summaryRow + 11, 28).Font.Italic
    AssertTrue stats, "batch.writer.statusLegend.colors", _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 7, 27) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 8, 27) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 9, 27) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 10, 27) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 11, 27)
    AssertTrue stats, "batch.writer.reserve.dataNotHeaderFill", _
        resultsSheet.Cells.Item(summaryRow + 12, 16).Interior.ColorIndex = -4142
    AssertWorstSummaryRowMatchesData stats, resultsSheet
    AssertTrue stats, "batch.writer.data.firstId", CStr(resultsSheet.Cells.Item(summaryRow + 12, 1).Value2) = "W1"
    AssertTrue stats, "batch.writer.data.capacityStatus", Len(CStr(resultsSheet.Cells.Item(summaryRow + 12, 6).Value2)) > 0
    AssertTrue stats, "batch.writer.data.capacityReserve", IsNumeric(resultsSheet.Cells.Item(summaryRow + 12, 17).Value2)
    AssertTrue stats, "batch.writer.data.statusColors", _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 3) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 5) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 6) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 7) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 8) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 9) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 10) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 11) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 12) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 13) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 14)
    Dim strengthAnchor As Object
    Set strengthAnchor = ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange
    AssertTrue stats, "batch.writer.strength.statusColors", _
        StatusCellHasExpectedFill(resultsSheet, strengthAnchor.Row, 2) And _
        StatusCellHasExpectedFill(resultsSheet, strengthAnchor.Row, 26) And _
        StatusCellHasExpectedFill(resultsSheet, strengthAnchor.Row, 42) And _
        CellHasNoFill(resultsSheet, strengthAnchor.Row, 1) And _
        CellHasNoFill(resultsSheet, strengthAnchor.Row, 10) And _
        CellHasNoFill(resultsSheet, strengthAnchor.Row, 28)
    AssertTrue stats, "batch.writer.strength.absentZonesBlank", _
        Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row + 2, 14).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row + 2, 15).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row + 2, 16).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row + 2, 18).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row + 2, 19).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row + 2, 20).Value2)) = 0
    AssertTrue stats, "batch.writer.strength.presentTensionKept", _
        Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row + 2, 17).Value2)) > 0 And _
        Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row + 2, 21).Value2)) > 0
    Dim crackAnchor As Object
    Set crackAnchor = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange
    AssertTrue stats, "batch.writer.crack.statusColors", _
        StatusCellHasExpectedFill(resultsSheet, crackAnchor.Row, 2) And _
        StatusCellHasExpectedFill(resultsSheet, crackAnchor.Row, 39) And _
        StatusCellHasExpectedFill(resultsSheet, crackAnchor.Row, 43) And _
        CellHasNoFill(resultsSheet, crackAnchor.Row, 1) And _
        CellHasNoFill(resultsSheet, crackAnchor.Row, 10) And _
        CellHasNoFill(resultsSheet, crackAnchor.Row, 40)
    AssertTrue stats, "batch.writer.crack.header.formationTitle", CStr(resultsSheet.Cells.Item(crackAnchor.Row - 4, 10).Value2) = "Момент образования трещин"
    AssertTrue stats, "batch.writer.crack.header.title", CStr(resultsSheet.Cells.Item(crackAnchor.Row - 4, 19).Value2) = "нормальные и продольные трещины"
    AssertTrue stats, "batch.writer.crack.header.mcrcNote", InStr(1, CStr(resultsSheet.Cells.Item(crackAnchor.Row - 2, 16).Value2), "моментного вектора", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.crack.header.formationStatus", _
        CStr(resultsSheet.Cells.Item(crackAnchor.Row - 3, 18).Value2) = "статус трещин" And _
        resultsSheet.Cells.Item(crackAnchor.Row - 3, 18).MergeArea.Rows.Count = 2
    AssertTrue stats, "batch.writer.crack.header.state", CStr(resultsSheet.Cells.Item(crackAnchor.Row - 1, 18).Value2) = "state"
    AssertTrue stats, "batch.writer.crack.header.normalStatusRu", _
        CStr(resultsSheet.Cells.Item(crackAnchor.Row - 1, 39).Value2) = "статус"
    AssertTrue stats, "batch.writer.crack.header.longStatusRu", _
        CStr(resultsSheet.Cells.Item(crackAnchor.Row - 1, 43).Value2) = "статус"
    AssertTrue stats, "batch.writer.crack.header.notesPlain", Not resultsSheet.Cells.Item(crackAnchor.Row - 2, 14).Font.Bold And _
        resultsSheet.Cells.Item(crackAnchor.Row - 2, 14).HorizontalAlignment = -4131
    AssertTrue stats, "batch.writer.crack.header.notesFill", CLng(resultsSheet.Cells.Item(crackAnchor.Row - 2, 14).Interior.Color) = RGB(217, 217, 217)
    AssertTrue stats, "batch.writer.crack.availableRowsBorder", _
        Len(CStr(resultsSheet.Cells.Item(crackAnchor.Row + 19, 1).Value2)) = 0 And _
        resultsSheet.Cells.Item(crackAnchor.Row + 19, 1).Borders(9).LineStyle <> -4142
    Dim stabilityAnchor As Object
    Set stabilityAnchor = ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange
    AssertTrue stats, "batch.writer.stability.statusColor", _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 2) And _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 37) And _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 43) And _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 52) And _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 58) And _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 71) And _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 83) And _
        CellHasNoFill(resultsSheet, stabilityAnchor.Row, 1)
    AssertTrue stats, "batch.writer.stability.header.summary", CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 5, 1).Value2) = _
        "Итог по расчету (с учетом " & ChrW$(&H3B7) & ")"
    AssertTrue stats, "batch.writer.stability.header.statusWidth", resultsSheet.Cells.Item(stabilityAnchor.Row - 4, 2).MergeArea.Columns.Count = 1
    AssertTrue stats, "batch.writer.stability.header.mainAxes", InStr(1, CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 4, 3).Value2), _
        "главных центральных осей", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.stability.header.noExtraTier", Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 3, 3).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 3, 5).Value2)) = 0
    AssertTrue stats, "batch.writer.stability.header.sp35", CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 5, 28).Value2) = "Расчет по СП 35"
    AssertTrue stats, "batch.writer.stability.header.sp63", CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 5, 60).Value2) = "Расчет по СП 63"
    AssertTrue stats, "batch.writer.stability.header.notes", Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 2, 3).Value2)) > 0 And _
        Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 2, 60).Value2)) > 0
    AssertTrue stats, "batch.writer.stability.header.notesPlain", Not resultsSheet.Cells.Item(stabilityAnchor.Row - 2, 3).Font.Bold And _
        resultsSheet.Cells.Item(stabilityAnchor.Row - 2, 3).HorizontalAlignment = -4131
    AssertTrue stats, "batch.writer.stability.header.notesFill", CLng(resultsSheet.Cells.Item(stabilityAnchor.Row - 2, 3).Interior.Color) = RGB(217, 217, 217)
    AssertTrue stats, "batch.writer.stability.header.sp35NcrBranch", CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 3, 32).Value2) = "при ec > r"
    AssertTrue stats, "batch.writer.stability.header.sp63PlaneHeight", resultsSheet.Cells.Item(stabilityAnchor.Row - 4, 60).MergeArea.Rows.Count = 2
    AssertTrue stats, "batch.writer.stability.header.sp35Ratio", CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 1, 36).Value2) = "0.7*Ncr/N"
    AssertClose stats, "batch.writer.stability.columnWidthA", CDbl(resultsSheet.Columns.Item(1).ColumnWidth), 10#, 0.01
    AssertClose stats, "batch.writer.stability.columnWidthN", CDbl(resultsSheet.Columns.Item(14).ColumnWidth), 10#, 0.01
    AssertClose stats, "batch.writer.stability.columnWidthAF", CDbl(resultsSheet.Columns.Item(32).ColumnWidth), 10#, 0.01
    AssertClose stats, "batch.writer.stability.columnWidthCC", CDbl(resultsSheet.Columns.Item(81).ColumnWidth), 10#, 0.01
    AssertTrue stats, "batch.writer.stability.availableRowsBorder", _
        Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row + 19, 1).Value2)) = 0 And _
        resultsSheet.Cells.Item(stabilityAnchor.Row + 19, 1).Borders(9).LineStyle <> -4142
    AssertTrue stats, "batch.writer.stability.lowerRanges", stabilityAnchor.Row + 20 < ThisWorkbook.Names.Item("rngNDMElementResults").RefersToRange.Row
    stage = "ReserveConsistency"
    AssertBatchSummaryReservesMatchDetailed stats, resultsSheet, "W1"
    AssertBatchSummaryReservesMatchDetailed stats, resultsSheet, "W2"
    AssertBatchSummaryReservesMatchDetailed stats, resultsSheet, "WT"
    Exit Sub

Failed:
    AppendLine stats, "FAIL-TRACE: TestBatchSummaryWriter." & stage & _
        "; err=" & CStr(Err.Number) & "; " & Err.Description
    Err.Raise Err.Number, Err.Source, "TestBatchSummaryWriter." & stage & ": " & Err.Description
End Sub

Private Function BatchSummaryStartRow() As Long
    BatchSummaryStartRow = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Row
End Function

' Сверяет одну строку компактной сводки с подробными таблицами Results.
Private Sub AssertBatchSummaryReservesMatchDetailed(ByRef stats As TBatchTestStats, _
        ByVal resultsSheet As Object, ByVal combinationID As String, _
        Optional ByVal checkStrength As Boolean = True, _
        Optional ByVal checkCrack As Boolean = True, _
        Optional ByVal checkStability As Boolean = True)
    Dim summaryRow As Long
    summaryRow = SummaryRowByCombination(resultsSheet, combinationID)
    Dim strengthRow As Long
    Dim crackRow As Long
    Dim stabilityRow As Long
    If checkStrength Then strengthRow = DetailedRowByCombination(resultsSheet, "rngStrengthSummaryAnchor", combinationID)
    If checkCrack Then crackRow = DetailedRowByCombination(resultsSheet, "rngCrackSummaryAnchor", combinationID)
    If checkStability Then stabilityRow = DetailedRowByCombination(resultsSheet, "rngStabilitySummaryAnchor", combinationID)

    AssertTrue stats, "batch.writer.reserve." & combinationID & ".summaryRow", summaryRow > 0
    If checkStrength Then AssertTrue stats, "batch.writer.reserve." & combinationID & ".strengthRow", strengthRow > 0
    If checkCrack Then AssertTrue stats, "batch.writer.reserve." & combinationID & ".crackRow", crackRow > 0
    If checkStability Then AssertTrue stats, "batch.writer.reserve." & combinationID & ".stabilityRow", stabilityRow > 0
    If summaryRow = 0 Then Exit Sub

    If checkStrength And strengthRow > 0 Then
        AssertOptionalReserve stats, "batch.writer.reserve." & combinationID & ".strain", _
            resultsSheet.Cells.Item(summaryRow, 16).Value2, _
            resultsSheet.Cells.Item(strengthRow, 25).Value2
        AssertOptionalReserve stats, "batch.writer.reserve." & combinationID & ".capacity", _
            resultsSheet.Cells.Item(summaryRow, 17).Value2, _
            resultsSheet.Cells.Item(strengthRow, 41).Value2
    End If
    If checkCrack And crackRow > 0 Then
        AssertOptionalReserve stats, "batch.writer.reserve." & combinationID & ".crack", _
            resultsSheet.Cells.Item(summaryRow, 18).Value2, _
            resultsSheet.Cells.Item(crackRow, 38).Value2
        AssertOptionalReserve stats, "batch.writer.reserve." & combinationID & ".longCrack", _
            resultsSheet.Cells.Item(summaryRow, 19).Value2, _
            resultsSheet.Cells.Item(crackRow, 42).Value2
    End If
    If checkStability And stabilityRow > 0 Then
        AssertOptionalReserve stats, "batch.writer.reserve." & combinationID & ".sp35p1eta", _
            resultsSheet.Cells.Item(summaryRow, 20).Value2, _
            resultsSheet.Cells.Item(stabilityRow, 36).Value2
        AssertOptionalReserve stats, "batch.writer.reserve." & combinationID & ".sp35p1table", _
            resultsSheet.Cells.Item(summaryRow, 21).Value2, _
            resultsSheet.Cells.Item(stabilityRow, 42).Value2
        AssertOptionalReserve stats, "batch.writer.reserve." & combinationID & ".sp35p2eta", _
            resultsSheet.Cells.Item(summaryRow, 22).Value2, _
            resultsSheet.Cells.Item(stabilityRow, 51).Value2
        AssertOptionalReserve stats, "batch.writer.reserve." & combinationID & ".sp35p2table", _
            resultsSheet.Cells.Item(summaryRow, 23).Value2, _
            resultsSheet.Cells.Item(stabilityRow, 57).Value2
        AssertOptionalReserve stats, "batch.writer.reserve." & combinationID & ".sp63p1", _
            resultsSheet.Cells.Item(summaryRow, 24).Value2, _
            resultsSheet.Cells.Item(stabilityRow, 70).Value2
        AssertOptionalReserve stats, "batch.writer.reserve." & combinationID & ".sp63p2", _
            resultsSheet.Cells.Item(summaryRow, 25).Value2, _
            resultsSheet.Cells.Item(stabilityRow, 82).Value2
    End If
End Sub

Private Function SummaryRowByCombination(ByVal resultsSheet As Object, ByVal combinationID As String) As Long
    Dim anchorRow As Long
    anchorRow = BatchSummaryStartRow()
    Dim rowIndex As Long
    For rowIndex = anchorRow + 12 To anchorRow + 200
        If StrComp(CStr(resultsSheet.Cells.Item(rowIndex, 1).Value2), combinationID, vbTextCompare) = 0 Then
            SummaryRowByCombination = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function DetailedRowByCombination(ByVal resultsSheet As Object, ByVal anchorName As String, _
        ByVal combinationID As String) As Long
    Dim anchor As Object
    Set anchor = ThisWorkbook.Names.Item(anchorName).RefersToRange
    Dim rowIndex As Long
    For rowIndex = anchor.Row To anchor.Row + 200
        If StrComp(CStr(resultsSheet.Cells.Item(rowIndex, anchor.Column).Value2), combinationID, vbTextCompare) = 0 Then
            DetailedRowByCombination = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

' Проверяет строку worst LC: в статусном блоке должен стоять номер LC,
' а в блоке запасов - коэффициент из той же строки LC и того же столбца проверки.
Private Sub AssertWorstSummaryRowMatchesData(ByRef stats As TBatchTestStats, ByVal resultsSheet As Object)
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "strain", 5, 16
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "capacity", 6, 17
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "crack", 7, 18
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "longCrack", 8, 19
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "sp35p1eta", 9, 20
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "sp35p1table", 10, 21
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "sp35p2eta", 11, 22
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "sp35p2table", 12, 23
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "sp63p1", 13, 24
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "sp63p2", 14, 25
End Sub

' Проверяет строку worst LC по расчетным столбцам compact summary.
Private Function BatchSummaryWorstRowContainsText(ByVal resultsSheet As Object, ByVal textValue As String) As Boolean
    Dim anchorRow As Long
    anchorRow = BatchSummaryStartRow()

    Dim columnIndex As Long
    For columnIndex = 5 To 25
        If StrComp(Trim$(CStr(resultsSheet.Cells.Item(anchorRow + 11, columnIndex).Value2)), _
                textValue, vbTextCompare) = 0 Then
            BatchSummaryWorstRowContainsText = True
            Exit Function
        End If
    Next columnIndex
End Function

' Читает отображаемое значение ячейки summary, корректно работая с объединенными
' областями шапки и правого словаря статусов.
Private Function BatchSummaryCellText(ByVal resultsSheet As Object, ByVal rowIndex As Long, ByVal columnIndex As Long) As String
    Dim cell As Object
    Set cell = resultsSheet.Cells.Item(rowIndex, columnIndex)
    If cell.MergeCells Then Set cell = cell.MergeArea.Cells.Item(1, 1)
    BatchSummaryCellText = Trim$(CStr(cell.Value2))
End Function

' Проверяет, что одна статусная ячейка окрашена по единой палитре Results.
Private Function StatusCellHasExpectedFill(ByVal resultsSheet As Object, _
        ByVal rowIndex As Long, ByVal columnIndex As Long) As Boolean
    Dim statusText As String
    statusText = CStr(resultsSheet.Cells.Item(rowIndex, columnIndex).Value2)
    StatusCellHasExpectedFill = (CLng(resultsSheet.Cells.Item(rowIndex, columnIndex).Interior.Color) = StatusFillColor(statusText))
End Function

' Проверяет, что расчетная ячейка рядом со статусом не получила статусную заливку всего блока.
Private Function CellHasNoFill(ByVal resultsSheet As Object, ByVal rowIndex As Long, ByVal columnIndex As Long) As Boolean
    CellHasNoFill = (CLng(resultsSheet.Cells.Item(rowIndex, columnIndex).Interior.ColorIndex) = -4142)
End Function

' Сверяет одну пару столбцов worst LC: номер сочетания и соответствующий запас.
Private Sub AssertWorstSummaryColumnMatchesData(ByRef stats As TBatchTestStats, _
        ByVal resultsSheet As Object, ByVal name As String, _
        ByVal statusColumn As Long, ByVal reserveColumn As Long)
    Dim anchorRow As Long
    anchorRow = BatchSummaryStartRow()
    Dim worstRow As Long
    worstRow = anchorRow + 11
    Dim worstCombination As String
    worstCombination = Trim$(CStr(resultsSheet.Cells.Item(worstRow, statusColumn).Value2))
    If Len(worstCombination) = 0 Or StrComp(worstCombination, "N/A", vbTextCompare) = 0 Then
        Dim reserveText As String
        reserveText = Trim$(CStr(resultsSheet.Cells.Item(worstRow, reserveColumn).Value2))
        AssertTrue stats, "batch.writer.worst." & name & ".notApplicable", _
            Len(reserveText) = 0 Or Not IsNumeric(reserveText)
        Exit Sub
    End If

    Dim dataRow As Long
    dataRow = SummaryRowByCombination(resultsSheet, worstCombination)
    AssertTrue stats, "batch.writer.worst." & name & ".combinationFound", dataRow > 0
    If dataRow > 0 Then
        AssertOptionalReserve stats, "batch.writer.worst." & name & ".reserve", _
            resultsSheet.Cells.Item(worstRow, reserveColumn).Value2, _
            resultsSheet.Cells.Item(dataRow, reserveColumn).Value2
    End If
End Sub

Private Sub AssertOptionalReserve(ByRef stats As TBatchTestStats, ByVal name As String, _
        ByVal summaryValue As Variant, ByVal detailedValue As Variant)
    Dim hasSummary As Boolean
    Dim hasDetailed As Boolean
    hasSummary = IsNumeric(summaryValue)
    hasDetailed = IsNumeric(detailedValue)
    If Not hasSummary And Not hasDetailed Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "PASS: " & name & "; both empty"
        Exit Sub
    End If
    AssertTrue stats, name & ".bothNumeric", hasSummary And hasDetailed
    If hasSummary And hasDetailed Then _
        AssertClose stats, name, CDbl(summaryValue), CDbl(detailedValue), MaxDouble(0.0000001, Abs(CDbl(detailedValue)) * 0.0000001)
End Sub

Private Function MaxDouble(ByVal firstValue As Double, ByVal secondValue As Double) As Double
    If firstValue > secondValue Then MaxDouble = firstValue Else MaxDouble = secondValue
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
    Set batch.ProfileCatalog = TestProfileCatalog()
    Set BuildBatchCalculator = batch
End Function

' Собирает ступенчатую тестовую модель, где локальный и глобальный варианты
' a_s дают разные расстояния. Это regression именно на передачу настройки из
' batch в расчет трещин, а не на сам геометрический ray-cast.
Private Function BuildSteppedCrackCoverBatch() As CBatchSectionCalculator
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
    Set section = BuildGeneratedSectionModel(mesh, rebars, "SteppedCoverTest")
    section.SourceType = "SteppedCoverTest"
    ' Узкий малоплощадный выступ почти не меняет равновесие, но сдвигает
    ' дальнюю растянутую опорную линию. Так batch-тест ловит именно передачу
    ' CoverDistanceMode без превращения проверки в тяжелый численный сценарий.
    section.AddConcreteElement -220#, -250#, 1#, 1, "", "", "Rectangle", 10#, 300#, 0#

    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize section, TestMaterialProvider()
    Set batch.ProfileCatalog = TestProfileCatalog()
    Set BuildSteppedCrackCoverBatch = batch
End Function

' Возвращает a_s из полного batch-расчета для указанного режима расстояния.
Private Function BatchSteppedCrackCoverA(ByVal coverMode As String) As Double
    SetSystemSetting "SLS.Crack.CoverDistanceMode", coverMode

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildSteppedCrackCoverBatch()
    batch.ApplySettings settings
    batch.AddCombination "COVER_" & coverMode, -20000#, -15000000#, 0#, "PR2", "cover distance mode"
    batch.Execute

    If batch.CrackStatus(1) <> "OK" And batch.CrackStatus(1) <> "FAIL" Then
        Err.Raise vbObjectError + 3935, "modTestBatchCalculation", _
            "Crack calculation did not run for cover mode " & coverMode & _
            ": crack=" & batch.CrackStatus(1) & "; direct=" & batch.DirectStateStatus(1) & _
            "; overall=" & batch.OverallStatus(1)
    End If
    BatchSteppedCrackCoverA = batch.CrackCoverA(1)
End Function

' Собирает Г-сечение из пользовательского примера: H1/B1/H2/B2 = 550/250/250/600,
' арматура Ø32 по всем внешним и внутренним граням. Этот сценарий нужен именно
' для проверки StateSolution при почти предельном осевом растяжении PR2.
Private Function BuildUserRectSetTensionBatch(ByRef referenceX As Double, ByRef referenceY As Double, _
        Optional ByVal providerOverride As CMaterialModelProvider = Nothing) As CBatchSectionCalculator
    Dim geom As ISectionGeometry
    Set geom = RectSetGeometry(250#, 550#, 600#, 250#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 50#, 50#, 1

    Dim rebars As CRebarLayout
    Set rebars = UserRectSetTensionRebars()

    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, rebars, "UserRectSetTension")

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
    Set batch.ProfileCatalog = TestProfileCatalog()
    Set BuildUserRectSetTensionBatch = batch
End Function

Private Function UserRectSetTensionRebars() As CRebarLayout
    Dim builder As CRectSetRebarLayoutBuilder
    Set builder = New CRectSetRebarLayoutBuilder
    Set UserRectSetTensionRebars = builder.Build(250#, 550#, 600#, 250#, 0#, 0#, _
        RectSetFaceSettingsForTest(5, 5), _
        RectSetFaceSettingsForTest(2, 2), _
        RectSetFaceSettingsForTest(2, 2), _
        RectSetFaceSettingsForTest(5, 5), _
        "A400")
End Function

Private Function RectSetFaceSettingsForTest(ByVal count1 As Long, ByVal count2 As Long) As Variant
    RectSetFaceSettingsForTest = Array(40#, 40#, 32#, 32#, count1, count2, 80#, 80#, 80#, 80#, _
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

' Возвращает значение конкретной строки профиля из rngCalculationProfiles.
' Тесты используют это для временного включения веток без изменения шаблона книги.
Private Function GetProfileValue(ByVal key As String, ByVal profileId As String) As String
    Dim profiles As Object
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange

    Dim rowIndex As Long
    rowIndex = ProfileKeyRow(profiles, key)
    GetProfileValue = CStr(profiles.Cells.Item(rowIndex, ProfileColumn(profiles, profileId)).Value2)
End Function

' Записывает значение в один профильный столбец. Восстановление старого
' значения остается на вызывающем тесте, чтобы сценарии были изолированными.
Private Sub SetProfileValue(ByVal key As String, ByVal profileId As String, ByVal value As String)
    Dim profiles As Object
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange

    profiles.Cells.Item(ProfileKeyRow(profiles, key), ProfileColumn(profiles, profileId)).Value2 = value
End Sub

' Находит строку параметра в вертикальной таблице профилей.
Private Function ProfileKeyRow(ByVal profiles As Object, ByVal key As String) As Long
    Dim rowIndex As Long
    For rowIndex = 1 To profiles.Rows.Count
        If StrComp(CStr(profiles.Cells.Item(rowIndex, 2).Value2), key, vbTextCompare) = 0 Then
            ProfileKeyRow = rowIndex
            Exit Function
        End If
    Next rowIndex
    Err.Raise vbObjectError + 3932, "modTestBatchCalculation", "Profile key not found: " & key
End Function

' Находит столбец PR1/PR2/... независимо от фактической ширины таблицы.
Private Function ProfileColumn(ByVal profiles As Object, ByVal profileId As String) As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    For rowIndex = 1 To profiles.Rows.Count
        For colIndex = 3 To profiles.Columns.Count
            If StrComp(CStr(profiles.Cells.Item(rowIndex, colIndex).Value2), profileId, vbTextCompare) = 0 Then
                ProfileColumn = colIndex
                Exit Function
            End If
        Next colIndex
    Next rowIndex
    Err.Raise vbObjectError + 3933, "modTestBatchCalculation", "Profile column not found: " & profileId
End Function

' Запускает короткий SP35-сценарий и возвращает предельную/критическую силу,
' чтобы тест мог сравнить влияние выбранного MaterialModel.Stability.ValueSet.
Private Function StabilitySP35CriticalForceForValueSet(ByVal valueSetText As String) As Double
    SetProfileValue "MaterialModel.Stability.ValueSet", "PR1", valueSetText

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.SetSP35Table721 ThisWorkbook.Names.Item("rngSP35Table721").RefersToRange.Value2
    batch.AddCombination "STAB_SP35_" & Replace$(valueSetText, "(", ""), -120000#, 0#, 0#, "PR1", "stability sp35"
    batch.Execute

    If batch.StabilityStatus(1) <> "OK" And batch.StabilityStatus(1) <> "FAIL" Then
        Err.Raise vbObjectError + 3934, "modTestBatchCalculation", _
            "SP35 stability status is not applicable: " & batch.StabilityStatus(1)
    End If
    StabilitySP35CriticalForceForValueSet = batch.StabilityCriticalForce(1)
End Function

' Выполняет короткий расчет устойчивости с одним изгибающим моментом и
' возвращает модуль фактически добавленного случайного эксцентриситета.
Private Function StabilityAccidentalForSingleMoment() As Double
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.SetSP35Table721 ThisWorkbook.Names.Item("rngSP35Table721").RefersToRange.Value2
    batch.AddCombination "ACC_E", -120000#, 1000000#, 0#, "PR1", "accidental eccentricity"
    batch.Execute

    StabilityAccidentalForSingleMoment = Abs(batch.StabilityAccidentalEcc1(1)) + _
        Abs(batch.StabilityAccidentalEcc2(1))
End Function

' Запускает минимальный LC устойчивости с текущими настройками Config.
' Нужен тестам валидации: они временно меняют одну настройку и проверяют,
' что ошибка доходит до пользовательского StabilityStatus.
Private Function StabilityStatusForCurrentSettings() As String
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.SetSP35Table721 ThisWorkbook.Names.Item("rngSP35Table721").RefersToRange.Value2
    batch.AddCombination "STAB_BAD_SETTING", -120000#, 1000000#, 0#, "PR1", "bad stability setting"
    batch.Execute

    StabilityStatusForCurrentSettings = batch.StabilityStatus(1)
End Function

' Возвращает детальную причину первой ошибки при текущих настройках Config.
' Это защищает пользовательское сообщение: короткий статус остается InputErr,
' но popup должен показывать конкретную настройку или расчетную причину.
Private Function StabilityFirstInvalidMessageForCurrentSettings() As String
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.SetSP35Table721 ThisWorkbook.Names.Item("rngSP35Table721").RefersToRange.Value2
    batch.AddCombination "STAB_BAD_SETTING", -120000#, 1000000#, 0#, "PR1", "bad stability setting"
    batch.Execute

    StabilityFirstInvalidMessageForCurrentSettings = batch.FirstInvalidInputMessage
End Function

' Запускает устойчивость с заданной длительной частью нагрузки и возвращает
' phi_l активной главной плоскости. Метод нужен тестам знаковой логики Ml/M.
Private Function StabilityPhiLForDurationLoad(ByVal combinationID As String, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, _
        ByVal sustainedN As Double, ByVal sustainedMx As Double, ByVal sustainedMy As Double) As Double
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.SetSP35Table721 ThisWorkbook.Names.Item("rngSP35Table721").RefersToRange.Value2
    batch.AddStabilityDurationLoad combinationID, sustainedN, sustainedMx, sustainedMy
    batch.AddCombination combinationID, nValue, mxValue, myValue, "PR1", "phi_l"
    batch.Execute

    StabilityPhiLForDurationLoad = ActiveMomentPlaneValue(batch, batch.StabilityPhiL1(1), batch.StabilityPhiL2(1))
End Function

' Возвращает значение из той плоскости, где фактически есть больший главный
' изгибающий момент. После явного соглашения "ось 1 = I1" у широкого
' прямоугольника глобальный Mx попадает во вторую главную плоскость, поэтому
' выбирать просто первое ненулевое значение нельзя.
Private Function ActiveMomentPlaneValue(ByVal batch As CBatchSectionCalculator, _
        ByVal firstValue As Double, ByVal secondValue As Double) As Double
    If Abs(batch.StabilityMoment1(1)) >= Abs(batch.StabilityMoment2(1)) Then
        ActiveMomentPlaneValue = firstValue
    Else
        ActiveMomentPlaneValue = secondValue
    End If
End Function

' Возвращает ненулевое значение активной плоскости. В тестах используется для
' сценариев, где нагрузка задана только в одной главной плоскости.
Private Function ActiveStabilityValue(ByVal firstValue As Double, ByVal secondValue As Double) As Double
    If Abs(firstValue) > 0.000000001 Then
        ActiveStabilityValue = Abs(firstValue)
    Else
        ActiveStabilityValue = Abs(secondValue)
    End If
End Function

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

Private Function RectSetGeometry(ByVal b1 As Double, ByVal h1 As Double, ByVal b2 As Double, ByVal h2 As Double) As ISectionGeometry
    Dim geom As CGeometryRectSet
    Set geom = New CGeometryRectSet
    geom.Initialize b1, h1, b2, h2, 0#, 0#
    Set RectSetGeometry = geom
End Function

Private Function CircleRebars(ByVal diameter As Double, ByVal centerX As Double, ByVal centerY As Double, _
        ByVal axisDistance As Double, ByVal barCount As Long, ByVal barDiameter As Double) As CRebarLayout
    Dim builder As CCircleRebarLayoutBuilder
    Set builder = New CCircleRebarLayoutBuilder
    Set CircleRebars = builder.Build(diameter, centerX, centerY, axisDistance, barCount, barDiameter, "A400")
End Function

Private Function RectSetRebars(ByVal b1 As Double, ByVal h1 As Double, ByVal b2 As Double, ByVal h2 As Double, _
        ByVal axisDistance As Double, ByVal barCount As Long, ByVal barDiameter As Double) As CRebarLayout
    Dim builder As CRectSetRebarLayoutBuilder
    Set builder = New CRectSetRebarLayoutBuilder
    Set RectSetRebars = builder.Build(b1, h1, b2, h2, 0#, 0#, _
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
    steelParameters.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#

    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.InitializeFromParameters concreteParameters, steelParameters
    Set TestMaterialProvider = provider
End Function

' Загружает профильный каталог из текущей книги для тестов batch-слоя.
' Это сохраняет тот же frontend-контракт PR1/PR2, который использует пользовательский сценарий.
Private Function TestProfileCatalog() As CCalculationProfileCatalog
    Dim profiles As CCalculationProfileCatalog
    Set profiles = New CCalculationProfileCatalog
    profiles.LoadFromWorkbook ThisWorkbook
    Set TestProfileCatalog = profiles
End Function

' Собирает круглое сечение для тестов устойчивости СП 35.
Private Function BuildCircleStabilitySection(ByVal diameter As Double, ByVal barCount As Long, _
        ByVal barDiameter As Double) As CSectionModel
    Dim geom As ISectionGeometry
    Set geom = CircleGeometry(diameter, 0#, 0#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 20#, 20#, 1

    Set BuildCircleStabilitySection = BuildGeneratedSectionModel(mesh, _
        CircleRebars(diameter, 0#, 0#, 40#, barCount, barDiameter), "CircleSP35Table")
End Function

' Возвращает Nult табличной ветви СП 35 и одновременно отдает площади,
' чтобы тест мог проверить нормативную поправку при As/Ab больше 3%.
Private Function SP35TableNultForCircle(ByVal diameter As Double, ByVal barCount As Long, _
        ByVal barDiameter As Double, ByRef concreteArea As Double, ByRef steelArea As Double, _
        ByRef phiValue As Double) As Double
    Dim section As CSectionModel
    Set section = BuildCircleStabilitySection(diameter, barCount, barDiameter)

    Dim concreteProps As CSectionPropertiesCalculator
    Set concreteProps = New CSectionPropertiesCalculator
    concreteProps.CalculateConcrete section
    concreteArea = concreteProps.Area

    Dim transformedProps As CSectionPropertiesCalculator
    Set transformedProps = New CSectionPropertiesCalculator
    transformedProps.CalculateTransformedByModuli section, 32500#, 200000#
    steelArea = transformedProps.RebarAreaTotal(section)

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize section, TestMaterialProvider()
    Set batch.ProfileCatalog = TestProfileCatalog()
    batch.ApplySettings settings
    batch.SetSP35Table721 ThisWorkbook.Names.Item("rngSP35Table721").RefersToRange.Value2
    batch.AddCombination "SP35_NULT_AREA", -100000#, 0#, 0#, "PR1", "sp35 nult area"
    batch.Execute

    phiValue = ActiveStabilityValue(batch.StabilityPhiValue1(1), batch.StabilityPhiValue2(1))
    SP35TableNultForCircle = ActiveStabilityValue(batch.StabilityNultimate1(1), batch.StabilityNultimate2(1))
End Function

' Запускает расчет по искусственной таблице 7.21 и возвращает phi_m.
' Таблица выбрана линейной, чтобы промежуточные значения проверялись точно.
Private Function SP35InterpolatedPhiM(ByVal q As Double) As Double
    Dim section As CSectionModel
    Set section = BuildCircleStabilitySection(500#, 4, 12#)

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateTransformedByModuli section, 32500#, 200000#

    Dim coreDistance As Double
    coreDistance = props.CoreDistanceAlong(section, 0#, 1#, False)
    SetSystemSetting "Stability.ElementLength", FormatNumberInvariant(15# * props.PrincipalRadius1)
    SetSystemSetting "Stability.AccidentalEccentricityUser1", FormatNumberInvariant(q * coreDistance)
    SetSystemSetting "Stability.AccidentalEccentricityUser2", FormatNumberInvariant(q * coreDistance)

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize section, TestMaterialProvider()
    Set batch.ProfileCatalog = TestProfileCatalog()
    batch.ApplySettings settings
    batch.SetSP35Table721 SP35InterpolationTable()
    batch.AddCombination "SP35_INTERP", -100000#, 0#, 0#, "PR1", "sp35 interpolation"
    batch.Execute

    SP35InterpolatedPhiM = ActiveStabilityValue(batch.StabilityPhiM1(1), batch.StabilityPhiM2(1))
End Function

' Искусственная таблица 7.21 с двумя строками по l0/i = 10 и 20.
Private Function SP35InterpolationTable() As Variant
    Dim tableData(1 To 3, 1 To 8) As Variant
    tableData(1, 1) = "l0/b": tableData(1, 2) = "l0/d": tableData(1, 3) = "l0/i"
    tableData(1, 4) = "phi_m q=0": tableData(1, 5) = "phi_m q=0.25"
    tableData(1, 6) = "phi_m q=0.50": tableData(1, 7) = "phi_m q=1.00"
    tableData(1, 8) = "phi_l"
    tableData(2, 3) = 10#: tableData(2, 4) = 1#: tableData(2, 5) = 2#
    tableData(2, 6) = 3#: tableData(2, 7) = 5#: tableData(2, 8) = 1.1
    tableData(3, 3) = 20#: tableData(3, 4) = 11#: tableData(3, 5) = 12#
    tableData(3, 6) = 13#: tableData(3, 7) = 15#: tableData(3, 8) = 1.3
    SP35InterpolationTable = tableData
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





