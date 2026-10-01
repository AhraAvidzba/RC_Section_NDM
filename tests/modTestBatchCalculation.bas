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
    Dim originalPr1Stability As String
    Dim hasOriginalPr1Stability As Boolean
    Dim t0 As Double
    t0 = Timer

    ' Общий batch-набор исторически проверяет прочность PR1 без фильтра устойчивости.
    ' Пользовательский дефолт книги при этом не меняем: значение возвращается в конце.
    originalPr1Stability = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    hasOriginalPr1Stability = True
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "No"

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
    AppendLine stats, "RUN: TestStabilityFailContinuesDownstream"
    TestStabilityFailContinuesDownstream stats
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
    AppendLine stats, "RUN: TestPR2AutoCrackStoresPreAndPostCrackStates"
    TestPR2AutoCrackStoresPreAndPostCrackStates stats
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
    AppendLine stats, "RUN: TestBatchSummaryPreservesSourceRowGaps"
    TestBatchSummaryPreservesSourceRowGaps stats
    AppendLine stats, "RUN: TestBatchSummaryRowsUseAvailableLoadRange"
    TestBatchSummaryRowsUseAvailableLoadRange stats
    AppendLine stats, "RUN: TestBatchCapacityUsesSystemSettings"
    TestBatchCapacityUsesSystemSettings stats
    AppendLine stats, "RUN: TestInvalidModeSettingsAreNotFallbacks"
    TestInvalidModeSettingsAreNotFallbacks stats
    AppendLine stats, "RUN: TestResultMetaStatusDictionary"
    TestResultMetaStatusDictionary stats
    AppendLine stats, "RUN: TestResultMetaAggregateSkipsNotApplicable"
    TestResultMetaAggregateSkipsNotApplicable stats
    AppendLine stats, "RUN: TestCombinationResultTreeDrivesDisplayFields"
    TestCombinationResultTreeDrivesDisplayFields stats
    AppendLine stats, "RUN: TestCrackAggregateIncludesCurrentStateFailure"
    TestCrackAggregateIncludesCurrentStateFailure stats
    AppendLine stats, "RUN: TestFormulaChecksDoNotCreateNumFail"
    TestFormulaChecksDoNotCreateNumFail stats
    AppendLine stats, "RUN: TestSectionStateResultStoresEquilibriumData"
    TestSectionStateResultStoresEquilibriumData stats
    AppendLine stats, "RUN: TestStateRequestEquivalenceIgnoresSolveOptions"
    TestStateRequestEquivalenceIgnoresSolveOptions stats
    AppendLine stats, "RUN: TestStateRepositoryReusesOnlyConvergedStates"
    TestStateRepositoryReusesOnlyConvergedStates stats
    AppendLine stats, "RUN: TestPrePostCrackStateNames"
    TestPrePostCrackStateNames stats
    AppendLine stats, "RUN: TestAudit02CurrentCrackedStateCacheHitCalculatesWidth"
    TestAudit02CurrentCrackedStateCacheHitCalculatesWidth stats
    AppendLine stats, "RUN: TestAudit02CanonicalResultsAndReset"
    TestAudit02CanonicalResultsAndReset stats
    AppendLine stats, "RUN: TestAudit02RepositoryContextAndRetry"
    TestAudit02RepositoryContextAndRetry stats
    AppendLine stats, "RUN: TestAudit02OnOffPhysicalResults"
    TestAudit02OnOffPhysicalResults stats
    TestAudit02InitialOffsetOnOffStatuses stats
    TestAudit03ResultLifecycle stats

    AppendLine stats, "TOTAL_BATCH: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
    RestoreBatchSuiteProfileDefaults originalPr1Stability, hasOriginalPr1Stability
    RunBatchCalculationTests = stats.Report
    Exit Function

Failed:
    RestoreBatchSuiteProfileDefaults originalPr1Stability, hasOriginalPr1Stability
    RunBatchCalculationTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

' Возвращает настройки профиля, временно измененные общим batch-прогоном.
' Отдельные stability-тесты внутри набора сами включают устойчивость и восстанавливают
' ее к этому временному тестовому базису, поэтому здесь нужен только финальный возврат.
Private Sub RestoreBatchSuiteProfileDefaults(ByVal originalPr1Stability As String, _
                                             ByVal hasOriginalPr1Stability As Boolean)
    If Not hasOriginalPr1Stability Then Exit Sub

    On Error Resume Next
    SetProfileValue "Calculation.Stability.Enabled", "PR1", originalPr1Stability
    On Error GoTo 0
End Sub

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
    AssertTrue stats, "batch.one.capacity.status", Len(batch.ResultAt(1).StrengthResult.Capacity.Status) > 0
    AssertTrue stats, "batch.one.crack.status", Len(batch.ResultAt(1).NormalCrackStatus) > 0
    AssertTrue stats, "batch.one.elapsed", batch.ElapsedSeconds >= 0#
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestProfileIdControlsLimitStateGroup(ByRef stats As TBatchTestStats)
    Dim group1 As CBatchSectionCalculator
    Set group1 = BuildBatchCalculator()
    group1.AddCombination "G1", -220000#, -7000000#, -5000000#, "PR1", "strength"
    group1.Execute

    AssertTrue stats, "batch.profileId.group1.capacity", group1.ResultAt(1).StrengthResult.Capacity.LambdaCapacity > 0#
    AssertTrue stats, "batch.profileId.group1.noCrack", group1.ResultAt(1).NormalCrackStatus = "N/A"

    Dim group2 As CBatchSectionCalculator
    Set group2 = BuildBatchCalculator()
    group2.AddCombination "G2", -220000#, -7000000#, -5000000#, "PR2", "crack"
    group2.Execute

    AssertTrue stats, "batch.profileId.group2.noCapacity", group2.ResultAt(1).StrengthResult.Capacity.Status = "N/A"
    AssertTrue stats, "batch.profileId.group2.crackedState", _
        Not group2.ResultAt(1).StateRepository.FindState(sstCrackedState) Is Nothing
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

    AssertTrue stats, "batch.settings.capacity.maxLambda", batch.ResultAt(1).StrengthResult.Capacity.Status = "NumFail"
    AssertClose stats, "batch.settings.capacity.maxLambda.noReserve", _
        batch.ResultAt(1).StrengthResult.Capacity.ReserveFactor, 0#, 0#

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
    AssertTrue stats, "batch.invalid.CapacitySolutionStrategy.status", batch.ResultAt(1).Status = "InputErr"
    AssertTrue stats, "batch.invalid.CapacitySolutionStrategy.noLambda", batch.ResultAt(1).StrengthResult.Capacity.LambdaCapacity = 0#
    AssertClose stats, "batch.invalid.CapacitySolutionStrategy.noReserve", _
        batch.ResultAt(1).StrengthResult.Capacity.ReserveFactor, 0#, 0#

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

    AssertTrue stats, "batch.governing.lambda.order", batch.ResultAt(2).StrengthResult.Capacity.LambdaCapacity > 0# And batch.ResultAt(2).StrengthResult.Capacity.LambdaCapacity < batch.ResultAt(1).StrengthResult.Capacity.LambdaCapacity
    AssertTrue stats, "batch.governing.lowestSafety", batch.GoverningCombinationID = "GOV"
    AssertTrue stats, "batch.governing.limitState", Len(batch.ResultAt(2).StrengthResult.Capacity.LimitState) > 0
    AssertTrue stats, "batch.governing.strength.status", _
        batch.ResultAt(2).StrengthResult.Capacity.Status = "OK" Or batch.ResultAt(2).StrengthResult.Capacity.Status = "FAIL" Or batch.ResultAt(2).StrengthResult.Capacity.Status = "NumFail"
End Sub

' Проверяет, что batch для чистой продольной силы автоматически выбирает
' траекторию lambda*N. Нулевые пользовательские моменты в этом режиме не
' являются ошибкой: до предела масштабируется именно продольная сила.
Private Sub TestBatchPureAxialCapacityUsesNult(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "N_ONLY", 50000#, 0#, 0#, "PR1", "pure axial"
    batch.Execute

    AssertTrue stats, "batch.nult.path", batch.ResultAt(1).StrengthResult.Capacity.PathResolved = "LambdaN"
    AssertTrue stats, "batch.nult.solutionMethod", batch.ResultAt(1).StrengthResult.Capacity.SolutionMethod = "LoadMultiplier"
    AppendLine stats, "INFO: batch.nult.status=" & batch.ResultAt(1).StrengthResult.Capacity.Status & _
        "; limitState=" & batch.ResultAt(1).StrengthResult.Capacity.LimitState & _
        "; solutionMethod=" & batch.ResultAt(1).StrengthResult.Capacity.SolutionMethod & _
        "; lambda=" & FormatNumberInvariant(batch.ResultAt(1).StrengthResult.Capacity.LambdaCapacity) & _
        "; Nult=" & FormatNumberInvariant(batch.ResultAt(1).StrengthResult.Capacity.NUltimate)
    AssertTrue stats, "batch.nult.status", batch.ResultAt(1).StrengthResult.Capacity.Status = "OK" Or batch.ResultAt(1).StrengthResult.Capacity.Status = "FAIL"
    AssertTrue stats, "batch.nult.axialUltimate", Abs(batch.ResultAt(1).StrengthResult.Capacity.NUltimate) > Abs(batch.N(1))
    AssertTrue stats, "batch.nult.noMomentUltimate", Abs(batch.ResultAt(1).StrengthResult.Capacity.MomentUltimate) < 0.000001
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

    AssertTrue stats, "batch.nEcc.path", batch.ResultAt(1).StrengthResult.Capacity.PathResolved = "LambdaN"
    AssertTrue stats, "batch.nEcc.status", batch.ResultAt(1).StrengthResult.Capacity.Status = "OK" Or batch.ResultAt(1).StrengthResult.Capacity.Status = "FAIL"
    AssertTrue stats, "batch.nEcc.method", batch.ResultAt(1).StrengthResult.Capacity.SolutionMethod = "UltimateStrain"
    AssertTrue stats, "batch.nEcc.loadPointMomentsRemainZero", _
        Abs(batch.ResultAt(1).StrengthResult.Capacity.MxUltimate) < 100000# And Abs(batch.ResultAt(1).StrengthResult.Capacity.MyUltimate) < 100000#
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
        "; capacityStatus=" & batch.ResultAt(1).StrengthResult.Capacity.Status & _
        "; limitState=" & batch.ResultAt(1).StrengthResult.Capacity.LimitState & _
        "; solutionMethod=" & batch.ResultAt(1).StrengthResult.Capacity.SolutionMethod & _
        "; lambda=" & FormatNumberInvariant(batch.ResultAt(1).StrengthResult.Capacity.LambdaCapacity) & _
        "; Nult=" & FormatNumberInvariant(batch.ResultAt(1).StrengthResult.Capacity.NUltimate)
    If batch.ResultAt(1).StrengthResult.Capacity.Status = "NumFail" Then
        AppendLine stats, "CAPACITY_DIAG: " & methodName & vbCrLf & _
            Right$(batch.ResultAt(1).StrengthResult.Capacity.DiagnosticLog, 12000)
    End If
    AssertTrue stats, "batch.rectset.n200." & methodName & ".notNumFail", batch.ResultAt(1).StrengthResult.Capacity.Status <> "NumFail"
    AssertTrue stats, "batch.rectset.n200." & methodName & ".capacityStatus", _
        batch.ResultAt(1).StrengthResult.Capacity.Status = "OK" Or batch.ResultAt(1).StrengthResult.Capacity.Status = "FAIL"
    AssertTrue stats, "batch.rectset.n200." & methodName & ".nult", Abs(batch.ResultAt(1).StrengthResult.Capacity.NUltimate) > Abs(batch.N(1))
    AssertTrue stats, "batch.rectset.n200." & methodName & ".solutionMethod", _
        batch.ResultAt(1).StrengthResult.Capacity.SolutionMethod = "LoadMultiplier"
End Sub

' Проверяет, что явный выбор lambda*Mxy масштабирует оба пользовательских
' момента при постоянной продольной силе. Это основной вариант для общего
' изгиба N + Mx + My, когда нужно найти предельный момент при заданной N.
Private Sub TestBatchExplicitCapacityLoadPathScalesMxy(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "M_BRANCH", -150000#, -3000000#, -3000000#, "PR1", "moment branch", ChrW$(&H3BB) & "*Mxy"
    batch.Execute

    AssertTrue stats, "batch.capacityPath.mxy.value", batch.ResultAt(1).CapacityPathDisplayName = ChrW$(&H3BB) & "*Mxy"
    AssertTrue stats, "batch.capacityPath.mxy.key", batch.ResultAt(1).StrengthResult.Capacity.PathResolved = "LambdaMxy"
    AssertTrue stats, "batch.capacityPath.mxy.mx", Abs(batch.ResultAt(1).StrengthResult.Capacity.MxUltimate) > 0#
    AssertTrue stats, "batch.capacityPath.mxy.my", Abs(batch.ResultAt(1).StrengthResult.Capacity.MyUltimate) > 0#
End Sub

' Проверяет, что явный выбор lambda*N масштабирует продольную силу даже при
' наличии пользовательских моментов. Моменты остаются постоянной частью
' траектории, а момент от эксцентриситета N масштабируется вместе с N.
Private Sub TestBatchExplicitCapacityLoadPathScalesNWithMoments(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "N_BRANCH", -150000#, -3000000#, -1000000#, "PR1", "axial branch", ChrW$(&H3BB) & "*N"
    batch.Execute

    AssertTrue stats, "batch.capacityPath.n.value", batch.ResultAt(1).CapacityPathDisplayName = ChrW$(&H3BB) & "*N"
    AssertTrue stats, "batch.capacityPath.n.key", batch.ResultAt(1).StrengthResult.Capacity.PathResolved = "LambdaN"
    AssertTrue stats, "batch.capacityPath.n.nult", Abs(batch.ResultAt(1).StrengthResult.Capacity.NUltimate) > 0#
    AssertTrue stats, "batch.capacityPath.n.status", batch.ResultAt(1).StrengthResult.Capacity.Status = "OK" Or batch.ResultAt(1).StrengthResult.Capacity.Status = "FAIL" Or batch.ResultAt(1).StrengthResult.Capacity.Status = "NumFail"
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
        batch.ResultAt(1).StrengthResult.Capacity.Status = "OK" Or batch.ResultAt(1).StrengthResult.Capacity.Status = "FAIL"
    AssertClose stats, "batch.capacityPath.n.ref.mxAtLoadPoint", _
        batch.ResultAt(1).StrengthResult.Capacity.MxUltimate, batch.UserMx(1), 100000#
    AssertClose stats, "batch.capacityPath.n.ref.myAtLoadPoint", _
        batch.ResultAt(1).StrengthResult.Capacity.MyUltimate, batch.UserMy(1), 100000#
    AssertTrue stats, "batch.capacityPath.n.ref.internalDiffers", _
        Abs(batch.ResultAt(1).StrengthResult.Capacity.CapacityStateMy - batch.ResultAt(1).StrengthResult.Capacity.MyUltimate) > 100000#
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

    AssertTrue stats, "batch.capacityPath.mx.key", batch.ResultAt(1).StrengthResult.Capacity.PathResolved = "LambdaMx"
    AssertTrue stats, "batch.capacityPath.my.key", batch.ResultAt(2).StrengthResult.Capacity.PathResolved = "LambdaMy"
    AssertTrue stats, "batch.capacityPath.mxy.key", batch.ResultAt(3).StrengthResult.Capacity.PathResolved = "LambdaMxy"
    AssertTrue stats, "batch.capacityPath.n.key", batch.ResultAt(4).StrengthResult.Capacity.PathResolved = "LambdaN"
    AssertTrue stats, "batch.capacityPath.all.key", batch.ResultAt(5).StrengthResult.Capacity.PathResolved = "LambdaNMxy"
    AssertTrue stats, "batch.capacityPath.noInputErr", _
        batch.ResultAt(1).StrengthResult.Capacity.Status <> "InputErr" And batch.ResultAt(2).StrengthResult.Capacity.Status <> "InputErr" And _
        batch.ResultAt(3).StrengthResult.Capacity.Status <> "InputErr" And batch.ResultAt(4).StrengthResult.Capacity.Status <> "InputErr" And _
        batch.ResultAt(5).StrengthResult.Capacity.Status <> "InputErr"
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

    AssertTrue stats, "batch.capacityPath.zero.mx.noInputErr", batch.ResultAt(1).StrengthResult.Capacity.Status <> "InputErr"
    AssertTrue stats, "batch.capacityPath.zero.my.noInputErr", batch.ResultAt(2).StrengthResult.Capacity.Status <> "InputErr"
    AssertTrue stats, "batch.capacityPath.zero.mxyMx.noInputErr", batch.ResultAt(3).StrengthResult.Capacity.Status <> "InputErr"
    AssertTrue stats, "batch.capacityPath.zero.mxyMy.noInputErr", batch.ResultAt(4).StrengthResult.Capacity.Status <> "InputErr"
    AssertTrue stats, "batch.capacityPath.zero.nmxyN.noInputErr", batch.ResultAt(5).StrengthResult.Capacity.Status <> "InputErr"
    AssertTrue stats, "batch.capacityPath.zero.nmxyMx.noInputErr", batch.ResultAt(6).StrengthResult.Capacity.Status <> "InputErr"
    AssertTrue stats, "batch.capacityPath.zero.nmxyMy.noInputErr", batch.ResultAt(7).StrengthResult.Capacity.Status <> "InputErr"
    AssertTrue stats, "batch.capacityPath.zero.nmxyEmpty.inputErr", batch.ResultAt(8).StrengthResult.Capacity.Status = "InputErr"
    AssertTrue stats, "batch.capacityPath.zero.nmxyN.pathKept", batch.ResultAt(5).StrengthResult.Capacity.PathResolved = "LambdaNMxy"
    AssertTrue stats, "batch.capacityPath.zero.nmxyMx.pathKept", batch.ResultAt(6).StrengthResult.Capacity.PathResolved = "LambdaNMxy"
    AssertTrue stats, "batch.capacityPath.zero.nmxyMy.pathKept", batch.ResultAt(7).StrengthResult.Capacity.PathResolved = "LambdaNMxy"
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

    AssertTrue stats, "batch.capacityPath.nmxyZeroM.path", batch.ResultAt(1).StrengthResult.Capacity.PathResolved = "LambdaNMxy"
    AssertTrue stats, "batch.capacityPath.nmxyZeroM.solutionMethod", batch.ResultAt(1).StrengthResult.Capacity.SolutionMethod = "LoadMultiplier"
    AssertTrue stats, "batch.capacityPath.nmxyZeroM.notNumFail", batch.ResultAt(1).StrengthResult.Capacity.Status <> "NumFail"
    AssertTrue stats, "batch.capacityPath.nmxyZeroM.nult", Abs(batch.ResultAt(1).StrengthResult.Capacity.NUltimate) > Abs(batch.N(1))

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

    AssertTrue stats, "batch.zeroMoment.capacityPath.axialDefault", batch.ResultAt(1).StrengthResult.Capacity.PathResolved = "LambdaN"
    AssertTrue stats, "batch.zeroMoment.capacityPath.notInputErr", batch.ResultAt(1).StrengthResult.Capacity.Status <> "InputErr"
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

    AssertTrue stats, "batch.zeroMoment.stability.status", batch.ResultAt(1).StabilityResult.Status = "OK" Or batch.ResultAt(1).StabilityResult.Status = "FAIL"
    AssertTrue stats, "batch.zeroMoment.stability.sign1", batch.ResultAt(1).StabilityResult.AccidentalEcc1 < 0#
    AssertTrue stats, "batch.zeroMoment.stability.sign2", batch.ResultAt(1).StabilityResult.AccidentalEcc2 < 0#
    AssertClose stats, "batch.zeroMoment.stability.finalUserMx", batch.UserMx(1), 0#, 0#
    AssertClose stats, "batch.zeroMoment.stability.finalUserMy", batch.UserMy(1), 0#, 0#
    AssertTrue stats, "batch.zeroMoment.stability.finalSummaryMxIncludesAccidental", _
        Abs(batch.ResultAt(1).StabilityResult.DesignMx) > 0#
    AssertTrue stats, "batch.zeroMoment.stability.finalSummaryMyIncludesAccidental", _
        Abs(batch.ResultAt(1).StabilityResult.DesignMy) > 0#
    AssertTrue stats, "batch.zeroMoment.stability.capacityPath", batch.ResultAt(1).StrengthResult.Capacity.PathResolved = "LambdaN"
    AssertTrue stats, "batch.zeroMoment.stability.capacityNotNumFail", batch.ResultAt(1).StrengthResult.Capacity.Status <> "NumFail"

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

    AssertTrue stats, "batch.capacityPath.invalid.status", batch.ResultAt(1).StrengthResult.Capacity.Status = "InputErr"
    AssertTrue stats, "batch.capacityPath.invalid.overall", batch.ResultAt(1).Status = "InputErr"
    AssertTrue stats, "batch.capacityPath.invalid.noPath", Len(batch.ResultAt(1).StrengthResult.Capacity.PathResolved) = 0
End Sub

' Проверяет, что физическое плато нормативной диаграммы не считается
' numerical extension. Extension начинается только после eps_ult; иначе
' нормальное состояние с запасом по capacity ошибочно получит FAIL.
Private Sub TestPR2PhysicalStateRunsCrackWithExtensionEnabled(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldExtension As String
    oldExtension = GetSystemSetting("General.DiagramExtension")

    On Error GoTo RestoreAndFail
    SetSystemSetting "General.DiagramExtension", "Yes"

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.ApplySettings settings
    batch.AddCombination "G2_PHYS", -220000#, -7000000#, -5000000#, "PR2", "physical state"
    batch.Execute

    Dim crackedState As CSectionStateResult
    Set crackedState = batch.ResultAt(1).StateRepository.FindState(sstCrackedState)

    AssertTrue stats, "batch.group2.physical.noExtension", Not batch.ResultAt(1).ExtensionUsed
    Dim stateOK As Boolean, statePhysical As Boolean
    If Not crackedState Is Nothing Then
        stateOK = (crackedState.Status = "OK")
        statePhysical = Not crackedState.ExtensionUsed
    End If
    AssertTrue stats, "batch.group2.physical.crackedStateOK", stateOK
    AssertTrue stats, "batch.group2.physical.crackBranchRuns", _
        batch.ResultAt(1).CrackResult.Longitudinal.Status <> "N/A" Or Not crackedState Is Nothing
    AssertTrue stats, "batch.group2.physical.crackedStatePhysical", _
        statePhysical

Restore:
    SetSystemSetting "General.DiagramExtension", oldExtension
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

    Dim localStatus As String
    localStatus = BatchSteppedCrackCoverStatus("NearestContour")

    Dim globalStatus As String
    globalStatus = BatchSteppedCrackCoverStatus("GlobalExtreme")

    AppendLine stats, "INFO: batch.crack.coverDistanceMode localStatus=" & localStatus & _
        "; globalStatus=" & globalStatus
    AssertTrue stats, "batch.crack.coverMode.localAccepted", localStatus <> "InputErr"
    AssertTrue stats, "batch.crack.coverMode.globalAccepted", globalStatus <> "InputErr"

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
' PreCrackState нужен для контроля состояния с работающим растянутым бетоном,
' PostCrackState - для ручной проверки sigma_s,crc после раскрытия трещины.
Private Sub TestPR2AutoCrackStoresPreAndPostCrackStates(ByRef stats As TBatchTestStats)
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
        batch.ResultAt(1).NormalCrackStatus = "OK" Or batch.ResultAt(1).NormalCrackStatus = "FAIL"
    AssertTrue stats, "batch.group2.autoMcrc.beforeState", _
        Not batch.ResultAt(1).StateRepository.FindState(sstPreCrackState) Is Nothing
    AssertTrue stats, "batch.group2.autoMcrc.afterState", _
        Not batch.ResultAt(1).StateRepository.FindState(sstPostCrackState) Is Nothing
    AssertTrue stats, "batch.group2.autoMcrc.afterRole", _
        batch.ResultAt(1).StateRepository.FindState(sstPostCrackState).MaterialModelRoleText = "CrackedState"

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

    Dim crackedState As CSectionStateResult
    Set crackedState = batch.ResultAt(1).StateRepository.FindState(sstCrackedState)
    Dim stateOK As Boolean
    If Not crackedState Is Nothing Then stateOK = (crackedState.Status = "OK")
    AssertTrue stats, "batch.group2.autoMcrcPure.crackedStateOK", _
        stateOK
    AssertTrue stats, "batch.group2.autoMcrcPure.crackCalculated", _
        batch.ResultAt(1).NormalCrackStatus = "OK" Or batch.ResultAt(1).NormalCrackStatus = "FAIL"
    AssertTrue stats, "batch.group2.autoMcrcPure.beforeState", _
        Not batch.ResultAt(1).StateRepository.FindState(sstPreCrackState) Is Nothing
    AssertTrue stats, "batch.group2.autoMcrcPure.afterState", _
        Not batch.ResultAt(1).StateRepository.FindState(sstPostCrackState) Is Nothing

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

    AppendLine stats, "INFO: batch.group2.rectsetSmallMoment crack=" & batch.ResultAt(1).NormalCrackStatus & _
        "; formed=" & CStr(batch.ResultAt(1).CrackResult.Formation.CrackFormed) & "; lambda=" & FormatNumberInvariant(batch.ResultAt(1).CrackResult.Formation.LambdaCrc)
    AssertTrue stats, "batch.group2.rectsetSmallMoment.notNumFail", batch.ResultAt(1).NormalCrackStatus <> "NumFail"
    AssertTrue stats, "batch.group2.rectsetSmallMoment.finished", _
        batch.ResultAt(1).NormalCrackStatus = "OK" Or batch.ResultAt(1).NormalCrackStatus = "FAIL"
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

    AppendLine stats, "INFO: " & prefix & "; status=" & batch.ResultAt(1).NormalCrackStatus & _
        "; method=" & batch.ResultAt(1).CrackResult.Formation.FormationMethod & _
        "; lambda=" & FormatNumberInvariant(batch.ResultAt(1).CrackResult.Formation.LambdaCrc) & _
        "; Ncrc=" & FormatNumberInvariant(batch.ResultAt(1).CrackResult.Formation.FormationNcrc) & _
        "; MxyCrc=" & FormatNumberInvariant(batch.ResultAt(1).CrackResult.Formation.Mcrc)
    If batch.ResultAt(1).CrackResult.Formation.FormationMethod <> expectedMethod Or batch.ResultAt(1).NormalCrackStatus = "NumFail" Then _
        AppendLine stats, "DIAG: " & prefix & vbCrLf & batch.DiagnosticLog

    AssertTrue stats, prefix & ".notNumFail", batch.ResultAt(1).NormalCrackStatus <> "NumFail"
    AssertTrue stats, prefix & ".method", batch.ResultAt(1).CrackResult.Formation.FormationMethod = expectedMethod
    AssertTrue stats, prefix & ".lambda", batch.ResultAt(1).CrackResult.Formation.LambdaCrc > 0# And batch.ResultAt(1).CrackResult.Formation.LambdaCrc <= 1#
    If expectNcrc Then _
        AssertTrue stats, prefix & ".ncrc", Abs(batch.ResultAt(1).CrackResult.Formation.FormationNcrc) > 0#
    If expectMcrc Then _
        AssertTrue stats, prefix & ".mcrc", batch.ResultAt(1).CrackResult.Formation.Mcrc > 0#
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
    Dim executionReport As CExecutionReport
    Set executionReport = New CExecutionReport
    executionReport.Initialize ThisWorkbook, settings
    Set batch.ExecutionReport = executionReport
    batch.Execute

    AppendLine stats, "INFO: " & prefix & "; status=" & batch.ResultAt(1).NormalCrackStatus & _
        "; lambda=" & FormatNumberInvariant(batch.ResultAt(1).CrackResult.Formation.LambdaCrc) & _
        "; method=" & batch.ResultAt(1).CrackResult.Formation.FormationMethod & _
        "; central=" & CStr(batch.ResultAt(1).CrackResult.Width.CentralTensionBranch) & _
        "; NcrcCentral=" & FormatNumberInvariant(batch.ResultAt(1).CrackResult.Formation.Ncrc) & _
        "; Ncrc=" & FormatNumberInvariant(batch.ResultAt(1).CrackResult.Formation.FormationNcrc) & _
        "; MxyCrc=" & FormatNumberInvariant(batch.ResultAt(1).CrackResult.Formation.Mcrc)
    If batch.ResultAt(1).NormalCrackStatus = "NumFail" Or _
            (expectNcrc And batch.ResultAt(1).CrackResult.Formation.FormationNcrc = 0#) Or _
            (expectMcrc And batch.ResultAt(1).CrackResult.Formation.Mcrc = 0#) Then
        AppendLine stats, "DIAG: " & prefix & vbCrLf & batch.DiagnosticLog
        AppendLine stats, "DIAG_META: " & prefix & "; formation=" & _
            MetaDebugText(batch.ResultAt(1).CrackFormationMeta) & "; current=" & _
            MetaDebugText(batch.ResultAt(1).CrackCurrentStateMeta) & "; width=" & _
            MetaDebugText(batch.ResultAt(1).CrackWidthMeta) & "; longitudinal=" & _
            MetaDebugText(batch.ResultAt(1).LongitudinalCrackMeta)
        Dim currentState As CSectionStateResult
        Set currentState = batch.ResultAt(1).StateRepository.FindState(sstCrackedState)
        If Not currentState Is Nothing Then
            AppendLine stats, "CURRENT_STATE_DIAG: " & prefix & _
                "; extensionSetting=" & GetSystemSetting("General.DiagramExtension") & _
                "; eps0=" & FormatNumberInvariant(currentState.Epsilon0) & _
                "; residualN=" & FormatNumberInvariant(currentState.ResidualN) & _
                "; residualMx=" & FormatNumberInvariant(currentState.ResidualMx) & _
                "; residualMy=" & FormatNumberInvariant(currentState.ResidualMy) & vbCrLf & currentState.DiagnosticLog
        End If
    End If
    AssertTrue stats, prefix & ".notNumFail", batch.ResultAt(1).NormalCrackStatus <> "NumFail"

    Dim writer As CCrackSummaryWriter
    Set writer = New CCrackSummaryWriter
    writer.WriteSummary ThisWorkbook, batch

    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim anchor As Object
    Set anchor = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange
    Dim detailRow As Long
    detailRow = DetailedRowByCombination(resultsSheet, "rngCrackSummaryAnchor", batch.CombinationID(1))
    AssertTrue stats, prefix & ".sheetRow", detailRow > 0
    If detailRow <= 0 Then Exit Sub

    Dim baseColumn As Long
    baseColumn = anchor.Column
    Dim isCentralAxial As Boolean
    isCentralAxial = batch.ResultAt(1).CrackResult.Width.CentralTensionBranch
    If Not isCentralAxial Then _
        AssertTrue stats, prefix & ".sheetMethod", _
            CellHasDisplayedResult(resultsSheet.Cells.Item(detailRow, baseColumn + 15 - 1).Value2)
    If expectNcrc Then
        AssertTrue stats, prefix & ".sheetNcrc", _
            CellHasDisplayedResult(resultsSheet.Cells.Item(detailRow, baseColumn + IIf(isCentralAxial, 13, 16) - 1).Value2)
    End If
    If expectMcrc Then
        AssertTrue stats, prefix & ".sheetMcrc", _
            CellHasDisplayedResult(resultsSheet.Cells.Item(detailRow, baseColumn + 17 - 1).Value2)
    End If
    If isCentralAxial Then
        AssertTrue stats, prefix & ".centralDepthsBlank", _
            Len(CStr(resultsSheet.Cells.Item(detailRow, baseColumn + 24 - 1).Value2)) = 0 And _
            Len(CStr(resultsSheet.Cells.Item(detailRow, baseColumn + 25 - 1).Value2)) = 0 And _
            Len(CStr(resultsSheet.Cells.Item(detailRow, baseColumn + 26 - 1).Value2)) = 0 And _
            Len(CStr(resultsSheet.Cells.Item(detailRow, baseColumn + 27 - 1).Value2)) = 0
    End If
End Sub

' Отличает реально выведенную величину от пустой ячейки и статуса N/A.
Private Function CellHasDisplayedResult(ByVal value As Variant) As Boolean
    Dim text As String
    text = Trim$(CStr(value))
    CellHasDisplayedResult = Len(text) > 0 And StrComp(text, "N/A", vbTextCompare) <> 0
End Function

' ДЛЯ ТЕСТОВ: возвращает компактную строку typed-result для диагностики
' регрессионных проверок, чтобы было видно, какой именно этап дал статус.
Private Function MetaDebugText(ByVal meta As CResultMeta) As String
    If meta Is Nothing Then
        MetaDebugText = "<nothing>"
    Else
        MetaDebugText = ResultInternalStatusToText(meta.InternalStatus) & "/" & _
            ResultCodeToText(meta.ResultCode) & "/" & meta.ResultComment
    End If
End Function

' Проверяет сценарий второй группы с большим осевым растяжением: расчет
' трещин не должен брать technical extension из настройки прямого НДС прочности.
Private Sub TestPR2AxialTensionBeyondPhysicalLimitUsesExtension(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldMaxIterations As String
    Dim oldLoadSteps As String
    Dim oldExtension As String
    oldMaxIterations = GetSystemSetting("Solver.MaxIterations")
    oldLoadSteps = GetSystemSetting("Solver.LoadSteps")
    oldExtension = GetSystemSetting("General.DiagramExtension")

    On Error GoTo RestoreAndFail
    SetSystemSetting "General.DiagramExtension", "Yes"
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

    AssertTrue stats, "batch.group2.extension.directPhysical", batch.ResultAt(1).StrengthResult.DirectState.Status <> "FAIL"
    AssertTrue stats, "batch.group2.extension.used", batch.ResultAt(1).ExtensionUsed
    AssertTrue stats, "batch.group2.extension.crackFail", batch.ResultAt(1).NormalCrackStatus = "FAIL"

Restore:
    SetSystemSetting "Solver.MaxIterations", oldMaxIterations
    SetSystemSetting "Solver.LoadSteps", oldLoadSteps
    SetSystemSetting "General.DiagramExtension", oldExtension
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
    oldExtension = GetSystemSetting("General.DiagramExtension")

    On Error GoTo RestoreAndFail
    SetSystemSetting "General.DiagramExtension", "Yes"
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

    AssertTrue stats, "batch.group1.extension.directFail", batch.ResultAt(1).StrengthResult.DirectState.Status = "FAIL"
    AssertTrue stats, "batch.group1.extension.used", batch.ResultAt(1).ExtensionUsed
    AssertTrue stats, "batch.group1.extension.noCrack", batch.ResultAt(1).NormalCrackStatus = "N/A"
    AssertTrue stats, "batch.group1.extension.overall", batch.ResultAt(1).Status = "FAIL"

Restore:
    SetSystemSetting "Solver.MaxIterations", oldMaxIterations
    SetSystemSetting "Solver.LoadSteps", oldLoadSteps
    SetSystemSetting "General.DiagramExtension", oldExtension
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
    oldExtension = GetSystemSetting("General.DiagramExtension")

    On Error GoTo RestoreAndFail
    SetSystemSetting "General.DiagramExtension", "Yes"
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
            " direct=" & batch.ResultAt(i).StrengthResult.DirectState.Status & _
            "; extensionUsed=" & CStr(batch.ResultAt(i).ExtensionUsed) & _
            "; epsSmax=" & FormatNumberInvariant(batch.ResultAt(i).StrengthResult.DirectState.StateResult.MaxSteelStrain)
        If batch.ResultAt(i).StrengthResult.DirectState.Status = "NumFail" Then
            AppendLine stats, "DIAG: batch.group1.nearLimit." & batch.CombinationID(i) & vbCrLf & batch.DiagnosticLog
        End If
        AssertTrue stats, "batch.group1.nearLimit.noNumFail." & batch.CombinationID(i), _
            batch.ResultAt(i).StrengthResult.DirectState.Status = "OK" Or batch.ResultAt(i).StrengthResult.DirectState.Status = "FAIL"
        If batch.ResultAt(i).StrengthResult.DirectState.Status = "FAIL" Then
            AssertTrue stats, "batch.group1.nearLimit.failUsesExtension." & batch.CombinationID(i), batch.ResultAt(i).ExtensionUsed
        End If
    Next i

Restore:
    SetSystemSetting "Solver.MaxIterations", oldMaxIterations
    SetSystemSetting "Solver.LoadSteps", oldLoadSteps
    SetSystemSetting "General.DiagramExtension", oldExtension
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
    oldExtension = GetSystemSetting("General.DiagramExtension")

    On Error GoTo RestoreAndFail
    SetSystemSetting "General.DiagramExtension", "Yes"
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
            " direct=" & batch.ResultAt(i).StrengthResult.DirectState.Status & _
            "; extensionUsed=" & CStr(batch.ResultAt(i).ExtensionUsed) & _
            "; epsCmin=" & FormatNumberInvariant(batch.ResultAt(i).StrengthResult.DirectState.StateResult.MinConcreteStrain)
        If batch.ResultAt(i).StrengthResult.DirectState.Status = "NumFail" Then
            AppendLine stats, "DIAG: batch.group1.compressionNearLimit." & batch.CombinationID(i) & vbCrLf & batch.DiagnosticLog
        End If
        AssertTrue stats, "batch.group1.compressionNearLimit.noNumFail." & batch.CombinationID(i), _
            batch.ResultAt(i).StrengthResult.DirectState.Status = "OK" Or batch.ResultAt(i).StrengthResult.DirectState.Status = "FAIL"
        If batch.ResultAt(i).StrengthResult.DirectState.Status = "FAIL" Then
            AssertTrue stats, "batch.group1.compressionNearLimit.failUsesExtension." & batch.CombinationID(i), batch.ResultAt(i).ExtensionUsed
        End If
    Next i

Restore:
    SetSystemSetting "Solver.MaxIterations", oldMaxIterations
    SetSystemSetting "Solver.LoadSteps", oldLoadSteps
    SetSystemSetting "General.DiagramExtension", oldExtension
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
    oldExtension = GetSystemSetting("General.DiagramExtension")

    On Error GoTo RestoreAndFail
    SetSystemSetting "General.DiagramExtension", "Yes"
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
            " direct=" & batch.ResultAt(i).StrengthResult.DirectState.Status & _
            "; extensionUsed=" & CStr(batch.ResultAt(i).ExtensionUsed) & _
            "; epsCmin=" & FormatNumberInvariant(batch.ResultAt(i).StrengthResult.DirectState.StateResult.MinConcreteStrain) & _
            "; epsSmax=" & FormatNumberInvariant(batch.ResultAt(i).StrengthResult.DirectState.StateResult.MaxSteelStrain)
        If i = firstOkIndex Then
            AssertTrue stats, testPrefix & ".lastPhysicalOK." & batch.CombinationID(i), _
                batch.ResultAt(i).StrengthResult.DirectState.Status = "OK" And Not batch.ResultAt(i).ExtensionUsed
        Else
            If batch.ResultAt(i).StrengthResult.DirectState.Status = "NumFail" Then
                AppendLine stats, "DIAG: " & testPrefix & "." & batch.CombinationID(i) & vbCrLf & batch.DiagnosticLog
            End If
            AssertTrue stats, testPrefix & ".stableFail." & batch.CombinationID(i), _
                batch.ResultAt(i).StrengthResult.DirectState.Status = "FAIL"
            AssertTrue stats, testPrefix & ".usesExtension." & batch.CombinationID(i), batch.ResultAt(i).ExtensionUsed
        End If
    Next i

Restore:
    SetSystemSetting "Solver.MaxIterations", oldMaxIterations
    SetSystemSetting "Solver.LoadSteps", oldLoadSteps
    SetSystemSetting "General.DiagramExtension", oldExtension
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

' Проверяет симметричный для сжатия сценарий второй группы: текущее
' CrackedState может искать равновесие через extension, но после этого
' crack-ветка должна остановиться со статусом FAIL.
Private Sub TestPR2AxialCompressionBeyondPhysicalLimitUsesExtension(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldMaxIterations As String
    Dim oldLoadSteps As String
    Dim oldExtension As String
    oldMaxIterations = GetSystemSetting("Solver.MaxIterations")
    oldLoadSteps = GetSystemSetting("Solver.LoadSteps")
    oldExtension = GetSystemSetting("General.DiagramExtension")

    On Error GoTo RestoreAndFail
    SetSystemSetting "General.DiagramExtension", "Yes"
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

    AssertTrue stats, "batch.group2.extension.compression.directPhysical", batch.ResultAt(1).StrengthResult.DirectState.Status <> "FAIL"
    AssertTrue stats, "batch.group2.extension.compression.used", batch.ResultAt(1).ExtensionUsed
    AssertTrue stats, "batch.group2.extension.compression.crackFail", batch.ResultAt(1).NormalCrackStatus = "FAIL"
    Dim formation As CCrackFormationResult
    Set formation = batch.ResultAt(1).CrackResult.Formation
    AssertTrue stats, "audit02.formation.physicalBlock.exists", Not formation Is Nothing
    If Not formation Is Nothing Then
        If formation.ResultMeta.InternalStatus <> rsCheckFailed Then _
            AppendLine stats, "DIAGNOSTIC physicalBlock: status=" & CStr(formation.ResultMeta.InternalStatus) & _
                "; code=" & CStr(formation.ResultMeta.ResultCode) & "; reason=" & formation.ResultMeta.ResultComment & _
                vbCrLf & formation.DiagnosticLog & vbCrLf & batch.DiagnosticLog
        AssertTrue stats, "audit02.formation.physicalBlock.status", formation.ResultMeta.InternalStatus = rsCheckFailed
        AssertTrue stats, "audit02.formation.physicalBlock.code", formation.ResultMeta.ResultCode = rcPhysicalLimitExceeded
        AssertTrue stats, "audit02.formation.physicalBlock.noPoint", Not formation.HasLimitPoint
        AssertTrue stats, "audit02.formation.physicalBlock.noPreState", formation.PreCrackState Is Nothing
    End If

Restore:
    SetSystemSetting "Solver.MaxIterations", oldMaxIterations
    SetSystemSetting "Solver.LoadSteps", oldLoadSteps
    SetSystemSetting "General.DiagramExtension", oldExtension
    Exit Sub

RestoreAndFail:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.group2.extension.compression; " & Err.Description
    Resume Restore
End Sub

' Проверяет неосевую перегрузку второй группы. При N + Mx + My текущее
' CrackedState может получить равновесие через extension, но расчет ширины
' трещины после такого состояния не продолжается.
Private Sub TestPR2BendingBeyondPhysicalLimitUsesExtension(ByRef stats As TBatchTestStats)
    Dim oldMode As String
    Dim oldCrackEnabled As String
    Dim oldMaxIterations As String
    Dim oldLoadSteps As String
    Dim oldExtension As String
    oldMaxIterations = GetSystemSetting("Solver.MaxIterations")
    oldLoadSteps = GetSystemSetting("Solver.LoadSteps")
    oldExtension = GetSystemSetting("General.DiagramExtension")

    On Error GoTo RestoreAndFail
    SetSystemSetting "General.DiagramExtension", "Yes"
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

    AssertTrue stats, "batch.group2.extension.bending.directPhysical", batch.ResultAt(1).StrengthResult.DirectState.Status <> "FAIL"
    AssertTrue stats, "batch.group2.extension.bending.used", batch.ResultAt(1).ExtensionUsed
    AssertTrue stats, "batch.group2.extension.bending.crackFail", batch.ResultAt(1).NormalCrackStatus = "FAIL"

Restore:
    SetSystemSetting "Solver.MaxIterations", oldMaxIterations
    SetSystemSetting "Solver.LoadSteps", oldLoadSteps
    SetSystemSetting "General.DiagramExtension", oldExtension
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
        batch.ResultAt(2).StrengthResult.Capacity.LambdaCapacity > 0# And batch.ResultAt(2).StrengthResult.Capacity.LambdaCapacity < batch.ResultAt(1).StrengthResult.Capacity.LambdaCapacity
    AssertTrue stats, "batch.governing.profiles.pr2Skipped", batch.ResultAt(3).StrengthResult.Capacity.Status = "N/A"
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

    AssertTrue stats, "batch.profile.capacity.pr1.runs", batch.ResultAt(1).StrengthResult.Capacity.Status <> "N/A"
    AssertTrue stats, "batch.profile.capacity.pr2.skipped", batch.ResultAt(2).StrengthResult.Capacity.Status = "N/A"

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

    AssertTrue stats, "batch.direct.lambda.positive", batch.ResultAt(1).StrengthResult.Capacity.LambdaCapacity > 0# And batch.ResultAt(2).StrengthResult.Capacity.LambdaCapacity > 0#
    AssertTrue stats, "batch.direct.capacity.ok", batch.ResultAt(1).StrengthResult.Capacity.Status = "OK" And batch.ResultAt(2).StrengthResult.Capacity.Status = "OK"
    AssertTrue stats, "batch.direct.crack.na", StrComp(batch.ResultAt(1).NormalCrackStatus, "N/A", vbTextCompare) = 0
    AssertTrue stats, "batch.direct.state.status", Len(batch.ResultAt(1).StrengthResult.DirectState.Status) > 0 And Len(batch.ResultAt(2).StrengthResult.DirectState.Status) > 0
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

    AppendLine stats, "INFO: batch.group1.tensionMoment.direct=" & batch.ResultAt(1).StrengthResult.DirectState.Status & _
        "; capacity=" & batch.ResultAt(1).StrengthResult.Capacity.Status & _
        "; epsCmin=" & FormatNumberInvariant(batch.ResultAt(1).StrengthResult.DirectState.StateResult.MinConcreteStrain) & _
        "; epsSmax=" & FormatNumberInvariant(batch.ResultAt(1).StrengthResult.DirectState.StateResult.MaxSteelStrain)
    If batch.ResultAt(1).StrengthResult.DirectState.Status = "NumFail" Then
        AppendLine stats, "DIAG: batch.group1.tensionMoment" & vbCrLf & batch.DiagnosticLog
    End If
    AssertTrue stats, "batch.group1.tensionMoment.directNotNumFail", _
        batch.ResultAt(1).StrengthResult.DirectState.Status = "OK" Or batch.ResultAt(1).StrengthResult.DirectState.Status = "FAIL"
    AssertTrue stats, "batch.group1.tensionMoment.capacityNotNumFail", _
        batch.ResultAt(1).StrengthResult.Capacity.Status = "OK" Or batch.ResultAt(1).StrengthResult.Capacity.Status = "FAIL"

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

    AssertTrue stats, "batch.pr1.capacityRuns", batch.ResultAt(1).StrengthResult.Capacity.Status <> "N/A"
    AssertTrue stats, "batch.pr1.directRuns", batch.ResultAt(1).StrengthResult.DirectState.Status <> "N/A"
    AssertTrue stats, "batch.pr1.crackNA", batch.ResultAt(1).NormalCrackStatus = "N/A"
    AssertTrue stats, "batch.pr1.hasState", batch.ResultAt(1).StrengthResult.DirectState.StateAvailable And batch.StateAvailableCount > 0

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

    AssertTrue stats, "batch.direct.failure.status", batch.ResultAt(1).Status = "NumFail" Or batch.ResultAt(1).Status = "FAIL"
    AssertTrue stats, "batch.direct.failure.directStatus", batch.ResultAt(1).StrengthResult.DirectState.Status = "NumFail" Or batch.ResultAt(1).StrengthResult.DirectState.Status = "FAIL"
    AssertTrue stats, "batch.direct.failure.crack", batch.ResultAt(1).NormalCrackStatus = "N/A"

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
        batch.ResultAt(1).CrackResult.Longitudinal.Status = "OK" Or batch.ResultAt(1).CrackResult.Longitudinal.Status = "FAIL"
    AssertTrue stats, "batch.longCrack.sigma", batch.ResultAt(1).CrackResult.Longitudinal.MaxCompressionStress > 0#
    AssertClose stats, "batch.longCrack.rbMc2", batch.ResultAt(1).CrackResult.Longitudinal.RbMc2, 14.6, 0.000000001
    AssertClose stats, "batch.longCrack.util", batch.ResultAt(1).CrackResult.Longitudinal.Utilization, _
        batch.ResultAt(1).CrackResult.Longitudinal.MaxCompressionStress / batch.ResultAt(1).CrackResult.Longitudinal.RbMc2, 0.000000001
End Sub

' Проверяет нормативную область применения: продольные трещины являются
' проверкой трещиностойкости по PR2, поэтому для PR1 этот блок остается
' N/A и не влияет на прочностной результат.
Private Sub TestLongitudinalCrackSkippedForPR1(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "LONG_G1", -90000#, 0#, 0#, "PR1", "longitudinal group1"
    batch.Execute

    AssertTrue stats, "batch.longCrack.group1.na", batch.ResultAt(1).CrackResult.Longitudinal.Status = "N/A"
    AssertClose stats, "batch.longCrack.group1.noRbMc2", batch.ResultAt(1).CrackResult.Longitudinal.RbMc2, 0#, 0.000000001
    AssertClose stats, "batch.longCrack.group1.noUtil", batch.ResultAt(1).CrackResult.Longitudinal.Utilization, 0#, 0.000000001
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

    AssertTrue stats, "batch.stability.sp63.status", batch.ResultAt(1).StabilityResult.Status = "OK" Or batch.ResultAt(1).StabilityResult.Status = "FAIL"
    AssertTrue stats, "batch.stability.sp63.ncr", batch.ResultAt(1).StabilityResult.CriticalForce > 0#
    AssertTrue stats, "batch.stability.sp63.eta", batch.ResultAt(1).StabilityResult.Eta1 > 0# Or batch.ResultAt(1).StabilityResult.Eta2 > 0#

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

    AssertTrue stats, "batch.stability.tension.na", batch.ResultAt(1).StabilityResult.Status = "N/A"
    AssertClose stats, "batch.stability.tension.noNcr", batch.ResultAt(1).StabilityResult.CriticalForce, 0#, 0.000000001
    AssertTrue stats, "batch.stability.tension.overall", batch.ResultAt(1).Status <> "InputErr"

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
    SetSystemSetting "Stability.Code", "SP35"
    SetSystemSetting "Stability.ElementLength", "1000"
    SetSystemSetting "Stability.Mu1", "1"
    SetSystemSetting "Stability.Mu2", "1"

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
    SetSystemSetting "Stability.ElementLength", oldLength
    SetSystemSetting "Stability.Mu1", oldMu1
    SetSystemSetting "Stability.Mu2", oldMu2
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
        batch.ResultAt(1).StabilityResult.Status = "OK" Or batch.ResultAt(1).StabilityResult.Status = "FAIL"
    AssertClose stats, "batch.stability.sp35.table.ncr1.na", batch.ResultAt(1).StabilityResult.Ncr1, 0#, 0.000000001
    AssertClose stats, "batch.stability.sp35.table.ncr2.na", batch.ResultAt(1).StabilityResult.Ncr2, 0#, 0.000000001
    AssertTrue stats, "batch.stability.sp35.table.nult1", batch.ResultAt(1).StabilityResult.Nultimate1 > 0#
    AssertTrue stats, "batch.stability.sp35.table.nult2", batch.ResultAt(1).StabilityResult.Nultimate2 > 0#

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
        batch.ResultAt(1).StabilityResult.PlaneBranch1 = "SP35-table" Or batch.ResultAt(1).StabilityResult.PlaneBranch2 = "SP35-table"
    AssertClose stats, "batch.stability.sp35.tablePhi.phiL1", _
        batch.ResultAt(1).StabilityResult.PhiL1, batch.ResultAt(1).StabilityResult.PhiLTable1, 0.000000001
    AssertTrue stats, "batch.stability.sp35.tablePhi.notTwo", batch.ResultAt(1).StabilityResult.PhiL1 < 1.99

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

    AssertClose stats, "batch.stability.sp35.boundary.reserve", batch.ResultAt(1).StabilityResult.ReserveFactor, 1#, 0.0000001
    AssertTrue stats, "batch.stability.sp35.boundary.status", batch.ResultAt(1).StabilityResult.Status = "OK"
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

' Проверяет orchestrator: если устойчивость дала FAIL, остальные запрошенные
' расчеты все равно выполняются, чтобы пользователь видел полную картину LC.
Private Sub TestStabilityFailContinuesDownstream(ByRef stats As TBatchTestStats)
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
    SetSystemSetting "Stability.ElementLength", "100000"
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
    batch.AddCombination "SP35_CONTINUE", -1000000#, 0#, 0#, "PR1", "stability fail but continue"
    batch.Execute

    AssertTrue stats, "batch.stability.continue.stabilityFail", batch.ResultAt(1).StabilityResult.Status = "FAIL"
    AssertTrue stats, "batch.stability.continue.overallFail", batch.ResultAt(1).Status = "FAIL"
    AssertTrue stats, "batch.stability.continue.directCalculated", batch.ResultAt(1).StrengthResult.DirectState.Status <> "N/A"
    AssertTrue stats, "batch.stability.continue.capacityCalculated", batch.ResultAt(1).StrengthResult.Capacity.Status <> "N/A"
    AssertTrue stats, "batch.stability.continue.longitudinalCalculated", batch.ResultAt(1).CrackResult.Longitudinal.Status <> "N/A"

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
    AppendLine stats, "FAIL: batch.stability.continueDownstream; " & Err.Description
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

    AssertTrue stats, "batch.stability.sp63.short.status", batch.ResultAt(1).StabilityResult.Status = "OK" Or batch.ResultAt(1).StabilityResult.Status = "FAIL"
    AssertTrue stats, "batch.stability.sp63.short.slenderness1", batch.ResultAt(1).StabilityResult.Slenderness1 > 0# And batch.ResultAt(1).StabilityResult.Slenderness1 <= 14#
    AssertTrue stats, "batch.stability.sp63.short.slenderness2", batch.ResultAt(1).StabilityResult.Slenderness2 > 0# And batch.ResultAt(1).StabilityResult.Slenderness2 <= 14#
    AssertClose stats, "batch.stability.sp63.short.eta1", batch.ResultAt(1).StabilityResult.Eta1, 1#, 0.000000001
    AssertClose stats, "batch.stability.sp63.short.eta2", batch.ResultAt(1).StabilityResult.Eta2, 1#, 0.000000001
    AssertTrue stats, "batch.stability.sp63.short.ncr", batch.ResultAt(1).StabilityResult.Ncr1 > 0# And batch.ResultAt(1).StabilityResult.Ncr2 > 0#

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

    AssertTrue stats, "batch.stability.mu.inputErr", batch.ResultAt(1).StabilityResult.Status = "InputErr"

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
        batch.ResultAt(1).StabilityResult.Status = "OK" Or batch.ResultAt(1).StabilityResult.Status = "FAIL"
    AssertTrue stats, "batch.stability.sp35Mixed.planeBranches", _
        StrComp(batch.ResultAt(1).StabilityResult.PlaneBranch1, batch.ResultAt(1).StabilityResult.PlaneBranch2, vbTextCompare) <> 0
    AssertTrue stats, "batch.stability.sp35Mixed.branch", batch.ResultAt(1).StabilityResult.Branch = "SP35-mixed"

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

    AssertTrue stats, "batch.stability.circleMx.status", batch.ResultAt(1).StabilityResult.Status = "OK" Or batch.ResultAt(1).StabilityResult.Status = "FAIL"
    AssertTrue stats, "batch.stability.circleMx.designMx", Abs(batch.ResultAt(1).StabilityResult.DesignMx) > 0#
    AssertClose stats, "batch.stability.circleMx.noDesignMy", batch.ResultAt(1).StabilityResult.DesignMy, 0#, 1#

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

    AssertTrue stats, "batch.stability.centroid.status", batch.ResultAt(1).StabilityResult.Status = "OK" Or batch.ResultAt(1).StabilityResult.Status = "FAIL"
    AssertTrue stats, "batch.stability.centroid.eFromTransformedCenter", _
        Sqr(batch.ResultAt(1).StabilityResult.Eccentricity1 ^ 2 + batch.ResultAt(1).StabilityResult.Eccentricity2 ^ 2) > centroidGap * 0.8
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
    coreDistance = ActiveStabilityValue(batch.ResultAt(1).StabilityResult.CoreDistance1, batch.ResultAt(1).StabilityResult.CoreDistance2)
    radiusValue = ActiveStabilityValue(batch.ResultAt(1).StabilityResult.Radius1, batch.ResultAt(1).StabilityResult.Radius2)
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
        batch.ResultAt(1).StabilityResult.Depth1 > 190# And batch.ResultAt(1).StabilityResult.Depth1 < 310#
    AssertTrue stats, "batch.stability.depth.concreteOnly2", _
        batch.ResultAt(1).StabilityResult.Depth2 > 190# And batch.ResultAt(1).StabilityResult.Depth2 < 310#

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

    AssertClose stats, "batch.stability.sustainedN.notClamped", batch.ResultAt(1).StabilityResult.SustainedN, 240000#, 0.000001

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
    AssertTrue stats, "batch.writer.stability.sp35.empty", Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row, 29).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row, 30).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row, 45).Value2)) = 0
    AssertTrue stats, "batch.writer.stability.sp35.naStatus", _
        CStr(resultsSheet.Cells.Item(stabilityAnchor.Row, 38).Value2) = "N/A" And _
        CStr(resultsSheet.Cells.Item(stabilityAnchor.Row, 44).Value2) = "N/A" And _
        CStr(resultsSheet.Cells.Item(stabilityAnchor.Row, 53).Value2) = "N/A" And _
        CStr(resultsSheet.Cells.Item(stabilityAnchor.Row, 59).Value2) = "N/A" And _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 38) And _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 44)
    AssertTrue stats, "batch.writer.stability.sp63.filled", Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row, 61).Value2)) > 0 Or _
        Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row, 72).Value2)) > 0
    AssertTrue stats, "batch.writer.stability.sp63.statusColor", _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 72) And _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 84) And _
        CellHasNoFill(resultsSheet, stabilityAnchor.Row, 61)

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
    AssertTrue stats, "batch.dynamic.last.status", Len(batch.ResultAt(24).Status) > 0
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
    AssertTrue stats, "batch.invalid.reader.status", batch.ResultAt(1).Status = "InputErr"
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
    AssertTrue stats, "batch.writer.header.statusGroup", CStr(resultsSheet.Cells.Item(summaryRow + 5, 6).Value2) = "статус проверки"
    AssertTrue stats, "batch.writer.header.reserveGroup", CStr(resultsSheet.Cells.Item(summaryRow + 5, 17).Value2) = "минимальные коэффициенты запаса"
    AssertTrue stats, "batch.writer.header.statusStrength", CStr(resultsSheet.Cells.Item(summaryRow + 6, 6).Value2) = "прочность"
    AssertTrue stats, "batch.writer.header.reserveStrength", CStr(resultsSheet.Cells.Item(summaryRow + 6, 17).Value2) = "прочность"
    AssertTrue stats, "batch.writer.header.centered", resultsSheet.Cells.Item(summaryRow + 5, 6).HorizontalAlignment = -4108 And _
        resultsSheet.Cells.Item(summaryRow + 7, 17).HorizontalAlignment = -4108
    AssertTrue stats, "batch.writer.header.commentLeft", resultsSheet.Cells.Item(summaryRow + 10, 17).HorizontalAlignment = -4131
    AssertTrue stats, "batch.writer.header.epsilon", CStr(resultsSheet.Cells.Item(summaryRow + 7, 6).Value2) = _
        "по деформациям " & ChrW$(&H3B5)
    AssertTrue stats, "batch.writer.header.sp63StatusPlaneMerge", _
        CStr(resultsSheet.Cells.Item(summaryRow + 7, 14).MergeArea.Cells.Item(1, 1).Value2) = "Плоскость 1" And _
        resultsSheet.Cells.Item(summaryRow + 7, 14).MergeArea.Rows.Count = 2 And _
        resultsSheet.Cells.Item(summaryRow + 7, 14).MergeArea.Columns.Count = 1 And _
        CStr(resultsSheet.Cells.Item(summaryRow + 9, 14).Value2) = "Ncr/N" And _
        Not resultsSheet.Cells.Item(summaryRow + 9, 14).MergeCells
    AssertTrue stats, "batch.writer.header.sp63ReservePlaneMerge", _
        CStr(resultsSheet.Cells.Item(summaryRow + 7, 25).MergeArea.Cells.Item(1, 1).Value2) = "Плоскость 1" And _
        resultsSheet.Cells.Item(summaryRow + 7, 25).MergeArea.Rows.Count = 2 And _
        resultsSheet.Cells.Item(summaryRow + 7, 25).MergeArea.Columns.Count = 1 And _
        CStr(resultsSheet.Cells.Item(summaryRow + 9, 25).Value2) = "Ncr/N" And _
        Not resultsSheet.Cells.Item(summaryRow + 9, 25).MergeCells
    AssertTrue stats, "batch.writer.header.id", _
        CStr(resultsSheet.Cells.Item(summaryRow + 5, 1).MergeArea.Cells.Item(1, 1).Value2) = "Combination ID"
    AssertTrue stats, "batch.writer.header.idMergeRows", resultsSheet.Cells.Item(summaryRow + 5, 1).MergeArea.Rows.Count = 5
    AssertTrue stats, "batch.writer.header.commentRow", InStr(1, CStr(resultsSheet.Cells.Item(summaryRow + 10, 17).Value2), "деформациям", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.worst.label", CStr(resultsSheet.Cells.Item(summaryRow + 11, 1).Value2) = "worst LC"
    AssertTrue stats, "batch.writer.worst.commentDash", CStr(resultsSheet.Cells.Item(summaryRow + 11, 2).Value2) = "-"
    AssertTrue stats, "batch.writer.worst.resultCommentDash", CStr(resultsSheet.Cells.Item(summaryRow + 11, 3).Value2) = "-"
    AssertTrue stats, "batch.writer.worst.overallDash", CStr(resultsSheet.Cells.Item(summaryRow + 11, 4).Value2) = "-"
    AssertTrue stats, "batch.writer.worst.noNa", Not BatchSummaryWorstRowContainsText(resultsSheet, "N/A")
    AssertTrue stats, "batch.writer.worst.bold", resultsSheet.Cells.Item(summaryRow + 11, 1).Font.Bold And _
        resultsSheet.Cells.Item(summaryRow + 11, 17).Font.Bold
    AssertTrue stats, "batch.writer.statusLegend.title", BatchSummaryCellText(resultsSheet, summaryRow + 5, 28) = "Расшифровка статусов"
    AssertTrue stats, "batch.writer.statusLegend.header", _
        BatchSummaryCellText(resultsSheet, summaryRow + 6, 28) = "Статус" And _
        BatchSummaryCellText(resultsSheet, summaryRow + 6, 29) = "Описание"
    AssertTrue stats, "batch.writer.statusLegend.values", _
        BatchSummaryCellText(resultsSheet, summaryRow + 7, 28) = "OK" And _
        BatchSummaryCellText(resultsSheet, summaryRow + 9, 28) = "BaseFail" And _
        BatchSummaryCellText(resultsSheet, summaryRow + 11, 28) = "InputErr" And _
        BatchSummaryCellText(resultsSheet, summaryRow + 12, 28) = "CalcErr"
    AssertTrue stats, "batch.writer.statusLegend.mergeOnlyTitle", _
        resultsSheet.Cells.Item(summaryRow + 5, 28).MergeArea.Columns.Count = 2 And _
        Not resultsSheet.Cells.Item(summaryRow + 6, 28).MergeCells And _
        Not resultsSheet.Cells.Item(summaryRow + 7, 29).MergeCells
    AssertTrue stats, "batch.writer.statusLegend.noWrap", _
        Not resultsSheet.Cells.Item(summaryRow + 7, 29).WrapText
    AssertTrue stats, "batch.writer.statusLegend.italicValues", _
        resultsSheet.Cells.Item(summaryRow + 7, 28).Font.Italic And _
        resultsSheet.Cells.Item(summaryRow + 13, 29).Font.Italic
    AssertTrue stats, "batch.writer.statusLegend.colors", _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 7, 28) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 8, 28) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 9, 28) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 10, 28) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 11, 28) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 28) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 13, 28)
    AssertTrue stats, "batch.writer.reserve.dataNotHeaderFill", _
        resultsSheet.Cells.Item(summaryRow + 12, 17).Interior.ColorIndex = -4142
    AssertWorstSummaryRowMatchesData stats, resultsSheet
    AssertTrue stats, "batch.writer.data.firstId", CStr(resultsSheet.Cells.Item(summaryRow + 12, 1).Value2) = "W1"
    AssertTrue stats, "batch.writer.data.capacityStatus", Len(CStr(resultsSheet.Cells.Item(summaryRow + 12, 7).Value2)) > 0
    AssertTrue stats, "batch.writer.data.capacityReserve", IsNumeric(resultsSheet.Cells.Item(summaryRow + 12, 18).Value2)
    AssertTrue stats, "batch.writer.data.statusColors", _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 4) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 6) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 7) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 8) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 9) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 10) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 11) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 12) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 13) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 14) And _
        StatusCellHasExpectedFill(resultsSheet, summaryRow + 12, 15)
    Dim strengthAnchor As Object
    Set strengthAnchor = ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange
    AssertTrue stats, "batch.writer.strength.statusColors", _
        StatusCellHasExpectedFill(resultsSheet, strengthAnchor.Row, 3) And _
        StatusCellHasExpectedFill(resultsSheet, strengthAnchor.Row, 30) And _
        StatusCellHasExpectedFill(resultsSheet, strengthAnchor.Row, 49) And _
        CellHasNoFill(resultsSheet, strengthAnchor.Row, 1) And _
        CellHasNoFill(resultsSheet, strengthAnchor.Row, 10) And _
        CellHasNoFill(resultsSheet, strengthAnchor.Row, 31)
    AssertTrue stats, "batch.writer.strength.absentZonesBlank", _
        Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row + 2, 18).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row + 2, 19).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row + 2, 20).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row + 2, 22).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row + 2, 23).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row + 2, 24).Value2)) = 0
    AssertTrue stats, "batch.writer.strength.presentTensionKept", _
        Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row + 2, 21).Value2)) > 0 And _
        Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row + 2, 25).Value2)) > 0
    Dim crackAnchor As Object
    Set crackAnchor = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange
    AssertTrue stats, "batch.writer.crack.statusColors", _
        StatusCellHasExpectedFill(resultsSheet, crackAnchor.Row, 3) And _
        StatusCellHasExpectedFill(resultsSheet, crackAnchor.Row, 19) And _
        StatusCellHasExpectedFill(resultsSheet, crackAnchor.Row, 20) And _
        StatusCellHasExpectedFill(resultsSheet, crackAnchor.Row, 23) And _
        StatusCellHasExpectedFill(resultsSheet, crackAnchor.Row, 45) And _
        StatusCellHasExpectedFill(resultsSheet, crackAnchor.Row, 49) And _
        CellHasNoFill(resultsSheet, crackAnchor.Row, 1) And _
        CellHasNoFill(resultsSheet, crackAnchor.Row, 10) And _
        CellHasNoFill(resultsSheet, crackAnchor.Row, 46)
    AssertTrue stats, "batch.writer.crack.header.formationTitle", CStr(resultsSheet.Cells.Item(crackAnchor.Row - 4, 11).Value2) = "Момент образования трещин"
    AssertTrue stats, "batch.writer.crack.header.crackedStateTitle", _
        CStr(resultsSheet.Cells.Item(crackAnchor.Row - 4, 22).Value2) = "равновесие при заданных нагрузках"
    AssertTrue stats, "batch.writer.crack.header.title", CStr(resultsSheet.Cells.Item(crackAnchor.Row - 4, 24).Value2) = "нормальные и продольные трещины"
    AssertTrue stats, "batch.writer.crack.header.mcrcNote", InStr(1, CStr(resultsSheet.Cells.Item(crackAnchor.Row - 2, 17).Value2), "моментного вектора", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.crack.header.formationStatus", _
        CStr(resultsSheet.Cells.Item(crackAnchor.Row - 3, 19).Value2) = "статус трещин" And _
        resultsSheet.Cells.Item(crackAnchor.Row - 3, 19).MergeArea.Columns.Count = 3
    AssertTrue stats, "batch.writer.crack.header.state", CStr(resultsSheet.Cells.Item(crackAnchor.Row - 1, 21).Value2) = "state"
    AssertTrue stats, "batch.writer.crack.header.crackedStateStatus", CStr(resultsSheet.Cells.Item(crackAnchor.Row - 1, 23).Value2) = "статус"
    AssertTrue stats, "batch.writer.crack.header.es", _
        CStr(resultsSheet.Cells.Item(crackAnchor.Row - 1, 41).Value2) = "Es, MPa"
    AssertTrue stats, "batch.writer.crack.header.normalStatusRu", _
        CStr(resultsSheet.Cells.Item(crackAnchor.Row - 1, 45).Value2) = "статус"
    AssertTrue stats, "batch.writer.crack.header.longStatusRu", _
        CStr(resultsSheet.Cells.Item(crackAnchor.Row - 1, 49).Value2) = "статус"
    AssertTrue stats, "batch.writer.crack.header.notesPlain", Not resultsSheet.Cells.Item(crackAnchor.Row - 2, 15).Font.Bold And _
        resultsSheet.Cells.Item(crackAnchor.Row - 2, 15).HorizontalAlignment = -4131
    AssertTrue stats, "batch.writer.crack.header.notesFill", CLng(resultsSheet.Cells.Item(crackAnchor.Row - 2, 15).Interior.Color) = RGB(217, 217, 217)
    AssertTrue stats, "batch.writer.crack.availableRowsBorder", _
        Len(CStr(resultsSheet.Cells.Item(crackAnchor.Row + 19, 1).Value2)) = 0 And _
        resultsSheet.Cells.Item(crackAnchor.Row + 19, 1).Borders(9).LineStyle <> -4142
    Dim stabilityAnchor As Object
    Set stabilityAnchor = ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange
    AssertTrue stats, "batch.writer.stability.statusColor", _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 3) And _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 38) And _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 44) And _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 53) And _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 59) And _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 72) And _
        StatusCellHasExpectedFill(resultsSheet, stabilityAnchor.Row, 84) And _
        CellHasNoFill(resultsSheet, stabilityAnchor.Row, 1)
    AssertTrue stats, "batch.writer.stability.header.summary", CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 5, 1).Value2) = _
        "Итог по расчету (с учетом " & ChrW$(&H3B7) & ")"
    AssertTrue stats, "batch.writer.stability.header.statusWidth", resultsSheet.Cells.Item(stabilityAnchor.Row - 4, 3).MergeArea.Columns.Count = 1
    AssertTrue stats, "batch.writer.stability.header.mainAxes", InStr(1, CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 4, 4).Value2), _
        "главных центральных осей", vbTextCompare) > 0
    AssertTrue stats, "batch.writer.stability.header.noExtraTier", Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 3, 4).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 3, 6).Value2)) = 0
    AssertTrue stats, "batch.writer.stability.header.sp35", CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 5, 29).Value2) = "Расчет по СП 35"
    AssertTrue stats, "batch.writer.stability.header.sp63", CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 5, 61).Value2) = "Расчет по СП 63"
    AssertTrue stats, "batch.writer.stability.header.notes", Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 2, 4).Value2)) > 0 And _
        Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 2, 61).Value2)) > 0
    AssertTrue stats, "batch.writer.stability.header.notesPlain", Not resultsSheet.Cells.Item(stabilityAnchor.Row - 2, 4).Font.Bold And _
        resultsSheet.Cells.Item(stabilityAnchor.Row - 2, 4).HorizontalAlignment = -4131
    AssertTrue stats, "batch.writer.stability.header.notesFill", CLng(resultsSheet.Cells.Item(stabilityAnchor.Row - 2, 4).Interior.Color) = RGB(217, 217, 217)
    AssertTrue stats, "batch.writer.stability.header.sp35NcrBranch", CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 3, 33).Value2) = "при ec > r"
    AssertTrue stats, "batch.writer.stability.header.sp63PlaneHeight", resultsSheet.Cells.Item(stabilityAnchor.Row - 4, 61).MergeArea.Rows.Count = 2
    AssertTrue stats, "batch.writer.stability.header.sp35Ratio", CStr(resultsSheet.Cells.Item(stabilityAnchor.Row - 1, 37).Value2) = "0.7*Ncr/N"
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

' Проверяет, что Results сохраняет пустые строки из rngLoadCombinations:
' LC с source-offset 1 и 4 должны попасть в первую и четвертую строки данных,
' а промежуточные строки остаются пустыми без N/A и статусной заливки.
Private Sub TestBatchSummaryPreservesSourceRowGaps(ByRef stats As TBatchTestStats)
    On Error GoTo Failed

    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "GAP1", -150000#, -2000000#, -1000000#, "PR1", "first source row", vbNullString, 1
    batch.AddCombination "GAP4", -90000#, 0#, 0#, "PR2", "fourth source row", vbNullString, 4
    batch.Execute

    Dim writer As CBatchResultWriter
    Set writer = New CBatchResultWriter
    writer.WriteSummary ThisWorkbook, batch

    Dim resultsSheet As Object
    Set resultsSheet = ThisWorkbook.Worksheets.Item("Results")
    Dim summaryRow As Long
    summaryRow = BatchSummaryStartRow()

    AssertTrue stats, "batch.writer.gaps.summary.first", _
        CStr(resultsSheet.Cells.Item(summaryRow + 12, 1).Value2) = "GAP1"
    AssertTrue stats, "batch.writer.gaps.summary.blank2", _
        Len(CStr(resultsSheet.Cells.Item(summaryRow + 13, 1).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(summaryRow + 13, 4).Value2)) = 0 And _
        CellHasNoFill(resultsSheet, summaryRow + 13, 4)
    AssertTrue stats, "batch.writer.gaps.summary.blank3", _
        Len(CStr(resultsSheet.Cells.Item(summaryRow + 14, 1).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(summaryRow + 14, 4).Value2)) = 0 And _
        CellHasNoFill(resultsSheet, summaryRow + 14, 4)
    AssertTrue stats, "batch.writer.gaps.summary.fourth", _
        CStr(resultsSheet.Cells.Item(summaryRow + 15, 1).Value2) = "GAP4"

    Dim strengthAnchor As Object
    Dim crackAnchor As Object
    Dim stabilityAnchor As Object
    Set strengthAnchor = ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange
    Set crackAnchor = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange
    Set stabilityAnchor = ThisWorkbook.Names.Item("rngStabilitySummaryAnchor").RefersToRange

    AssertTrue stats, "batch.writer.gaps.strength", _
        CStr(resultsSheet.Cells.Item(strengthAnchor.Row, 1).Value2) = "GAP1" And _
        Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row + 1, 1).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(strengthAnchor.Row + 2, 1).Value2)) = 0 And _
        CStr(resultsSheet.Cells.Item(strengthAnchor.Row + 3, 1).Value2) = "GAP4"
    AssertTrue stats, "batch.writer.gaps.crack", _
        CStr(resultsSheet.Cells.Item(crackAnchor.Row, 1).Value2) = "GAP1" And _
        Len(CStr(resultsSheet.Cells.Item(crackAnchor.Row + 1, 1).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(crackAnchor.Row + 2, 1).Value2)) = 0 And _
        CStr(resultsSheet.Cells.Item(crackAnchor.Row + 3, 1).Value2) = "GAP4"
    AssertTrue stats, "batch.writer.gaps.stability", _
        CStr(resultsSheet.Cells.Item(stabilityAnchor.Row, 1).Value2) = "GAP1" And _
        Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row + 1, 1).Value2)) = 0 And _
        Len(CStr(resultsSheet.Cells.Item(stabilityAnchor.Row + 2, 1).Value2)) = 0 And _
        CStr(resultsSheet.Cells.Item(stabilityAnchor.Row + 3, 1).Value2) = "GAP4"
    Exit Sub

Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: batch.writer.gaps; " & Err.Description
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
            resultsSheet.Cells.Item(summaryRow, 17).Value2, _
            resultsSheet.Cells.Item(strengthRow, 29).Value2
        AssertOptionalReserve stats, "batch.writer.reserve." & combinationID & ".capacity", _
            resultsSheet.Cells.Item(summaryRow, 18).Value2, _
            resultsSheet.Cells.Item(strengthRow, 48).Value2
    End If
    If checkCrack And crackRow > 0 Then
        AssertOptionalReserve stats, "batch.writer.reserve." & combinationID & ".crack", _
            resultsSheet.Cells.Item(summaryRow, 19).Value2, _
            resultsSheet.Cells.Item(crackRow, 44).Value2
        AssertOptionalReserve stats, "batch.writer.reserve." & combinationID & ".longCrack", _
            resultsSheet.Cells.Item(summaryRow, 20).Value2, _
            resultsSheet.Cells.Item(crackRow, 48).Value2
    End If
    If checkStability And stabilityRow > 0 Then
        AssertOptionalReserve stats, "batch.writer.reserve." & combinationID & ".sp35p1eta", _
            resultsSheet.Cells.Item(summaryRow, 21).Value2, _
            resultsSheet.Cells.Item(stabilityRow, 37).Value2
        AssertOptionalReserve stats, "batch.writer.reserve." & combinationID & ".sp35p1table", _
            resultsSheet.Cells.Item(summaryRow, 22).Value2, _
            resultsSheet.Cells.Item(stabilityRow, 43).Value2
        AssertOptionalReserve stats, "batch.writer.reserve." & combinationID & ".sp35p2eta", _
            resultsSheet.Cells.Item(summaryRow, 23).Value2, _
            resultsSheet.Cells.Item(stabilityRow, 52).Value2
        AssertOptionalReserve stats, "batch.writer.reserve." & combinationID & ".sp35p2table", _
            resultsSheet.Cells.Item(summaryRow, 24).Value2, _
            resultsSheet.Cells.Item(stabilityRow, 58).Value2
        AssertOptionalReserve stats, "batch.writer.reserve." & combinationID & ".sp63p1", _
            resultsSheet.Cells.Item(summaryRow, 25).Value2, _
            resultsSheet.Cells.Item(stabilityRow, 71).Value2
        AssertOptionalReserve stats, "batch.writer.reserve." & combinationID & ".sp63p2", _
            resultsSheet.Cells.Item(summaryRow, 26).Value2, _
            resultsSheet.Cells.Item(stabilityRow, 83).Value2
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
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "strain", 6, 17
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "capacity", 7, 18
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "crack", 8, 19
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "longCrack", 9, 20
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "sp35p1eta", 10, 21
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "sp35p1table", 11, 22
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "sp35p2eta", 12, 23
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "sp35p2table", 13, 24
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "sp63p1", 14, 25
    AssertWorstSummaryColumnMatchesData stats, resultsSheet, "sp63p2", 15, 26
End Sub

' Проверяет строку worst LC по расчетным столбцам compact summary.
Private Function BatchSummaryWorstRowContainsText(ByVal resultsSheet As Object, ByVal textValue As String) As Boolean
    Dim anchorRow As Long
    anchorRow = BatchSummaryStartRow()

    Dim columnIndex As Long
    For columnIndex = 6 To 26
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

' Возвращает статус полного batch-расчета для указанного режима расстояния.
' Численное отличие NearestContour/GlobalExtreme проверяется в прямых crack-тестах;
' здесь важно, что настройка проходит через batch и не превращается в InputErr.
Private Function BatchSteppedCrackCoverStatus(ByVal coverMode As String) As String
    SetSystemSetting "SLS.Crack.CoverDistanceMode", coverMode

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim batch As CBatchSectionCalculator
    Set batch = BuildSteppedCrackCoverBatch()
    batch.ApplySettings settings
    batch.AddCombination "COVER_" & coverMode, -20000#, -32000000#, 0#, "PR2", "cover distance mode"
    batch.Execute

    BatchSteppedCrackCoverStatus = batch.ResultAt(1).NormalCrackStatus
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

    If batch.ResultAt(1).StabilityResult.Status <> "OK" And batch.ResultAt(1).StabilityResult.Status <> "FAIL" Then
        Err.Raise vbObjectError + 3934, "modTestBatchCalculation", _
            "SP35 stability status is not applicable: " & batch.ResultAt(1).StabilityResult.Status
    End If
    StabilitySP35CriticalForceForValueSet = batch.ResultAt(1).StabilityResult.CriticalForce
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

    StabilityAccidentalForSingleMoment = Abs(batch.ResultAt(1).StabilityResult.AccidentalEcc1) + _
        Abs(batch.ResultAt(1).StabilityResult.AccidentalEcc2)
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

    StabilityStatusForCurrentSettings = batch.ResultAt(1).StabilityResult.Status
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

    StabilityPhiLForDurationLoad = ActiveMomentPlaneValue(batch, batch.ResultAt(1).StabilityResult.PhiL1, batch.ResultAt(1).StabilityResult.PhiL2)
End Function

' Возвращает значение из той плоскости, где фактически есть больший главный
' изгибающий момент. После явного соглашения "ось 1 = I1" у широкого
' прямоугольника глобальный Mx попадает во вторую главную плоскость, поэтому
' выбирать просто первое ненулевое значение нельзя.
Private Function ActiveMomentPlaneValue(ByVal batch As CBatchSectionCalculator, _
        ByVal firstValue As Double, ByVal secondValue As Double) As Double
    If Abs(batch.ResultAt(1).StabilityResult.Moment1) >= Abs(batch.ResultAt(1).StabilityResult.Moment2) Then
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

' Проверяет утвержденный словарь ResultMeta -> внешний статус.
' Отдельно фиксируется правило этапа 1: capacity с INITIAL_STATE_BEYOND_LIMIT
' выводится как BaseFail, а не как NumFail.
Private Sub TestResultMetaStatusDictionary(ByRef stats As TBatchTestStats)
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy

    Dim meta As CResultMeta
    Set meta = New CResultMeta

    meta.SetResult rsSuccess, rcCheckPassed, rkGeneric, vbNullString
    AssertEquals stats, "resultMeta.ok", policy.ExternalStatus(meta), "OK"

    meta.SetResult rsCheckFailed, rcCheckFailed, rkGeneric, vbNullString
    AssertEquals stats, "resultMeta.fail", policy.ExternalStatus(meta), "FAIL"

    meta.SetResult rsNumericalFailure, rcNumericalFailure, rkDirectState, vbNullString
    AssertEquals stats, "resultMeta.numFail", policy.ExternalStatus(meta), "NumFail"

    meta.SetResult rsInvalidInput, rcInvalidInput, rkGeneric, vbNullString
    AssertEquals stats, "resultMeta.inputErr", policy.ExternalStatus(meta), "InputErr"

    meta.SetResult rsInternalError, rcInternalError, rkGeneric, vbNullString
    AssertEquals stats, "resultMeta.calcErr", policy.ExternalStatus(meta), "CalcErr"

    meta.SetResult rsCheckFailed, rcInitialStateBeyondLimit, rkCapacity, vbNullString
    AssertEquals stats, "resultMeta.baseFail", policy.ExternalStatus(meta), "BaseFail"
End Sub

' Проверяет, что неприменимые ветви не ухудшают итоговый статус LC.
' Это защищает batch summary от ситуации, когда незапрошенный расчет перебивает
' реально выполненные OK/FAIL-ветки.
Private Sub TestResultMetaAggregateSkipsNotApplicable(ByRef stats As TBatchTestStats)
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy

    Dim directMeta As CResultMeta
    Set directMeta = New CResultMeta
    directMeta.SetResult rsSuccess, rcCheckPassed, rkDirectState, vbNullString

    Dim skippedMeta As CResultMeta
    Set skippedMeta = New CResultMeta
    skippedMeta.SetNotApplicable rkCapacity, vbNullString

    Dim failedMeta As CResultMeta
    Set failedMeta = New CResultMeta
    failedMeta.SetResult rsCheckFailed, rcCheckFailed, rkCrackWidth, vbNullString

    AssertEquals stats, "resultMeta.aggregate.okWithNA", _
        policy.AggregateMeta(directMeta, skippedMeta, Nothing), "OK"
    AssertEquals stats, "resultMeta.aggregate.failWithNA", _
        policy.AggregateMeta(directMeta, skippedMeta, failedMeta), "FAIL"
    AssertEquals stats, "resultMeta.aggregate.onlyNA", _
        policy.AggregateMeta(skippedMeta, Nothing, Nothing), "N/A"
End Sub

' Проверяет финальный result-tree: display вычисляется из канонической meta,
' а сводные комментарии собираются только из соответствующих typed ветвей.
Private Sub TestCombinationResultTreeDrivesDisplayFields(ByRef stats As TBatchTestStats)
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy

    Dim result As CCombinationResult
    Set result = New CCombinationResult
    result.Clear "LambdaMxy"

    Dim okMeta As CResultMeta
    Set okMeta = New CResultMeta
    okMeta.SetResult rsSuccess, rcCheckPassed, rkDirectState, vbNullString
    result.SetDirectStateMeta okMeta
    AssertEquals stats, "combinationTree.direct.flat", result.StrengthResult.DirectState.Status, "OK"
    AssertEquals stats, "combinationTree.direct.meta", _
        policy.ExternalStatus(result.DirectStateMeta), "OK"

    Dim warningResult As CCombinationResult
    Set warningResult = New CCombinationResult
    warningResult.Clear "LambdaMxy"
    Dim warningMeta As CResultMeta
    Set warningMeta = New CResultMeta
    warningMeta.SetResult rsSuccessWithWarning, rcCheckPassed, rkDirectState, _
        "Равновесие найдено через резервный старт."
    warningResult.SetDirectStateMeta warningMeta
    Dim warningOverall As CResultMeta
    Set warningOverall = warningResult.OverallMeta
    AssertEquals stats, "combinationTree.warning.externalOk", _
        policy.ExternalStatus(warningOverall), "OK"
    AssertTrue stats, "combinationTree.warning.commentKept", _
        InStr(1, warningOverall.ResultComment, "Равновесие найдено через резервный старт", vbTextCompare) > 0

    Dim failMeta As CResultMeta
    Set failMeta = New CResultMeta
    failMeta.SetResult rsCheckFailed, rcCheckFailed, rkDirectState, "Проверка не проходит."
    result.MergeDirectStateMeta failMeta
    AssertEquals stats, "combinationTree.direct.merge.flat", result.StrengthResult.DirectState.Status, "FAIL"
    AssertEquals stats, "combinationTree.direct.merge.meta", _
        policy.ExternalStatus(result.StrengthResult.DirectMeta), "FAIL"

    Dim capacity As CCapacityResult
    Set capacity = New CCapacityResult
    capacity.InitializeSkipped "capacity не запрошена", "None"
    result.StoreCapacityResult capacity
    AssertTrue stats, "combinationTree.capacity.exists", Not result.StrengthResult.Capacity Is Nothing
    AssertEquals stats, "combinationTree.capacity.meta", _
        policy.ExternalStatus(result.CapacityMeta), "N/A"

    Dim stabilityMeta As CResultMeta
    Set stabilityMeta = New CResultMeta
    stabilityMeta.SetResult rsInvalidInput, rcInvalidInput, rkStability, "ошибка настройки устойчивости"
    result.SetStabilityMeta stabilityMeta, "ошибка настройки устойчивости"
    AssertEquals stats, "combinationTree.stability.flat", result.StabilityResult.Status, "InputErr"
    AssertEquals stats, "combinationTree.stability.meta", _
        policy.ExternalStatus(result.StabilityMeta), "InputErr"

    Dim overall As CResultMeta
    Set overall = result.OverallMeta
    AssertTrue stats, "combinationTree.overallComment.direct", _
        InStr(1, overall.ResultComment, "НДС:", vbTextCompare) > 0
    AssertTrue stats, "combinationTree.overallComment.stability", _
        InStr(1, overall.ResultComment, "Устойчивость: ошибка настройки устойчивости", vbTextCompare) > 0

    Dim baseResult As CCombinationResult
    Set baseResult = New CCombinationResult
    baseResult.Clear "LambdaMxy"
    baseResult.SetDirectStateMeta failMeta
    Dim baseCapacity As CCapacityResult
    Set baseCapacity = New CCapacityResult
    baseCapacity.InitializeBaseFail "Начальное состояние при lambda = 0 не проходит физический критерий: ConcreteStrainLimit."
    baseResult.StoreCapacityResult baseCapacity
    AssertEquals stats, "combinationTree.strength.baseFail.meta", _
        policy.ExternalStatus(baseResult.StrengthMeta), "BaseFail"
    AssertTrue stats, "combinationTree.strength.baseFail.comment", _
        InStr(1, baseResult.StrengthMeta.ResultComment, "НДС: Проверка не проходит.", vbTextCompare) > 0 And _
        InStr(1, baseResult.StrengthMeta.ResultComment, _
            "Несущая: Несущая способность не проходит уже для исходной части выбранного пути", vbTextCompare) > 0
End Sub

' Проверяет, что общий статус detailed-блока трещин не скрывает ошибку
' обязательного CurrentCrackedState. Формульные проверки при такой ошибке
' остаются N/A/blocked, но сводный crack-meta обязан сохранить NumFail/InputErr/CalcErr.
Private Sub TestCrackAggregateIncludesCurrentStateFailure(ByRef stats As TBatchTestStats)
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy

    Dim currentStateMeta As CResultMeta
    Set currentStateMeta = New CResultMeta
    currentStateMeta.SetResult rsNumericalFailure, rcNumericalFailure, rkDirectState, _
        "CrackedState от заданного сочетания не найден численно."

    Dim crackTree As CCrackResult
    Set crackTree = New CCrackResult
    crackTree.Initialize Nothing, currentStateMeta, Nothing, Nothing

    Dim result As CCombinationResult
    Set result = New CCombinationResult
    result.Clear "LambdaMxy"
    result.StoreCrackAggregateResult crackTree

    AssertEquals stats, "combinationTree.crack.currentState.meta", _
        policy.ExternalStatus(result.CrackCurrentStateMeta), "NumFail"
    AssertEquals stats, "combinationTree.crack.aggregate.currentStateFailure", _
        policy.ExternalStatus(result.CrackMeta), "NumFail"
    AssertEquals stats, "combinationTree.crack.flat.currentStateFailure", _
        result.NormalCrackStatus, "NumFail"
    AssertTrue stats, "combinationTree.crack.comment.currentStateFailure", _
        InStr(1, result.CrackMeta.ResultComment, "CrackedState от заданного сочетания", vbTextCompare) > 0

    Dim longitudinalMeta As CResultMeta
    Set longitudinalMeta = New CResultMeta
    longitudinalMeta.SetResult rsCheckFailed, rcCheckFailed, rkLongitudinalCrack, _
        "Продольные трещины: напряжение сжатого бетона выше допустимого."

    Dim longitudinalResult As CLongitudinalCrackResult
    Set longitudinalResult = New CLongitudinalCrackResult
    longitudinalResult.Initialize longitudinalMeta, 20#, 14.6, 20# / 14.6

    Dim summaryTree As CCrackResult
    Set summaryTree = New CCrackResult
    summaryTree.Initialize Nothing, Nothing, Nothing, longitudinalResult

    Dim summaryResult As CCombinationResult
    Set summaryResult = New CCombinationResult
    summaryResult.Clear "LambdaMxy"
    summaryResult.StoreCrackAggregateResult summaryTree

    AssertEquals stats, "combinationTree.crack.normalOnly.na", _
        policy.ExternalStatus(summaryResult.CrackMeta), "N/A"
    AssertEquals stats, "combinationTree.crack.summary.longitudinalFail", _
        policy.ExternalStatus(summaryResult.CrackSummaryMeta), "FAIL"
    AssertTrue stats, "combinationTree.crack.summary.comment", _
        InStr(1, summaryResult.CrackSummaryMeta.ResultComment, "Продольные трещины", vbTextCompare) > 0
End Sub

' Проверяет, что проверки без поиска равновесия не создают NumFail.
' Продольные трещины и устойчивость являются инженерскими проверками готовых
' величин; их отрицательный результат должен быть FAIL/InputErr/N/A, но не NumFail.
Private Sub TestFormulaChecksDoNotCreateNumFail(ByRef stats As TBatchTestStats)
    Dim policy As CBatchStatusPolicy
    Set policy = New CBatchStatusPolicy

    AssertEquals stats, "status.longitudinal.pass", _
        policy.StatusFromMeta(policy.LongitudinalCrackMeta(10#, 14.6)), "OK"
    AssertEquals stats, "status.longitudinal.fail", _
        policy.StatusFromMeta(policy.LongitudinalCrackMeta(20#, 14.6)), "FAIL"
    AssertEquals stats, "status.longitudinal.na", _
        policy.StatusFromMeta(policy.LongitudinalCrackMeta(0#, 14.6)), "N/A"
    AssertEquals stats, "status.longitudinal.input", _
        policy.StatusFromMeta(policy.LongitudinalCrackMeta(10#, 0#)), "InputErr"

    AssertTrue stats, "status.longitudinal.noNumFail", _
        policy.StatusFromMeta(policy.LongitudinalCrackMeta(20#, 14.6)) <> "NumFail"
End Sub

' Проверяет Stage 2: named-state хранит не только плоскость деформаций, но и
' целевые/внутренние усилия, невязки и diagnostic metadata solver-а.
Private Sub TestSectionStateResultStoresEquilibriumData(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "STATE_META", -100000#, 12000000#, 3000000#, "PR1", "state metadata"
    batch.Execute

    Dim stateResult As CSectionStateResult
    Set stateResult = batch.ResultAt(1).StateRepository.FindState(sstStrengthState)

    AssertTrue stats, "stateMeta.exists", Not stateResult Is Nothing
    If stateResult Is Nothing Then Exit Sub

    AssertClose stats, "stateMeta.epsilon0", stateResult.Epsilon0, batch.ResultAt(1).StrengthResult.DirectState.StateResult.Epsilon0, 0.000000000001
    AssertClose stats, "stateMeta.kappaX", stateResult.KappaX, batch.ResultAt(1).StrengthResult.DirectState.StateResult.KappaX, 0.000000000001
    AssertClose stats, "stateMeta.kappaY", stateResult.KappaY, batch.ResultAt(1).StrengthResult.DirectState.StateResult.KappaY, 0.000000000001
    AssertClose stats, "stateMeta.targetN", stateResult.TargetN, batch.N(1), 0.001
    AssertClose stats, "stateMeta.targetMx", stateResult.TargetMx, batch.Mx(1), 0.001
    AssertClose stats, "stateMeta.targetMy", stateResult.TargetMy, batch.My(1), 0.001
    AssertClose stats, "stateMeta.nint", stateResult.Nint, batch.ResultAt(1).StrengthResult.DirectState.StateResult.Nint, 0.001
    AssertClose stats, "stateMeta.mxint", stateResult.Mxint, batch.ResultAt(1).StrengthResult.DirectState.StateResult.Mxint, 0.001
    AssertClose stats, "stateMeta.myint", stateResult.Myint, batch.ResultAt(1).StrengthResult.DirectState.StateResult.Myint, 0.001
    AssertTrue stats, "stateMeta.solverCalls", stateResult.SolverCallCount >= 1
    AssertTrue stats, "stateMeta.iterations", stateResult.IterationCount >= 0
    AssertTrue stats, "stateMeta.resultMeta", Not stateResult.ResultMeta Is Nothing
End Sub

' Проверяет Stage 3: технические solve-options не входят в ключ
' эквивалентности named-state, если они не меняют физический результат.
Private Sub TestStateRequestEquivalenceIgnoresSolveOptions(ByRef stats As TBatchTestStats)
    Dim spec As CMaterialModelSpec
    Set spec = New CMaterialModelSpec
    spec.Initialize "ULS(I)", "ThreeLine", "Ignore", "TwoLine"

    Dim requestA As CStateRequest
    Set requestA = New CStateRequest
    requestA.Initialize sstStrengthState, cpStrength, spec, 1000#, 2000#, 3000#, True, False

    Dim requestB As CStateRequest
    Set requestB = New CStateRequest
    requestB.Initialize sstStrengthState, cpStrength, spec, 1000#, 2000#, 3000#, True, True

    Dim requestDifferent As CStateRequest
    Set requestDifferent = New CStateRequest
    requestDifferent.Initialize sstStrengthState, cpStrength, spec, 1000#, 2500#, 3000#, True, False

    AssertEquals stats, "stateRequest.key.solveOptionsIgnored", _
        requestA.EquivalenceKey, requestB.EquivalenceKey
    AssertTrue stats, "stateRequest.key.loadsMatter", _
        requestA.EquivalenceKey <> requestDifferent.EquivalenceKey
End Sub

' Проверяет Stage 3: repository может переиспользовать только реально найденное
' физическое НДС. Неуспешная попытка хранится для вывода, но не блокирует
' будущий solve того же физического запроса.
Private Sub TestStateRepositoryReusesOnlyConvergedStates(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "STATE_REUSE", -100000#, 12000000#, 3000000#, "PR1", "state repository"
    batch.Execute

    Dim successfulState As CSectionStateResult
    Set successfulState = batch.ResultAt(1).StateRepository.FindState(sstStrengthState)
    AssertTrue stats, "stateRepository.successSource", Not successfulState Is Nothing
    If successfulState Is Nothing Then Exit Sub

    Dim successfulRequest As CStateRequest
    Set successfulRequest = New CStateRequest
    successfulRequest.Initialize sstStrengthState, cpStrength, successfulState.MaterialSpec, _
        successfulState.TargetN, successfulState.TargetMx, successfulState.TargetMy, _
        successfulState.ExtensionUsed, False

    Dim successfulRepository As CStateRepository
    Set successfulRepository = New CStateRepository
    successfulRepository.StoreForRequest successfulRequest, successfulState

    AssertTrue stats, "stateRepository.successReusable", _
        successfulRepository.FindEquivalent(successfulRequest) Is successfulState

    Dim retryRepository As CStateRepository
    Set retryRepository = New CStateRepository

    Dim failedSolver As CSectionSolver
    Set failedSolver = New CSectionSolver

    Dim failedState As CSectionStateResult
    Set failedState = New CSectionStateResult
    failedState.InitializeFromSolver sstStrengthState, cpStrength, successfulState.MaterialSpec, failedSolver, False
    retryRepository.StoreForRequest successfulRequest, failedState

    AssertTrue stats, "stateRepository.failedNotReusable", retryRepository.FindEquivalent(successfulRequest) Is Nothing
    AssertTrue stats, "stateRepository.failedStillInSnapshot", retryRepository.StateCount = 1

    retryRepository.StoreForRequest successfulRequest, successfulState
    AssertTrue stats, "stateRepository.retrySuccessReusable", _
        retryRepository.FindEquivalent(successfulRequest) Is successfulState
    AssertTrue stats, "stateRepository.retrySuccessSnapshot", _
        retryRepository.FindState(sstStrengthState) Is successfulState
End Sub

' Проверяет финальный контракт именования: после миграции принимаются только
' канонические PreCrackState/PostCrackState без старых Before/AfterMcrc alias.
Private Sub TestPrePostCrackStateNames(ByRef stats As TBatchTestStats)
    AssertTrue stats, "stateName.preCanonical", _
        SectionStateTypeFromText("PreCrackState") = sstPreCrackState
    AssertTrue stats, "stateName.postCanonical", _
        SectionStateTypeFromText("PostCrackState") = sstPostCrackState
    AssertEquals stats, "stateName.preText", SectionStateTypeToText(sstPreCrackState), "PreCrackState"
    AssertEquals stats, "stateName.postText", SectionStateTypeToText(sstPostCrackState), "PostCrackState"
End Sub

' ============================== ДЛЯ ТЕСТОВ AUDIT02 ==============================

' Изолирует повторное чтение сводной meta от solver-а и Excel writer-ов.
' Возвращает время и последний комментарий для диагностики расхода строк;
' физические числа и штатный regression-набор здесь не изменяются.
Public Function RunAudit02ResultMetaStress() As String
    Dim result As CCombinationResult
    Set result = New CCombinationResult
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    meta.SetResult rsInvalidInput, rcInvalidInput, rkDirectState, "Тестовая ошибка исходных данных."
    result.SetDirectStateMeta meta
    meta.SetResult rsInvalidInput, rcInvalidInput, rkStability, "Тестовая ошибка настройки устойчивости."
    result.SetStabilityMeta meta
    Dim started As Double
    started = Timer
    Dim i As Long
    Dim statusText As String
    Dim commentText As String
    For i = 1 To 5000
        statusText = result.Status
        commentText = result.OverallMeta.ResultComment
    Next i
    RunAudit02ResultMetaStress = "INFO: metaStress; reads=5000; elapsedSec=" & _
        FormatNumberInvariant(Timer - started) & "; status=" & statusText & "; comment=" & commentText
End Function

' Повторяет реальный batch-маршрут после сохранения Pre/Post/current states.
' Центральное растяжение не требует поисковых проб, поэтому при cache-hit
' можно точно доказать отсутствие любого нового equilibrium solve.
Private Sub TestAudit02CurrentCrackedStateCacheHitCalculatesWidth(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "AUDIT02_CURRENT_CACHE", 200000#, 0#, 0#, "PR2", "current state cache"
    batch.Execute
    Dim first As CCrackResult
    Set first = batch.ResultAt(1).CrackResult
    AssertTrue stats, "audit02.currentCache.firstExists", Not first Is Nothing
    If first Is Nothing Then Exit Sub
    AssertTrue stats, "audit02.currentCache.crackFormed", first.Formation.CrackFormed
    AssertTrue stats, "audit02.currentCache.firstWidth", first.Width.CrackWidth > 0#
    Dim countBefore As Long
    countBefore = SectionEquilibriumSolveCount()
    batch.TestRepeatCrackCalculation 1
    Dim second As CCrackResult
    Set second = batch.ResultAt(1).CrackResult
    AssertTrue stats, "audit02.currentCache.secondExists", Not second Is Nothing
    If second Is Nothing Then Exit Sub
    AssertTrue stats, "audit02.currentCache.noHeavySolve", SectionEquilibriumSolveCount() = countBefore
    AssertClose stats, "audit02.currentCache.width", second.Width.CrackWidth, first.Width.CrackWidth, 0.000000001
    AssertClose stats, "audit02.currentCache.sigma", second.Width.SigmaS, first.Width.SigmaS, 0.000000001
    AssertClose stats, "audit02.currentCache.sigmaCrc", second.Width.SigmaSCrc, first.Width.SigmaSCrc, 0.000000001
    AssertClose stats, "audit02.currentCache.psi", second.Width.PsiS, first.Width.PsiS, 0.000000001
    AssertTrue stats, "audit02.currentCache.stateOK", second.CurrentStateMeta.InternalStatus = rsSuccess
End Sub

' Проверяет канонический результат реального расчета и повторное использование
' контейнеров: State не копируется в direct-result, meta не меняется снаружи,
' а нейтральная повторная инициализация не оставляет чисел прошлого сочетания.
Private Sub TestAudit02CanonicalResultsAndReset(ByRef stats As TBatchTestStats)
    Dim batch As CBatchSectionCalculator
    Set batch = BuildBatchCalculator()
    batch.AddCombination "AUDIT02_CANONICAL", 200000#, 0#, 0#, "PR2", "canonical result"
    batch.Execute
    Dim calculated As CCombinationResult
    Set calculated = batch.ResultAt(1)
    Dim current As CSectionStateResult
    Set current = calculated.StateRepository.FindState(sstCrackedState)
    AssertTrue stats, "audit02.canonical.currentExists", Not current Is Nothing
    If current Is Nothing Then Exit Sub

    Dim result As CCombinationResult
    Set result = New CCombinationResult
    result.StoreSectionStateResult current
    result.AddState current
    AssertTrue stats, "audit02.canonical.directIdentity", result.StrengthResult.DirectState.StateResult Is current
    AssertTrue stats, "audit02.canonical.repositoryIdentity", result.StateRepository.FindState(sstCrackedState) Is current
    AssertClose stats, "audit02.canonical.directN", result.StrengthResult.DirectState.StateResult.Nint, current.Nint, 0#
    Dim detached As CResultMeta
    Set detached = result.DirectStateMeta
    detached.SetResult rsInternalError, rcInternalError, rkDirectState, "Изменение внешней копии."
    AssertEquals stats, "audit02.canonical.directMetaIsolated", result.StrengthResult.DirectState.Status, "OK"

    result.StoreCrackAggregateResult calculated.CrackResult
    AssertTrue stats, "audit02.canonical.crackIdentity", result.CrackResult Is calculated.CrackResult
    AssertTrue stats, "audit02.canonical.widthIdentity", result.CrackResult.Width Is calculated.CrackResult.Width
    AssertTrue stats, "audit02.canonical.noDirectOverwrite", result.StrengthResult.DirectState.StateResult Is current

    Dim width As CCrackWidthResult
    Set width = result.CrackResult.Width
    AssertTrue stats, "audit02.canonical.widthPopulated", width.CrackWidth > 0#
    Dim neutral As CResultMeta
    Set neutral = New CResultMeta
    neutral.SetNotApplicable rkCrackWidth
    width.InitializeFromCalculator Nothing, neutral
    AssertClose stats, "audit02.canonical.widthReset", width.CrackWidth, 0#, 0#
    AssertClose stats, "audit02.canonical.sigmaReset", width.SigmaS, 0#, 0#
    AssertClose stats, "audit02.canonical.spacingReset", width.CrackSpacing, 0#, 0#
    AssertEquals stats, "audit02.canonical.rebarsReset", width.TensionRebarIds, vbNullString
    AssertTrue stats, "audit02.canonical.widthMetaReset", width.ResultMeta.InternalStatus = rsNotApplicable

    Dim stability As CStabilityResult
    Set stability = New CStabilityResult
    stability.Ncr2 = 123#
    stability.PhiValue1 = 0.8
    stability.StiffnessD2 = 456#
    stability.PlaneBranch2 = "branch"
    stability.PlaneApplicable2 = True
    stability.PlanePassed2 = True
    stability.Clear
    AssertClose stats, "audit02.canonical.stabilityNcrReset", stability.Ncr2, 0#, 0#
    AssertClose stats, "audit02.canonical.stabilityPhiReset", stability.PhiValue1, 0#, 0#
    AssertClose stats, "audit02.canonical.stabilityDReset", stability.StiffnessD2, 0#, 0#
    AssertEquals stats, "audit02.canonical.stabilityBranchReset", stability.PlaneBranch2, vbNullString
    AssertTrue stats, "audit02.canonical.stabilityFlagsReset", Not stability.PlaneApplicable2 And Not stability.PlanePassed2
    AssertEquals stats, "audit02.canonical.stabilityStatusReset", stability.Status, "N/A"
    Set detached = stability.Meta
    detached.SetResult rsInternalError, rcInternalError, rkStability, "Изменение внешней копии."
    AssertEquals stats, "audit02.canonical.stabilityMetaIsolated", stability.Status, "N/A"

    result.Clear "LambdaN"
    AssertTrue stats, "audit02.canonical.clearNoState", result.StrengthResult.DirectState.StateResult Is Nothing
    AssertTrue stats, "audit02.canonical.clearRepository", result.StateRepository.StateCount = 0
    AssertClose stats, "audit02.canonical.clearWidth", result.CrackResult.Width.CrackWidth, 0#, 0#
    AssertEquals stats, "audit02.canonical.clearStatus", result.Status, "N/A"
End Sub

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

    phiValue = ActiveStabilityValue(batch.ResultAt(1).StabilityResult.PhiValue1, batch.ResultAt(1).StabilityResult.PhiValue2)
    SP35TableNultForCircle = ActiveStabilityValue(batch.ResultAt(1).StabilityResult.Nultimate1, batch.ResultAt(1).StabilityResult.Nultimate2)
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

    SP35InterpolatedPhiM = ActiveStabilityValue(batch.ResultAt(1).StabilityResult.PhiM1, batch.ResultAt(1).StabilityResult.PhiM2)
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

Private Sub AssertEquals(ByRef stats As TBatchTestStats, ByVal name As String, _
        ByVal actual As String, ByVal expected As String)
    If actual = expected Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name & "; actual=" & actual & "; expected=" & expected
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

' ============================== ДЛЯ ТЕСТОВ ==============================

' Воспроизводит осевой RectSet-маршрут Off с полным численным журналом.
' Возвращает диагностику без изменения ожидаемых значений штатных тестов;
' все временные настройки восстанавливаются до выхода, в том числе при ошибке.
Public Function RunAudit02OffAxialDiagnostic() As String
    On Error GoTo Failed
    Dim keys As Variant
    keys = Array("General.DiagramExtension", "General.ExecutionReportEnabled", _
        "Capacity.SolutionStrategy", "Capacity.BaseLoadSteps", "Capacity.MaxRetries")
    Dim previous(0 To 4) As String
    Dim i As Long
    For i = 0 To 4
        previous(i) = GetSystemSetting(CStr(keys(i)))
    Next i
    SetSystemSetting "General.DiagramExtension", "No"
    SetSystemSetting "General.ExecutionReportEnabled", "Yes"
    SetSystemSetting "Capacity.SolutionStrategy", "LoadMultiplier"
    SetSystemSetting "Capacity.BaseLoadSteps", "1"
    SetSystemSetting "Capacity.MaxRetries", "0"
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
    Dim report As CExecutionReport
    Set report = New CExecutionReport
    report.Initialize ThisWorkbook, settings
    Set batch.ExecutionReport = report
    batch.AddCombination "AUDIT02_OFF_N200", 200# * 9806.65, 0#, 0#, _
        "PR1", "Off axial diagnostic", ChrW$(&H3BB) & "*N"
    batch.ApplyLoadReference referenceX, referenceY, referenceX, referenceY
    batch.Execute
    RunAudit02OffAxialDiagnostic = "STATUS: " & batch.ResultAt(1).StrengthResult.Capacity.Status & vbCrLf & _
        batch.ResultAt(1).StrengthResult.Capacity.ResultMeta.ResultComment & vbCrLf & _
        batch.ResultAt(1).StrengthResult.Capacity.DiagnosticLog
Restore:
    For i = 0 To 4
        SetSystemSetting CStr(keys(i)), previous(i)
    Next i
    Exit Function
Failed:
    RunAudit02OffAxialDiagnostic = "RUNTIME ERROR: " & Err.Description
    Resume Restore
End Function

' Воспроизводит formation-маршрут пакетного Г-сечения с полным журналом
' общего Search. Использует тот же материал и ту же нагрузку, что regression;
' временно включается только отчет, физические настройки не подменяются.
Public Function RunAudit02FormationDiagnostic() As String
    On Error GoTo Failed
    Dim oldReport As String
    oldReport = GetSystemSetting("General.ExecutionReportEnabled")
    SetSystemSetting "General.ExecutionReportEnabled", "Yes"
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Dim referenceX As Double
    Dim referenceY As Double
    Dim batch As CBatchSectionCalculator
    Set batch = BuildUserRectSetTensionBatch(referenceX, referenceY)
    batch.ApplySettings settings
    Dim report As CExecutionReport
    Set report = New CExecutionReport
    report.Initialize ThisWorkbook, settings
    Set batch.ExecutionReport = report
    batch.AddCombination "AUDIT02_FORMATION", 20# * 9806.65, 0#, 0#, "PR2", "Formation diagnostic"
    batch.Execute
    RunAudit02FormationDiagnostic = "STATUS: " & batch.ResultAt(1).NormalCrackStatus & vbCrLf & _
        MetaDebugText(batch.ResultAt(1).CrackFormationMeta) & vbCrLf & _
        batch.ResultAt(1).CrackResult.Formation.DiagnosticLog
Restore:
    SetSystemSetting "General.ExecutionReportEnabled", oldReport
    Exit Function
Failed:
    RunAudit02FormationDiagnostic = "RUNTIME ERROR: " & Err.Description
    Resume Restore
End Function

' ============================== ДЛЯ ТЕСТОВ AUDIT02 CACHE ==============================

' Запускает направленную приемку scoped repository отдельно от Excel writer-ов.
' Проверяет реальные solve/cache-hit, смену контекста и повтор после неудачи.
Public Function RunAudit02RepositoryContextTests() As String
    On Error GoTo Failed
    Dim stats As TBatchTestStats
    TestAudit02RepositoryContextAndRetry stats
    AppendLine stats, "TOTAL_AUDIT02_CACHE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit02RepositoryContextTests = stats.Report
    Exit Function
Failed:
    RunAudit02RepositoryContextTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' Подтверждает, что reuse зависит от физической модели и точности приемки,
' а не от retry/warm-start. Уже выданный snapshot не меняется при новом solve.
Private Sub TestAudit02RepositoryContextAndRetry(ByRef stats As TBatchTestStats)
    Dim section As CSectionModel
    Set section = BuildCircleStabilitySection(300#, 8, 16#)
    Dim concrete As CConcreteMaterialParameters
    Set concrete = New CConcreteMaterialParameters
    concrete.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    Dim steel As CSteelMaterialParameters
    Set steel = New CSteelMaterialParameters
    steel.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#
    Dim materials As CMaterialModelProvider
    Set materials = New CMaterialModelProvider
    materials.InitializeFromParameters concrete, steel
    Dim repository As CStateRepository
    Set repository = New CStateRepository
    Dim provider As CStateProvider
    Set provider = New CStateProvider
    provider.Initialize section, materials, repository
    Dim spec As CMaterialModelSpec
    Set spec = New CMaterialModelSpec
    spec.Initialize "ULS(I)", "ThreeLine", "Ignore", "TwoLine"
    Dim request As CStateRequest
    Set request = New CStateRequest
    request.Initialize sstStrengthState, cpStrength, spec, -100000#, 4000000#, 3000000#, True, False
    Dim first As CSectionStateResult
    Set first = provider.GetOrSolve(request)
    AssertTrue stats, "audit02.cache.initial.converged", first.Converged
    AssertTrue stats, "audit02.cache.initial.runner", Not provider.LastRunner Is Nothing
    Dim firstEps As Double
    firstEps = first.Epsilon0
    Dim firstMetaCode As EResultCode
    firstMetaCode = first.ResultCode
    Dim firstLog As String
    firstLog = first.DiagnosticLog

    Dim retryRequest As CStateRequest
    Set retryRequest = New CStateRequest
    retryRequest.Initialize sstStrengthState, cpStrength, spec, -100000#, 4000000#, 3000000#, True, True
    provider.LoadSteps = 3
    provider.DiagnosticsEnabled = False
    Dim state As CSectionStateResult
    Set state = provider.GetOrSolve(retryRequest, provider.SolverSnapshot(first))
    AssertTrue stats, "audit02.cache.options.sameState", state Is first
    AssertTrue stats, "audit02.cache.options.reused", provider.LastStateWasReused
    AssertTrue stats, "audit02.cache.options.noSolve", provider.LastRunner Is Nothing

    materials.InitializeFromParameters concrete, steel, diagramExtensionEnabled:=False
    Set state = provider.GetOrSolve(request)
    AssertTrue stats, "audit02.cache.mode.invalidated", Not provider.LastStateWasReused
    AssertTrue stats, "audit02.cache.mode.newState", Not state Is first
    AssertTrue stats, "audit02.cache.mode.physical", state.Converged And state.WithinPhysicalRange
    AssertTrue stats, "audit02.cache.mode.noExtension", Not state.ExtensionUsed
    AssertClose stats, "audit02.cache.mode.epsilon", state.Epsilon0, firstEps, 0.00000000001

    Dim materialState As CSectionStateResult
    Set materialState = state
    concrete.Initialize 15.5, 1.1, 22#, 1.8, 27000#, 27000#, rbMc2:=14.6
    AssertClose stats, "audit02.cache.material.ownedInput", materials.ConcreteParameters.Eb, 32500#, 0#
    Dim detached As CConcreteMaterialParameters
    Set detached = materials.ConcreteParameters
    detached.Initialize 15.5, 1.1, 22#, 1.8, 25000#, 25000#, rbMc2:=14.6
    AssertClose stats, "audit02.cache.material.detachedGetter", materials.ConcreteParameters.Eb, 32500#, 0#
    materials.InitializeFromParameters concrete, steel, diagramExtensionEnabled:=False
    Set state = provider.GetOrSolve(request)
    AssertTrue stats, "audit02.cache.material.invalidated", Not provider.LastStateWasReused
    AssertTrue stats, "audit02.cache.material.changedPlane", Abs(state.Epsilon0 - materialState.Epsilon0) > 0.000000001
    AssertClose stats, "audit02.cache.material.targetN", state.Nint, request.TargetN, 5#
    AssertClose stats, "audit02.cache.material.targetMx", state.Mxint, request.TargetMx, 5000#
    AssertClose stats, "audit02.cache.material.targetMy", state.Myint, request.TargetMy, 5000#

    Dim revision As Long
    revision = section.Revision
    section.SourceType = "Changed presentation label"
    AssertTrue stats, "audit02.cache.geometry.labelNotPhysical", section.Revision = revision
    section.AddRebarElement 0#, 0#, 20#, 0#, "A400"
    AssertTrue stats, "audit02.cache.geometry.revision", section.Revision > revision
    Dim beforeGeometry As CSectionStateResult
    Set beforeGeometry = state
    Set state = provider.GetOrSolve(request)
    AssertTrue stats, "audit02.cache.geometry.invalidated", Not provider.LastStateWasReused
    AssertTrue stats, "audit02.cache.geometry.changedPlane", Abs(state.Epsilon0 - beforeGeometry.Epsilon0) > 0.0000000001

    provider.ToleranceN = 0.01
    provider.ToleranceMx = 1#
    provider.ToleranceMy = 1#
    Set state = provider.GetOrSolve(request)
    AssertTrue stats, "audit02.cache.tolerance.invalidated", Not provider.LastStateWasReused
    AssertClose stats, "audit02.cache.tolerance.N", state.Nint, request.TargetN, 0.01
    AssertClose stats, "audit02.cache.tolerance.Mx", state.Mxint, request.TargetMx, 1#
    AssertClose stats, "audit02.cache.tolerance.My", state.Myint, request.TargetMy, 1#
    Dim stricterState As CSectionStateResult
    Set stricterState = state
    Dim otherRequest As CStateRequest
    Set otherRequest = New CStateRequest
    otherRequest.Initialize sstStrengthState, cpStrength, spec, -150000#, 2000000#, 1000000#, True, True
    Set state = provider.GetOrSolve(otherRequest)
    AssertTrue stats, "audit02.cache.named.latestSolve", repository.FindState(sstStrengthState) Is state
    Set state = provider.GetOrSolve(request)
    AssertTrue stats, "audit02.cache.named.oldLoadReusable", state Is stricterState
    AssertTrue stats, "audit02.cache.named.latestCacheHit", repository.FindState(sstStrengthState) Is state
    AssertTrue stats, "audit02.cache.named.singleSlot", repository.StateCount = 1

    Dim slsSpec As CMaterialModelSpec
    Set slsSpec = New CMaterialModelSpec
    slsSpec.Initialize "SLS(II)", "ThreeLine", "UseDiagram", "TwoLine"
    otherRequest.Initialize sstPreCrackState, cpMcrc, slsSpec, -100000#, 4000000#, 3000000#, True, True
    Set state = provider.GetOrSolve(otherRequest)
    AssertTrue stats, "audit02.cache.role.notMixed", Not provider.LastStateWasReused
    AssertEquals stats, "audit02.cache.role.spec", state.MaterialSpec.SpecKey, slsSpec.SpecKey
    Set state = provider.GetOrSolve(otherRequest)
    AssertTrue stats, "audit02.cache.pre.reused", provider.LastStateWasReused
    AssertTrue stats, "audit02.cache.pre.noRunner", provider.LastRunner Is Nothing
    AssertTrue stats, "audit02.cache.pre.restored", Not provider.SolverSnapshot(state) Is Nothing

    Dim postSpec As CMaterialModelSpec
    Set postSpec = New CMaterialModelSpec
    postSpec.Initialize "SLS(II)", "TwoLine", "Ignore", "TwoLine"
    otherRequest.Initialize sstPostCrackState, cpCrackedNDS, postSpec, -100000#, 4000000#, 3000000#, True, True
    Set state = provider.GetOrSolve(otherRequest)
    AssertTrue stats, "audit02.cache.post.physical", state.Converged And state.WithinPhysicalRange
    Dim solvesBeforePostReuse As Long
    solvesBeforePostReuse = SectionEquilibriumSolveCount()
    Set state = provider.GetOrSolve(otherRequest)
    AssertTrue stats, "audit02.cache.post.reused", provider.LastStateWasReused
    AssertTrue stats, "audit02.cache.post.noRunner", provider.LastRunner Is Nothing
    AssertTrue stats, "audit02.cache.post.restored", Not provider.SolverSnapshot(state) Is Nothing
    AssertTrue stats, "audit02.cache.post.noSolve", SectionEquilibriumSolveCount() = solvesBeforePostReuse

    repository.Clear
    provider.LoadSteps = 1
    provider.MaxIterations = 1
    provider.LineSearchEnabled = False
    provider.DampingInitial = 0.1
    provider.MaxDeltaEpsilon0 = 0.00000001
    provider.MaxDeltaKappa = 0.000000000001
    Set state = provider.GetOrSolve(request)
    AssertTrue stats, "audit02.cache.retry.firstFailed", Not state.Converged
    AssertTrue stats, "audit02.cache.retry.notReusable", repository.FindEquivalent(request) Is Nothing
    provider.LoadSteps = 8
    provider.MaxIterations = 80
    provider.LineSearchEnabled = True
    provider.DampingInitial = 1#
    provider.MaxDeltaEpsilon0 = 0.0005
    provider.MaxDeltaKappa = 0.00001
    Set state = provider.GetOrSolve(retryRequest)
    AssertTrue stats, "audit02.cache.retry.newSolve", Not provider.LastStateWasReused
    AssertTrue stats, "audit02.cache.retry.success", state.Converged
    AssertTrue stats, "audit02.cache.retry.latestNamed", repository.FindState(sstStrengthState) Is state
    Set state = provider.GetOrSolve(request)
    AssertTrue stats, "audit02.cache.retry.reusedAfterSuccess", provider.LastStateWasReused
    AssertClose stats, "audit02.cache.snapshot.epsilonUnchanged", first.Epsilon0, firstEps, 0#
    AssertTrue stats, "audit02.cache.snapshot.codeUnchanged", first.ResultCode = firstMetaCode
    AssertEquals stats, "audit02.cache.snapshot.logUnchanged", first.DiagnosticLog, firstLog
End Sub

' ============================== ДЛЯ ТЕСТОВ AUDIT02 ON/OFF ==============================

' Возвращает полный численный отчет парного сравнения физических сценариев.
' Переключатель меняет только provider диаграмм, а не усилия, профили или допуски.
Public Function RunAudit02OnOffComparisonTests() As String
    On Error GoTo Failed
    Dim stats As TBatchTestStats
    TestAudit02OnOffPhysicalResults stats
    TestAudit02InitialOffsetOnOffStatuses stats
    AppendLine stats, "TOTAL_AUDIT02_PAIR: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit02OnOffComparisonTests = stats.Report
    Exit Function
Failed:
    RunAudit02OnOffComparisonTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' Сравнивает одноосные, двухосные и осевые LC двух material roles. Второй
' проход включает устойчивость; ее формулы не получают численного Extension.
' Все изменения тестового Config восстанавливаются и при runtime-ошибке.
Private Sub TestAudit02OnOffPhysicalResults(ByRef stats As TBatchTestStats)
    On Error GoTo Failed
    Dim oldStability As String
    oldStability = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    Dim oldLength As String
    oldLength = GetSystemSetting("Stability.ElementLength")
    ' Небольшое сечение с L=8000 может не иметь допустимого текущего LC после
    ' учета устойчивости. Эта матрица сравнивает именно физические решения.
    SetSystemSetting "Stability.ElementLength", 1000#
    Dim maxima As Object
    Set maxima = CreateObject("Scripting.Dictionary")
    Dim phase As Long
    For phase = 0 To 1
        If phase = 0 Then
            SetProfileValue "Calculation.Stability.Enabled", "PR1", "No"
        Else
            SetProfileValue "Calculation.Stability.Enabled", "PR1", "Yes"
        End If
        Dim offBatch As CBatchSectionCalculator
        Dim onBatch As CBatchSectionCalculator
        Set offBatch = BuildAudit02PairBatch(False)
        Set onBatch = BuildAudit02PairBatch(True)
        Dim scenario As Variant
        For Each scenario In Array( _
                Array("ULS_BIAX", -100000#, 8000000#, 3000000#, "PR1"), _
                Array("ULS_X", -100000#, 4000000#, 0#, "PR1"), _
                Array("ULS_Y", -100000#, 0#, 4000000#, "PR1"), _
                Array("ULS_AXIAL", 50000#, 0#, 0#, "PR1"), _
                Array("SLS_BEND", -20000#, -15000000#, 0#, "PR2"), _
                Array("SLS_AXIAL", 200000#, 0#, 0#, "PR2"), _
                Array("SLS_NOCRACK", -100000#, 0#, 0#, "PR2"))
            offBatch.AddCombination CStr(scenario(0)), CDbl(scenario(1)), CDbl(scenario(2)), _
                CDbl(scenario(3)), CStr(scenario(4)), "Audit02 paired physical result"
            onBatch.AddCombination CStr(scenario(0)), CDbl(scenario(1)), CDbl(scenario(2)), _
                CDbl(scenario(3)), CStr(scenario(4)), "Audit02 paired physical result"
        Next scenario
        offBatch.Execute
        onBatch.Execute
        Dim i As Long
        For i = 1 To offBatch.Count
            Dim prefix As String
            prefix = "audit02.pair." & CStr(phase) & "." & offBatch.CombinationID(i)
            Dim offResult As CCombinationResult
            Dim onResult As CCombinationResult
            Set offResult = offBatch.ResultAt(i)
            Set onResult = onBatch.ResultAt(i)
            AppendLine stats, "PAIR_STATUS|" & prefix & "|Off=" & offResult.Status & "|On=" & onResult.Status & _
                "|OffComment=" & offResult.OverallMeta.ResultComment & "|OnComment=" & onResult.OverallMeta.ResultComment
            AssertEquals stats, prefix & ".overall", onResult.Status, offResult.Status
            AssertTrue stats, prefix & ".offNoNumericalFailure", _
                offResult.Status <> "NumFail" And offResult.Status <> "CalcErr" And offResult.Status <> "InputErr"
            AssertEquals stats, prefix & ".direct", onResult.StrengthResult.DirectState.Status, offResult.StrengthResult.DirectState.Status
            AssertEquals stats, prefix & ".capacity", onResult.StrengthResult.Capacity.Status, offResult.StrengthResult.Capacity.Status
            AssertEquals stats, prefix & ".criterion", onResult.StrengthResult.Capacity.LimitState, offResult.StrengthResult.Capacity.LimitState
            AssertEquals stats, prefix & ".normal", onResult.NormalCrackStatus, offResult.NormalCrackStatus
            AssertEquals stats, prefix & ".longitudinal", onResult.CrackResult.Longitudinal.Status, offResult.CrackResult.Longitudinal.Status
            AssertEquals stats, prefix & ".stability", onResult.StabilityResult.Status, offResult.StabilityResult.Status
            AssertTrue stats, prefix & ".crackAvailability", onResult.CrackResult.Formation.CrackFormed = offResult.CrackResult.Formation.CrackFormed
            AssertTrue stats, prefix & ".pointAvailability", onResult.CrackResult.Formation.HasLimitPoint = offResult.CrackResult.Formation.HasLimitPoint
            AssertTrue stats, prefix & ".stateCount", onResult.StateRepository.StateCount = offResult.StateRepository.StateCount

            Dim stateType As Variant
            For Each stateType In Array(sstStrengthState, sstCapacityState, sstPreCrackState, sstPostCrackState, sstCrackedState)
                Audit02ComparePairState stats, maxima, prefix, _
                    offResult.StateRepository.FindState(CLng(stateType)), onResult.StateRepository.FindState(CLng(stateType))
            Next stateType
            Dim offNumbers As Variant
            Dim onNumbers As Variant
            offNumbers = Audit02PairResultNumbers(offResult)
            onNumbers = Audit02PairResultNumbers(onResult)
            Dim metric As Long
            For metric = LBound(offNumbers) To UBound(offNumbers)
                Audit02AssertPair stats, maxima, prefix, CStr(offNumbers(metric)(0)), _
                    CDbl(onNumbers(metric)(1)), CDbl(offNumbers(metric)(1)), CDbl(offNumbers(metric)(2))
            Next metric
            Audit02AssertPair stats, maxima, prefix, "Reference.N", onBatch.StrengthReferenceN(i), offBatch.StrengthReferenceN(i), 0.001
            Audit02AssertPair stats, maxima, prefix, "Reference.Mx", onBatch.StrengthReferenceMx(i), offBatch.StrengthReferenceMx(i), 0.001
            Audit02AssertPair stats, maxima, prefix, "Reference.My", onBatch.StrengthReferenceMy(i), offBatch.StrengthReferenceMy(i), 0.001
        Next i
    Next phase
    Dim key As Variant
    For Each key In maxima.Keys
        AppendLine stats, "MAXDIFF|" & CStr(key) & "|abs=" & FormatNumberInvariant(CDbl(maxima(key)))
    Next key
Restore:
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldStability
    SetSystemSetting "Stability.ElementLength", oldLength
    Exit Sub
Failed:
    Dim reason As String
    reason = Err.Description
    AssertTrue stats, "audit02.pair.runtime: " & reason, False
    Resume Restore
End Sub

' Строит одинаковую геометрию и физические параметры, явно выбирая On/Off.
' Настройки прочих алгоритмов и профили читаются одинаково для обеих половин пары.
Private Function BuildAudit02PairBatch(ByVal extensionEnabled As Boolean) As CBatchSectionCalculator
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
    Dim materials As CMaterialModelProvider
    Set materials = New CMaterialModelProvider
    Dim parameters As CMaterialModelProvider
    Set parameters = TestMaterialProvider()
    materials.InitializeFromParameters parameters.ConcreteParameters, parameters.SteelParameters, _
        diagramExtensionEnabled:=extensionEnabled
    Dim batch As CBatchSectionCalculator
    Set batch = New CBatchSectionCalculator
    batch.Initialize BuildGeneratedSectionModel(mesh, rebars, "Audit02PairedRectangle"), materials
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    batch.ApplySettings settings
    batch.SetSP35Table721 ThisWorkbook.Names.Item("rngSP35Table721").RefersToRange.Value2
    Set batch.ProfileCatalog = TestProfileCatalog()
    batch.ApplyLoadReference 0#, 0#
    Set BuildAudit02PairBatch = batch
End Function

' ДЛЯ ТЕСТОВ
' При постоянной сжимающей N за физической несущей обычные диаграммы не
' дают равновесия, а продолжение позволяет подтвердить недопустимый старт.
' Проверяем разницу NumFail/BaseFail по машинным кодам, не по тексту причины,
' и согласованность capacity-статуса в typed result, detailed и batch summary.
Private Sub TestAudit02InitialOffsetOnOffStatuses(ByRef stats As TBatchTestStats)
    On Error GoTo Failed
    Dim oldStability As String
    oldStability = GetProfileValue("Calculation.Stability.Enabled", "PR1")
    SetProfileValue "Calculation.Stability.Enabled", "PR1", "No"
    Dim offBatch As CBatchSectionCalculator
    Dim onBatch As CBatchSectionCalculator
    Set offBatch = BuildAudit02PairBatch(False)
    Set onBatch = BuildAudit02PairBatch(True)
    offBatch.AddCombination "OFFSET_2MN", -2000000#, 8000000#, 0#, "PR1", "Постоянная N выше физического предела", "lambda*Mxy"
    offBatch.AddCombination "OFFSET_3MN", -3000000#, 8000000#, 0#, "PR1", "Постоянная N выше физического предела", "lambda*Mxy"
    onBatch.AddCombination "OFFSET_2MN", -2000000#, 8000000#, 0#, "PR1", "Постоянная N выше физического предела", "lambda*Mxy"
    onBatch.AddCombination "OFFSET_3MN", -3000000#, 8000000#, 0#, "PR1", "Постоянная N выше физического предела", "lambda*Mxy"
    offBatch.Execute
    onBatch.Execute
    Dim i As Long
    For i = 1 To 2
        Dim offResult As CCapacityResult
        Dim onResult As CCapacityResult
        Set offResult = offBatch.ResultAt(i).StrengthResult.Capacity
        Set onResult = onBatch.ResultAt(i).StrengthResult.Capacity
        Dim prefix As String
        prefix = "audit02.offsetPair." & offBatch.CombinationID(i)
        AppendLine stats, "OFFSET_PAIR|" & offBatch.CombinationID(i) & "|Off=" & offResult.Status & _
            "|OffCode=" & CStr(offResult.ResultMeta.ResultCode) & "|On=" & onResult.Status & _
            "|OnCode=" & CStr(onResult.ResultMeta.ResultCode) & "|OffComment=" & offResult.ResultMeta.ResultComment & _
            "|OnComment=" & onResult.ResultMeta.ResultComment
        AssertEquals stats, prefix & ".offStatus", offResult.Status, "NumFail"
        AssertEquals stats, prefix & ".onStatus", onResult.Status, "BaseFail"
        ' Audit03 F03: сохраняется точная причина касательной системы; внешний
        ' NumFail, физические усилия и expected/tolerance парного теста прежние.
        AssertTrue stats, prefix & ".offCode", offResult.ResultMeta.ResultCode = rcSingularTangent
        AssertTrue stats, prefix & ".onCode", onResult.ResultMeta.ResultCode = rcInitialStateBeyondLimit
        AssertTrue stats, prefix & ".offNoPoint", Not offResult.SearchResult.HasLimitPoint
        AssertTrue stats, prefix & ".onNoPoint", Not onResult.SearchResult.HasLimitPoint
        AssertTrue stats, prefix & ".onAuxiliary", onBatch.ResultAt(i).StrengthResult.DirectState.StateResult.Converged
        AssertTrue stats, prefix & ".onExtended", onBatch.ResultAt(i).StrengthResult.DirectState.StateResult.ExtensionUsed
        AssertTrue stats, prefix & ".onNotPhysical", Not onBatch.ResultAt(i).StrengthResult.DirectState.StateResult.WithinPhysicalRange
        AssertTrue stats, prefix & ".comments", Len(offResult.ResultMeta.ResultComment) > 0 And Len(onResult.ResultMeta.ResultComment) > 0
    Next i
    Dim mode As Long
    For mode = 0 To 1
        Dim batch As CBatchSectionCalculator
        If mode = 0 Then Set batch = offBatch Else Set batch = onBatch
        Dim writer As CBatchResultWriter
        Set writer = New CBatchResultWriter
        writer.WriteSummary ThisWorkbook, batch
        Dim sheet As Object
        Set sheet = ThisWorkbook.Worksheets.Item("Results")
        For i = 1 To 2
            Dim summaryRow As Long, detailedRow As Long
            summaryRow = SummaryRowByCombination(sheet, batch.CombinationID(i))
            detailedRow = DetailedRowByCombination(sheet, "rngStrengthSummaryAnchor", batch.CombinationID(i))
            prefix = "audit02.offsetPair.sheet." & CStr(mode) & "." & CStr(i)
            AssertEquals stats, prefix & ".summary", CStr(sheet.Cells.Item(summaryRow, 7).Value2), _
                batch.ResultAt(i).StrengthResult.Capacity.Status
            AssertEquals stats, prefix & ".detailed", CStr(sheet.Cells.Item(detailedRow, 49).Value2), _
                batch.ResultAt(i).StrengthResult.Capacity.Status
            AssertEquals stats, prefix & ".overall", CStr(sheet.Cells.Item(summaryRow, 4).Value2), batch.ResultAt(i).Status
            AssertTrue stats, prefix & ".colorSummary", StatusCellHasExpectedFill(sheet, summaryRow, 7)
            AssertTrue stats, prefix & ".colorDetailed", StatusCellHasExpectedFill(sheet, detailedRow, 49)
        Next i
    Next mode
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldStability
    Exit Sub
Failed:
    Dim reason As String
    reason = CStr(Err.Number) & "; " & Err.Description
    SetProfileValue "Calculation.Stability.Enabled", "PR1", oldStability
    AssertTrue stats, "audit02.offsetPair.runtime." & reason, False
End Sub

' Сравнивает доступность и весь численный снимок named-state, исключая
' итерации/время/журнал: эти показатели не являются условием одинаковой физики.
Private Sub Audit02ComparePairState(ByRef stats As TBatchTestStats, ByVal maxima As Object, _
        ByVal prefix As String, ByVal offState As CSectionStateResult, ByVal onState As CSectionStateResult)
    If offState Is Nothing Then
        AssertTrue stats, prefix & ".missingStateMatched", onState Is Nothing
        Exit Sub
    End If
    Dim key As String
    key = offState.StateTypeText
    AssertTrue stats, prefix & "." & key & ".exists", Not onState Is Nothing
    If onState Is Nothing Then Exit Sub
    AssertEquals stats, prefix & "." & key & ".status", onState.Status, offState.Status
    AssertTrue stats, prefix & "." & key & ".converged", onState.Converged And offState.Converged
    AssertTrue stats, prefix & "." & key & ".physical", onState.WithinPhysicalRange And offState.WithinPhysicalRange
    AssertTrue stats, prefix & "." & key & ".noFinalExtension", Not onState.ExtensionUsed And Not offState.ExtensionUsed
    Dim first As Variant
    Dim second As Variant
    first = Audit02PairStateNumbers(offState)
    second = Audit02PairStateNumbers(onState)
    Dim i As Long
    For i = LBound(first) To UBound(first)
        Audit02AssertPair stats, maxima, prefix, key & "." & CStr(first(i)(0)), _
            CDbl(second(i)(1)), CDbl(first(i)(1)), CDbl(first(i)(2))
    Next i
End Sub

' Подготавливает численные поля State с прежними допусками solver/strain-
' тестов; технические отличия поисковой стратегии сюда не включаются.
Private Function Audit02PairStateNumbers(ByVal state As CSectionStateResult) As Variant
    Audit02PairStateNumbers = Array( _
        Array("TargetN", state.TargetN, 5#), Array("TargetMx", state.TargetMx, 5000#), Array("TargetMy", state.TargetMy, 5000#), _
        Array("Nint", state.Nint, 5#), Array("Mxint", state.Mxint, 5000#), Array("Myint", state.Myint, 5000#), _
        Array("epsilon0", state.Epsilon0, 0.000001), Array("kappaX", state.KappaX, 0.00000000001), Array("kappaY", state.KappaY, 0.00000000001), _
        Array("epsBmin", state.MinConcreteStrain, 0.000001), Array("epsBmax", state.MaxConcreteStrain, 0.000001), _
        Array("epsSmin", state.MinSteelStrain, 0.000001), Array("epsSmax", state.MaxSteelStrain, 0.000001), _
        Array("sigmaBmin", state.MinConcreteStress, 0.001), Array("sigmaBmax", state.MaxConcreteStress, 0.001), _
        Array("sigmaSmin", state.MinSteelStress, 0.001), Array("sigmaSmax", state.MaxSteelStress, 0.001), _
        Array("ResidualN", state.ResidualN, 5#), Array("ResidualMx", state.ResidualMx, 5000#), Array("ResidualMy", state.ResidualMy, 5000#))
End Function

' Подготавливает метрики всех инженерных ветвей для парного сравнения.
' Неприменимые ветви сравниваются также по доступности и typed status выше.
Private Function Audit02PairResultNumbers(ByVal result As CCombinationResult) As Variant
    Dim capacity As CCapacityResult
    Set capacity = result.StrengthResult.Capacity
    Dim formation As CCrackFormationResult
    Set formation = result.CrackResult.Formation
    Dim width As CCrackWidthResult
    Set width = result.CrackResult.Width
    Dim stability As CStabilityResult
    Set stability = result.StabilityResult
    Audit02PairResultNumbers = Array( _
        Array("Capacity.lambda", capacity.LambdaCapacity, 0.01), Array("Capacity.N", capacity.NUltimate, 5#), _
        Array("Capacity.Mx", capacity.MxUltimate, 5000#), Array("Capacity.My", capacity.MyUltimate, 5000#), _
        Array("Capacity.util", capacity.UtilCapacity, 0.000001), Array("Formation.lambda", formation.LambdaCrc, 0.001), _
        Array("Formation.Ncrc", formation.Ncrc, 0.001), Array("Formation.N", formation.FormationNcrc, 0.001), _
        Array("Formation.Mcrc", formation.Mcrc, 50000#), Array("Formation.Ared", formation.Ared, 0.001), _
        Array("Width.acrc", width.CrackWidth, 0.000000001), Array("Width.sigmaS", width.SigmaS, 0.001), _
        Array("Width.sigmaSCrc", width.SigmaSCrc, 0.001), Array("Width.psi", width.PsiS, 0.000000001), _
        Array("Width.ls", width.CrackSpacing, 0.001), Array("Width.lsRaw", width.CrackSpacingRaw, 0.001), _
        Array("Width.Abt", width.Abt, 0.001), Array("Width.As", width.AsTension, 0.001), Array("Width.ds", width.DsEquivalent, 0.001), _
        Array("Width.xt", width.TensionDepth, 0.001), Array("Width.hbt", width.EffectiveZoneDepth, 0.001), _
        Array("Width.a", width.CoverA, 0.001), Array("Width.h", width.SectionDepthH, 0.001), Array("Width.Es", width.SteelEs, 0.001), _
        Array("Width.util", width.Utilization, 0.000000001), Array("Longitudinal.sigma", result.CrackResult.Longitudinal.MaxCompressionStress, 0.001), _
        Array("Longitudinal.util", result.CrackResult.Longitudinal.Utilization, 0.000000001), _
        Array("Stability.N", stability.DesignN, 0.001), Array("Stability.Mx", stability.DesignMx, 0.001), _
        Array("Stability.My", stability.DesignMy, 0.001), Array("Stability.Ncr1", stability.Ncr1, 0.001), _
        Array("Stability.Ncr2", stability.Ncr2, 0.001), Array("Stability.reserve", stability.SummaryReserve, 0.000000001))
End Function

' Фиксирует отклонение каждой применимой величины и максимальное абсолютное
' отклонение по метрике во всех сценариях; существующие expected не меняются.
Private Sub Audit02AssertPair(ByRef stats As TBatchTestStats, ByVal maxima As Object, _
        ByVal prefix As String, ByVal metric As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal tolerance As Double)
    AssertClose stats, prefix & "." & metric, actual, expected, tolerance
    Dim difference As Double
    difference = Abs(actual - expected)
    If Not maxima.Exists(metric) Then maxima.Add metric, 0#
    If difference > CDbl(maxima(metric)) Then maxima(metric) = difference
End Sub

' ===========================================================================
' ДЛЯ ТЕСТОВ: AUDIT03 - ЖИЗНЕННЫЙ ЦИКЛ МЕТАДАННЫХ
' ===========================================================================

' Выполняет направленную проверку flags/clone/reset и ранних result-фабрик.
' Это отдельный вход для воспроизводимого F04 без полного batch-прогона.
Public Function RunAudit03LifecycleTests() As String
    On Error GoTo Failed
    Dim stats As TBatchTestStats
    TestAudit03ResultLifecycle stats
    AppendLine stats, "TOTAL_AUDIT03_LIFECYCLE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03LifecycleTests = stats.Report
    Exit Function
Failed:
    RunAudit03LifecycleTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' Проверяет каждый внутренний исход: зависимая/выключенная ветка не объявляет
' собственный расчет выполненным, а реальная численная попытка сохраняется.
' Повторное заполнение и независимый clone не переносят флаги прошлого LC.
Private Sub TestAudit03ResultLifecycle(ByRef stats As TBatchTestStats)
    Dim statuses As Variant
    statuses = Array(rsSuccess, rsCheckFailed, rsNumericalFailure, rsInvalidInput, _
        rsInvalidConfiguration, rsBlockedByDependency, rsNotApplicable, rsNotRequested, _
        rsSuccessWithWarning, rsInternalError)
    Dim value As Variant, attempted As Variant
    Dim meta As CResultMeta, snapshot As CResultMeta
    Set meta = New CResultMeta
    For Each value In statuses
        For Each attempted In Array(False, True)
            Dim expectedApplies As Boolean, expectedCalculated As Boolean
            expectedApplies = (value <> rsNotApplicable And value <> rsNotRequested)
            expectedCalculated = CBool(attempted)
            Select Case value
                Case rsInvalidInput, rsInvalidConfiguration, rsBlockedByDependency, rsNotApplicable, rsNotRequested
                    expectedCalculated = False
            End Select
            Dim prefix As String
            prefix = "audit03.lifecycle." & CStr(value) & "." & CStr(attempted)
            meta.SetResult CLng(value), rcCheckPassed, rkCrackWidth, "Текущий результат.", _
                "Текущая диагностика.", True, CBool(attempted)
            AssertTrue stats, prefix & ".applies", meta.Applies = expectedApplies
            AssertTrue stats, prefix & ".calculated", meta.Calculated = expectedCalculated
            Set snapshot = meta.Clone
            meta.Clear
            AssertTrue stats, prefix & ".cloneStatus", snapshot.InternalStatus = value
            AssertTrue stats, prefix & ".cloneApplies", snapshot.Applies = expectedApplies
            AssertTrue stats, prefix & ".cloneCalculated", snapshot.Calculated = expectedCalculated
            AssertTrue stats, prefix & ".clearCalculated", Not meta.Calculated
            AssertTrue stats, prefix & ".clearCode", meta.ResultCode = rcNone
            AssertEquals stats, prefix & ".clearComment", meta.ResultComment, vbNullString
        Next attempted
    Next value

    Dim capacity As CCapacityResult
    Set capacity = New CCapacityResult
    capacity.InitializeBaseFail "Исходная часть нагрузки уже не проходит проверку."
    AssertTrue stats, "audit03.lifecycle.capacity.baseCalculated", capacity.ResultMeta.Calculated
    capacity.InitializeCalcError "Не получен результат поиска."
    AssertTrue stats, "audit03.lifecycle.capacity.earlyInternal", Not capacity.ResultMeta.Calculated
    AssertTrue stats, "audit03.lifecycle.capacity.earlyApplies", capacity.ResultMeta.Applies

    Dim state As CSectionStateResult, direct As CDirectStateResult
    Set state = New CSectionStateResult
    AssertTrue stats, "audit03.lifecycle.state.empty", Not state.ResultMeta.Calculated
    Set direct = New CDirectStateResult
    direct.SetMeta Nothing
    AssertTrue stats, "audit03.lifecycle.direct.missing", Not direct.Meta.Calculated

    Dim formation As CCrackFormationResult, width As CCrackWidthResult
    Set formation = New CCrackFormationResult
    formation.InitializeFromCalculator Nothing, Nothing
    AssertTrue stats, "audit03.lifecycle.formation.missing", Not formation.ResultMeta.Calculated
    Set width = New CCrackWidthResult
    width.InitializeFromCalculator Nothing, Nothing
    AssertTrue stats, "audit03.lifecycle.width.missing", Not width.ResultMeta.Calculated

    Dim longitudinal As CLongitudinalCrackResult, stability As CStabilityResult
    Set longitudinal = New CLongitudinalCrackResult
    longitudinal.Initialize Nothing, 0#, 0#, 0#
    AssertTrue stats, "audit03.lifecycle.longitudinal.missing", Not longitudinal.ResultMeta.Calculated
    Set stability = New CStabilityResult
    stability.SetMeta Nothing
    AssertTrue stats, "audit03.lifecycle.stability.missing", Not stability.Meta.Calculated

    Dim search As CLimitSearchResult
    Set search = New CLimitSearchResult
    AssertTrue stats, "audit03.lifecycle.search.empty", Not search.Meta.Calculated
    search.Initialize Nothing, "Auto", vbNullString, False, 0#, 0#, 0#, 0#, Nothing, _
        vbNullString, vbNullString, vbNullString, False
    AssertTrue stats, "audit03.lifecycle.search.notStarted", Not search.Meta.Calculated
    search.Initialize Nothing, "Auto", "LoadMultiplier", False, 0#, 0#, 0#, 0#, Nothing, _
        vbNullString, vbNullString, vbNullString, True
    AssertTrue stats, "audit03.lifecycle.search.actualAttempt", search.Meta.Calculated
End Sub
