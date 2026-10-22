Attribute VB_Name = "modTestCapacitySolver"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: Тесты поиска несущей способности
' ==========================================================================
' Модуль проверяет LoadMultiplier, UltimateStrain и одномерные методы поиска
' lambda. Важна не только близость результата, но и запуск правильной ветки.

Private Type TCapacityTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет физические пределы, знаки и общие методы поиска.
' Отчёт различает подтверждённую предельную точку, численный отказ и ошибку ввода.
Public Function RunCapacitySolverTests() As String
    On Error GoTo Failed

    Dim stats As TCapacityTestStats
    Dim t0 As Double
    t0 = Timer

    TestLimitSearchAdapterInterfaces stats
    TestRequestReinitialization stats
    TestCommonStrainLimits stats
    TestCenteredAxialPathClassification stats
    TestAppendLine stats, "RUN: TestMxPositiveAndNegative"
    TestMxPositiveAndNegative stats
    TestAppendLine stats, "RUN: TestMyPositiveAndNegative"
    TestMyPositiveAndNegative stats
    TestAppendLine stats, "RUN: TestMxySignedCombinations"
    TestMxySignedCombinations stats
    TestAppendLine stats, "RUN: TestConcreteTensionBehaviorAffectsSolverAndCapacity"
    TestConcreteTensionBehaviorAffectsSolverAndCapacity stats
    TestAppendLine stats, "RUN: TestCircleCapacitySymmetry"
    TestCircleCapacitySymmetry stats
    TestAppendLine stats, "RUN: TestLambdaLessThanOne"
    TestLambdaLessThanOne stats
    TestAppendLine stats, "RUN: TestLambdaNearOne"
    TestLambdaNearOne stats
    TestAppendLine stats, "RUN: TestZeroAxialForce"
    TestZeroAxialForce stats
    TestAppendLine stats, "RUN: TestAxialLoadMultiplierFindsNult"
    TestAxialLoadMultiplierFindsNult stats
    TestAxialConcretePlateauStart stats
    TestUltimateMomentTranslationInvariant stats
    TestUltimateStrainToleranceNoHiddenFloor stats
    TestUltimateNumericalContextCache stats
    TestNumericalPathPreparation stats
    TestAppendLine stats, "RUN: TestProjectedLimitRecovery"
    TestProjectedLimitRecovery stats
    TestWarmStartRestartContract stats
    TestUltimateStepScale stats
    TestAppendLine stats, "RUN: TestConcreteLimitState"
    TestConcreteLimitState stats
    TestAppendLine stats, "RUN: TestConcreteTensionLimitState"
    TestConcreteTensionLimitState stats
    TestAppendLine stats, "RUN: TestConcreteTensionLimitIgnored"
    TestConcreteTensionLimitIgnored stats
    TestAppendLine stats, "RUN: TestSteelLimitState"
    TestSteelLimitState stats
    TestAppendLine stats, "RUN: TestLoadMultiplierPureBendingUsesStateGuess"
    TestLoadMultiplierPureBendingUsesStateGuess stats
    TestAppendLine stats, "RUN: TestAsymmetricCoupledCurvatures"
    TestAsymmetricCoupledCurvatures stats
    TestAppendLine stats, "RUN: TestAsymmetricMxy"
    TestAsymmetricMxy stats
    TestAppendLine stats, "RUN: TestInvalidBaseMoment"
    TestInvalidBaseMoment stats
    TestAppendLine stats, "RUN: TestCapacitySolutionStrategyComparisons"
    TestCapacitySolutionStrategyComparisons stats
    TestAppendLine stats, "RUN: TestLoadMultiplierWithWorkbookTfDefaults"
    TestLoadMultiplierWithWorkbookTfDefaults stats
    TestAppendLine stats, "RUN: TestLoadMultiplierSearchMethods"
    TestLoadMultiplierSearchMethods stats
    TestAppendLine stats, "RUN: TestLimitSearchSecantFinalizesCheckedRoot"
    TestLimitSearchSecantFinalizesCheckedRoot stats
    TestAppendLine stats, "RUN: TestLimitSearchBisectionIterationLimitFails"
    TestLimitSearchBisectionIterationLimitFails stats
    TestAppendLine stats, "RUN: TestLimitSearchNumericalZeroIsNotRoot"
    TestLimitSearchNumericalZeroIsNotRoot stats
    TestLimitSearchInitialRecoveryBudget stats
    TestLimitSearchAlternatingRecoveryBudget stats
    TestSmallLambdaBracketRecovery stats
    TestCapacityMomentMeritScaling stats
    TestForcePathMaterialPlateau stats
    TestForcePathMaterialPlateau stats, 5
    TestCapacityDiagnosticBuffer stats
    TestAppendLine stats, "RUN: TestAudit02GenericUltimateSearch"
    TestAudit02GenericUltimateSearch stats
    TestAppendLine stats, "RUN: TestAudit02GenericLoadMultiplierMatrix"
    TestAudit02GenericLoadMultiplierMatrix stats
    TestAudit03SearchArithmetic stats
    TestAudit03RealCapacityPrecision stats
    TestAudit03CapacityTypedFailures stats
    TestAudit03CapacitySearchLifecycle stats
    TestAudit03UltimateGuards stats
    TestAudit03MultiplierTypedFaults stats
    TestAppendLine stats, "RUN: TestCapacityLoadPathMethodMatrix"
    TestCapacityLoadPathMethodMatrix stats
    TestAppendLine stats, "RUN: TestCapacityLoadPathZeroComponentMatrix"
    TestCapacityLoadPathZeroComponentMatrix stats
    TestAppendLine stats, "RUN: TestRectanglesCapacityLoadPathSmoke"
    TestRectanglesCapacityLoadPathSmoke stats
    TestAppendLine stats, "RUN: TestNultBaseLoadStepsSensitivity"
    TestNultBaseLoadStepsSensitivity stats
    TestAppendLine stats, "RUN: TestSearchMethodInputErrors"
    TestSearchMethodInputErrors stats
    TestAppendLine stats, "RUN: TestInitialLambdaFailureStatusMapping"
    TestInitialLambdaFailureStatusMapping stats
    TestUltimateInitialPhysicalRejection stats
    TestUltimateInitialNumericalFailure stats
    TestAppendLine stats, "RUN: TestSearchMethodPerformanceComparison"
    TestSearchMethodPerformanceComparison stats

    TestAppendLine stats, "RUN: TestAudit02CapacitySearchBoundary"
    TestAudit02CapacitySearchBoundary stats
    TestAppendLine stats, "RUN: TestAudit02AsymmetricSteelLimits"
    TestAudit02AsymmetricSteelLimits stats
    TestAppendLine stats, "RUN: TestAudit02UnconvergedProbeIsNumerical"
    TestAudit02UnconvergedProbeIsNumerical stats
    TestAudit02PositiveUnconvergedProbe stats
    TestAudit02InitialOffsetBoundary stats

    TestAppendLine stats, "TOTAL_CAPACITY: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & TestFormatNumberInvariant(Timer - t0)
    RunCapacitySolverTests = stats.Report
    Exit Function

Failed:
    RunCapacitySolverTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Отдельно воспроизводит расчёт в tf/tf*m с Rectangles/Generated и проверяет,
' что после чтения контрольных настроек возвращается исходный выбор сечения.
Public Function RunWorkbookTfFixtureTests() As String
    Dim stats As TCapacityTestStats
    On Error GoTo Failed
    TestLoadMultiplierWithWorkbookTfDefaults stats
    TestAppendLine stats, "TOTAL_WORKBOOK_TF_FIXTURE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunWorkbookTfFixtureTests = stats.Report
    Exit Function
Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: workbook.tfFixture.runtime; error=" & CStr(Err.Number) & "; description=" & Err.Description
    TestAppendLine stats, "TOTAL_WORKBOOK_TF_FIXTURE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunWorkbookTfFixtureTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' направленные численные примеры без полного набора поисковых задач.
Public Function RunRetiredGeometryCapacityTests() As String
    Dim stats As TCapacityTestStats
    On Error GoTo Failed
    TestMxPositiveAndNegative stats
    TestMyPositiveAndNegative stats
    TestMxySignedCombinations stats
    TestCircleCapacitySymmetry stats
    TestAsymmetricCoupledCurvatures stats
    TestAsymmetricMxy stats
    TestAppendLine stats, "TOTAL_RETIRED_GEOMETRY_CAPACITY: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunRetiredGeometryCapacityTests = stats.Report
    Exit Function
Failed:
    RunRetiredGeometryCapacityTests = stats.Report & "FAIL: runtime=" & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' выходные компоненты после уточнения физического предела.
' Проверяет реальные усилия отдельно от Nult и компонентов траектории.
Public Function RunCapacityOutputContractTests() As String
    Dim stats As TCapacityTestStats
    On Error GoTo Failed
    TestCircleCapacitySymmetry stats
    TestCapacityLoadPathMethodMatrix stats
    TestCapacityLoadPathZeroComponentMatrix stats
    TestAppendLine stats, "TOTAL_CAPACITY_OUTPUT_CONTRACT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunCapacityOutputContractTests = stats.Report
    Exit Function
Failed:
    RunCapacityOutputContractTests = stats.Report & "FAIL: runtime=" & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' постоянная перегрузка пути проверяется до поиска предельной
' плоскости; неудачный НДС не становится физическим отказом.
Public Function RunOverloadInitialPointTests() As String
    Dim stats As TCapacityTestStats
    On Error GoTo Failed
    TestUltimateInitialPhysicalRejection stats
    TestUltimateInitialNumericalFailure stats
    TestInitialLambdaFailureStatusMapping stats
    TestAppendLine stats, "TOTAL_OVERLOAD_INITIAL_POINT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunOverloadInitialPointTests = stats.Report
    Exit Function
Failed:
    RunOverloadInitialPointTests = stats.Report & "FAIL: overload.initial.runtime; " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Линейные материалы позволяют получить точное равновесие при заведомом
' превышении физической деформации, без эвристической оценки несущей.
Private Sub TestUltimateInitialPhysicalRejection(ByRef stats As TCapacityTestStats)
    Dim section As CSectionModel, strategy As Variant, cap As CCapacitySolver
    Dim request As CLimitSearchRequest, result As CLimitSearchResult, coordinator As CLimitSearchCoordinator
    Dim forceOnly As Variant, prefix As String, nBase As Double, mxOffset As Double, mxBase As Double
    Set section = New CSectionModel
    section.AddConcreteElement -100#, -100#, 10000#
    section.AddConcreteElement 100#, -100#, 10000#
    section.AddConcreteElement -100#, 100#, 10000#
    section.AddConcreteElement 100#, 100#, 10000#
    section.AddRebarElement -80#, -80#, 20#, 0#, "A400"
    section.AddRebarElement 80#, 80#, 20#, 0#, "A400"
    For Each forceOnly In Array(False, True)
    For Each strategy In Array("Auto", "UltimateStrain", "LoadMultiplier")
        nBase = 0#: mxOffset = 0#: mxBase = 1000000#
        If CBool(forceOnly) Then nBase = 100000#: mxOffset = 1000000#: mxBase = 0#
        prefix = "overload.initial." & CStr(strategy) & ".forceOnly=" & CStr(forceOnly)
        Set cap = New CCapacitySolver: TestConfigureCapacity cap
        Set request = New CLimitSearchRequest
        InitializeCapacitySearchRequest request, cap, section, TestLinearConcrete(), TestLinearSteel(), CStr(strategy), _
            -10000000#, nBase, mxOffset, mxBase, 0#, 0#, CBool(forceOnly), False
        Set coordinator = New CLimitSearchCoordinator
        ResetSectionEquilibriumSolveCount
        Set result = coordinator.ExecuteCapacity(request)
        TestAssertTrue stats, prefix & ".baseFailure", result.Meta.ResultCode = rcInitialStateBeyondLimit
        TestAssertTrue stats, prefix & ".notNumericalFailure", result.Meta.InternalStatus = rsCheckFailed
        TestAssertTrue stats, prefix & ".noInventedLimit", Not result.HasLimitPoint
        TestAssertClose stats, prefix & ".zeroLambda", result.LambdaUltimate, 0#, 0#
        TestAssertTrue stats, prefix & ".equilibriumChecked", SectionEquilibriumSolveCount() = 1
        TestAssertTrue stats, prefix & ".noUltimateIterations", UltimateStrainIterationCount() = 0
        TestAssertTrue stats, prefix & ".reason", Len(result.Meta.ResultComment) > 0
    Next strategy
    Next forceOnly
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Неподтвержденное равновесие при нагрузке выше плато без продолжения
' не разрешает объявить физический отказ. UltimateStrain вправе искать предел.
Private Sub TestUltimateInitialNumericalFailure(ByRef stats As TCapacityTestStats)
    Dim section As CSectionModel, cap As CCapacitySolver, request As CLimitSearchRequest, path As CLoadPathVector
    Dim value As Double, state As String, solver As CSectionSolver
    Set section = New CSectionModel
    section.AddConcreteElement -100#, -100#, 10000#
    section.AddConcreteElement 100#, -100#, 10000#
    section.AddConcreteElement -100#, 100#, 10000#
    section.AddConcreteElement 100#, 100#, 10000#
    Set cap = New CCapacitySolver: TestConfigureCapacity cap
    cap.SolverMaxIterations = 1
    cap.MaxRetries = 3
    Set request = New CLimitSearchRequest
    InitializeCapacitySearchRequest request, cap, section, TestProvisionalConcrete(), TestLinearSteel(), "UltimateStrain", _
        -1000000#, 0#, 0#, 1000000#, 0#, 0#, False, False
    Set path = New CLoadPathVector: path.Initialize -1000000#, 0#, 0#, 1000000#, 0#, 0#
    TestAssertTrue stats, "overload.initial.numerical.fixturePrepared", request.Problem.Prepare(request)
    TestAssertTrue stats, "overload.initial.numerical.fixtureEvaluated", _
        request.Problem.EvaluatePoint(request, path, Nothing, 0#, value, state, solver)
    TestAssertTrue stats, "overload.initial.numerical.fixtureReallyFailed", request.Problem.IsNumericalFailureState(state)
    ResetSectionEquilibriumSolveCount
    TestAssertTrue stats, "overload.initial.numerical.searchStillAllowed", request.Problem.PrepareUltimate(request, path)
    TestAssertTrue stats, "overload.initial.numerical.notPhysicalFailure", Not cap.LimitSearchInitialLambdaRejected
    TestAssertTrue stats, "overload.initial.numerical.noRepeatedPreflight", SectionEquilibriumSolveCount() = 1
    TestAssertTrue stats, "overload.initial.numerical.noInventedLimit", Not cap.Converged
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' стартовая плоскость внутри физических пределов, включая перенос на 1000 м.
Public Function RunLimitGuessScalingTests() As String
    Dim stats As TCapacityTestStats
    On Error GoTo Failed
    Dim section As CSectionModel, frame As CSectionNumericalFrame, builder As CStateGuessBuilder
    Dim solver As CSectionSolver, limits As CSectionStrainLimits, shift As Variant, multiplier As Variant
    Dim eps0 As Double, kx As Double, ky As Double, utilization As Double, prefix As String
    For Each shift In Array(0#, 1000000#, -1000000#)
        Set section = New CSectionModel
        section.AddConcreteElement CDbl(shift) - 100#, -CDbl(shift) - 100#, 10000#
        section.AddConcreteElement CDbl(shift) + 100#, -CDbl(shift) + 100#, 10000#
        section.AddRebarElement CDbl(shift) - 80#, -CDbl(shift) - 80#, 20#, 0#, "A400"
        section.AddRebarElement CDbl(shift) + 80#, -CDbl(shift) + 80#, 20#, 0#, "A400"
        Set frame = New CSectionNumericalFrame: frame.Initialize section
        Set builder = New CStateGuessBuilder: Set solver = New CSectionSolver
        Set limits = New CSectionStrainLimits
        limits.Configure -0.0035, 0.00015, 0.001, 0.0025
        For Each multiplier In Array(1#, 1000000#)
            kx = CDbl(multiplier) * 0.00001: ky = CDbl(multiplier) * -0.000006
            eps0 = frame.OriginalStrain(-CDbl(multiplier) * 0.004, kx, ky)
            solver.EvaluatePhysicalStrainPlane section, TestLinearConcrete(), TestLinearSteel(), eps0, kx, ky
            limits.Evaluate section, solver: utilization = limits.Utilization
            builder.ScaleToPhysicalLimit section, utilization, eps0, kx, ky
            solver.EvaluatePhysicalStrainPlane section, TestLinearConcrete(), TestLinearSteel(), eps0, kx, ky
            limits.Evaluate section, solver
            prefix = "limitGuess." & CStr(shift) & "." & CStr(multiplier)
            TestAssertClose stats, prefix & ".ratio", limits.Utilization, 0.9, 0.000000001
            TestAssertClose stats, prefix & ".center", frame.CenterStrain(eps0, kx, ky), _
                -CDbl(multiplier) * 0.004 * 0.9 / utilization, 0.00000000001
            TestAssertClose stats, prefix & ".kx", kx, CDbl(multiplier) * 0.00001 * 0.9 / utilization, 0.000000000001
            TestAssertClose stats, prefix & ".ky", ky, CDbl(multiplier) * -0.000006 * 0.9 / utilization, 0.000000000001
            TestAssertClose stats, prefix & ".coordinates", section.ConcreteX(1), CDbl(shift) - 100#, 0#
        Next multiplier
        eps0 = 0.0002: kx = 0.000001: ky = -0.000003
        builder.ScaleToPhysicalLimit section, 0#, eps0, kx, ky
        TestAssertClose stats, "limitGuess.zero." & CStr(shift), eps0, 0.0002, 0#
    Next shift
    TestCapacityInitialPlaneBounds stats
    TestAppendLine stats, "TOTAL_LIMIT_GUESS_SCALING: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunLimitGuessScalingTests = stats.Report
    Exit Function
Failed:
    RunLimitGuessScalingTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Старт моментного и общего пути не уходит за физические пределы даже при
' большой перегрузке. Готовое НДС сохраняется дословно, без изменения пути.
Private Sub TestCapacityInitialPlaneBounds(ByRef stats As TCapacityTestStats)
    Dim section As CSectionModel, frame As CSectionNumericalFrame, shift As Variant, n As Variant, mode As Variant
    Dim cap As CCapacitySolver, path As CLoadPathVector, concrete As Object, steel As Object
    Dim limits As CSectionStrainLimits, initialState As CSectionSolver, supplied As CSectionSolver
    Dim eps0 As Double, kx As Double, ky As Double, prefix As String
    Set concrete = TestLinearConcrete(): Set steel = TestLinearSteel()
    For Each shift In Array(0#, 1000000#, -1000000#)
        Set section = New CSectionModel
        section.AddConcreteElement CDbl(shift) - 100#, -CDbl(shift) - 100#, 10000#
        section.AddConcreteElement CDbl(shift) + 100#, -CDbl(shift) - 100#, 10000#
        section.AddConcreteElement CDbl(shift) - 100#, -CDbl(shift) + 100#, 10000#
        section.AddConcreteElement CDbl(shift) + 100#, -CDbl(shift) + 100#, 10000#
        section.AddRebarElement CDbl(shift) - 80#, -CDbl(shift) - 80#, 20#, 0#, "A400"
        section.AddRebarElement CDbl(shift) + 80#, -CDbl(shift) + 80#, 20#, 0#, "A400"
        Set frame = New CSectionNumericalFrame: frame.Initialize section
        For Each n In Array(-100000#, -1000000000000#, 1000000000000#)
            For Each mode In Array("Moment", "All")
                Set cap = New CCapacitySolver: TestConfigureCapacity cap
                cap.ConcreteTensionLimitEnabled = True: cap.ConcreteTensionLimit = 0.00015
                cap.SteelCompressionLimit = 0.001: cap.SteelTensionLimit = 0.0025
                Set path = New CLoadPathVector
                If CStr(mode) = "Moment" Then
                    path.Initialize CDbl(n), 0#, CDbl(n) * frame.OriginY, 10000000#, _
                        CDbl(n) * frame.OriginX, -3000000#
                Else
                    path.Initialize 0#, CDbl(n), 0#, CDbl(n) * frame.OriginY + 10000000#, _
                        0#, CDbl(n) * frame.OriginX - 3000000#
                End If
                prefix = "capacityInitialBounds." & CStr(shift) & "." & CStr(n) & "." & CStr(mode)
                TestAssertTrue stats, prefix & ".prepared", cap.LimitSearchPrepareCapacityUltimate(section, concrete, steel, path)
                cap.LimitSearchInitializeUltimateGuess section, concrete, steel, path, Nothing, eps0, kx, ky
                Set initialState = New CSectionSolver
                initialState.EvaluatePhysicalStrainPlane section, concrete, steel, eps0, kx, ky
                Set limits = New CSectionStrainLimits
                limits.Configure -0.0035, 0.00015, 0.001, 0.0025, True
                limits.Evaluate section, initialState
                TestAssertTrue stats, prefix & ".physical", limits.Utilization <= 1.000000001
                TestAssertClose stats, prefix & ".Nunchanged", path.TargetN(1#), CDbl(n), 0#
                Set supplied = New CSectionSolver
                supplied.SetInitialState frame.OriginalStrain(-0.008, 0.000001, -0.000002), 0.000001, -0.000002
                cap.LimitSearchInitializeUltimateGuess section, concrete, steel, path, supplied, eps0, kx, ky
                TestAssertClose stats, prefix & ".suppliedEps", eps0, supplied.Epsilon0, 0#
                TestAssertClose stats, prefix & ".suppliedKx", kx, supplied.KappaX, 0#
                TestAssertClose stats, prefix & ".suppliedKy", ky, supplied.KappaY, 0#
            Next mode
        Next n
    Next shift
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' общий критерий бетона и арматуры без запуска поиска нагрузки.
Public Function RunCommonStrainLimitTests() As String
    Dim stats As TCapacityTestStats
    On Error GoTo Failed
    TestCommonStrainLimits stats
    TestCenteredAxialPathClassification stats
    TestAppendLine stats, "TOTAL_COMMON_STRAIN_LIMITS: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunCommonStrainLimitTests = stats.Report
    Exit Function
Failed:
    RunCommonStrainLimitTests = stats.Report & "FAIL: runtime=" & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Осевая стратегия не зависит от положения чертежа; реальные моменты
' и постоянная часть траектории не должны ошибочно считаться чистым N.
Private Sub TestCenteredAxialPathClassification(ByRef stats As TCapacityTestStats)
    Dim section As CSectionModel, frame As CSectionNumericalFrame, shift As Variant
    Dim n As Double, mx As Double, my As Double, prefix As String
    For Each shift In Array(0#, 1000000#, -1000000#)
        Set section = New CSectionModel
        section.AddConcreteElement CDbl(shift) - 50#, CDbl(shift) - 100#, 100#
        section.AddConcreteElement CDbl(shift) + 50#, CDbl(shift) + 100#, 100#
        Set frame = New CSectionNumericalFrame: frame.Initialize section
        n = -1000000#: mx = n * frame.OriginY: my = n * frame.OriginX
        prefix = "centeredAxial." & CStr(shift)
        Dim math As CLoadPathMath
        Set math = New CLoadPathMath
        TestAssertTrue stats, prefix & ".pureN", math.IsPureAxialLoadPath(0#, n, 0#, mx, 0#, my, frame.OriginX, frame.OriginY)
        TestAssertTrue stats, prefix & ".Mx", Not math.IsPureAxialLoadPath(0#, n, 0#, mx + 1#, 0#, my, frame.OriginX, frame.OriginY)
        TestAssertTrue stats, prefix & ".My", Not math.IsPureAxialLoadPath(0#, n, 0#, mx, 0#, my - 1#, frame.OriginX, frame.OriginY)
        TestAssertTrue stats, prefix & ".constantMoment", Not math.IsPureAxialLoadPath(0#, n, 1#, mx, 0#, my, frame.OriginX, frame.OriginY)
        TestAssertTrue stats, prefix & ".constantN", Not math.IsPureAxialLoadPath(-1#, n, -frame.OriginY, mx, -frame.OriginX, my, frame.OriginX, frame.OriginY)
        TestAssertTrue stats, prefix & ".emptyBase", Not math.IsPureAxialLoadPath(0#, 0#, 0#, 0#, 0#, 0#, frame.OriginX, frame.OriginY)
    Next shift
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Раздельно проверяет четыре ограничения, отключенную растянутую ветвь,
' разные пределы стали и неизменность критерия при переносе координат.
Private Sub TestCommonStrainLimits(ByRef stats As TCapacityTestStats)
    Dim section As CSectionModel, solver As CSectionSolver, limits As CSectionStrainLimits
    Dim shift As Variant, i As Long, j As Long, state As String, expected As Double
    Dim strain As Double, utilization As Double, criticalStrain As Double, element As String
    Dim enabled As Variant, plane As Variant, planes As Variant, prefix As String
    Set limits = New CSectionStrainLimits
    planes = Array(Array(0#, 0#, 0#), Array(-0.001, 0#, 0#), Array(0.0001, 0#, 0#), _
        Array(-0.0002, 0.000007, -0.000002), Array(0.0001, -0.000004, 0.000006))
    For Each shift In Array(0#, 1000000#, -1000000#)
        Set section = New CSectionModel
        section.AddConcreteElement CDbl(shift) - 100#, CDbl(shift) / 2# - 100#, 10000#
        section.AddConcreteElement CDbl(shift) + 100#, CDbl(shift) / 2# - 100#, 10000#
        section.AddConcreteElement CDbl(shift) - 100#, CDbl(shift) / 2# + 100#, 10000#
        section.AddConcreteElement CDbl(shift) + 100#, CDbl(shift) / 2# + 100#, 10000#
        section.AddRebarElement CDbl(shift) - 80#, CDbl(shift) / 2# - 80#, 20#, 0#, "A400"
        section.AddRebarElement CDbl(shift) + 80#, CDbl(shift) / 2# + 80#, 20#, 0#, "A400"
        Set solver = New CSectionSolver
        For Each enabled In Array(False, True)
            j = 0
            For Each plane In planes
                j = j + 1
                solver.EvaluatePhysicalCenterPlane section, TestLinearConcrete(), TestLinearSteel(), _
                    CDbl(plane(0)), CDbl(plane(1)), CDbl(plane(2))
                limits.Configure -0.0035, 0.00015, 0.001, 0.0025, CBool(enabled)
                limits.Evaluate section, solver
                utilization = 0#: criticalStrain = 0#: element = vbNullString: state = "NumericalFailure"
                ' Независимый тестовый oracle: прямой перебор по известным пределам.
                For i = 1 To section.ConcreteCount
                    strain = solver.StrainAt(section.ConcreteX(i), section.ConcreteY(i))
                    expected = 0#
                    If strain < 0# Then expected = strain / -0.0035
                    If strain > 0# And CBool(enabled) Then expected = strain / 0.00015
                    If expected > utilization Then
                        utilization = expected: criticalStrain = strain
                        element = section.ConcreteID(i)
                        If strain < 0# Then
                            element = element & ":Compression": state = "ConcreteStrainLimit"
                        Else
                            element = element & ":Tension": state = "ConcreteTensionStrainLimit"
                        End If
                    End If
                Next i
                For i = 1 To section.RebarCount
                    strain = solver.StrainAt(section.RebarX(i), section.RebarY(i))
                    If strain < 0# Then expected = Abs(strain) / 0.001 Else expected = strain / 0.0025
                    If expected > utilization Then
                        utilization = expected: criticalStrain = strain
                        element = section.RebarID(i) & ":Steel": state = "SteelStrainLimit"
                    End If
                Next i
                prefix = "commonLimits." & CStr(shift) & "." & CStr(enabled) & "." & CStr(j)
                TestAssertClose stats, prefix & ".ratio", limits.Utilization, utilization, 0#
                TestAssertClose stats, prefix & ".strain", limits.CriticalStrain, criticalStrain, 0#
                TestAssertEquals stats, prefix & ".element", limits.CriticalElement, element
                TestAssertEquals stats, prefix & ".state", limits.LimitState, state
            Next plane
        Next enabled
        TestAssertCommonLimitCase stats, "compressionConcrete." & CStr(shift), section, -0.0035, _
            -0.0035, 0.00015, 0.025, 0.025, True, "ConcreteStrainLimit", "сжатие бетона"
        TestAssertCommonLimitCase stats, "tensionConcrete." & CStr(shift), section, 0.00015, _
            -0.0035, 0.00015, 0.025, 0.025, True, "ConcreteTensionStrainLimit", "растяжение бетона"
        TestAssertCommonLimitCase stats, "compressionSteel." & CStr(shift), section, -0.001, _
            -0.0035, 0.00015, 0.001, 0.0025, True, "SteelStrainLimit", "сжатие арматуры"
        TestAssertCommonLimitCase stats, "tensionSteel." & CStr(shift), section, 0.0025, _
            -0.0035, 0.00015, 0.001, 0.0025, False, "SteelStrainLimit", "растяжение арматуры"
        TestAssertCommonLimitCase stats, "tieConcreteFirst." & CStr(shift), section, -0.0035, _
            -0.0035, 0.00015, 0.0035, 0.025, True, "ConcreteStrainLimit", "сжатие бетона"
    Next shift
    TestStrainExtremaOwnership stats
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет принадлежность кэша экстремальных деформаций модели и её ревизии, выбор первого элемента при равенстве и полный поиск при потере точности.
Private Sub TestStrainExtremaOwnership(ByRef stats As TCapacityTestStats)
    Dim section As CSectionModel, foreign As CSectionModel, solver As CSectionSolver
    Dim limits As CSectionStrainLimits, cMin As Double, cMax As Double, sMin As Double, sMax As Double
    Dim cMinIndex As Long, cMaxIndex As Long, sMinIndex As Long, sMaxIndex As Long
    Set section = New CSectionModel: Set solver = New CSectionSolver
    section.AddConcreteElement 100#, 0#, 10000#
    section.AddConcreteElement -100#, 0#, 10000#
    section.AddRebarElement 80#, 0#, 20#, 0#, "A400"
    section.AddRebarElement -80#, 0#, 20#, 0#, "A400"
    solver.EvaluatePhysicalCenterPlane section, TestLinearConcrete(), TestLinearSteel(), 0#, 0#, 0.00001
    TestAssertTrue stats, "strainExtrema.current", solver.TryGetStrainExtrema(section, cMin, cMax, _
        cMinIndex, cMaxIndex, sMin, sMax, sMinIndex, sMaxIndex)
    TestAssertClose stats, "strainExtrema.concreteMinIndex", CDbl(cMinIndex), 2#, 0#
    TestAssertClose stats, "strainExtrema.concreteMaxIndex", CDbl(cMaxIndex), 1#, 0#
    TestAssertTrue stats, "strainBounds.current", solver.TryGetConcreteStrainBounds(section, cMin, cMax)
    TestAssertClose stats, "strainBounds.minimum", cMin, solver.StrainAt(-100#, 0#), 0#
    TestAssertClose stats, "strainBounds.maximum", cMax, solver.StrainAt(100#, 0#), 0#
    Set limits = New CSectionStrainLimits
    limits.Configure -0.001, 0.001, 0.025, 0.025, True
    limits.Evaluate section, solver
    TestAssertEquals stats, "strainExtrema.tie.firstConcrete", limits.CriticalElement, section.ConcreteID(1) & ":Tension"
    limits.Configure -0.0035, 0.00015, 0.0008, 0.0008, False
    limits.Evaluate section, solver
    TestAssertEquals stats, "strainExtrema.tie.firstSteel", limits.CriticalElement, section.RebarID(1) & ":Steel"
    TestAssertEquals stats, "strainExtrema.tie.firstSteelDirection", limits.Description, "растяжение арматуры"

    Set foreign = New CSectionModel
    foreign.AddConcreteElement 100#, 0#, 10000#
    foreign.AddConcreteElement -100#, 0#, 10000#
    TestAssertTrue stats, "strainExtrema.foreign", Not solver.TryGetStrainExtrema(foreign, cMin, cMax, _
        cMinIndex, cMaxIndex, sMin, sMax, sMinIndex, sMaxIndex)
    TestAssertTrue stats, "strainBounds.foreign", Not solver.TryGetConcreteStrainBounds(foreign, cMin, cMax)
    section.AddRebarElement 200#, 0#, 20#, 0#, "A400"
    TestAssertTrue stats, "strainExtrema.revision", Not solver.TryGetStrainExtrema(section, cMin, cMax, _
        cMinIndex, cMaxIndex, sMin, sMax, sMinIndex, sMaxIndex)
    TestAssertTrue stats, "strainBounds.revision", Not solver.TryGetConcreteStrainBounds(section, cMin, cMax)
    limits.Evaluate section, solver
    TestAssertEquals stats, "strainExtrema.revision.fullScan", limits.CriticalElement, section.RebarID(3) & ":Steel"
    TestAssertClose stats, "strainExtrema.revision.ratio", limits.Utilization, 2.5, 0#
    solver.ClearResult
    TestAssertTrue stats, "strainExtrema.cleared", Not solver.TryGetStrainExtrema(section, cMin, cMax, _
        cMinIndex, cMaxIndex, sMin, sMax, sMinIndex, sMaxIndex)
    TestAssertTrue stats, "strainBounds.cleared", Not solver.TryGetConcreteStrainBounds(section, cMin, cMax)
    solver.EvaluatePhysicalCenterPlane section, TestLinearConcrete(), TestLinearSteel(), 0.001, 0#, 5E-21 ' Вклад кривизны теряется рядом с epsilon0: кэш значений допустим, но индексы экстремумов требуют полного поиска.
    TestAssertTrue stats, "strainExtrema.rounding.fullScan", Not solver.TryGetStrainExtrema(section, cMin, cMax, _
        cMinIndex, cMaxIndex, sMin, sMax, sMinIndex, sMaxIndex)
    TestAssertTrue stats, "strainBounds.rounding.valuesAvailable", solver.TryGetConcreteStrainBounds(section, cMin, cMax)
    TestAssertClose stats, "strainBounds.rounding.minimum", cMin, solver.MinConcreteStrain, 0#
    TestAssertClose stats, "strainBounds.rounding.maximum", cMax, solver.MaxConcreteStrain, 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет заданный деформационный предел на однородной плоскости: коэффициент использования, вид предела и понятное описание.
Private Sub TestAssertCommonLimitCase(ByRef stats As TCapacityTestStats, ByVal name As String, _
        ByVal section As CSectionModel, ByVal strain As Double, ByVal concreteCompression As Double, _
        ByVal concreteTension As Double, ByVal steelCompression As Double, ByVal steelTension As Double, _
        ByVal tensionEnabled As Boolean, ByVal state As String, ByVal description As String)
    Dim solver As CSectionSolver, limits As CSectionStrainLimits
    Set solver = New CSectionSolver: Set limits = New CSectionStrainLimits
    solver.EvaluatePhysicalCenterPlane section, TestLinearConcrete(), TestLinearSteel(), strain, 0#, 0#
    limits.Configure concreteCompression, concreteTension, steelCompression, steelTension, tensionEnabled
    limits.Evaluate section, solver
    TestAssertClose stats, name & ".ratio", limits.Utilization, 1#, 0.000000000001
    TestAssertEquals stats, name & ".state", limits.LimitState, state
    TestAssertEquals stats, name & ".description", limits.Description, description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет реальные COM-интерфейсы после compile/save/reopen.
Public Function RunLimitSearchInterfaceRegressionTests() As String
    Dim stats As TCapacityTestStats
    On Error GoTo Failed
    TestLimitSearchAdapterInterfaces stats
    TestAppendLine stats, "TOTAL_LIMIT_INTERFACES: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunLimitSearchInterfaceRegressionTests = stats.Report
    Exit Function
Failed:
    RunLimitSearchInterfaceRegressionTests = stats.Report & "FAIL: runtime=" & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Интерфейс и его запрос не должны зависеть от реализующих интерфейс классов.
' Проверяем обе реализации и передачу готового критерия через общую фабрику.
Private Sub TestLimitSearchAdapterInterfaces(ByRef stats As TCapacityTestStats)
    Dim capacity As CCapacitySolver, formation As CCrackFormationCalculator
    Dim capacityProblem As CCapacityLimitSearchProblem, formationProblem As CCrackLimitSearchProblem
    Dim callback As ILimitSearchProblem, section As CSectionModel, request As CLimitSearchRequest
    Set capacity = New CCapacitySolver
    Set formation = New CCrackFormationCalculator
    Set capacityProblem = New CCapacityLimitSearchProblem
    Set formationProblem = New CCrackLimitSearchProblem
    capacityProblem.Initialize capacity
    formationProblem.Initialize formation
    TestAssertTrue stats, "limitInterface.capacity.supports", TypeOf capacityProblem Is ILimitSearchProblem
    TestAssertTrue stats, "limitInterface.formation.supports", TypeOf formationProblem Is ILimitSearchProblem
    If Not TypeOf capacityProblem Is ILimitSearchProblem Then Exit Sub
    If Not TypeOf formationProblem Is ILimitSearchProblem Then Exit Sub
    Set callback = capacityProblem
    TestAssertTrue stats, "limitInterface.capacity.cast", callback.ResultKind = rkCapacity
    Set callback = formationProblem
    TestAssertTrue stats, "limitInterface.formation.cast", callback.ResultKind = rkCrackFormation
    Set section = New CSectionModel
    section.AddConcreteElement 0#, 0#, 10000#
    Set request = New CLimitSearchRequest
    InitializeCapacitySearchRequest request, capacity, section, TestLinearConcrete(), TestLinearSteel(), _
        "LoadMultiplier", 0#, -100#, 0#, 20#, 0#, 30#, False, False
    TestAssertTrue stats, "limitInterface.capacity.requestKind", request.Problem.ResultKind = rkCapacity
    TestAssertTrue stats, "limitInterface.capacity.requestSection", request.Section Is section
    Set request = New CLimitSearchRequest
    InitializeCrackFormationSearchRequest request, formation, section, TestLinearConcrete(), TestLinearSteel(), _
        "Auto", "LambdaMxy", -100#, 20#, 30#, 20#, 30#, 0#, 0#
    TestAssertTrue stats, "limitInterface.formation.requestKind", request.Problem.ResultKind = rkCrackFormation
    TestAssertTrue stats, "limitInterface.formation.requestSection", request.Section Is section
    TestAssertClose stats, "limitInterface.formation.targetMx", request.TargetMx, 20#, 0#
    TestAssertEquals stats, "limitInterface.formation.path", request.CrackFormationPath, "LambdaMxy"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строгий абсолютный допуск не должен ослабляться скрытым относительным минимумом.
Private Sub TestUltimateStrainToleranceNoHiddenFloor(ByRef stats As TCapacityTestStats)
    Dim section As CSectionModel, cap As CCapacitySolver, path As CLoadPathVector
    Dim solver As CSectionSolver, r1 As Double, r2 As Double, rLimit As Double
    Dim lambdaValue As Double, state As String
    Set section = New CSectionModel
    section.AddConcreteElement 0#, 0#, 10000#
    Set cap = New CCapacitySolver
    cap.ConcreteCompressionLimit = -0.0035
    cap.StrainTolerance = 0.000000001
    Set path = New CLoadPathVector
    path.Initialize 0#, -100000#, 0#, 0#, 0#, 0#
    TestAssertTrue stats, "ultimateTolerance.physicalCriterion", _
        cap.LimitSearchEvaluateUltimateResidual(section, TestProvisionalConcrete(), TestProvisionalSteel(), _
            path, -0.0035, 0#, 0#, solver, r1, r2, rLimit, lambdaValue, state)
    TestAssertTrue stats, "ultimateTolerance.rejectAboveAbsolute", _
        Not cap.LimitSearchUltimateResidualReached(0#, 0#, 0.0000005)
    TestAssertTrue stats, "ultimateTolerance.acceptBelowAbsolute", _
        cap.LimitSearchUltimateResidualReached(0#, 0#, 0.00000025)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Смена изменяемого пути, ревизии геометрии или допуска не переиспользует
' прежний численный контекст; отдельная оценка совпадает с новым calculator-ом.
Private Sub TestUltimateNumericalContextCache(ByRef stats As TCapacityTestStats)
    Dim section As CSectionModel, cap As CCapacitySolver, fresh As CCapacitySolver, path As CLoadPathVector
    Dim solver As CSectionSolver, reference As CSectionSolver, phase As Long, prefix As String
    Dim r1 As Double, r2 As Double, r3 As Double, lambda As Double, state As String
    Dim s1 As Double, s2 As Double, s3 As Double, expected As Double, referenceState As String
    Set section = New CSectionModel
    section.AddConcreteElement -25#, -20#, 10000#
    section.AddConcreteElement 25#, 20#, 10000#
    Set cap = New CCapacitySolver: Set path = New CLoadPathVector
    path.Initialize 0#, -100000#, 100#, 30#, -200#, 20#
    For phase = 1 To 4
        If phase = 2 Then path.Initialize 700#, -200000#, -500#, 60#, 300#, -40#
        If phase = 3 Then section.AddConcreteElement 70#, -30#, 5000#
        If phase = 4 Then
            cap.SolverToleranceN = 0.001: cap.SolverToleranceMx = 0.01: cap.SolverToleranceMy = 0.02
        End If
        Set fresh = New CCapacitySolver
        If phase = 4 Then
            fresh.SolverToleranceN = 0.001: fresh.SolverToleranceMx = 0.01: fresh.SolverToleranceMy = 0.02
        End If
        prefix = "ultimateContext." & CStr(phase)
        TestAssertTrue stats, prefix & ".warm", cap.LimitSearchEvaluateUltimateResidual(section, _
            TestProvisionalConcrete(), TestProvisionalSteel(), path, -0.003, 0.000003, -0.000002, solver, r1, r2, r3, lambda, state)
        TestAssertTrue stats, prefix & ".cold", fresh.LimitSearchEvaluateUltimateResidual(section, _
            TestProvisionalConcrete(), TestProvisionalSteel(), path, -0.003, 0.000003, -0.000002, reference, s1, s2, s3, expected, referenceState)
        TestAssertClose stats, prefix & ".lambda", lambda, expected, 0#
        TestAssertClose stats, prefix & ".path1", r1, s1, 0#
        TestAssertClose stats, prefix & ".path2", r2, s2, 0#
        TestAssertClose stats, prefix & ".limit", r3, s3, 0#
        TestAssertEquals stats, prefix & ".state", state, referenceState
        TestAssertTrue stats, prefix & ".tolerance", _
            cap.LimitSearchUltimateResidualReached(0.00000001, 0#, 0#) = _
            fresh.LimitSearchUltimateResidualReached(0.00000001, 0#, 0#)
    Next phase
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' общая подготовка сверяется с независимым явным переносом,
' включая изменение каждого поля пути, модели, допуска и способа нормировки.
Public Function RunNumericalPathPreparationTests() As String
    Dim stats As TCapacityTestStats
    On Error GoTo Failed
    TestNumericalPathPreparation stats
    TestAppendLine stats, "TOTAL_NUMERICAL_PATH_PREPARATION: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunNumericalPathPreparationTests = stats.Report
    Exit Function
Failed:
    RunNumericalPathPreparationTests = stats.Report & "FAIL: numericalPath.runtime=" & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает подготовленный численный путь с заново вычисленным переносом и масштабированием после изменения каждого значимого входа.
Private Sub TestNumericalPathPreparation(ByRef stats As TCapacityTestStats)
    Dim section As CSectionModel, other As CSectionModel, path As CLoadPathVector, transferred As CLoadPathVector
    Dim cached As CLoadPathMath, oracle As CLoadPathMath, frame As CSectionNumericalFrame, solver As CSectionSolver
    Dim mode As Variant, phase As Long, i As Long, commonScale As Boolean, values(1 To 6) As Double
    Dim tn As Double, tx As Double, ty As Double, n As Double, mx As Double, my As Double, forceScale As Double
    Dim actualLambda As Double, actual1 As Double, actual2 As Double, expectedLambda As Double, expected1 As Double, expected2 As Double
    Dim prefix As String, expectedTolerance As Double
    For Each mode In Array(False, True)
        Set section = New CSectionModel: Set other = New CSectionModel
        section.AddConcreteElement 1000100#, -999900#, 2500#
        section.AddConcreteElement 999900#, -1000100#, 2500#
        section.AddConcreteElement 1000100#, -1000100#, 2500#
        other.AddConcreteElement -2000200#, 1999900#, 2300#
        other.AddConcreteElement -1999900#, 2000100#, 2700#
        other.AddConcreteElement -2000100#, 2000150#, 2600#
        section.AddRebarElement 1000000#, -1000000#, 22#, 0#, "A400"
        other.AddRebarElement -2000000#, 2000000#, 28#, 0#, "A400"
        values(1) = -2000#: values(2) = -20000#: values(3) = 2000000000#: values(4) = 20000000000#
        values(5) = -2000000000#: values(6) = -20000000000#
        tn = 0.001: tx = 200#: ty = 150#: commonScale = CBool(mode)
        Set path = New CLoadPathVector: Set cached = New CLoadPathMath
        For phase = 1 To 18
            Select Case phase
                Case 3: section.AddConcreteElement 1000000#, -1000075#, 1700#
                Case 4: Set section = other
                Case 5 To 10: values(phase - 4) = values(phase - 4) + 123.456 ' По очереди меняется каждая из шести компонент: смещение и база N, Mx, My.
                Case 11: tn = 0.007
                Case 12: tx = 345#
                Case 13: ty = 567#
                Case 14, 15: commonScale = Not commonScale
                Case 16: cached.Configure 77#, 88#, 99#
                Case 17: values(2) = 1E-20
            End Select
            path.Initialize values(1), values(2), values(3), values(4), values(5), values(6)
            Set frame = New CSectionNumericalFrame: frame.Initialize section
            Set transferred = New CLoadPathVector
            transferred.Initialize values(1), values(2), _
                frame.CenterMx(values(1), values(3)) / frame.LengthScale, frame.CenterMx(values(2), values(4)) / frame.LengthScale, _
                frame.CenterMy(values(1), values(5)) / frame.LengthScale, frame.CenterMy(values(2), values(6)) / frame.LengthScale
            forceScale = 1#
            For i = 1 To 6
                Select Case i
                    Case 1: n = transferred.NOffset
                    Case 2: n = transferred.NBase
                    Case 3: n = transferred.MxOffset
                    Case 4: n = transferred.MxBase
                    Case 5: n = transferred.MyOffset
                    Case 6: n = transferred.MyBase
                End Select
                If Abs(n) > forceScale Then forceScale = Abs(n)
            Next i
            Set oracle = New CLoadPathMath
            If commonScale Then oracle.Configure forceScale, forceScale, forceScale Else oracle.Configure tn, tx / frame.LengthScale, ty / frame.LengthScale
            Set solver = New CSectionSolver
            solver.EvaluatePhysicalCenterPlane section, TestProvisionalConcrete(), TestProvisionalSteel(), -0.0006, 0.000001, -0.0000007
            solver.NumericalEfforts n, mx, my
            mx = mx / frame.LengthScale: my = my / frame.LengthScale
            expectedLambda = oracle.Lambda(n, mx, my, transferred)
            If commonScale Then
                oracle.BuildResiduals n, mx, my, transferred, expectedLambda, expected1, expected2, forceScale
            Else
                oracle.BuildResiduals n, mx, my, transferred, expectedLambda, expected1, expected2
            End If
            cached.PrepareNumericalPath section, path, tn, tx, ty, commonScale
            cached.BuildNumericalResiduals solver, actualLambda, actual1, actual2
            prefix = "numericalPreparation." & CStr(mode) & "." & CStr(phase)
            TestAssertClose stats, prefix & ".lambda", actualLambda, expectedLambda, 0#
            TestAssertClose stats, prefix & ".residual1", actual1, expected1, 0#
            TestAssertClose stats, prefix & ".residual2", actual2, expected2, 0#
            TestAssertClose stats, prefix & ".lambdaOnly", cached.NumericalLambda(solver), expectedLambda, 0#
            TestAssertClose stats, prefix & ".forceScale", cached.NumericalForceScale, forceScale, 0#
            expectedTolerance = tn
            If tx / frame.LengthScale < expectedTolerance Then expectedTolerance = tx / frame.LengthScale
            If ty / frame.LengthScale < expectedTolerance Then expectedTolerance = ty / frame.LengthScale
            TestAssertClose stats, prefix & ".pathTolerance", cached.NumericalPathTolerance, _
                expectedTolerance / forceScale, 0#
            TestAssertClose stats, prefix & ".sourceNOffset", path.NOffset, values(1), 0#
            TestAssertClose stats, prefix & ".sourceNBase", path.NBase, values(2), 0#
            TestAssertClose stats, prefix & ".sourceMxOffset", path.MxOffset, values(3), 0#
            TestAssertClose stats, prefix & ".sourceMxBase", path.MxBase, values(4), 0#
            TestAssertClose stats, prefix & ".sourceMyOffset", path.MyOffset, values(5), 0#
            TestAssertClose stats, prefix & ".sourceMyBase", path.MyBase, values(6), 0#
        Next phase
    Next mode
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' направленные проверки осевого предела и соседних критериев.
Public Function RunAxialCapacityRegressionTests() As String
    Dim stats As TCapacityTestStats
    On Error GoTo Failed
    TestAxialConcretePlateauStart stats
    TestAxialLoadMultiplierFindsNult stats
    TestNultBaseLoadStepsSensitivity stats
    TestConcreteLimitState stats
    TestConcreteTensionLimitState stats
    TestSteelLimitState stats
    TestAudit02AsymmetricSteelLimits stats
    TestAppendLine stats, "TOTAL_AXIAL_CAPACITY: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAxialCapacityRegressionTests = stats.Report
    Exit Function
Failed:
    RunAxialCapacityRegressionTests = stats.Report & "FAIL: runtime=" & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' все 24 теплых старта на плато, бюджет возврата к общему старту
' и шесть штатных поисков известного нелинейного предела с переносами.
Public Function RunProjectedLimitRecoveryTests() As String
    Dim stats As TCapacityTestStats
    On Error GoTo Failed
    TestProjectedLimitRecovery stats
    TestWarmStartRestartContract stats
    TestAppendLine stats, "TOTAL_PROJECTED_RECOVERY: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunProjectedLimitRecoveryTests = stats.Report
    Exit Function
Failed:
    RunProjectedLimitRecoveryTests = stats.Report & "FAIL: projectedRecovery.runtime=" & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' масштаб пути не меняет известный предел и его равновесие.
Public Function RunUltimateStepScaleRegressionTests() As String
    Dim stats As TCapacityTestStats
    On Error GoTo Failed
    TestUltimateStepScale stats
    TestAppendLine stats, "TOTAL_ULTIMATE_STEP_SCALE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunUltimateStepScaleRegressionTests = stats.Report
    Exit Function
Failed:
    RunUltimateStepScaleRegressionTests = stats.Report & "FAIL: ultimateStepScale.runtime=" & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Одна физическая плоскость задает широкий диапазон lambda при тех же усилиях.
' Проверяются холодный старт, плато и исходные координаты удаленного чертежа.
Private Sub TestUltimateStepScale(ByRef stats As TCapacityTestStats)
    Dim shift As Variant, sign As Variant, factor As Variant, x As Long, y As Long, startIndex As Long
    Dim section As CSectionModel, concrete As CMaterialDiagram, steel As CMaterialDiagram
    Dim known As CSectionSolver, start As CSectionSolver, cap As CCapacitySolver
    Dim adapter As CCapacityLimitSearchProblem, problem As ILimitSearchProblem, search As CUltimateStrainSearch
    Dim path As CLoadPathVector, prefix As String, succeeded As Boolean, expectedLambda As Double
    For Each shift In Array(0#, -1000000#, 1000000#)
        Set section = New CSectionModel
        For x = -4 To 4
            For y = -2 To 2
                section.AddConcreteElement CDbl(shift) + 50# * x, -CDbl(shift) + 50# * y, 2500#
            Next y
        Next x
        section.AddRebarElement CDbl(shift) - 160#, -CDbl(shift) - 80#, 32#, 0#, "A400"
        section.AddRebarElement CDbl(shift) + 160#, -CDbl(shift) - 80#, 32#, 0#, "A400"
        section.AddRebarElement CDbl(shift) - 160#, -CDbl(shift) + 80#, 32#, 0#, "A400"
        section.AddRebarElement CDbl(shift) + 160#, -CDbl(shift) + 80#, 32#, 0#, "A400"
        Set concrete = TestProvisionalConcrete(): Set steel = TestProvisionalSteel()
        concrete.Freeze: steel.Freeze
        For Each sign In Array(-1#, 1#)
            Set known = New CSectionSolver
            known.EvaluatePhysicalCenterPlane section, concrete, steel, -0.0021, CDbl(sign) * 0.000006, CDbl(sign) * 0.000004
            For Each factor In Array(0.000001, 0.01, 1#, 10000#, 1000000#)
                expectedLambda = 1# / CDbl(factor)
                Set path = New CLoadPathVector
                path.Initialize 0#, known.Nint * CDbl(factor), 0#, known.Mxint * CDbl(factor), 0#, known.Myint * CDbl(factor)
                For startIndex = 1 To 2
                    Set cap = New CCapacitySolver
                    cap.ConcreteCompressionLimit = -0.0035: cap.ConcreteTensionLimit = 0#
                    cap.SteelCompressionLimit = 0.025: cap.SteelTensionLimit = 0.025
                    cap.SolverToleranceN = 0.001: cap.SolverToleranceMx = 200#: cap.SolverToleranceMy = 200#
                    cap.SolverMaxIterations = 60: cap.StrainTolerance = 0.000001
                    prefix = "ultimateStepScale." & CStr(shift) & "." & CStr(sign) & "." & CStr(factor) & "." & CStr(startIndex)
                    Set start = Nothing
                    If startIndex = 2 Then
                        Set start = New CSectionSolver
                        start.EvaluatePhysicalCenterPlane section, concrete, steel, -0.003, 0#, 0#
                        TestAssertTrue stats, prefix & ".startEquilibrium", start.ConfirmEquilibrium(start.Nint, start.Mxint, start.Myint)
                    End If
                    Set adapter = New CCapacityLimitSearchProblem: adapter.Initialize cap: Set problem = adapter
                    TestAssertTrue stats, prefix & ".prepared", cap.LimitSearchPrepareCapacityUltimate(section, concrete, steel, path)
                    Set search = New CUltimateStrainSearch
                    succeeded = search.RunNewton(problem, section, concrete, steel, path, start, "Не найден известный предел масштабированного пути.")
                    cap.LimitSearchCompleteCapacityUltimate path
                    TestAssertTrue stats, prefix & ".knownLimitFound." & cap.StopReason, succeeded And cap.Converged
                    If succeeded Then
                        TestAssertClose stats, prefix & ".lambda", cap.LambdaUltimate, expectedLambda, expectedLambda * 0.0001
                        TestAssertEquals stats, prefix & ".criterion", cap.LimitState, "ConcreteStrainLimit"
                        TestAssertTrue stats, prefix & ".equilibrium", cap.LastSolver.ConfirmEquilibrium(path.TargetN(cap.LambdaUltimate), _
                            path.TargetMx(cap.LambdaUltimate), path.TargetMy(cap.LambdaUltimate))
                        TestAssertClose stats, prefix & ".physicalStrain", cap.LastSolver.MinConcreteStrain, -0.0035, 0.000001
                    End If
                Next startIndex
            Next factor
        Next sign
    Next shift
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет восстановление заранее известного предела из разных начальных плоскостей при двух знаках изгиба и больших переносах координат.
Private Sub TestProjectedLimitRecovery(ByRef stats As TCapacityTestStats)
    Dim shift As Variant, sign As Variant, x As Long, y As Long, startIndex As Long
    Dim section As CSectionModel, concrete As CMaterialDiagram, steel As CMaterialDiagram
    Dim known As CSectionSolver, start As CSectionSolver, cap As CCapacitySolver
    Dim adapter As CCapacityLimitSearchProblem, problem As ILimitSearchProblem, search As CUltimateStrainSearch
    Dim path As CLoadPathVector, epsilon As Double, kx As Double, ky As Double, prefix As String, succeeded As Boolean
    For Each shift In Array(0#, -1000000#, 1000000#)
        Set section = New CSectionModel
        For x = -4 To 4
            For y = -2 To 2
                section.AddConcreteElement CDbl(shift) + 50# * x, -CDbl(shift) + 50# * y, 2500#
            Next y
        Next x
        section.AddRebarElement CDbl(shift) - 160#, -CDbl(shift) - 80#, 32#, 0#, "A400"
        section.AddRebarElement CDbl(shift) + 160#, -CDbl(shift) - 80#, 32#, 0#, "A400"
        section.AddRebarElement CDbl(shift) - 160#, -CDbl(shift) + 80#, 32#, 0#, "A400"
        section.AddRebarElement CDbl(shift) + 160#, -CDbl(shift) + 80#, 32#, 0#, "A400"
        Set concrete = TestProvisionalConcrete(): Set steel = TestProvisionalSteel()
        concrete.Freeze: steel.Freeze
        For Each sign In Array(-1#, 1#)
            Set known = New CSectionSolver
            known.EvaluatePhysicalCenterPlane section, concrete, steel, -0.0021, CDbl(sign) * 0.000006, CDbl(sign) * 0.000004
            Set path = New CLoadPathVector: path.Initialize 0#, known.Nint, 0#, known.Mxint, 0#, known.Myint
            For startIndex = 1 To 5
                Set cap = New CCapacitySolver
                cap.ConcreteCompressionLimit = -0.0035: cap.ConcreteTensionLimit = 0#
                cap.SteelCompressionLimit = 0.025: cap.SteelTensionLimit = 0.025
                cap.SolverToleranceN = 0.001: cap.SolverToleranceMx = 200#: cap.SolverToleranceMy = 200#
                cap.SolverMaxIterations = 60: cap.StrainTolerance = 0.000001
                prefix = "projectedRecovery." & CStr(shift) & "." & CStr(sign) & "." & CStr(startIndex)
                If startIndex = 5 Then
                    prefix = prefix & ".public"
                    cap.SolveByUltimateLoadPath section, concrete, steel, 0#, path.NBase, 0#, path.MxBase, 0#, path.MyBase
                    succeeded = cap.Converged
                Else
                    Select Case startIndex
                        Case 1: epsilon = -0.003: kx = 0#: ky = 0#
                        Case 2: epsilon = -0.00176: kx = 0.0000004: ky = 0.0000003
                        Case 3: epsilon = -0.0032: kx = 0.000001: ky = -0.0000008
                        Case 4: epsilon = 0.0001: kx = 0#: ky = 0#
                    End Select
                    Set start = New CSectionSolver
                    start.EvaluatePhysicalCenterPlane section, concrete, steel, epsilon, kx, ky
                    TestAssertTrue stats, prefix & ".startEquilibrium", start.ConfirmEquilibrium(start.Nint, start.Mxint, start.Myint)
                    Set adapter = New CCapacityLimitSearchProblem: adapter.Initialize cap: Set problem = adapter
                    TestAssertTrue stats, prefix & ".prepared", cap.LimitSearchPrepareCapacityUltimate(section, concrete, steel, path)
                    Set search = New CUltimateStrainSearch
                    succeeded = search.RunNewton(problem, section, concrete, steel, path, start, "Не найден известный предел.")
                    cap.LimitSearchCompleteCapacityUltimate path
                End If
                TestAssertTrue stats, prefix & ".knownLimitFound", succeeded And cap.Converged
                If succeeded Then
                    TestAssertClose stats, prefix & ".lambda", cap.LambdaUltimate, 1#, 0.0001
                    TestAssertTrue stats, prefix & ".N", Abs(cap.LastSolver.Nint - path.TargetN(cap.LambdaUltimate)) <= 0.001
                    TestAssertTrue stats, prefix & ".Mx", Abs(cap.LastSolver.Mxint - path.TargetMx(cap.LambdaUltimate)) <= 200#
                    TestAssertTrue stats, prefix & ".My", Abs(cap.LastSolver.Myint - path.TargetMy(cap.LambdaUltimate)) <= 200#
                End If
            Next startIndex
        Next sign
    Next shift
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' один возврат не сбрасывает бюджет и не скрывает терминальные ошибки.
Private Sub TestWarmStartRestartContract(ByRef stats As TCapacityTestStats)
    Dim kind As Variant, scenario As Long, fake As CTestLimitSearchProblem
    Dim section As CSectionModel, start As CSectionSolver, path As CLoadPathVector
    Dim search As CUltimateStrainSearch, succeeded As Boolean, prefix As String, budget As Long
    Set section = New CSectionModel
    section.AddConcreteElement -10#, -10#, 100#: section.AddConcreteElement 10#, 10#, 100#
    Set path = New CLoadPathVector: path.Initialize 0#, 1#, 0#, 0#, 0#, 0#
    For Each kind In Array(rkCapacity, rkCrackFormation)
        For scenario = 1 To 7
            Set fake = New CTestLimitSearchProblem: Set search = New CUltimateStrainSearch
            Set start = New CSectionSolver
            start.EvaluatePhysicalStrainPlane section, TestLinearConcrete(), TestLinearSteel(), -0.001, 0#, 0#
            If scenario <> 3 Then TestAssertTrue stats, "warmContract.confirmStart", start.ConfirmEquilibrium(start.Nint, start.Mxint, start.Myint)
            budget = 2: If scenario = 2 Then budget = 1
            fake.TestConfigureUltimateWarmCase budget, CLng(kind)
            Select Case scenario
                Case 4
                    Set start = Nothing
                Case 5
                    fake.TestConfigureUltimateWarmCase budget, CLng(kind), coldStrain:=-0.001
                Case 6
                    start.EvaluatePhysicalStrainPlane section, TestLinearConcrete(), TestLinearSteel(), 0#, 0#, 0#
                    TestAssertTrue stats, "warmContract.confirmTerminalStart", start.ConfirmEquilibrium(start.Nint, start.Mxint, start.Myint)
                    fake.TestConfigureUltimateWarmCase budget, CLng(kind), 7
                Case 7
                    fake.TestConfigureUltimateWarmCase budget, CLng(kind), 1
            End Select
            ResetSectionEquilibriumSolveCount
            succeeded = search.RunNewton(fake, Nothing, Nothing, Nothing, path, start, "Тестовый бюджет исчерпан.")
            prefix = "warmContract." & CStr(kind) & "." & CStr(scenario)
            Dim expectedStarts As Long
            expectedStarts = 1
            If scenario = 1 Or scenario = 7 Then expectedStarts = 2
            TestAssertTrue stats, prefix & ".actualStartCount", UltimateStrainStartCount() = expectedStarts
            TestAssertTrue stats, prefix & ".iterationCount", UltimateStrainIterationCount() >= 1 And UltimateStrainIterationCount() <= budget
            TestAssertTrue stats, prefix & ".noEquilibriumSolve", SectionEquilibriumSolveCount() = 0 And SectionEquilibriumIterationCount() = 0
            TestAssertTrue stats, prefix & ".guessBudget", fake.TestUltimateGuessCalls >= 1 And fake.TestUltimateGuessCalls <= 2
            If scenario = 1 Or scenario = 4 Then
                TestAssertTrue stats, prefix & ".accepted", succeeded And fake.TestConverged
                TestAssertTrue stats, prefix & ".positiveLambda", fake.TestFinalLambda = 1#
                If scenario = 1 Then
                    TestAssertTrue stats, prefix & ".budgetNotReset", fake.TestFinalIterations = 2 And fake.TestUltimateGuessCalls = 2
                Else
                    TestAssertTrue stats, prefix & ".coldNoRestart", fake.TestFinalIterations = 1 And fake.TestUltimateGuessCalls = 1
                End If
            Else
                TestAssertTrue stats, prefix & ".noFalseSuccess", Not succeeded And Not fake.TestConverged
                If scenario = 6 Then
                    TestAssertTrue stats, prefix & ".terminalPreserved", fake.TestFailureCode = sfcInvalidConfiguration And fake.TestUltimateGuessCalls = 1
                Else
                    TestAssertTrue stats, prefix & ".numericalNotPhysical", fake.TestFailureCode = sfcNumericalFailure
                End If
                If scenario = 2 Or scenario = 3 Then TestAssertTrue stats, prefix & ".restartForbidden", fake.TestUltimateGuessCalls = 1
                If scenario = 7 Then TestAssertTrue stats, prefix & ".singleRestart", fake.TestUltimateGuessCalls = 2
            End If
        Next scenario
    Next kind
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' реальные циклы, ошибочный ввод и теплые перезапуски.
' Пробы интегрирования не являются решениями НДС или отдельными итерациями.
Public Function RunSolverWorkStatsTests() As String
    Dim stats As TCapacityTestStats
    On Error GoTo Failed
    Dim section As CSectionModel, solver As CSectionSolver, method As Variant
    Set section = New CSectionModel
    section.AddConcreteElement -10#, -10#, 100#
    section.AddConcreteElement 10#, -10#, 100#
    section.AddConcreteElement -10#, 10#, 100#
    section.AddConcreteElement 10#, 10#, 100#
    Dim calls As Long, iterations As Long
    ResetSectionEquilibriumSolveCount
    For Each method In Array("Newton", "Secant")
        Set solver = New CSectionSolver
        solver.SolverMethod = CStr(method)
        solver.LoadSteps = 3
        solver.Solve section, TestLinearConcrete(), TestLinearSteel(), -1000#, 100#, -200#
        calls = calls + 1: iterations = iterations + solver.Iterations
        TestAssertTrue stats, "workStats." & CStr(method) & ".solved", solver.Converged
        TestAssertTrue stats, "workStats." & CStr(method) & ".calls", SectionEquilibriumSolveCount() = calls
        TestAssertTrue stats, "workStats." & CStr(method) & ".allIterations", SectionEquilibriumIterationCount() = iterations
    Next method
    solver.EvaluatePhysicalStrainPlane section, TestLinearConcrete(), TestLinearSteel(), -0.001, 0#, 0#
    TestAssertTrue stats, "workStats.integrationNotSolve", SectionEquilibriumSolveCount() = calls And SectionEquilibriumIterationCount() = iterations
    solver.MaxIterations = 0
    solver.Solve section, TestLinearConcrete(), TestLinearSteel(), -1000#, 100#, -200#
    TestAssertTrue stats, "workStats.invalidInputNotCounted", SectionEquilibriumSolveCount() = calls And SectionEquilibriumIterationCount() = iterations
    TestWarmStartRestartContract stats
    ResetSectionEquilibriumSolveCount
    TestAssertTrue stats, "workStats.resetAll", SectionEquilibriumSolveCount() = 0 And SectionEquilibriumIterationCount() = 0 And _
        UltimateStrainStartCount() = 0 And UltimateStrainIterationCount() = 0
    RunSolverWorkStatsTests = stats.Report & "TOTAL_SOLVER_WORK_STATS: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    Exit Function
Failed:
    RunSolverWorkStatsTests = stats.Report & "FAIL: workStats.runtime; " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' действительный двухосный UltimateStrain, переносы и разные пределы.
Public Function RunUltimateFrameRegressionTests() As String
    Dim stats As TCapacityTestStats
    On Error GoTo Failed
    TestUltimateMomentTranslationInvariant stats
    TestUltimateStrainToleranceNoHiddenFloor stats
    TestUltimateNumericalContextCache stats
    TestProjectedLimitRecovery stats
    TestAudit02AsymmetricSteelLimits stats
    TestAppendLine stats, "TOTAL_ULTIMATE_FRAME: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunUltimateFrameRegressionTests = stats.Report
    Exit Function
Failed:
    RunUltimateFrameRegressionTests = stats.Report & "FAIL: ultimateFrame.runtime=" & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Плато бетона при слегка несимметричной арматуре требует уточнения момента.
' Предел бетона остается управляющим и при переносе всей модели от начала СК.
Private Sub TestAxialConcretePlateauStart(ByRef stats As TCapacityTestStats)
    Dim shift As Variant, method As Variant, x As Long, y As Long
    Dim section As CSectionModel, props As CSectionPropertiesCalculator
    Dim cap As CCapacitySolver, request As CLimitSearchRequest, search As CLoadMultiplierSearch
    Dim loadState As CSectionLoadState, result As CLimitSearchResult
    Dim referenceLambda As Double, prefix As String
    For Each shift In Array(0#, -3200#, 5000#)
        Set section = New CSectionModel: section.Clear
        For x = -4 To 4
            For y = -2 To 2
                section.AddConcreteElement CDbl(shift) + 50# * x, CDbl(shift) / 4# + 50# * y, 2500#
            Next y
        Next x
        section.AddRebarElement CDbl(shift) - 160#, CDbl(shift) / 4# - 80#, 32#, 0#, "A400"
        section.AddRebarElement CDbl(shift) + 160#, CDbl(shift) / 4# - 80#, 32#, 0#, "A400"
        section.AddRebarElement CDbl(shift) - 160#, CDbl(shift) / 4# + 80#, 32#, 0#, "A400"
        section.AddRebarElement CDbl(shift) + 160#, CDbl(shift) / 4# + 80.2, 32#, 0#, "A400"
        Set props = New CSectionPropertiesCalculator
        props.CalculateTransformed section, TestProvisionalConcrete(), TestProvisionalSteel()
        Set loadState = New CSectionLoadState
        loadState.Initialize -500000#, 0#, 0#, props.CentroidX, props.CentroidY
        For Each method In Array("Bisection", "Brent", "Secant")
            Set cap = New CCapacitySolver
            cap.ConcreteCompressionLimit = -0.0035
            cap.SteelCompressionLimit = 0.025
            cap.SteelTensionLimit = 0.025
            cap.MaxRetries = 0
            cap.SolverBaseLoadSteps = 1
            cap.SearchMethod = CStr(method)
            Set request = New CLimitSearchRequest
            InitializeCapacitySearchRequest request, cap, section, TestProvisionalConcrete(), TestProvisionalSteel(), "LoadMultiplier", _
                0#, loadState.N, 0#, loadState.MxAboutPoint(0#), 0#, loadState.MyAboutPoint(0#), True, True
            Set search = New CLoadMultiplierSearch
            Set result = search.Execute(request)
            prefix = "axialPlateau." & CStr(shift) & "." & CStr(method)
            TestAppendLine stats, prefix & "|lambda=" & TestFormatNumberInvariant(cap.LambdaUltimate) & "|reason=" & cap.StopReason
            TestAssertTrue stats, prefix & ".converged", cap.Converged
            TestAssertEquals stats, prefix & ".criterion", cap.LimitState, "ConcreteStrainLimit"
            If cap.Converged Then
                TestAssertEquilibrium stats, prefix, cap.LastSolver, cap.NUltimate, cap.MxUltimate, cap.MyUltimate
                TestAssertClose stats, prefix & ".concreteLimit", cap.LastSolver.MinConcreteStrain, -0.0035, 0.00000001
                If referenceLambda = 0# Then referenceLambda = cap.LambdaUltimate
                TestAssertClose stats, prefix & ".translationInvariant", cap.LambdaUltimate, referenceLambda, 0.00001
            End If
        Next method
    Next shift
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Перенос начала координат не должен менять предел или нарушать равновесие.
' Проверяются четыре направления двухосного изгиба на несимметричной арматуре.
Private Sub TestUltimateMomentTranslationInvariant(ByRef stats As TCapacityTestStats)
    Dim shift As Variant, loads As Variant, x As Long, y As Long, index As Long
    Dim section As CSectionModel, props As CSectionPropertiesCalculator, cap As CCapacitySolver
    Dim referenceLambda(0 To 3) As Double, prefix As String, targetN As Double
    loads = Array(Array(10000000#, 5000000#), Array(-10000000#, 5000000#), _
        Array(10000000#, -5000000#), Array(-10000000#, -5000000#))
    targetN = -100000#
    For Each shift In Array(0#, -3200#, 5000#, -1000000#, 1000000#)
        Set section = New CSectionModel
        For x = -4 To 4
            For y = -2 To 2
                section.AddConcreteElement CDbl(shift) + 50# * x, CDbl(shift) / 4# + 50# * y, 2500#
            Next y
        Next x
        section.AddRebarElement CDbl(shift) - 160#, CDbl(shift) / 4# - 80#, 32#, 0#, "A400"
        section.AddRebarElement CDbl(shift) + 160#, CDbl(shift) / 4# - 80#, 32#, 0#, "A400"
        section.AddRebarElement CDbl(shift) - 160#, CDbl(shift) / 4# + 80#, 32#, 0#, "A400"
        section.AddRebarElement CDbl(shift) + 160#, CDbl(shift) / 4# + 80.2, 32#, 0#, "A400"
        Set props = New CSectionPropertiesCalculator
        props.CalculateConcrete section
        For index = 0 To 3
            Set cap = New CCapacitySolver
            TestConfigureCapacity cap
            cap.ConcreteCompressionLimit = -0.0035
            cap.SteelStrainLimit = 0.025
            cap.SolveByUltimateLoadPath section, TestProvisionalConcrete(), TestProvisionalSteel(), _
                targetN, 0#, targetN * props.CentroidY, CDbl(loads(index)(0)), _
                targetN * props.CentroidX, CDbl(loads(index)(1))
            prefix = "ultimateTranslation." & CStr(shift) & "." & CStr(index)
            TestAppendLine stats, prefix & "|lambda=" & TestFormatNumberInvariant(cap.LambdaUltimate) & "|reason=" & cap.StopReason
            TestAssertTrue stats, prefix & ".converged", cap.Converged
            TestAssertEquals stats, prefix & ".method", cap.SolutionMethod, "UltimateStrain"
            If cap.Converged Then
                TestAssertEquilibrium stats, prefix, cap.LastSolver, targetN, _
                    targetN * props.CentroidY + cap.LambdaUltimate * CDbl(loads(index)(0)), _
                    targetN * props.CentroidX + cap.LambdaUltimate * CDbl(loads(index)(1))
                TestAssertClose stats, prefix & ".physicalLimit", cap.TestCriticalStrainUtilization, 1#, 0.0001
                If CDbl(shift) = 0# Then referenceLambda(index) = cap.LambdaUltimate
                TestAssertClose stats, prefix & ".translationInvariant", cap.LambdaUltimate, referenceLambda(index), 0.0001
            End If
        Next index
    Next shift
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Решает два направления Mx при постоянной сжимающей N: проверяет знаки
' предельных моментов, физический критерий и равновесие каждой конечной точки.
Private Sub TestMxPositiveAndNegative(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim capNeg As CCapacitySolver
    Set capNeg = New CCapacitySolver
    TestConfigureCapacity capNeg
    capNeg.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -300000#, -10000000#, 0#

    Dim capPos As CCapacitySolver
    Set capPos = New CCapacitySolver
    TestConfigureCapacity capPos
    capPos.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -300000#, 10000000#, 0#

    TestAssertTrue stats, "capacity.mx.negative.converged", capNeg.Converged
    TestAssertTrue stats, "capacity.mx.positive.converged", capPos.Converged
    TestAssertTrue stats, "capacity.mx.negative.sign", capNeg.MxUltimate < 0#
    TestAssertTrue stats, "capacity.mx.positive.sign", capPos.MxUltimate > 0#
    TestAssertTrue stats, "capacity.mx.limit.physical", TestIsPhysicalLimitState(capNeg.LimitState) And TestIsPhysicalLimitState(capPos.LimitState)
    TestAssertEquilibrium stats, "capacity.mx.negative", capNeg.LastSolver, -300000#, capNeg.MxUltimate, 0#
    TestAssertEquilibrium stats, "capacity.mx.positive", capPos.LastSolver, -300000#, capPos.MxUltimate, 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет симметрию несущей круглого сечения по X/Y и сохранение направления
' косого изгиба. Допуск симметрии учитывает дискретизацию исходной тестовой сетки.
Private Sub TestCircleCapacitySymmetry(ByRef stats As TCapacityTestStats)
    Dim geom As CGeometryRegion
    Set geom = TestFixtureDisk(300#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 10#, 10#, 1

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", 0#, 90#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B2", 90#, 0#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B3", 0#, -90#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B4", -90#, 0#, 20#, 0#, "A400", "", geom

    Dim capMx As CCapacitySolver
    Set capMx = New CCapacitySolver
    TestConfigureCapacity capMx
    capMx.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -220000#, -6000000#, 0#

    Dim capMy As CCapacitySolver
    Set capMy = New CCapacitySolver
    TestConfigureCapacity capMy
    capMy.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -220000#, 0#, -6000000#

    Dim capMxy As CCapacitySolver
    Set capMxy = New CCapacitySolver
    TestConfigureCapacity capMxy
    capMxy.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -220000#, -5000000#, -5000000#

    TestAssertTrue stats, "circle.capacity.mx.converged", capMx.Converged
    TestAssertTrue stats, "circle.capacity.my.converged", capMy.Converged
    TestAssertTrue stats, "circle.capacity.mxy.converged", capMxy.Converged
    TestAssertRelative stats, "circle.capacity.mx.my.symmetry", Abs(capMx.MomentUltimate), Abs(capMy.MomentUltimate), 0.03
    TestAssertTrue stats, "circle.capacity.mxy.lambda.positive", capMxy.LambdaUltimate > 0#
    TestAssertClose stats, "circle.capacity.mxy.direction", capMxy.MxUltimate * -5000000# - capMxy.MyUltimate * -5000000#, 0#, 1000#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает Ignore/UseDiagram для растянутого бетона при одинаковых нагрузках.
' Активная ветвь должна менять равновесную плоскость в одно-/двухосных задачах.
Private Sub TestConcreteTensionBehaviorAffectsSolverAndCapacity(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim ignoreConcrete As CMaterialDiagram
    Set ignoreConcrete = TestProvisionalConcrete()

    Dim useConcrete As CMaterialDiagram
    Set useConcrete = TestProvisionalConcreteWithTension()

    Dim solverIgnore As CSectionSolver
    Set solverIgnore = New CSectionSolver
    TestConfigureSectionSolver solverIgnore
    solverIgnore.Solve TestBuildGeneratedSectionModel(mesh, rebars), ignoreConcrete, TestProvisionalSteel(), -50000#, -5000000#, -3000000#

    Dim solverUse As CSectionSolver
    Set solverUse = New CSectionSolver
    TestConfigureSectionSolver solverUse
    solverUse.Solve TestBuildGeneratedSectionModel(mesh, rebars), useConcrete, TestProvisionalSteel(), -50000#, -5000000#, -3000000#

    TestAssertTrue stats, "tensionMode.solver.ignore.converged", solverIgnore.Converged
    TestAssertTrue stats, "tensionMode.solver.use.converged", solverUse.Converged
    TestAssertTrue stats, "tensionMode.solver.affectsState", Abs(solverIgnore.KappaX - solverUse.KappaX) > 0.000000001 Or Abs(solverIgnore.KappaY - solverUse.KappaY) > 0.000000001

    TestAssertTensionBehaviorAffectsSolverMode stats, "tensionBehavior.strength.mx", mesh, rebars, -50000#, -5000000#, 0#
    TestAssertTensionBehaviorAffectsSolverMode stats, "tensionBehavior.strength.my", mesh, rebars, -50000#, 0#, -5000000#
    TestAssertTensionBehaviorAffectsSolverMode stats, "tensionBehavior.strength.mxy", mesh, rebars, -50000#, -5000000#, -3000000#
End Sub


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Решает один вектор усилий с отключенным и активным растянутым бетоном;
' проверяет сходимость обеих моделей и значимое различие параметров плоскости.
Private Sub TestAssertTensionBehaviorAffectsSolverMode(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double)
    Dim ignoreConcrete As CMaterialDiagram
    Set ignoreConcrete = TestProvisionalConcrete()

    Dim useConcrete As CMaterialDiagram
    Set useConcrete = TestProvisionalConcreteWithTension()

    Dim solverIgnore As CSectionSolver
    Set solverIgnore = New CSectionSolver
    TestConfigureSectionSolver solverIgnore
    solverIgnore.Solve TestBuildGeneratedSectionModel(mesh, rebars), ignoreConcrete, TestProvisionalSteel(), nValue, mxValue, myValue

    Dim solverUse As CSectionSolver
    Set solverUse = New CSectionSolver
    TestConfigureSectionSolver solverUse
    solverUse.Solve TestBuildGeneratedSectionModel(mesh, rebars), useConcrete, TestProvisionalSteel(), nValue, mxValue, myValue

    TestAssertTrue stats, prefix & ".ignore.converged", solverIgnore.Converged
    TestAssertTrue stats, prefix & ".use.converged", solverUse.Converged
    TestAssertTrue stats, prefix & ".stateDiffers", Abs(solverIgnore.KappaX - solverUse.KappaX) > 0.000000001 Or _
        Abs(solverIgnore.KappaY - solverUse.KappaY) > 0.000000001 Or Abs(solverIgnore.Epsilon0 - solverUse.Epsilon0) > 0.000000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет все четыре квадранта Mx/My при одинаковой постоянной N.
' Предельные моменты сохраняют знаки и отношение компонентов заданного пути.
Private Sub TestMxySignedCombinations(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    TestCheckMxyCombination stats, "capacity.mxy.pp", mesh, rebars, -250000#, 6000000#, 4000000#
    TestCheckMxyCombination stats, "capacity.mxy.np", mesh, rebars, -250000#, -6000000#, 4000000#
    TestCheckMxyCombination stats, "capacity.mxy.pn", mesh, rebars, -250000#, 6000000#, -4000000#
    TestCheckMxyCombination stats, "capacity.mxy.nn", mesh, rebars, -250000#, -6000000#, -4000000#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет одну косую траекторию: положительную lambda, направление моментов
' и равновесие конечной точки. Не заменяет косую задачу независимыми осями.
Private Sub TestCheckMxyCombination(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal nValue As Double, ByVal mxBase As Double, ByVal myBase As Double)
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    cap.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), nValue, mxBase, myBase

    TestAssertTrue stats, prefix & ".converged", cap.Converged
    TestAssertTrue stats, prefix & ".lambda.positive", cap.LambdaUltimate > 0#
    TestAssertTrue stats, prefix & ".mx.sign", Sgn(cap.MxUltimate) = Sgn(mxBase)
    TestAssertTrue stats, prefix & ".my.sign", Sgn(cap.MyUltimate) = Sgn(myBase)
    TestAssertClose stats, prefix & ".direction", cap.MxUltimate * myBase - cap.MyUltimate * mxBase, 0#, 1000#
    TestAssertEquilibrium stats, prefix, cap.LastSolver, nValue, cap.MxUltimate, cap.MyUltimate
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Решает оба направления My при постоянной N и проверяет знаки конечных
' моментов, физический управляющий предел и компонентное равновесие.
Private Sub TestMyPositiveAndNegative(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim capNeg As CCapacitySolver
    Set capNeg = New CCapacitySolver
    TestConfigureCapacity capNeg
    capNeg.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -300000#, 0#, -10000000#

    Dim capPos As CCapacitySolver
    Set capPos = New CCapacitySolver
    TestConfigureCapacity capPos
    capPos.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -300000#, 0#, 10000000#

    TestAssertTrue stats, "capacity.my.negative.converged", capNeg.Converged
    TestAssertTrue stats, "capacity.my.positive.converged", capPos.Converged
    TestAssertTrue stats, "capacity.my.negative.sign", capNeg.MyUltimate < 0#
    TestAssertTrue stats, "capacity.my.positive.sign", capPos.MyUltimate > 0#
    TestAssertTrue stats, "capacity.my.limit.physical", TestIsPhysicalLimitState(capNeg.LimitState) And TestIsPhysicalLimitState(capPos.LimitState)
    TestAssertEquilibrium stats, "capacity.my.negative", capNeg.LastSolver, -300000#, 0#, capNeg.MyUltimate
    TestAssertEquilibrium stats, "capacity.my.positive", capPos.LastSolver, -300000#, 0#, capPos.MyUltimate
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Удваивает найденный предельный момент и ожидает lambda < 1.
' Успешный поиск предела не должен скрывать непрохождение заданной нагрузки.
Private Sub TestLambdaLessThanOne(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim reference As CCapacitySolver
    Set reference = New CCapacitySolver
    TestConfigureCapacity reference
    reference.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -300000#, -10000000#, 0#

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    cap.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -300000#, reference.MxUltimate * 2#, 0#

    TestAssertTrue stats, "capacity.lambda.lessThanOne.converged", cap.Converged
    TestAssertTrue stats, "capacity.lambda.lessThanOne.value", cap.LambdaUltimate < 1#
    TestCheckAudit02SearchVsCapacity stats, cap, TestBuildGeneratedSectionModel(mesh, rebars), _
        -300000#, reference.MxUltimate * 2#, 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторно ищет предел от момента уже найденной опорной точки;
' lambda должна быть около 1 в исходном допуске, без изменения критерия.
Private Sub TestLambdaNearOne(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim reference As CCapacitySolver
    Set reference = New CCapacitySolver
    TestConfigureCapacity reference
    reference.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -300000#, -10000000#, 0#

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    cap.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -300000#, reference.MxUltimate, 0#

    TestAssertTrue stats, "capacity.lambda.nearOne.converged", cap.Converged
    TestAssertClose stats, "capacity.lambda.nearOne.value", cap.LambdaUltimate, 1#, 0.03
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет чистый изгиб при N=0 на линейных материалах с заданными пределами.
' Нулевая продольная сила не должна нарушать поиск или равновесие точки.
Private Sub TestZeroAxialForce(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    cap.ConcreteCompressionLimit = -0.0006
    cap.SteelStrainLimit = 0.001
    cap.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestLinearConcrete(), TestLinearSteel(), 0#, -5000000#, 0#

    TestAssertTrue stats, "capacity.zeroN.converged", cap.Converged
    TestAssertEquilibrium stats, "capacity.zeroN", cap.LastSolver, 0#, cap.MxUltimate, 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет траекторию изменения чистой продольной силы.
' Меняется только N; нулевые Mx и My допустимы.
Private Sub TestAxialLoadMultiplierFindsNult(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim section As CSectionModel
    Set section = TestBuildGeneratedSectionModel(mesh, rebars)

    Dim compression As CCapacitySolver
    Set compression = New CCapacitySolver
    TestConfigureCapacity compression
    compression.SolveByLoadPathMultiplier section, TestProvisionalConcrete(), TestProvisionalSteel(), _
        0#, -100000#, 0#, 0#, 0#, 0#

    Dim tension As CCapacitySolver
    Set tension = New CCapacitySolver
    TestConfigureCapacity tension
    tension.SolveByLoadPathMultiplier section, TestProvisionalConcrete(), TestProvisionalSteel(), _
        0#, 50000#, 0#, 0#, 0#, 0#

    TestAssertTrue stats, "capacity.nult.compression.converged", compression.Converged
    TestAssertTrue stats, "capacity.nult.compression.sign", compression.NUltimate < 0#
    TestAssertTrue stats, "capacity.nult.compression.safety", Abs(compression.NUltimate / -100000#) > 1#
    TestAssertEquilibrium stats, "capacity.nult.compression", compression.LastSolver, compression.NUltimate, 0#, 0#

    TestAssertTrue stats, "capacity.nult.tension.converged", tension.Converged
    TestAssertTrue stats, "capacity.nult.tension.sign", tension.NUltimate > 0#
    TestAssertTrue stats, "capacity.nult.tension.safety", Abs(tension.NUltimate / 50000#) > 1#
    TestAssertEquilibrium stats, "capacity.nult.tension", tension.LastSolver, tension.NUltimate, 0#, 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Делает сжатие бетона управляющим ограничением, оставляя большой предел
' арматуры; поиск должен завершиться именно ConcreteStrainLimit.
Private Sub TestConcreteLimitState(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    cap.ConcreteCompressionLimit = -0.0006
    cap.SteelStrainLimit = 1#
    cap.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestLinearConcrete(), TestLinearSteel(), 0#, -10000000#, 0#

    TestAssertTrue stats, "capacity.concreteLimit.converged", cap.Converged
    TestAssertEquals stats, "capacity.concreteLimit.state", cap.LimitState, "ConcreteStrainLimit"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Включает активную tensile-ветвь и малый растягивающий предел бетона.
' Управляющий код должен относиться к растяжению, а не сжатию или арматуре.
Private Sub TestConcreteTensionLimitState(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim concrete As CMaterialDiagram
    Set concrete = TestProvisionalConcreteWithTension()

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    cap.InitialLambdaStep = 0.1
    cap.ConcreteCompressionLimit = -1#
    cap.ConcreteTensionLimit = 0.00000001
    cap.ConcreteTensionLimitEnabled = True
    cap.SteelStrainLimit = 1#
    cap.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), concrete, TestProvisionalSteel(), -50000#, -5000000#, 0#

    TestAssertTrue stats, "capacity.concreteTensionLimit.converged", cap.Converged
    TestAssertEquals stats, "capacity.concreteTensionLimit.state", cap.LimitState, "ConcreteTensionStrainLimit"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' При отключённом растянутом бетоне малый предел его растяжения
' не ограничивает несущую способность. В этом примере предел задаёт арматура.
Private Sub TestConcreteTensionLimitIgnored(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim concrete As CMaterialDiagram
    Set concrete = TestProvisionalConcrete()

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    cap.InitialLambdaStep = 0.1
    cap.ConcreteCompressionLimit = -1#
    cap.ConcreteTensionLimit = 0.00000001
    cap.ConcreteTensionLimitEnabled = False
    cap.SteelStrainLimit = 0.00000001
    cap.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), concrete, TestProvisionalSteel(), -50000#, -5000000#, 0#

    TestAssertTrue stats, "capacity.concreteTensionLimit.ignore.notTension", cap.LimitState <> "ConcreteTensionStrainLimit"
    TestAssertEquals stats, "capacity.concreteTensionLimit.ignore.state", cap.LimitState, "SteelStrainLimit"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Делает деформацию арматуры управляющей на линейных материалах;
' ожидается SteelStrainLimit при сохраненном поиске равновесия.
Private Sub TestSteelLimitState(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    cap.ConcreteCompressionLimit = -1#
    cap.SteelStrainLimit = 0.0005
    cap.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestLinearConcrete(), TestLinearSteel(), 0#, -10000000#, 0#

    TestAssertTrue stats, "capacity.steelLimit.converged", cap.Converged
    TestAssertEquals stats, "capacity.steelLimit.state", cap.LimitState, "SteelStrainLimit"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет старт чистого изгиба из подготовленной StateGuessBuilder плоскости.
' При жестком бюджете итераций должен находиться физический предел, а журнал
' подтверждать использование начального приближения вместо нулевого излома.
Private Sub TestLoadMultiplierPureBendingUsesStateGuess(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    cap.SolverMaxIterations = 1
    cap.MaxRetries = 1
    cap.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), 0#, -10000000#, 0#

    If Not cap.Converged Then TestAppendLine stats, "DIAGNOSTIC pureBendingGuess: " & cap.StopReason & vbCrLf & cap.DiagnosticLog

    TestAssertTrue stats, "capacity.pureBendingGuess.converged", cap.Converged
    TestAssertTrue stats, "capacity.pureBendingGuess.physical", TestIsPhysicalLimitState(cap.LimitState)
    TestAssertTrue stats, "capacity.pureBendingGuess.usedGuess", InStr(1, cap.DiagnosticLog, "для lambda-точки чистого изгиба применена стартовая плоскость", vbTextCompare) > 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет связность кривизн несимметричного сечения с ненулевым Ixy:
' даже при My=0 равновесная KappaY может быть ненулевой.
Private Sub TestAsymmetricCoupledCurvatures(ByRef stats As TCapacityTestStats)
    Dim geom As CGeometryRegion
    Set geom = TestCapacityAsymmetricGeometry()

    Dim mesh As CFiberMeshBuilder
    Set mesh = TestBuildMesh(geom, 10#)

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete TestBuildGeneratedSectionModel(mesh, Nothing)

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", 50#, 50#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B2", 550#, 50#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B3", 50#, 700#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B4", 200#, 700#, 20#, 0#, "A400", "", geom

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    cap.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -260000#, -8000000#, 0#

    TestAssertTrue stats, "capacity.asym.Ixy.nonzero", Abs(props.Ixyc) > 1000000#
    TestAssertTrue stats, "capacity.asym.converged", cap.Converged
    TestAssertTrue stats, "capacity.asym.coupledKappaY", Abs(cap.LastSolver.KappaY) > 0.000000001
    TestAssertEquilibrium stats, "capacity.asym", cap.LastSolver, -260000#, cap.MxUltimate, 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Решает косой изгиб несимметричного Rectangles и проверяет обе кривизны,
' отношение предельных моментов и равновесие относительно исходных осей.
Private Sub TestAsymmetricMxy(ByRef stats As TCapacityTestStats)
    Dim geom As CGeometryRegion
    Set geom = TestCapacityAsymmetricGeometry()

    Dim mesh As CFiberMeshBuilder
    Set mesh = TestBuildMesh(geom, 10#)

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete TestBuildGeneratedSectionModel(mesh, Nothing)

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", 50#, 50#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B2", 550#, 50#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B3", 50#, 700#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B4", 200#, 700#, 20#, 0#, "A400", "", geom

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    cap.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -260000#, -6000000#, -4000000#

    TestAssertTrue stats, "capacity.mxy.asym.Ixy.nonzero", Abs(props.Ixyc) > 1000000#
    TestAssertTrue stats, "capacity.mxy.asym.converged", cap.Converged
    TestAssertTrue stats, "capacity.mxy.asym.kappaX.nonzero", Abs(cap.LastSolver.KappaX) > 0.000000001
    TestAssertTrue stats, "capacity.mxy.asym.kappaY.nonzero", Abs(cap.LastSolver.KappaY) > 0.000000001
    TestAssertClose stats, "capacity.mxy.asym.direction", cap.MxUltimate * -4000000# - cap.MyUltimate * -6000000#, 0#, 1000#
    TestAssertEquilibrium stats, "capacity.mxy.asym", cap.LastSolver, -260000#, cap.MxUltimate, cap.MyUltimate
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Передает нулевую масштабируемую часть именно моментного пути.
' Наличие постоянной N не делает такой путь допустимым: ожидается InvalidInput.
Private Sub TestInvalidBaseMoment(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 200#, 100#, 20#, 60#, 30#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    cap.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -100000#, 0#, 0#

    TestAssertTrue stats, "capacity.invalidBaseMoment.notConverged", Not cap.Converged
    TestAssertEquals stats, "capacity.invalidBaseMoment.state", cap.LimitState, "InvalidInput"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает LoadMultiplier и UltimateStrain на прямоугольнике и круге,
' включая определяющие деформации и причины численной неудачи.
Private Sub TestCapacitySolutionStrategyComparisons(ByRef stats As TCapacityTestStats)
    TestMethodPureCompression stats
    TestCompareCapacitySolutionStrategys stats, "method.n_plus_mx", -300000#, -10000000#, 0#, False
    TestCompareCapacitySolutionStrategys stats, "method.n_plus_my", -300000#, 0#, -10000000#, False
    TestCompareCapacitySolutionStrategys stats, "method.biaxial", -250000#, -6000000#, -4000000#, False
    TestCompareCircleCapacitySolutionStrategys stats
    TestMethodStrainLimitState stats
    TestUltimateStrainNumericalFailure stats
    TestUltimateStrainDoesNotApplyMaxLambda stats
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что моментные entrypoint-ы обоих методов отклоняют чистую N
' без базового момента; допустимая осевая lambda*N тестируется отдельным путем.
Private Sub TestMethodPureCompression(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim loadMethod As CCapacitySolver
    Set loadMethod = New CCapacitySolver
    TestConfigureCapacity loadMethod
    loadMethod.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -300000#, 0#, 0#

    Dim strainMethod As CCapacitySolver
    Set strainMethod = New CCapacitySolver
    TestConfigureCapacity strainMethod
    strainMethod.SolveByUltimateStrain TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -300000#, 0#, 0#

    TestAssertEquals stats, "method.pure_compression.load.invalid", loadMethod.LimitState, "InvalidInput"
    TestAssertEquals stats, "method.pure_compression.strain.invalid", strainMethod.LimitState, "InvalidInput"
    TestAppendComparison stats, "method.pure_compression.load", loadMethod, 0#
    TestAppendComparison stats, "method.pure_compression.strain", strainMethod, 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Решает одинаковую моментную траекторию двумя стратегиями и сравнивает
' lambda/управляющий предел; время сохраняется как наблюдение, не порог PASS.
Private Sub TestCompareCapacitySolutionStrategys(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal nValue As Double, ByVal mxBase As Double, ByVal myBase As Double, ByVal tightLimits As Boolean)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim loadMethod As CCapacitySolver
    Dim strainMethod As CCapacitySolver
    Dim t0 As Double
    Dim loadElapsed As Double
    Dim strainElapsed As Double

    Set loadMethod = New CCapacitySolver
    TestConfigureCapacity loadMethod
    If tightLimits Then
        loadMethod.ConcreteCompressionLimit = -0.0006
        loadMethod.SteelStrainLimit = 0.001
    End If
    t0 = Timer
    loadMethod.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), nValue, mxBase, myBase
    loadElapsed = Timer - t0

    Set strainMethod = New CCapacitySolver
    TestConfigureCapacity strainMethod
    If tightLimits Then
        strainMethod.ConcreteCompressionLimit = -0.0006
        strainMethod.SteelStrainLimit = 0.001
    End If
    t0 = Timer
    strainMethod.SolveByUltimateStrain TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), nValue, mxBase, myBase
    strainElapsed = Timer - t0

    TestAssertTrue stats, prefix & ".load.converged", loadMethod.Converged
    TestAssertTrue stats, prefix & ".strain.converged", strainMethod.Converged
    TestAssertClose stats, prefix & ".lambda", strainMethod.LambdaUltimate, loadMethod.LambdaUltimate, 0.02
    TestAssertEquals stats, prefix & ".limitState", strainMethod.LimitState, loadMethod.LimitState
    TestAssertTrue stats, prefix & ".criticalElement", Len(strainMethod.TestCriticalElement) > 0
    TestAssertTrue stats, prefix & ".solverCalls", strainMethod.Iterations > 0
    TestAppendComparison stats, prefix & ".load", loadMethod, loadElapsed
    TestAppendComparison stats, prefix & ".strain", strainMethod, strainElapsed
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сопоставляет стратегии на симметричном круге с повернутыми X/Y задачами;
' проверяет конечные моменты и определение критического элемента.
Private Sub TestCompareCircleCapacitySolutionStrategys(ByRef stats As TCapacityTestStats)
    Dim geom As CGeometryRegion
    Set geom = TestFixtureDisk(300#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 12#, 12#, 1

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", 0#, 90#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B2", 90#, 0#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B3", 0#, -90#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B4", -90#, 0#, 20#, 0#, "A400", "", geom

    Dim loadMethod As CCapacitySolver
    Set loadMethod = New CCapacitySolver
    TestConfigureCapacity loadMethod
    loadMethod.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -220000#, -6000000#, 0#

    Dim strainMethod As CCapacitySolver
    Set strainMethod = New CCapacitySolver
    TestConfigureCapacity strainMethod
    strainMethod.SolveByUltimateStrain TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -220000#, 0#, -6000000#

    TestAssertTrue stats, "method.circle.load.converged", loadMethod.Converged
    TestAssertTrue stats, "method.circle.strain.converged", strainMethod.Converged
    TestAssertRelative stats, "method.circle.symmetry", Abs(loadMethod.MomentUltimate), Abs(strainMethod.MomentUltimate), 0.04
    TestAssertTrue stats, "method.circle.criticalElement", Len(strainMethod.TestCriticalElement) > 0
    TestAppendComparison stats, "method.circle.load", loadMethod, 0#
    TestAppendComparison stats, "method.circle.strain", strainMethod, 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает обе стратегии при специально заданном бетонном strain-пределе.
' UltimateStrain должен подтвердить тот же управляющий критерий и utilization.
Private Sub TestMethodStrainLimitState(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim loadMethod As CCapacitySolver
    Set loadMethod = New CCapacitySolver
    TestConfigureCapacity loadMethod
    loadMethod.ConcreteCompressionLimit = -0.0006
    loadMethod.SteelStrainLimit = 1#
    loadMethod.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestLinearConcrete(), TestLinearSteel(), 0#, -10000000#, 0#

    Dim strainMethod As CCapacitySolver
    Set strainMethod = New CCapacitySolver
    TestConfigureCapacity strainMethod
    strainMethod.ConcreteCompressionLimit = -0.0006
    strainMethod.SteelStrainLimit = 1#
    strainMethod.SolveByUltimateStrain TestBuildGeneratedSectionModel(mesh, rebars), TestLinearConcrete(), TestLinearSteel(), 0#, -10000000#, 0#

    TestAssertTrue stats, "method.strain_limit.load.converged", loadMethod.Converged
    TestAssertTrue stats, "method.strain_limit.strain.converged", strainMethod.Converged
    TestAssertClose stats, "method.strain_limit.lambda", strainMethod.LambdaUltimate, loadMethod.LambdaUltimate, 0.02
    TestAssertEquals stats, "method.strain_limit.limitState", strainMethod.LimitState, "ConcreteStrainLimit"
    TestAssertTrue stats, "method.strain_limit.criticalElement", Len(strainMethod.TestCriticalElement) > 0
    TestAssertTrue stats, "method.strain_limit.criticalUtilization", strainMethod.TestCriticalStrainUtilization > 0.98
    TestAppendComparison stats, "method.strain_limit.load", loadMethod, 0#
    TestAppendComparison stats, "method.strain_limit.strain", strainMethod, 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет невозможный момент: неподтвержденная проба не становится
' физической верхней границей независимо от доступного бюджета итераций.
Private Sub TestUltimateStrainNumericalFailure(ByRef stats As TCapacityTestStats)
    Dim section As CSectionModel ' Все элементы находятся в начале координат: Mx=My=0 при любой плоскости.
    Set section = New CSectionModel
    section.AddConcreteElement 0#, 0#, 60000#
    section.AddRebarElement 0#, 0#, 20#, 0#, "A400"

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    cap.SolverMaxIterations = 100
    cap.MaxRetries = 1
    cap.SolveByUltimateStrain section, TestProvisionalConcrete(), TestProvisionalSteel(), 0#, -10000000#, 0#

    TestAssertTrue stats, "method.numericalFailure.notConverged", Not cap.Converged
    TestAssertTrue stats, "method.numericalFailure.state", cap.LimitState = "NumericalFailure" Or cap.LimitState = "SingularTangent"
    TestAssertTrue stats, "method.numericalFailure.noPhysicalUpper", Not TestIsPhysicalLimitState(cap.LimitState)
    TestAppendComparison stats, "method.numericalFailure.strain", cap, 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отрицательный контракт с доказанно невозможным моментом.
Public Function RunUltimateImpossibleEquilibriumTests() As String
    Dim stats As TCapacityTestStats
    On Error GoTo Failed
    TestUltimateStrainNumericalFailure stats
    TestAppendLine stats, "TOTAL_ULTIMATE_IMPOSSIBLE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunUltimateImpossibleEquilibriumTests = stats.Report
    Exit Function
Failed:
    RunUltimateImpossibleEquilibriumTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что MaxLambda не отбрасывает уже найденное UltimateStrain-решение.
' Эта настройка ограничивает только расширение скобки LoadMultiplier.
Private Sub TestUltimateStrainDoesNotApplyMaxLambda(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    cap.MaxLambda = 0.5
    cap.SolveByUltimateStrain TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -300000#, -1000000#, 0#

    TestAssertTrue stats, "method.maxLambdaIgnored.converged", cap.Converged
    TestAssertTrue stats, "method.maxLambdaIgnored.lambdaAboveLimit", cap.LambdaUltimate > 0.5
    TestAssertTrue stats, "method.maxLambdaIgnored.momentStored", Abs(cap.MomentUltimate) > 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет пользовательский сценарий с дефолтными единицами книги.
' После перехода INPUT-нагрузок на tf и tf*m численные значения допусков
' solver-а тоже должны быть заданы в этих единицах. Иначе 1000 в строке
' Solver.ToleranceMx превращается в 1000 tf*m, LoadMultiplier принимает
' грубо несбалансированные probe-точки и уходит далеко за реальный предел.
' Rectangles/Generated выбираются только на время чтения настроек; исходный
' выбор пользователя восстанавливается до формирования контрольной модели.
Private Sub TestLoadMultiplierWithWorkbookTfDefaults(ByRef stats As TCapacityTestStats)
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader

    Dim inputs As Object, typeCell As Object, sourceCell As Object, data As Variant, row As Long
    Set inputs = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    data = inputs.Value2 ' Ключи читаются одним блоком; сохранённый тип пользовательского сечения не является условием теста.
    For row = 2 To UBound(data, 1)
        If CStr(data(row, 1)) = "Geometry.Type" Then Set typeCell = inputs.Cells(row, 2)
        If CStr(data(row, 1)) = "Geometry.Source" Then Set sourceCell = inputs.Cells(row, 2)
    Next row
    If typeCell Is Nothing Or sourceCell Is Nothing Then Err.Raise vbObjectError + 5590, _
        "modTestCapacitySolver", "Для контрольного сечения не найдены настройки Geometry.Type и Geometry.Source."
    Dim savedType As Variant, savedSource As Variant
    savedType = typeCell.Formula: savedSource = sourceCell.Formula ' Восстанавливаются и значения, и возможные пользовательские формулы.
    On Error GoTo RestoreFixture
    typeCell.Value2 = "Rectangles": sourceCell.Value2 = "Generated"
    settings.LoadFromWorkbook ThisWorkbook
    typeCell.Formula = savedType: sourceCell.Formula = savedSource
    On Error GoTo 0

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    TestAssertEquals stats, "capacity.tfDefaults.geometry", settings.GetString("Geometry.Type", ""), "Rectangles"
    TestAssertEquals stats, "capacity.tfDefaults.forceUnit", settings.GetString("Units.Force.Input", ""), "tf"
    TestAssertEquals stats, "capacity.tfDefaults.momentUnit", settings.GetString("Units.Moment.Input", ""), "tf*m"

    Dim registry As CSectionTypeRegistry
    Set registry = New CSectionTypeRegistry

    Dim section As CSectionModel
    Set section = BuildWorkbookSectionModel(ThisWorkbook, settings, units)

    Dim materialProvider As CMaterialModelProvider
    Set materialProvider = New CMaterialModelProvider
    materialProvider.Initialize settings, units

    Dim concrete As Object
    Set concrete = materialProvider.ConcreteMaterial(cpStrength)

    Dim steel As Object
    Set steel = materialProvider.SteelMaterial(cpStrength)

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateTransformed section, concrete, steel

    Dim nValue As Double
    Dim mxValue As Double
    Dim myValue As Double
    Dim mxOffset As Double
    Dim myOffset As Double
    nValue = units.InputForceToInternal(100#)
    mxValue = units.InputMomentMxToInternal(50#)
    myValue = 0#
    Dim loadState As CSectionLoadState
    Set loadState = New CSectionLoadState
    loadState.Initialize nValue, mxValue, myValue, props.CentroidX, props.CentroidY
    mxOffset = loadState.AxialMxAboutPoint(0#)
    myOffset = loadState.AxialMyAboutPoint(0#)

    Dim reference As CCapacitySolver
    Set reference = New CCapacitySolver
    reference.ApplySettings settings, units
    reference.SolveByUltimateStrain section, concrete, steel, nValue, mxValue, myValue, mxOffset, myOffset
    TestAssertTrue stats, "capacity.tfDefaults.ultimate.converged", reference.Converged
    TestAssertTrue stats, "capacity.tfDefaults.ultimate.lambda", reference.LambdaUltimate > 0# And reference.LambdaUltimate < 10#

    TestAssertLoadMultiplierTfDefault stats, "Bisection", settings, units, section, concrete, steel, _
        nValue, mxValue, myValue, mxOffset, myOffset, reference.LambdaUltimate
    TestAssertLoadMultiplierTfDefault stats, "Brent", settings, units, section, concrete, steel, _
        nValue, mxValue, myValue, mxOffset, myOffset, reference.LambdaUltimate
    TestAssertLoadMultiplierTfDefault stats, "Secant", settings, units, section, concrete, steel, _
        nValue, mxValue, myValue, mxOffset, myOffset, reference.LambdaUltimate
    TestAssertEquals stats, "capacity.tfDefaults.restored.geometry", CStr(typeCell.Formula), CStr(savedType)
    TestAssertEquals stats, "capacity.tfDefaults.restored.source", CStr(sourceCell.Formula), CStr(savedSource)
    Exit Sub

RestoreFixture:
    Dim errorNumber As Long, errorSource As String, errorDescription As String
    errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
    typeCell.Formula = savedType: sourceCell.Formula = savedSource
    Err.Raise errorNumber, errorSource, errorDescription ' Ошибка чтения остаётся ошибкой теста, но выбранное пользователем сечение сохраняется.
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет один метод LoadMultiplier на настройках, прочитанных в tf/tf*m.
' Сравнение с независимой веткой UltimateStrain сохраняет заданный допуск lambda.
Private Sub TestAssertLoadMultiplierTfDefault(ByRef stats As TCapacityTestStats, ByVal methodName As String, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByVal section As CSectionModel, ByVal concrete As Object, ByVal steel As Object, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, _
        ByVal mxOffset As Double, ByVal myOffset As Double, ByVal referenceLambda As Double)
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    cap.ApplySettings settings, units
    cap.SearchMethod = methodName
    cap.SolveByLoadMultiplier section, concrete, steel, nValue, mxValue, myValue, mxOffset, myOffset

    TestAssertTrue stats, "capacity.tfDefaults." & methodName & ".converged", cap.Converged
    TestAssertTrue stats, "capacity.tfDefaults." & methodName & ".physical", TestIsPhysicalLimitState(cap.LimitState)
    TestAssertClose stats, "capacity.tfDefaults." & methodName & ".lambda", cap.LambdaUltimate, referenceLambda, 0.2
    TestAssertTrue stats, "capacity.tfDefaults." & methodName & ".notOvershot", cap.LambdaUltimate < 2# * referenceLambda
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает Bisection/Brent/Secant на одной моментной траектории;
' каждый метод должен выполнить свою ветвь и подтвердить близкую lambda.
Private Sub TestLoadMultiplierSearchMethods(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim bisection As CCapacitySolver
    Dim brent As CCapacitySolver
    Dim secant As CCapacitySolver
    Set bisection = TestRunSearchMethod("Bisection", mesh, rebars)
    Set brent = TestRunSearchMethod("Brent", mesh, rebars)
    Set secant = TestRunSearchMethod("Secant", mesh, rebars)

    TestAssertTrue stats, "search.bisection.converged", bisection.Converged
    TestAssertTrue stats, "search.brent.converged", brent.Converged
    TestAssertTrue stats, "search.secant.converged", secant.Converged
    TestAssertTrue stats, "search.bisection.branch", InStr(1, bisection.DiagnosticLog, "searchMethod=Bisection", vbTextCompare) > 0
    TestAssertTrue stats, "search.brent.branch", InStr(1, brent.DiagnosticLog, "searchMethod=Brent", vbTextCompare) > 0
    TestAssertTrue stats, "search.secant.branch", InStr(1, secant.DiagnosticLog, "searchMethod=Secant", vbTextCompare) > 0
    TestAssertTrue stats, "search.brent.independent", InStr(1, brent.StopReason, "Brent", vbTextCompare) > 0
    TestAssertTrue stats, "search.secant.independent", InStr(1, secant.StopReason, "Secant", vbTextCompare) > 0
    TestAssertClose stats, "search.brent.lambda", brent.LambdaUltimate, bisection.LambdaUltimate, 0.02
    TestAssertClose stats, "search.secant.lambda", secant.LambdaUltimate, bisection.LambdaUltimate, 0.02
    TestAssertEquilibrium stats, "search.brent", brent.LastSolver, -300000#, brent.MxUltimate, 0#
    TestAssertEquilibrium stats, "search.secant", secant.LastSolver, -300000#, secant.MxUltimate, 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет дефект T01: Secant обязан финализировать именно проверенную точку,
' где целевая функция стала нулевой, а не старую нижнюю границу скобки.
Private Sub TestLimitSearchSecantFinalizesCheckedRoot(ByRef stats As TCapacityTestStats)
    On Error GoTo Failed
    Dim stage As String

    stage = "configure"
    Dim problem As CTestLimitSearchProblem
    Set problem = New CTestLimitSearchProblem
    problem.TestConfigure "Secant", 0.75, 0.000000001, 8, 1#, 16#, rkCapacity

    stage = "request"
    Dim request As CLimitSearchRequest
    Set request = New CLimitSearchRequest
    Dim callback As ILimitSearchProblem
    Set callback = problem
    request.InitializeWithProblem callback, rkCapacity, "LoadMultiplier", _
        0#, 1#, 0#, 0#, 0#, 0#

    stage = "execute"
    Dim search As CLoadMultiplierSearch
    Set search = New CLoadMultiplierSearch
    Dim result As CLimitSearchResult
    Set result = search.Execute(request)

    stage = "assert"
    TestAssertTrue stats, "limitSearch.secant.resultObject", Not result Is Nothing
    TestAssertTrue stats, "limitSearch.secant.converged", problem.TestConverged
    TestAssertEquals stats, "limitSearch.secant.method", problem.TestFinalMethod, "Secant"
    TestAssertClose stats, "limitSearch.secant.root", problem.TestFinalLambda, 0.75, 0.000000001
    Exit Sub

Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: limitSearch.secant.runtime; stage=" & stage & _
        "; error=" & CStr(Err.Number) & "; description=" & Err.Description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Если бисекция исчерпала итерации до достижения допуска,
' недостаточно уточнённая граница не должна выдаваться за найденный предел.
Private Sub TestLimitSearchBisectionIterationLimitFails(ByRef stats As TCapacityTestStats)
    On Error GoTo Failed
    Dim stage As String

    stage = "configure"
    Dim problem As CTestLimitSearchProblem
    Set problem = New CTestLimitSearchProblem
    problem.TestConfigure "Bisection", 0.2, 0.001, 1, 1#, 16#, rkCrackFormation

    stage = "request"
    Dim request As CLimitSearchRequest
    Set request = New CLimitSearchRequest
    Dim callback As ILimitSearchProblem
    Set callback = problem
    request.InitializeWithProblem callback, rkCrackFormation, "LoadMultiplier", _
        0#, 1#, 0#, 0#, 0#, 0#, "LambdaMxy"

    stage = "execute"
    Dim search As CLoadMultiplierSearch
    Set search = New CLoadMultiplierSearch
    Dim result As CLimitSearchResult
    Set result = search.Execute(request)

    stage = "assert"
    TestAssertTrue stats, "limitSearch.bisection.resultObject", Not result Is Nothing
    TestAssertTrue stats, "limitSearch.bisection.iterationLimitHit", problem.TestIterationLimitHit
    TestAssertTrue stats, "limitSearch.bisection.notConverged", Not problem.TestConverged
    TestAssertEquals stats, "limitSearch.bisection.limitState", problem.TestLimitState, "NumericalFailure"
    TestAssertTrue stats, "limitSearch.bisection.noFakeLimit", Abs(problem.TestFinalLambda) <= 0.000000001
    Exit Sub

Failed:
    stats.Failed = stats.Failed + 1
    TestAppendLine stats, "FAIL: limitSearch.bisection.runtime; stage=" & stage & _
        "; error=" & CStr(Err.Number) & "; description=" & Err.Description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' нулевая невязка неудачной пробы не означает найденный корень.
' Проверяет сочетание возвращённых признаков: True, NumericalFailure и f = 0.
Private Sub TestLimitSearchNumericalZeroIsNotRoot(ByRef stats As TCapacityTestStats)
    Dim method As Variant, problem As CTestLimitSearchProblem, request As CLimitSearchRequest
    Dim search As CLoadMultiplierSearch, result As CLimitSearchResult, prefix As String
    For Each method In Array("Bisection", "Brent", "Secant")
        Set problem = New CTestLimitSearchProblem
        problem.TestConfigure CStr(method), 0.75, 0.000001, 16, 1#, 16#
        problem.TestFailureAt = 0.75
        Set request = New CLimitSearchRequest
        request.InitializeWithProblem problem, rkCapacity, "LoadMultiplier", 0#, 1#, 0#, 0#, 0#, 0#
        Set search = New CLoadMultiplierSearch
        Set result = search.Execute(request)
        prefix = "limitSearch.failedZero." & CStr(method)
        TestAssertTrue stats, prefix & ".numerical", result.Meta.InternalStatus = rsNumericalFailure
        TestAssertTrue stats, prefix & ".noPoint", Not result.HasLimitPoint
        TestAssertTrue stats, prefix & ".notConverged", Not problem.TestConverged
        TestAssertTrue stats, prefix & ".finiteWork", problem.TestProbeCalls < 30
    Next method
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Отказ всех ненулевых проб должен завершаться за заданный бюджет,
' даже когда исходная нулевая точка допустима и обычной скобки еще нет.
Private Sub TestLimitSearchInitialRecoveryBudget(ByRef stats As TCapacityTestStats)
    Dim method As Variant, kind As Variant, problem As CTestLimitSearchProblem
    Dim request As CLimitSearchRequest, search As CLoadMultiplierSearch, result As CLimitSearchResult
    Dim prefix As String
    For Each kind In Array(rkCapacity, rkCrackFormation)
        For Each method In Array("Bisection", "Brent", "Secant")
            Set problem = New CTestLimitSearchProblem
            problem.TestConfigure CStr(method), 0.75, 0.000001, 8, 1#, 16#, CLng(kind)
            problem.TestFailureAbove = 1E-100
            Set request = New CLimitSearchRequest
            request.InitializeWithProblem problem, CLng(kind), "LoadMultiplier", 0#, 1#, 0#, 0#, 0#, 0#
            Set search = New CLoadMultiplierSearch
            Set result = search.Execute(request)
            prefix = "limitSearch.initialRecovery." & CStr(kind) & "." & CStr(method)
            TestAssertTrue stats, prefix & ".numerical", result.Meta.InternalStatus = rsNumericalFailure
            TestAssertTrue stats, prefix & ".noPoint", Not result.HasLimitPoint
            TestAssertTrue stats, prefix & ".notConverged", Not problem.TestConverged
            TestAssertTrue stats, prefix & ".finiteBudget", problem.TestProbeCalls <= 9
        Next method
    Next kind
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Чередование успешных проб и численных неудач не обновляет бюджет скобки.
' Исчерпание retries не создает физический предел.
Private Sub TestLimitSearchAlternatingRecoveryBudget(ByRef stats As TCapacityTestStats)
    Dim method As Variant, kind As Variant
    Dim problem As CTestLimitSearchProblem, request As CLimitSearchRequest
    Dim search As CLoadMultiplierSearch, result As CLimitSearchResult, prefix As String
    For Each kind In Array(rkCapacity, rkCrackFormation)
        For Each method In Array("Bisection", "Brent", "Secant")
            Set problem = New CTestLimitSearchProblem
            problem.TestConfigure CStr(method), 2#, 0.000001, 8, 2#, 16#, CLng(kind)
            problem.TestFailureAbove = 0.6
            Set request = New CLimitSearchRequest
            request.InitializeWithProblem problem, CLng(kind), "LoadMultiplier", 0#, 1#, 0#, 0#, 0#, 0#
            Set search = New CLoadMultiplierSearch: Set result = search.Execute(request)
            prefix = "limitSearch.alternatingRecovery." & CStr(kind) & "." & CStr(method)
            TestAssertTrue stats, prefix & ".numerical", result.Meta.InternalStatus = rsNumericalFailure
            TestAssertTrue stats, prefix & ".noPoint", Not result.HasLimitPoint And Not result.Converged
            TestAssertTrue stats, prefix & ".intermediateSuccess", problem.TestAcceptableProbeCalls >= 3
            TestAssertTrue stats, prefix & ".finiteBudget", problem.TestProbeCalls <= 20
            TestAssertTrue stats, prefix & ".fallbackOnce", problem.TestFallbackCalls = 1
        Next method
    Next kind
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' быстрый контроль бюджета скобки без тяжелых расчетов сечений.
Public Function RunBracketRecoveryRegressionTests() As String
    Dim stats As TCapacityTestStats
    TestLimitSearchNumericalZeroIsNotRoot stats
    TestLimitSearchInitialRecoveryBudget stats
    TestLimitSearchAlternatingRecoveryBudget stats
    TestSmallLambdaBracketRecovery stats
    TestAppendLine stats, "TOTAL_BRACKET_RECOVERY: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunBracketRecoveryRegressionTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' быстрый контроль неудачной пробы и недостижимой точности.
Public Function RunLimitSearchFailureRegressionTests() As String
    Dim stats As TCapacityTestStats
    TestLimitSearchNumericalZeroIsNotRoot stats
    TestLimitSearchInitialRecoveryBudget stats
    TestAudit03RealCapacityPrecision stats
    TestRectanglesCapacityLoadPathSmoke stats
    TestUltimateMomentTranslationInvariant stats
    TestUltimateStrainToleranceNoHiddenFloor stats
    TestUltimateNumericalContextCache stats
    TestAppendLine stats, "TOTAL_LIMIT_FAILURE_REGRESSION: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunLimitSearchFailureRegressionTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет все пользовательские траектории LoadPath на всех
' доступных способах поиска несущей способности. Тест намеренно работает
' на уровне CCapacitySolver: batch уже переводит пользовательские строки
' lambda*Mx/lambda*My/... в универсальную форму Offset + lambda*Base.
Private Sub TestCapacityLoadPathMethodMatrix(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    TestCheckCapacityLoadPathMethods stats, "mx", mesh, rebars, -120000#, 0#, 0#, -2000000#, 0#, 0#
    TestCheckCapacityLoadPathMethods stats, "my", mesh, rebars, -120000#, 0#, 0#, 0#, 0#, -2000000#
    TestCheckCapacityLoadPathMethods stats, "mxy", mesh, rebars, -120000#, 0#, 0#, -2000000#, 0#, -1200000#
    TestCheckCapacityLoadPathMethods stats, "n", mesh, rebars, 0#, -50000#, 0#, 0#, 0#, 0#
    TestCheckCapacityLoadPathMethods stats, "nmxy", mesh, rebars, 0#, -50000#, 0#, -1200000#, 0#, -800000#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет вырожденные, но допустимые lambda-траектории: в Base-векторе
' могут быть нулевые компоненты, если хотя бы одна компонента нагрузки реально
' масштабируется. Это защищает общий контракт solver-а Offset + lambda*Base:
' lambda*Mx не обязан иметь N, lambda*Mxy может содержать только один момент, а lambda*NMxy
' может фактически свестись к чистому N, чистому Mx или чистому My.
Private Sub TestCapacityLoadPathZeroComponentMatrix(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    TestCheckCapacityLoadPathMethods stats, "zero.mxNoN", mesh, rebars, 0#, 0#, 0#, -2000000#, 0#, 0#
    TestCheckCapacityLoadPathMethods stats, "zero.myNoN", mesh, rebars, 0#, 0#, 0#, 0#, 0#, -2000000#
    TestCheckCapacityLoadPathMethods stats, "zero.mxyOnlyMx", mesh, rebars, 0#, 0#, 0#, -2000000#, 0#, 0#
    TestCheckCapacityLoadPathMethods stats, "zero.mxyOnlyMy", mesh, rebars, 0#, 0#, 0#, 0#, 0#, -2000000#
    TestCheckCapacityLoadPathMethods stats, "zero.nmxyOnlyN", mesh, rebars, 0#, -50000#, 0#, 0#, 0#, 0#
    TestCheckCapacityLoadPathMethods stats, "zero.nmxyOnlyMx", mesh, rebars, 0#, 0#, 0#, -2000000#, 0#, 0#
    TestCheckCapacityLoadPathMethods stats, "zero.nmxyOnlyMy", mesh, rebars, 0#, 0#, 0#, 0#, 0#, -2000000#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет пользовательское Г-сечение именно на осевой траектории lambda*N.
' В этой задаче важно, что N приложена в бетонном центре тяжести: после
' переноса к координатам расчетных элементов внутри solver-а появляются
' связанные Mx/My, хотя пользователь масштабирует только продольную силу.
Private Sub TestRectanglesCapacityLoadPathSmoke(ByRef stats As TCapacityTestStats)
    Dim section As CSectionModel
    Set section = TestRectanglesCapacitySection()

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete section

    Dim tensionLoad As CSectionLoadState
    Set tensionLoad = New CSectionLoadState
    tensionLoad.Initialize 200# * 9806.65, 0#, 0#, props.CentroidX, props.CentroidY
    TestCheckRectanglesCapacityPathMethods stats, "rectangles.n.tension", section, 0#, tensionLoad.N, _
        0#, tensionLoad.AxialMxAboutPoint(0#), 0#, tensionLoad.AxialMyAboutPoint(0#), True

    Dim compressionLoad As CSectionLoadState
    Set compressionLoad = New CSectionLoadState
    compressionLoad.Initialize -200# * 9806.65, 0#, 0#, props.CentroidX, props.CentroidY
    TestCheckRectanglesCapacityPathMethods stats, "rectangles.n.compression", section, 0#, compressionLoad.N, _
        0#, compressionLoad.AxialMxAboutPoint(0#), 0#, compressionLoad.AxialMyAboutPoint(0#), True

    ' Пользовательский сценарий из книги: Г-сечение, N задана относительно
    ' бетонного центра тяжести, а предельная способность ищется по lambda*Mx.
    ' Здесь обязана включаться быстрая моментная постановка
    ' UltimateStrain: N постоянна, направление Mx/My сохраняется.
    TestCheckRectanglesMomentUltimatePath stats, "rectangles.moment.mx.userCase", section, _
        -200# * 9806.65, 50# * 9806.65 * 1000#, 0#, props.CentroidX, props.CentroidY
    TestCheckRectanglesMomentUltimatePath stats, "rectangles.moment.my.userCase", section, _
        -200# * 9806.65, 0#, 50# * 9806.65 * 1000#, props.CentroidX, props.CentroidY
    TestCheckRectanglesMomentUltimatePath stats, "rectangles.moment.mxy.userCase", section, _
        -200# * 9806.65, 50# * 9806.65 * 1000#, 25# * 9806.65 * 1000#, props.CentroidX, props.CentroidY
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет моментную ветку UltimateStrain на несимметричном Г-сечении.
' Внутри solver-а момент от N добавляется как постоянный offset, а
' пользовательский момент масштабируется через lambda. Такой тест защищает
' рабочую моментную постановку от случайного ухода в общий load-path residual.
Private Sub TestCheckRectanglesMomentUltimatePath(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal section As CSectionModel, ByVal nValue As Double, _
        ByVal userMxBase As Double, ByVal userMyBase As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double)
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap

    Dim mxOffset As Double
    Dim myOffset As Double
    Dim loadState As CSectionLoadState
    Set loadState = New CSectionLoadState
    loadState.Initialize nValue, userMxBase, userMyBase, referenceX, referenceY
    mxOffset = loadState.AxialMxAboutPoint(0#)
    myOffset = loadState.AxialMyAboutPoint(0#)

    cap.SolveByUltimateLoadPath section, TestProvisionalConcrete(), TestProvisionalSteel(), _
        nValue, 0#, mxOffset, userMxBase, myOffset, userMyBase

    TestAppendLine stats, "INFO: " & prefix & _
        "; status=" & cap.LimitState & _
        "; converged=" & CStr(cap.Converged) & _
        "; solutionMethod=" & cap.SolutionMethod & _
        "; lambda=" & TestFormatNumberInvariant(cap.LambdaUltimate)
    TestAssertTrue stats, prefix & ".converged", cap.Converged
    TestAssertEquals stats, prefix & ".solutionMethod", cap.SolutionMethod, "UltimateStrain"
    TestAssertTrue stats, prefix & ".lambda", cap.LambdaUltimate > 0#
    TestAssertTrue stats, prefix & ".physical", TestIsPhysicalLimitState(cap.LimitState)
    TestAssertLoadComponent stats, prefix & ".equilibrium.N", cap.LastSolver.Nint, nValue, 10#, 0.0001
    TestAssertLoadComponent stats, prefix & ".equilibrium.Mx", cap.LastSolver.Mxint, _
        mxOffset + cap.LambdaUltimate * userMxBase, 10000#, 0.0002
    TestAssertLoadComponent stats, prefix & ".equilibrium.My", cap.LastSolver.Myint, _
        myOffset + cap.LambdaUltimate * userMyBase, 10000#, 0.0002
    TestAssertClose stats, prefix & ".mxUltimate.user", cap.MxUltimate, cap.LambdaUltimate * userMxBase, 20000#
    TestAssertClose stats, prefix & ".myUltimate.user", cap.MyUltimate, cap.LambdaUltimate * userMyBase, 20000#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Запускает одинаковый Rectangles load path всеми разрешенными стратегиями поиска.
' Для силовой траектории явно передает принятый переход к LoadMultiplier.
Private Sub TestCheckRectanglesCapacityPathMethods(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal section As CSectionModel, ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        Optional ByVal allowForcePathFallback As Boolean = False)
    TestCheckRectanglesCapacityPathMethod stats, prefix & ".auto", "Auto", "", section, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowForcePathFallback
    TestCheckRectanglesCapacityPathMethod stats, prefix & ".ultimate", "UltimateStrain", "", section, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowForcePathFallback
    TestCheckRectanglesCapacityPathMethod stats, prefix & ".bisection", "LoadMultiplier", "Bisection", section, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowForcePathFallback
    TestCheckRectanglesCapacityPathMethod stats, prefix & ".brent", "LoadMultiplier", "Brent", section, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowForcePathFallback
    TestCheckRectanglesCapacityPathMethod stats, prefix & ".secant", "LoadMultiplier", "Secant", section, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowForcePathFallback
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет выбранный поиск на готовом сечении Rectangles и проверяет физический
' предел. Усилия итогового НДС отдельно проверяются на заданной траектории нагрузки.
Private Sub TestCheckRectanglesCapacityPathMethod(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal methodName As String, ByVal searchMethod As String, ByVal section As CSectionModel, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        Optional ByVal allowForcePathFallback As Boolean = False)
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    If Len(searchMethod) > 0 Then cap.SearchMethod = searchMethod

    Select Case methodName
        Case "Auto"
            If allowForcePathFallback Then
                cap.SolveByLoadPathMultiplier section, TestProvisionalConcrete(), TestProvisionalSteel(), _
                    nOffset, nBase, mxOffset, mxBase, myOffset, myBase, True
            Else
                cap.SolveByAutoLoadPath section, TestProvisionalConcrete(), TestProvisionalSteel(), _
                    nOffset, nBase, mxOffset, mxBase, myOffset, myBase
            End If
        Case "UltimateStrain"
            If allowForcePathFallback Then
                cap.SolveByLoadPathMultiplier section, TestProvisionalConcrete(), TestProvisionalSteel(), _
                    nOffset, nBase, mxOffset, mxBase, myOffset, myBase, True
            Else
                cap.SolveByUltimateLoadPath section, TestProvisionalConcrete(), TestProvisionalSteel(), _
                    nOffset, nBase, mxOffset, mxBase, myOffset, myBase
            End If
        Case Else
            cap.SolveByLoadPathMultiplier section, TestProvisionalConcrete(), TestProvisionalSteel(), _
                nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowForcePathFallback
    End Select

    TestAppendLine stats, "INFO: " & prefix & _
        "; status=" & cap.LimitState & _
        "; converged=" & CStr(cap.Converged) & _
        "; solutionMethod=" & cap.SolutionMethod & _
        "; lambda=" & TestFormatNumberInvariant(cap.LambdaUltimate)
    TestAssertTrue stats, prefix & ".converged", cap.Converged
    TestAssertTrue stats, prefix & ".lambda", cap.LambdaUltimate > 0#
    TestAssertTrue stats, prefix & ".physical", TestIsPhysicalLimitState(cap.LimitState)
    TestAssertLoadPathResult stats, prefix, cap, nOffset, nBase, mxOffset, mxBase, myOffset, myBase
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Прогоняет чистый Nult через разные значения BaseLoadSteps. Настройка не
' является физическим параметром, но для осевой траектории на плато диаграммы
' она может заметно влиять на устойчивость отдельных probe-точек.
Private Sub TestNultBaseLoadStepsSensitivity(ByRef stats As TCapacityTestStats)
    Dim section As CSectionModel
    Set section = TestRectanglesCapacitySection()

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete section

    Dim tensionLoad As CSectionLoadState
    Set tensionLoad = New CSectionLoadState
    tensionLoad.Initialize 200# * 9806.65, 0#, 0#, props.CentroidX, props.CentroidY
    TestCheckNultBaseLoadSteps stats, "tension", section, tensionLoad.N, _
        tensionLoad.AxialMxAboutPoint(0#), tensionLoad.AxialMyAboutPoint(0#)

    Dim compressionLoad As CSectionLoadState
    Set compressionLoad = New CSectionLoadState
    compressionLoad.Initialize -200# * 9806.65, 0#, 0#, props.CentroidX, props.CentroidY
    TestCheckNultBaseLoadSteps stats, "compression", section, compressionLoad.N, _
        compressionLoad.AxialMxAboutPoint(0#), compressionLoad.AxialMyAboutPoint(0#)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторяет осевой путь с 1/2/4/8 ступенями внутренних решений равновесия.
' Один шаг диагностируется отдельно; успешность обязательна для остальных вариантов.
Private Sub TestCheckNultBaseLoadSteps(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal section As CSectionModel, ByVal nBase As Double, _
        ByVal mxBase As Double, ByVal myBase As Double)
    Dim steps As Variant
    steps = Array(1, 2, 4, 8)

    Dim i As Long
    For i = LBound(steps) To UBound(steps)
        Dim cap As CCapacitySolver
        Set cap = New CCapacitySolver
        TestConfigureCapacity cap
        cap.SolverBaseLoadSteps = CLng(steps(i))
        cap.SolveByLoadPathMultiplier section, TestProvisionalConcrete(), TestProvisionalSteel(), _
            0#, nBase, 0#, mxBase, 0#, myBase, True

        TestAppendLine stats, "INFO: capacity.nult.baseLoadSteps." & prefix & "." & CStr(steps(i)) & _
            "; converged=" & CStr(cap.Converged) & _
            "; state=" & cap.LimitState & _
            "; lambda=" & TestFormatNumberInvariant(cap.LambdaUltimate) & _
            "; iterations=" & CStr(cap.Iterations) & _
            "; retries=" & CStr(cap.TestRetryCount)
        If CLng(steps(i)) > 1 Then
            TestAssertTrue stats, "capacity.nult.baseLoadSteps." & prefix & "." & CStr(steps(i)) & ".converged", cap.Converged
            TestAssertTrue stats, "capacity.nult.baseLoadSteps." & prefix & "." & CStr(steps(i)) & ".lambda", cap.LambdaUltimate > 0#
        End If
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает стратегии поиска на одной сетке, арматуре и Offset/Base-траектории.
' Численные методы LoadMultiplier проверяются каждый отдельным запуском.
Private Sub TestCheckCapacityLoadPathMethods(ByRef stats As TCapacityTestStats, ByVal pathName As String, _
        ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double)
    TestCheckCapacityLoadPathMethod stats, pathName & ".auto", "Auto", "", mesh, rebars, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase
    TestCheckCapacityLoadPathMethod stats, pathName & ".ultimate", "UltimateStrain", "", mesh, rebars, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase
    TestCheckCapacityLoadPathMethod stats, pathName & ".bisection", "LoadMultiplier", "Bisection", mesh, rebars, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase
    TestCheckCapacityLoadPathMethod stats, pathName & ".brent", "LoadMultiplier", "Brent", mesh, rebars, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase
    TestCheckCapacityLoadPathMethod stats, pathName & ".secant", "LoadMultiplier", "Secant", mesh, rebars, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Собирает CSectionModel и запускает только запрошенный метод поиска предела.
' При отказе сохраняет диагностику; при успехе проверяет равновесие и путь нагрузки.
Private Sub TestCheckCapacityLoadPathMethod(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal methodName As String, ByVal searchMethod As String, _
        ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double)
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    If Len(searchMethod) > 0 Then cap.SearchMethod = searchMethod

    Select Case methodName
        Case "Auto"
            cap.SolveByAutoLoadPath TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), _
                nOffset, nBase, mxOffset, mxBase, myOffset, myBase
        Case "UltimateStrain"
            cap.SolveByUltimateLoadPath TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), _
                nOffset, nBase, mxOffset, mxBase, myOffset, myBase
        Case Else
            cap.SolveByLoadPathMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), _
                nOffset, nBase, mxOffset, mxBase, myOffset, myBase
    End Select

    TestAssertTrue stats, "capacity.pathMatrix." & prefix & ".converged", cap.Converged
    TestAssertTrue stats, "capacity.pathMatrix." & prefix & ".lambda", cap.LambdaUltimate > 0#
    TestAssertTrue stats, "capacity.pathMatrix." & prefix & ".physical", TestIsPhysicalLimitState(cap.LimitState)
    If Not cap.Converged Then
        TestAppendLine stats, "DIAG: capacity.pathMatrix." & prefix & _
            "; limitState=" & cap.LimitState & _
            "; stopReason=" & cap.StopReason & _
            "; log=" & Replace(cap.DiagnosticLog, vbCrLf, " | ")
    End If
    TestAssertCapacitySolutionMethod stats, "capacity.pathMatrix." & prefix & ".solutionMethod", cap, methodName, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase
    If cap.Converged Then
        TestAssertLoadPathResult stats, "capacity.pathMatrix." & prefix, cap, _
            nOffset, nBase, mxOffset, mxBase, myOffset, myBase
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит фиксированное двухпрямоугольное сечение с арматурой по четырем граням.
' Геометрия, шаг сетки и диаметры независимы от пользовательского Config.
Private Function TestRectanglesCapacitySection() As CSectionModel
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim row As Variant, rows As Variant, mesh As TSectionMeshOptions, builder As CSectionBuilder
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "RectanglesCapacityTest"
    Set path = facade.AdjacentRectangles("Boundary", Array(Array(0#, 0#, 600#, 250#), Array(0#, 250#, 250#, 550#)))
    facade.Outer "OUTER", path.ID
    rows = Array(Array("H1.Left", 40#, 250#, 40#, 800#, 5), Array("H1.Right", 210#, 800#, 210#, 250#, 5), _
        Array("B1.Top", 0#, 760#, 250#, 760#, 2), Array("B1.Bottom", 250#, 290#, 0#, 290#, 2), _
        Array("H2.Left", 40#, 0#, 40#, 250#, 2), Array("H2.Right", 560#, 250#, 560#, 0#, 2), _
        Array("B2.Top", 0#, 210#, 600#, 210#, 5), Array("B2.Bottom", 600#, 40#, 0#, 40#, 5))
    For Each row In rows
        Set path = facade.LinePath(CStr(row(0)), CDbl(row(1)), CDbl(row(2)), CDbl(row(3)), CDbl(row(4)))
        Set options = New CRebarRuleOptions
        options.StartSetback = 80#: options.EndSetback = 80#
        options.PositionOrder = FromPathMiddle: options.Publication = PublishPerSourcePosition
        options.Coincidence = KeepPublishedOwner: options.AnnotationGroup = CStr(row(0))
        facade.RebarByCount CStr(row(0)) & ".Main", path.ID, CLng(row(5)), 32#, options, "A400"
    Next row
    mesh.StepX = 50#: mesh.StepY = 50#: mesh.BoundarySubdivisions = 1: mesh.MaterialID = 1
    Set builder = New CSectionBuilder: Set TestRectanglesCapacitySection = builder.Build(facade.Finish, mesh)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выводится название выбранного метода UltimateStrain или LoadMultiplier.
' Внутреннее имя общей процедуры поиска не подменяет пользовательскую настройку.
Private Sub TestAssertCapacitySolutionMethod(ByRef stats As TCapacityTestStats, ByVal name As String, _
        ByVal cap As CCapacitySolver, ByVal methodName As String, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double)
    Dim isPureAxial As Boolean
    isPureAxial = Abs(nBase) > 0.000000001 And _
        Abs(nOffset) <= 0.000000001 And _
        Abs(mxOffset) <= 0.000000001 And Abs(mxBase) <= 0.000000001 And _
        Abs(myOffset) <= 0.000000001 And Abs(myBase) <= 0.000000001

    If methodName = "LoadMultiplier" Then
        TestAssertEquals stats, name, cap.SolutionMethod, "LoadMultiplier"
    ElseIf methodName = "UltimateStrain" Then
        If isPureAxial Then
            TestAssertEquals stats, name, cap.SolutionMethod, "LoadMultiplier"
        Else
            TestAssertEquals stats, name, cap.SolutionMethod, "UltimateStrain"
        End If
    Else
        TestAssertTrue stats, name, cap.SolutionMethod = "UltimateStrain" Or cap.SolutionMethod = "LoadMultiplier"
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимо вычисляет N, Mx и My по Offset + lambda*Base и сравнивает с НДС.
' Предельная сила Nult не выводится, если N на выбранном пути не меняется.
Private Sub TestAssertLoadPathResult(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal cap As CCapacitySolver, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double)
    Dim expectedN As Double
    Dim expectedMx As Double
    Dim expectedMy As Double
    expectedN = nOffset + cap.LambdaUltimate * nBase
    expectedMx = mxOffset + cap.LambdaUltimate * mxBase
    expectedMy = myOffset + cap.LambdaUltimate * myBase

    TestAssertLoadComponent stats, prefix & ".equilibrium.N", cap.LastSolver.Nint, expectedN, 10#, 0.0001
    TestAssertLoadComponent stats, prefix & ".equilibrium.Mx", cap.LastSolver.Mxint, expectedMx, 10000#, 0.0002
    TestAssertLoadComponent stats, prefix & ".equilibrium.My", cap.LastSolver.Myint, expectedMy, 10000#, 0.0002
    If Abs(nBase) > 0.000000001 Then
        TestAssertClose stats, prefix & ".nUltimate", cap.NUltimate, expectedN, 20#
    Else
        TestAssertClose stats, prefix & ".nUltimateBlank", cap.NUltimate, 0#, 0#
    End If
    TestAssertClose stats, prefix & ".mxUltimate", cap.MxUltimate, expectedMx, 20000#
    TestAssertClose stats, prefix & ".myUltimate", cap.MyUltimate, expectedMy, 20000#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет пустое/неизвестное имя Search и контролируемый отказ Secant
' при малом бюджете: ни один сценарий не переключается молча на другой метод.
Private Sub TestSearchMethodInputErrors(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim emptyMethod As CCapacitySolver
    Set emptyMethod = New CCapacitySolver
    TestConfigureCapacity emptyMethod
    emptyMethod.SearchMethod = ""
    emptyMethod.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -300000#, -10000000#, 0#

    Dim invalidMethod As CCapacitySolver
    Set invalidMethod = New CCapacitySolver
    TestConfigureCapacity invalidMethod
    invalidMethod.SearchMethod = "FalseMethod"
    invalidMethod.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -300000#, -10000000#, 0#

    Dim secantFail As CCapacitySolver
    Set secantFail = New CCapacitySolver
    TestConfigureCapacity secantFail
    secantFail.SearchMethod = "Secant"
    secantFail.SolverMaxIterations = 1
    secantFail.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -300000#, -10000000#, 0#

    TestAssertTrue stats, "search.empty.inputError", Not emptyMethod.Converged And emptyMethod.LimitState = "InvalidInput"
    TestAssertTrue stats, "search.invalid.inputError", Not invalidMethod.Converged And invalidMethod.LimitState = "InvalidInput"
    TestAssertTrue stats, "search.secant.controlledFailure", Not secantFail.Converged And _
        (secantFail.LimitState = "NumericalFailure" Or secantFail.LimitState = "InvalidInput")
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что lambda=0 разделяет физический BaseFail и численную ошибку.
' Search сообщает только состояние первой точки, а внешний статус обязан
' формироваться из InternalStatus/ResultCode, а не из текста комментария.
Private Sub TestInitialLambdaFailureStatusMapping(ByRef stats As TCapacityTestStats)
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy

    Dim numericalCap As CCapacitySolver
    Set numericalCap = New CCapacitySolver
    numericalCap.LimitSearchHandleCapacityInitialFailure "NumericalFailure"

    Dim numericalResult As CLimitSearchResult
    Set numericalResult = TestAudit02CapacitySnapshot(numericalCap, "LoadMultiplier")
    TestAssertEquals stats, "capacity.initialLambda.numerical.external", _
        policy.ExternalStatus(numericalResult.Meta), "NumFail"

    Dim physicalCap As CCapacitySolver
    Set physicalCap = New CCapacitySolver
    physicalCap.LimitSearchHandleCapacityInitialFailure "ConcreteStrainLimit"

    Dim physicalResult As CLimitSearchResult
    Set physicalResult = TestAudit02CapacitySnapshot(physicalCap, "LoadMultiplier")
    TestAssertEquals stats, "capacity.initialLambda.physical.external", _
        policy.ExternalStatus(physicalResult.Meta), "BaseFail"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет три замера каждого поискового метода на одинаковых данных.
' Единичный замер не служит доказательством ускорения.
Private Sub TestSearchMethodPerformanceComparison(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    TestAppendSearchPerformance stats, "Bisection", mesh, rebars
    TestAppendSearchPerformance stats, "Brent", mesh, rebars
    TestAppendSearchPerformance stats, "Secant", mesh, rebars
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Решает общий контрольный моментный путь указанным Bisection/Brent/Secant
' и возвращает solver с фактическими счетчиками и конечной диагностикой.
Private Function TestRunSearchMethod(ByVal methodName As String, ByVal mesh As CFiberMeshBuilder, _
        ByVal rebars As CRebarLayout) As CCapacitySolver
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    cap.SearchMethod = methodName
    cap.SolveByLoadMultiplier TestBuildGeneratedSectionModel(mesh, rebars), TestProvisionalConcrete(), TestProvisionalSteel(), -300000#, -10000000#, 0#
    Set TestRunSearchMethod = cap
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет повторные измерения метода на свежем solver-е и добавляет
' среднее время вместе с результатом; не изменяет настройки ради быстродействия.
Private Sub TestAppendSearchPerformance(ByRef stats As TCapacityTestStats, ByVal methodName As String, _
        ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout)
    Dim runs As Long
    runs = 3
    Dim totalSec As Double
    Dim cap As CCapacitySolver
    Dim i As Long
    For i = 1 To runs
        Dim t0 As Double
        t0 = Timer
        Set cap = TestRunSearchMethod(methodName, mesh, rebars)
        totalSec = totalSec + (Timer - t0)
    Next i

    TestAssertTrue stats, "search.perf." & methodName & ".converged", cap.Converged
    TestAppendLine stats, "PERF_CAPACITY_SEARCH|" & methodName & _
        "|lambda=" & TestFormatNumberInvariant(cap.LambdaUltimate) & _
        "|iterations=" & CStr(cap.Iterations) & _
        "|functionEvaluations=" & CStr(cap.Iterations) & _
        "|avgSec=" & TestFormatNumberInvariant(totalSec / runs) & _
        "|residual=" & TestFormatNumberInvariant(cap.MomentEquilibriumResidual) & _
        "|converged=" & CStr(cap.Converged) & _
        "|status=" & cap.LimitState
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет результаты и стоимость поиска для сравнения стратегий без пересчета.
' Указывает найденную lambda, управляющий элемент и число обращений к решателю.
Private Sub TestAppendComparison(ByRef stats As TCapacityTestStats, ByVal name As String, _
        ByVal cap As CCapacitySolver, ByVal elapsedSec As Double)
    TestAppendLine stats, "COMPARE|" & name & _
        "|converged=" & CStr(cap.Converged) & _
        "|lambda=" & TestFormatNumberInvariant(cap.LambdaUltimate) & _
        "|limitState=" & cap.LimitState & _
        "|criticalStrain=" & TestFormatNumberInvariant(cap.TestCriticalStrain) & _
        "|criticalElement=" & cap.TestCriticalElement & _
        "|solverCalls=" & CStr(cap.Iterations) & _
        "|elapsedSec=" & TestFormatNumberInvariant(elapsedSec)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задаёт фиксированные пределы и бюджет контрольных задач несущей способности.
' Это тестовые данные, а не нормативные значения по умолчанию рабочего расчёта.
Private Sub TestConfigureCapacity(ByVal cap As CCapacitySolver)
    cap.ConcreteCompressionLimit = -0.0015
    cap.SteelStrainLimit = 0.00175
    cap.LambdaTolerance = 0.01
    cap.MaxLambda = 64#
    cap.InitialLambdaStep = 1#
    cap.MaxRetries = 4
    cap.SolverBaseLoadSteps = 8
    cap.SolverMaxIterations = 60
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Настраивает независимую проверку равновесия во внутренних единицах N и N*mm.
' Значения не читаются из Config и не меняют допуски проверяемого поискового метода.
Private Sub TestConfigureSectionSolver(ByVal solver As CSectionSolver)
    solver.LoadSteps = 8
    solver.MaxIterations = 80
    solver.ToleranceN = 5#
    solver.ToleranceMx = 5000#
    solver.ToleranceMy = 5000#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создает фиксированный сжатый бетон без растянутой ветви и расширения.
Private Function TestProvisionalConcrete() As CMaterialDiagram
    Dim concrete As CMaterialDiagram
    Set concrete = New CMaterialDiagram
    concrete.TestInitialize -0.0015, -15.5, -0.0035, -15.5
    Set TestProvisionalConcrete = concrete
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создает физическую TwoLine-диаграмму с растяжением через обычный material provider.
' Параметры берутся из fixtures ниже, а не из листа Config.
Private Function TestProvisionalConcreteWithTension() As CMaterialDiagram
    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.TestInitializeFromParameters TestConcreteParameters(), TestSteelParameters(), "TwoLine", "UseDiagram"
    Set TestProvisionalConcreteWithTension = provider.ConcreteMaterial(cpStrength)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задаёт два набора характеристик бетона и сопротивление продольным трещинам
' для воспроизводимой проверки. Нормативное обоснование не заменяется этими числами.
Private Function TestConcreteParameters() As CConcreteMaterialParameters
    Dim parameters As CConcreteMaterialParameters
    Set parameters = New CConcreteMaterialParameters
    parameters.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    Set TestConcreteParameters = parameters
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задает фиксированные модули и сопротивления арматуры обоих расчетных наборов.
Private Function TestSteelParameters() As CSteelMaterialParameters
    Dim parameters As CSteelMaterialParameters
    Set parameters = New CSteelMaterialParameters
    parameters.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#
    Set TestSteelParameters = parameters
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает симметричную физическую диаграмму арматуры с постоянным плато.
Private Function TestProvisionalSteel() As CMaterialDiagram
    Dim steel As CMaterialDiagram
    Set steel = New CMaterialDiagram
    steel.TestInitialize 0.00175, 350#, 0.025
    Set TestProvisionalSteel = steel
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает линейный бетон для аналитически проверяемых задач равновесия.
Private Function TestLinearConcrete() As CTestLinearConcreteMaterial
    Dim concrete As CTestLinearConcreteMaterial
    Set concrete = New CTestLinearConcreteMaterial
    concrete.TestInitialize 32500#
    Set TestLinearConcrete = concrete
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает линейную арматуру; предел для поиска задается отдельно в Capacity.
Private Function TestLinearSteel() As CTestLinearSteelMaterial
    Dim steel As CTestLinearSteelMaterial
    Set steel = New CTestLinearSteelMaterial
    steel.TestInitialize 200000#
    Set TestLinearSteel = steel
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Готовит симметричный прямоугольный fixture и возвращает его сетку и арматуру.
' Центр находится в начале координат, четыре одинаковых стержня зеркальны по осям.
Private Sub TestPrepareSymmetricSection(ByVal width As Double, ByVal height As Double, ByVal stepSize As Double, _
        ByVal xAbs As Double, ByVal yAbs As Double, ByRef mesh As CFiberMeshBuilder, ByRef rebars As CRebarLayout)
    Dim geom As CGeometryRegion
    Set geom = TestRectangleGeometry(width, height)
    Set mesh = TestBuildMesh(geom, stepSize)
    Set rebars = TestSymmetricRebars(geom, xAbs, yAbs)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Расставляет четыре стержня диаметром 20 мм в зеркальных точках (+/-x, +/-y).
' Каждый стержень проходит обычную проверку принадлежности геометрии.
Private Function TestSymmetricRebars(ByVal geom As CGeometryRegion, ByVal xAbs As Double, ByVal yAbs As Double) As CRebarLayout
    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -xAbs, -yAbs, 20#, 0#, "A400", "", geom
    rebars.AddBar "B2", xAbs, -yAbs, 20#, 0#, "A400", "", geom
    rebars.AddBar "B3", -xAbs, yAbs, 20#, 0#, "A400", "", geom
    rebars.AddBar "B4", xAbs, yAbs, 20#, 0#, "A400", "", geom
    Set TestSymmetricRebars = rebars
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает прямоугольник без скруглений, поворота и смещения для fixtures.
Private Function TestRectangleGeometry(ByVal width As Double, ByVal height As Double) As CGeometryRegion
    Dim geom As CGeometryRegion
    Set geom = TestFixtureBox(width, height)
    Set TestRectangleGeometry = geom
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает тестовую несимметричную геометрию с заметным Ixy для проверки,
' что capacity-решатель учитывает связанную кривизну в общей постановке N+Mx+My.
Private Function TestCapacityAsymmetricGeometry() As CGeometryRegion
    Dim geom As CGeometryRegion
    Set geom = TestFixtureBoxes(Array(Array(0#, 0#, 600#, 250#), Array(0#, 250#, 250#, 550#)))
    Set TestCapacityAsymmetricGeometry = geom
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит тестовую сетку общей бетонной области CGeometryRegion с одинаковым
' шагом по осям; затем сетка и арматура объединяются в CSectionModel.
Private Function TestBuildMesh(ByVal geom As CGeometryRegion, ByVal stepSize As Double) As CFiberMeshBuilder
    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, stepSize, stepSize, 1
    Set TestBuildMesh = mesh
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Отличает физические деформационные критерии от технического исхода поиска.
' Это проверка диагностического поля теста, не назначение пользовательского статуса.
Private Function TestIsPhysicalLimitState(ByVal state As String) As Boolean
    TestIsPhysicalLimitState = (state = "ConcreteStrainLimit" Or state = "ConcreteTensionStrainLimit" Or _
        state = "SteelStrainLimit")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет все компоненты внутренних усилий с независимо заданной нагрузкой.
' Нулевая сила и нулевой момент проверяются по своим абсолютным допускам.
Private Sub TestAssertEquilibrium(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal solver As CSectionSolver, ByVal n As Double, ByVal mx As Double, ByVal my As Double)
    TestAssertLoadComponent stats, prefix & ".N", solver.Nint, n, 10#, 0.0001
    TestAssertLoadComponent stats, prefix & ".Mx", solver.Mxint, mx, 10000#, 0.0001
    TestAssertLoadComponent stats, prefix & ".My", solver.Myint, my, 10000#, 0.0001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Не делит на почти нулевой эталон: выбирает абсолютную или относительную ошибку.
Private Sub TestAssertLoadComponent(ByRef stats As TCapacityTestStats, ByVal name As String, _
        ByVal actual As Double, ByVal expected As Double, ByVal zeroTolerance As Double, _
        ByVal relTolerance As Double)
    If Abs(expected) <= zeroTolerance Then
        TestAssertClose stats, name, actual, expected, zeroTolerance
    Else
        TestAssertRelative stats, name, actual, expected, relTolerance
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Группа assertions ведет счет и протоколирует условия и фактические отклонения.
' Допуски численных сравнений задаются самим контрольным случаем.
Private Sub TestAssertTrue(ByRef stats As TCapacityTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        TestAppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        TestAppendLine stats, "FAIL: " & name
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает строковый результат без нормализации; при отказе сохраняет обе
' строки, чтобы отличить неверный статус или код от численной ошибки теста.
Private Sub TestAssertEquals(ByRef stats As TCapacityTestStats, ByVal name As String, ByVal actual As String, ByVal expected As String)
    If actual = expected Then
        stats.Passed = stats.Passed + 1
        TestAppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        TestAppendLine stats, "FAIL: " & name & "; actual=" & actual & "; expected=" & expected
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает по абсолютному допуску с записью actual/expected и величины ошибки.
Private Sub TestAssertClose(ByRef stats As TCapacityTestStats, ByVal name As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal tolerance As Double)
    Dim diff As Double
    diff = Abs(actual - expected)
    If diff <= tolerance Then
        stats.Passed = stats.Passed + 1
        TestAppendLine stats, "OK: " & name & "; actual=" & TestFormatNumberInvariant(actual) & _
            "; expected=" & TestFormatNumberInvariant(expected) & "; absDiff=" & TestFormatNumberInvariant(diff)
    Else
        stats.Failed = stats.Failed + 1
        TestAppendLine stats, "FAIL: " & name & "; actual=" & TestFormatNumberInvariant(actual) & _
            "; expected=" & TestFormatNumberInvariant(expected) & "; absDiff=" & TestFormatNumberInvariant(diff)
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет относительное отклонение; для почти нулевого эталона избегает деления.
Private Sub TestAssertRelative(ByRef stats As TCapacityTestStats, ByVal name As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal relTolerance As Double)
    Dim relDiff As Double
    If Abs(expected) <= 0.000000001 Then
        relDiff = Abs(actual - expected)
    Else
        relDiff = Abs((actual - expected) / expected)
    End If

    If relDiff <= relTolerance Then
        stats.Passed = stats.Passed + 1
        TestAppendLine stats, "OK: " & name & "; actual=" & TestFormatNumberInvariant(actual) & _
            "; expected=" & TestFormatNumberInvariant(expected) & "; relDiff=" & TestFormatNumberInvariant(relDiff)
    Else
        stats.Failed = stats.Failed + 1
        TestAppendLine stats, "FAIL: " & name & "; actual=" & TestFormatNumberInvariant(actual) & _
            "; expected=" & TestFormatNumberInvariant(expected) & "; relDiff=" & TestFormatNumberInvariant(relDiff)
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Дополняет тестовый протокол; числовые поля ниже используют точку в любой локали.
Private Sub TestAppendLine(ByRef stats As TCapacityTestStats, ByVal text As String)
    stats.Report = stats.Report & text & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Форматирует число для текстового протокола с десятичной точкой; исходные расчётные значения не меняются.
Private Function TestFormatNumberInvariant(ByVal value As Double) As String
    TestFormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ============================== ДЛЯ ТЕСТОВ ==============================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет предел в последнем интервале удвоения, точно на MaxLambda и
' за MaxLambda. Последняя ситуация не должна выдавать диагностическую нижнюю
' точку как найденную несущую способность.
Private Sub TestAudit02CapacitySearchBoundary(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars
    Dim section As CSectionModel
    Set section = TestBuildGeneratedSectionModel(mesh, rebars)
    Dim stiffness As Double
    stiffness = TestAudit02AxialStiffness(section)

    Dim roots As Variant
    roots = Array(3#, 5.25, 6#, 6.0001, 8#)
    Dim root As Variant, method As Variant
    For Each method In Array("Bisection", "Brent", "Secant")
    For Each root In roots
        Dim cap As CCapacitySolver
        Set cap = New CCapacitySolver
        TestConfigureCapacity cap
        cap.ConcreteCompressionLimit = -1#
        cap.SteelStrainLimit = 0.001
        cap.MaxLambda = 6#
        cap.LambdaTolerance = 0.000001
        cap.SolverToleranceN = 0.01
        cap.SearchMethod = CStr(method)
        cap.SolverToleranceMx = 1#: cap.SolverToleranceMy = 1#
        cap.SolveByLoadPathMultiplier section, TestLinearConcrete(), TestLinearSteel(), _
            0#, stiffness * 0.001 / CDbl(root), 0#, 0#, 0#, 0#
        Dim prefix As String
        prefix = "audit02.boundary." & CStr(method) & "." & CStr(root)
        If CDbl(root) <= 6# Then
            TestAssertTrue stats, prefix & ".converged", cap.Converged
            TestAssertClose stats, prefix & ".lambda", cap.LambdaUltimate, CDbl(root), 0.00001
            TestAssertTrue stats, prefix & ".pointConverged", cap.LastSolver.Converged
            TestAssertClose stats, prefix & ".residualN", cap.LastSolver.ResidualN, 0#, 0.01
            TestAssertClose stats, prefix & ".residualMx", cap.LastSolver.ResidualMx, 0#, 1#
            TestAssertClose stats, prefix & ".residualMy", cap.LastSolver.ResidualMy, 0#, 1#
            TestAssertClose stats, prefix & ".steelLimit", cap.LastSolver.MaxAbsSteelStrain, 0.001, 0.0000000001
        Else
            TestAssertTrue stats, prefix & ".noLimit", Not cap.Converged
            TestAssertTrue stats, prefix & ".boundFlag", cap.SearchBoundReached
            TestAssertClose stats, prefix & ".notCapacity", cap.LambdaUltimate, 0#, 0#
            Dim searchResult As CLimitSearchResult
            Set searchResult = TestAudit02CapacitySnapshot(cap, "LoadMultiplier")
            TestAssertTrue stats, prefix & ".code", searchResult.Meta.ResultCode = rcSearchBoundReached
        End If
    Next root
    Next method
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отдельно проверяет физический предел на технической границе поиска.
Public Function RunCapacitySearchBoundaryTests() As String
    Dim stats As TCapacityTestStats
    TestAudit02CapacitySearchBoundary stats
    TestAppendLine stats, "TOTAL_CAPACITY_SEARCH_BOUNDARY: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunCapacitySearchBoundaryTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет независимые пределы стали двух знаков на линейном материале.
' Числа заданы аналитически из общей осевой жесткости и физических деформаций.
Private Sub TestAudit02AsymmetricSteelLimits(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars
    Dim section As CSectionModel
    Set section = TestBuildGeneratedSectionModel(mesh, rebars)
    Dim methodIndex As Long
    For methodIndex = 1 To 2
    Dim direction As Long
    For direction = -1 To 1 Step 2
        Dim cap As CCapacitySolver
        Set cap = New CCapacitySolver
        TestConfigureCapacity cap
        cap.ConcreteCompressionLimit = -1#
        cap.SteelCompressionLimit = 0.0003
        cap.SteelTensionLimit = 0.0018
        cap.LambdaTolerance = 0.000001
        cap.StrainTolerance = 0.000000001
        cap.SolverToleranceN = 0.01
        If methodIndex = 1 Then
            cap.SolveByLoadPathMultiplier section, TestLinearConcrete(), TestLinearSteel(), _
                0#, direction * TestAudit02AxialStiffness(section) * 0.0001, 0#, 0#, 0#, 0#
        Else
            cap.SolveByUltimateLoadPath section, TestLinearConcrete(), TestLinearSteel(), _
                0#, direction * TestAudit02AxialStiffness(section) * 0.0001, 0#, 0#, 0#, 0#, False
        End If
        Dim expectedLambda As Double
        If direction < 0 Then expectedLambda = 3# Else expectedLambda = 18#
        Dim prefix As String
        prefix = "audit02.steelSign." & CStr(methodIndex) & "." & CStr(direction)
        TestAssertTrue stats, prefix & ".converged", cap.Converged
        TestAssertClose stats, prefix & ".lambda", cap.LambdaUltimate, expectedLambda, 0.00001
        TestAssertEquals stats, prefix & ".criterion", cap.LimitState, "SteelStrainLimit"
    Next direction
    Next methodIndex
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет реальную неудачу равновесия после нескольких итераций.
' Большие деформации пробной плоскости не должны создавать физический предел
' и BaseFail; основанием остаются машинные коды численной неудачи solver-а.
Private Sub TestAudit02UnconvergedProbeIsNumerical(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars
    Dim section As CSectionModel
    Set section = TestBuildGeneratedSectionModel(mesh, rebars)
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    cap.MaxRetries = 0
    cap.SolverBaseLoadSteps = 1
    cap.SolverMaxIterations = 3
    cap.ConcreteCompressionLimit = -0.000000001
    Dim path As CLoadPathVector
    Set path = New CLoadPathVector
    path.Initialize -1000000000#, 0#, 0#, 1000000#, 0#, 0#
    Dim functionValue As Double
    Dim probeState As String
    cap.LimitSearchEvaluateCapacityLoadMultiplier section, TestProvisionalConcrete(), _
        TestProvisionalSteel(), path, 0#, functionValue, probeState
    TestAssertTrue stats, "audit02.failedProbe.solverExists", Not cap.LastSolver Is Nothing
    If cap.LastSolver Is Nothing Then Exit Sub
    TestAssertTrue stats, "audit02.failedProbe.unconverged", Not cap.LastSolver.Converged
    TestAssertTrue stats, "audit02.failedProbe.multipleIterations", cap.LastSolver.Iterations > 1
    TestAssertTrue stats, "audit02.failedProbe.strainBeyondLimit", cap.LastSolver.MinConcreteStrain < -0.000000001
    TestAssertTrue stats, "audit02.failedProbe.numerical", cap.LimitSearchCapacityStateIsNumericalFailure(probeState)
    cap.LimitSearchHandleCapacityInitialFailure probeState
    Dim result As CLimitSearchResult
    Set result = TestAudit02CapacitySnapshot(cap, "LoadMultiplier")
    TestAssertTrue stats, "audit02.failedProbe.noBaseFail", result.Meta.InternalStatus = rsNumericalFailure
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Несошедшаяся положительная lambda не подтверждает физическую верхнюю
' границу, даже когда последняя итерационная плоскость превысила предел.
Private Sub TestAudit02PositiveUnconvergedProbe(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars
    Dim section As CSectionModel
    Set section = TestBuildGeneratedSectionModel(mesh, rebars)
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    cap.MaxRetries = 0
    cap.SolverBaseLoadSteps = 1
    cap.SolverMaxIterations = 3
    cap.ConcreteCompressionLimit = -0.000000001
    Dim path As CLoadPathVector
    Set path = New CLoadPathVector
    path.Initialize 0#, -1000000000#, 0#, 1000000#, 0#, 0#
    Dim residual As Double
    Dim state As String
    cap.LimitSearchEvaluateCapacityLoadMultiplier section, TestProvisionalConcrete(), _
        TestProvisionalSteel(), path, 1#, residual, state
    TestAssertTrue stats, "audit02.failedPositive.noEquilibrium", Not cap.LastSolver.Converged
    TestAssertTrue stats, "audit02.failedPositive.multiple", cap.LastSolver.Iterations > 1
    TestAssertTrue stats, "audit02.failedPositive.exceeded", cap.LastSolver.MinConcreteStrain < -0.000000001
    TestAssertTrue stats, "audit02.failedPositive.numerical", cap.LimitSearchCapacityStateIsNumericalFailure(state)
    TestAssertTrue stats, "audit02.failedPositive.notLimit", Not cap.LimitSearchCapacityStateIsPhysicalLimit(state)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Проверяет Offset ниже, точно на и выше физического критерия на линейном
' материале. Точная начальная предельная точка уже исчерпала Capacity-путь,
' но сама оценка плоскости остается физической, без технического продолжения.
Private Sub TestAudit02InitialOffsetBoundary(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars
    Dim section As CSectionModel
    Set section = TestBuildGeneratedSectionModel(mesh, rebars)
    Dim ratio As Variant
    For Each ratio In Array(0.5, 1#, 1.1)
        Dim cap As CCapacitySolver
        Set cap = New CCapacitySolver
        TestConfigureCapacity cap
        cap.ConcreteCompressionLimit = -1#
        cap.SteelStrainLimit = 0.001
        cap.SolverToleranceN = 0.01
        Dim path As CLoadPathVector
        Set path = New CLoadPathVector
        path.Initialize CDbl(ratio) * TestAudit02AxialStiffness(section) * 0.001, _
            TestAudit02AxialStiffness(section) * 0.0001, 0#, 0#, 0#, 0#
        Dim residual As Double
        Dim state As String
        cap.LimitSearchEvaluateCapacityLoadMultiplier section, TestLinearConcrete(), TestLinearSteel(), _
            path, 0#, residual, state
        Dim prefix As String
        prefix = "audit02.offset." & CStr(ratio)
        TestAssertTrue stats, prefix & ".equilibrium", cap.LastSolver.Converged
        If CDbl(ratio) < 1# Then
            TestAssertTrue stats, prefix & ".acceptable", cap.LimitSearchCapacityStateIsAcceptable(state)
        ElseIf CDbl(ratio) = 1# Then
            ' Равновесие решается с прежним абсолютным допуском усилий.
            ' Уточненная плоскость может оказаться по любую сторону точной
            ' аналитической границы; классификация должна отражать именно ее,
            ' а не объявлять превышение по одному значению входной нагрузки.
            TestAssertClose stats, prefix & ".exactUtilization", cap.TestCriticalStrainUtilization, 1#, 0.000000000001
            TestAssertTrue stats, prefix & ".classification", _
                cap.LimitSearchCapacityStateIsPhysicalLimit(state) = (cap.TestCriticalStrainUtilization >= 1#)
            TestAssertTrue stats, prefix & ".notNumerical", Not cap.LimitSearchCapacityStateIsNumericalFailure(state)
        Else
            TestAssertTrue stats, prefix & ".limit", cap.LimitSearchCapacityStateIsPhysicalLimit(state)
            cap.LimitSearchHandleCapacityInitialFailure state
            Dim result As CLimitSearchResult
            Set result = TestAudit02CapacitySnapshot(cap, "LoadMultiplier")
            TestAssertTrue stats, prefix & ".code", result.Meta.ResultCode = rcInitialStateBeyondLimit
            Dim policy As CResultStatusPolicy
            Set policy = New CResultStatusPolicy
            TestAssertEquals stats, prefix & ".display", policy.ExternalStatus(result.Meta), "BaseFail"
        End If
    Next ratio
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' общий метод Ньютона работает через ILimitSearchProblem
' без экземпляров расчёта несущей способности и трещин. Аналитический корень
' проверяется отдельно; ожидаемые значения физических примеров не меняются.
Private Sub TestAudit02GenericUltimateSearch(ByRef stats As TCapacityTestStats)
    Dim problem As CTestLimitSearchProblem
    Set problem = New CTestLimitSearchProblem
    problem.TestConfigure "Bisection", 1.25, 0.000000001, 20
    Dim loadPath As CLoadPathVector
    Set loadPath = New CLoadPathVector
    loadPath.Initialize 0#, 1#, 0#, 0#, 0#, 0#
    Dim search As CUltimateStrainSearch
    Set search = New CUltimateStrainSearch
    TestAssertTrue stats, "audit02.genericUltimate.succeeded", _
        search.RunNewton(problem, Nothing, Nothing, Nothing, loadPath, Nothing, "Тестовый поиск не сошелся.")
    TestAssertClose stats, "audit02.genericUltimate.lambda", problem.TestFinalLambda, 1.25, 0.000000001
    TestAssertEquals stats, "audit02.genericUltimate.method", problem.TestFinalMethod, "UltimateStrain"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает точную осевую жесткость тестового линейного сечения, Н.
' Бетонная сетка включает площадь стержней, поэтому арматура добавляет
' разность Es-Eb, как и интегрирование напряжений в CSectionSolver.
Private Function TestAudit02AxialStiffness(ByVal section As CSectionModel) As Double
    Dim i As Long
    For i = 1 To section.ConcreteCount
        TestAudit02AxialStiffness = TestAudit02AxialStiffness + 32500# * section.ConcreteArea(i)
    Next i
    For i = 1 To section.RebarCount
        TestAudit02AxialStiffness = TestAudit02AxialStiffness + (200000# - 32500#) * section.RebarArea(i)
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что найденный предел является успехом Search, а непрохождение
' текущего сочетания определяет только Capacity. Числа старого сценария и
' его допуски не изменяются; отдельно проверяется изоляция выданной meta.
Private Sub TestCheckAudit02SearchVsCapacity(ByRef stats As TCapacityTestStats, _
        ByVal cap As CCapacitySolver, ByVal section As CSectionModel, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double)
    Dim spec As CMaterialModelSpec
    Set spec = New CMaterialModelSpec
    spec.Initialize "ULS(I)", "TwoLine", "Ignore", "TwoLine"
    Dim search As CLimitSearchResult
    Set search = TestAudit02CapacitySnapshot(cap, "LoadMultiplier", section, _
        TestProvisionalConcrete(), TestProvisionalSteel(), spec)
    TestAssertTrue stats, "audit02.searchEngineering.searchSuccess", search.Meta.InternalStatus = rsSuccess
    TestAssertTrue stats, "audit02.searchEngineering.searchCode", search.Meta.ResultCode = rcCheckPassed
    TestAssertTrue stats, "audit02.searchEngineering.searchSucceeded", search.Succeeded
    Dim loadState As CSectionLoadState
    Set loadState = New CSectionLoadState
    loadState.Initialize targetN, targetMx, targetMy, 0#, 0#
    Dim path As CLoadPathDescriptor
    Set path = New CLoadPathDescriptor
    path.InitializeFromLoadState "lambda*Mxy", loadState
    Dim result As CCapacityResult
    Set result = New CCapacityResult
    result.InitializeFromSearch search, path, False, 0#, 0#, Nothing, 300#, 200#
    TestAssertTrue stats, "audit02.searchEngineering.capacityFailed", result.ResultMeta.InternalStatus = rsCheckFailed
    TestAssertEquals stats, "audit02.searchEngineering.capacityDisplay", result.Status, "FAIL"
    TestAssertClose stats, "audit02.searchEngineering.lambdaUnchanged", result.LambdaCapacity, cap.LambdaUltimate, 0#
    Dim detachedMeta As CResultMeta
    Set detachedMeta = search.Meta
    detachedMeta.SetResult rsInternalError, rcInternalError, rkCapacity, "Изменение копии в тесте."
    TestAssertTrue stats, "audit02.searchEngineering.metaSnapshot", search.Meta.InternalStatus = rsSuccess
    TestAssertTrue stats, "audit02.searchEngineering.pointState", Not search.PointState Is Nothing
    If Not search.PointState Is Nothing Then
        TestAssertTrue stats, "audit02.searchEngineering.sharedState", result.StateResult Is search.PointState
        Dim savedN As Double, savedEps As Double
        savedN = search.PointState.Nint
        savedEps = search.PointState.Epsilon0
        cap.LastSolver.EvaluateStrainPlane section, TestProvisionalConcrete(), TestProvisionalSteel(), 0#, 0#, 0#
        TestAssertClose stats, "audit02.searchEngineering.solverMutation.N", search.NUltimate, savedN, 0#
        TestAssertClose stats, "audit02.searchEngineering.solverMutation.plane", search.PointState.Epsilon0, savedEps, 0#
    End If
    Dim detachedSearch As CLimitSearchResult
    Set detachedSearch = result.SearchResult
    detachedSearch.InitializeInvalidConfiguration "Unknown", "Изменение выданной копии Search."
    TestAssertTrue stats, "audit03.searchEngineering.outputSnapshot", result.SearchResult.HasLimitPoint
    TestAssertClose stats, "audit03.searchEngineering.outputLambda", result.SearchResult.LambdaUltimate, cap.LambdaUltimate, 0#
    search.InitializeInvalidConfiguration "Unknown", "Повторное заполнение исходного Search."
    TestAssertTrue stats, "audit03.searchEngineering.inputSnapshot", result.SearchResult.HasLimitPoint
    TestAssertTrue stats, "audit03.searchEngineering.inputMeta", result.SearchResult.Meta.InternalStatus = rsSuccess
    TestAssertClose stats, "audit03.searchEngineering.inputLambda", result.SearchResult.LambdaUltimate, cap.LambdaUltimate, 0#
    TestAssertEquals stats, "audit03.searchEngineering.capacityStatusPreserved", result.Status, "FAIL"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' создаёт сохранённый результат несущей способности тем же способом,
' что обычный поиск. Общему результату не передаются инженерные расчётчики.
Private Function TestAudit02CapacitySnapshot(ByVal cap As CCapacitySolver, ByVal strategy As String, _
        Optional ByVal section As CSectionModel = Nothing, _
        Optional ByVal concrete As Object = Nothing, Optional ByVal steel As Object = Nothing, _
        Optional ByVal spec As CMaterialModelSpec = Nothing) As CLimitSearchResult
    Dim problem As CCapacityLimitSearchProblem
    Set problem = New CCapacityLimitSearchProblem
    problem.Initialize cap
    Set TestAudit02CapacitySnapshot = problem.BuildSnapshot(strategy, section, concrete, steel, spec)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' общий подбор интервала и восстановление после неудачи работают
' для обоих видов задач без железобетонных расчётчиков. Проверяются одномерные
' методы, MaxLambda, отсутствие точки за границей и повторное заполнение результата.
Private Sub TestAudit02GenericLoadMultiplierMatrix(ByRef stats As TCapacityTestStats)
    Dim kind As Variant, methodName As Variant, root As Variant
    For Each kind In Array(rkCapacity, rkCrackFormation)
        For Each methodName In Array("Bisection", "Brent", "Secant")
            For Each root In Array(0.75, 5.25, 6#, 8#)
                Dim problem As CTestLimitSearchProblem
                Set problem = New CTestLimitSearchProblem
                problem.TestConfigure CStr(methodName), CDbl(root), 0.000001, 80, 1#, 6#, CLng(kind)
                Dim request As CLimitSearchRequest
                Set request = New CLimitSearchRequest
                request.InitializeWithProblem problem, CLng(kind), "LoadMultiplier", 0#, 1#, 0#, 0#, 0#, 0#
                Dim search As CLoadMultiplierSearch
                Set search = New CLoadMultiplierSearch
                Dim result As CLimitSearchResult
                Set result = search.Execute(request)
                Dim prefix As String
                prefix = "audit02.genericMultiplier." & CStr(kind) & "." & CStr(methodName) & "." & CStr(root)
                If CDbl(root) <= 6# Then
                    TestAssertTrue stats, prefix & ".success", result.Succeeded
                    TestAssertTrue stats, prefix & ".point", result.HasLimitPoint
                    TestAssertClose stats, prefix & ".lambda", result.LambdaUltimate, CDbl(root), 0.000001
                    TestAssertClose stats, prefix & ".loads", result.NUltimate, result.LambdaUltimate, 0#
                    TestAssertEquals stats, prefix & ".method", result.ActualMethod, CStr(methodName)
                Else
                    TestAssertTrue stats, prefix & ".boundCode", result.Meta.ResultCode = rcSearchBoundReached
                    TestAssertTrue stats, prefix & ".noPoint", Not result.HasLimitPoint
                    TestAssertClose stats, prefix & ".noCapacity", result.LambdaUltimate, 0#, 0#
                    TestAssertClose stats, prefix & ".maxProbed", problem.TestProbeMaxLambda, 6#, 0#
                End If
                result.InitializeInvalidConfiguration "Unknown", "Проверка очистки", CLng(kind)
                TestAssertTrue stats, prefix & ".resetPoint", Not result.HasLimitPoint
                TestAssertTrue stats, prefix & ".resetState", result.PointState Is Nothing
                TestAssertTrue stats, prefix & ".resetExecution", Not result.SearchExecuted
                TestAssertClose stats, prefix & ".resetLoads", result.NUltimate, 0#, 0#
            Next root
        Next methodName
        Set problem = New CTestLimitSearchProblem
        problem.TestConfigure "Bisection", 1.1, 0.000001, 80, 1#, 6#, CLng(kind)
        problem.TestFailureAbove = 1.2
        Set request = New CLimitSearchRequest
        request.InitializeWithProblem problem, CLng(kind), "LoadMultiplier", 0#, 1#, 0#, 0#, 0#, 0#
        Set result = search.Execute(request)
        TestAssertTrue stats, "audit02.genericMultiplier.recovery." & CStr(kind) & ".point", result.HasLimitPoint
        TestAssertClose stats, "audit02.genericMultiplier.recovery." & CStr(kind) & ".lambda", result.LambdaUltimate, 1.1, 0.000001
        TestAssertTrue stats, "audit02.genericMultiplier.recovery." & CStr(kind) & ".upperFailed", problem.TestProbeMaxLambda > 1.2
    Next kind

    Set request = New CLimitSearchRequest
    Set result = search.Execute(request)
    TestAssertTrue stats, "audit02.genericMultiplier.missingProblem.internal", result.Meta.InternalStatus = rsInternalError
    TestAssertTrue stats, "audit02.genericMultiplier.missingProblem.noPoint", Not result.HasLimitPoint
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' допуск lambda не отменяет физическую приемку. Все методы и
' оба вида задач сужают ту же скобку после численной неудачи уточнения.
Public Function RunBracketFinalizationRecoveryTests() As String
    Dim stats As TCapacityTestStats, kind As Variant, methodName As Variant
    Dim problem As CTestLimitSearchProblem, request As CLimitSearchRequest
    Dim search As CLoadMultiplierSearch, result As CLimitSearchResult, prefix As String
    Set search = New CLoadMultiplierSearch
    For Each kind In Array(rkCapacity, rkCrackFormation)
        For Each methodName In Array("Bisection", "Brent", "Secant")
            Set problem = New CTestLimitSearchProblem
            problem.TestConfigure CStr(methodName), 1.3, 0.02, 80, 1#, 6#, CLng(kind)
            problem.TestStepCriterion = True
            problem.TestFinalizationMaxWidth = 0.0002
            Set request = New CLimitSearchRequest
            request.InitializeWithProblem problem, CLng(kind), "LoadMultiplier", 0#, 1#, 0#, 0#, 0#, 0#
            Set result = search.Execute(request)
            prefix = "bracketFinalize." & CStr(kind) & "." & CStr(methodName)
            TestAssertTrue stats, prefix & ".success", result.Succeeded
            TestAssertTrue stats, prefix & ".point", result.HasLimitPoint
            TestAssertClose stats, prefix & ".physicalAccuracy", result.LambdaUltimate, 1.3, 0.0002
            TestAssertTrue stats, prefix & ".repeatedAcceptance", problem.TestBracketFinalizationCalls > 1
            TestAssertTrue stats, prefix & ".budget", problem.TestFinalIterations <= 80
            TestAssertTrue stats, prefix & ".recoveryLogged", InStr(1, problem.TestDiagnosticLog, "bracketFinalizationRecovery", vbBinaryCompare) > 0
            TestAssertEquals stats, prefix & ".method", result.ActualMethod, CStr(methodName)
        Next methodName
    Next kind
    Set problem = New CTestLimitSearchProblem
    problem.TestConfigure "Bisection", 1.3, 0.02, 7
    problem.TestStepCriterion = True: problem.TestFinalizationMaxWidth = 0.0002
    Set request = New CLimitSearchRequest
    request.InitializeWithProblem problem, rkCapacity, "LoadMultiplier", 0#, 1#, 0#, 0#, 0#, 0#
    Set result = search.Execute(request)
    TestAssertTrue stats, "bracketFinalize.exhausted.noPoint", Not result.HasLimitPoint
    TestAssertTrue stats, "bracketFinalize.exhausted.numerical", result.Meta.InternalStatus = rsNumericalFailure
    TestAssertTrue stats, "bracketFinalize.exhausted.noBudgetReset", problem.TestBracketFinalizationCalls = 1

    Set problem = New CTestLimitSearchProblem
    problem.TestConfigure "Bisection", 1.3, 0.02, 0
    problem.TestStepCriterion = True: problem.TestFinalizationMaxWidth = -1#
    Set request = New CLimitSearchRequest
    request.InitializeWithProblem problem, rkCapacity, "LoadMultiplier", 0#, 1#, 0#, 0#, 0#, 0#
    Set result = search.Execute(request)
    TestAssertTrue stats, "bracketFinalize.machine.noPoint", Not result.HasLimitPoint
    TestAssertTrue stats, "bracketFinalize.machine.numerical", result.Meta.InternalStatus = rsNumericalFailure
    TestAssertTrue stats, "bracketFinalize.machine.bounded", problem.TestProbeCalls < 100

    Set problem = New CTestLimitSearchProblem
    problem.TestConfigureMultiplierFaultCase 11, "Bisection", rkCapacity
    Set request = New CLimitSearchRequest
    request.InitializeWithProblem problem, rkCapacity, "LoadMultiplier", 0#, 1#, 0#, 0#, 0#, 0#
    Set result = search.Execute(request)
    TestAssertTrue stats, "bracketFinalize.terminal.configuration", result.Meta.InternalStatus = rsInvalidConfiguration
    TestAssertTrue stats, "bracketFinalize.terminal.once", problem.TestBracketFinalizationCalls = 1
    Set problem = New CTestLimitSearchProblem
    problem.TestConfigure "Bisection", 1.3, 0.02, 80
    problem.TestStepCriterion = True: problem.TestDeferOverallCompletion = True
    Set request = New CLimitSearchRequest
    request.InitializeWithProblem problem, rkCrackFormation, "LoadMultiplier", 0#, 1#, 0#, 0#, 0#, 0#
    Set result = search.Execute(request)
    TestAssertTrue stats, "bracketFinalize.deferred.pointAccepted", result.Succeeded And result.HasLimitPoint
    TestAssertTrue stats, "bracketFinalize.deferred.ownerNotComplete", Not problem.TestConverged
    TestAssertTrue stats, "bracketFinalize.deferred.noRepeat", problem.TestBracketFinalizationCalls = 1
    TestAssertTrue stats, "bracketFinalize.deferred.noFallback", problem.TestFallbackCalls = 0
    TestAppendLine stats, "TOTAL BRACKET FINALIZATION RECOVERY: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunBracketFinalizationRecoveryTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ============================== ДЛЯ ТЕСТОВ ==============================
' Отдельный entrypoint для watchdog: соседние Double и нулевой bisection-budget
' на неисправном алгоритме зависают, а не возвращают правдоподобный предел.
Public Function TestRunAudit03SearchStagnation() As String
    Dim stats As TCapacityTestStats
    TestAudit03SearchArithmeticCase stats, "Bisection", rkCapacity, 1
    TestAppendLine stats, "TOTAL_AUDIT03_STAGNATION: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    TestRunAudit03SearchStagnation = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отдельный полный прогон арифметических краев generic Search.
Public Function RunAudit03SearchTests() As String
    Dim stats As TCapacityTestStats
    TestAudit03SearchArithmetic stats
    TestAppendLine stats, "TOTAL_AUDIT03_SEARCH: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03SearchTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' обычный поиск несущей способности без искусственного ограничения
' делений интервала. Проверяет передачу окончательных кодов ошибок.
Public Function TestRunAudit03CapacityContracts() As String
    Dim stats As TCapacityTestStats
    TestAudit03RealCapacityPrecision stats
    TestAudit03CapacityTypedFailures stats
    TestAppendLine stats, "TOTAL_AUDIT03_CAPACITY_CONTRACTS: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    TestRunAudit03CapacityContracts = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' линейные материалы исключают неудачу равновесия. Если допуск
' меньше разрешения Double, поиск должен честно сообщить причину остановки.
' Разные значения lambda не должны объединяться в одну сохранённую пробу.
Private Sub TestAudit03RealCapacityPrecision(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder, rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars
    Dim section As CSectionModel
    Set section = TestBuildGeneratedSectionModel(mesh, rebars)
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    TestConfigureCapacity cap
    cap.SearchMethod = "Bisection"
    cap.LambdaTolerance = 1E-18
    cap.MaxRetries = 0
    cap.DiagnosticsEnabled = True
    TestSaveAudit03SearchProgress stats, "START: audit03.capacity.realPrecision"
    cap.SolveByLoadPathMultiplier section, TestLinearConcrete(), TestLinearSteel(), 0#, 0#, 0#, 10000000#, 0#, 0#
    Dim result As CLimitSearchResult
    Set result = TestAudit02CapacitySnapshot(cap, "LoadMultiplier")
    TestAssertTrue stats, "audit03.capacity.realPrecision.numerical", result.Meta.InternalStatus = rsNumericalFailure
    TestAssertTrue stats, "audit03.capacity.realPrecision.code", result.Meta.ResultCode = rcNumericalFailure
    TestAssertTrue stats, "audit03.capacity.realPrecision.noPoint", Not result.HasLimitPoint
    TestAssertTrue stats, "audit03.capacity.realPrecision.notBound", Not cap.SearchBoundReached
    TestAssertTrue stats, "audit03.capacity.realPrecision.finiteWork", cap.Iterations > 40 And cap.Iterations < 100
    TestAssertTrue stats, "audit03.capacity.realPrecision.calculated", result.Meta.Calculated
    TestAssertTrue stats, "audit03.capacity.realPrecision.reason", Len(result.Meta.ResultComment) > 0
    TestAppendLine stats, "COMMENT: audit03.capacity.realPrecision; " & result.Meta.ResultComment
    TestAppendLine stats, "DIAGNOSTIC: " & cap.DiagnosticLog
    TestSaveAudit03SearchProgress stats, "DONE: audit03.capacity.realPrecision"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' configuration/input/internal/singular причины проходят от
' реального Capacity solve в Search-result без retries и назначения BaseFail
' программной ошибке. Повторный запуск не наследует meta предыдущей задачи.
Private Sub TestAudit03CapacityTypedFailures(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder, rebars As CRebarLayout
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars
    Dim validSection As CSectionModel
    Set validSection = TestBuildGeneratedSectionModel(mesh, rebars)
    Dim scenario As Long
    For scenario = 1 To 9
        Dim cap As CCapacitySolver
        Set cap = New CCapacitySolver
        TestConfigureCapacity cap
        cap.SearchMethod = "Bisection"
        cap.DiagnosticsEnabled = True
        Dim section As CSectionModel
        Set section = validSection
        Dim concrete As Object, steel As Object
        Set concrete = TestLinearConcrete()
        Set steel = TestLinearSteel()
        Dim expectedStatus As EResultInternalStatus, expectedCode As EResultCode
        Dim expectedFailure As ESolverFailureCode, nOffset As Double
        nOffset = 0#
        Select Case scenario
            Case 1
                cap.SolverMethod = "Invalid"
                expectedStatus = rsInvalidConfiguration: expectedCode = rcInvalidConfiguration
                expectedFailure = sfcInvalidConfiguration
            Case 2
                cap.SolverMaxIterations = 0
                expectedStatus = rsInvalidConfiguration: expectedCode = rcInvalidConfiguration
                expectedFailure = sfcInvalidConfiguration
            Case 3
                Set section = Nothing
                expectedStatus = rsInternalError: expectedCode = rcInternalError
                expectedFailure = sfcInternalError
            Case 4
                Set concrete = Nothing
                expectedStatus = rsInternalError: expectedCode = rcInternalError
                expectedFailure = sfcInternalError
            Case 5
                Set section = New CSectionModel
                expectedStatus = rsInvalidInput: expectedCode = rcInvalidInput
                expectedFailure = sfcInvalidInput
            Case 6
                Set section = New CSectionModel
                section.AddConcreteElement 0#, 0#, 10000#, sourceName:="Singular"
                nOffset = -100000#
                cap.MaxRetries = 0
                expectedStatus = rsNumericalFailure: expectedCode = rcSingularTangent
                expectedFailure = sfcSingularTangent
            Case 7, 8, 9
                If scenario = 7 Then cap.SolverMethod = "Invalid"
                If scenario = 8 Then cap.SolverMaxIterations = 0
                If scenario = 9 Then cap.SolverToleranceN = -1#
                expectedStatus = rsInvalidConfiguration: expectedCode = rcInvalidConfiguration
                expectedFailure = sfcInvalidConfiguration
        End Select
        Dim prefix As String
        prefix = "audit03.capacity.typed." & CStr(scenario)
        TestSaveAudit03SearchProgress stats, "START: " & prefix
        If scenario <= 6 Then
            cap.SolveByLoadPathMultiplier section, concrete, steel, nOffset, 0#, 0#, 10000000#, 0#, 0#
        Else
            cap.SolveByUltimateLoadPath section, concrete, steel, nOffset, 0#, 0#, 10000000#, 0#, 0#
        End If
        Dim result As CLimitSearchResult
        Set result = TestAudit02CapacitySnapshot(cap, "LoadMultiplier")
        TestAssertTrue stats, prefix & ".status", result.Meta.InternalStatus = expectedStatus
        TestAssertTrue stats, prefix & ".code", result.Meta.ResultCode = expectedCode
        TestAssertTrue stats, prefix & ".failure", cap.FailureCode = expectedFailure
        TestAssertTrue stats, prefix & ".noRetry", cap.TestRetryCount = 0
        TestAssertTrue stats, prefix & ".noPoint", Not result.HasLimitPoint
        TestAssertTrue stats, prefix & ".notBaseFail", result.Meta.ResultCode <> rcInitialStateBeyondLimit
        TestAssertTrue stats, prefix & ".reason", Len(result.Meta.ResultComment) > 0
        If scenario <> 6 Then TestAssertTrue stats, prefix & ".notCalculated", Not result.Meta.Calculated
        TestAppendLine stats, "COMMENT: " & prefix & "; " & result.Meta.ResultComment
        TestSaveAudit03SearchProgress stats, "DONE: " & prefix
    Next scenario
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отдельный watchdog воспроизводит бесконечный line search
' автоматического общего Newton при неизменной норме.
Public Function TestRunAudit03UltimateStagnation() As String
    Dim stats As TCapacityTestStats
    TestAudit03UltimateGuardCase stats, rkCapacity, 1
    TestScaledUltimateStepAcceptance stats
    TestRejectedUltimateFinalPoint stats
    TestAppendLine stats, "TOTAL_AUDIT03_ULTIMATE_STAGNATION: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    TestRunAudit03UltimateStagnation = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Даже очень малый шаг принимается, если все конечные критерии выполнены.
' Отказ конечной приемки нельзя скрыть малым шагом или нулевой внутренней невязкой.
Private Sub TestScaledUltimateStepAcceptance(ByRef stats As TCapacityTestStats)
    Dim kind As Variant, scenario As Long, problem As CTestLimitSearchProblem
    Dim search As CUltimateStrainSearch, path As CLoadPathVector, succeeded As Boolean, prefix As String
    For Each kind In Array(rkCapacity, rkCrackFormation)
        For scenario = 22 To 23
            Set problem = New CTestLimitSearchProblem
            problem.TestConfigureUltimateGuardCase scenario, CLng(kind)
            Set search = New CUltimateStrainSearch: Set path = New CLoadPathVector
            path.Initialize 0#, 1#, 0#, 0#, 0#, 0#
            succeeded = search.RunNewton(problem, Nothing, Nothing, Nothing, path, Nothing, "Тестовый поиск не сошелся.")
            prefix = "scaledUltimateStep." & CStr(kind) & "." & CStr(scenario)
            TestAssertTrue stats, prefix & ".finiteCalls", problem.TestUltimateCalls > 1 And problem.TestUltimateCalls <= 9
            If scenario = 22 Then
                TestAssertTrue stats, prefix & ".smallStepMayReachValidLimit", succeeded And problem.TestConverged
                TestAssertTrue stats, prefix & ".noFailure", problem.TestFailureCode = sfcNone
            Else
                TestAssertTrue stats, prefix & ".noFalseSuccess", Not succeeded And Not problem.TestConverged
                TestAssertTrue stats, prefix & ".numericalNotPhysicalFailure", problem.TestFailureCode = sfcNumericalFailure
                TestAssertTrue stats, prefix & ".numericalReason", Len(problem.TestDiagnosticLog) > 0
            End If
        Next scenario
    Next kind
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Недопустимый корень не уточняется повторно; это численный отказ попытки,
' а не физическое разрушение и не запрет последующего разрешенного поиска.
Private Sub TestRejectedUltimateFinalPoint(ByRef stats As TCapacityTestStats)
    Dim kind As Variant, scenario As Long, problem As CTestLimitSearchProblem, path As CLoadPathVector
    Dim search As CUltimateStrainSearch, succeeded As Boolean, prefix As String
    For Each kind In Array(rkCapacity, rkCrackFormation)
        For scenario = 24 To 26
            Set problem = New CTestLimitSearchProblem
            problem.TestConfigureUltimateGuardCase scenario, CLng(kind)
            Set search = New CUltimateStrainSearch: Set path = New CLoadPathVector
            path.Initialize 0#, 1#, 0#, 0#, 0#, 0#
            succeeded = search.RunNewton(problem, Nothing, Nothing, Nothing, path, Nothing, "Не сошлось.")
            prefix = "ultimate.rejectedPoint." & CStr(kind) & "." & CStr(scenario)
            TestAssertTrue stats, prefix & ".noFalseSuccess", Not succeeded And Not problem.TestConverged
            If scenario = 25 Then
                TestAssertTrue stats, prefix & ".finalIteration", problem.TestUltimateCalls = 6
            Else
                TestAssertTrue stats, prefix & ".singleEvaluation", problem.TestUltimateCalls = 1
            End If
            If scenario = 26 Then
                TestAssertTrue stats, prefix & ".invalidDisposition", problem.TestFailureCode = sfcInternalError
            Else
                TestAssertTrue stats, prefix & ".numericalNotPhysicalFailure", problem.TestFailureCode = sfcNumericalFailure
            End If
            TestAssertTrue stats, prefix & ".reasonPreserved", Len(problem.TestFailureReason) > 0
        Next scenario
    Next kind
    TestCapacityRejectedDirectionRecovery stats
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Предел противоположного направления не принимается как несущая. После
' его отказа пригодная точка того же пути принимается без прежнего failure-code.
Private Sub TestCapacityRejectedDirectionRecovery(ByRef stats As TCapacityTestStats)
    Dim shift As Variant, x As Variant, y As Variant, section As CSectionModel
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram, tension As CSectionSolver
    Dim compression As CSectionSolver, cap As CCapacitySolver, adapter As CCapacityLimitSearchProblem
    Dim problem As ILimitSearchProblem, path As CLoadPathVector, outcome As EUltimateFinalization
    Dim lambdaValue As Double, prefix As String
    For Each shift In Array(0#, -1000000#, 1000000#)
        Set section = New CSectionModel
        For Each x In Array(-100#, 100#)
            For Each y In Array(-50#, 50#)
                section.AddConcreteElement CDbl(shift) + CDbl(x), -CDbl(shift) + CDbl(y), 2500#
                section.AddRebarElement CDbl(shift) + 0.8 * CDbl(x), -CDbl(shift) + 0.6 * CDbl(y), 20#, 0#, "A400"
            Next y
        Next x
        Set concrete = TestProvisionalConcrete(): Set steel = TestProvisionalSteel()
        Set tension = New CSectionSolver
        tension.EvaluatePhysicalCenterPlane section, concrete, steel, 0.025, 0#, 0#
        Set path = New CLoadPathVector
        path.Initialize 0#, -tension.Nint, 0#, -tension.Mxint, 0#, -tension.Myint
        Set cap = New CCapacitySolver
        cap.ConcreteCompressionLimit = -0.0035: cap.ConcreteTensionLimit = 0#
        cap.SteelCompressionLimit = 0.025: cap.SteelTensionLimit = 0.025
        cap.SolverToleranceN = 0.001: cap.SolverToleranceMx = 200#: cap.SolverToleranceMy = 200#
        Set adapter = New CCapacityLimitSearchProblem: adapter.Initialize cap: Set problem = adapter
        prefix = "ultimate.rejectedDirection." & CStr(shift)
        TestAssertTrue stats, prefix & ".prepared", cap.LimitSearchPrepareCapacityUltimate(section, concrete, steel, path)
        outcome = problem.LimitSearchFinalizeUltimateResult(section, concrete, steel, path, -1#, tension, "SteelStrainLimit", 1)
        TestAssertTrue stats, prefix & ".rejected", outcome = ufRejectedPoint And Not cap.Converged
        TestAssertTrue stats, prefix & ".numericalFailure", cap.FailureCode = sfcNumericalFailure
        TestAssertTrue stats, prefix & ".retryAllowed", Not problem.StoppedTerminally
        TestAssertTrue stats, prefix & ".diagnostic", InStr(1, cap.DiagnosticLog, "lambda: -1", vbBinaryCompare) > 0
        Set compression = New CSectionSolver
        compression.EvaluatePhysicalCenterPlane section, concrete, steel, -0.0035, 0#, 0#
        lambdaValue = compression.Nint / path.NBase
        outcome = problem.LimitSearchFinalizeUltimateResult(section, concrete, steel, path, lambdaValue, compression, "ConcreteStrainLimit", 1)
        TestAssertTrue stats, prefix & ".acceptedRecovery", outcome = ufAccepted And cap.Converged
        TestAssertTrue stats, prefix & ".failureCleared", cap.FailureCode = sfcNone
        TestAssertTrue stats, prefix & ".recoveryEquilibrium", cap.LastSolver.ConfirmEquilibrium( _
            path.TargetN(lambdaValue), path.TargetMx(lambdaValue), path.TargetMy(lambdaValue))
    Next shift
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Численно неудачная верхняя проба не закрывает recovery по допуску lambda.
' Подтвержденная скобка нужна даже когда ее масштаб меньше этого допуска.
Private Sub TestSmallLambdaBracketRecovery(ByRef stats As TCapacityTestStats)
    Dim kind As Variant, methodName As Variant, problem As CTestLimitSearchProblem
    Dim request As CLimitSearchRequest, search As CLoadMultiplierSearch, result As CLimitSearchResult
    Dim prefix As String
    For Each kind In Array(rkCapacity, rkCrackFormation)
        For Each methodName In Array("Bisection", "Brent", "Secant")
            Set problem = New CTestLimitSearchProblem
            problem.TestConfigure CStr(methodName), 0.0000012, 0.000001, 80, 0.000001, 0.000008, CLng(kind)
            problem.TestFailureAbove = 0.00000125
            Set request = New CLimitSearchRequest
            request.InitializeWithProblem problem, CLng(kind), "LoadMultiplier", 0#, 1#, 0#, 0#, 0#, 0#
            Set search = New CLoadMultiplierSearch: Set result = search.Execute(request)
            prefix = "smallLambdaRecovery." & CStr(kind) & "." & CStr(methodName)
            TestAssertTrue stats, prefix & ".success", result.Succeeded And result.HasLimitPoint
            TestAssertTrue stats, prefix & ".recoveryProbed", problem.TestProbeCalls >= 5
            TestAssertTrue stats, prefix & ".failedUpper", problem.TestProbeMaxLambda > 0.00000125
            TestAssertTrue stats, prefix & ".noPrematureStop", _
                InStr(1, problem.TestDiagnosticLog, "bracketNumericalFailure", vbBinaryCompare) = 0
            TestAssertClose stats, prefix & ".withinRequestedTolerance", result.LambdaUltimate, 0.0000012, 0.000001
        Next methodName
    Next kind
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' малые lambda, конечный бюджет и машинная стагнация общего поиска.
Public Function RunSmallLambdaRecoveryTests() As String
    Dim stats As TCapacityTestStats
    TestSmallLambdaBracketRecovery stats
    TestLimitSearchInitialRecoveryBudget stats
    TestAudit03SearchArithmetic stats
    TestAppendLine stats, "TOTAL_SMALL_LAMBDA_RECOVERY: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunSmallLambdaRecoveryTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Изменение допуска N меняет приемку, но не относительный масштаб нормы
' одной и той же физической плоскости при N=const и масштабировании моментов.
Private Sub TestCapacityMomentMeritScaling(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder, rebars As CRebarLayout, section As CSectionModel
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars
    Set section = TestBuildGeneratedSectionModel(mesh, rebars)
    Dim tolerance As Variant, cap As CCapacitySolver, path As CLoadPathVector
    Dim eps0 As Double, kx As Double, ky As Double, solver As CSectionSolver
    Dim rN As Double, rDir As Double, rLimit As Double, lambdaValue As Double, state As String
    Dim referenceNorm As Double, norm As Double
    Set path = New CLoadPathVector: path.Initialize -100000#, 0#, 0#, 10000000#, 0#, 3000000#
    For Each tolerance In Array(1#, 0.00000001)
        Set cap = New CCapacitySolver: TestConfigureCapacity cap
        cap.SolverToleranceN = CDbl(tolerance)
        TestAssertTrue stats, "momentMerit.prepared." & CStr(tolerance), _
            cap.LimitSearchPrepareCapacityUltimate(section, TestLinearConcrete(), TestLinearSteel(), path)
        cap.LimitSearchInitializeUltimateGuess section, TestLinearConcrete(), TestLinearSteel(), path, Nothing, eps0, kx, ky
        TestAssertTrue stats, "momentMerit.evaluated." & CStr(tolerance), _
            cap.LimitSearchEvaluateUltimateResidual(section, TestLinearConcrete(), TestLinearSteel(), path, _
                -0.001, 0.00001, 0.000002, solver, rN, rDir, rLimit, lambdaValue, state)
        norm = cap.LimitSearchUltimateResidualNorm(rN, rDir, rLimit)
        TestAssertTrue stats, "momentMerit.notAccepted." & CStr(tolerance), _
            Not cap.LimitSearchUltimateResidualReached(rN, rDir, rLimit)
        If CDbl(tolerance) = 1# Then
            referenceNorm = norm
        Else
            TestAssertClose stats, "momentMerit.toleranceIndependent", norm, referenceNorm, 0.000000000001
        End If
    Next tolerance
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Силовая траектория с малыми неизменными моментами проходит материальное
' плато. Вырождение Якобиана не доказывает отсутствия физического предела.
Private Sub TestForcePathMaterialPlateau(ByRef stats As TCapacityTestStats, Optional ByVal loadSlot As Long = 0)
    Dim extended As Variant, shift As Variant, section As CSectionModel, x As Long, y As Long
    Dim concreteParameters As CConcreteMaterialParameters, steelParameters As CSteelMaterialParameters
    Dim provider As CMaterialModelProvider, spec As CMaterialModelSpec, concrete As Object, steel As Object
    Dim cap As CCapacitySolver, physical As CStateSolutionRunner, prefix As String, referenceLambda As Double
    Dim nBase As Double, mx As Double, my As Double
    nBase = 10000#: mx = 205939.65: my = -102969.825
    If loadSlot > 0 Then
        nBase = (1# + loadSlot / 100#) * 9806.65
        mx = (0.02 + loadSlot / 1000#) * 9806650#
        my = -mx / 2#
    End If
    Set concreteParameters = New CConcreteMaterialParameters
    concreteParameters.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    Set steelParameters = New CSteelMaterialParameters
    steelParameters.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#
    Set spec = New CMaterialModelSpec: spec.Initialize "ULS(I)", "ThreeLine", "Ignore", "TwoLine"
    Set physical = New CStateSolutionRunner
    For Each extended In Array(False, True)
        Set provider = New CMaterialModelProvider
        provider.TestInitializeFromParameters concreteParameters, steelParameters, diagramExtensionEnabled:=CBool(extended)
        Set concrete = provider.ConcreteMaterialFromSpec(spec): Set steel = provider.SteelMaterialFromSpec(spec)
        For Each shift In Array(0#, 1000000#, -1000000#)
            Set section = New CSectionModel
            For x = 0 To 9
                For y = 0 To 9
                    section.AddConcreteElement CDbl(shift) - 135# + 30# * x, CDbl(shift) / 4# - 90# + 20# * y, 600#
                Next y
            Next x
            section.AddRebarElement CDbl(shift) - 90#, CDbl(shift) / 4# - 60#, 20#, 0#, "A400"
            section.AddRebarElement CDbl(shift) + 90#, CDbl(shift) / 4# - 60#, 20#, 0#, "A400"
            section.AddRebarElement CDbl(shift) - 90#, CDbl(shift) / 4# + 60#, 20#, 0#, "A400"
            section.AddRebarElement CDbl(shift) + 90#, CDbl(shift) / 4# + 60#, 20#, 0#, "A400"
            Set cap = New CCapacitySolver
            cap.SolverToleranceN = 0.01: cap.SolverToleranceMx = 1#: cap.SolverToleranceMy = 1#
            cap.StrainTolerance = 0.0000001
            prefix = "forcePlateau." & CStr(loadSlot) & "." & CStr(extended) & "." & CStr(shift)
            TestSaveAudit03SearchProgress stats, "RUN: " & prefix
            cap.SolveByUltimateLoadPath section, concrete, steel, 0#, nBase, _
                mx, nBase * CDbl(shift) / 4#, my, nBase * CDbl(shift), False
            TestAssertTrue stats, prefix & ".converged", cap.Converged
            If cap.Converged Then
                TestAssertTrue stats, prefix & ".N", Abs(cap.LastSolver.Nint - cap.LambdaUltimate * nBase) <= 0.01
                TestAssertTrue stats, prefix & ".Mx", Abs(cap.LastSolver.Mxint - mx - _
                    cap.LambdaUltimate * nBase * CDbl(shift) / 4#) <= 1#
                TestAssertTrue stats, prefix & ".My", Abs(cap.LastSolver.Myint - my - _
                    cap.LambdaUltimate * nBase * CDbl(shift)) <= 1#
                TestAssertTrue stats, prefix & ".physical", physical.StateWithinPhysicalRange(section, cap.LastSolver, concrete, steel)
                ' N + Mx/90 <= 4*As*Rs: бетон только сжат, а все y >= -90.
                ' Предел достижим при текучести всех стержней и сжатии нижнего
                ' ряда бетона. Независимый контроль не зависит от Newton.
                TestAssertClose stats, prefix & ".forceBound", cap.NUltimate, _
                    4# * section.RebarArea(1) * 350# - mx / 90#, 0.02
                If referenceLambda = 0# Then referenceLambda = cap.LambdaUltimate
                TestAssertTrue stats, prefix & ".invariant", Abs(cap.LambdaUltimate - referenceLambda) * nBase <= 0.02
            Else
                TestAppendLine stats, prefix & "|" & cap.StopReason & vbCrLf & cap.DiagnosticLog
            End If
            TestSaveAudit03SearchProgress stats, "DONE: " & prefix
        Next shift
    Next extended
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' плато на той же траектории силы с ненулевыми моментами.
Public Function RunForcePathPlateauTests() As String
    Dim stats As TCapacityTestStats
    TestForcePathMaterialPlateau stats
    TestForcePathMaterialPlateau stats, 5
    TestAppendLine stats, "TOTAL_FORCE_PATH_PLATEAU: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunForcePathPlateauTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' буферизация не меняет строки, порядок, повторное чтение и reset.
Public Function RunCapacityDiagnosticBufferTests() As String
    Dim stats As TCapacityTestStats
    TestCapacityDiagnosticBuffer stats
    TestAppendLine stats, "TOTAL_CAPACITY_DIAGNOSTIC_BUFFER: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunCapacityDiagnosticBufferTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет рост, повторное чтение, отключение, объединение и сброс журнала поиска без потери записей.
Private Sub TestCapacityDiagnosticBuffer(ByRef stats As TCapacityTestStats)
    Dim cap As CCapacitySolver, expected As String, i As Long
    Set cap = New CCapacitySolver
    cap.LimitSearchAppendCapacityDiagnostic "first"
    expected = "search=first" & vbCrLf
    TestAssertEquals stats, "diagnosticBuffer.first", cap.DiagnosticLog, expected
    TestAssertEquals stats, "diagnosticBuffer.repeat", cap.DiagnosticLog, expected
    For i = 1 To 1025
        cap.LimitSearchAppendCapacityDiagnostic CStr(i)
        expected = expected & "search=" & CStr(i) & vbCrLf
    Next i
    TestAssertEquals stats, "diagnosticBuffer.growth", cap.DiagnosticLog, expected
    cap.DiagnosticsEnabled = False
    cap.LimitSearchAppendCapacityDiagnostic "ignored"
    TestAssertEquals stats, "diagnosticBuffer.disabled", cap.DiagnosticLog, expected
    cap.DiagnosticsEnabled = True
    cap.LimitSearchMergeCapacityFallback "A", "B", "prior" & vbCrLf, "first", "next", vbNullString
    expected = "first; статус=A; причина=B" & vbCrLf & "prior" & vbCrLf & "next" & vbCrLf & expected
    TestAssertEquals stats, "diagnosticBuffer.merge", cap.DiagnosticLog, expected
    cap.LimitSearchAppendCapacityDiagnostic "last"
    TestAssertEquals stats, "diagnosticBuffer.appendAfterMerge", cap.DiagnosticLog, expected & "search=last" & vbCrLf
    cap.LimitSearchPrepareCapacityUltimate Nothing, Nothing, Nothing, Nothing
    TestAssertEquals stats, "diagnosticBuffer.reset", cap.DiagnosticLog, vbNullString
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отдельная регрессия масштаба нормы без изменения допусков приемки.
Public Function RunCapacityMomentMeritTests() As String
    Dim stats As TCapacityTestStats
    TestCapacityMomentMeritScaling stats
    TestCapacityJacobianMeritSelection stats
    TestAppendLine stats, "TOTAL_CAPACITY_MOMENT_MERIT: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunCapacityMomentMeritTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' выбор масштаба зависит только от численной формы пути, не от сечения.
Private Sub TestCapacityJacobianMeritSelection(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder, rebars As CRebarLayout, section As CSectionModel
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars
    Set section = TestBuildGeneratedSectionModel(mesh, rebars)
    Dim cap As CCapacitySolver, problem As CCapacityLimitSearchProblem, callback As ILimitSearchProblem
    Set cap = New CCapacitySolver: TestConfigureCapacity cap
    Set problem = New CCapacityLimitSearchProblem: problem.Initialize cap: Set callback = problem
    Dim values As Variant, item As Variant, path As CLoadPathVector, index As Long, prefix As String
    values = Array(Array(-100000#, 0#, 0#, 100#, 0#, 200#, True), _
        Array(0#, -100000#, 0#, 100#, 0#, 200#, False), _
        Array(0#, -100000#, 100#, 0#, -200#, 0#, True), _
        Array(0#, 0#, 0#, 100#, 0#, 200#, True), _
        Array(0#, -100000#, 0#, 0#, 0#, 0#, False), _
        Array(-100000#, 100000#, 0#, 100#, 0#, 200#, True))
    For Each item In values
        index = index + 1: prefix = "capacity.merit.selection." & CStr(index)
        Set path = New CLoadPathVector
        path.Initialize item(0), item(1), item(2), item(3), item(4), item(5)
        TestAssertTrue stats, prefix & ".prepared", cap.LimitSearchPrepareCapacityUltimate( _
            section, TestLinearConcrete(), TestLinearSteel(), path)
        TestAssertTrue stats, prefix & ".scale", callback.LimitSearchUseJacobianMeritScaling = CBool(item(6))
    Next item
    cap.LimitSearchPrepareCapacityUltimate Nothing, Nothing, Nothing, Nothing
    TestAssertTrue stats, "capacity.merit.selection.invalidReset", Not callback.LimitSearchUseJacobianMeritScaling
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' конечность, typed failures и реальные Capacity-маршруты.
Public Function TestRunAudit03UltimateContracts() As String
    Dim stats As TCapacityTestStats
    TestAudit03UltimateGuards stats
    TestAudit03CapacityTypedFailures stats
    TestScaledUltimateStepAcceptance stats
    TestRejectedUltimateFinalPoint stats
    TestAppendLine stats, "TOTAL_AUDIT03_ULTIMATE_CONTRACTS: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    TestRunAudit03UltimateContracts = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' оба domain-kind используют один общий Newton, без выдуманных
' материалов. Критерий и исходная точность линейной задачи не ослабляются.
Private Sub TestAudit03UltimateGuards(ByRef stats As TCapacityTestStats)
    Dim kind As Variant, scenario As Long, guardScenario As Variant
    For Each kind In Array(rkCapacity, rkCrackFormation)
        For Each guardScenario In Array(1, 4, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21)
            scenario = CLng(guardScenario)
            TestAudit03UltimateGuardCase stats, CLng(kind), scenario
        Next guardScenario
    Next kind
    TestAudit02GenericUltimateSearch stats
    TestUltimateFinalAllowedCorrection stats
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Последняя поправка общего поиска проходит те же residual и domain-проверки.
Private Sub TestUltimateFinalAllowedCorrection(ByRef stats As TCapacityTestStats)
    Dim kind As Variant, problem As CTestLimitSearchProblem, path As CLoadPathVector
    Dim search As CUltimateStrainSearch
    For Each kind In Array(rkCapacity, rkCrackFormation)
        Set problem = New CTestLimitSearchProblem
        problem.TestConfigure "Bisection", 1.25, 0.000000001, 1, resultKind:=CLng(kind)
        Set path = New CLoadPathVector: path.Initialize 0#, 1#, 0#, 0#, 0#, 0#
        Set search = New CUltimateStrainSearch
        TestAssertTrue stats, "ultimate.finalCorrection." & CStr(kind), _
            search.RunNewton(problem, Nothing, Nothing, Nothing, path, Nothing, "Не сошлось.")
        TestAssertClose stats, "ultimate.finalCorrection." & CStr(kind) & ".lambda", problem.TestFinalLambda, 1.25, 0.00000001
        TestAssertTrue stats, "ultimate.finalCorrection." & CStr(kind) & ".budget", problem.TestFinalIterations = 1
    Next kind
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет реальный callback call path, отсутствие фиктивной
' точки, точную причину и конечное число проб. Nothing context - нарушение API.
Private Sub TestAudit03UltimateGuardCase(ByRef stats As TCapacityTestStats, _
        ByVal kind As EResultKind, ByVal scenario As Long)
    On Error GoTo Failed
    Dim prefix As String
    prefix = "audit03.ultimate." & CStr(kind) & "." & CStr(scenario)
    TestSaveAudit03SearchProgress stats, "START: " & prefix
    Dim problem As CTestLimitSearchProblem
    Set problem = New CTestLimitSearchProblem
    problem.TestConfigureUltimateGuardCase scenario, kind
    Dim search As CUltimateStrainSearch
    Set search = New CUltimateStrainSearch
    Dim path As CLoadPathVector
    Set path = New CLoadPathVector
    path.Initialize 0#, 1#, 0#, 0#, 0#, 0#
    If scenario = 12 Then
        On Error Resume Next
        Dim ignored As Boolean
        ignored = search.RunNewton(Nothing, Nothing, Nothing, Nothing, path, Nothing, "Тестовый поиск не сошелся.")
        Dim errorNumber As Long
        errorNumber = Err.Number
        On Error GoTo Failed
        TestAssertTrue stats, prefix & ".controlledContractError", errorNumber = vbObjectError + 4212
        TestSaveAudit03SearchProgress stats, "DONE: " & prefix
        Exit Sub
    End If
    If scenario = 13 Then Set path = Nothing
    TestAssertTrue stats, prefix & ".notAccepted", _
        Not search.RunNewton(problem, Nothing, Nothing, Nothing, path, Nothing, "Тестовый поиск не сошелся.")
    Dim expectedFailure As ESolverFailureCode
    expectedFailure = sfcInvalidConfiguration
    Select Case scenario
        Case 1, 11, 14, 16, 17, 18, 20: expectedFailure = sfcNumericalFailure
        Case 8, 10, 13, 15, 19, 21: expectedFailure = sfcInternalError
        Case 9: expectedFailure = sfcSingularTangent
    End Select
    TestAssertTrue stats, prefix & ".typedCause", problem.TestFailureCode = expectedFailure
    TestAssertTrue stats, prefix & ".noPoint", Not problem.TestConverged
    TestAssertTrue stats, prefix & ".finiteProbes", problem.TestUltimateCalls <= NONLINEAR_MAX_BACKTRACKS + 4
    If scenario = 4 Or scenario = 13 Then TestAssertTrue stats, prefix & ".noNumericalAttempt", problem.TestUltimateCalls = 0
    If scenario >= 7 And scenario <= 11 Then TestAssertTrue stats, prefix & ".failedProbeNotOverwritten", problem.TestUltimateCalls = 2
    If scenario >= 14 And scenario <= 17 Then TestAssertTrue stats, prefix & ".failedProbeNotOverwritten", problem.TestUltimateCalls = 2
    If scenario >= 18 Then TestAssertTrue stats, prefix & ".failedGuessNotRetried", problem.TestUltimateCalls = 0
    Dim callback As ILimitSearchProblem
    Set callback = problem
    Dim request As CLimitSearchRequest
    Set request = New CLimitSearchRequest
    request.InitializeWithProblem problem, kind, "UltimateStrain", 0#, 1#, 0#, 0#, 0#, 0#
    Dim result As CLimitSearchResult
    Set result = callback.BuildResult(request, False)
    Dim expectedMeta As CResultMeta
    Set expectedMeta = New CResultMeta
    expectedMeta.SetSolverFailure expectedFailure, kind, "Контроль typed-причины.", _
        (problem.TestUltimateCalls > 0 Or scenario >= 18)
    TestAssertTrue stats, prefix & ".status", result.Meta.InternalStatus = expectedMeta.InternalStatus
    TestAssertTrue stats, prefix & ".code", result.Meta.ResultCode = expectedMeta.ResultCode
    TestAssertTrue stats, prefix & ".resultKind", result.Meta.ResultKind = kind
    TestAssertTrue stats, prefix & ".noPhysicalPoint", Not result.HasLimitPoint
    TestAssertTrue stats, prefix & ".lifecycle", result.Meta.Calculated = expectedMeta.Calculated
    TestAssertTrue stats, prefix & ".reason", Len(result.Meta.ResultComment) > 0
    ' Последний SetFailure проверяется независимо от fake-result, который
    ' включает весь DiagnosticLog и мог бы скрыть потерю исходной причины.
    Dim callbackReason As String, runtimeCode As Long, standardDescription As String
    Select Case scenario
        Case 10: callbackReason = "Тестовое нарушение программного контракта.": runtimeCode = 9
        Case 11: callbackReason = "Тестовое переполнение численной операции.": runtimeCode = 6
        Case 14, 18: runtimeCode = 6
        Case 15, 19: runtimeCode = 9
        Case 16: runtimeCode = 11
        Case 17: callbackReason = "Тестовое деление на ноль.": runtimeCode = 11
        Case 20: callbackReason = "Тестовое переполнение стартовой плоскости.": runtimeCode = 6
        Case 21: callbackReason = "Тестовое нарушение подготовки стартовой плоскости.": runtimeCode = 9
    End Select
    If runtimeCode > 0 Then
        TestAssertTrue stats, prefix & ".rawRuntimeCode", _
            InStr(1, problem.TestDiagnosticLog, "ultimateRuntimeError=" & CStr(runtimeCode), vbBinaryCompare) > 0
        TestAssertTrue stats, prefix & ".rawRuntimeStage", InStr(1, problem.TestDiagnosticLog, "; stage=", vbBinaryCompare) > 0
        If Len(callbackReason) > 0 Then
            TestAssertTrue stats, prefix & ".exactCallbackReason", InStr(1, problem.TestFailureReason, callbackReason, vbBinaryCompare) > 0
            TestAssertTrue stats, prefix & ".rawCallbackReason", InStr(1, problem.TestDiagnosticLog, "; description=" & callbackReason, vbBinaryCompare) > 0
        Else
            Err.Clear
            standardDescription = Error$(runtimeCode)
            TestAssertTrue stats, prefix & ".rawStandardReason", _
                InStr(1, problem.TestDiagnosticLog, "; description=" & standardDescription, vbBinaryCompare) > 0
            If expectedFailure = sfcNumericalFailure Then
                TestAssertTrue stats, prefix & ".localizedReason", InStr(1, problem.TestFailureReason, "диапазон арифметики", vbTextCompare) > 0
            Else
                TestAssertTrue stats, prefix & ".localizedReason", InStr(1, problem.TestFailureReason, "внутренняя ошибка вычислений", vbTextCompare) > 0
            End If
        End If
    End If
    TestAppendLine stats, "COMMENT: " & prefix & "; calls=" & CStr(problem.TestUltimateCalls) & "; " & result.Meta.ResultComment
    ' Новый запуск на той же задаче обязан сбросить failure и стартовые guards;
    ' успешная линейная система не наследует причину предыдущей остановки.
    problem.TestConfigure "Bisection", 1.25, 0.000000001, 20, resultKind:=kind
    If path Is Nothing Then Set path = New CLoadPathVector
    path.Initialize 0#, 1#, 0#, 0#, 0#, 0#
    TestAssertTrue stats, prefix & ".recovery", search.RunNewton(problem, Nothing, Nothing, Nothing, path, Nothing, "Тестовый поиск не сошелся.")
    Set result = callback.BuildResult(request, True)
    TestAssertTrue stats, prefix & ".recoveryStatus", result.Succeeded And result.Meta.InternalStatus = rsSuccess
    TestAssertClose stats, prefix & ".recoveryLambda", result.LambdaUltimate, 1.25, 0.00000001
    TestAssertTrue stats, prefix & ".recoveryReasonCleared", problem.TestFailureCode = sfcNone And Len(problem.TestFailureReason) = 0
    TestSaveAudit03SearchProgress stats, "DONE: " & prefix
    Exit Sub
Failed:
    TestAssertTrue stats, prefix & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
    TestSaveAudit03SearchProgress stats, "FAILED: " & prefix
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет соседние Double, большие положительные границы и отсутствие
' прогресса recovery для обоих инженерных потребителей каждого 1D-метода.
Private Sub TestAudit03SearchArithmetic(ByRef stats As TCapacityTestStats)
    Dim kind As Variant, methodName As Variant, scenario As Long
    For Each kind In Array(rkCapacity, rkCrackFormation)
        For Each methodName In Array("Bisection", "Brent", "Secant")
            For scenario = 1 To 4
                TestAudit03SearchArithmeticCase stats, CStr(methodName), CLng(kind), scenario
            Next scenario
        Next methodName
    Next kind
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Реальная generic-задача различает невозможную точность, техническую границу,
' точный предел на MaxLambda и ошибочный recovery без фальшивого State.
Private Sub TestAudit03SearchArithmeticCase(ByRef stats As TCapacityTestStats, _
        ByVal methodName As String, ByVal kind As EResultKind, ByVal scenario As Long)
    On Error GoTo Failed
    Dim problem As CTestLimitSearchProblem
    Set problem = New CTestLimitSearchProblem
    Dim prefix As String
    prefix = "audit03.search." & CStr(kind) & "." & methodName & "." & CStr(scenario)
    TestSaveAudit03SearchProgress stats, "START: " & prefix
    Dim adjacent As Double
    adjacent = 1# + 2# ^ (-52)
    Select Case scenario
        Case 1
            TestAssertTrue stats, prefix & ".adjacentDistinct", adjacent > 1# And adjacent - 1# < 0.000000000000001
            problem.TestConfigure methodName, adjacent, 1E-18, 80, 1#, adjacent, kind
            If methodName = "Bisection" Then problem.TestConfigure methodName, adjacent, 1E-18, 0, 1#, adjacent, kind
            problem.TestStepCriterion = True
        Case 2
            problem.TestConfigure methodName, 1.2E+308, 1E+294, 120, 1E+308, 1.4E+308, kind
        Case 3
            problem.TestConfigure methodName, 3#, 0.000001, 80, 1#, 6#, kind
            problem.TestFailureAbove = 0.5
            problem.TestRecoveryWithoutProgress = True
        Case 4
            problem.TestConfigure methodName, 6#, 0.000001, 80, 1#, 6#, kind
    End Select
    Dim request As CLimitSearchRequest
    Set request = New CLimitSearchRequest
    request.InitializeWithProblem problem, kind, "LoadMultiplier", 0#, 1#, 0#, 0#, 0#, 0#
    Dim search As CLoadMultiplierSearch
    Set search = New CLoadMultiplierSearch
    Dim result As CLimitSearchResult
    Set result = search.Execute(request)
    If scenario = 1 Or scenario = 3 Then
        TestAssertTrue stats, prefix & ".numerical", result.Meta.InternalStatus = rsNumericalFailure
        TestAssertTrue stats, prefix & ".notBound", result.Meta.ResultCode <> rcSearchBoundReached
        TestAssertTrue stats, prefix & ".noPoint", Not result.HasLimitPoint
        TestAssertTrue stats, prefix & ".noFakeLimit", result.LambdaUltimate = 0#
        TestAssertTrue stats, prefix & ".finiteWork", problem.TestProbeCalls <= 120
        TestAssertTrue stats, prefix & ".reason", Len(result.DiagnosticLog) > 0
    Else
        TestAssertTrue stats, prefix & ".success", result.Succeeded And result.HasLimitPoint
        If scenario = 2 Then
            TestAssertTrue stats, prefix & ".accuracy", Abs(result.LambdaUltimate / 1.2E+308 - 1#) <= 0.00000000000002
            TestAssertTrue stats, prefix & ".bounded", problem.TestProbeMaxLambda <= 1.4E+308
        Else
            TestAssertClose stats, prefix & ".maxLambda", result.LambdaUltimate, 6#, 0.000001
            TestAssertClose stats, prefix & ".lastProbe", problem.TestProbeMaxLambda, 6#, 0#
        End If
    End If
    TestAppendLine stats, "COMMENT: " & prefix & "; " & result.Meta.ResultComment & "; " & result.DiagnosticLog
    TestSaveAudit03SearchProgress stats, "DONE: " & prefix
    Exit Sub
Failed:
    TestAssertTrue stats, prefix & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
    TestSaveAudit03SearchProgress stats, "FAILED: " & prefix
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет уже выполненные assertions до следующего COM-шага.
' При модальной ошибке или watchdog виден конкретный случай, а не только suite.
Private Sub TestSaveAudit03SearchProgress(ByRef stats As TCapacityTestStats, ByVal stage As String)
    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open ThisWorkbook.Path & "\Audit03_Search_Progress.txt" For Output As #fileNumber
    Print #fileNumber, stats.Report
    Print #fileNumber, stage
    Close #fileNumber
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ============================== ДЛЯ ТЕСТОВ ==============================
' проверяет реальный lifecycle Capacity/Search, в том числе снимок
' до запуска и reset одного solver-а после успешной численной попытки.
Public Function TestRunAudit03CapacityLifecycle() As String
    On Error GoTo Failed
    Dim stats As TCapacityTestStats
    TestAudit03CapacitySearchLifecycle stats
    TestAppendLine stats, "TOTAL_AUDIT03_CAPACITY_LIFECYCLE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    TestRunAudit03CapacityLifecycle = stats.Report
    Exit Function
Failed:
    TestRunAudit03CapacityLifecycle = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' оба численных метода действительно ищут предел, а невалидные
' допуски/бюджет и отсутствующая зависимость останавливаются до Search.
' При повторе старый результат должен сохранять независимый факт выполнения.
Private Sub TestAudit03CapacitySearchLifecycle(ByRef stats As TCapacityTestStats)
    Dim problem As CCapacityLimitSearchProblem, result As CLimitSearchResult
    Set problem = New CCapacityLimitSearchProblem
    Set result = problem.BuildSnapshot("Auto")
    TestAssertTrue stats, "audit03.capacity.lifecycle.missing.notExecuted", Not result.SearchExecuted
    TestAssertTrue stats, "audit03.capacity.lifecycle.missing.notCalculated", Not result.Meta.Calculated
    Dim mesh As CFiberMeshBuilder, rebars As CRebarLayout, section As CSectionModel
    TestPrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars
    Set section = TestBuildGeneratedSectionModel(mesh, rebars)
    Dim cap As CCapacitySolver, retained As CLimitSearchResult, strategy As Variant
    Set cap = New CCapacitySolver
    Set result = TestAudit02CapacitySnapshot(cap, "Auto")
    TestAssertTrue stats, "audit03.capacity.lifecycle.empty.notExecuted", Not result.SearchExecuted
    TestAssertTrue stats, "audit03.capacity.lifecycle.empty.notCalculated", Not result.Meta.Calculated
    For Each strategy In Array("LoadMultiplier", "UltimateStrain")
        TestConfigureCapacity cap
        If CStr(strategy) = "LoadMultiplier" Then
            cap.SolveByLoadPathMultiplier section, TestLinearConcrete(), TestLinearSteel(), 0#, 0#, 0#, 10000000#, 0#, 0#
        Else
            cap.SolveByUltimateLoadPath section, TestLinearConcrete(), TestLinearSteel(), 0#, 0#, 0#, 10000000#, 0#, 0#
        End If
        Set result = TestAudit02CapacitySnapshot(cap, CStr(strategy))
        Set retained = result.Clone
        Dim prefix As String
        prefix = "audit03.capacity.lifecycle." & CStr(strategy)
        TestAssertTrue stats, prefix & ".attempted", result.SearchExecuted
        TestAssertTrue stats, prefix & ".calculated", result.Meta.Calculated
        TestAssertTrue stats, prefix & ".point", result.HasLimitPoint
        If CStr(strategy) = "LoadMultiplier" Then
            cap.LambdaTolerance = 0#
            cap.SolveByLoadPathMultiplier section, TestLinearConcrete(), TestLinearSteel(), 0#, 0#, 0#, 10000000#, 0#, 0#
        Else
            cap.SolverMaxIterations = 0
            cap.SolveByUltimateLoadPath section, TestLinearConcrete(), TestLinearSteel(), 0#, 0#, 0#, 10000000#, 0#, 0#
        End If
        Set result = TestAudit02CapacitySnapshot(cap, CStr(strategy))
        TestAssertTrue stats, prefix & ".invalidBeforeSearch", Not result.SearchExecuted
        TestAssertTrue stats, prefix & ".invalidNotCalculated", Not result.Meta.Calculated
        TestAssertTrue stats, prefix & ".invalidNoPoint", Not result.HasLimitPoint
        TestAssertTrue stats, prefix & ".priorSnapshot", retained.SearchExecuted And retained.HasLimitPoint
        cap.SolveByLoadPathMultiplier Nothing, TestLinearConcrete(), TestLinearSteel(), 0#, 0#, 0#, 10000000#, 0#, 0#
        Set result = TestAudit02CapacitySnapshot(cap, CStr(strategy))
        TestAssertTrue stats, prefix & ".missingReset", Not result.SearchExecuted
        TestAssertTrue stats, prefix & ".missingNoPoint", Not result.HasLimitPoint
    Next strategy
    ' Принятая последняя поправка завершает UltimateStrain. Неверный допуск
    ' неиспользованного LoadMultiplier не отменяет подтвержденную точку.
    TestConfigureCapacity cap
    cap.SolverMaxIterations = 1
    cap.LambdaTolerance = 0#
    Dim request As CLimitSearchRequest, coordinator As CLimitSearchCoordinator
    Set request = New CLimitSearchRequest: Set coordinator = New CLimitSearchCoordinator
    InitializeCapacitySearchRequest request, cap, section, TestProvisionalConcrete(), TestProvisionalSteel(), _
        "Auto", -20000#, 0#, 0#, 10000000#, 0#, 0#, False, False
    Set result = coordinator.ExecuteCapacity(request)
    TestAssertTrue stats, "audit03.capacity.lifecycle.fallback.firstAttemptKept", result.SearchExecuted
    TestAssertTrue stats, "audit03.capacity.lifecycle.fallback.finalSuccess", result.Succeeded
    TestAssertTrue stats, "audit03.capacity.lifecycle.fallback.point", result.HasLimitPoint
    TestAssertClose stats, "audit03.capacity.lifecycle.fallback.physicalLimit", cap.TestCriticalStrainUtilization, 1#, 0.00001
    TestAssertEquilibrium stats, "audit03.capacity.lifecycle.fallback.equilibrium", cap.LastSolver, _
        -20000#, result.LambdaUltimate * 10000000#, 0#
    TestConfigureCapacity cap ' Прямой содержательный API должен сохранять тот же lifecycle, что coordinator.
    cap.SolverMaxIterations = 1
    cap.LambdaTolerance = 0#
    cap.SolveByAutoLoadPath section, TestProvisionalConcrete(), TestProvisionalSteel(), _
        -20000#, 0#, 0#, 10000000#, 0#, 0#, False, False
    Set result = TestAudit02CapacitySnapshot(cap, "Auto")
    TestAssertTrue stats, "audit03.capacity.lifecycle.directAuto.firstAttemptKept", result.SearchExecuted
    TestAssertTrue stats, "audit03.capacity.lifecycle.directAuto.finalSuccess", result.Succeeded
    TestAssertTrue stats, "audit03.capacity.lifecycle.directAuto.point", result.HasLimitPoint
    TestAssertClose stats, "audit03.capacity.lifecycle.directAuto.physicalLimit", cap.TestCriticalStrainUtilization, 1#, 0.00001
    TestAssertEquilibrium stats, "audit03.capacity.lifecycle.directAuto.equilibrium", cap.LastSolver, _
        -20000#, result.LambdaUltimate * 10000000#, 0#
    ' Текущее сочетание осевое относительно расчетного центра, но содержит
    ' взаимно компенсирующиеся пользовательский момент и момент от N.
    ' Во всех трех Auto-путях остаются моментные компоненты траектории,
    ' хотя при lambda=1 они компенсированы. Предел найден на последнем
    ' Auto-пути; итог сохраняет историю неуспешных попыток и найденную точку.
    TestConfigureCapacity cap
    cap.SolverMaxIterations = 1
    cap.LambdaTolerance = 0#
    Dim autoLoad As CSectionLoadState
    Set autoLoad = New CSectionLoadState
    autoLoad.Initialize -20000#, 10000000#, 0#, 0#, 500#
    InitializeCapacitySearchRequest request, cap, section, TestProvisionalConcrete(), TestProvisionalSteel(), _
        "UltimateStrain", 0#, 0#, 0#, 1#, 0#, 0#, False, False
    request.ConfigureLoadPathSelection "Auto", autoLoad, True
    Set result = coordinator.ExecuteCapacity(request)
    TestAppendLine stats, "AUTO_PATH_LIFECYCLE: path=" & request.SelectedLoadPath.Key & _
        "; status=" & CStr(result.Meta.InternalStatus) & "; code=" & CStr(result.Meta.ResultCode) & _
        "; searched=" & CStr(result.SearchExecuted) & "; comment=" & result.Meta.ResultComment
    TestAppendLine stats, "AUTO_PATH_LIFECYCLE_DIAGNOSTICS: " & result.DiagnosticLog
    TestAssertTrue stats, "audit03.capacity.lifecycle.autoPaths.firstAttemptKept", result.SearchExecuted
    TestAssertTrue stats, "audit03.capacity.lifecycle.autoPaths.finalSuccess", result.Succeeded
    TestAssertTrue stats, "audit03.capacity.lifecycle.autoPaths.finalPathNMxy", request.SelectedLoadPath.Key = "LambdaNMxy"
    TestAssertTrue stats, "audit03.capacity.lifecycle.autoPaths.point", result.HasLimitPoint
    TestAssertEquilibrium stats, "audit03.capacity.lifecycle.autoPaths.equilibrium", cap.LastSolver, _
        -20000# * result.LambdaUltimate, 0#, 0#
    Dim fake As CTestLimitSearchProblem, multiplier As CLoadMultiplierSearch
    Set fake = New CTestLimitSearchProblem: Set multiplier = New CLoadMultiplierSearch
    fake.TestConfigure "Bisection", 1.25, 0.000001, 80
    request.InitializeWithProblem fake, rkCapacity, "LoadMultiplier", 0#, 1#, 0#, 0#, 0#, 0#
    Set result = multiplier.Execute(request)
    TestAssertTrue stats, "audit03.capacity.lifecycle.generic.executed", result.SearchExecuted
    fake.TestConfigure "Bisection", 1.25, 0#, 80
    Set result = multiplier.Execute(request)
    TestAssertTrue stats, "audit03.capacity.lifecycle.generic.invalidNotExecuted", Not result.SearchExecuted
    TestAssertTrue stats, "audit03.capacity.lifecycle.generic.invalidNoProbe", fake.TestProbeCalls = 0
    Set result = New CLimitSearchResult
    result.Initialize Nothing, "Auto", vbNullString, False, 0#, 0#, 0#, 0#, Nothing, vbNullString, vbNullString
    TestAssertTrue stats, "audit03.capacity.lifecycle.default.notExecuted", Not result.SearchExecuted
    TestAssertTrue stats, "audit03.capacity.lifecycle.default.notCalculated", Not result.Meta.Calculated
    ' Ранний отчет берет объяснение из собственного инженерного результата,
    ' а не из технического LimitState и не из несуществующего поля meta.
    Dim capacityResult As CCapacityResult, load As CSectionLoadState, path As CLoadPathDescriptor
    Set capacityResult = New CCapacityResult
    Set load = New CSectionLoadState: Set path = New CLoadPathDescriptor
    load.Initialize 0#, 10000000#, 0#, 0#, 0#
    path.InitializeFromLoadState "LambdaMx", load
    capacityResult.InitializeInvalidInput "В Config задайте положительный допуск поиска.", path.Key
    Dim earlyReport As String
    earlyReport = capacityResult.EarlyStopReport(path)
    TestAssertTrue stats, "audit03.capacity.lifecycle.report.reason", InStr(1, earlyReport, capacityResult.ResultMeta.ResultComment, vbBinaryCompare) > 0
    TestAssertTrue stats, "audit03.capacity.lifecycle.report.russian", InStr(1, earlyReport, "Поиск несущей способности", vbBinaryCompare) = 1
    TestAssertTrue stats, "audit03.capacity.lifecycle.report.noRawStatus", InStr(1, earlyReport, "InvalidInput", vbTextCompare) = 0
    TestAudit03CapacityConfigurationMessages stats, section
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' каждый невалидный поисковый параметр должен сохранить typed
' причину, назвать свою строку Config и завершиться до первой численной пробы.
Private Sub TestAudit03CapacityConfigurationMessages(ByRef stats As TCapacityTestStats, _
        ByVal section As CSectionModel)
    Dim key As Variant, cap As CCapacitySolver, result As CLimitSearchResult
    Dim prefix As String, strategy As String
    For Each key In Array("Capacity.InitialLambda", "Capacity.ToleranceLambda", _
            "Capacity.MaxLambda", "Capacity.MaxRetries", "Capacity.BaseLoadSteps", _
            "Capacity.ToleranceStrain")
        Set cap = New CCapacitySolver
        TestConfigureCapacity cap
        strategy = "LoadMultiplier"
        Select Case CStr(key)
            Case "Capacity.InitialLambda": cap.InitialLambdaStep = 0#
            Case "Capacity.ToleranceLambda": cap.LambdaTolerance = 0#
            Case "Capacity.MaxLambda": cap.MaxLambda = 0#
            Case "Capacity.MaxRetries": cap.MaxRetries = -1
            Case "Capacity.BaseLoadSteps": cap.SolverBaseLoadSteps = 0
            Case "Capacity.ToleranceStrain"
                cap.StrainTolerance = 0#
                strategy = "UltimateStrain"
        End Select
        If strategy = "UltimateStrain" Then
            cap.SolveByUltimateLoadPath section, TestLinearConcrete(), TestLinearSteel(), 0#, 0#, 0#, 10000000#, 0#, 0#
        Else
            cap.SolveByLoadPathMultiplier section, TestLinearConcrete(), TestLinearSteel(), 0#, 0#, 0#, 10000000#, 0#, 0#
        End If
        Set result = TestAudit02CapacitySnapshot(cap, strategy)
        prefix = "audit03.capacity.configMessage." & CStr(key)
        TestAssertTrue stats, prefix & ".status", result.Meta.InternalStatus = rsInvalidConfiguration
        TestAssertTrue stats, prefix & ".code", result.Meta.ResultCode = rcInvalidConfiguration
        TestAssertTrue stats, prefix & ".sheet", InStr(1, result.Meta.ResultComment, "Config", vbBinaryCompare) > 0
        TestAssertTrue stats, prefix & ".key", InStr(1, result.Meta.ResultComment, CStr(key), vbBinaryCompare) > 0
        TestAssertTrue stats, prefix & ".noSearch", Not result.SearchExecuted And cap.Iterations = 0
        TestAppendLine stats, "CAPACITY_CONFIG_MESSAGE: " & CStr(key) & "|" & result.Meta.ResultComment
    Next key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ============================== ДЛЯ ТЕСТОВ ==============================
' Проверяет фактический общий LoadMultiplier для обоих инженерных видов и
' всех трех методов. Исключение callback-а не должно выходить из Execute,
' а терминальная причина запрещает последующий recovery/finalization.
Public Function TestRunAudit03MultiplierTypedFaults() As String
    On Error GoTo Failed
    Dim stats As TCapacityTestStats
    TestAudit03MultiplierTypedFaults stats
    TestAppendLine stats, "TOTAL_AUDIT03_MULTIPLIER_TYPED: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    TestRunAudit03MultiplierTypedFaults = stats.Report
    Exit Function
Failed:
    TestRunAudit03MultiplierTypedFaults = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' каждый направленный отказ имеет отдельный expected enum,
' физическую точку не создает и затем проходит настоящий повтор успешного
' поиска на том же problem/request. Никакой текстовый mapping не используется.
Private Sub TestAudit03MultiplierTypedFaults(ByRef stats As TCapacityTestStats)
    Dim kind As Variant, method As Variant, scenario As Long
    Dim expectedStatus As EResultInternalStatus, expectedCode As EResultCode
    Dim expectedFailure As ESolverFailureCode, policy As CResultStatusPolicy
    Dim fake As CTestLimitSearchProblem, request As CLimitSearchRequest
    Dim search As CLoadMultiplierSearch, result As CLimitSearchResult, prefix As String
    Dim errorNumber As Long, errorText As String, expectedDisplay As String
    Set policy = New CResultStatusPolicy
    For Each kind In Array(rkCapacity, rkCrackFormation)
        For Each method In Array("Bisection", "Brent", "Secant")
            For scenario = 1 To 16
                Set fake = New CTestLimitSearchProblem: Set request = New CLimitSearchRequest
                Set search = New CLoadMultiplierSearch
                fake.TestConfigureMultiplierFaultCase scenario, CStr(method), CLng(kind)
                request.InitializeWithProblem fake, CLng(kind), "LoadMultiplier", 0#, 1#, 0#, 0#, 0#, 0#
                prefix = "audit03.multiplierTyped." & CStr(kind) & "." & CStr(method) & "." & CStr(scenario)
                expectedStatus = rsInternalError: expectedCode = rcInternalError
                expectedFailure = sfcInternalError: expectedDisplay = "CalcErr"
                Select Case scenario
                    Case 3, 16
                        expectedStatus = rsNumericalFailure: expectedCode = rcNumericalFailure
                        expectedFailure = sfcNumericalFailure: expectedDisplay = "NumFail"
                    Case 4, 7, 9, 11, 12
                        expectedStatus = rsInvalidConfiguration: expectedCode = rcInvalidConfiguration
                        expectedFailure = sfcInvalidConfiguration: expectedDisplay = "InputErr"
                    Case 5
                        expectedStatus = rsInvalidInput: expectedCode = rcInvalidInput
                        expectedFailure = sfcInvalidInput: expectedDisplay = "InputErr"
                End Select
                errorNumber = 0: errorText = vbNullString
                Set result = TestAudit03ExecuteMultiplierFault(search, request, errorNumber, errorText)
                TestAssertTrue stats, prefix & ".noUnhandled", errorNumber = 0
                TestAssertTrue stats, prefix & ".resultExists", Not result Is Nothing
                If Not result Is Nothing Then
                    TestAssertTrue stats, prefix & ".status", result.Meta.InternalStatus = expectedStatus
                    TestAssertTrue stats, prefix & ".code", result.Meta.ResultCode = expectedCode
                    TestAssertTrue stats, prefix & ".display", policy.ExternalStatus(result.Meta) = expectedDisplay
                    TestAssertTrue stats, prefix & ".noPoint", Not result.HasLimitPoint And Not fake.TestConverged
                    TestAssertTrue stats, prefix & ".typedCause", fake.TestFailureCode = expectedFailure
                    TestAssertTrue stats, prefix & ".comment", InStr(1, result.Meta.ResultComment, _
                        "Контрольный отказ LoadMultiplier, сценарий " & CStr(scenario), vbBinaryCompare) > 0
                    If scenario = 1 Or scenario = 2 Or scenario = 3 Or scenario = 13 Or scenario = 14 Or scenario = 16 Then
                        TestAssertTrue stats, prefix & ".exactCallbackReason", InStr(1, fake.TestFailureReason, _
                            "Контрольный отказ LoadMultiplier, сценарий " & CStr(scenario), vbBinaryCompare) > 0
                        TestAssertTrue stats, prefix & ".rawRuntimeDiagnostic", _
                            InStr(1, result.DiagnosticLog, "loadMultiplierRuntimeError=", vbBinaryCompare) > 0
                    End If
                    If scenario = 13 Or scenario = 14 Then
                        TestAssertTrue stats, prefix & ".notCalculated", Not result.Meta.Calculated And Not result.SearchExecuted
                    ElseIf expectedStatus = rsInvalidInput Or expectedStatus = rsInvalidConfiguration Then
                        TestAssertTrue stats, prefix & ".validationLifecycle", Not result.Meta.Calculated And result.SearchExecuted
                    Else
                        TestAssertTrue stats, prefix & ".calculated", result.Meta.Calculated And result.SearchExecuted
                    End If
                Else
                    TestAppendLine stats, "FAULT_EXCEPTION: " & prefix & "; " & CStr(errorNumber) & "; " & errorText
                End If
                If expectedStatus <> rsNumericalFailure And scenario <> 15 Then
                    TestAssertTrue stats, prefix & ".noTerminalFallback", fake.TestFallbackCalls = 0
                End If
                TestAssertTrue stats, prefix & ".finiteWork", fake.TestProbeCalls < 100
                fake.TestConfigure CStr(method), 1.25, 0.000001, 80, resultKind:=CLng(kind)
                Set result = search.Execute(request)
                TestAssertTrue stats, prefix & ".recoverySuccess", result.Meta.InternalStatus = rsSuccess And result.HasLimitPoint
                TestAssertTrue stats, prefix & ".recoveryCauseReset", fake.TestFailureCode = sfcNone
                TestAssertTrue stats, prefix & ".recoveryCommentClean", InStr(1, result.Meta.ResultComment, "Контрольный отказ", vbBinaryCompare) = 0
            Next scenario
        Next method
    Next kind
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' перехватывает только исключение проверяемого публичного Execute,
' чтобы отрицательный reproducer дошел до всех сценариев и записал каждый
' необработанный отказ. Такой отказ никогда не считается успешным result.
Private Function TestAudit03ExecuteMultiplierFault(ByVal search As CLoadMultiplierSearch, _
        ByVal request As CLimitSearchRequest, ByRef errorNumber As Long, ByRef errorText As String) As CLimitSearchResult
    On Error GoTo Failed
    Set TestAudit03ExecuteMultiplierFault = search.Execute(request)
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description
End Function
















' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет только повторное заполнение request без решения равновесия.
' Инициализация трещин должна убрать данные прежнего запроса несущей способности.
Public Function RunLimitRequestLifecycleTests() As String
    On Error GoTo Failed
    Dim stats As TCapacityTestStats
    TestRequestReinitialization stats
    TestAppendLine stats, "TOTAL_LIMIT_REQUEST_LIFECYCLE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunLimitRequestLifecycleTests = stats.Report
    Exit Function
Failed:
    RunLimitRequestLifecycleTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сопоставляет переиспользованный запрос со свежим для трех прежних осевых
' режимов. Настройки/объекты нового расчета сохраняются точно; solver не вызывается.
Private Sub TestRequestReinitialization(ByRef stats As TCapacityTestStats)
    Dim problem As CTestLimitSearchProblem, request As CLimitSearchRequest, fresh As CLimitSearchRequest
    Dim section As CSectionModel, concrete As Object, steel As Object, prior As Variant, prefix As String
    Set problem = New CTestLimitSearchProblem: Set section = New CSectionModel
    Set concrete = TestLinearConcrete(): Set steel = TestLinearSteel()
    For Each prior In Array(Array(False, False), Array(True, False), Array(True, True))
        Set request = New CLimitSearchRequest: Set fresh = New CLimitSearchRequest
        request.InitializeCapacity problem, section, concrete, steel, "LoadMultiplier", 11#, 12#, 13#, 14#, 15#, 16#, CBool(prior(0)), CBool(prior(1))
        request.InitializeCrackFormation problem, section, concrete, steel, "UltimateStrain", "LambdaMxy", 17#, 18#, 19#, 20#, 21#, 22#, 23#
        fresh.InitializeCrackFormation problem, section, concrete, steel, "UltimateStrain", "LambdaMxy", 17#, 18#, 19#, 20#, 21#, 22#, 23#
        prefix = "request.reinitialize." & CStr(prior(0)) & "." & CStr(prior(1))
        TestAssertClose stats, prefix & ".nOffset", request.NOffset, fresh.NOffset, 0#
        TestAssertClose stats, prefix & ".nBase", request.NBase, fresh.NBase, 0#
        TestAssertClose stats, prefix & ".mxOffset", request.MxOffset, fresh.MxOffset, 0#
        TestAssertClose stats, prefix & ".mxBase", request.MxBase, fresh.MxBase, 0#
        TestAssertClose stats, prefix & ".myOffset", request.MyOffset, fresh.MyOffset, 0#
        TestAssertClose stats, prefix & ".myBase", request.MyBase, fresh.MyBase, 0#
        TestAssertTrue stats, prefix & ".forceOnly", request.ForceOnly = fresh.ForceOnly
        TestAssertTrue stats, prefix & ".pureAxial", request.PureAxialForMethod = fresh.PureAxialForMethod
        TestAssertClose stats, prefix & ".targetN", request.TargetN, fresh.TargetN, 0#
        TestAssertClose stats, prefix & ".targetMx", request.TargetMx, fresh.TargetMx, 0#
        TestAssertClose stats, prefix & ".targetMy", request.TargetMy, fresh.TargetMy, 0#
        TestAssertClose stats, prefix & ".momentMx", request.MomentMxAtReference, fresh.MomentMxAtReference, 0#
        TestAssertClose stats, prefix & ".momentMy", request.MomentMyAtReference, fresh.MomentMyAtReference, 0#
        TestAssertClose stats, prefix & ".referenceX", request.ReferenceX, fresh.ReferenceX, 0#
        TestAssertClose stats, prefix & ".referenceY", request.ReferenceY, fresh.ReferenceY, 0#
        TestAssertTrue stats, prefix & ".kind", request.ResultKind = fresh.ResultKind
        TestAssertTrue stats, prefix & ".strategy", request.SolutionStrategy = fresh.SolutionStrategy
        TestAssertTrue stats, prefix & ".problem", request.Problem Is fresh.Problem
        TestAssertTrue stats, prefix & ".section", request.Section Is fresh.Section
        TestAssertTrue stats, prefix & ".concrete", request.ConcreteMaterial Is fresh.ConcreteMaterial
        TestAssertTrue stats, prefix & ".steel", request.SteelMaterial Is fresh.SteelMaterial
        TestAssertTrue stats, prefix & ".selection", Not request.HasLoadPathSelection And Not fresh.HasLoadPathSelection
    Next prior
End Sub
