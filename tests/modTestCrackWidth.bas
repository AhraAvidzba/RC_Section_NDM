Attribute VB_Name = "modTestCrackWidth"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: расчёт нормальных трещин
' ==========================================================================
' Проверяет выбор растянутой области, ширину трещины по СП 63, способы задания
' psis, центральное растяжение и запись результата. Расчёт трещин использует
' обычный CSectionSolver; отдельного решения равновесия в тесте нет.

Private Type TCrackTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет материалы, образование трещины, подготовку областей и формулу ширины.
' Возвращает отчёт отдельных проверок; ошибка подготовки теста сохраняется отдельно.
Public Function RunCrackWidthTests() As String
    On Error GoTo Failed

    Dim stats As TCrackTestStats
    Dim t0 As Double
    t0 = Timer

    TestConcreteTensionBranches stats
    TestCrackWidthFormulaCalculatorPure stats
    TestCrackUserPsiMx stats
    TestCrackUserPsiMxy stats
    TestCrackUserCoefficients stats
    TestCoverModeNormalization stats
    TestEffectiveAndFullTensionZones stats
    TestAutoPsiSkipsLambdaWhenFirstCheckPasses stats
    TestAlwaysCalcPsiAppliesSigmaCrcWhenAutoPasses stats
    TestSigmaSCrcAveragingModeAllSelected stats
    TestAutoPsiAndLambdaAfterFailedFirstCheck stats
    TestAutoMcrcPureBendingConverges stats
    TestAutoMcrcFixedNIndependentOfMomentMagnitude stats
    TestCrackInitiationLoadPaths stats
    TestFormationFailureStillCalculatesWidth stats
    TestFormationFallbackWithUniformCurrentState stats
    TestNearUniformCrackTranslations stats
    TestAudit03ConfirmedNotCrackedSkipsWidth stats
    TestCrackFormationSearchBoundKeepsTechnicalCode stats
    TestCrackFormationNoCrackDoesNotBuildPostState stats
    TestCrackFormationCacheHitWithoutLastRunner stats
    TestLimitSearchResultKeepsCrackDiagnosticSnapshot stats
    TestDangerousLoadsDoNotNumFail stats
    TestAutoMcrcOneSignTensionUsesFormula854 stats
    TestCentralTensionBranch stats
    TestNoTensionRebar stats
    TestAudit02FormationOutcomeSemantics stats
    TestAudit02IndependentFormation stats
    TestAudit02PsiSignedInputsAndFallbackModes stats
    TestAudit03FormationTypedFailures stats
    TestAudit03FormationSearchLifecycle stats
    TestAudit03FormationResidualContracts stats
    TestAudit03FormationAdapterFailures stats
    TestAudit03ExtremeFormationInputs stats
    TestAppendLine stats, "TOTAL_CRACK: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & TestFormatNumberInvariant(Timer - t0)
    RunCrackWidthTests = stats.Report
    Exit Function

Failed:
    RunCrackWidthTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function



' -------------------- ДЛЯ ТЕСТОВ --------------------
' ------------------------------
' Материал бетона
' ------------------------------
' Проверяем, что растянутую часть диаграммы не обязательно задавать руками:
' Mcrc-диаграмма provider-а включает растянутую ветвь бетона по параметрам II ГПС.
Private Sub TestConcreteTensionBranches(ByRef stats As TCrackTestStats)
    Dim concrete As CMaterialDiagram
    Set concrete = TestProvisionalConcrete()
    TestAssertClose stats, "crack.concrete.ignoreStress", concrete.GetStress(0.0001), 0#, 0.000000000001
    TestAssertClose stats, "crack.concrete.ignoreTangent", concrete.GetTangentModulus(0.0001), 0#, 0.000000000001

    Set concrete = TestProvisionalConcreteWithTension()
    TestAssertClose stats, "crack.concrete.useStress", concrete.GetStress(0.00002), 0.65, 0.000000000001
    TestAssertClose stats, "crack.concrete.useLimit", concrete.GetStress(0.001), 1.8, 0.000000000001
    TestAssertClose stats, "crack.concrete.useTangent", concrete.GetTangentModulus(0.00002), 32500#, 0.000000001
End Sub


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет формулу ширины по готовым числам независимо от хранения НДС, статусов и выбора
' арматуры.
Private Sub TestCrackWidthFormulaCalculatorPure(ByRef stats As TCrackTestStats)
    Dim formula As CCrackWidthCalculator
    Set formula = New CCrackWidthCalculator

    TestAssertClose stats, "crack.formula.width", _
        formula.CrackWidthFromData(1.4, 0.5, 1#, 0.8, 200#, 200000#, 320#), _
        0.1792, 0.000000000001
    TestAssertClose stats, "crack.formula.utilization", _
        formula.UtilizationFromData(0.1792, 0.4), 0.448, 0.000000000001
    TestAssertClose stats, "crack.formula.noSteelEs", _
        formula.CrackWidthFromData(1#, 1#, 1#, 1#, 100#, 0#, 100#), _
        0#, 0.000000000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ------------------------------
' User-режим psi_s
' ------------------------------
' Для обычного изгиба проверяем, что a_crc собирается из расчетных sigma_s,
' ls и коэффициентов, а не из ручного CrackSpacing.
Private Sub TestCrackUserPsiMx(ByRef stats As TCrackTestStats)
    Dim stage As String
    On Error GoTo Failed
    stage = "serviceState"
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = TestSolveServiceState(section, -20000#, -15000000#, 0#)

    Dim crack As CCrackWidthCalculator
    stage = "calculateCrack"
    Set crack = TestCalculateCrack(solver, section, -20000#, -15000000#, 0#, "User", "Effective")
    stage = "assertions"
    TestAssertCrackCommon stats, "crack.user.mx", crack
    TestAssertClose stats, "crack.user.psi", crack.PsiS, 1#, 0.000000001
    TestAssertTrue stats, "crack.user.sigmaCrcAvailable", crack.SigmaSCrc > 0#
    Exit Sub
Failed:
    Err.Raise Err.Number, "modTestCrackWidth.TestCrackUserPsiMx." & stage, Err.Source & ": " & Err.Description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Для косого изгиба круга проверяет положительные sigma_s и ls и согласованную
' сборку ширины с заданным пользователем psi_s.
Private Sub TestCrackUserPsiMxy(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = TestSolveCircleServiceState(section, -10000#, -8000000#, -8000000#)

    Dim crack As CCrackWidthCalculator
    Set crack = TestCalculateCrack(solver, section, -10000#, -8000000#, -8000000#, "User", "Effective")
    TestAssertCrackCommon stats, "crack.user.circleMxy", crack
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Передает нестандартные пользовательские phi и psi. Они должны попасть
' в итоговую формулу, а найденные Pre/Post и sigma_s,crc остаться доступными.
Private Sub TestCrackUserCoefficients(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = TestSolveServiceState(section, -20000#, -15000000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = TestCalculateCrack(solver, section, -20000#, -15000000#, 0#, "User", "Effective", _
        allowable:=0.3, phi1Value:=1.6, phi2Value:=0.7, phi3Mode:="User", phi3Value:=1.1, psiSValue:=0.8)
    TestAssertCrackCommon stats, "crack.user.coeffs", crack
    TestAssertClose stats, "crack.user.coeffs.phi1", crack.Phi1, 1.6, 0.000000001
    TestAssertClose stats, "crack.user.coeffs.phi2", crack.Phi2, 0.7, 0.000000001
    TestAssertClose stats, "crack.user.coeffs.phi3", crack.Phi3, 1.1, 0.000000001
    TestAssertClose stats, "crack.user.coeffs.psi", crack.PsiS, 0.8, 0.000000001
    TestAssertTrue stats, "crack.user.coeffs.lambda", crack.FormationResult.LambdaCrc > 0# And crack.FormationResult.LambdaCrc <= 1#
    TestAssertTrue stats, "crack.user.coeffs.preCrackState", Not crack.FormationResult.PreCrackState Is Nothing
    TestAssertTrue stats, "crack.user.coeffs.postCrackState", Not crack.FormationResult.PostCrackState Is Nothing
    TestAssertTrue stats, "crack.user.coeffs.sigmaCrcAvailable", crack.SigmaSCrc > 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ------------------------------
' Зона Abt
' ------------------------------
' Проверяет допустимые короткие aliases режимов расстояния до контура:
' nearest и global должны нормализоваться в разные канонические режимы.
Private Sub TestCoverModeNormalization(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = TestSolveServiceState(section, -20000#, -15000000#, 0#)

    Dim localCrack As CCrackWidthCalculator
    Set localCrack = TestCalculateCrack(solver, section, -20000#, -15000000#, 0#, _
        "User", "Effective", coverMode:="nearest")
    TestAssertCrackCommon stats, "crack.cover.local", localCrack
    TestAssertTrue stats, "crack.cover.local.mode", localCrack.CoverMode = "NearestContour"

    Dim globalCrack As CCrackWidthCalculator
    Set globalCrack = TestCalculateCrack(solver, section, -20000#, -15000000#, 0#, _
        "User", "Effective", coverMode:="global")
    TestAssertCrackCommon stats, "crack.cover.global", globalCrack
    TestAssertTrue stats, "crack.cover.global.mode", globalCrack.CoverMode = "GlobalExtreme"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Для одного косого НДС сравнивает ограниченную полосу Effective и всю
' растянутую зону FullTension; глубина Effective проверяется независимой формулой.
Private Sub TestEffectiveAndFullTensionZones(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = TestSolveServiceState(section, -20000#, -15000000#, -3000000#)

    Dim effectiveCrack As CCrackWidthCalculator
    Set effectiveCrack = TestCalculateCrack(solver, section, -20000#, -15000000#, -3000000#, "User", "Effective")
    Dim fullCrack As CCrackWidthCalculator
    Set fullCrack = TestCalculateCrack(solver, section, -20000#, -15000000#, -3000000#, "User", "FullTension")

    TestAssertCrackCommon stats, "crack.zone.effective", effectiveCrack
    TestAssertCrackCommon stats, "crack.zone.full", fullCrack
    TestAssertTrue stats, "crack.zone.fullAbtPositive", fullCrack.Abt > 0#
    TestAssertTrue stats, "crack.zone.effectiveAbtPositive", effectiveCrack.Abt > 0#
    TestAssertClose stats, "crack.zone.effectiveDepthFormula", effectiveCrack.EffectiveZoneDepth, _
        MinTest(MaxTest(effectiveCrack.TensionDepth, 2# * effectiveCrack.CoverA), 0.5 * effectiveCrack.SectionDepthH), 0.000000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ------------------------------
' Auto-режим psi_s
' ------------------------------
' Если первая проверка Width проходит, Auto сохраняет psi_s = 1.
' Самостоятельный Formation при этом предоставляет свои Pre/Post и sigma_s,crc.
Private Sub TestAutoPsiSkipsLambdaWhenFirstCheckPasses(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = TestSolveServiceState(section, -20000#, -15000000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = TestCalculateCrack(solver, section, -20000#, -15000000#, 0#, "Auto", "Effective", allowable:=1#)
    TestAssertCrackCommon stats, "crack.auto.pass", crack
    TestAssertClose stats, "crack.auto.pass.psi", crack.PsiS, 1#, 0.000000001
    TestAssertTrue stats, "crack.auto.pass.lambda", crack.FormationResult.LambdaCrc > 0# And crack.FormationResult.LambdaCrc <= 1#
    TestAssertTrue stats, "crack.auto.pass.sigmaCrcAvailable", crack.SigmaSCrc > 0#
    TestAssertTrue stats, "crack.auto.pass.preCrackState", Not crack.FormationResult.PreCrackState Is Nothing
    TestAssertTrue stats, "crack.auto.pass.postCrackState", Not crack.FormationResult.PostCrackState Is Nothing
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает Auto и AlwaysCalc при прошедшей первой проверке: только AlwaysCalc
' должен уточнить psi_s по sigma_s,crc и уменьшить ширину на том же НДС.
Private Sub TestAlwaysCalcPsiAppliesSigmaCrcWhenAutoPasses(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = TestSolveServiceState(section, -20000#, -15000000#, 0#)

    Dim autoCrack As CCrackWidthCalculator
    Set autoCrack = TestCalculateCrack(solver, section, -20000#, -15000000#, 0#, "Auto", "Effective", allowable:=1#)

    Dim alwaysCrack As CCrackWidthCalculator
    Set alwaysCrack = TestCalculateCrack(solver, section, -20000#, -15000000#, 0#, "AlwaysCalc", "Effective", allowable:=1#)

    TestAssertCrackCommon stats, "crack.alwaysCalc.pass", alwaysCrack
    TestAssertTrue stats, "crack.alwaysCalc.sigmaCrc", alwaysCrack.SigmaSCrc > 0#
    TestAssertTrue stats, "crack.alwaysCalc.psiReduced", alwaysCrack.PsiS < 1#
    TestAssertTrue stats, "crack.alwaysCalc.widthLessThanAuto", alwaysCrack.CrackWidth < autoCrack.CrackWidth
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Режим усреднения относится только к sigma_s,crc. Проверяем, что
' AllSelected проходит нормализацию и расчет как самостоятельная настройка.
Private Sub TestSigmaSCrcAveragingModeAllSelected(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = TestSolveServiceState(section, -20000#, -15000000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = TestCalculateCrack(solver, section, -20000#, -15000000#, 0#, "AlwaysCalc", "Effective", _
        allowable:=1#, sigmaSCrcAveragingMode:="AllSelected")

    TestAssertCrackCommon stats, "crack.sigmaCrcMode.allSelected", crack
    TestAssertTrue stats, "crack.sigmaCrcMode.allSelected.normalized", _
        StrComp(crack.SigmaSCrcAveragingMode, "AllSelected", vbTextCompare) = 0
    TestAssertTrue stats, "crack.sigmaCrcMode.allSelected.sigmaCrcAvailable", crack.SigmaSCrc > 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' При не прошедшей первой проверке Auto использует найденный PostCrackState
' для psi_s. Проверяются предел psi_s, роли Pre/Post и растягивающий критерий бетона.
Private Sub TestAutoPsiAndLambdaAfterFailedFirstCheck(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = TestSolveServiceState(section, -20000#, -15000000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = TestCalculateCrack(solver, section, -20000#, -15000000#, 0#, "Auto", "Effective", allowable:=0.0001)
    TestAssertCrackCommon stats, "crack.auto.fail", crack
    TestAssertTrue stats, "crack.auto.fail.lambda", crack.FormationResult.LambdaCrc > 0# And crack.FormationResult.LambdaCrc <= 1#
    TestAssertTrue stats, "crack.auto.fail.psiRange", crack.PsiS >= 0# And crack.PsiS <= 1#
    TestAssertTrue stats, "crack.auto.fail.sigmaCrc", crack.SigmaSCrc >= 0#
    TestAssertTrue stats, "crack.auto.fail.psiReduced", crack.PsiS < 1#
    TestAssertTrue stats, "crack.auto.fail.preCrackState", Not crack.FormationResult.PreCrackState Is Nothing
    TestAssertTrue stats, "crack.auto.fail.postCrackState", Not crack.FormationResult.PostCrackState Is Nothing
    TestAssertTrue stats, "crack.auto.fail.preCrackState.type", _
        crack.FormationResult.PreCrackState.StateTypeText = "PreCrackState"
    TestAssertTrue stats, "crack.auto.fail.postCrackState.type", _
        crack.FormationResult.PostCrackState.StateTypeText = "PostCrackState"
    TestAssertTrue stats, "crack.auto.fail.preCrackState.role", _
        crack.FormationResult.PreCrackState.MaterialModelRoleText = "CrackInitiation"
    TestAssertTrue stats, "crack.auto.fail.postCrackState.role", _
        crack.FormationResult.PostCrackState.MaterialModelRoleText = "CrackedState"

    Dim minConcreteStrain As Double
    Dim maxConcreteStrain As Double
    TestConcreteStateStrainBounds crack.FormationResult.PreCrackState, section, minConcreteStrain, maxConcreteStrain
    TestAssertTrue stats, "crack.auto.fail.preCrack.twoSign", minConcreteStrain < 0#
    TestAssertClose stats, "crack.auto.fail.preCrack.epsBtUlt", maxConcreteStrain, 0.00015, 0.000001
    TestAssertTrue stats, "crack.auto.fail.preCrack.notElasticRbtEb", maxConcreteStrain > 1.8 / 32500#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ------------------------------
' Auto-режим Mcrc при чистом изгибе
' ------------------------------
' Проверяет production-сценарий, где исходное CrackedState при N=0 сначала
' находится через CStateSolutionRunner с удобной стартовой плоскостью, а затем
' CCrackFormationCalculator ищет PreCrackState/PostCrackState без отдельного
' пользовательского N и повторной отдельной реализации равновесия.
Private Sub TestAutoMcrcPureBendingConverges(ByRef stats As TCrackTestStats)
    TestCheckAutoMcrcPureBending stats, "crack.auto.pureMx", 0#, -15000000#, 0#
    TestCheckAutoMcrcPureBending stats, "crack.auto.pureMy", 0#, 0#, -15000000#
    TestCheckAutoMcrcPureBending stats, "crack.auto.pureMxy", 0#, -12000000#, -9000000#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Решает заданный чисто изгибный вариант через production-runner и проверяет
' точку Formation, обе стороны трещинообразования и уменьшение psi_s.
Private Sub TestCheckAutoMcrcPureBending(ByRef stats As TCrackTestStats, ByVal prefix As String, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = TestSolveServiceStateWithRunner(section, nValue, mxValue, myValue)

    Dim crack As CCrackWidthCalculator
    Set crack = TestCalculateCrack(solver, section, nValue, mxValue, myValue, _
        "Auto", "Effective", allowable:=0.0001)
    TestAssertCrackCommon stats, prefix, crack
    TestAssertTrue stats, prefix & ".lambda", crack.FormationResult.LambdaCrc > 0# And crack.FormationResult.LambdaCrc <= 1#
    TestAssertTrue stats, prefix & ".preCrackState", Not crack.FormationResult.PreCrackState Is Nothing
    TestAssertTrue stats, prefix & ".postCrackState", Not crack.FormationResult.PostCrackState Is Nothing
    TestAssertTrue stats, prefix & ".psiReduced", crack.PsiS < 1#

    Dim minConcreteStrain As Double
    Dim maxConcreteStrain As Double
    TestConcreteStateStrainBounds crack.FormationResult.PreCrackState, section, minConcreteStrain, maxConcreteStrain
    TestAssertTrue stats, prefix & ".preCrack.twoSign", minConcreteStrain < 0#
    TestAssertClose stats, prefix & ".preCrack.epsBtUlt", maxConcreteStrain, 0.00015, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет смысл Mcrc для общей N+M ветки: при одной и той же продольной
' силе и направлении изгиба lambda масштабирует только моментный вектор.
' Поэтому Mcrc не должен зависеть от того, насколько далеко текущий LC
' находится за порогом трещинообразования.
Private Sub TestAutoMcrcFixedNIndependentOfMomentMagnitude(ByRef stats As TCrackTestStats)
    Dim sectionLow As CSectionModel
    Dim sectionHigh As CSectionModel
    Dim solverLow As CSectionSolver
    Dim solverHigh As CSectionSolver
    Dim nValue As Double
    Dim mxLow As Double
    Dim mxHigh As Double
    nValue = -20000#
    mxLow = -12000000#
    mxHigh = -18000000#

    Set solverLow = TestSolveServiceStateWithRunner(sectionLow, nValue, mxLow, 0#)
    Set solverHigh = TestSolveServiceStateWithRunner(sectionHigh, nValue, mxHigh, 0#)

    Dim crackLow As CCrackWidthCalculator
    Dim crackHigh As CCrackWidthCalculator
    Set crackLow = TestCalculateCrack(solverLow, sectionLow, nValue, mxLow, 0#, "Auto", "Effective", allowable:=0.0001)
    Set crackHigh = TestCalculateCrack(solverHigh, sectionHigh, nValue, mxHigh, 0#, "Auto", "Effective", allowable:=0.0001)

    TestAssertCrackCommon stats, "crack.auto.mcrcFixedN.low", crackLow
    TestAssertCrackCommon stats, "crack.auto.mcrcFixedN.high", crackHigh
    TestAssertTrue stats, "crack.auto.mcrcFixedN.lowState", Not crackLow.FormationResult.PreCrackState Is Nothing
    TestAssertTrue stats, "crack.auto.mcrcFixedN.highState", Not crackHigh.FormationResult.PreCrackState Is Nothing
    TestAssertClose stats, "crack.auto.mcrcFixedN.mcrcInvariant", _
        Abs(crackLow.FormationResult.Mcrc), Abs(crackHigh.FormationResult.Mcrc), 25000#
    TestAssertClose stats, "crack.auto.mcrcFixedN.lambdaLowMoment", _
        Abs(crackLow.FormationResult.LambdaCrc * mxLow), Abs(crackLow.FormationResult.Mcrc), 25000#
    TestAssertClose stats, "crack.auto.mcrcFixedN.lambdaHighMoment", _
        Abs(crackHigh.FormationResult.LambdaCrc * mxHigh), Abs(crackHigh.FormationResult.Mcrc), 25000#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет три пользовательских пути поиска образования нормальной трещины.
' lambda*Mxy сохраняет N постоянной, lambda*N нужен для центрального
' растяжения, а lambda*NMxy масштабирует весь вектор N/Mx/My как единую траекторию.
Private Sub TestCrackInitiationLoadPaths(ByRef stats As TCrackTestStats)
    Dim sectionMxy As CSectionModel
    Dim solverMxy As CSectionSolver
    Set solverMxy = TestSolveServiceStateWithRunner(sectionMxy, -20000#, -15000000#, 0#)

    Dim crackMxy As CCrackWidthCalculator
    Set crackMxy = TestCalculateCrack(solverMxy, sectionMxy, -20000#, -15000000#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="lambda*Mxy")
    TestAssertCrackCommon stats, "crack.path.mxy", crackMxy
    TestAssertTrue stats, "crack.path.mxy.lambda", crackMxy.FormationResult.LambdaCrc > 0# And crackMxy.FormationResult.LambdaCrc <= 1#
    TestAssertClose stats, "crack.path.mxy.nFixed", crackMxy.FormationResult.FormationNcrc, -20000#, 0.001

    Dim sectionN As CSectionModel
    Dim solverN As CSectionSolver
    Set solverN = TestSolveServiceState(sectionN, 200000#, 0#, 0#)

    Dim crackN As CCrackWidthCalculator
    Set crackN = TestCalculateCrack(solverN, sectionN, 200000#, 0#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="lambda*N")
    TestAssertCrackCommon stats, "crack.path.n", crackN
    TestAssertTrue stats, "crack.path.n.central", crackN.CentralTensionBranch
    TestAssertClose stats, "crack.path.n.formationN", crackN.FormationResult.FormationNcrc, crackN.FormationResult.Ncrc, 0.001

    Dim crackAutoN As CCrackWidthCalculator
    Set crackAutoN = TestCalculateCrack(solverN, sectionN, 200000#, 0#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="Auto")
    TestAssertCrackCommon stats, "crack.path.autoN", crackAutoN
    TestAssertTrue stats, "crack.path.autoN.central", crackAutoN.CentralTensionBranch
    TestAssertTrue stats, "crack.path.autoN.method", crackAutoN.FormationResult.FormationMethod = ChrW$(&H3BB) & "*N"
    TestAssertClose stats, "crack.path.autoN.formationN", crackAutoN.FormationResult.FormationNcrc, crackAutoN.FormationResult.Ncrc, 0.001

    Dim sectionNMxy As CSectionModel
    Dim solverNMxy As CSectionSolver
    Set solverNMxy = TestSolveServiceStateWithRunner(sectionNMxy, -20000#, -15000000#, 0#)

    Dim crackNMxy As CCrackWidthCalculator
    Set crackNMxy = TestCalculateCrack(solverNMxy, sectionNMxy, -20000#, -15000000#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="lambda*NMxy")
    TestAssertCrackCommon stats, "crack.path.nmxy", crackNMxy
    TestAssertTrue stats, "crack.path.nmxy.lambda", crackNMxy.FormationResult.LambdaCrc > 0# And crackNMxy.FormationResult.LambdaCrc <= 1#
    TestAssertClose stats, "crack.path.nmxy.nScaled", crackNMxy.FormationResult.FormationNcrc, crackNMxy.FormationResult.LambdaCrc * -20000#, 0.001
    TestAssertClose stats, "crack.path.nmxy.mScaled", crackNMxy.FormationResult.Mcrc, Abs(crackNMxy.FormationResult.LambdaCrc * -15000000#), 50000#

    Dim crackNFixedMomentFallback As CCrackWidthCalculator
    Set crackNFixedMomentFallback = TestCalculateCrack(solverMxy, sectionMxy, -20000#, -15000000#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="lambda*N")
    TestAssertTrue stats, "crack.path.nFixedMomentFallback.converged", crackNFixedMomentFallback.Converged
    TestAssertTrue stats, "crack.path.nFixedMomentFallback.formed", crackNFixedMomentFallback.CrackFormed
    TestAssertTrue stats, "crack.path.nFixedMomentFallback.baseFailure", _
        crackNFixedMomentFallback.FormationResult.ResultMeta.InternalStatus = rsCheckFailed And _
        crackNFixedMomentFallback.FormationResult.ResultMeta.ResultCode = rcInitialStateBeyondLimit
    TestAssertClose stats, "crack.path.nFixedMomentFallback.lambda0", crackNFixedMomentFallback.FormationResult.LambdaCrc, 0#, 0.000000001
    TestAssertClose stats, "crack.path.nFixedMomentFallback.noFormationN", crackNFixedMomentFallback.FormationResult.FormationNcrc, 0#, 0.000000001
    TestAssertClose stats, "crack.path.nFixedMomentFallback.noMcrc", crackNFixedMomentFallback.FormationResult.Mcrc, 0#, 0.000000001
    TestAssertTrue stats, "crack.path.nFixedMomentFallback.noBeforeState", crackNFixedMomentFallback.FormationResult.PreCrackState Is Nothing
    TestAssertTrue stats, "crack.path.nFixedMomentFallback.noAfterState", crackNFixedMomentFallback.FormationResult.PostCrackState Is Nothing
    TestAssertClose stats, "crack.path.nFixedMomentFallback.psi1", crackNFixedMomentFallback.PsiS, 1#, 0.000000001
    TestAssertCrackCalculatorNotNumFail stats, "crack.path.nFixedMomentFallback.status", crackNFixedMomentFallback

    Dim crackFallback As CCrackWidthCalculator
    Set crackFallback = TestCalculateCrack(solverN, sectionN, 200000#, 0#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="lambda*Mxy")
    TestAssertTrue stats, "crack.path.mxyAxialFallback.converged", crackFallback.Converged
    TestAssertTrue stats, "crack.path.mxyAxialFallback.formed", crackFallback.CrackFormed
    TestAssertTrue stats, "crack.path.mxyAxialFallback.emptyScaledLoad", _
        crackFallback.FormationResult.ResultMeta.InternalStatus = rsInvalidInput And _
        crackFallback.FormationResult.ResultMeta.ResultCode = rcInvalidInput
    TestAssertClose stats, "crack.path.mxyAxialFallback.lambda0", crackFallback.FormationResult.LambdaCrc, 0#, 0.000000001
    TestAssertClose stats, "crack.path.mxyAxialFallback.psi1", crackFallback.PsiS, 1#, 0.000000001
    TestAssertCrackCalculatorNotNumFail stats, "crack.path.mxyAxialFallback.status", crackFallback
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' настоящая формула получает физическое текущее НДС и отсутствие
' точки Formation с разными typed причинами. Ширина сохраняет psi_s = 1 и
' warning, а NumericalFailure остается только у поиска образования.
Private Sub TestFormationFailureStillCalculatesWidth(ByRef stats As TCrackTestStats)
    Dim section As CSectionModel, solver As CSectionSolver, current As CSectionStateResult
    Set solver = TestSolveServiceState(section, 200000#, 0#, 0#)
    Set current = New CSectionStateResult
    current.InitializeFromSolver sstCrackedState, cpCrackedNDS, TestCrackedStateSpec(), solver, False, True
    Dim owner As CCrackFormationCalculator, formation As CCrackFormationResult, meta As CResultMeta
    Dim width As CCrackWidthCalculator, load As CSectionLoadState, mode As Variant, failure As Variant, prefix As String
    Set load = New CSectionLoadState: load.Initialize 200000#, 0#, 0#, 0#, 0#
    For Each failure In Array(sfcNumericalFailure, sfcInvalidInput, sfcInvalidConfiguration, sfcInternalError)
        Set owner = New CCrackFormationCalculator
        owner.LimitSearchSetFailure CLng(failure), "Контрольная причина отсутствия точки образования трещины."
        Set formation = New CCrackFormationResult
        Set meta = New CResultMeta
        meta.SetSolverFailure CLng(failure), rkCrackFormation, "Контрольная причина отсутствия точки образования трещины.", True
        formation.InitializeFromCalculator owner, meta
        For Each mode In Array("User", "Auto", "AlwaysCalc")
            prefix = "crack.formationFailureFallback." & CStr(failure) & "." & CStr(mode)
            Set width = New CCrackWidthCalculator
            width.PsiMode = CStr(mode): width.PsiS = 0.25: width.AllowableCrackWidth = 10#
            width.Calculate current, section, TestMaterialProvider(), TestCrackedStateSpec(), formation, load
            TestAssertTrue stats, prefix & ".calculated", width.ResultMeta.Calculated
            TestAssertTrue stats, prefix & ".warning", width.ResultMeta.InternalStatus = rsSuccessWithWarning
            TestAssertClose stats, prefix & ".psi1", width.PsiS, 1#, 0#
            TestAssertTrue stats, prefix & ".comment", InStr(1, width.ResultMeta.ResultComment, "Предупреждение", vbTextCompare) > 0
            TestAssertCrackCommon stats, prefix, width
            TestAssertTrue stats, prefix & ".formationCausePreserved", formation.ResultMeta.InternalStatus = meta.InternalStatus
            width.AllowableCrackWidth = 0.0000001
            width.Calculate current, section, TestMaterialProvider(), TestCrackedStateSpec(), formation, load
            TestAssertTrue stats, prefix & ".exceededCalculated", width.ResultMeta.Calculated
            TestAssertTrue stats, prefix & ".exceededStatus", width.ResultMeta.InternalStatus = rsCheckFailed
            TestAssertTrue stats, prefix & ".exceededCode", width.ResultMeta.ResultCode = rcCrackWidthExceeded
            TestAssertTrue stats, prefix & ".exceededWarning", _
                InStr(1, width.ResultMeta.ResultComment, "Предупреждение", vbTextCompare) > 0
            TestAssertClose stats, prefix & ".exceededPsi1", width.PsiS, 1#, 0#
        Next mode
    Next failure
    Set width = New CCrackWidthCalculator
    width.Calculate Nothing, section, TestMaterialProvider(), TestCrackedStateSpec(), formation, load
    TestAssertTrue stats, "crack.formationFailureFallback.currentMissingBlocks", width.CrackWidthInternalStatus = rsBlockedByDependency
    current.InitializeFromSolver sstCrackedState, cpCrackedNDS, TestCrackedStateSpec(), solver, False, False
    width.Calculate current, section, TestMaterialProvider(), TestCrackedStateSpec(), formation, load
    TestAssertTrue stats, "crack.formationFailureFallback.currentPhysicalFailureBlocks", _
        width.CrackWidthInternalStatus = rsBlockedByDependency And Not width.ResultMeta.Calculated
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отсутствие Formation не превращает нулевую кривизну в ошибку
' деления. Реальные нулевые/малые и сжатые состояния дают неприменимость Width,
' а равномерное растяжение использует всю растянутую арматуру и psi_s = 1.
Private Sub TestFormationFallbackWithUniformCurrentState(ByRef stats As TCrackTestStats)
    Dim section As CSectionModel, solver As CSectionSolver, current As CSectionStateResult
    Dim owner As CCrackFormationCalculator, formation As CCrackFormationResult, meta As CResultMeta
    Dim width As CCrackWidthCalculator, load As CSectionLoadState, item As Variant, mode As Variant, prefix As String
    Set owner = New CCrackFormationCalculator
    owner.LimitSearchSetFailure sfcNumericalFailure, "Контрольная несходимость поиска точки образования трещины."
    Set meta = New CResultMeta
    meta.SetSolverFailure sfcNumericalFailure, rkCrackFormation, "Контрольная несходимость поиска точки образования трещины.", True
    Set formation = New CCrackFormationResult: formation.InitializeFromCalculator owner, meta
    For Each item In Array(Array(0#, 0#, 0#), Array(0.000000001, 0#, 0#), _
            Array(-0.000000001, 0#, 0#), Array(0#, 0.000000001, 0#), Array(0#, -0.000000001, 0#), _
            Array(0#, 0#, 0.000000001), Array(0#, 0#, -0.000000001), Array(-200000#, 0#, 0#), Array(200000#, 0#, 0#))
        Set solver = TestSolveServiceState(section, CDbl(item(0)), CDbl(item(1)), CDbl(item(2)))
        Set current = New CSectionStateResult
        current.InitializeFromSolver sstCrackedState, cpCrackedNDS, TestCrackedStateSpec(), solver, False, True
        Set load = New CSectionLoadState: load.Initialize CDbl(item(0)), CDbl(item(1)), CDbl(item(2)), 0#, 0#
        For Each mode In Array("User", "Auto", "AlwaysCalc")
            prefix = "crack.uniformFallback." & CStr(item(0)) & "." & CStr(item(1)) & "." & CStr(item(2)) & "." & CStr(mode)
            Set width = New CCrackWidthCalculator
            width.PsiMode = CStr(mode): width.PsiS = 0.25: width.AllowableCrackWidth = 10#
            width.Calculate current, section, TestMaterialProvider(), TestCrackedStateSpec(), formation, load
            TestAssertTrue stats, prefix & ".converged", width.Converged
            TestAssertTrue stats, prefix & ".causePreserved", formation.ResultMeta.InternalStatus = rsNumericalFailure
            If CDbl(item(0)) = 200000# Then
                TestAssertTrue stats, prefix & ".calculated", width.ResultMeta.Calculated
                TestAssertTrue stats, prefix & ".warning", width.ResultMeta.InternalStatus = rsSuccessWithWarning
                TestAssertClose stats, prefix & ".psi1", width.PsiS, 1#, 0#
                TestAssertClose stats, prefix & ".allConcrete", width.Abt, 60000#, 0.000000001
                TestAssertCrackCommon stats, prefix, width
            Else
                TestAssertTrue stats, prefix & ".notApplicable", width.ResultMeta.InternalStatus = rsNotApplicable
                TestAssertTrue stats, prefix & ".notCalculated", Not width.ResultMeta.Calculated
                TestAssertClose stats, prefix & ".zeroWidth", width.CrackWidth, 0#, 0#
                TestAssertTrue stats, prefix & ".reason", Len(width.ResultMeta.ResultComment) > 0
            End If
        Next mode
    Next item
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Малый момент не достигает критерия до технического MaxLambda. Результат
' обязан сохранить SEARCH_BOUND_REACHED без фиктивного Post и NumFail у Width.
Private Sub TestCrackFormationSearchBoundKeepsTechnicalCode(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = TestSolveServiceStateWithRunner(section, -20000#, -6000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = TestCalculateCrack(solver, section, -20000#, -6000#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="lambda*Mxy", _
        formationStrategy:="LoadMultiplier")

    TestAssertTrue stats, "crack.searchBound.formationNotConverged", Not crack.FormationResult.Converged
    TestAssertTrue stats, "crack.searchBound.notFormed", Not crack.CrackFormed
    TestAssertTrue stats, "crack.searchBound.internalStatus", _
        crack.FormationResult.ResultMeta.InternalStatus = rsCheckFailed
    TestAssertTrue stats, "crack.searchBound.resultCode", _
        crack.FormationResult.ResultMeta.ResultCode = rcSearchBoundReached
    TestAssertTrue stats, "crack.searchBound.readableInteger", _
        InStr(1, crack.FormationResult.ResultMeta.ResultComment, ".,", vbBinaryCompare) = 0
    TestAssertTrue stats, "crack.searchBound.readableCriterion", _
        InStr(1, crack.FormationResult.ResultMeta.ResultComment, "условие образования нормальной трещины", vbTextCompare) > 0
    TestAssertTrue stats, "crack.searchBound.noPostState", crack.FormationResult.PostCrackState Is Nothing
    TestAssertCrackCalculatorNotNumFail stats, "crack.searchBound.widthNotNumFail", crack
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Фиксированный малый момент и масштабирование сжимающей N не образуют трещину.
' Проверяются доказанное недостижение критерия, неприменимость Width и отсутствие Pre/Post.
Private Sub TestCrackFormationNoCrackDoesNotBuildPostState(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = TestSolveServiceStateWithRunner(section, -20000#, -10000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = TestCalculateCrack(solver, section, -20000#, -10000#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="lambda*N")

    TestAssertTrue stats, "crack.noCrack.converged", crack.Converged
    TestAssertTrue stats, "crack.noCrack.notFormed", Not crack.CrackFormed
    TestAssertTrue stats, "crack.noCrack.criterion", _
        crack.FormationResult.ResultMeta.ResultCode = rcCriterionNotReached
    TestAssertTrue stats, "crack.noCrack.fixedPathSuccess", crack.FormationResult.ResultMeta.InternalStatus = rsSuccess
    TestAssertTrue stats, "crack.noCrack.confirmed", crack.FormationResult.ConfirmedNotCracked
    TestAssertTrue stats, "crack.noCrack.widthNotApplicable", crack.CrackWidthInternalStatus = rsNotApplicable
    TestAssertTrue stats, "crack.noCrack.noBeforeState", crack.FormationResult.PreCrackState Is Nothing
    TestAssertTrue stats, "crack.noCrack.noPostState", crack.FormationResult.PostCrackState Is Nothing
    TestAssertCrackCalculatorNotNumFail stats, "crack.noCrack.notNumFail", crack
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторный Formation использует repository без нового LastRunner. Width
' все равно получает sigma_s,crc и Pre/Post из сохраненных результатов.
Private Sub TestCrackFormationCacheHitWithoutLastRunner(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = TestSolveServiceStateWithRunner(section, -20000#, -15000000#, 0#)

    Dim repository As CStateRepository
    Set repository = New CStateRepository

    Dim provider As CStateProvider
    Set provider = New CStateProvider
    provider.Initialize section, TestMaterialProvider(), repository
    ConfigureTestStateProvider provider

    Dim firstCrack As CCrackWidthCalculator
    Set firstCrack = TestCalculateCrack(solver, section, -20000#, -15000000#, 0#, _
        "AlwaysCalc", "Effective", allowable:=1#, stateProvider:=provider)
    TestAssertCrackCommon stats, "crack.cache.first", firstCrack
    TestAssertTrue stats, "crack.cache.first.sigmaCrc", firstCrack.SigmaSCrc > 0#
    TestAssertTrue stats, "crack.cache.first.preState", Not firstCrack.FormationResult.PreCrackState Is Nothing
    TestAssertTrue stats, "crack.cache.first.postState", Not firstCrack.FormationResult.PostCrackState Is Nothing

    Dim secondCrack As CCrackWidthCalculator
    Set secondCrack = TestCalculateCrack(solver, section, -20000#, -15000000#, 0#, _
        "AlwaysCalc", "Effective", allowable:=1#, stateProvider:=provider)
    TestAssertCrackCommon stats, "crack.cache.second", secondCrack
    TestAssertTrue stats, "crack.cache.second.reused", provider.LastStateWasReused
    TestAssertTrue stats, "crack.cache.second.noLastRunner", provider.LastRunner Is Nothing
    TestAssertTrue stats, "crack.cache.second.sigmaCrc", secondCrack.SigmaSCrc > 0#
    TestAssertTrue stats, "crack.cache.second.preState", Not secondCrack.FormationResult.PreCrackState Is Nothing
    TestAssertTrue stats, "crack.cache.second.postState", Not secondCrack.FormationResult.PostCrackState Is Nothing
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет независимость сохранённой диагностики результата поиска.
' Последующие попытки образования трещины не должны менять уже выданный отчёт.
Private Sub TestLimitSearchResultKeepsCrackDiagnosticSnapshot(ByRef stats As TCrackTestStats)
    Dim crack As CCrackFormationCalculator
    Set crack = New CCrackFormationCalculator
    crack.LimitSearchAppendDiagnostic "diagnostic-before"

    Dim result As CLimitSearchResult
    Set result = crack.BuildSearchSnapshot("lambda*Mxy")

    crack.LimitSearchAppendDiagnostic "diagnostic-after"

    TestAssertTrue stats, "limitSearch.crackSnapshot.hasBefore", _
        InStr(1, result.DiagnosticLog, "diagnostic-before", vbTextCompare) > 0
    TestAssertTrue stats, "limitSearch.crackSnapshot.noAfter", _
        InStr(1, result.DiagnosticLog, "diagnostic-after", vbTextCompare) = 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет стартовую плоскость чистого и косого изгиба и обоих осевых знаков.
' Эти конечные нагрузки должны дать Formation либо обоснованное отсутствие трещины.
Private Sub TestDangerousLoadsDoNotNumFail(ByRef stats As TCrackTestStats)
    TestCheckDangerousCrackLoad stats, "crack.danger.pureMx", 0#, -15000000#, 0#, True
    TestCheckDangerousCrackLoad stats, "crack.danger.pureMy", 0#, 0#, -15000000#, True
    TestCheckDangerousCrackLoad stats, "crack.danger.pureMxy", 0#, -12000000#, -9000000#, True
    TestCheckDangerousCrackLoad stats, "crack.danger.centralTension", 200000#, 0#, 0#, True, "lambda*N"
    TestCheckDangerousCrackLoad stats, "crack.danger.pureCompression", -100000#, 0#, 0#, False
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Рассчитывает один вариант обычным способом и проверяет код исхода образования
' трещины. Текст пояснения не используется для определения численной ошибки.
Private Sub TestCheckDangerousCrackLoad(ByRef stats As TCrackTestStats, ByVal prefix As String, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, _
        ByVal shouldForm As Boolean, Optional ByVal formationPath As String = "lambda*Mxy")
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = TestSolveServiceStateWithRunner(section, nValue, mxValue, myValue)

    Dim crack As CCrackWidthCalculator
    Set crack = TestCalculateCrack(solver, section, nValue, mxValue, myValue, _
        "Auto", "Effective", allowable:=0.0001, formationPath:=formationPath)
    TestAssertTrue stats, prefix & ".converged", crack.Converged
    TestAssertTrue stats, prefix & ".formationNotNumFail", _
        crack.FormationResult.ResultMeta.InternalStatus <> rsNumericalFailure
    TestAssertCrackCalculatorNotNumFail stats, prefix & ".notNumFail", crack
    If shouldForm Then
        TestAssertTrue stats, prefix & ".formed", crack.CrackFormed
        TestAssertTrue stats, prefix & ".lambda", crack.FormationResult.LambdaCrc > 0# And crack.FormationResult.LambdaCrc <= 1#
        TestAssertTrue stats, prefix & ".preCrackState", Not crack.FormationResult.PreCrackState Is Nothing
        TestAssertTrue stats, prefix & ".postCrackState", Not crack.FormationResult.PostCrackState Is Nothing
    Else
        TestAssertTrue stats, prefix & ".notFormed", Not crack.CrackFormed
        TestAssertClose stats, prefix & ".width", crack.CrackWidth, 0#, 0.000000000001
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Для целиком растянутого бетона проверяет зависимость предельной деформации
' Formation от отношения минимальной и максимальной деформаций по формуле 8.54.
Private Sub TestAutoMcrcOneSignTensionUsesFormula854(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = TestSolveServiceState(section, 100000#, 3000000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = TestCalculateCrack(solver, section, 100000#, 3000000#, 0#, _
        "Auto", "Effective", allowable:=0.0001)
    TestAssertCrackCommon stats, "crack.auto.oneSign", crack
    TestAssertTrue stats, "crack.auto.oneSign.preCrackState", Not crack.FormationResult.PreCrackState Is Nothing

    Dim minConcreteStrain As Double
    Dim maxConcreteStrain As Double
    TestConcreteStateStrainBounds crack.FormationResult.PreCrackState, section, minConcreteStrain, maxConcreteStrain
    TestAssertTrue stats, "crack.auto.oneSign.allTension", minConcreteStrain > 0#

    Dim expectedUlt As Double
    expectedUlt = 0.00015 - (0.00015 - 0.0001) * (minConcreteStrain / maxConcreteStrain)
    TestAssertClose stats, "crack.auto.oneSign.formula854", maxConcreteStrain, expectedUlt, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ------------------------------
' Центральное растяжение
' ------------------------------
' При центральном растяжении по внешним нагрузкам применяется отдельная ветка Ncrc = Ared*Rbt,ser.
Private Sub TestCentralTensionBranch(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = TestSolveServiceState(section, 200000#, 0#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = TestCalculateCrack(solver, section, 200000#, 0#, 0#, "Auto", "Effective", _
        allowable:=0.0001, formationPath:="lambda*N")
    Dim fullZoneCrack As CCrackWidthCalculator
    Set fullZoneCrack = TestCalculateCrack(solver, section, 200000#, 0#, 0#, "Auto", "FullTension", _
        allowable:=0.0001, formationPath:="lambda*N")
    TestAssertCrackCommon stats, "crack.central", crack
    TestAssertTrue stats, "crack.central.branch", crack.CentralTensionBranch
    TestAssertTrue stats, "crack.central.ncrc", crack.FormationResult.Ncrc > 0#
    TestAssertTrue stats, "crack.central.lambda", crack.FormationResult.LambdaCrc > 0# And crack.FormationResult.LambdaCrc <= 1#
    TestAssertClose stats, "crack.central.lambdaFromNcrc", crack.FormationResult.LambdaCrc * 200000#, crack.FormationResult.Ncrc, 0.001
    TestAssertTrue stats, "crack.central.preCrackState", Not crack.FormationResult.PreCrackState Is Nothing
    TestAssertTrue stats, "crack.central.postCrackState", Not crack.FormationResult.PostCrackState Is Nothing
    TestAssertClose stats, "crack.central.phi3", crack.Phi3, 1.2, 0.000000001
    TestAssertClose stats, "crack.central.zoneModeInvariant", fullZoneCrack.CrackWidth, crack.CrackWidth, 0.000000001

    Dim shiftedSolver As CSectionSolver
    Set shiftedSolver = TestSolveServiceState(section, 200000#, 2500000#, -1500000#)
    Dim shiftedMomentCrack As CCrackWidthCalculator
    Set shiftedMomentCrack = TestCalculateCrack(shiftedSolver, section, 200000#, 2500000#, -1500000#, _
        "User", "Effective")
    TestAssertTrue stats, "crack.central.eccentricLoadRejected", Not shiftedMomentCrack.CentralTensionBranch

    Dim offsetLoad As CSectionLoadState
    Set offsetLoad = New CSectionLoadState
    offsetLoad.Initialize 200000#, 0#, 0#, -7.5, 12.5
    Dim offsetCrack As CCrackWidthCalculator
    Set offsetCrack = TestCalculateCrack(shiftedSolver, section, offsetLoad.N, offsetLoad.InternalMx, offsetLoad.InternalMy, _
        "User", "Effective", _
        loadStateForClassification:=offsetLoad, centralReferenceX:=0#, centralReferenceY:=0#)
    TestAssertTrue stats, "crack.central.offsetLoadStateRejected", Not offsetCrack.CentralTensionBranch
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ------------------------------
' Нет растянутой арматуры
' ------------------------------
' Чистое сжатие должно завершаться корректно, но без раскрытой трещины.
Private Sub TestNoTensionRebar(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = TestSolveServiceState(section, -100000#, 0#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = TestCalculateCrack(solver, section, -100000#, 0#, 0#, "User", "Effective")
    TestAssertTrue stats, "crack.noTension.converged", crack.Converged
    TestAssertTrue stats, "crack.noTension.notFormed", Not crack.CrackFormed
    TestAssertClose stats, "crack.noTension.width", crack.CrackWidth, 0#, 0.000000000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт прямоугольное сечение для II ГПС и рассчитывает текущее НДС
' без растянутого бетона. Если равновесие не подтверждено, останавливается
' подготовка теста; формула ширины по непригодному НДС не проверяется.
Private Function TestSolveServiceState(ByRef section As CSectionModel, ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double) As CSectionSolver
    Dim stage As String
    On Error GoTo Failed
    stage = "shape"
    Dim geom As CGeometryRegion
    Set geom = TestFixtureBox(300#, 200#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    stage = "mesh"
    mesh.BuildMesh geom, 20#, 20#, 1

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    stage = "rebars"
    rebars.AddBar "B1", -90#, -60#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B2", 90#, -60#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B3", -90#, 60#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B4", 90#, 60#, 20#, 0#, "Rebar", "", geom

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    ConfigureTestSolver solver
    stage = "section"
    Set section = TestBuildGeneratedSectionModel(mesh, rebars)
    Dim provider As CMaterialModelProvider
    Set provider = TestMaterialProvider()
    stage = "solve"
    solver.Solve section, provider.ConcreteMaterial(cpCrackedNDS), provider.SteelMaterial(cpCrackedNDS), nValue, mxValue, myValue
    If Not solver.Converged Then Err.Raise vbObjectError + 3800, "modTestCrackWidth", "Service state did not converge: " & solver.StopReason
    Set TestSolveServiceState = solver
    Exit Function
Failed:
    Err.Raise Err.Number, "modTestCrackWidth.TestSolveServiceState." & stage, Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт круг с четырьмя стержнями и рассчитывает текущее НДС для II ГПС.
' Возвращает согласованные модель и решатель; несходимость явно отмечается как ошибка
' подготовки.
Private Function TestSolveCircleServiceState(ByRef section As CSectionModel, ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double) As CSectionSolver
    Dim geom As CGeometryRegion
    Set geom = TestFixtureDisk(300#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 15#, 15#, 1

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", 0#, 90#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B2", 90#, 0#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B3", 0#, -90#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B4", -90#, 0#, 20#, 0#, "Rebar", "", geom

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    ConfigureTestSolver solver
    Set section = TestBuildGeneratedSectionModel(mesh, rebars)
    Dim provider As CMaterialModelProvider
    Set provider = TestMaterialProvider()
    solver.Solve section, provider.ConcreteMaterial(cpCrackedNDS), provider.SteelMaterial(cpCrackedNDS), nValue, mxValue, myValue
    If Not solver.Converged Then Err.Raise vbObjectError + 3801, "modTestCrackWidth", "Circle service state did not converge: " & solver.StopReason
    Set TestSolveCircleServiceState = solver
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Решает текущее CrackedState тем же production-runner-ом, который batch
' использует перед расчетом трещин. Это важно для чистого изгиба: стартовая
' плоскость подбирается централизованно в CStateGuessBuilder.
Private Function TestSolveServiceStateWithRunner(ByRef section As CSectionModel, ByVal nValue As Double, _
        ByVal mxValue As Double, ByVal myValue As Double) As CSectionSolver
    Dim geom As CGeometryRegion
    Set geom = TestFixtureBox(300#, 200#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 20#, 20#, 1

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -90#, -60#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B2", 90#, -60#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B3", -90#, 60#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B4", 90#, 60#, 20#, 0#, "Rebar", "", geom

    Set section = TestBuildGeneratedSectionModel(mesh, rebars)

    Dim provider As CMaterialModelProvider
    Set provider = TestMaterialProvider()
    Dim crackedSpec As CMaterialModelSpec
    Set crackedSpec = TestCrackedStateSpec()

    Dim runner As CStateSolutionRunner
    Set runner = New CStateSolutionRunner
    ConfigureTestStateRunner runner
    runner.Solve section, provider.ConcreteMaterialForEquilibriumFromSpec(crackedSpec), _
        provider.SteelMaterialForEquilibriumFromSpec(crackedSpec), nValue, mxValue, myValue, False

    If runner.ResultSolver Is Nothing Or Not runner.Converged Then
        Err.Raise vbObjectError + 3802, "modTestCrackWidth", _
            "Runner service state did not converge: " & runner.StopReason
    End If
    Set TestSolveServiceStateWithRunner = runner.ResultSolver
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Однотипный setup равновесия для fixtures: восемь шагов, 80 итераций,
' компонентные допуски 5 Н и 5000 Н*мм. Units и знаки уже внутренние.
Private Sub ConfigureTestSolver(ByVal solver As CSectionSolver)
    solver.LoadSteps = 8
    solver.MaxIterations = 80
    solver.ToleranceN = 5#
    solver.ToleranceMx = 5000#
    solver.ToleranceMy = 5000#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Передает общему runner тот же фиксированный бюджет и компонентные допуски fixture.
' Это позволяет сравнивать prepared-State маршрут без изменения расчетной постановки.
Private Sub ConfigureTestStateRunner(ByVal runner As CStateSolutionRunner)
    runner.LoadSteps = 8
    runner.MaxIterations = 80
    runner.ToleranceN = 5#
    runner.ToleranceMx = 5000#
    runner.ToleranceMy = 5000#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Настраивает provider для обычного решения и reuse состояний в crack-тестах.
' Фиксированные допуски во внутренних единицах не зависят от Config.
Private Sub ConfigureTestStateProvider(ByVal provider As CStateProvider)
    provider.LoadSteps = 8
    provider.MaxIterations = 80
    provider.ToleranceN = 5#
    provider.ToleranceMx = 5000#
    provider.ToleranceMy = 5000#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет Formation отдельно от Width и передает последнему готовые состояния.
' Optional provider позволяет проверить reuse; параметры режимов задаются явно,
' чтобы тест не зависел от текущих пользовательских defaults книги.
Private Function TestCalculateCrack(ByVal solver As CSectionSolver, ByVal section As CSectionModel, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, _
        ByVal psiMode As String, ByVal zoneMode As String, _
        Optional ByVal allowable As Double = 0.3, _
        Optional ByVal phi1Value As Double = 1.4, Optional ByVal phi2Value As Double = 0.5, _
        Optional ByVal phi3Mode As String = "Auto", Optional ByVal phi3Value As Double = 1#, _
        Optional ByVal psiSValue As Double = 1#, Optional ByVal coverMode As String = "NearestContour", _
        Optional ByVal loadStateForClassification As CSectionLoadState = Nothing, _
        Optional ByVal centralReferenceX As Double = 0#, Optional ByVal centralReferenceY As Double = 0#, _
        Optional ByVal formationPath As String = "lambda*Mxy", _
        Optional ByVal sigmaSCrcAveragingMode As String = "TensionOnly", _
        Optional ByVal formationStrategy As String = "Auto", _
        Optional ByVal stateProvider As CStateProvider = Nothing) As CCrackWidthCalculator
    Dim stage As String
    On Error GoTo Failed
    stage = "createCalculators"
    Dim crack As CCrackWidthCalculator
    Set crack = New CCrackWidthCalculator
    Dim formationCalculator As CCrackFormationCalculator
    Set formationCalculator = New CCrackFormationCalculator
    stage = "configure"
    crack.AllowableCrackWidth = allowable
    crack.PsiMode = psiMode
    crack.SigmaSCrcAveragingMode = sigmaSCrcAveragingMode
    crack.TensionZoneMode = zoneMode
    crack.Phi1 = phi1Value
    crack.Phi2 = phi2Value
    crack.Phi3Mode = phi3Mode
    crack.Phi3 = phi3Value
    crack.PsiS = psiSValue
    crack.CoverMode = coverMode
    formationCalculator.CrackFormationPath = formationPath
    formationCalculator.CrackFormationSolutionStrategy = formationStrategy
    formationCalculator.SolverLoadSteps = 8
    formationCalculator.SolverMaxIterations = 100
    formationCalculator.SolverToleranceN = 5#
    formationCalculator.SolverToleranceMx = 5000#
    formationCalculator.SolverToleranceMy = 5000#
    If Not stateProvider Is Nothing Then Set formationCalculator.StateProvider = stateProvider
    stage = "materials"
    Dim provider As CMaterialModelProvider
    Set provider = TestMaterialProvider()
    Dim crackedSpec As CMaterialModelSpec
    Set crackedSpec = TestCrackedStateSpec()
    Dim initiationSpec As CMaterialModelSpec
    Set initiationSpec = TestCrackInitiationSpec()

    stage = "classificationLoad"
    Dim classificationLoad As CSectionLoadState
    If loadStateForClassification Is Nothing Then
        Set classificationLoad = New CSectionLoadState
        classificationLoad.Initialize nValue, mxValue, myValue, 0#, 0#
    Else
        Set classificationLoad = loadStateForClassification
    End If
    Dim formation As CCrackFormationResult
    stage = "formation"
    Set formation = formationCalculator.CheckFormation(section, provider, crackedSpec, initiationSpec, _
        nValue, mxValue, myValue, classificationLoad, centralReferenceX, centralReferenceY)
    Dim currentState As CSectionStateResult
    stage = "currentState"
    Set currentState = New CSectionStateResult
    Dim stateRunner As CStateSolutionRunner
    Set stateRunner = New CStateSolutionRunner
    currentState.InitializeFromSolver sstCrackedState, cpCrackedNDS, crackedSpec, solver, _
        stateRunner.StateUsesExtension(section, solver, provider.ConcreteMaterialForEquilibriumFromSpec(crackedSpec), _
            provider.SteelMaterialForEquilibriumFromSpec(crackedSpec)), _
        stateRunner.StateWithinPhysicalRange(section, solver, provider.ConcreteMaterialForEquilibriumFromSpec(crackedSpec), _
            provider.SteelMaterialForEquilibriumFromSpec(crackedSpec))
    stage = "width"
    crack.Calculate currentState, section, provider, crackedSpec, formation, classificationLoad
    Set TestCalculateCrack = crack
    Exit Function
Failed:
    Err.Raise Err.Number, "modTestCrackWidth.TestCalculateCrack." & stage, Err.Source & ": " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает SLS spec текущего трещиноватого состояния: растянутый бетон Ignore.
Private Function TestCrackedStateSpec() As CMaterialModelSpec
    Dim spec As CMaterialModelSpec
    Set spec = New CMaterialModelSpec
    spec.Initialize "SLS(II)", "TwoLine", "Ignore", "TwoLine"
    Set TestCrackedStateSpec = spec
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает SLS spec Formation с активной трехлинейной растянутой ветвью бетона.
Private Function TestCrackInitiationSpec() As CMaterialModelSpec
    Dim spec As CMaterialModelSpec
    Set spec = New CMaterialModelSpec
    spec.Initialize "SLS(II)", "ThreeLine", "UseDiagram", "TwoLine"
    Set TestCrackInitiationSpec = spec
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Однотипные материалные fixtures получают готовые физические диаграммы
' соответствующей роли от provider, не конструируя отдельную расчетную методику.
Private Function TestProvisionalConcrete() As CMaterialDiagram
    Dim provider As CMaterialModelProvider
    Set provider = TestMaterialProvider()
    Set TestProvisionalConcrete = provider.ConcreteMaterial(cpCrackedNDS)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Получает физическую диаграмму бетона с растяжением для роли образования
' трещины; точки предельной деформации здесь не расширяются для равновесия.
Private Function TestProvisionalConcreteWithTension() As CMaterialDiagram
    Dim provider As CMaterialModelProvider
    Set provider = TestMaterialProvider()
    Set TestProvisionalConcreteWithTension = provider.ConcreteMaterial(cpMcrc)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Получает диаграмму арматуры роли текущего cracked-НДС из общего fixture
' provider-а, согласованную с материалами остальных crack-тестов.
Private Function TestProvisionalSteel() As CMaterialDiagram
    Dim provider As CMaterialModelProvider
    Set provider = TestMaterialProvider()
    Set TestProvisionalSteel = provider.SteelMaterial(cpCrackedNDS)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создает воспроизводимые ULS/SLS характеристики бетона и ненапрягаемой арматуры
' во внутренних единицах; пользовательский Config эти unit-fixtures не меняет.
Private Function TestMaterialProvider() As CMaterialModelProvider
    Dim steelParameters As CSteelMaterialParameters
    Set steelParameters = New CSteelMaterialParameters
    steelParameters.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#

    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.TestInitializeFromParameters TestConcreteParameters(), steelParameters
    Set TestMaterialProvider = provider
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задает фиксированные сопротивления и модуль бетонной fixture для независимых
' численных ожиданий; значения не выдаются за нормативную таблицу марки бетона.
Private Function TestConcreteParameters() As CConcreteMaterialParameters
    Dim parameters As CConcreteMaterialParameters
    Set parameters = New CConcreteMaterialParameters
    parameters.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    Set TestConcreteParameters = parameters
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет общие инварианты сформированной нормальной трещины и вручную
' собирает ширину из полученных данных. При отказе печатает Formation и Pre-state.
Private Sub TestAssertCrackCommon(ByRef stats As TCrackTestStats, ByVal prefix As String, ByVal crack As CCrackWidthCalculator)
    If Not crack.Converged Then
        TestAppendLine stats, "DIAGNOSTIC: " & prefix & "; width=" & crack.StopReason & _
            "; formation=" & crack.FormationResult.StopReason
        If prefix = "crack.user.mx" Then TestAppendLine stats, crack.FormationResult.DiagnosticLog
        Dim preState As CSectionStateResult
        Set preState = crack.FormationResult.PreCrackState
        If Not preState Is Nothing Then
            TestAppendLine stats, "PRE_STATE: physical=" & CStr(preState.WithinPhysicalRange) & _
                "; extension=" & CStr(preState.ExtensionUsed) & _
                "; epsBmin=" & TestFormatNumberInvariant(preState.MinConcreteStrain) & _
                "; epsBmax=" & TestFormatNumberInvariant(preState.MaxConcreteStrain) & _
                "; epsSmax=" & TestFormatNumberInvariant(preState.MaxSteelStrain)
        End If
    End If
    TestAssertTrue stats, prefix & ".converged", crack.Converged
    TestAssertTrue stats, prefix & ".formed", crack.CrackFormed
    TestAssertTrue stats, prefix & ".tensionBars", crack.TensionRebarCount > 0
    TestAssertTrue stats, prefix & ".sigmaPositive", crack.SigmaS > 0#
    TestAssertTrue stats, prefix & ".spacingPositive", crack.CrackSpacing > 0#
    TestAssertClose stats, prefix & ".widthFormula", crack.CrackWidth, _
        crack.Phi1 * crack.Phi2 * crack.Phi3 * crack.PsiS * (crack.SigmaS / 200000#) * crack.CrackSpacing, 0.000000001
    TestAssertClose stats, prefix & ".utilization", crack.Utilization, crack.CrackWidth / crack.AllowableCrackWidth, 0.000000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что физический fallback образования трещины не превращается
' в пользовательский NumFail. NumFail допустим только при реальной численной
' несходимости state/search, а не при штатном резервном psi_s = 1.
Private Sub TestAssertCrackCalculatorNotNumFail(ByRef stats As TCrackTestStats, _
        ByVal name As String, ByVal crack As CCrackWidthCalculator)
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    TestAssertTrue stats, name, policy.ExternalStatus(crack.ResultMeta) <> policy.NumFail
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Однотипные assertions увеличивают счетчики и сохраняют конкретный test-ID.
' Численное сравнение использует переданный абсолютный допуск без его изменения.
Private Sub TestAssertTrue(ByRef stats As TCrackTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        TestAppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        TestAppendLine stats, "FAIL: " & name
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет абсолютную разницу по допуску конкретного теста и сохраняет
' actual/expected/absDiff; инженерные критерии этим допуском не изменяются.
Private Sub TestAssertClose(ByRef stats As TCrackTestStats, ByVal name As String, ByVal actual As Double, _
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
' Добавляет одну диагностическую строку в возвращаемый отчет теста.
Private Sub TestAppendLine(ByRef stats As TCrackTestStats, ByVal text As String)
    stats.Report = stats.Report & text & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Форматирует число для текстового протокола с десятичной точкой; исходные расчётные значения не меняются.
Private Function TestFormatNumberInvariant(ByVal value As Double) As String
    TestFormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает большее из двух чисел для независимого контрольного вычисления.
Private Function MaxTest(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxTest = a Else MaxTest = b
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает меньшее из двух чисел для независимого контрольного вычисления.
Private Function MinTest(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinTest = a Else MinTest = b
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимо вычисляет extrema деформаций бетонных центров по сохраненной
' плоскости epsilon0 + kappaX*y + kappaY*x, не вызывая solver или Formation.
Private Sub TestConcreteStateStrainBounds(ByVal state As CSectionStateResult, ByVal section As CSectionModel, _
        ByRef minStrain As Double, ByRef maxStrain As Double)
    Dim i As Long
    For i = 1 To section.ConcreteCount
        Dim strain As Double
        strain = state.Epsilon0 + state.KappaX * section.ConcreteY(i) + state.KappaY * section.ConcreteX(i)
        If i = 1 Then
            minStrain = strain
            maxStrain = strain
        Else
            If strain < minStrain Then minStrain = strain
            If strain > maxStrain Then maxStrain = strain
        End If
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ============================== ДЛЯ ТЕСТОВ ==============================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет самостоятельный formation-сценарий без Width и сохранность
' результата при новом LC на том же калькуляторе, включая сброс результата.
Private Sub TestAudit02IndependentFormation(ByRef stats As TCrackTestStats)
    Dim section As CSectionModel
    Dim solver As CSectionSolver
    Set solver = TestSolveServiceState(section, -20000#, -15000000#, 0#)
    Dim provider As CMaterialModelProvider
    Set provider = TestMaterialProvider()
    Dim load As CSectionLoadState
    Set load = New CSectionLoadState
    load.Initialize -20000#, -15000000#, 0#, 0#, 0#
    Dim calculator As CCrackFormationCalculator
    Set calculator = New CCrackFormationCalculator
    calculator.CrackFormationPath = "lambda*Mxy"
    calculator.SolverLoadSteps = 8
    calculator.SolverMaxIterations = 100
    calculator.SolverToleranceN = 5#
    calculator.SolverToleranceMx = 5000#
    calculator.SolverToleranceMy = 5000#
    Dim result As CCrackFormationResult
    Set result = calculator.CheckFormation(section, provider, TestCrackedStateSpec(), _
        TestCrackInitiationSpec(), load.N, load.InternalMx, load.InternalMy, load, 0#, 0#)
    TestAssertTrue stats, "audit02.formationOnly.success", result.Converged
    TestAssertTrue stats, "audit02.formationOnly.cracked", result.CrackFormed
    TestAssertTrue stats, "audit02.formationOnly.hasPoint", result.HasLimitPoint
    TestAssertTrue stats, "audit02.formationOnly.pre", Not result.PreCrackState Is Nothing
    TestAssertTrue stats, "audit02.formationOnly.post", Not result.PostCrackState Is Nothing
    If Not result.PreCrackState Is Nothing Then
        TestAssertTrue stats, "audit02.formationOnly.prePhysical", result.PreCrackState.WithinPhysicalRange
    End If
    If Not result.PostCrackState Is Nothing Then
        TestAssertTrue stats, "audit02.formationOnly.postPhysical", result.PostCrackState.WithinPhysicalRange
    End If
    TestAssertTrue stats, "audit02.formationOnly.searchPoint", result.SearchResult.HasLimitPoint
    TestAssertTrue stats, "audit02.formationOnly.actualMethod", result.SearchResult.ActualMethod <> "Auto"
    Dim lambdaSnapshot As Double
    lambdaSnapshot = result.LambdaCrc
    load.Initialize -100000#, 0#, 0#, 0#, 0#
    Dim second As CCrackFormationResult
    Set second = calculator.CheckFormation(section, provider, TestCrackedStateSpec(), _
        TestCrackInitiationSpec(), load.N, load.InternalMx, load.InternalMy, load, 0#, 0#)
    TestAssertTrue stats, "audit02.formationOnly.reuseEmptyFixedPath", _
        second.ResultMeta.InternalStatus = rsInvalidInput And second.ResultMeta.ResultCode = rcInvalidInput
    TestAssertTrue stats, "audit02.formationOnly.reuseNoCrack", Not second.CrackFormed
    TestAssertTrue stats, "audit02.formationOnly.reuseNoPoint", Not second.HasLimitPoint
    TestAssertClose stats, "audit02.formationOnly.snapshot", result.LambdaCrc, lambdaSnapshot, 0#
    result.InitializeFromCalculator Nothing, Nothing
    TestAssertTrue stats, "audit02.formationOnly.resetPoint", Not result.HasLimitPoint
    TestAssertTrue stats, "audit02.formationOnly.resetPre", result.PreCrackState Is Nothing
    TestAssertTrue stats, "audit02.formationOnly.resetSearch", result.SearchResult Is Nothing
    TestAssertTrue stats, "audit02.formationOnly.resetMeta", result.ResultMeta.InternalStatus = rsInternalError
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Различает найденный порог за текущим LC и трещину от постоянной части пути.
' Проверяет сохранение машинного кода/предупреждения и отсутствие фиктивных
' состояний в консервативном psi=1 fallback. Путь без масштабируемой
' составляющей сохраняет ошибку ввода, но не заменяет пригодное текущее НДС.
Private Sub TestAudit02FormationOutcomeSemantics(ByRef stats As TCrackTestStats)
    Dim section As CSectionModel
    Dim solver As CSectionSolver
    Set solver = TestSolveServiceStateWithRunner(section, -20000#, -1000000#, 0#)
    Dim crack As CCrackWidthCalculator
    Set crack = TestCalculateCrack(solver, section, -20000#, -1000000#, 0#, _
        "AlwaysCalc", "Effective", formationPath:="lambda*Mxy")
    TestAssertTrue stats, "audit02.formation.aboveCurrent.converged", crack.Converged
    TestAssertTrue stats, "audit02.formation.aboveCurrent.notCracked", Not crack.CrackFormed
    TestAssertTrue stats, "audit02.formation.aboveCurrent.lambda", crack.FormationResult.LambdaCrc > 1#
    TestAssertTrue stats, "audit02.formation.aboveCurrent.foundCode", crack.FormationResult.ResultMeta.ResultCode = rcCheckPassed
    Dim search As CLimitSearchResult
    Set search = New CLimitSearchResult
    Set search = crack.FormationResult.SearchResult
    TestAssertTrue stats, "audit02.formation.aboveCurrent.searchSuccess", search.Meta.InternalStatus = rsSuccess
    TestAssertClose stats, "audit02.formation.aboveCurrent.pointLambda", search.LambdaUltimate, crack.FormationResult.LambdaCrc, 0#
    TestAssertClose stats, "audit02.formation.aboveCurrent.pointMoment", _
        Sqr(search.MxUltimate * search.MxUltimate + search.MyUltimate * search.MyUltimate), _
        crack.FormationResult.Mcrc, 0.00000000000001 * Abs(crack.FormationResult.Mcrc) ' Норма найденных усилий и lambda-момент допускают округление Double.
    TestAssertTrue stats, "audit02.formation.aboveCurrent.noPost", crack.FormationResult.PostCrackState Is Nothing

    Set solver = TestSolveServiceStateWithRunner(section, 200000#, 0#, 0#)
    Set crack = TestCalculateCrack(solver, section, 200000#, 0#, 0#, _
        "AlwaysCalc", "Effective", formationPath:="lambda*Mxy")
    Set search = crack.FormationResult.SearchResult
    TestAssertTrue stats, "audit02.formation.constant.emptyPath", search.Meta.InternalStatus = rsInvalidInput
    TestAssertTrue stats, "audit02.formation.constant.code", search.Meta.ResultCode = rcInvalidInput
    TestAssertTrue stats, "audit02.formation.constant.comment", Len(search.Meta.ResultComment) > 0
    TestAssertTrue stats, "audit02.formation.constant.noPre", crack.FormationResult.PreCrackState Is Nothing
    TestAssertTrue stats, "audit02.formation.constant.noPost", crack.FormationResult.PostCrackState Is Nothing
    TestAssertClose stats, "audit02.formation.constant.noPoint", search.LambdaUltimate, 0#, 0#
    TestAssertClose stats, "audit02.formation.constant.psi1", crack.PsiS, 1#, 0#
    search.Initialize Nothing, "Auto", vbNullString, False, 0#, 0#, 0#, 0#, Nothing, vbNullString, vbNullString
    TestAssertTrue stats, "audit02.formation.missing.internalError", search.Meta.InternalStatus = rsInternalError
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ДЛЯ ТЕСТОВ
' Проверяет реальные правила psi для сжатого/нулевого sigma_s,crc и нулевого
' текущего напряжения. Затем три режима проходят полный расчет по разрешенной
' ветви трещины от постоянной части без фиктивного PostCrackState.
Private Sub TestAudit02PsiSignedInputsAndFallbackModes(ByRef stats As TCrackTestStats)
    Dim formulaOwner As CCrackWidthCalculator
    Set formulaOwner = New CCrackWidthCalculator
    TestAssertClose stats, "audit02.psi.negativeCrc", formulaOwner.Audit02AutoPsiForTests(100#, -25#), 1#, 0#
    TestAssertClose stats, "audit02.psi.zeroCrc", formulaOwner.Audit02AutoPsiForTests(100#, 0#), 1#, 0#
    TestAssertClose stats, "audit02.psi.zeroCurrent", formulaOwner.Audit02AutoPsiForTests(0#, 25#), 1#, 0#
    TestAssertClose stats, "audit02.psi.positiveCrc", formulaOwner.Audit02AutoPsiForTests(100#, 25#), 0.8, 0.000000000001
    TestAssertClose stats, "audit02.psi.lowerBound", formulaOwner.Audit02AutoPsiForTests(100#, 200#), 0#, 0#
    Dim section As CSectionModel
    Dim solver As CSectionSolver
    Set solver = TestSolveServiceStateWithRunner(section, 200000#, 0#, 0#)
    Dim mode As Variant
    For Each mode In Array("User", "Auto", "AlwaysCalc")
        Dim crack As CCrackWidthCalculator
        Set crack = TestCalculateCrack(solver, section, 200000#, 0#, 0#, _
            CStr(mode), "Effective", allowable:=0.0001, psiSValue:=0.65, formationPath:="lambda*Mxy")
        Dim prefix As String
        prefix = "audit02.psi.fallback." & CStr(mode)
        TestAssertTrue stats, prefix & ".converged", crack.Converged
        TestAssertTrue stats, prefix & ".formationCode", crack.FormationResult.ResultMeta.ResultCode = rcInvalidInput
        TestAssertTrue stats, prefix & ".noPost", crack.FormationResult.PostCrackState Is Nothing
        TestAssertClose stats, prefix & ".value", crack.PsiS, 1#, 0#
        TestAssertTrue stats, prefix & ".warningComment", _
            InStr(1, crack.ResultMeta.ResultComment, "Предупреждение", vbTextCompare) > 0
        TestAssertTrue stats, prefix & ".upperBound", crack.PsiS <= 1#
        TestAssertCrackCalculatorNotNumFail stats, prefix & ".notNumerical", crack
    Next mode
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' точка за текущей нагрузкой и отказ недопустимого явного пути.
Public Function RunFormationOutcomeTests() As String
    Dim stats As TCrackTestStats
    On Error GoTo Failed
    TestAudit02FormationOutcomeSemantics stats
    TestAppendLine stats, "TOTAL_FORMATION_OUTCOME: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunFormationOutcomeTests = stats.Report
    Exit Function
Failed:
    RunFormationOutcomeTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ============================== ДЛЯ ТЕСТОВ AUDIT03 ==============================
' Отдельный entrypoint проверяет terminal ошибки state-solve для всех физических
' путей трещинообразования, включая отсутствие незаявленных Auto-переходов.
Public Function TestRunAudit03FormationContracts() As String
    Dim stats As TCrackTestStats
    TestAudit03FormationTypedFailures stats
    TestAppendLine stats, "TOTAL_AUDIT03_FORMATION_CONTRACTS: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    TestRunAudit03FormationContracts = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ошибочный метод/бюджет SectionSolver проходят через настоящие
' lambda-пробы; config failure не становится NumFail и не запускает новые пути.
Private Sub TestAudit03FormationTypedFailures(ByRef stats As TCrackTestStats)
    Dim section As CSectionModel, serviceSolver As CSectionSolver
    Set serviceSolver = TestSolveServiceState(section, -20000#, -15000000#, 0#)
    Dim provider As CMaterialModelProvider
    Set provider = TestMaterialProvider()
    Dim load As CSectionLoadState
    Set load = New CSectionLoadState
    load.Initialize -20000#, -15000000#, 0#, 0#, 0#
    Dim path As Variant, scenario As Long, strategy As Variant
    For Each strategy In Array("LoadMultiplier", "UltimateStrain")
    For Each path In Array("lambda*Mxy", "lambda*N", "lambda*NMxy", "Auto")
        For scenario = 1 To 3
            Dim calculator As CCrackFormationCalculator
            Set calculator = New CCrackFormationCalculator
            calculator.CrackFormationPath = CStr(path)
            calculator.CrackFormationSolutionStrategy = CStr(strategy)
            calculator.SolverLoadSteps = 8
            calculator.SolverMaxIterations = 100
            calculator.SolverToleranceN = 5#
            calculator.SolverToleranceMx = 5000#
            calculator.SolverToleranceMy = 5000#
            If scenario = 1 Then
                calculator.SolverMethod = "Invalid"
            ElseIf scenario = 2 Then
                calculator.SolverMaxIterations = 0
            Else
                calculator.SolverToleranceN = -1#
            End If
            Dim result As CCrackFormationResult
            Set result = calculator.CheckFormation(section, provider, TestCrackedStateSpec(), _
                TestCrackInitiationSpec(), load.N, load.InternalMx, load.InternalMy, load, 0#, 0#)
            Dim prefix As String
            prefix = "audit03.formation.typed." & CStr(strategy) & "." & CStr(path) & "." & CStr(scenario)
            TestAssertTrue stats, prefix & ".status", result.ResultMeta.InternalStatus = rsInvalidConfiguration
            TestAssertTrue stats, prefix & ".code", result.ResultMeta.ResultCode = rcInvalidConfiguration
            TestAssertTrue stats, prefix & ".notCalculated", Not result.ResultMeta.Calculated
            If CStr(strategy) = "LoadMultiplier" Then
                TestAssertTrue stats, prefix & ".oneAttempt", result.SolverCallCount = 1
            Else
                TestAssertTrue stats, prefix & ".noNumericalAttempt", result.SolverCallCount = 0
            End If
            TestAssertTrue stats, prefix & ".noPoint", Not result.HasLimitPoint
            TestAssertTrue stats, prefix & ".noPre", result.PreCrackState Is Nothing
            TestAssertTrue stats, prefix & ".noPost", result.PostCrackState Is Nothing
            TestAssertTrue stats, prefix & ".reason", Len(result.ResultMeta.ResultComment) > 0
            TestAppendLine stats, "COMMENT: " & prefix & "; " & result.ResultMeta.ResultComment
        Next scenario
    Next path
    Next strategy
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет факт численного Search отдельно от аналитического
' Ncrc и получения Pre/PostState. Повторные вызовы одного calculator-а не
' должны переносить флаг выполнения или найденную точку между сочетаниями.
Public Function TestRunAudit03FormationLifecycle() As String
    On Error GoTo Failed
    Dim stats As TCrackTestStats
    TestAudit03FormationSearchLifecycle stats
    TestAppendLine stats, "TOTAL_AUDIT03_FORMATION_LIFECYCLE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    TestRunAudit03FormationLifecycle = stats.Report
    Exit Function
Failed:
    TestRunAudit03FormationLifecycle = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' реальные нагрузки охватывают ранний выход, аналитическую точку
' с/без состояний, численный поиск и ошибочный фиксированный путь до проб.
' Невалидный метод solver-а отдельно доказывает попытку поиска без успеха.
Private Sub TestAudit03FormationSearchLifecycle(ByRef stats As TCrackTestStats)
    Dim section As CSectionModel, setupSolver As CSectionSolver
    Set setupSolver = TestSolveServiceState(section, 0#, 0#, 0#)
    Dim provider As CMaterialModelProvider, load As CSectionLoadState
    Set provider = TestMaterialProvider()
    Set load = New CSectionLoadState
    Dim calculator As CCrackFormationCalculator
    Set calculator = New CCrackFormationCalculator
    calculator.SolverLoadSteps = 8
    calculator.SolverMaxIterations = 100
    calculator.SolverToleranceN = 5#
    calculator.SolverToleranceMx = 5000#
    calculator.SolverToleranceMy = 5000#
    Dim strategy As Variant, scenario As Long, nValue As Double, mxValue As Double
    Dim result As CCrackFormationResult, search As CLimitSearchResult, retained As CLimitSearchResult
    For Each strategy In Array("Auto", "UltimateStrain", "LoadMultiplier")
        calculator.CrackFormationSolutionStrategy = CStr(strategy)
        For scenario = 0 To 5
            calculator.CrackFormationPath = "Auto"
            calculator.SolverMethod = "Newton"
            nValue = 0#: mxValue = 0#
            Select Case scenario
                Case 1: nValue = 1000#
                Case 2: nValue = 250000#
                Case 3: nValue = -20000#: mxValue = -15000000#
                Case 4
                    nValue = -20000#: mxValue = -15000000#
                    calculator.CrackFormationPath = "lambda*My"
                Case 5
                    nValue = -20000#: mxValue = -15000000#
                    calculator.SolverMethod = "Invalid"
            End Select
            load.Initialize nValue, mxValue, 0#, 0#, 0#
            Set result = calculator.CheckFormation(section, provider, TestCrackedStateSpec(), _
                TestCrackInitiationSpec(), nValue, mxValue, 0#, load, 0#, 0#)
            Set search = result.SearchResult
            Dim prefix As String, expectedExecuted As Boolean
            prefix = "audit03.formation.lifecycle." & CStr(strategy) & "." & CStr(scenario)
            ' UltimateStrain проверяет настройки до численных проб; LoadMultiplier
            ' получает тот же типизированный отказ в своей первой пробе.
            expectedExecuted = (scenario = 3 Or (scenario = 5 And CStr(strategy) = "LoadMultiplier"))
            TestAssertTrue stats, prefix & ".searchExists", Not search Is Nothing
            If Not search Is Nothing Then
                TestAssertTrue stats, prefix & ".searchExecuted", search.SearchExecuted = expectedExecuted
                TestAssertTrue stats, prefix & ".cloneFlag", search.Clone.SearchExecuted = expectedExecuted
                TestAssertTrue stats, prefix & ".metaSnapshot", search.Meta.Calculated = result.ResultMeta.Calculated
                If scenario = 3 Then Set retained = search
            End If
            If scenario = 0 Or scenario = 4 Then
                TestAssertTrue stats, prefix & ".notCalculated", Not result.ResultMeta.Calculated
                TestAssertTrue stats, prefix & ".noSolver", result.SolverCallCount = 0
                TestAssertTrue stats, prefix & ".noPoint", Not result.HasLimitPoint
                TestAssertTrue stats, prefix & ".noPre", result.PreCrackState Is Nothing
                TestAssertTrue stats, prefix & ".noPost", result.PostCrackState Is Nothing
            ElseIf scenario = 1 Then
                TestAssertTrue stats, prefix & ".analyticalCalculated", result.ResultMeta.Calculated
                TestAssertTrue stats, prefix & ".analyticalNoSolver", result.SolverCallCount = 0
                TestAssertTrue stats, prefix & ".thresholdBeyondLoad", result.LambdaCrc > 1#
            ElseIf scenario = 2 Then
                TestAssertTrue stats, prefix & ".analyticalCalculated", result.ResultMeta.Calculated
                TestAssertTrue stats, prefix & ".namedSolves", result.SolverCallCount > 0
                TestAssertTrue stats, prefix & ".preState", Not result.PreCrackState Is Nothing
                TestAssertTrue stats, prefix & ".postState", Not result.PostCrackState Is Nothing
            ElseIf scenario = 3 Then
                TestAssertTrue stats, prefix & ".numericCalculated", result.ResultMeta.Calculated
                TestAssertTrue stats, prefix & ".numericPoint", result.HasLimitPoint
            ElseIf scenario = 5 Then
                TestAssertTrue stats, prefix & ".typedConfiguration", result.ResultMeta.InternalStatus = rsInvalidConfiguration
                TestAssertTrue stats, prefix & ".validationLifecycle", (result.SolverCallCount > 0) = expectedExecuted
                TestAssertTrue stats, prefix & ".noPoint", Not result.HasLimitPoint
            End If
            If scenario >= 4 Then TestAssertTrue stats, prefix & ".priorSnapshot", retained.SearchExecuted And retained.HasLimitPoint
            TestAppendLine stats, "FORMATION_LIFECYCLE: " & prefix & "; SearchExecuted=" & CStr(search.SearchExecuted) & _
                "; Calculated=" & CStr(result.ResultMeta.Calculated) & "; solves=" & CStr(result.SolverCallCount) & _
                "; comment=" & result.ResultMeta.ResultComment
        Next scenario
        calculator.SolverMethod = "Newton": calculator.CrackFormationPath = "Auto" ' Нулевой повтор после отказа дополнительно проверяет reset между LC.
        load.Initialize 0#, 0#, 0#, 0#, 0#
        Set result = calculator.CheckFormation(section, provider, TestCrackedStateSpec(), _
            TestCrackInitiationSpec(), 0#, 0#, 0#, load, 0#, 0#)
        TestAssertTrue stats, "audit03.formation.lifecycle." & CStr(strategy) & ".finalReset", Not result.SearchResult.SearchExecuted
    Next strategy
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ищет воспроизводимые случаи неположительного sigma_s,crc
' на настоящих Formation/PostCrackState, а не на подставленных числах формулы.
' Явные параметры и конечная сетка нагрузок не меняют пользовательский Config.
Public Function RunAudit03SigmaSCrcBoundaryTests() As String
    On Error GoTo Failed
    Dim stats As TCrackTestStats
    TestAudit03SigmaSCrcBoundary stats
    TestAppendLine stats, "TOTAL_AUDIT03_SIGMA_CRC_BOUNDARY: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03SigmaSCrcBoundaryTests = stats.Report
    Exit Function
Failed:
    RunAudit03SigmaSCrcBoundaryTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сначала находит допустимый CurrentCrackedState общим runner-ом.
' Только затем Formation и Width проверяются в обоих режимах psi/усреднения;
' физически непригодный текущий state не используется как вход формулы.
Private Sub TestAudit03SigmaSCrcBoundary(ByRef stats As TCrackTestStats, Optional ByVal includeUser As Boolean = False)
    Dim section As CSectionModel
    Set section = TestAudit03SigmaBoundarySection()
    Dim provider As CMaterialModelProvider
    Set provider = TestMaterialProvider()
    Dim spec As CMaterialModelSpec
    Set spec = TestCrackedStateSpec()
    Dim nFactor As Variant, eccentricityFactor As Variant, psiMode As Variant, averagingMode As Variant
    Dim nValue As Double, mxValue As Double, cases As Long, nonpositiveCases As Long, noPostTensionCases As Long
    Dim noCurrentTensionCases As Long ' Отрицательный контроль: при отсутствии текущего растяжения формула не нужна.
    Dim modes As Variant
    modes = Array("AlwaysCalc", "Auto")
    If includeUser Then modes = Array("AlwaysCalc", "Auto", "User")
    For Each nFactor In Array(0.1, 0.25, 0.4, 0.6, 0.75)
        For Each eccentricityFactor In Array(1.25, 1.75, 2.5, 3.5)
            nValue = -CDbl(nFactor) * 60000# * 22#
            mxValue = -nValue * (200# / 6#) * CDbl(eccentricityFactor)
            Dim runner As CStateSolutionRunner
            Set runner = New CStateSolutionRunner
            ConfigureTestStateRunner runner
            runner.Solve section, provider.ConcreteMaterialForEquilibriumFromSpec(spec), _
                provider.SteelMaterialForEquilibriumFromSpec(spec), nValue, mxValue, 0#, True
            TestAppendLine stats, "SIGMA_CRC_CURRENT: N=" & TestFormatNumberInvariant(nValue) & _
                "|Mx=" & TestFormatNumberInvariant(mxValue) & "|converged=" & CStr(runner.Converged) & _
                "|physical=" & CStr(runner.WithinPhysicalRange) & _
                "|currentSteelTension=" & CStr(runner.ResultSolver.HasSteelTension) & _
                "|maxSteelStrain=" & TestFormatNumberInvariant(runner.ResultSolver.MaxSteelStrain) & "|reason=" & runner.StopReason
            If runner.Converged And runner.WithinPhysicalRange Then
                For Each psiMode In modes
                    For Each averagingMode In Array("AllSelected", "TensionOnly")
                        Dim crack As CCrackWidthCalculator
                        Set crack = TestCalculateCrack(runner.ResultSolver, section, nValue, mxValue, 0#, _
                            CStr(psiMode), "Effective", allowable:=0.00000001, _
                            sigmaSCrcAveragingMode:=CStr(averagingMode), psiSValue:=0.65)
                        cases = cases + 1
                        Dim prefix As String
                        prefix = "audit03.sigmaCrc.boundary." & CStr(cases)
                        TestAppendLine stats, "SIGMA_CRC_BOUNDARY: " & prefix & "|N=" & TestFormatNumberInvariant(nValue) & _
                            "|Mx=" & TestFormatNumberInvariant(mxValue) & "|psiMode=" & CStr(psiMode) & _
                            "|averaging=" & CStr(averagingMode) & "|point=" & CStr(crack.FormationResult.HasLimitPoint) & _
                            "|widthCalculated=" & CStr(crack.ResultMeta.Calculated) & "|sigmaS=" & TestFormatNumberInvariant(crack.SigmaS) & _
                            "|sigmaSCrc=" & TestFormatNumberInvariant(crack.SigmaSCrc) & "|psi=" & TestFormatNumberInvariant(crack.PsiS) & _
                            "|comment=" & crack.ResultMeta.ResultComment
                        If includeUser And crack.FormationResult.HasLimitPoint And crack.FormationResult.CrackFormed Then
                            If Not crack.FormationResult.PostCrackState.HasSteelTension Then
                                If runner.ResultSolver.HasSteelTension Then
                                    noPostTensionCases = noPostTensionCases + 1
                                    TestAssertTrue stats, prefix & ".noPostTension.widthCalculated", crack.ResultMeta.Calculated
                                    TestAssertTrue stats, prefix & ".noPostTension.widthReady", crack.Converged
                                    TestAssertTrue stats, prefix & ".noPostTension.currentSigmaPositive", crack.SigmaS > 0#
                                    TestAssertTrue stats, prefix & ".noPostTension.notNA", crack.ResultMeta.InternalStatus <> rsNotApplicable
                                    TestAssertClose stats, prefix & ".noPostTension.psiOne", crack.PsiS, 1#, 0#
                                    TestAssertTrue stats, prefix & ".noPostTension.warning", InStr(1, crack.ResultMeta.ResultComment, "Предупреждение", vbTextCompare) > 0
                                    TestAssertClose stats, prefix & ".noPostTension.formula", crack.CrackWidth, _
                                        crack.Phi1 * crack.Phi2 * crack.Phi3 * crack.SigmaS / crack.SteelEs * crack.CrackSpacing, 0.000000001
                                    Dim widthResult As CCrackWidthResult, region As CConcreteRegion
                                    Set widthResult = New CCrackWidthResult
                                    widthResult.InitializeFromCalculator crack, crack.ResultMeta, prefix
                                    TestAssertTrue stats, prefix & ".noPostTension.resultComment", _
                                        widthResult.ResultMeta.ResultComment = crack.ResultMeta.ResultComment
                                    Set region = widthResult.InteractionRegion
                                    TestAssertTrue stats, prefix & ".noPostTension.resultRegion", Not region Is Nothing
                                    If Not region Is Nothing Then
                                        TestAssertClose stats, prefix & ".noPostTension.regionArea", region.Area, crack.Abt, 0#
                                        TestAssertTrue stats, prefix & ".noPostTension.regionOwner", region.OwnerID = prefix
                                    End If
                                Else
                                    noCurrentTensionCases = noCurrentTensionCases + 1
                                    TestAssertTrue stats, prefix & ".noCurrentTension.notCalculated", Not crack.ResultMeta.Calculated
                                    TestAssertTrue stats, prefix & ".noCurrentTension.notApplicable", crack.ResultMeta.InternalStatus = rsNotApplicable
                                    TestAssertClose stats, prefix & ".noCurrentTension.psiUnchanged", crack.PsiS, 0.65, 0#
                                    TestAssertTrue stats, prefix & ".noCurrentTension.noPsiWarning", _
                                        InStr(1, crack.ResultMeta.ResultComment, "psi_s принят равным 1", vbTextCompare) = 0
                                End If
                            End If
                        End If
                        If crack.FormationResult.HasLimitPoint And crack.ResultMeta.Calculated Then
                            TestAssertTrue stats, prefix & ".psiRange", crack.PsiS >= 0# And crack.PsiS <= 1#
                            TestAssertTrue stats, prefix & ".widthNotNumerical", crack.ResultMeta.InternalStatus <> rsNumericalFailure
                            If CStr(averagingMode) = "TensionOnly" Then _
                                TestAssertTrue stats, prefix & ".tensionOnlyNonnegative", crack.SigmaSCrc >= 0#
                            If crack.SigmaSCrc <= 0.000000001 And CStr(psiMode) <> "User" Then
                                nonpositiveCases = nonpositiveCases + 1
                                TestAssertClose stats, prefix & ".nonpositivePsiOne", crack.PsiS, 1#, 0.000000000001
                                TestAssertTrue stats, prefix & ".nonpositiveReason", _
                                    InStr(1, crack.ResultMeta.ResultComment, "неполож", vbTextCompare) > 0
                            End If
                        End If
                    Next averagingMode
                Next psiMode
            End If
        Next eccentricityFactor
    Next nFactor
    TestAssertTrue stats, "audit03.sigmaCrc.boundary.actualPreparedNonpositive", nonpositiveCases > 0
    If includeUser Then
        TestAssertTrue stats, "sp35.sp63.noPostTension.actualReadyFormation", noPostTensionCases > 0
        TestAssertTrue stats, "sp35.sp63.noPostTension.noCurrentTensionControl", noCurrentTensionCases > 0
    End If
    TestAppendLine stats, "SIGMA_CRC_BOUNDARY_COUNTS: calculatedVariants=" & CStr(cases) & _
        "|nonpositivePrepared=" & CStr(nonpositiveCases) & "|noPostTensionWithCurrentTension=" & CStr(noPostTensionCases) & _
        "|noCurrentTension=" & CStr(noCurrentTensionCases)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' стержни находятся глубже крайнего бетонного волокна.
' При большом сжатии в момент образования трещины они могут быть сжатыми,
' хотя в более позднем текущем НДС попадут в выбранную растянутую группу.
Private Function TestAudit03SigmaBoundarySection() As CSectionModel
    Dim geom As CGeometryRegion
    Set geom = TestFixtureBox(300#, 200#)
    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 20#, 20#, 1
    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -90#, -60#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B2", 90#, -60#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B3", -90#, 60#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B4", 90#, 60#, 20#, 0#, "Rebar", "", geom
    Set TestAudit03SigmaBoundarySection = TestBuildGeneratedSectionModel(mesh, rebars)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ============================== ДЛЯ ТЕСТОВ: FORMATION GATE ==============================
' различает доказанное отсутствие нормальной трещины и отсутствие
' пригодной точки из-за ошибки поиска. Во втором случае проверяется psi=1,
' в первом формула не должна получать даже отсутствующее текущее состояние.
Public Function RunAudit03CrackFormationGateTests() As String
    Dim stats As TCrackTestStats
    TestAudit03ConfirmedNotCrackedSkipsWidth stats
    TestFormationFailureStillCalculatesWidth stats
    TestFormationFallbackWithUniformCurrentState stats
    TestAppendLine stats, "TOTAL_AUDIT03_CRACK_FORMATION_GATE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03CrackFormationGateTests = stats.Report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' успешное CRITERION_NOT_REACHED является подтвержденным
' NotCracked даже без Pre/Post и пороговой точки. Настройки коэффициентов и
' PsiMode не должны включать формулу; продольные трещины используют только
' свое текущее НДС и независимо могут дать как OK, так и FAIL.
Private Sub TestAudit03ConfirmedNotCrackedSkipsWidth(ByRef stats As TCrackTestStats)
    Dim owner As CCrackFormationCalculator, formation As CCrackFormationResult, meta As CResultMeta
    Dim width As CCrackWidthCalculator, mode As Variant, status As Variant, prefix As String
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    Set owner = New CCrackFormationCalculator
    owner.LimitSearchFinishNoCrack "Подтверждено: нормальная трещина при заданной нагрузке не образуется."
    For Each status In Array(rsSuccess, rsSuccessWithWarning)
        Set meta = New CResultMeta
        meta.SetResult CLng(status), rcCriterionNotReached, rkCrackFormation, owner.StopReason
        Set formation = New CCrackFormationResult
        formation.InitializeFromCalculator owner, meta
        TestAssertTrue stats, "crack.confirmedNotCracked.noPoint." & CStr(status), Not formation.HasLimitPoint
        For Each mode In Array("User", "Auto", "AlwaysCalc")
            prefix = "crack.confirmedNotCracked." & CStr(status) & "." & CStr(mode)
            Set width = New CCrackWidthCalculator
            width.PsiMode = CStr(mode): width.PsiS = 0.25
            width.Calculate Nothing, Nothing, Nothing, Nothing, formation, Nothing
            TestAssertTrue stats, prefix & ".notFormed", Not width.CrackFormed
            TestAssertTrue stats, prefix & ".notApplicable", width.ResultMeta.InternalStatus = rsNotApplicable
            TestAssertTrue stats, prefix & ".notCalculated", Not width.ResultMeta.Calculated
            TestAssertTrue stats, prefix & ".external", policy.ExternalStatus(width.ResultMeta) = "N/A"
            TestAssertTrue stats, prefix & ".reason", InStr(1, width.ResultMeta.ResultComment, "не образуется", vbTextCompare) > 0
            TestAssertTrue stats, prefix & ".noFallbackWarning", InStr(1, width.ResultMeta.ResultComment, "psi_s", vbTextCompare) = 0
            TestAssertClose stats, prefix & ".psiPreserved", width.PsiS, 0.25, 0#
            TestAssertClose stats, prefix & ".zeroWidth", width.CrackWidth, 0#, 0#
        Next mode
    Next status
    Dim section As CSectionModel, solver As CSectionSolver, current As CSectionStateResult
    Set solver = TestSolveServiceState(section, -200000#, 0#, 0#)
    Set current = New CSectionStateResult
    current.InitializeFromSolver sstCrackedState, cpCrackedNDS, TestCrackedStateSpec(), solver, False, True
    Dim longitudinal As CLongitudinalCrackCalculator, result As CLongitudinalCrackResult
    Set longitudinal = New CLongitudinalCrackCalculator
    Set result = longitudinal.Calculate(current, 100#)
    TestAssertTrue stats, "crack.confirmedNotCracked.longitudinalPass", result.ResultMeta.InternalStatus = rsSuccess
    TestAssertTrue stats, "crack.confirmedNotCracked.longitudinalCalculated", result.ResultMeta.Calculated
    Set result = longitudinal.Calculate(current, 0.000001)
    TestAssertTrue stats, "crack.confirmedNotCracked.longitudinalFail", result.ResultMeta.InternalStatus = rsCheckFailed
    TestAssertTrue stats, "crack.confirmedNotCracked.longitudinalFailCalculated", result.ResultMeta.Calculated
    TestAssertTrue stats, "crack.confirmedNotCracked.formationPreserved", formation.ResultMeta.InternalStatus = rsSuccessWithWarning
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ============================== ДЛЯ ТЕСТОВ AUDIT03 ==============================
' Выполняет отдельный reproducer границы fixed-plane callback Formation.
' Здесь не ищется еще одно равновесие: проверяется сохранение точной причины
' EvaluateStrainPlane и отсутствие фиктивного предельного состояния.
Public Function TestRunAudit03FormationResidualContracts() As String
    On Error GoTo Failed
    Dim stats As TCrackTestStats
    TestAudit03FormationResidualContracts stats
    TestAppendLine stats, "TOTAL_AUDIT03_FORMATION_RESIDUAL: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    TestRunAudit03FormationResidualContracts = stats.Report
    Exit Function
Failed:
    TestRunAudit03FormationResidualContracts = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' реальные section/materials и все пять путей используются для
' отсутствующих зависимостей, пустого бетона, переполнения и уже установленной
' terminal причины. Повтор CheckFormation на том же calculator проверяет reset.
Private Sub TestAudit03FormationResidualContracts(ByRef stats As TCrackTestStats)
    Dim validSection As CSectionModel, setupSolver As CSectionSolver
    Set setupSolver = TestSolveServiceState(validSection, -20000#, -15000000#, 0#)
    Dim provider As CMaterialModelProvider, load As CSectionLoadState
    Set provider = TestMaterialProvider(): Set load = New CSectionLoadState
    load.Initialize -20000#, -15000000#, -7500000#, 0#, 0#
    Dim recoveryLoad As CSectionLoadState
    Set recoveryLoad = New CSectionLoadState
    recoveryLoad.Initialize -20000#, -15000000#, 0#, 0#, 0#
    Dim path As Variant, scenario As Long, descriptor As CLoadPathDescriptor
    Dim calculator As CCrackFormationCalculator, section As CSectionModel
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram, vector As CLoadPathVector
    Dim solver As CSectionSolver, search As CLimitSearchResult, recovery As CCrackFormationResult
    Dim eps0 As Double, kx As Double, r1 As Double, r2 As Double, rLimit As Double, lambdaValue As Double
    Dim state As String, evaluated As Boolean, prefix As String, expectedStatus As EResultInternalStatus
    Dim expectedCode As EResultCode, policy As CResultStatusPolicy, expectedDisplay As String
    Set policy = New CResultStatusPolicy
    For Each path In Array("LambdaMx", "LambdaMy", "LambdaMxy", "LambdaN", "LambdaNMxy")
        For scenario = 1 To 8
            Set calculator = New CCrackFormationCalculator
            calculator.DiagnosticsEnabled = True
            Set section = validSection
            Set concrete = provider.ConcreteMaterialFromSpec(TestCrackInitiationSpec())
            Set steel = provider.SteelMaterialFromSpec(TestCrackInitiationSpec())
            Set descriptor = New CLoadPathDescriptor
            descriptor.InitializeFromLoadState CStr(path), load
            Set vector = descriptor.Vector
            eps0 = 0#: kx = 0#: state = vbNullString: Set solver = Nothing
            expectedStatus = rsInternalError: expectedCode = rcInternalError: expectedDisplay = "CalcErr"
            Select Case scenario
                Case 1: Set section = Nothing
                Case 2
                    Set section = New CSectionModel
                    expectedStatus = rsInvalidInput: expectedCode = rcInvalidInput: expectedDisplay = "InputErr"
                Case 3: Set concrete = Nothing
                Case 4: Set steel = Nothing
                Case 5: Set vector = Nothing
                Case 6
                    eps0 = 1E+308: kx = 1E+308
                    expectedStatus = rsNumericalFailure: expectedCode = rcNumericalFailure: expectedDisplay = "NumFail"
                Case 7
                    calculator.LimitSearchSetFailure sfcInvalidConfiguration, "Контрольная первичная ошибка настройки Formation."
                    Set vector = Nothing
                    expectedStatus = rsInvalidConfiguration: expectedCode = rcInvalidConfiguration: expectedDisplay = "InputErr"
                Case 8
                    calculator.LimitSearchSetFailure sfcInternalError, "Контрольная первичная ошибка контракта Formation."
                    Set vector = Nothing
            End Select
            prefix = "audit03.formationResidual." & CStr(path) & "." & CStr(scenario)
            evaluated = calculator.LimitSearchEvaluateUltimateResidual(section, concrete, steel, vector, _
                eps0, kx, 0#, solver, r1, r2, rLimit, lambdaValue, state)
            Set search = calculator.BuildSearchSnapshot("UltimateStrain")
            TestAssertTrue stats, prefix & ".noResidual", Not evaluated
            TestAssertTrue stats, prefix & ".status", search.Meta.InternalStatus = expectedStatus
            TestAssertTrue stats, prefix & ".code", search.Meta.ResultCode = expectedCode
            TestAssertTrue stats, prefix & ".display", policy.ExternalStatus(search.Meta) = expectedDisplay
            TestAssertTrue stats, prefix & ".noPoint", Not search.HasLimitPoint And Not calculator.HasLimitPoint
            TestAssertTrue stats, prefix & ".noState", search.PointState Is Nothing
            TestAssertTrue stats, prefix & ".reason", Len(Trim$(search.Meta.ResultComment)) > 0
            If scenario = 5 Or scenario = 6 Then
                TestAssertTrue stats, prefix & ".localizedRuntimeReason", _
                    InStr(1, search.Meta.ResultComment, "Object variable", vbTextCompare) = 0 And _
                    InStr(1, search.Meta.ResultComment, "Overflow", vbTextCompare) = 0
                Dim runtimeNumber As Long, runtimeDescription As String
                If scenario = 5 Then
                    runtimeNumber = 91
                    ' Общая проверка нагрузки сохраняет код 91, но поясняет отсутствие
                    ' данных своим русским текстом вместо системного английского.
                    runtimeDescription = "не передана модель сечения или описание изменения N, Mx и My."
                Else
                    runtimeNumber = 6: runtimeDescription = Error$(runtimeNumber)
                End If
                TestAssertTrue stats, prefix & ".runtimeDiagnostic", _
                    InStr(1, search.DiagnosticLog, "formationResidualRuntimeError=" & CStr(runtimeNumber), vbBinaryCompare) > 0 And _
                    InStr(1, search.DiagnosticLog, runtimeDescription, vbBinaryCompare) > 0
            End If
            If scenario >= 7 Then
                TestAssertTrue stats, prefix & ".noTerminalProbe", calculator.SolverCallCount = 0
                TestAssertTrue stats, prefix & ".primaryReason", InStr(1, search.Meta.ResultComment, "Контрольная первичная", vbBinaryCompare) = 1
            Else
                TestAssertTrue stats, prefix & ".oneEvaluation", calculator.SolverCallCount = 1
            End If
            TestAppendLine stats, "FORMATION_RESIDUAL_COMMENT: " & prefix & "; " & search.Meta.ResultComment
            Set recovery = calculator.CheckFormation(validSection, provider, TestCrackedStateSpec(), _
                TestCrackInitiationSpec(), -20000#, -15000000#, 0#, recoveryLoad, 0#, 0#)
            TestAssertTrue stats, prefix & ".recoverySuccess", recovery.ResultMeta.InternalStatus = rsSuccess Or _
                recovery.ResultMeta.InternalStatus = rsSuccessWithWarning
            TestAssertTrue stats, prefix & ".recoveryPoint", recovery.HasLimitPoint
            TestAssertTrue stats, prefix & ".recoveryReasonClean", InStr(1, recovery.ResultMeta.ResultComment, "Контрольная первичная", vbBinaryCompare) = 0
        Next scenario
    Next path
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ==================== ДЛЯ ТЕСТОВ: ПРИЧИНЫ ОСТАНОВКИ FORMATION-АДАПТЕРА ====================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет typed-протокол callback-ов общего Search без решения равновесия.
' Ошибка входа и неизвестный внутренний probe-исход не являются NumFail;
' вторичная численная остановка не должна подменять первичную terminal-причину.
Private Sub TestAudit03FormationAdapterFailures(ByRef stats As TCrackTestStats)
    Dim calculator As CCrackFormationCalculator, adapter As CCrackLimitSearchProblem
    Dim context As ILimitSearchProblem, result As CLimitSearchResult, policy As CResultStatusPolicy
    Dim scenario As Long, prefix As String, reason As String, expectedDisplay As String
    Dim expectedStatus As EResultInternalStatus, expectedCode As EResultCode
    Set policy = New CResultStatusPolicy
    For scenario = 1 To 5
        Set calculator = New CCrackFormationCalculator
        Set adapter = New CCrackLimitSearchProblem: adapter.Initialize calculator
        Set context = adapter
        reason = "Контрольная причина адаптера образования трещины."
        expectedStatus = rsInvalidInput: expectedCode = rcInvalidInput: expectedDisplay = "InputErr"
        Select Case scenario
            Case 1, 4: context.SetInvalidInput reason
            Case 2
                context.HandleUnexpectedState "Неизвестная контрольная проба"
                reason = "Неизвестная контрольная проба"
                expectedStatus = rsInternalError: expectedCode = rcInternalError: expectedDisplay = "CalcErr"
            Case 3
                context.SetNumericalFailure reason
                expectedStatus = rsNumericalFailure: expectedCode = rcNumericalFailure: expectedDisplay = "NumFail"
            Case 5
                context.SetFailure sfcInternalError, reason
                expectedStatus = rsInternalError: expectedCode = rcInternalError: expectedDisplay = "CalcErr"
        End Select
        If scenario >= 4 Then context.SetNumericalFailure "Вторичная контрольная численная остановка."
        Set result = calculator.BuildSearchSnapshot("LoadMultiplier")
        prefix = "audit03.formationAdapterFailure." & CStr(scenario)
        TestAssertTrue stats, prefix & ".status", result.Meta.InternalStatus = expectedStatus
        TestAssertTrue stats, prefix & ".code", result.Meta.ResultCode = expectedCode
        TestAssertTrue stats, prefix & ".display", policy.ExternalStatus(result.Meta) = expectedDisplay
        TestAssertTrue stats, prefix & ".reason", InStr(1, result.Meta.ResultComment, reason, vbBinaryCompare) > 0
        TestAssertTrue stats, prefix & ".noPoint", Not result.HasLimitPoint
        TestAssertTrue stats, prefix & ".noState", result.PointState Is Nothing
        TestAssertTrue stats, prefix & ".noSolve", calculator.SolverCallCount = 0
        If scenario <> 3 Then TestAssertTrue stats, prefix & ".terminal", context.StoppedTerminally
        TestAppendLine stats, "FORMATION_ADAPTER_FAILURE: " & prefix & "; " & result.Meta.ResultComment
    Next scenario
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает отдельный протокол callback-классификации, не меняя материалы,
' настройки книги или существующие численные expected и tolerance.
Public Function RunAudit03FormationAdapterFailuresTests() As String
    On Error GoTo Failed
    Dim stats As TCrackTestStats
    TestAudit03FormationAdapterFailures stats
    TestAppendLine stats, "TOTAL_FORMATION_ADAPTER_FAILURES: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03FormationAdapterFailuresTests = stats.Report
    Exit Function
Failed:
    RunAudit03FormationAdapterFailuresTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ==================== ДЛЯ ТЕСТОВ: КОНЕЧНЫЕ ЭКСТРЕМАЛЬНЫЕ МОМЕНТЫ ====================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает отдельный протокол Formation для больших конечных Double.
' Проверяет завершение и typed-результат, но не обещает сходимость за пределами
' диаграммы и не требует искусственной точки трещинообразования.
Public Function RunAudit03ExtremeFormationTests() As String
    On Error GoTo Failed
    Dim stats As TCrackTestStats
    TestAudit03ExtremeFormationInputs stats
    TestAppendLine stats, "TOTAL_AUDIT03_EXTREME_FORMATION: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03ExtremeFormationTests = stats.Report
    Exit Function
Failed:
    RunAudit03ExtremeFormationTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Off/On, три стратегии и шесть направлений моментного вектора.
' После каждого набора тот же calculator решает обычный LC: большой запрос
' не должен оставлять чужую причину отказа или разрушать следующий расчет.
Private Sub TestAudit03ExtremeFormationInputs(ByRef stats As TCrackTestStats)
    Dim section As CSectionModel, setupSolver As CSectionSolver
    Set setupSolver = TestSolveServiceState(section, -20000#, -15000000#, 0#)
    Dim steelParameters As CSteelMaterialParameters, provider As CMaterialModelProvider
    Set steelParameters = New CSteelMaterialParameters
    steelParameters.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#
    Dim extended As Variant, strategy As Variant, magnitude As Variant, direction As Long
    Dim calculator As CCrackFormationCalculator, load As CSectionLoadState
    Dim result As CCrackFormationResult, mx As Double, my As Double, prefix As String
    Dim errorNumber As Long, errorDescription As String
    Dim vector As CLoadPathVector, concrete As CMaterialDiagram, steel As CMaterialDiagram
    Dim probeSolver As CSectionSolver, functionValue As Double, probeState As String, evaluated As Boolean
    For Each extended In Array(False, True)
        Set provider = New CMaterialModelProvider
        provider.TestInitializeFromParameters TestConcreteParameters(), steelParameters, diagramExtensionEnabled:=CBool(extended)
        For Each strategy In Array("UltimateStrain", "LoadMultiplier", "Auto")
            Set calculator = New CCrackFormationCalculator
            calculator.CrackFormationPath = "Auto"
            calculator.CrackFormationSolutionStrategy = CStr(strategy)
            calculator.SolverLoadSteps = 1: calculator.SolverMaxIterations = 3
            For Each magnitude In Array(1E+160, 1E+308)
                For direction = 0 To 5
                    mx = 0#: my = 0#
                    Select Case direction
                        Case 0: mx = CDbl(magnitude)
                        Case 1: mx = -CDbl(magnitude)
                        Case 2: my = CDbl(magnitude)
                        Case 3: my = -CDbl(magnitude)
                        Case 4: mx = CDbl(magnitude): my = -CDbl(magnitude)
                        Case 5: mx = -CDbl(magnitude): my = CDbl(magnitude)
                    End Select
                    prefix = "audit03.extremeFormation." & CStr(extended) & "." & CStr(strategy) & "." & _
                        CStr(magnitude) & "." & CStr(direction)
                    Set load = New CSectionLoadState
                    load.Initialize -20000#, mx, my, 0#, 0#
                    Set result = Nothing
                    On Error Resume Next
                    Err.Clear
                    Set result = calculator.CheckFormation(section, provider, TestCrackedStateSpec(), _
                        TestCrackInitiationSpec(), -20000#, mx, my, load, 0#, 0#)
                    errorNumber = Err.Number: errorDescription = Err.Description
                    On Error GoTo 0
                    TestAssertTrue stats, prefix & ".noRuntimeError", errorNumber = 0
                    TestAssertTrue stats, prefix & ".hasResult", Not result Is Nothing
                    If Not result Is Nothing Then
                        TestAssertTrue stats, prefix & ".notInternalError", result.ResultMeta.InternalStatus <> rsInternalError
                        TestAssertTrue stats, prefix & ".reasonPresent", Len(Trim$(result.ResultMeta.ResultComment)) > 0
                        TestAssertTrue stats, prefix & ".localizedReason", _
                            InStr(1, result.ResultMeta.ResultComment, "Overflow", vbTextCompare) = 0 And _
                            InStr(1, result.ResultMeta.ResultComment, "Object variable", vbTextCompare) = 0
                        If Not result.HasLimitPoint Then
                            TestAssertTrue stats, prefix & ".noPreState", result.PreCrackState Is Nothing
                            TestAssertTrue stats, prefix & ".noPostState", result.PostCrackState Is Nothing
                        End If
                        TestAppendLine stats, "EXTREME_FORMATION_RESULT: " & prefix & "|status=" & _
                            CStr(result.ResultMeta.InternalStatus) & "|code=" & CStr(result.ResultMeta.ResultCode) & _
                            "|comment=" & result.ResultMeta.ResultComment
                    End If
                    TestAppendLine stats, "EXTREME_FORMATION_CALL: " & prefix & "|error=" & CStr(errorNumber) & "|" & errorDescription
                    If CStr(strategy) = "LoadMultiplier" Then
                        Set vector = New CLoadPathVector
                        vector.Initialize -20000#, 0#, 0#, mx, 0#, my
                        Set concrete = provider.ConcreteMaterialForEquilibriumFromSpec(TestCrackInitiationSpec())
                        Set steel = provider.SteelMaterialForEquilibriumFromSpec(TestCrackInitiationSpec())
                        Set probeSolver = Nothing: probeState = vbNullString
                        On Error Resume Next
                        Err.Clear
                        evaluated = calculator.LimitSearchEvaluateCrackFormationLoadMultiplier(section, concrete, steel, _
                            vector, Nothing, 1#, functionValue, probeState, probeSolver)
                        errorNumber = Err.Number: errorDescription = Err.Description
                        On Error GoTo 0
                        TestAssertTrue stats, prefix & ".directProbeNoRuntimeError", errorNumber = 0
                        TestAssertTrue stats, prefix & ".directProbeAttempt", Not probeSolver Is Nothing
                        TestAppendLine stats, "EXTREME_FORMATION_PROBE: " & prefix & "|error=" & CStr(errorNumber) & _
                            "|evaluated=" & CStr(evaluated) & "|" & errorDescription
                    End If
                Next direction
            Next magnitude
            calculator.SolverLoadSteps = 8: calculator.SolverMaxIterations = 100
            load.Initialize -20000#, -15000000#, 0#, 0#, 0#
            Set result = calculator.CheckFormation(section, provider, TestCrackedStateSpec(), _
                TestCrackInitiationSpec(), -20000#, -15000000#, 0#, load, 0#, 0#)
            prefix = "audit03.extremeFormation.recovery." & CStr(extended) & "." & CStr(strategy)
            TestAssertTrue stats, prefix & ".success", result.ResultMeta.InternalStatus = rsSuccess Or _
                result.ResultMeta.InternalStatus = rsSuccessWithWarning
            TestAssertTrue stats, prefix & ".point", result.HasLimitPoint
        Next strategy
    Next extended
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ============================== ДЛЯ ТЕСТОВ: POST БЕЗ РАСТЯЖЕНИЯ АРМАТУРЫ ==============================
' проверяет дополнение цели СП 35 на настоящих Pre/Post/current
' состояниях. Нет растяжения в Post, но текущая арматура растянута: все режимы
' получают psi_s=1, выполненную формулу и предупреждение, а не N/A/NumFail.
Public Function RunNoFormationTensionWidthTests() As String
    On Error GoTo Failed
    Dim stats As TCrackTestStats
    TestAudit03SigmaSCrcBoundary stats, True
    TestAppendLine stats, "TOTAL_NO_FORMATION_TENSION: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunNoFormationTensionWidthTests = stats.Report
    Exit Function
Failed:
    RunNoFormationTensionWidthTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' адресная проверка потребителя геометрии СП 63. Использует
' штатные случаи центрального/косого растяжения и двух режимов площади,
' не запускает полный набор Formation и старую матрицу on/off.
Public Function RunCrackGeometryTests() As String
    On Error GoTo Failed
    Dim stats As TCrackTestStats
    TestEffectiveAndFullTensionZones stats
    TestCentralTensionBranch stats
    TestCrackUserPsiMxy stats
    TestNearUniformCrackTranslations stats
    TestAppendLine stats, "TOTAL_CRACK_GEOMETRY: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunCrackGeometryTests = stats.Report
    Exit Function
Failed:
    RunCrackGeometryTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет почти равномерное растяжение и сжатие при переносах WCS по обеим
' осям. Ширина, статусы и фактические деформации должны сохраняться; epsilon0
' при этом меняется по закону переноса плоскости и может менять знак.
Public Function RunCrackTranslationTests() As String
    On Error GoTo Failed
    Dim stats As TCrackTestStats
    TestNearUniformCrackTranslations stats
    TestAppendLine stats, "TOTAL_CRACK_TRANSLATION: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunCrackTranslationTests = stats.Report
    Exit Function
Failed:
    RunCrackTranslationTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задает физическую плоскость с кривизной 5e-13/мм без поиска равновесия,
' затем подтверждает ее по собственным внутренним усилиям. Независимая ширина
' СП 63: 1.4*0.5*1.2*(20/200000)*400 = 0.0336 мм при psi_s=1.
' СП 35: (20/200000)*1.5*Sqr(75)*10 = 0.012990381... мм по всей площади.
' Для центра 0/±1e-11 плоскость пересекает ноль: перенос не должен менять
' выбор центральной ветки или превращать малую деформацию в ошибку ввода.
Private Sub TestNearUniformCrackTranslations(ByRef stats As TCrackTestStats)
    Dim provider As CMaterialModelProvider, spec As CMaterialModelSpec
    Set provider = TestMaterialProvider(): Set spec = TestCrackedStateSpec()
    Dim owner As CCrackFormationCalculator, meta As CResultMeta, formation As CCrackFormationResult
    Set owner = New CCrackFormationCalculator
    owner.LimitSearchSetFailure sfcNumericalFailure, "Контрольное отсутствие точки образования трещины."
    Set meta = New CResultMeta
    meta.SetSolverFailure sfcNumericalFailure, rkCrackFormation, "Контрольное отсутствие точки образования трещины.", True
    Set formation = New CCrackFormationResult: formation.InitializeFromCalculator owner, meta
    Dim axis As Long, shift As Variant, center As Variant, standard As Variant, x As Variant, y As Variant
    Dim standardIndex As Long, baselineStatus(1 To 2) As EResultInternalStatus, baselineCode(1 To 2) As EResultCode
    Dim baselineCalculated(1 To 2) As Boolean, baselineWidth(1 To 2) As Double
    Dim dx As Double, dy As Double, kx As Double, ky As Double, epsilonCenter As Double
    Dim section As CSectionModel, solver As CSectionSolver, current As CSectionStateResult
    Dim width As CCrackWidthCalculator, load As CSectionLoadState, prefix As String
    For Each center In Array(0.0001, -0.0001, 0#, 0.00000000001, -0.00000000001)
        epsilonCenter = CDbl(center)
        For axis = 1 To 2
            kx = 0#: ky = 0#
            If axis = 1 Then ky = 0.0000000000005 Else kx = 0.0000000000005
            For Each shift In Array(0#, 1000000000#, -1000000000#)
                dx = 0#: dy = 0#
                If axis = 1 Then dx = CDbl(shift) Else dy = CDbl(shift)
                Set section = New CSectionModel
                For Each x In Array(-75#, 75#)
                    For Each y In Array(-50#, 50#)
                        section.AddConcreteElement dx + CDbl(x), dy + CDbl(y), 15000#, width:=150#, height:=100#
                    Next y
                Next x
                For Each x In Array(-90#, 90#)
                    For Each y In Array(-60#, 60#)
                        section.AddRebarElement dx + CDbl(x), dy + CDbl(y), 20#, 0#, "Rebar"
                    Next y
                Next x
                Set solver = New CSectionSolver
                solver.EvaluateStrainPlane section, provider.ConcreteMaterialFromSpec(spec), provider.SteelMaterialFromSpec(spec), _
                    epsilonCenter - ky * dx - kx * dy, kx, ky
                prefix = "crack.translate." & CStr(center) & ".axis" & CStr(axis) & "." & CStr(shift)
                TestAssertTrue stats, prefix & ".equilibrium", solver.ConfirmEquilibrium(solver.Nint, solver.Mxint, solver.Myint)
                Set current = New CSectionStateResult
                current.InitializeFromSolver sstCrackedState, cpCrackedNDS, spec, solver, False, True
                TestAssertTrue stats, prefix & ".physical", current.WithinPhysicalRange
                TestAssertClose stats, prefix & ".maxStrain", current.MaxConcreteStrain, epsilonCenter + ky * 75# + kx * 50#, 0.000000000000001
                TestAssertClose stats, prefix & ".minStrain", current.MinConcreteStrain, epsilonCenter - ky * 75# - kx * 50#, 0.000000000000001
                Set load = New CSectionLoadState
                load.Initialize solver.Nint, solver.Mxint - solver.Nint * dy, solver.Myint - solver.Nint * dx, dx, dy
                standardIndex = 0
                For Each standard In Array("SP63", "SP35")
                    standardIndex = standardIndex + 1
                    prefix = "crack.translate." & CStr(standard) & "." & CStr(center) & ".axis" & CStr(axis) & "." & CStr(shift)
                    Set width = New CCrackWidthCalculator
                    width.StandardCode = CStr(standard)
                    width.PsiMode = "AlwaysCalc": width.AllowableCrackWidth = 0.001
                    width.Calculate current, section, provider, spec, formation, load
                    If Abs(epsilonCenter) < 0.000001 Then
                        If CDbl(shift) = 0# Then ' Начало координат в центре воспроизводит прежний критерий; другие WCS обязаны дать тот же результат.
                            baselineStatus(standardIndex) = width.CrackWidthInternalStatus
                            baselineCode(standardIndex) = width.CrackWidthResultCode
                            baselineCalculated(standardIndex) = width.ResultMeta.Calculated
                            baselineWidth(standardIndex) = width.CrackWidth
                        End If
                        TestAssertTrue stats, prefix & ".noInputOrInternalError", _
                            width.CrackWidthInternalStatus <> rsInvalidInput And width.CrackWidthInternalStatus <> rsInternalError
                        TestAssertTrue stats, prefix & ".centralCriterion", width.CentralTensionBranch = (epsilonCenter > 0.000000000001)
                        TestAssertTrue stats, prefix & ".sameStatus", width.CrackWidthInternalStatus = baselineStatus(standardIndex)
                        TestAssertTrue stats, prefix & ".sameCode", width.CrackWidthResultCode = baselineCode(standardIndex)
                        TestAssertTrue stats, prefix & ".sameCalculated", width.ResultMeta.Calculated = baselineCalculated(standardIndex)
                        TestAssertClose stats, prefix & ".sameWidth", width.CrackWidth, baselineWidth(standardIndex), 0.000000000001
                    ElseIf epsilonCenter > 0# Then
                        TestAssertTrue stats, prefix & ".calculated", width.ResultMeta.Calculated
                        TestAssertTrue stats, prefix & ".checkFailed", width.CrackWidthInternalStatus = rsCheckFailed
                        TestAssertTrue stats, prefix & ".widthExceeded", width.CrackWidthResultCode = rcCrackWidthExceeded
                        If CStr(standard) = "SP63" Then
                            TestAssertClose stats, prefix & ".width", width.CrackWidth, 0.0336, 0.0000001
                        Else
                            TestAssertClose stats, prefix & ".width", width.CrackWidth, 20# / 200000# * 1.5 * Sqr(75#) * 10#, 0.0000001
                        End If
                    Else
                        TestAssertTrue stats, prefix & ".notCalculated", Not width.ResultMeta.Calculated
                        TestAssertTrue stats, prefix & ".notApplicable", width.CrackWidthInternalStatus = rsNotApplicable
                        TestAssertTrue stats, prefix & ".notCracked", width.CrackWidthResultCode = rcCrackNotFormed
                        TestAssertClose stats, prefix & ".width", width.CrackWidth, 0#, 0#
                    End If
                Next standard
            Next shift
        Next axis
    Next center
End Sub
