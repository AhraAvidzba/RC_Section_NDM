Attribute VB_Name = "modTestCrackWidth"
Option Explicit

' ==========================================================================
' Регрессионные тесты расчета нормальных трещин
' ==========================================================================
' Тесты проверяют SLS-расчет CCrackWidthCalculator: выбор расчетной
' растянутой зоны, формулу СП 63 для a_crc, режимы psi_s, центральное
' растяжение и writer основного результата. Проверки не меняют CSectionSolver:
' расчет трещин использует его как готовый общий решатель равновесия.

Private Type TCrackTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

' Выполняет проверки материалов, Formation, подготовки crack data и формулы
' Width. Возвращает отчет assertions; ошибка setup явно остается runtime error.
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
    AppendLine stats, "TOTAL_CRACK: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
    RunCrackWidthTests = stats.Report
    Exit Function

Failed:
    RunCrackWidthTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function


' ------------------------------
' Материал бетона
' ------------------------------
' Проверяем, что растянутую часть диаграммы не обязательно задавать руками:
' Mcrc-диаграмма provider-а включает растянутую ветвь бетона по параметрам II ГПС.
Private Sub TestConcreteTensionBranches(ByRef stats As TCrackTestStats)
    Dim concrete As CMaterialDiagram
    Set concrete = ProvisionalConcrete()
    AssertClose stats, "crack.concrete.ignoreStress", concrete.GetStress(0.0001), 0#, 0.000000000001
    AssertClose stats, "crack.concrete.ignoreTangent", concrete.GetTangentModulus(0.0001), 0#, 0.000000000001

    Set concrete = ProvisionalConcreteWithTension()
    AssertClose stats, "crack.concrete.useStress", concrete.GetStress(0.00002), 0.65, 0.000000000001
    AssertClose stats, "crack.concrete.useLimit", concrete.GetStress(0.001), 1.8, 0.000000000001
    AssertClose stats, "crack.concrete.useTangent", concrete.GetTangentModulus(0.00002), 32500#, 0.000000001
End Sub


' Проверяет формульный API Width: он получает только
' готовые числа и не зависит от State-объектов, статусов и выбора арматуры.
Private Sub TestCrackWidthFormulaCalculatorPure(ByRef stats As TCrackTestStats)
    Dim formula As CCrackWidthCalculator
    Set formula = New CCrackWidthCalculator

    AssertClose stats, "crack.formula.width", _
        formula.CrackWidthFromData(1.4, 0.5, 1#, 0.8, 200#, 200000#, 320#), _
        0.1792, 0.000000000001
    AssertClose stats, "crack.formula.utilization", _
        formula.UtilizationFromData(0.1792, 0.4), 0.448, 0.000000000001
    AssertClose stats, "crack.formula.noSteelEs", _
        formula.CrackWidthFromData(1#, 1#, 1#, 1#, 100#, 0#, 100#), _
        0#, 0.000000000001
End Sub

' ------------------------------
' User-режим psi_s
' ------------------------------
' Для обычного изгиба проверяем, что a_crc собирается из расчетных sigma_s,
' ls и коэффициентов, а не из ручного CrackSpacing.
Private Sub TestCrackUserPsiMx(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceState(section, -20000#, -15000000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, -20000#, -15000000#, 0#, "User", "Effective")
    AssertCrackCommon stats, "crack.user.mx", crack
    AssertClose stats, "crack.user.psi", crack.PsiS, 1#, 0.000000001
    AssertTrue stats, "crack.user.sigmaCrcAvailable", crack.SigmaSCrc > 0#
End Sub

' Для косого изгиба круга проверяет положительные sigma_s и ls и согласованную
' сборку ширины с заданным пользователем psi_s.
Private Sub TestCrackUserPsiMxy(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveCircleServiceState(section, -10000#, -8000000#, -8000000#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, -10000#, -8000000#, -8000000#, "User", "Effective")
    AssertCrackCommon stats, "crack.user.circleMxy", crack
End Sub

' Передает нестандартные пользовательские phi и psi. Они должны попасть
' в итоговую формулу, а найденные Pre/Post и sigma_s,crc остаться доступными.
Private Sub TestCrackUserCoefficients(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceState(section, -20000#, -15000000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, -20000#, -15000000#, 0#, "User", "Effective", _
        allowable:=0.3, phi1Value:=1.6, phi2Value:=0.7, phi3Mode:="User", phi3Value:=1.1, psiSValue:=0.8)
    AssertCrackCommon stats, "crack.user.coeffs", crack
    AssertClose stats, "crack.user.coeffs.phi1", crack.Phi1, 1.6, 0.000000001
    AssertClose stats, "crack.user.coeffs.phi2", crack.Phi2, 0.7, 0.000000001
    AssertClose stats, "crack.user.coeffs.phi3", crack.Phi3, 1.1, 0.000000001
    AssertClose stats, "crack.user.coeffs.psi", crack.PsiS, 0.8, 0.000000001
    AssertTrue stats, "crack.user.coeffs.lambda", crack.FormationResult.LambdaCrc > 0# And crack.FormationResult.LambdaCrc <= 1#
    AssertTrue stats, "crack.user.coeffs.preCrackState", Not crack.FormationResult.PreCrackState Is Nothing
    AssertTrue stats, "crack.user.coeffs.postCrackState", Not crack.FormationResult.PostCrackState Is Nothing
    AssertTrue stats, "crack.user.coeffs.sigmaCrcAvailable", crack.SigmaSCrc > 0#
End Sub

' ------------------------------
' Зона Abt
' ------------------------------
' Проверяет допустимые короткие aliases режимов расстояния до контура:
' nearest и global должны нормализоваться в разные канонические режимы.
Private Sub TestCoverModeNormalization(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceState(section, -20000#, -15000000#, 0#)

    Dim localCrack As CCrackWidthCalculator
    Set localCrack = CalculateCrack(solver, section, -20000#, -15000000#, 0#, _
        "User", "Effective", coverMode:="nearest")
    AssertCrackCommon stats, "crack.cover.local", localCrack
    AssertTrue stats, "crack.cover.local.mode", localCrack.CoverMode = "NearestContour"

    Dim globalCrack As CCrackWidthCalculator
    Set globalCrack = CalculateCrack(solver, section, -20000#, -15000000#, 0#, _
        "User", "Effective", coverMode:="global")
    AssertCrackCommon stats, "crack.cover.global", globalCrack
    AssertTrue stats, "crack.cover.global.mode", globalCrack.CoverMode = "GlobalExtreme"
End Sub

' Для одного косого НДС сравнивает ограниченную полосу Effective и всю
' растянутую зону FullTension; глубина Effective проверяется независимой формулой.
Private Sub TestEffectiveAndFullTensionZones(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceState(section, -20000#, -15000000#, -3000000#)

    Dim effectiveCrack As CCrackWidthCalculator
    Set effectiveCrack = CalculateCrack(solver, section, -20000#, -15000000#, -3000000#, "User", "Effective")
    Dim fullCrack As CCrackWidthCalculator
    Set fullCrack = CalculateCrack(solver, section, -20000#, -15000000#, -3000000#, "User", "FullTension")

    AssertCrackCommon stats, "crack.zone.effective", effectiveCrack
    AssertCrackCommon stats, "crack.zone.full", fullCrack
    AssertTrue stats, "crack.zone.fullAbtPositive", fullCrack.Abt > 0#
    AssertTrue stats, "crack.zone.effectiveAbtPositive", effectiveCrack.Abt > 0#
    AssertClose stats, "crack.zone.effectiveDepthFormula", effectiveCrack.EffectiveZoneDepth, _
        MinTest(MaxTest(effectiveCrack.TensionDepth, 2# * effectiveCrack.CoverA), 0.5 * effectiveCrack.SectionDepthH), 0.000000001
End Sub

' ------------------------------
' Auto-режим psi_s
' ------------------------------
' Если первая проверка Width проходит, Auto сохраняет psi_s = 1.
' Самостоятельный Formation при этом предоставляет свои Pre/Post и sigma_s,crc.
Private Sub TestAutoPsiSkipsLambdaWhenFirstCheckPasses(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceState(section, -20000#, -15000000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, -20000#, -15000000#, 0#, "Auto", "Effective", allowable:=1#)
    AssertCrackCommon stats, "crack.auto.pass", crack
    AssertClose stats, "crack.auto.pass.psi", crack.PsiS, 1#, 0.000000001
    AssertTrue stats, "crack.auto.pass.lambda", crack.FormationResult.LambdaCrc > 0# And crack.FormationResult.LambdaCrc <= 1#
    AssertTrue stats, "crack.auto.pass.sigmaCrcAvailable", crack.SigmaSCrc > 0#
    AssertTrue stats, "crack.auto.pass.preCrackState", Not crack.FormationResult.PreCrackState Is Nothing
    AssertTrue stats, "crack.auto.pass.postCrackState", Not crack.FormationResult.PostCrackState Is Nothing
End Sub

' Сравнивает Auto и AlwaysCalc при прошедшей первой проверке: только AlwaysCalc
' должен уточнить psi_s по sigma_s,crc и уменьшить ширину на том же НДС.
Private Sub TestAlwaysCalcPsiAppliesSigmaCrcWhenAutoPasses(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceState(section, -20000#, -15000000#, 0#)

    Dim autoCrack As CCrackWidthCalculator
    Set autoCrack = CalculateCrack(solver, section, -20000#, -15000000#, 0#, "Auto", "Effective", allowable:=1#)

    Dim alwaysCrack As CCrackWidthCalculator
    Set alwaysCrack = CalculateCrack(solver, section, -20000#, -15000000#, 0#, "AlwaysCalc", "Effective", allowable:=1#)

    AssertCrackCommon stats, "crack.alwaysCalc.pass", alwaysCrack
    AssertTrue stats, "crack.alwaysCalc.sigmaCrc", alwaysCrack.SigmaSCrc > 0#
    AssertTrue stats, "crack.alwaysCalc.psiReduced", alwaysCrack.PsiS < 1#
    AssertTrue stats, "crack.alwaysCalc.widthLessThanAuto", alwaysCrack.CrackWidth < autoCrack.CrackWidth
End Sub

' Режим усреднения относится только к sigma_s,crc. Проверяем, что
' AllSelected проходит нормализацию и расчет как самостоятельная настройка.
Private Sub TestSigmaSCrcAveragingModeAllSelected(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceState(section, -20000#, -15000000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, -20000#, -15000000#, 0#, "AlwaysCalc", "Effective", _
        allowable:=1#, sigmaSCrcAveragingMode:="AllSelected")

    AssertCrackCommon stats, "crack.sigmaCrcMode.allSelected", crack
    AssertTrue stats, "crack.sigmaCrcMode.allSelected.normalized", _
        StrComp(crack.SigmaSCrcAveragingMode, "AllSelected", vbTextCompare) = 0
    AssertTrue stats, "crack.sigmaCrcMode.allSelected.sigmaCrcAvailable", crack.SigmaSCrc > 0#
End Sub

' При не прошедшей первой проверке Auto использует найденный PostCrackState
' для psi_s. Проверяются предел psi_s, роли Pre/Post и растягивающий критерий бетона.
Private Sub TestAutoPsiAndLambdaAfterFailedFirstCheck(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceState(section, -20000#, -15000000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, -20000#, -15000000#, 0#, "Auto", "Effective", allowable:=0.0001)
    AssertCrackCommon stats, "crack.auto.fail", crack
    AssertTrue stats, "crack.auto.fail.lambda", crack.FormationResult.LambdaCrc > 0# And crack.FormationResult.LambdaCrc <= 1#
    AssertTrue stats, "crack.auto.fail.psiRange", crack.PsiS >= 0# And crack.PsiS <= 1#
    AssertTrue stats, "crack.auto.fail.sigmaCrc", crack.SigmaSCrc >= 0#
    AssertTrue stats, "crack.auto.fail.psiReduced", crack.PsiS < 1#
    AssertTrue stats, "crack.auto.fail.preCrackState", Not crack.FormationResult.PreCrackState Is Nothing
    AssertTrue stats, "crack.auto.fail.postCrackState", Not crack.FormationResult.PostCrackState Is Nothing
    AssertTrue stats, "crack.auto.fail.preCrackState.type", _
        crack.FormationResult.PreCrackState.StateTypeText = "PreCrackState"
    AssertTrue stats, "crack.auto.fail.postCrackState.type", _
        crack.FormationResult.PostCrackState.StateTypeText = "PostCrackState"
    AssertTrue stats, "crack.auto.fail.preCrackState.role", _
        crack.FormationResult.PreCrackState.MaterialModelRoleText = "CrackInitiation"
    AssertTrue stats, "crack.auto.fail.postCrackState.role", _
        crack.FormationResult.PostCrackState.MaterialModelRoleText = "CrackedState"

    Dim minConcreteStrain As Double
    Dim maxConcreteStrain As Double
    ConcreteStateStrainBounds crack.FormationResult.PreCrackState, section, minConcreteStrain, maxConcreteStrain
    AssertTrue stats, "crack.auto.fail.preCrack.twoSign", minConcreteStrain < 0#
    AssertClose stats, "crack.auto.fail.preCrack.epsBtUlt", maxConcreteStrain, 0.00015, 0.000001
    AssertTrue stats, "crack.auto.fail.preCrack.notElasticRbtEb", maxConcreteStrain > 1.8 / 32500#
End Sub

' ------------------------------
' Auto-режим Mcrc при чистом изгибе
' ------------------------------
' Проверяет production-сценарий, где исходное CrackedState при N=0 сначала
' находится через CStateSolutionRunner с удобной стартовой плоскостью, а затем
' CCrackFormationCalculator ищет PreCrackState/PostCrackState без отдельного
' пользовательского N и повторной отдельной реализации равновесия.
Private Sub TestAutoMcrcPureBendingConverges(ByRef stats As TCrackTestStats)
    CheckAutoMcrcPureBending stats, "crack.auto.pureMx", 0#, -15000000#, 0#
    CheckAutoMcrcPureBending stats, "crack.auto.pureMy", 0#, 0#, -15000000#
    CheckAutoMcrcPureBending stats, "crack.auto.pureMxy", 0#, -12000000#, -9000000#
End Sub

' Решает заданный чисто изгибный вариант через production-runner и проверяет
' точку Formation, обе стороны трещинообразования и уменьшение psi_s.
Private Sub CheckAutoMcrcPureBending(ByRef stats As TCrackTestStats, ByVal prefix As String, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceStateWithRunner(section, nValue, mxValue, myValue)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, nValue, mxValue, myValue, _
        "Auto", "Effective", allowable:=0.0001)
    AssertCrackCommon stats, prefix, crack
    AssertTrue stats, prefix & ".lambda", crack.FormationResult.LambdaCrc > 0# And crack.FormationResult.LambdaCrc <= 1#
    AssertTrue stats, prefix & ".preCrackState", Not crack.FormationResult.PreCrackState Is Nothing
    AssertTrue stats, prefix & ".postCrackState", Not crack.FormationResult.PostCrackState Is Nothing
    AssertTrue stats, prefix & ".psiReduced", crack.PsiS < 1#

    Dim minConcreteStrain As Double
    Dim maxConcreteStrain As Double
    ConcreteStateStrainBounds crack.FormationResult.PreCrackState, section, minConcreteStrain, maxConcreteStrain
    AssertTrue stats, prefix & ".preCrack.twoSign", minConcreteStrain < 0#
    AssertClose stats, prefix & ".preCrack.epsBtUlt", maxConcreteStrain, 0.00015, 0.000001
End Sub

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

    Set solverLow = SolveServiceStateWithRunner(sectionLow, nValue, mxLow, 0#)
    Set solverHigh = SolveServiceStateWithRunner(sectionHigh, nValue, mxHigh, 0#)

    Dim crackLow As CCrackWidthCalculator
    Dim crackHigh As CCrackWidthCalculator
    Set crackLow = CalculateCrack(solverLow, sectionLow, nValue, mxLow, 0#, "Auto", "Effective", allowable:=0.0001)
    Set crackHigh = CalculateCrack(solverHigh, sectionHigh, nValue, mxHigh, 0#, "Auto", "Effective", allowable:=0.0001)

    AssertCrackCommon stats, "crack.auto.mcrcFixedN.low", crackLow
    AssertCrackCommon stats, "crack.auto.mcrcFixedN.high", crackHigh
    AssertTrue stats, "crack.auto.mcrcFixedN.lowState", Not crackLow.FormationResult.PreCrackState Is Nothing
    AssertTrue stats, "crack.auto.mcrcFixedN.highState", Not crackHigh.FormationResult.PreCrackState Is Nothing
    AssertClose stats, "crack.auto.mcrcFixedN.mcrcInvariant", _
        Abs(crackLow.FormationResult.Mcrc), Abs(crackHigh.FormationResult.Mcrc), 25000#
    AssertClose stats, "crack.auto.mcrcFixedN.lambdaLowMoment", _
        Abs(crackLow.FormationResult.LambdaCrc * mxLow), Abs(crackLow.FormationResult.Mcrc), 25000#
    AssertClose stats, "crack.auto.mcrcFixedN.lambdaHighMoment", _
        Abs(crackHigh.FormationResult.LambdaCrc * mxHigh), Abs(crackHigh.FormationResult.Mcrc), 25000#
End Sub

' Проверяет три пользовательских пути поиска образования нормальной трещины.
' λ*Mxy сохраняет N постоянной, λ*N нужен для центрального
' растяжения, а λ*NMxy масштабирует весь вектор N/Mx/My как единую траекторию.
Private Sub TestCrackInitiationLoadPaths(ByRef stats As TCrackTestStats)
    Dim sectionMxy As CSectionModel
    Dim solverMxy As CSectionSolver
    Set solverMxy = SolveServiceStateWithRunner(sectionMxy, -20000#, -15000000#, 0#)

    Dim crackMxy As CCrackWidthCalculator
    Set crackMxy = CalculateCrack(solverMxy, sectionMxy, -20000#, -15000000#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="lambda*Mxy")
    AssertCrackCommon stats, "crack.path.mxy", crackMxy
    AssertTrue stats, "crack.path.mxy.lambda", crackMxy.FormationResult.LambdaCrc > 0# And crackMxy.FormationResult.LambdaCrc <= 1#
    AssertClose stats, "crack.path.mxy.nFixed", crackMxy.FormationResult.FormationNcrc, -20000#, 0.001

    Dim sectionN As CSectionModel
    Dim solverN As CSectionSolver
    Set solverN = SolveServiceState(sectionN, 200000#, 0#, 0#)

    Dim crackN As CCrackWidthCalculator
    Set crackN = CalculateCrack(solverN, sectionN, 200000#, 0#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="lambda*N")
    AssertCrackCommon stats, "crack.path.n", crackN
    AssertTrue stats, "crack.path.n.central", crackN.CentralTensionBranch
    AssertClose stats, "crack.path.n.formationN", crackN.FormationResult.FormationNcrc, crackN.FormationResult.Ncrc, 0.001

    Dim crackAutoN As CCrackWidthCalculator
    Set crackAutoN = CalculateCrack(solverN, sectionN, 200000#, 0#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="Auto")
    AssertCrackCommon stats, "crack.path.autoN", crackAutoN
    AssertTrue stats, "crack.path.autoN.central", crackAutoN.CentralTensionBranch
    AssertTrue stats, "crack.path.autoN.method", crackAutoN.FormationResult.FormationMethod = ChrW$(&H3BB) & "*N"
    AssertClose stats, "crack.path.autoN.formationN", crackAutoN.FormationResult.FormationNcrc, crackAutoN.FormationResult.Ncrc, 0.001

    Dim sectionNMxy As CSectionModel
    Dim solverNMxy As CSectionSolver
    Set solverNMxy = SolveServiceStateWithRunner(sectionNMxy, -20000#, -15000000#, 0#)

    Dim crackNMxy As CCrackWidthCalculator
    Set crackNMxy = CalculateCrack(solverNMxy, sectionNMxy, -20000#, -15000000#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="lambda*NMxy")
    AssertCrackCommon stats, "crack.path.nmxy", crackNMxy
    AssertTrue stats, "crack.path.nmxy.lambda", crackNMxy.FormationResult.LambdaCrc > 0# And crackNMxy.FormationResult.LambdaCrc <= 1#
    AssertClose stats, "crack.path.nmxy.nScaled", crackNMxy.FormationResult.FormationNcrc, crackNMxy.FormationResult.LambdaCrc * -20000#, 0.001
    AssertClose stats, "crack.path.nmxy.mScaled", crackNMxy.FormationResult.Mcrc, Abs(crackNMxy.FormationResult.LambdaCrc * -15000000#), 50000#

    Dim crackNFixedMomentFallback As CCrackWidthCalculator
    Set crackNFixedMomentFallback = CalculateCrack(solverMxy, sectionMxy, -20000#, -15000000#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="lambda*N")
    AssertTrue stats, "crack.path.nFixedMomentFallback.converged", crackNFixedMomentFallback.Converged
    AssertTrue stats, "crack.path.nFixedMomentFallback.formed", crackNFixedMomentFallback.CrackFormed
    AssertClose stats, "crack.path.nFixedMomentFallback.lambda0", crackNFixedMomentFallback.FormationResult.LambdaCrc, 0#, 0.000000001
    AssertClose stats, "crack.path.nFixedMomentFallback.noFormationN", crackNFixedMomentFallback.FormationResult.FormationNcrc, 0#, 0.000000001
    AssertClose stats, "crack.path.nFixedMomentFallback.noMcrc", crackNFixedMomentFallback.FormationResult.Mcrc, 0#, 0.000000001
    AssertTrue stats, "crack.path.nFixedMomentFallback.noBeforeState", crackNFixedMomentFallback.FormationResult.PreCrackState Is Nothing
    AssertTrue stats, "crack.path.nFixedMomentFallback.noAfterState", crackNFixedMomentFallback.FormationResult.PostCrackState Is Nothing
    AssertClose stats, "crack.path.nFixedMomentFallback.psi1", crackNFixedMomentFallback.PsiS, 1#, 0.000000001
    AssertCrackCalculatorNotNumFail stats, "crack.path.nFixedMomentFallback.status", crackNFixedMomentFallback

    Dim crackFallback As CCrackWidthCalculator
    Set crackFallback = CalculateCrack(solverN, sectionN, 200000#, 0#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="lambda*Mxy")
    AssertTrue stats, "crack.path.mxyAxialFallback.converged", crackFallback.Converged
    AssertTrue stats, "crack.path.mxyAxialFallback.formed", crackFallback.CrackFormed
    AssertClose stats, "crack.path.mxyAxialFallback.lambda0", crackFallback.FormationResult.LambdaCrc, 0#, 0.000000001
    AssertClose stats, "crack.path.mxyAxialFallback.psi1", crackFallback.PsiS, 1#, 0.000000001
    AssertCrackCalculatorNotNumFail stats, "crack.path.mxyAxialFallback.status", crackFallback
End Sub

' Малый момент не достигает критерия до технического MaxLambda. Результат
' обязан сохранить SEARCH_BOUND_REACHED без фиктивного Post и NumFail у Width.
Private Sub TestCrackFormationSearchBoundKeepsTechnicalCode(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceStateWithRunner(section, -20000#, -6000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, -20000#, -6000#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="lambda*Mxy", _
        formationStrategy:="LoadMultiplier")

    AssertTrue stats, "crack.searchBound.notConverged", Not crack.Converged
    AssertTrue stats, "crack.searchBound.notFormed", Not crack.CrackFormed
    AssertTrue stats, "crack.searchBound.internalStatus", _
        crack.FormationResult.ResultMeta.InternalStatus = rsCheckFailed
    AssertTrue stats, "crack.searchBound.resultCode", _
        crack.FormationResult.ResultMeta.ResultCode = rcSearchBoundReached
    AssertTrue stats, "crack.searchBound.noPostState", crack.FormationResult.PostCrackState Is Nothing
    AssertCrackCalculatorNotNumFail stats, "crack.searchBound.widthNotNumFail", crack
End Sub

' Фиксированный малый момент и масштабирование сжимающей N не образуют трещину.
' Проверяются доказанное недостижение критерия, неприменимость Width и отсутствие Pre/Post.
Private Sub TestCrackFormationNoCrackDoesNotBuildPostState(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceStateWithRunner(section, -20000#, -10000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, -20000#, -10000#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="lambda*N")

    AssertTrue stats, "crack.noCrack.converged", crack.Converged
    AssertTrue stats, "crack.noCrack.notFormed", Not crack.CrackFormed
    AssertTrue stats, "crack.noCrack.criterion", _
        crack.FormationResult.ResultMeta.ResultCode = rcCriterionNotReached
    AssertTrue stats, "crack.noCrack.widthNotApplicable", _
        crack.CrackWidthResultCode = rcCrackNotFormed
    AssertTrue stats, "crack.noCrack.noBeforeState", crack.FormationResult.PreCrackState Is Nothing
    AssertTrue stats, "crack.noCrack.noPostState", crack.FormationResult.PostCrackState Is Nothing
    AssertCrackCalculatorNotNumFail stats, "crack.noCrack.notNumFail", crack
End Sub

' Повторный Formation использует repository без нового LastRunner. Width
' все равно получает sigma_s,crc и Pre/Post из сохраненных результатов.
Private Sub TestCrackFormationCacheHitWithoutLastRunner(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceStateWithRunner(section, -20000#, -15000000#, 0#)

    Dim repository As CStateRepository
    Set repository = New CStateRepository

    Dim provider As CStateProvider
    Set provider = New CStateProvider
    provider.Initialize section, TestMaterialProvider(), repository
    ConfigureTestStateProvider provider

    Dim firstCrack As CCrackWidthCalculator
    Set firstCrack = CalculateCrack(solver, section, -20000#, -15000000#, 0#, _
        "AlwaysCalc", "Effective", allowable:=1#, stateProvider:=provider)
    AssertCrackCommon stats, "crack.cache.first", firstCrack
    AssertTrue stats, "crack.cache.first.sigmaCrc", firstCrack.SigmaSCrc > 0#
    AssertTrue stats, "crack.cache.first.preState", Not firstCrack.FormationResult.PreCrackState Is Nothing
    AssertTrue stats, "crack.cache.first.postState", Not firstCrack.FormationResult.PostCrackState Is Nothing

    Dim secondCrack As CCrackWidthCalculator
    Set secondCrack = CalculateCrack(solver, section, -20000#, -15000000#, 0#, _
        "AlwaysCalc", "Effective", allowable:=1#, stateProvider:=provider)
    AssertCrackCommon stats, "crack.cache.second", secondCrack
    AssertTrue stats, "crack.cache.second.reused", provider.LastStateWasReused
    AssertTrue stats, "crack.cache.second.noLastRunner", provider.LastRunner Is Nothing
    AssertTrue stats, "crack.cache.second.sigmaCrc", secondCrack.SigmaSCrc > 0#
    AssertTrue stats, "crack.cache.second.preState", Not secondCrack.FormationResult.PreCrackState Is Nothing
    AssertTrue stats, "crack.cache.second.postState", Not secondCrack.FormationResult.PostCrackState Is Nothing
End Sub

' Проверяет, что общий LimitSearch-result хранит собственный снимок diagnostics.
' Если result будет читать живой Formation-калькулятор, последующие Auto-попытки
' смогут задним числом менять уже сохраненный отчет.
Private Sub TestLimitSearchResultKeepsCrackDiagnosticSnapshot(ByRef stats As TCrackTestStats)
    Dim crack As CCrackFormationCalculator
    Set crack = New CCrackFormationCalculator
    crack.LimitSearchAppendDiagnostic "diagnostic-before"

    Dim result As CLimitSearchResult
    Set result = crack.BuildSearchSnapshot("lambda*Mxy")

    crack.LimitSearchAppendDiagnostic "diagnostic-after"

    AssertTrue stats, "limitSearch.crackSnapshot.hasBefore", _
        InStr(1, result.DiagnosticLog, "diagnostic-before", vbTextCompare) > 0
    AssertTrue stats, "limitSearch.crackSnapshot.noAfter", _
        InStr(1, result.DiagnosticLog, "diagnostic-after", vbTextCompare) = 0
End Sub

' Проверяет стартовую плоскость чистого и косого изгиба и обоих осевых знаков.
' Эти конечные нагрузки должны дать Formation либо обоснованное отсутствие трещины.
Private Sub TestDangerousLoadsDoNotNumFail(ByRef stats As TCrackTestStats)
    CheckDangerousCrackLoad stats, "crack.danger.pureMx", 0#, -15000000#, 0#, True
    CheckDangerousCrackLoad stats, "crack.danger.pureMy", 0#, 0#, -15000000#, True
    CheckDangerousCrackLoad stats, "crack.danger.pureMxy", 0#, -12000000#, -9000000#, True
    CheckDangerousCrackLoad stats, "crack.danger.centralTension", 200000#, 0#, 0#, True, "lambda*N"
    CheckDangerousCrackLoad stats, "crack.danger.pureCompression", -100000#, 0#, 0#, False
End Sub

' Выполняет один вариант через общий runner и проверяет typed Formation-исход.
' Текст комментария не используется как признак отсутствия численной ошибки.
Private Sub CheckDangerousCrackLoad(ByRef stats As TCrackTestStats, ByVal prefix As String, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, _
        ByVal shouldForm As Boolean, Optional ByVal formationPath As String = "lambda*Mxy")
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceStateWithRunner(section, nValue, mxValue, myValue)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, nValue, mxValue, myValue, _
        "Auto", "Effective", allowable:=0.0001, formationPath:=formationPath)
    AssertTrue stats, prefix & ".converged", crack.Converged
    AssertTrue stats, prefix & ".formationNotNumFail", _
        crack.FormationResult.ResultMeta.InternalStatus <> rsNumericalFailure
    AssertCrackCalculatorNotNumFail stats, prefix & ".notNumFail", crack
    If shouldForm Then
        AssertTrue stats, prefix & ".formed", crack.CrackFormed
        AssertTrue stats, prefix & ".lambda", crack.FormationResult.LambdaCrc > 0# And crack.FormationResult.LambdaCrc <= 1#
        AssertTrue stats, prefix & ".preCrackState", Not crack.FormationResult.PreCrackState Is Nothing
        AssertTrue stats, prefix & ".postCrackState", Not crack.FormationResult.PostCrackState Is Nothing
    Else
        AssertTrue stats, prefix & ".notFormed", Not crack.CrackFormed
        AssertClose stats, prefix & ".width", crack.CrackWidth, 0#, 0.000000000001
    End If
End Sub

' Для целиком растянутого бетона проверяет зависимость предельной деформации
' Formation от отношения минимальной и максимальной деформаций по формуле 8.54.
Private Sub TestAutoMcrcOneSignTensionUsesFormula854(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceState(section, 100000#, 3000000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, 100000#, 3000000#, 0#, _
        "Auto", "Effective", allowable:=0.0001)
    AssertCrackCommon stats, "crack.auto.oneSign", crack
    AssertTrue stats, "crack.auto.oneSign.preCrackState", Not crack.FormationResult.PreCrackState Is Nothing

    Dim minConcreteStrain As Double
    Dim maxConcreteStrain As Double
    ConcreteStateStrainBounds crack.FormationResult.PreCrackState, section, minConcreteStrain, maxConcreteStrain
    AssertTrue stats, "crack.auto.oneSign.allTension", minConcreteStrain > 0#

    Dim expectedUlt As Double
    expectedUlt = 0.00015 - (0.00015 - 0.0001) * (minConcreteStrain / maxConcreteStrain)
    AssertClose stats, "crack.auto.oneSign.formula854", maxConcreteStrain, expectedUlt, 0.000001
End Sub

' ------------------------------
' Центральное растяжение
' ------------------------------
' При центральном растяжении по внешним нагрузкам применяется отдельная ветка Ncrc = Ared*Rbt,ser.
Private Sub TestCentralTensionBranch(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceState(section, 200000#, 0#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, 200000#, 0#, 0#, "Auto", "Effective", _
        allowable:=0.0001, formationPath:="lambda*N")
    Dim fullZoneCrack As CCrackWidthCalculator
    Set fullZoneCrack = CalculateCrack(solver, section, 200000#, 0#, 0#, "Auto", "FullTension", _
        allowable:=0.0001, formationPath:="lambda*N")
    AssertCrackCommon stats, "crack.central", crack
    AssertTrue stats, "crack.central.branch", crack.CentralTensionBranch
    AssertTrue stats, "crack.central.ncrc", crack.FormationResult.Ncrc > 0#
    AssertTrue stats, "crack.central.lambda", crack.FormationResult.LambdaCrc > 0# And crack.FormationResult.LambdaCrc <= 1#
    AssertClose stats, "crack.central.lambdaFromNcrc", crack.FormationResult.LambdaCrc * 200000#, crack.FormationResult.Ncrc, 0.001
    AssertTrue stats, "crack.central.preCrackState", Not crack.FormationResult.PreCrackState Is Nothing
    AssertTrue stats, "crack.central.postCrackState", Not crack.FormationResult.PostCrackState Is Nothing
    AssertClose stats, "crack.central.phi3", crack.Phi3, 1.2, 0.000000001
    AssertClose stats, "crack.central.zoneModeInvariant", fullZoneCrack.CrackWidth, crack.CrackWidth, 0.000000001

    Dim shiftedSolver As CSectionSolver
    Set shiftedSolver = SolveServiceState(section, 200000#, 2500000#, -1500000#)
    Dim shiftedMomentCrack As CCrackWidthCalculator
    Set shiftedMomentCrack = CalculateCrack(shiftedSolver, section, 200000#, 2500000#, -1500000#, _
        "User", "Effective")
    AssertTrue stats, "crack.central.eccentricLoadRejected", Not shiftedMomentCrack.CentralTensionBranch

    Dim offsetLoad As CSectionLoadState
    Set offsetLoad = New CSectionLoadState
    offsetLoad.Initialize 200000#, 0#, 0#, -7.5, 12.5
    Dim offsetCrack As CCrackWidthCalculator
    Set offsetCrack = CalculateCrack(shiftedSolver, section, offsetLoad.N, offsetLoad.InternalMx, offsetLoad.InternalMy, _
        "User", "Effective", _
        loadStateForClassification:=offsetLoad, centralReferenceX:=0#, centralReferenceY:=0#)
    AssertTrue stats, "crack.central.offsetLoadStateRejected", Not offsetCrack.CentralTensionBranch
End Sub

' ------------------------------
' Нет растянутой арматуры
' ------------------------------
' Чистое сжатие должно завершаться корректно, но без раскрытой трещины.
Private Sub TestNoTensionRebar(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceState(section, -100000#, 0#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, -100000#, 0#, 0#, "User", "Effective")
    AssertTrue stats, "crack.noTension.converged", crack.Converged
    AssertTrue stats, "crack.noTension.notFormed", Not crack.CrackFormed
    AssertClose stats, "crack.noTension.width", crack.CrackWidth, 0#, 0.000000000001
End Sub

' Создает прямоугольную SLS-fixture и решает текущее НДС с отключенным растянутым
' бетоном. Неподтвержденное равновесие останавливает setup теста, не формулу Width.
Private Function SolveServiceState(ByRef section As CSectionModel, ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double) As CSectionSolver
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 300#, 200#, 0#, 0#, 0#, 0#

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 20#, 20#, 1

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -90#, -60#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B2", 90#, -60#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B3", -90#, 60#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B4", 90#, 60#, 20#, 0#, "Rebar", "", geom

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    ConfigureTestSolver solver
    Set section = BuildGeneratedSectionModel(mesh, rebars)
    Dim provider As CMaterialModelProvider
    Set provider = TestMaterialProvider()
    solver.Solve section, provider.ConcreteMaterial(cpCrackedNDS), provider.SteelMaterial(cpCrackedNDS), nValue, mxValue, myValue
    If Not solver.Converged Then Err.Raise vbObjectError + 3800, "modTestCrackWidth", "Service state did not converge: " & solver.StopReason
    Set SolveServiceState = solver
End Function

' Создает круг с четырьмя стержнями и решает его текущее SLS-НДС.
' Возвращает согласованные section и solver; ошибка сходимости явна в setup.
Private Function SolveCircleServiceState(ByRef section As CSectionModel, ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double) As CSectionSolver
    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByDiameter 300#

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
    Set section = BuildGeneratedSectionModel(mesh, rebars)
    Dim provider As CMaterialModelProvider
    Set provider = TestMaterialProvider()
    solver.Solve section, provider.ConcreteMaterial(cpCrackedNDS), provider.SteelMaterial(cpCrackedNDS), nValue, mxValue, myValue
    If Not solver.Converged Then Err.Raise vbObjectError + 3801, "modTestCrackWidth", "Circle service state did not converge: " & solver.StopReason
    Set SolveCircleServiceState = solver
End Function

' Решает текущее CrackedState тем же production-runner-ом, который batch
' использует перед расчетом трещин. Это важно для чистого изгиба: стартовая
' плоскость подбирается централизованно в CStateGuessBuilder.
Private Function SolveServiceStateWithRunner(ByRef section As CSectionModel, ByVal nValue As Double, _
        ByVal mxValue As Double, ByVal myValue As Double) As CSectionSolver
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 300#, 200#, 0#, 0#, 0#, 0#

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 20#, 20#, 1

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -90#, -60#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B2", 90#, -60#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B3", -90#, 60#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B4", 90#, 60#, 20#, 0#, "Rebar", "", geom

    Set section = BuildGeneratedSectionModel(mesh, rebars)

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
    Set SolveServiceStateWithRunner = runner.ResultSolver
End Function

' Однотипный setup равновесия для fixtures: восемь шагов, 80 итераций,
' компонентные допуски 5 Н и 5000 Н*мм. Units и знаки уже внутренние.
Private Sub ConfigureTestSolver(ByVal solver As CSectionSolver)
    solver.LoadSteps = 8
    solver.MaxIterations = 80
    solver.ToleranceN = 5#
    solver.ToleranceMx = 5000#
    solver.ToleranceMy = 5000#
End Sub

' Передает общему runner тот же фиксированный бюджет и компонентные допуски fixture.
' Это позволяет сравнивать prepared-State маршрут без изменения расчетной постановки.
Private Sub ConfigureTestStateRunner(ByVal runner As CStateSolutionRunner)
    runner.LoadSteps = 8
    runner.MaxIterations = 80
    runner.ToleranceN = 5#
    runner.ToleranceMx = 5000#
    runner.ToleranceMy = 5000#
End Sub

' Настраивает provider для обычного решения и reuse состояний в crack-тестах.
' Фиксированные допуски во внутренних единицах не зависят от Config.
Private Sub ConfigureTestStateProvider(ByVal provider As CStateProvider)
    provider.LoadSteps = 8
    provider.MaxIterations = 80
    provider.ToleranceN = 5#
    provider.ToleranceMx = 5000#
    provider.ToleranceMy = 5000#
End Sub

' Выполняет Formation отдельно от Width и передает последнему готовые состояния.
' Optional provider позволяет проверить reuse; параметры режимов задаются явно,
' чтобы тест не зависел от текущих пользовательских defaults книги.
Private Function CalculateCrack(ByVal solver As CSectionSolver, ByVal section As CSectionModel, _
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
    Dim crack As CCrackWidthCalculator
    Set crack = New CCrackWidthCalculator
    Dim formationCalculator As CCrackFormationCalculator
    Set formationCalculator = New CCrackFormationCalculator
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
    Dim provider As CMaterialModelProvider
    Set provider = TestMaterialProvider()
    Dim crackedSpec As CMaterialModelSpec
    Set crackedSpec = TestCrackedStateSpec()
    Dim initiationSpec As CMaterialModelSpec
    Set initiationSpec = TestCrackInitiationSpec()

    Dim classificationLoad As CSectionLoadState
    If loadStateForClassification Is Nothing Then
        Set classificationLoad = New CSectionLoadState
        classificationLoad.Initialize nValue, mxValue, myValue, 0#, 0#
    Else
        Set classificationLoad = loadStateForClassification
    End If
    Dim formation As CCrackFormationResult
    Set formation = formationCalculator.CheckFormation(section, provider, crackedSpec, initiationSpec, _
        nValue, mxValue, myValue, classificationLoad, centralReferenceX, centralReferenceY)
    Dim currentState As CSectionStateResult
    Set currentState = New CSectionStateResult
    Dim stateRunner As CStateSolutionRunner
    Set stateRunner = New CStateSolutionRunner
    currentState.InitializeFromSolver sstCrackedState, cpCrackedNDS, crackedSpec, solver, _
        stateRunner.StateUsesExtension(section, solver, provider.ConcreteMaterialForEquilibriumFromSpec(crackedSpec), _
            provider.SteelMaterialForEquilibriumFromSpec(crackedSpec)), _
        stateRunner.StateWithinPhysicalRange(section, solver, provider.ConcreteMaterialForEquilibriumFromSpec(crackedSpec), _
            provider.SteelMaterialForEquilibriumFromSpec(crackedSpec))
    crack.Calculate currentState, section, provider, crackedSpec, formation, classificationLoad
    Set CalculateCrack = crack
End Function

' Возвращает SLS spec текущего трещиноватого состояния: растянутый бетон Ignore.
Private Function TestCrackedStateSpec() As CMaterialModelSpec
    Dim spec As CMaterialModelSpec
    Set spec = New CMaterialModelSpec
    spec.Initialize "SLS(II)", "TwoLine", "Ignore", "TwoLine"
    Set TestCrackedStateSpec = spec
End Function

' Возвращает SLS spec Formation с активной трехлинейной растянутой ветвью бетона.
Private Function TestCrackInitiationSpec() As CMaterialModelSpec
    Dim spec As CMaterialModelSpec
    Set spec = New CMaterialModelSpec
    spec.Initialize "SLS(II)", "ThreeLine", "UseDiagram", "TwoLine"
    Set TestCrackInitiationSpec = spec
End Function

' Однотипные материалные fixtures получают готовые физические диаграммы
' соответствующей роли от provider, не конструируя отдельную расчетную методику.
Private Function ProvisionalConcrete() As CMaterialDiagram
    Dim provider As CMaterialModelProvider
    Set provider = TestMaterialProvider()
    Set ProvisionalConcrete = provider.ConcreteMaterial(cpCrackedNDS)
End Function

Private Function ProvisionalConcreteWithTension() As CMaterialDiagram
    Dim provider As CMaterialModelProvider
    Set provider = TestMaterialProvider()
    Set ProvisionalConcreteWithTension = provider.ConcreteMaterial(cpMcrc)
End Function

Private Function ProvisionalSteel() As CMaterialDiagram
    Dim provider As CMaterialModelProvider
    Set provider = TestMaterialProvider()
    Set ProvisionalSteel = provider.SteelMaterial(cpCrackedNDS)
End Function

' Создает воспроизводимые ULS/SLS характеристики бетона и ненапрягаемой арматуры
' во внутренних единицах; пользовательский Config эти unit-fixtures не меняет.
Private Function TestMaterialProvider() As CMaterialModelProvider
    Dim steelParameters As CSteelMaterialParameters
    Set steelParameters = New CSteelMaterialParameters
    steelParameters.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#

    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.InitializeFromParameters TestConcreteParameters(), steelParameters
    Set TestMaterialProvider = provider
End Function

' Задает фиксированные сопротивления и модуль бетонной fixture для независимых
' численных ожиданий; значения не выдаются за нормативную таблицу марки бетона.
Private Function TestConcreteParameters() As CConcreteMaterialParameters
    Dim parameters As CConcreteMaterialParameters
    Set parameters = New CConcreteMaterialParameters
    parameters.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    Set TestConcreteParameters = parameters
End Function

' Проверяет общие инварианты сформированной нормальной трещины и вручную
' собирает ширину из полученных данных. При отказе печатает Formation и Pre-state.
Private Sub AssertCrackCommon(ByRef stats As TCrackTestStats, ByVal prefix As String, ByVal crack As CCrackWidthCalculator)
    If Not crack.Converged Then
        AppendLine stats, "DIAGNOSTIC: " & prefix & "; width=" & crack.StopReason & _
            "; formation=" & crack.FormationResult.StopReason
        If prefix = "crack.user.mx" Then AppendLine stats, crack.FormationResult.DiagnosticLog
        Dim preState As CSectionStateResult
        Set preState = crack.FormationResult.PreCrackState
        If Not preState Is Nothing Then
            AppendLine stats, "PRE_STATE: physical=" & CStr(preState.WithinPhysicalRange) & _
                "; extension=" & CStr(preState.ExtensionUsed) & _
                "; epsBmin=" & FormatNumberInvariant(preState.MinConcreteStrain) & _
                "; epsBmax=" & FormatNumberInvariant(preState.MaxConcreteStrain) & _
                "; epsSmax=" & FormatNumberInvariant(preState.MaxSteelStrain)
        End If
    End If
    AssertTrue stats, prefix & ".converged", crack.Converged
    AssertTrue stats, prefix & ".formed", crack.CrackFormed
    AssertTrue stats, prefix & ".tensionBars", crack.TensionRebarCount > 0
    AssertTrue stats, prefix & ".sigmaPositive", crack.SigmaS > 0#
    AssertTrue stats, prefix & ".spacingPositive", crack.CrackSpacing > 0#
    AssertClose stats, prefix & ".widthFormula", crack.CrackWidth, _
        crack.Phi1 * crack.Phi2 * crack.Phi3 * crack.PsiS * (crack.SigmaS / 200000#) * crack.CrackSpacing, 0.000000001
    AssertClose stats, prefix & ".utilization", crack.Utilization, crack.CrackWidth / crack.AllowableCrackWidth, 0.000000001
End Sub

' Проверяет, что физический fallback образования трещины не превращается
' в пользовательский NumFail. NumFail допустим только при реальной численной
' несходимости state/search, а не при штатном резервном psi_s = 1.
Private Sub AssertCrackCalculatorNotNumFail(ByRef stats As TCrackTestStats, _
        ByVal name As String, ByVal crack As CCrackWidthCalculator)
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    AssertTrue stats, name, policy.ExternalStatus(crack.ResultMeta) <> policy.NumFail
End Sub

' Однотипные assertions увеличивают счетчики и сохраняют конкретный test-ID.
' Численное сравнение использует переданный абсолютный допуск без его изменения.
Private Sub AssertTrue(ByRef stats As TCrackTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name
    End If
End Sub

Private Sub AssertClose(ByRef stats As TCrackTestStats, ByVal name As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal tolerance As Double)
    Dim diff As Double
    diff = Abs(actual - expected)
    If diff <= tolerance Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name & "; actual=" & FormatNumberInvariant(actual) & _
            "; expected=" & FormatNumberInvariant(expected) & "; absDiff=" & FormatNumberInvariant(diff)
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name & "; actual=" & FormatNumberInvariant(actual) & _
            "; expected=" & FormatNumberInvariant(expected) & "; absDiff=" & FormatNumberInvariant(diff)
    End If
End Sub

' Добавляет одну диагностическую строку в возвращаемый отчет теста.
Private Sub AppendLine(ByRef stats As TCrackTestStats, ByVal text As String)
    stats.Report = stats.Report & text & vbCrLf
End Sub

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

Private Function MaxTest(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxTest = a Else MaxTest = b
End Function

Private Function MinTest(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinTest = a Else MinTest = b
End Function

' Независимо вычисляет extrema деформаций бетонных центров по сохраненной
' плоскости epsilon0 + kappaX*y + kappaY*x, не вызывая solver или Formation.
Private Sub ConcreteStateStrainBounds(ByVal state As CSectionStateResult, ByVal section As CSectionModel, _
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

' ============================== ДЛЯ ТЕСТОВ ==============================

' Проверяет самостоятельный formation-сценарий без Width и сохранность
' результата при новом LC на том же калькуляторе, включая сброс результата.
Private Sub TestAudit02IndependentFormation(ByRef stats As TCrackTestStats)
    Dim section As CSectionModel
    Dim solver As CSectionSolver
    Set solver = SolveServiceState(section, -20000#, -15000000#, 0#)
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
    AssertTrue stats, "audit02.formationOnly.success", result.Converged
    AssertTrue stats, "audit02.formationOnly.cracked", result.CrackFormed
    AssertTrue stats, "audit02.formationOnly.hasPoint", result.HasLimitPoint
    AssertTrue stats, "audit02.formationOnly.pre", Not result.PreCrackState Is Nothing
    AssertTrue stats, "audit02.formationOnly.post", Not result.PostCrackState Is Nothing
    If Not result.PreCrackState Is Nothing Then
        AssertTrue stats, "audit02.formationOnly.prePhysical", result.PreCrackState.WithinPhysicalRange
    End If
    If Not result.PostCrackState Is Nothing Then
        AssertTrue stats, "audit02.formationOnly.postPhysical", result.PostCrackState.WithinPhysicalRange
    End If
    AssertTrue stats, "audit02.formationOnly.searchPoint", result.SearchResult.HasLimitPoint
    AssertTrue stats, "audit02.formationOnly.actualMethod", result.SearchResult.ActualMethod <> "Auto"
    Dim lambdaSnapshot As Double
    lambdaSnapshot = result.LambdaCrc
    load.Initialize -100000#, 0#, 0#, 0#, 0#
    Dim second As CCrackFormationResult
    Set second = calculator.CheckFormation(section, provider, TestCrackedStateSpec(), _
        TestCrackInitiationSpec(), load.N, load.InternalMx, load.InternalMy, load, 0#, 0#)
    AssertTrue stats, "audit02.formationOnly.reuseNoCrack", second.Converged And Not second.CrackFormed
    AssertTrue stats, "audit02.formationOnly.reuseNoPoint", Not second.HasLimitPoint
    AssertClose stats, "audit02.formationOnly.snapshot", result.LambdaCrc, lambdaSnapshot, 0#
    result.InitializeFromCalculator Nothing, Nothing
    AssertTrue stats, "audit02.formationOnly.resetPoint", Not result.HasLimitPoint
    AssertTrue stats, "audit02.formationOnly.resetPre", result.PreCrackState Is Nothing
    AssertTrue stats, "audit02.formationOnly.resetSearch", result.SearchResult Is Nothing
    AssertTrue stats, "audit02.formationOnly.resetMeta", result.ResultMeta.InternalStatus = rsInternalError
End Sub

' Различает найденный порог за текущим LC и трещину от постоянной части пути.
' Проверяет сохранение машинного кода/предупреждения и отсутствие фиктивных
' состояний в разрешенном psi=1 fallback, без изменения численных expected.
Private Sub TestAudit02FormationOutcomeSemantics(ByRef stats As TCrackTestStats)
    Dim section As CSectionModel
    Dim solver As CSectionSolver
    Set solver = SolveServiceStateWithRunner(section, -20000#, -1000000#, 0#)
    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, -20000#, -1000000#, 0#, _
        "AlwaysCalc", "Effective", formationPath:="lambda*Mxy")
    AssertTrue stats, "audit02.formation.aboveCurrent.converged", crack.Converged
    AssertTrue stats, "audit02.formation.aboveCurrent.notCracked", Not crack.CrackFormed
    AssertTrue stats, "audit02.formation.aboveCurrent.lambda", crack.FormationResult.LambdaCrc > 1#
    AssertTrue stats, "audit02.formation.aboveCurrent.foundCode", crack.FormationResult.ResultMeta.ResultCode = rcCheckPassed
    Dim search As CLimitSearchResult
    Set search = New CLimitSearchResult
    Set search = crack.FormationResult.SearchResult
    AssertTrue stats, "audit02.formation.aboveCurrent.searchSuccess", search.Meta.InternalStatus = rsSuccess
    AssertClose stats, "audit02.formation.aboveCurrent.pointLambda", search.LambdaUltimate, crack.FormationResult.LambdaCrc, 0#
    AssertClose stats, "audit02.formation.aboveCurrent.pointMoment", _
        Sqr(search.MxUltimate * search.MxUltimate + search.MyUltimate * search.MyUltimate), _
        crack.FormationResult.Mcrc, 0#
    AssertTrue stats, "audit02.formation.aboveCurrent.noPost", crack.FormationResult.PostCrackState Is Nothing

    Set solver = SolveServiceStateWithRunner(section, 200000#, 0#, 0#)
    Set crack = CalculateCrack(solver, section, 200000#, 0#, 0#, _
        "AlwaysCalc", "Effective", formationPath:="lambda*Mxy")
    Set search = crack.FormationResult.SearchResult
    AssertTrue stats, "audit02.formation.constant.warning", search.Meta.InternalStatus = rsSuccessWithWarning
    AssertTrue stats, "audit02.formation.constant.code", search.Meta.ResultCode = rcInitialStateBeyondLimit
    AssertTrue stats, "audit02.formation.constant.comment", Len(search.Meta.ResultComment) > 0
    AssertTrue stats, "audit02.formation.constant.noPre", crack.FormationResult.PreCrackState Is Nothing
    AssertTrue stats, "audit02.formation.constant.noPost", crack.FormationResult.PostCrackState Is Nothing
    AssertClose stats, "audit02.formation.constant.noPoint", search.LambdaUltimate, 0#, 0#
    AssertClose stats, "audit02.formation.constant.psi1", crack.PsiS, 1#, 0#
    search.Initialize Nothing, "Auto", vbNullString, False, 0#, 0#, 0#, 0#, Nothing, vbNullString, vbNullString
    AssertTrue stats, "audit02.formation.missing.internalError", search.Meta.InternalStatus = rsInternalError
End Sub

' ДЛЯ ТЕСТОВ
' Проверяет реальные правила psi для сжатого/нулевого sigma_s,crc и нулевого
' текущего напряжения. Затем три режима проходят полный расчет по разрешенной
' ветви трещины от постоянной части без фиктивного PostCrackState.
Private Sub TestAudit02PsiSignedInputsAndFallbackModes(ByRef stats As TCrackTestStats)
    Dim formulaOwner As CCrackWidthCalculator
    Set formulaOwner = New CCrackWidthCalculator
    AssertClose stats, "audit02.psi.negativeCrc", formulaOwner.Audit02AutoPsiForTests(100#, -25#), 1#, 0#
    AssertClose stats, "audit02.psi.zeroCrc", formulaOwner.Audit02AutoPsiForTests(100#, 0#), 1#, 0#
    AssertClose stats, "audit02.psi.zeroCurrent", formulaOwner.Audit02AutoPsiForTests(0#, 25#), 1#, 0#
    AssertClose stats, "audit02.psi.positiveCrc", formulaOwner.Audit02AutoPsiForTests(100#, 25#), 0.8, 0.000000000001
    AssertClose stats, "audit02.psi.lowerBound", formulaOwner.Audit02AutoPsiForTests(100#, 200#), 0#, 0#
    Dim section As CSectionModel
    Dim solver As CSectionSolver
    Set solver = SolveServiceStateWithRunner(section, 200000#, 0#, 0#)
    Dim mode As Variant
    For Each mode In Array("User", "Auto", "AlwaysCalc")
        Dim crack As CCrackWidthCalculator
        Set crack = CalculateCrack(solver, section, 200000#, 0#, 0#, _
            CStr(mode), "Effective", allowable:=0.0001, psiSValue:=0.65, formationPath:="lambda*Mxy")
        Dim prefix As String
        prefix = "audit02.psi.fallback." & CStr(mode)
        AssertTrue stats, prefix & ".converged", crack.Converged
        AssertTrue stats, prefix & ".warningCode", crack.FormationResult.ResultMeta.ResultCode = rcInitialStateBeyondLimit
        AssertTrue stats, prefix & ".noPost", crack.FormationResult.PostCrackState Is Nothing
        If CStr(mode) = "User" Then
            AssertClose stats, prefix & ".value", crack.PsiS, 0.65, 0#
        Else
            AssertClose stats, prefix & ".value", crack.PsiS, 1#, 0#
        End If
        AssertTrue stats, prefix & ".upperBound", crack.PsiS <= 1#
        AssertCrackCalculatorNotNumFail stats, prefix & ".notNumerical", crack
    Next mode
End Sub

' ============================== ДЛЯ ТЕСТОВ AUDIT03 ==============================
' Отдельный entrypoint проверяет terminal ошибки state-solve для всех физических
' путей трещинообразования, включая отсутствие незаявленных Auto-переходов.
Public Function RunAudit03FormationContracts() As String
    Dim stats As TCrackTestStats
    TestAudit03FormationTypedFailures stats
    AppendLine stats, "TOTAL_AUDIT03_FORMATION_CONTRACTS: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03FormationContracts = stats.Report
End Function

' ДЛЯ ТЕСТОВ: ошибочный метод/бюджет SectionSolver проходят через настоящие
' lambda-пробы; config failure не становится NumFail и не запускает новые пути.
Private Sub TestAudit03FormationTypedFailures(ByRef stats As TCrackTestStats)
    Dim section As CSectionModel, serviceSolver As CSectionSolver
    Set serviceSolver = SolveServiceState(section, -20000#, -15000000#, 0#)
    Dim provider As CMaterialModelProvider
    Set provider = TestMaterialProvider()
    Dim load As CSectionLoadState
    Set load = New CSectionLoadState
    load.Initialize -20000#, -15000000#, 0#, 0#, 0#
    Dim path As Variant, scenario As Long
    For Each path In Array("lambda*Mxy", "lambda*N", "lambda*NMxy", "Auto")
        For scenario = 1 To 3
            Dim calculator As CCrackFormationCalculator
            Set calculator = New CCrackFormationCalculator
            calculator.CrackFormationPath = CStr(path)
            calculator.CrackFormationSolutionStrategy = "LoadMultiplier"
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
                calculator.SolverMinLineSearchAlpha = 0#
            End If
            Dim result As CCrackFormationResult
            Set result = calculator.CheckFormation(section, provider, TestCrackedStateSpec(), _
                TestCrackInitiationSpec(), load.N, load.InternalMx, load.InternalMy, load, 0#, 0#)
            Dim prefix As String
            prefix = "audit03.formation.typed." & CStr(path) & "." & CStr(scenario)
            AssertTrue stats, prefix & ".status", result.ResultMeta.InternalStatus = rsInvalidConfiguration
            AssertTrue stats, prefix & ".code", result.ResultMeta.ResultCode = rcInvalidConfiguration
            AssertTrue stats, prefix & ".notCalculated", Not result.ResultMeta.Calculated
            AssertTrue stats, prefix & ".oneAttempt", result.SolverCallCount = 1
            AssertTrue stats, prefix & ".noPoint", Not result.HasLimitPoint
            AssertTrue stats, prefix & ".noPre", result.PreCrackState Is Nothing
            AssertTrue stats, prefix & ".noPost", result.PostCrackState Is Nothing
            AssertTrue stats, prefix & ".reason", Len(result.ResultMeta.ResultComment) > 0
            AppendLine stats, "COMMENT: " & prefix & "; " & result.ResultMeta.ResultComment
        Next scenario
    Next path
End Sub

' ДЛЯ ТЕСТОВ: ищет воспроизводимые случаи неположительного sigma_s,crc
' на настоящих Formation/PostCrackState, а не на подставленных числах формулы.
' Явные параметры и конечная сетка нагрузок не меняют пользовательский Config.
Public Function RunAudit03SigmaSCrcBoundaryTests() As String
    On Error GoTo Failed
    Dim stats As TCrackTestStats
    TestAudit03SigmaSCrcBoundary stats
    AppendLine stats, "TOTAL_AUDIT03_SIGMA_CRC_BOUNDARY: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03SigmaSCrcBoundaryTests = stats.Report
    Exit Function
Failed:
    RunAudit03SigmaSCrcBoundaryTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' ДЛЯ ТЕСТОВ: сначала находит допустимый CurrentCrackedState общим runner-ом.
' Только затем Formation и Width проверяются в обоих режимах psi/усреднения;
' физически непригодный текущий state не используется как вход формулы.
Private Sub TestAudit03SigmaSCrcBoundary(ByRef stats As TCrackTestStats)
    Dim section As CSectionModel
    Set section = Audit03SigmaBoundarySection()
    Dim provider As CMaterialModelProvider
    Set provider = TestMaterialProvider()
    Dim spec As CMaterialModelSpec
    Set spec = TestCrackedStateSpec()
    Dim nFactor As Variant, eccentricityFactor As Variant, psiMode As Variant, averagingMode As Variant
    Dim nValue As Double, mxValue As Double, cases As Long, nonpositiveCases As Long
    For Each nFactor In Array(0.1, 0.25, 0.4, 0.6, 0.75)
        For Each eccentricityFactor In Array(1.25, 1.75, 2.5, 3.5)
            nValue = -CDbl(nFactor) * 60000# * 22#
            mxValue = -nValue * (200# / 6#) * CDbl(eccentricityFactor)
            Dim runner As CStateSolutionRunner
            Set runner = New CStateSolutionRunner
            ConfigureTestStateRunner runner
            runner.Solve section, provider.ConcreteMaterialForEquilibriumFromSpec(spec), _
                provider.SteelMaterialForEquilibriumFromSpec(spec), nValue, mxValue, 0#, True
            AppendLine stats, "SIGMA_CRC_CURRENT: N=" & FormatNumberInvariant(nValue) & _
                "|Mx=" & FormatNumberInvariant(mxValue) & "|converged=" & CStr(runner.Converged) & _
                "|physical=" & CStr(runner.WithinPhysicalRange) & "|reason=" & runner.StopReason
            If runner.Converged And runner.WithinPhysicalRange Then
                For Each psiMode In Array("AlwaysCalc", "Auto")
                    For Each averagingMode In Array("AllSelected", "TensionOnly")
                        Dim crack As CCrackWidthCalculator
                        Set crack = CalculateCrack(runner.ResultSolver, section, nValue, mxValue, 0#, _
                            CStr(psiMode), "Effective", allowable:=0.00000001, _
                            sigmaSCrcAveragingMode:=CStr(averagingMode))
                        cases = cases + 1
                        Dim prefix As String
                        prefix = "audit03.sigmaCrc.boundary." & CStr(cases)
                        AppendLine stats, "SIGMA_CRC_BOUNDARY: " & prefix & "|N=" & FormatNumberInvariant(nValue) & _
                            "|Mx=" & FormatNumberInvariant(mxValue) & "|psiMode=" & CStr(psiMode) & _
                            "|averaging=" & CStr(averagingMode) & "|point=" & CStr(crack.FormationResult.HasLimitPoint) & _
                            "|widthCalculated=" & CStr(crack.ResultMeta.Calculated) & "|sigmaS=" & FormatNumberInvariant(crack.SigmaS) & _
                            "|sigmaSCrc=" & FormatNumberInvariant(crack.SigmaSCrc) & "|psi=" & FormatNumberInvariant(crack.PsiS) & _
                            "|comment=" & crack.ResultMeta.ResultComment
                        If crack.FormationResult.HasLimitPoint And crack.ResultMeta.Calculated Then
                            AssertTrue stats, prefix & ".psiRange", crack.PsiS >= 0# And crack.PsiS <= 1#
                            AssertTrue stats, prefix & ".widthNotNumerical", crack.ResultMeta.InternalStatus <> rsNumericalFailure
                            If CStr(averagingMode) = "TensionOnly" Then _
                                AssertTrue stats, prefix & ".tensionOnlyNonnegative", crack.SigmaSCrc >= 0#
                            If crack.SigmaSCrc <= 0.000000001 Then
                                nonpositiveCases = nonpositiveCases + 1
                                AssertClose stats, prefix & ".nonpositivePsiOne", crack.PsiS, 1#, 0.000000000001
                                AssertTrue stats, prefix & ".nonpositiveReason", _
                                    InStr(1, crack.ResultMeta.ResultComment, "неполож", vbTextCompare) > 0
                            End If
                        End If
                    Next averagingMode
                Next psiMode
            End If
        Next eccentricityFactor
    Next nFactor
    AssertTrue stats, "audit03.sigmaCrc.boundary.actualPreparedNonpositive", nonpositiveCases > 0
    AppendLine stats, "SIGMA_CRC_BOUNDARY_COUNTS: calculatedVariants=" & CStr(cases) & _
        "|nonpositivePrepared=" & CStr(nonpositiveCases)
End Sub

' ДЛЯ ТЕСТОВ: стержни находятся глубже крайнего бетонного волокна.
' При большом сжатии в момент образования трещины они могут быть сжатыми,
' хотя в более позднем текущем НДС попадут в выбранную растянутую группу.
Private Function Audit03SigmaBoundarySection() As CSectionModel
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 300#, 200#, 0#, 0#, 0#, 0#
    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 20#, 20#, 1
    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -90#, -60#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B2", 90#, -60#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B3", -90#, 60#, 20#, 0#, "Rebar", "", geom
    rebars.AddBar "B4", 90#, 60#, 20#, 0#, "Rebar", "", geom
    Set Audit03SigmaBoundarySection = BuildGeneratedSectionModel(mesh, rebars)
End Function
