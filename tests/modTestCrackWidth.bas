Attribute VB_Name = "modTestCrackWidth"
Option Explicit

' ==========================================================================
' Регрессионные тесты расчета нормальных трещин
' ==========================================================================
' Тесты проверяют новую SLS-методику CCrackWidthCalculator: выбор расчетной
' растянутой зоны, формулу СП 63 для a_crc, режимы psi_s, центральное
' растяжение и writer основного результата. Проверки не меняют CSectionSolver:
' расчет трещин использует его как готовый общий решатель равновесия.

Private Type TCrackTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

Public Function RunCrackWidthTests() As String
    On Error GoTo Failed

    Dim stats As TCrackTestStats
    Dim t0 As Double
    t0 = Timer

    TestConcreteTensionBranches stats
    TestCrackUserPsiMx stats
    TestCrackUserPsiMxy stats
    TestCrackUserCoefficients stats
    TestCoverModeNormalization stats
    TestEffectiveAndFullTensionZones stats
    TestAutoPsiSkipsLambdaWhenFirstCheckPasses stats
    TestAutoPsiAndLambdaAfterFailedFirstCheck stats
    TestAutoMcrcPureBendingConverges stats
    TestAutoMcrcFixedNIndependentOfMomentMagnitude stats
    TestCrackInitiationLoadPaths stats
    TestDangerousLoadsDoNotNumFail stats
    TestAutoMcrcOneSignTensionUsesFormula854 stats
    TestCentralTensionBranch stats
    TestNoTensionRebar stats
    AppendLine stats, "TOTAL_CRACK: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
    RunCrackWidthTests = stats.Report
    Exit Function

Failed:
    RunCrackWidthTests = "RUNTIME ERROR: " & CStr(Err.Number) & _
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
    AssertClose stats, "crack.user.noSigmaCrc", crack.SigmaSCrc, 0#, 0.000000001
End Sub

Private Sub TestCrackUserPsiMxy(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveCircleServiceState(section, -10000#, -8000000#, -8000000#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, -10000#, -8000000#, -8000000#, "User", "Effective")
    AssertCrackCommon stats, "crack.user.circleMxy", crack
End Sub

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
    AssertTrue stats, "crack.user.coeffs.lambda", crack.LambdaCrc > 0# And crack.LambdaCrc <= 1#
    AssertTrue stats, "crack.user.coeffs.beforeMcrcState", Not crack.BeforeMcrcState Is Nothing
    AssertTrue stats, "crack.user.coeffs.afterMcrcState", Not crack.AfterMcrcState Is Nothing
    AssertClose stats, "crack.user.coeffs.noSigmaCrc", crack.SigmaSCrc, 0#, 0.000000001
End Sub

' ------------------------------
' Зона Abt
' ------------------------------
' FullTension не должен давать меньшую площадь бетона, чем Effective, потому
' что Effective является ограниченной полосой у растянутой поверхности.
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
' Режим Auto после непрохождения первой проверки должен найти lambda_crc в пределах текущей нагрузки, решить
' состояние после образования трещины и получить psi_s по sigma_s,crc.
Private Sub TestAutoPsiSkipsLambdaWhenFirstCheckPasses(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceState(section, -20000#, -15000000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, -20000#, -15000000#, 0#, "Auto", "Effective", allowable:=1#)
    AssertCrackCommon stats, "crack.auto.pass", crack
    AssertClose stats, "crack.auto.pass.psi", crack.PsiS, 1#, 0.000000001
    AssertTrue stats, "crack.auto.pass.lambda", crack.LambdaCrc > 0# And crack.LambdaCrc <= 1#
    AssertClose stats, "crack.auto.pass.noSigmaCrc", crack.SigmaSCrc, 0#, 0.000000001
    AssertTrue stats, "crack.auto.pass.beforeMcrcState", Not crack.BeforeMcrcState Is Nothing
    AssertTrue stats, "crack.auto.pass.afterMcrcState", Not crack.AfterMcrcState Is Nothing
End Sub

Private Sub TestAutoPsiAndLambdaAfterFailedFirstCheck(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceState(section, -20000#, -15000000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, -20000#, -15000000#, 0#, "Auto", "Effective", allowable:=0.0001)
    AssertCrackCommon stats, "crack.auto.fail", crack
    AssertTrue stats, "crack.auto.fail.lambda", crack.LambdaCrc > 0# And crack.LambdaCrc <= 1#
    AssertTrue stats, "crack.auto.fail.psiRange", crack.PsiS >= 0# And crack.PsiS <= 1#
    AssertTrue stats, "crack.auto.fail.sigmaCrc", crack.SigmaSCrc >= 0#
    AssertTrue stats, "crack.auto.fail.psiReduced", crack.PsiS < 1#
    AssertTrue stats, "crack.auto.fail.beforeMcrcState", Not crack.BeforeMcrcState Is Nothing
    AssertTrue stats, "crack.auto.fail.afterMcrcState", Not crack.AfterMcrcState Is Nothing
    AssertTrue stats, "crack.auto.fail.beforeMcrcState.type", _
        crack.BeforeMcrcState.StateTypeText = "BeforeMcrcState"
    AssertTrue stats, "crack.auto.fail.afterMcrcState.type", _
        crack.AfterMcrcState.StateTypeText = "AfterMcrcState"
    AssertTrue stats, "crack.auto.fail.beforeMcrcState.role", _
        crack.BeforeMcrcState.MaterialModelRoleText = "CrackInitiation"
    AssertTrue stats, "crack.auto.fail.afterMcrcState.role", _
        crack.AfterMcrcState.MaterialModelRoleText = "CrackedState"

    Dim minConcreteStrain As Double
    Dim maxConcreteStrain As Double
    ConcreteStateStrainBounds crack.BeforeMcrcState, section, minConcreteStrain, maxConcreteStrain
    AssertTrue stats, "crack.auto.fail.beforeMcrc.twoSign", minConcreteStrain < 0#
    AssertClose stats, "crack.auto.fail.beforeMcrc.epsBtUlt", maxConcreteStrain, 0.00015, 0.000001
    AssertTrue stats, "crack.auto.fail.beforeMcrc.notElasticRbtEb", maxConcreteStrain > 1.8 / 32500#
End Sub

' ------------------------------
' Auto-режим Mcrc при чистом изгибе
' ------------------------------
' Проверяет production-сценарий, где исходное CrackedState при N=0 сначала
' находится через CStateSolutionRunner с удобной стартовой плоскостью, а затем
' CCrackWidthCalculator ищет BeforeMcrcState/AfterMcrcState без отдельного
' пользовательского N. Такой случай раньше был численно чувствителен в capacity.
Private Sub TestAutoMcrcPureBendingConverges(ByRef stats As TCrackTestStats)
    CheckAutoMcrcPureBending stats, "crack.auto.pureMx", 0#, -15000000#, 0#
    CheckAutoMcrcPureBending stats, "crack.auto.pureMy", 0#, 0#, -15000000#
    CheckAutoMcrcPureBending stats, "crack.auto.pureMxy", 0#, -12000000#, -9000000#
End Sub

Private Sub CheckAutoMcrcPureBending(ByRef stats As TCrackTestStats, ByVal prefix As String, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceStateWithRunner(section, nValue, mxValue, myValue)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, nValue, mxValue, myValue, _
        "Auto", "Effective", allowable:=0.0001)
    AssertCrackCommon stats, prefix, crack
    AssertTrue stats, prefix & ".lambda", crack.LambdaCrc > 0# And crack.LambdaCrc <= 1#
    AssertTrue stats, prefix & ".beforeMcrcState", Not crack.BeforeMcrcState Is Nothing
    AssertTrue stats, prefix & ".afterMcrcState", Not crack.AfterMcrcState Is Nothing
    AssertTrue stats, prefix & ".psiReduced", crack.PsiS < 1#

    Dim minConcreteStrain As Double
    Dim maxConcreteStrain As Double
    ConcreteStateStrainBounds crack.BeforeMcrcState, section, minConcreteStrain, maxConcreteStrain
    AssertTrue stats, prefix & ".beforeMcrc.twoSign", minConcreteStrain < 0#
    AssertClose stats, prefix & ".beforeMcrc.epsBtUlt", maxConcreteStrain, 0.00015, 0.000001
End Sub

' Проверяет опасные для сходимости сочетания, где нулевая стартовая плоскость
' раньше могла приводить к NumFail: чистый изгиб по каждой оси, косой чистый
' изгиб, центральное растяжение и чистое сжатие без момента. Тест не проверяет
' конкретную ширину, а фиксирует главный контракт: служебные НДС трещин идут
' через общий CStateSolutionRunner и не падают на выборе стартовой плоскости.
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
    AssertTrue stats, "crack.auto.mcrcFixedN.lowState", Not crackLow.BeforeMcrcState Is Nothing
    AssertTrue stats, "crack.auto.mcrcFixedN.highState", Not crackHigh.BeforeMcrcState Is Nothing
    AssertClose stats, "crack.auto.mcrcFixedN.mcrcInvariant", _
        Abs(crackLow.Mcrc), Abs(crackHigh.Mcrc), 25000#
    AssertClose stats, "crack.auto.mcrcFixedN.lambdaLowMoment", _
        Abs(crackLow.LambdaCrc * mxLow), Abs(crackLow.Mcrc), 25000#
    AssertClose stats, "crack.auto.mcrcFixedN.lambdaHighMoment", _
        Abs(crackHigh.LambdaCrc * mxHigh), Abs(crackHigh.Mcrc), 25000#
End Sub

' Проверяет три пользовательских пути поиска образования нормальной трещины.
' λ*Mxy оставляет старую изгибную схему, λ*N нужен для центрального
' растяжения, а λ*NMxy масштабирует весь вектор N/Mx/My как единую траекторию.
Private Sub TestCrackInitiationLoadPaths(ByRef stats As TCrackTestStats)
    Dim sectionMxy As CSectionModel
    Dim solverMxy As CSectionSolver
    Set solverMxy = SolveServiceStateWithRunner(sectionMxy, -20000#, -15000000#, 0#)

    Dim crackMxy As CCrackWidthCalculator
    Set crackMxy = CalculateCrack(solverMxy, sectionMxy, -20000#, -15000000#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="lambda*Mxy")
    AssertCrackCommon stats, "crack.path.mxy", crackMxy
    AssertTrue stats, "crack.path.mxy.lambda", crackMxy.LambdaCrc > 0# And crackMxy.LambdaCrc <= 1#
    AssertClose stats, "crack.path.mxy.nFixed", crackMxy.FormationNcrc, -20000#, 0.001

    Dim sectionN As CSectionModel
    Dim solverN As CSectionSolver
    Set solverN = SolveServiceState(sectionN, 200000#, 0#, 0#)

    Dim crackN As CCrackWidthCalculator
    Set crackN = CalculateCrack(solverN, sectionN, 200000#, 0#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="lambda*N")
    AssertCrackCommon stats, "crack.path.n", crackN
    AssertTrue stats, "crack.path.n.central", crackN.CentralTensionBranch
    AssertClose stats, "crack.path.n.formationN", crackN.FormationNcrc, crackN.Ncrc, 0.001

    Dim crackAutoN As CCrackWidthCalculator
    Set crackAutoN = CalculateCrack(solverN, sectionN, 200000#, 0#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="Auto")
    AssertCrackCommon stats, "crack.path.autoN", crackAutoN
    AssertTrue stats, "crack.path.autoN.central", crackAutoN.CentralTensionBranch
    AssertTrue stats, "crack.path.autoN.method", crackAutoN.CrackFormationMethod = ChrW$(&H3BB) & "*N"
    AssertClose stats, "crack.path.autoN.formationN", crackAutoN.FormationNcrc, crackAutoN.Ncrc, 0.001

    Dim sectionNMxy As CSectionModel
    Dim solverNMxy As CSectionSolver
    Set solverNMxy = SolveServiceStateWithRunner(sectionNMxy, -20000#, -15000000#, 0#)

    Dim crackNMxy As CCrackWidthCalculator
    Set crackNMxy = CalculateCrack(solverNMxy, sectionNMxy, -20000#, -15000000#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="lambda*NMxy")
    AssertCrackCommon stats, "crack.path.nmxy", crackNMxy
    AssertTrue stats, "crack.path.nmxy.lambda", crackNMxy.LambdaCrc > 0# And crackNMxy.LambdaCrc <= 1#
    AssertClose stats, "crack.path.nmxy.nScaled", crackNMxy.FormationNcrc, crackNMxy.LambdaCrc * -20000#, 0.001
    AssertClose stats, "crack.path.nmxy.mScaled", crackNMxy.Mcrc, Abs(crackNMxy.LambdaCrc * -15000000#), 50000#

    Dim crackNFixedMomentFallback As CCrackWidthCalculator
    Set crackNFixedMomentFallback = CalculateCrack(solverMxy, sectionMxy, -20000#, -15000000#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="lambda*N")
    AssertTrue stats, "crack.path.nFixedMomentFallback.converged", crackNFixedMomentFallback.Converged
    AssertTrue stats, "crack.path.nFixedMomentFallback.formed", crackNFixedMomentFallback.CrackFormed
    AssertClose stats, "crack.path.nFixedMomentFallback.lambda0", crackNFixedMomentFallback.LambdaCrc, 0#, 0.000000001
    AssertClose stats, "crack.path.nFixedMomentFallback.noFormationN", crackNFixedMomentFallback.FormationNcrc, 0#, 0.000000001
    AssertClose stats, "crack.path.nFixedMomentFallback.noMcrc", crackNFixedMomentFallback.Mcrc, 0#, 0.000000001
    AssertTrue stats, "crack.path.nFixedMomentFallback.noBeforeState", crackNFixedMomentFallback.BeforeMcrcState Is Nothing
    AssertTrue stats, "crack.path.nFixedMomentFallback.noAfterState", crackNFixedMomentFallback.AfterMcrcState Is Nothing
    AssertClose stats, "crack.path.nFixedMomentFallback.psi1", crackNFixedMomentFallback.PsiS, 1#, 0.000000001

    Dim crackFallback As CCrackWidthCalculator
    Set crackFallback = CalculateCrack(solverN, sectionN, 200000#, 0#, 0#, _
        "Auto", "Effective", allowable:=0.0001, formationPath:="lambda*Mxy")
    AssertTrue stats, "crack.path.mxyAxialFallback.converged", crackFallback.Converged
    AssertTrue stats, "crack.path.mxyAxialFallback.formed", crackFallback.CrackFormed
    AssertClose stats, "crack.path.mxyAxialFallback.lambda0", crackFallback.LambdaCrc, 0#, 0.000000001
    AssertClose stats, "crack.path.mxyAxialFallback.psi1", crackFallback.PsiS, 1#, 0.000000001
End Sub

Private Sub TestDangerousLoadsDoNotNumFail(ByRef stats As TCrackTestStats)
    CheckDangerousCrackLoad stats, "crack.danger.pureMx", 0#, -15000000#, 0#, True
    CheckDangerousCrackLoad stats, "crack.danger.pureMy", 0#, 0#, -15000000#, True
    CheckDangerousCrackLoad stats, "crack.danger.pureMxy", 0#, -12000000#, -9000000#, True
    CheckDangerousCrackLoad stats, "crack.danger.centralTension", 200000#, 0#, 0#, True, "lambda*N"
    CheckDangerousCrackLoad stats, "crack.danger.pureCompression", -100000#, 0#, 0#, False
End Sub

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
    AssertTrue stats, prefix & ".notNumFail", InStr(1, crack.StopReason, "NumericalFailure", vbTextCompare) = 0
    If shouldForm Then
        AssertTrue stats, prefix & ".formed", crack.CrackFormed
        AssertTrue stats, prefix & ".lambda", crack.LambdaCrc > 0# And crack.LambdaCrc <= 1#
        AssertTrue stats, prefix & ".beforeMcrcState", Not crack.BeforeMcrcState Is Nothing
        AssertTrue stats, prefix & ".afterMcrcState", Not crack.AfterMcrcState Is Nothing
    Else
        AssertTrue stats, prefix & ".notFormed", Not crack.CrackFormed
        AssertClose stats, prefix & ".width", crack.CrackWidth, 0#, 0.000000000001
    End If
End Sub

Private Sub TestAutoMcrcOneSignTensionUsesFormula854(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceState(section, 100000#, 3000000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, 100000#, 3000000#, 0#, _
        "Auto", "Effective", allowable:=0.0001)
    AssertCrackCommon stats, "crack.auto.oneSign", crack
    AssertTrue stats, "crack.auto.oneSign.beforeMcrcState", Not crack.BeforeMcrcState Is Nothing

    Dim minConcreteStrain As Double
    Dim maxConcreteStrain As Double
    ConcreteStateStrainBounds crack.BeforeMcrcState, section, minConcreteStrain, maxConcreteStrain
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
    AssertTrue stats, "crack.central.ncrc", crack.Ncrc > 0#
    AssertTrue stats, "crack.central.lambda", crack.LambdaCrc > 0# And crack.LambdaCrc <= 1#
    AssertClose stats, "crack.central.lambdaFromNcrc", crack.LambdaCrc * 200000#, crack.Ncrc, 0.001
    AssertTrue stats, "crack.central.beforeMcrcState", Not crack.BeforeMcrcState Is Nothing
    AssertTrue stats, "crack.central.afterMcrcState", Not crack.AfterMcrcState Is Nothing
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
    runner.Solve section, provider.ConcreteStateMaterialFromSpec(crackedSpec), _
        provider.SteelStateMaterialFromSpec(crackedSpec), nValue, mxValue, myValue, False

    If runner.ResultSolver Is Nothing Or Not runner.Converged Then
        Err.Raise vbObjectError + 3802, "modTestCrackWidth", _
            "Runner service state did not converge: " & runner.StopReason
    End If
    Set SolveServiceStateWithRunner = runner.ResultSolver
End Function

Private Sub ConfigureTestSolver(ByVal solver As CSectionSolver)
    solver.LoadSteps = 8
    solver.MaxIterations = 80
    solver.ToleranceN = 5#
    solver.ToleranceMx = 5000#
    solver.ToleranceMy = 5000#
End Sub

Private Sub ConfigureTestStateRunner(ByVal runner As CStateSolutionRunner)
    runner.LoadSteps = 8
    runner.MaxIterations = 80
    runner.ToleranceN = 5#
    runner.ToleranceMx = 5000#
    runner.ToleranceMy = 5000#
End Sub

Private Function CalculateCrack(ByVal solver As CSectionSolver, ByVal section As CSectionModel, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, _
        ByVal psiMode As String, ByVal zoneMode As String, _
        Optional ByVal allowable As Double = 0.3, _
        Optional ByVal phi1Value As Double = 1.4, Optional ByVal phi2Value As Double = 0.5, _
        Optional ByVal phi3Mode As String = "Auto", Optional ByVal phi3Value As Double = 1#, _
        Optional ByVal psiSValue As Double = 1#, Optional ByVal coverMode As String = "NearestContour", _
        Optional ByVal loadStateForClassification As CSectionLoadState = Nothing, _
        Optional ByVal centralReferenceX As Double = 0#, Optional ByVal centralReferenceY As Double = 0#, _
        Optional ByVal formationPath As String = "lambda*Mxy") As CCrackWidthCalculator
    Dim crack As CCrackWidthCalculator
    Set crack = New CCrackWidthCalculator
    crack.AllowableCrackWidth = allowable
    crack.PsiMode = psiMode
    crack.TensionZoneMode = zoneMode
    crack.Phi1 = phi1Value
    crack.Phi2 = phi2Value
    crack.Phi3Mode = phi3Mode
    crack.Phi3 = phi3Value
    crack.PsiS = psiSValue
    crack.CoverMode = coverMode
    crack.CrackFormationPath = formationPath
    crack.SolverLoadSteps = 8
    crack.SolverMaxIterations = 100
    crack.SolverToleranceN = 5#
    crack.SolverToleranceMx = 5000#
    crack.SolverToleranceMy = 5000#
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
    crack.Calculate solver, section, provider, crackedSpec, initiationSpec, nValue, mxValue, myValue, _
        classificationLoad, centralReferenceX, centralReferenceY
    Set CalculateCrack = crack
End Function

Private Function TestCrackedStateSpec() As CMaterialModelSpec
    Dim spec As CMaterialModelSpec
    Set spec = New CMaterialModelSpec
    spec.Initialize "SLS(II)", "TwoLine", "Ignore", "TwoLine"
    Set TestCrackedStateSpec = spec
End Function

Private Function TestCrackInitiationSpec() As CMaterialModelSpec
    Dim spec As CMaterialModelSpec
    Set spec = New CMaterialModelSpec
    spec.Initialize "SLS(II)", "ThreeLine", "UseDiagram", "TwoLine"
    Set TestCrackInitiationSpec = spec
End Function

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

Private Function TestMaterialProvider() As CMaterialModelProvider
    Dim steelParameters As CSteelMaterialParameters
    Set steelParameters = New CSteelMaterialParameters
    steelParameters.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#

    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.InitializeFromParameters TestConcreteParameters(), steelParameters
    Set TestMaterialProvider = provider
End Function

Private Function TestConcreteParameters() As CConcreteMaterialParameters
    Dim parameters As CConcreteMaterialParameters
    Set parameters = New CConcreteMaterialParameters
    parameters.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    Set TestConcreteParameters = parameters
End Function

Private Sub AssertCrackCommon(ByRef stats As TCrackTestStats, ByVal prefix As String, ByVal crack As CCrackWidthCalculator)
    AssertTrue stats, prefix & ".converged", crack.Converged
    AssertTrue stats, prefix & ".formed", crack.CrackFormed
    AssertTrue stats, prefix & ".tensionBars", crack.TensionRebarCount > 0
    AssertTrue stats, prefix & ".sigmaPositive", crack.SigmaS > 0#
    AssertTrue stats, prefix & ".spacingPositive", crack.CrackSpacing > 0#
    AssertClose stats, prefix & ".widthFormula", crack.CrackWidth, _
        crack.Phi1 * crack.Phi2 * crack.Phi3 * crack.PsiS * (crack.SigmaS / 200000#) * crack.CrackSpacing, 0.000000001
    AssertClose stats, prefix & ".utilization", crack.Utilization, crack.CrackWidth / crack.AllowableCrackWidth, 0.000000001
End Sub

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
