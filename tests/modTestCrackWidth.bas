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
    TestCrackFixed1Mx stats
    TestCrackFixed1Mxy stats
    TestEffectiveAndFullTensionZones stats
    TestAutoPsiSkipsLambdaWhenFirstCheckPasses stats
    TestAutoPsiAndLambdaAfterFailedFirstCheck stats
    TestCentralTensionBranch stats
    TestNoTensionRebar stats
    TestCrackWriter stats

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
' Fixed1-режим psi_s
' ------------------------------
' Для обычного изгиба проверяем, что a_crc собирается из расчетных sigma_s,
' ls и коэффициентов, а не из старого ручного CrackSpacing.
Private Sub TestCrackFixed1Mx(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceState(section, -80000#, -5000000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, -80000#, -5000000#, 0#, "Fixed1", "Effective")
    AssertCrackCommon stats, "crack.fixed1.mx", crack
    AssertClose stats, "crack.fixed1.psi", crack.PsiS, 1#, 0.000000001
    AssertClose stats, "crack.fixed1.noSigmaCrc", crack.SigmaSCrc, 0#, 0.000000001
End Sub

Private Sub TestCrackFixed1Mxy(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveCircleServiceState(section, -10000#, -8000000#, -8000000#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, -10000#, -8000000#, -8000000#, "Fixed1", "Effective")
    AssertCrackCommon stats, "crack.fixed1.circleMxy", crack
End Sub

' ------------------------------
' Зона Abt
' ------------------------------
' FullTension не должен давать меньшую площадь бетона, чем Effective, потому
' что Effective является ограниченной полосой у растянутой поверхности.
Private Sub TestEffectiveAndFullTensionZones(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceState(section, -30000#, -5500000#, -1500000#)

    Dim effectiveCrack As CCrackWidthCalculator
    Set effectiveCrack = CalculateCrack(solver, section, -30000#, -5500000#, -1500000#, "Fixed1", "Effective")
    Dim fullCrack As CCrackWidthCalculator
    Set fullCrack = CalculateCrack(solver, section, -30000#, -5500000#, -1500000#, "Fixed1", "FullTension")

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
    Set solver = SolveServiceState(section, -80000#, -5000000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, -80000#, -5000000#, 0#, "Auto", "Effective")
    AssertCrackCommon stats, "crack.auto.pass", crack
    AssertClose stats, "crack.auto.pass.psi", crack.PsiS, 1#, 0.000000001
    AssertClose stats, "crack.auto.pass.noLambda", crack.LambdaCrc, 0#, 0.000000001
    AssertClose stats, "crack.auto.pass.noSigmaCrc", crack.SigmaSCrc, 0#, 0.000000001
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
    Set crack = CalculateCrack(solver, section, 200000#, 0#, 0#, "Auto", "Effective", 0#, 0#, allowable:=0.0001)
    Dim fullZoneCrack As CCrackWidthCalculator
    Set fullZoneCrack = CalculateCrack(solver, section, 200000#, 0#, 0#, "Auto", "FullTension", 0#, 0#, allowable:=0.0001)
    AssertCrackCommon stats, "crack.central", crack
    AssertTrue stats, "crack.central.branch", crack.CentralTensionBranch
    AssertTrue stats, "crack.central.ncrc", crack.Ncrc > 0#
    AssertTrue stats, "crack.central.lambda", crack.LambdaCrc > 0# And crack.LambdaCrc <= 1#
    AssertClose stats, "crack.central.phi3", crack.Phi3, 1.2, 0.000000001
    AssertClose stats, "crack.central.zoneModeInvariant", fullZoneCrack.CrackWidth, crack.CrackWidth, 0.000000001

    Dim shiftedSolver As CSectionSolver
    Set shiftedSolver = SolveServiceState(section, 200000#, 2500000#, -1500000#)
    Dim shiftedMomentCrack As CCrackWidthCalculator
    Set shiftedMomentCrack = CalculateCrack(shiftedSolver, section, 200000#, 2500000#, -1500000#, _
        "Fixed1", "Effective", 2500000#, -1500000#)
    AssertTrue stats, "crack.central.eccentricLoadRejected", Not shiftedMomentCrack.CentralTensionBranch
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
    Set crack = CalculateCrack(solver, section, -100000#, 0#, 0#, "Fixed1", "Effective")
    AssertTrue stats, "crack.noTension.converged", crack.Converged
    AssertTrue stats, "crack.noTension.notFormed", Not crack.CrackFormed
    AssertClose stats, "crack.noTension.width", crack.CrackWidth, 0#, 0.000000000001
End Sub

' ------------------------------
' Writer основного отчета
' ------------------------------
' Проверяем, что расширенный блок результата выводит расчетную ширину,
' допускаемую ширину и диагностические поля новой методики.
Private Sub TestCrackWriter(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim section As CSectionModel
    Set solver = SolveServiceState(section, -80000#, -5000000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, section, -80000#, -5000000#, 0#, "Fixed1", "Effective")
    Dim writer As CCapacityResultWriter
    Set writer = New CCapacityResultWriter
    writer.WriteCrackResult ThisWorkbook, crack

    Dim target As Object
    Set target = ThisWorkbook.Names.Item("rngResultSection").RefersToRange
    AssertClose stats, "crack.writer.width", CDbl(target.Cells.Item(18, 5).Value2), crack.CrackWidth, 0.000000001
    AssertClose stats, "crack.writer.allowable", CDbl(target.Cells.Item(19, 5).Value2), crack.AllowableCrackWidth, 0.000000001
    AssertTrue stats, "crack.writer.sigma", CDbl(target.Cells.Item(21, 5).Value2) > 0#
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
    rebars.AddBar "B1", -90#, -60#, 20#, 0#, "Ribbed", "", geom
    rebars.AddBar "B2", 90#, -60#, 20#, 0#, "Ribbed", "", geom
    rebars.AddBar "B3", -90#, 60#, 20#, 0#, "Ribbed", "", geom
    rebars.AddBar "B4", 90#, 60#, 20#, 0#, "Ribbed", "", geom

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
    rebars.AddBar "B1", 0#, 90#, 20#, 0#, "Ribbed", "", geom
    rebars.AddBar "B2", 90#, 0#, 20#, 0#, "Ribbed", "", geom
    rebars.AddBar "B3", 0#, -90#, 20#, 0#, "Ribbed", "", geom
    rebars.AddBar "B4", -90#, 0#, 20#, 0#, "Ribbed", "", geom

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

Private Sub ConfigureTestSolver(ByVal solver As CSectionSolver)
    solver.LoadSteps = 8
    solver.MaxIterations = 80
    solver.ToleranceN = 5#
    solver.ToleranceMx = 5000#
    solver.ToleranceMy = 5000#
End Sub

Private Function CalculateCrack(ByVal solver As CSectionSolver, ByVal section As CSectionModel, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, _
        ByVal psiMode As String, ByVal zoneMode As String, _
        Optional ByVal centroidMxForCentralCheck As Variant, Optional ByVal centroidMyForCentralCheck As Variant, _
        Optional ByVal allowable As Double = 0.3) As CCrackWidthCalculator
    Dim crack As CCrackWidthCalculator
    Set crack = New CCrackWidthCalculator
    crack.AllowableCrackWidth = allowable
    crack.PsiMode = psiMode
    crack.TensionZoneMode = zoneMode
    crack.SolverLoadSteps = 8
    crack.SolverMaxIterations = 100
    crack.SolverToleranceN = 5#
    crack.SolverToleranceMx = 5000#
    crack.SolverToleranceMy = 5000#
    Dim provider As CMaterialModelProvider
    Set provider = TestMaterialProvider()

    If IsMissing(centroidMxForCentralCheck) Or IsMissing(centroidMyForCentralCheck) Then
        crack.Calculate solver, section, provider, cpCrackedNDS, nValue, mxValue, myValue
    Else
        crack.Calculate solver, section, provider, cpCrackedNDS, nValue, mxValue, myValue, _
            centroidMxForCentralCheck, centroidMyForCentralCheck
    End If
    Set CalculateCrack = crack
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
    steelParameters.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#, "Ribbed"

    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.InitializeFromParameters TestConcreteParameters(), steelParameters
    Set TestMaterialProvider = provider
End Function

Private Function TestConcreteParameters() As CConcreteMaterialParameters
    Dim parameters As CConcreteMaterialParameters
    Set parameters = New CConcreteMaterialParameters
    parameters.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#
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
