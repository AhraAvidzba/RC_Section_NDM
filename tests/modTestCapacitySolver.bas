Attribute VB_Name = "modTestCapacitySolver"
Option Explicit

' ==========================================================================
' Тесты поиска несущей способности
' ==========================================================================
' Модуль проверяет LoadMultiplier, UltimateStrain и одномерные методы поиска
' lambda. Важна не только близость результата, но и запуск правильной ветки.

Private Type TCapacityTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

' Запускает связанный набор операций и возвращает пользователю итоговый статус выполнения.
Public Function RunCapacitySolverTests() As String
    On Error GoTo Failed

    Dim stats As TCapacityTestStats
    Dim t0 As Double
    t0 = Timer

    AppendLine stats, "RUN: TestMxPositiveAndNegative"
    TestMxPositiveAndNegative stats
    AppendLine stats, "RUN: TestMyPositiveAndNegative"
    TestMyPositiveAndNegative stats
    AppendLine stats, "RUN: TestMxySignedCombinations"
    TestMxySignedCombinations stats
    AppendLine stats, "RUN: TestConcreteTensionBehaviorAffectsSolverAndCapacity"
    TestConcreteTensionBehaviorAffectsSolverAndCapacity stats
    AppendLine stats, "RUN: TestCircleCapacitySymmetry"
    TestCircleCapacitySymmetry stats
    AppendLine stats, "RUN: TestLambdaLessThanOne"
    TestLambdaLessThanOne stats
    AppendLine stats, "RUN: TestLambdaNearOne"
    TestLambdaNearOne stats
    AppendLine stats, "RUN: TestZeroAxialForce"
    TestZeroAxialForce stats
    AppendLine stats, "RUN: TestAxialLoadMultiplierFindsNult"
    TestAxialLoadMultiplierFindsNult stats
    AppendLine stats, "RUN: TestConcreteLimitState"
    TestConcreteLimitState stats
    AppendLine stats, "RUN: TestConcreteTensionLimitState"
    TestConcreteTensionLimitState stats
    AppendLine stats, "RUN: TestConcreteTensionLimitIgnored"
    TestConcreteTensionLimitIgnored stats
    AppendLine stats, "RUN: TestSteelLimitState"
    TestSteelLimitState stats
    AppendLine stats, "RUN: TestLoadMultiplierPureBendingUsesStateGuess"
    TestLoadMultiplierPureBendingUsesStateGuess stats
    AppendLine stats, "RUN: TestAsymmetricCoupledCurvatures"
    TestAsymmetricCoupledCurvatures stats
    AppendLine stats, "RUN: TestAsymmetricMxy"
    TestAsymmetricMxy stats
    AppendLine stats, "RUN: TestInvalidBaseMoment"
    TestInvalidBaseMoment stats
    AppendLine stats, "RUN: TestCapacitySolutionStrategyComparisons"
    TestCapacitySolutionStrategyComparisons stats
    AppendLine stats, "RUN: TestLoadMultiplierWithWorkbookTfDefaults"
    TestLoadMultiplierWithWorkbookTfDefaults stats
    AppendLine stats, "RUN: TestLoadMultiplierSearchMethods"
    TestLoadMultiplierSearchMethods stats
    AppendLine stats, "RUN: TestCapacityLoadPathMethodMatrix"
    TestCapacityLoadPathMethodMatrix stats
    AppendLine stats, "RUN: TestCapacityLoadPathZeroComponentMatrix"
    TestCapacityLoadPathZeroComponentMatrix stats
    AppendLine stats, "RUN: TestLShapeCapacityLoadPathSmoke"
    TestLShapeCapacityLoadPathSmoke stats
    AppendLine stats, "RUN: TestNultBaseLoadStepsSensitivity"
    TestNultBaseLoadStepsSensitivity stats
    AppendLine stats, "RUN: TestSearchMethodInputErrors"
    TestSearchMethodInputErrors stats
    AppendLine stats, "RUN: TestSearchMethodPerformanceComparison"
    TestSearchMethodPerformanceComparison stats

    AppendLine stats, "TOTAL_CAPACITY: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
    RunCapacitySolverTests = stats.Report
    Exit Function

Failed:
    RunCapacitySolverTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestMxPositiveAndNegative(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim capNeg As CCapacitySolver
    Set capNeg = New CCapacitySolver
    ConfigureCapacity capNeg
    capNeg.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, -10000000#, 0#

    Dim capPos As CCapacitySolver
    Set capPos = New CCapacitySolver
    ConfigureCapacity capPos
    capPos.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, 10000000#, 0#

    AssertTrue stats, "capacity.mx.negative.converged", capNeg.Converged
    AssertTrue stats, "capacity.mx.positive.converged", capPos.Converged
    AssertTrue stats, "capacity.mx.negative.sign", capNeg.MxUltimate < 0#
    AssertTrue stats, "capacity.mx.positive.sign", capPos.MxUltimate > 0#
    AssertTrue stats, "capacity.mx.limit.physical", IsPhysicalLimitState(capNeg.LimitState) And IsPhysicalLimitState(capPos.LimitState)
    AssertEquilibrium stats, "capacity.mx.negative", capNeg.LastSolver, -300000#, capNeg.MxUltimate, 0#
    AssertEquilibrium stats, "capacity.mx.positive", capPos.LastSolver, -300000#, capPos.MxUltimate, 0#
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestCircleCapacitySymmetry(ByRef stats As TCapacityTestStats)
    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByDiameter 300#

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
    ConfigureCapacity capMx
    capMx.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -220000#, -6000000#, 0#

    Dim capMy As CCapacitySolver
    Set capMy = New CCapacitySolver
    ConfigureCapacity capMy
    capMy.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -220000#, 0#, -6000000#

    Dim capMxy As CCapacitySolver
    Set capMxy = New CCapacitySolver
    ConfigureCapacity capMxy
    capMxy.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -220000#, -5000000#, -5000000#

    AssertTrue stats, "circle.capacity.mx.converged", capMx.Converged
    AssertTrue stats, "circle.capacity.my.converged", capMy.Converged
    AssertTrue stats, "circle.capacity.mxy.converged", capMxy.Converged
    AssertRelative stats, "circle.capacity.mx.my.symmetry", Abs(capMx.MomentUltimate), Abs(capMy.MomentUltimate), 0.03
    AssertTrue stats, "circle.capacity.mxy.lambda.positive", capMxy.LambdaUltimate > 0#
    AssertClose stats, "circle.capacity.mxy.direction", capMxy.MxUltimate * -5000000# - capMxy.MyUltimate * -5000000#, 0#, 1000#
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestConcreteTensionBehaviorAffectsSolverAndCapacity(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim ignoreConcrete As CMaterialDiagram
    Set ignoreConcrete = ProvisionalConcrete()

    Dim useConcrete As CMaterialDiagram
    Set useConcrete = ProvisionalConcreteWithTension()

    Dim solverIgnore As CSectionSolver
    Set solverIgnore = New CSectionSolver
    ConfigureSectionSolver solverIgnore
    solverIgnore.Solve BuildGeneratedSectionModel(mesh, rebars), ignoreConcrete, ProvisionalSteel(), -50000#, -5000000#, -3000000#

    Dim solverUse As CSectionSolver
    Set solverUse = New CSectionSolver
    ConfigureSectionSolver solverUse
    solverUse.Solve BuildGeneratedSectionModel(mesh, rebars), useConcrete, ProvisionalSteel(), -50000#, -5000000#, -3000000#

    AssertTrue stats, "tensionMode.solver.ignore.converged", solverIgnore.Converged
    AssertTrue stats, "tensionMode.solver.use.converged", solverUse.Converged
    AssertTrue stats, "tensionMode.solver.affectsState", Abs(solverIgnore.KappaX - solverUse.KappaX) > 0.000000001 Or Abs(solverIgnore.KappaY - solverUse.KappaY) > 0.000000001

    AssertTensionBehaviorAffectsSolverMode stats, "tensionBehavior.strength.mx", mesh, rebars, -50000#, -5000000#, 0#
    AssertTensionBehaviorAffectsSolverMode stats, "tensionBehavior.strength.my", mesh, rebars, -50000#, 0#, -5000000#
    AssertTensionBehaviorAffectsSolverMode stats, "tensionBehavior.strength.mxy", mesh, rebars, -50000#, -5000000#, -3000000#
End Sub

Private Sub AssertTensionBehaviorAffectsSolverMode(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double)
    Dim ignoreConcrete As CMaterialDiagram
    Set ignoreConcrete = ProvisionalConcrete()

    Dim useConcrete As CMaterialDiagram
    Set useConcrete = ProvisionalConcreteWithTension()

    Dim solverIgnore As CSectionSolver
    Set solverIgnore = New CSectionSolver
    ConfigureSectionSolver solverIgnore
    solverIgnore.Solve BuildGeneratedSectionModel(mesh, rebars), ignoreConcrete, ProvisionalSteel(), nValue, mxValue, myValue

    Dim solverUse As CSectionSolver
    Set solverUse = New CSectionSolver
    ConfigureSectionSolver solverUse
    solverUse.Solve BuildGeneratedSectionModel(mesh, rebars), useConcrete, ProvisionalSteel(), nValue, mxValue, myValue

    AssertTrue stats, prefix & ".ignore.converged", solverIgnore.Converged
    AssertTrue stats, prefix & ".use.converged", solverUse.Converged
    AssertTrue stats, prefix & ".stateDiffers", Abs(solverIgnore.KappaX - solverUse.KappaX) > 0.000000001 Or _
        Abs(solverIgnore.KappaY - solverUse.KappaY) > 0.000000001 Or Abs(solverIgnore.Epsilon0 - solverUse.Epsilon0) > 0.000000001
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestMxySignedCombinations(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    CheckMxyCombination stats, "capacity.mxy.pp", mesh, rebars, -250000#, 6000000#, 4000000#
    CheckMxyCombination stats, "capacity.mxy.np", mesh, rebars, -250000#, -6000000#, 4000000#
    CheckMxyCombination stats, "capacity.mxy.pn", mesh, rebars, -250000#, 6000000#, -4000000#
    CheckMxyCombination stats, "capacity.mxy.nn", mesh, rebars, -250000#, -6000000#, -4000000#
End Sub

Private Sub CheckMxyCombination(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal nValue As Double, ByVal mxBase As Double, ByVal myBase As Double)
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), nValue, mxBase, myBase

    AssertTrue stats, prefix & ".converged", cap.Converged
    AssertTrue stats, prefix & ".lambda.positive", cap.LambdaUltimate > 0#
    AssertTrue stats, prefix & ".mx.sign", Sgn(cap.MxUltimate) = Sgn(mxBase)
    AssertTrue stats, prefix & ".my.sign", Sgn(cap.MyUltimate) = Sgn(myBase)
    AssertClose stats, prefix & ".direction", cap.MxUltimate * myBase - cap.MyUltimate * mxBase, 0#, 1000#
    AssertEquilibrium stats, prefix, cap.LastSolver, nValue, cap.MxUltimate, cap.MyUltimate
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestMyPositiveAndNegative(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim capNeg As CCapacitySolver
    Set capNeg = New CCapacitySolver
    ConfigureCapacity capNeg
    capNeg.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, 0#, -10000000#

    Dim capPos As CCapacitySolver
    Set capPos = New CCapacitySolver
    ConfigureCapacity capPos
    capPos.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, 0#, 10000000#

    AssertTrue stats, "capacity.my.negative.converged", capNeg.Converged
    AssertTrue stats, "capacity.my.positive.converged", capPos.Converged
    AssertTrue stats, "capacity.my.negative.sign", capNeg.MyUltimate < 0#
    AssertTrue stats, "capacity.my.positive.sign", capPos.MyUltimate > 0#
    AssertTrue stats, "capacity.my.limit.physical", IsPhysicalLimitState(capNeg.LimitState) And IsPhysicalLimitState(capPos.LimitState)
    AssertEquilibrium stats, "capacity.my.negative", capNeg.LastSolver, -300000#, 0#, capNeg.MyUltimate
    AssertEquilibrium stats, "capacity.my.positive", capPos.LastSolver, -300000#, 0#, capPos.MyUltimate
End Sub
' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestLambdaLessThanOne(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim reference As CCapacitySolver
    Set reference = New CCapacitySolver
    ConfigureCapacity reference
    reference.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, -10000000#, 0#

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, reference.MxUltimate * 2#, 0#

    AssertTrue stats, "capacity.lambda.lessThanOne.converged", cap.Converged
    AssertTrue stats, "capacity.lambda.lessThanOne.value", cap.LambdaUltimate < 1#
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestLambdaNearOne(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim reference As CCapacitySolver
    Set reference = New CCapacitySolver
    ConfigureCapacity reference
    reference.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, -10000000#, 0#

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, reference.MxUltimate, 0#

    AssertTrue stats, "capacity.lambda.nearOne.converged", cap.Converged
    AssertClose stats, "capacity.lambda.nearOne.value", cap.LambdaUltimate, 1#, 0.03
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestZeroAxialForce(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.ConcreteCompressionLimit = -0.0006
    cap.SteelStrainLimit = 0.001#
    cap.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), LinearConcrete(), LinearSteel(), 0#, -5000000#, 0#

    AssertTrue stats, "capacity.zeroN.converged", cap.Converged
    AssertEquilibrium stats, "capacity.zeroN", cap.LastSolver, 0#, cap.MxUltimate, 0#
End Sub

' Проверяет универсальную lambda-траекторию для чистой продольной силы. Здесь
' масштабируется только N, поэтому нулевые Mx/My являются нормальным входом,
' а не ошибкой контракта.
Private Sub TestAxialLoadMultiplierFindsNult(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, rebars)

    Dim compression As CCapacitySolver
    Set compression = New CCapacitySolver
    ConfigureCapacity compression
    compression.SolveByLoadPathMultiplier section, ProvisionalConcrete(), ProvisionalSteel(), _
        0#, -100000#, 0#, 0#, 0#, 0#

    Dim tension As CCapacitySolver
    Set tension = New CCapacitySolver
    ConfigureCapacity tension
    tension.SolveByLoadPathMultiplier section, ProvisionalConcrete(), ProvisionalSteel(), _
        0#, 50000#, 0#, 0#, 0#, 0#

    AssertTrue stats, "capacity.nult.compression.converged", compression.Converged
    AssertTrue stats, "capacity.nult.compression.sign", compression.NUltimate < 0#
    AssertTrue stats, "capacity.nult.compression.safety", Abs(compression.NUltimate / -100000#) > 1#
    AssertEquilibrium stats, "capacity.nult.compression", compression.LastSolver, compression.NUltimate, 0#, 0#

    AssertTrue stats, "capacity.nult.tension.converged", tension.Converged
    AssertTrue stats, "capacity.nult.tension.sign", tension.NUltimate > 0#
    AssertTrue stats, "capacity.nult.tension.safety", Abs(tension.NUltimate / 50000#) > 1#
    AssertEquilibrium stats, "capacity.nult.tension", tension.LastSolver, tension.NUltimate, 0#, 0#
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestConcreteLimitState(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.ConcreteCompressionLimit = -0.0006
    cap.SteelStrainLimit = 1#
    cap.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), LinearConcrete(), LinearSteel(), 0#, -10000000#, 0#

    AssertTrue stats, "capacity.concreteLimit.converged", cap.Converged
    AssertEquals stats, "capacity.concreteLimit.state", cap.LimitState, "ConcreteStrainLimit"
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestConcreteTensionLimitState(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim concrete As CMaterialDiagram
    Set concrete = ProvisionalConcreteWithTension()

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.InitialLambdaStep = 0.1
    cap.ConcreteCompressionLimit = -1#
    cap.ConcreteTensionLimit = 0.00000001
    cap.ConcreteTensionLimitEnabled = True
    cap.SteelStrainLimit = 1#
    cap.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), concrete, ProvisionalSteel(), -50000#, -5000000#, 0#

    AssertTrue stats, "capacity.concreteTensionLimit.converged", cap.Converged
    AssertEquals stats, "capacity.concreteTensionLimit.state", cap.LimitState, "ConcreteTensionStrainLimit"
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestConcreteTensionLimitIgnored(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim concrete As CMaterialDiagram
    Set concrete = ProvisionalConcrete()

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.InitialLambdaStep = 0.1
    cap.ConcreteCompressionLimit = -1#
    cap.ConcreteTensionLimit = 0.00000001
    cap.ConcreteTensionLimitEnabled = False
    cap.SteelStrainLimit = 0.00000001
    cap.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), concrete, ProvisionalSteel(), -50000#, -5000000#, 0#

    AssertTrue stats, "capacity.concreteTensionLimit.ignore.notTension", cap.LimitState <> "ConcreteTensionStrainLimit"
    AssertEquals stats, "capacity.concreteTensionLimit.ignore.state", cap.LimitState, "SteelStrainLimit"
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestSteelLimitState(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.ConcreteCompressionLimit = -1#
    cap.SteelStrainLimit = 0.0005
    cap.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), LinearConcrete(), LinearSteel(), 0#, -10000000#, 0#

    AssertTrue stats, "capacity.steelLimit.converged", cap.Converged
    AssertEquals stats, "capacity.steelLimit.state", cap.LimitState, "SteelStrainLimit"
End Sub

' Проверяет, что LoadMultiplier не падает на чистом изгибе из-за старта из нулевого излома диаграммы.
' Этот искусственно жесткий сценарий проверяет устойчивость numerical extension после появления
' CStateGuessBuilder он стал важной регрессией устойчивости для lambda*Mx/lambda*My без постоянной продольной силы.
Private Sub TestLoadMultiplierPureBendingUsesStateGuess(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SolverMaxIterations = 1
    cap.MaxRetries = 1
    cap.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), 0#, -10000000#, 0#

    AssertTrue stats, "capacity.pureBendingGuess.converged", cap.Converged
    AssertTrue stats, "capacity.pureBendingGuess.physical", IsPhysicalLimitState(cap.LimitState)
    AssertTrue stats, "capacity.pureBendingGuess.usedGuess", InStr(1, cap.DiagnosticLog, "pureBendingProbeGuess applied", vbTextCompare) > 0
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestAsymmetricCoupledCurvatures(ByRef stats As TCapacityTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 320#, 220#, 0#, 70#, 20#, 0#

    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 10#)

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete BuildGeneratedSectionModel(mesh, Nothing)

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -110#, -70#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B2", 95#, -65#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B3", -90#, 65#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B4", 80#, 70#, 20#, 0#, "A400", "", geom

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -260000#, -8000000#, 0#

    AssertTrue stats, "capacity.asym.Ixy.nonzero", Abs(props.Ixyc) > 1000000#
    AssertTrue stats, "capacity.asym.converged", cap.Converged
    AssertTrue stats, "capacity.asym.coupledKappaY", Abs(cap.LastSolver.KappaY) > 0.000000001
    AssertEquilibrium stats, "capacity.asym", cap.LastSolver, -260000#, cap.MxUltimate, 0#
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestAsymmetricMxy(ByRef stats As TCapacityTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 320#, 220#, 0#, 70#, 20#, 0#

    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 10#)

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete BuildGeneratedSectionModel(mesh, Nothing)

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -110#, -70#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B2", 95#, -65#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B3", -90#, 65#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B4", 80#, 70#, 20#, 0#, "A400", "", geom

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -260000#, -6000000#, -4000000#

    AssertTrue stats, "capacity.mxy.asym.Ixy.nonzero", Abs(props.Ixyc) > 1000000#
    AssertTrue stats, "capacity.mxy.asym.converged", cap.Converged
    AssertTrue stats, "capacity.mxy.asym.kappaX.nonzero", Abs(cap.LastSolver.KappaX) > 0.000000001
    AssertTrue stats, "capacity.mxy.asym.kappaY.nonzero", Abs(cap.LastSolver.KappaY) > 0.000000001
    AssertClose stats, "capacity.mxy.asym.direction", cap.MxUltimate * -4000000# - cap.MyUltimate * -6000000#, 0#, 1000#
    AssertEquilibrium stats, "capacity.mxy.asym", cap.LastSolver, -260000#, cap.MxUltimate, cap.MyUltimate
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestInvalidBaseMoment(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 200#, 100#, 20#, 60#, 30#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -100000#, 0#, 0#

    AssertTrue stats, "capacity.invalidBaseMoment.notConverged", Not cap.Converged
    AssertEquals stats, "capacity.invalidBaseMoment.state", cap.LimitState, "InvalidInput"
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestCapacitySolutionStrategyComparisons(ByRef stats As TCapacityTestStats)
    TestMethodPureCompression stats
    CompareCapacitySolutionStrategys stats, "method.n_plus_mx", -300000#, -10000000#, 0#, False
    CompareCapacitySolutionStrategys stats, "method.n_plus_my", -300000#, 0#, -10000000#, False
    CompareCapacitySolutionStrategys stats, "method.biaxial", -250000#, -6000000#, -4000000#, False
    CompareCircleCapacitySolutionStrategys stats
    TestMethodStrainLimitState stats
    TestUltimateStrainNumericalFailure stats
    TestUltimateStrainInvalidLambdaClearsMoments stats
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestMethodPureCompression(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim loadMethod As CCapacitySolver
    Set loadMethod = New CCapacitySolver
    ConfigureCapacity loadMethod
    loadMethod.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, 0#, 0#

    Dim strainMethod As CCapacitySolver
    Set strainMethod = New CCapacitySolver
    ConfigureCapacity strainMethod
    strainMethod.SolveByUltimateStrain BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, 0#, 0#

    AssertEquals stats, "method.pure_compression.load.invalid", loadMethod.LimitState, "InvalidInput"
    AssertEquals stats, "method.pure_compression.strain.invalid", strainMethod.LimitState, "InvalidInput"
    AppendComparison stats, "method.pure_compression.load", loadMethod, 0#
    AppendComparison stats, "method.pure_compression.strain", strainMethod, 0#
End Sub

Private Sub CompareCapacitySolutionStrategys(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal nValue As Double, ByVal mxBase As Double, ByVal myBase As Double, ByVal tightLimits As Boolean)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim loadMethod As CCapacitySolver
    Dim strainMethod As CCapacitySolver
    Dim t0 As Double
    Dim loadElapsed As Double
    Dim strainElapsed As Double

    Set loadMethod = New CCapacitySolver
    ConfigureCapacity loadMethod
    If tightLimits Then
        loadMethod.ConcreteCompressionLimit = -0.0006
        loadMethod.SteelStrainLimit = 0.001#
    End If
    t0 = Timer
    loadMethod.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), nValue, mxBase, myBase
    loadElapsed = Timer - t0

    Set strainMethod = New CCapacitySolver
    ConfigureCapacity strainMethod
    If tightLimits Then
        strainMethod.ConcreteCompressionLimit = -0.0006
        strainMethod.SteelStrainLimit = 0.001#
    End If
    t0 = Timer
    strainMethod.SolveByUltimateStrain BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), nValue, mxBase, myBase
    strainElapsed = Timer - t0

    AssertTrue stats, prefix & ".load.converged", loadMethod.Converged
    AssertTrue stats, prefix & ".strain.converged", strainMethod.Converged
    AssertClose stats, prefix & ".lambda", strainMethod.LambdaUltimate, loadMethod.LambdaUltimate, 0.02
    AssertEquals stats, prefix & ".limitState", strainMethod.LimitState, loadMethod.LimitState
    AssertTrue stats, prefix & ".criticalElement", Len(strainMethod.CriticalElement) > 0
    AssertTrue stats, prefix & ".solverCalls", strainMethod.Iterations > 0
    AppendComparison stats, prefix & ".load", loadMethod, loadElapsed
    AppendComparison stats, prefix & ".strain", strainMethod, strainElapsed
End Sub

Private Sub CompareCircleCapacitySolutionStrategys(ByRef stats As TCapacityTestStats)
    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByDiameter 300#

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
    ConfigureCapacity loadMethod
    loadMethod.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -220000#, -6000000#, 0#

    Dim strainMethod As CCapacitySolver
    Set strainMethod = New CCapacitySolver
    ConfigureCapacity strainMethod
    strainMethod.SolveByUltimateStrain BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -220000#, 0#, -6000000#

    AssertTrue stats, "method.circle.load.converged", loadMethod.Converged
    AssertTrue stats, "method.circle.strain.converged", strainMethod.Converged
    AssertRelative stats, "method.circle.symmetry", Abs(loadMethod.MomentUltimate), Abs(strainMethod.MomentUltimate), 0.04
    AssertTrue stats, "method.circle.criticalElement", Len(strainMethod.CriticalElement) > 0
    AppendComparison stats, "method.circle.load", loadMethod, 0#
    AppendComparison stats, "method.circle.strain", strainMethod, 0#
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestMethodStrainLimitState(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim loadMethod As CCapacitySolver
    Set loadMethod = New CCapacitySolver
    ConfigureCapacity loadMethod
    loadMethod.ConcreteCompressionLimit = -0.0006
    loadMethod.SteelStrainLimit = 1#
    loadMethod.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), LinearConcrete(), LinearSteel(), 0#, -10000000#, 0#

    Dim strainMethod As CCapacitySolver
    Set strainMethod = New CCapacitySolver
    ConfigureCapacity strainMethod
    strainMethod.ConcreteCompressionLimit = -0.0006
    strainMethod.SteelStrainLimit = 1#
    strainMethod.SolveByUltimateStrain BuildGeneratedSectionModel(mesh, rebars), LinearConcrete(), LinearSteel(), 0#, -10000000#, 0#

    AssertTrue stats, "method.strain_limit.load.converged", loadMethod.Converged
    AssertTrue stats, "method.strain_limit.strain.converged", strainMethod.Converged
    AssertClose stats, "method.strain_limit.lambda", strainMethod.LambdaUltimate, loadMethod.LambdaUltimate, 0.02
    AssertEquals stats, "method.strain_limit.limitState", strainMethod.LimitState, "ConcreteStrainLimit"
    AssertTrue stats, "method.strain_limit.criticalElement", Len(strainMethod.CriticalElement) > 0
    AssertTrue stats, "method.strain_limit.criticalUtilization", strainMethod.CriticalStrainUtilization > 0.98
    AppendComparison stats, "method.strain_limit.load", loadMethod, 0#
    AppendComparison stats, "method.strain_limit.strain", strainMethod, 0#
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestUltimateStrainNumericalFailure(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SolverMaxIterations = 1
    cap.MaxRetries = 1
    cap.SolveByUltimateStrain BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), 0#, -10000000#, 0#

    AssertTrue stats, "method.numericalFailure.notConverged", Not cap.Converged
    AssertTrue stats, "method.numericalFailure.state", cap.LimitState = "NumericalFailure" Or cap.LimitState = "SingularTangent"
    AssertTrue stats, "method.numericalFailure.noPhysicalUpper", Not IsPhysicalLimitState(cap.LimitState)
    AppendComparison stats, "method.numericalFailure.strain", cap, 0#
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestUltimateStrainInvalidLambdaClearsMoments(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.MaxLambda = 0.5
    cap.SolveByUltimateStrain BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, -1000000#, 0#

    AssertTrue stats, "method.invalidLambda.notConverged", Not cap.Converged
    AssertEquals stats, "method.invalidLambda.state", cap.LimitState, "NumericalFailure"
    AssertClose stats, "method.invalidLambda.lambdaZero", cap.LambdaUltimate, 0#, 0#
    AssertClose stats, "method.invalidLambda.mxZero", cap.MxUltimate, 0#, 0#
    AssertClose stats, "method.invalidLambda.momentZero", cap.MomentUltimate, 0#, 0#
End Sub

' Проверяет пользовательский сценарий с дефолтными единицами книги.
' После перехода INPUT-нагрузок на tf и tf*m численные значения допусков
' solver-а тоже должны быть заданы в этих единицах. Иначе 1000 в строке
' Solver.ToleranceMx превращается в 1000 tf*m, LoadMultiplier принимает
' грубо несбалансированные probe-точки и уходит далеко за реальный предел.
Private Sub TestLoadMultiplierWithWorkbookTfDefaults(ByRef stats As TCapacityTestStats)
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    AssertEquals stats, "capacity.tfDefaults.geometry", settings.GetString("Geometry.Type", ""), "LShape"
    AssertEquals stats, "capacity.tfDefaults.forceUnit", settings.GetString("Units.Force.Input", ""), "tf"
    AssertEquals stats, "capacity.tfDefaults.momentUnit", settings.GetString("Units.Moment.Input", ""), "tf*m"

    Dim registry As CSectionTypeRegistry
    Set registry = New CSectionTypeRegistry

    Dim section As CSectionModel
    Set section = registry.BuildGeneratedModel(settings, units)

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
    mxOffset = nValue * props.CentroidY
    myOffset = nValue * props.CentroidX

    Dim reference As CCapacitySolver
    Set reference = New CCapacitySolver
    reference.ApplySettings settings, units
    reference.SolveByUltimateStrain section, concrete, steel, nValue, mxValue, myValue, mxOffset, myOffset
    AssertTrue stats, "capacity.tfDefaults.ultimate.converged", reference.Converged
    AssertTrue stats, "capacity.tfDefaults.ultimate.lambda", reference.LambdaUltimate > 0# And reference.LambdaUltimate < 10#

    AssertLoadMultiplierTfDefault stats, "Bisection", settings, units, section, concrete, steel, _
        nValue, mxValue, myValue, mxOffset, myOffset, reference.LambdaUltimate
    AssertLoadMultiplierTfDefault stats, "Brent", settings, units, section, concrete, steel, _
        nValue, mxValue, myValue, mxOffset, myOffset, reference.LambdaUltimate
    AssertLoadMultiplierTfDefault stats, "Secant", settings, units, section, concrete, steel, _
        nValue, mxValue, myValue, mxOffset, myOffset, reference.LambdaUltimate
End Sub

Private Sub AssertLoadMultiplierTfDefault(ByRef stats As TCapacityTestStats, ByVal methodName As String, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByVal section As CSectionModel, ByVal concrete As Object, ByVal steel As Object, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double, _
        ByVal mxOffset As Double, ByVal myOffset As Double, ByVal referenceLambda As Double)
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    cap.ApplySettings settings, units
    cap.SearchMethod = methodName
    cap.SolveByLoadMultiplier section, concrete, steel, nValue, mxValue, myValue, mxOffset, myOffset

    AssertTrue stats, "capacity.tfDefaults." & methodName & ".converged", cap.Converged
    AssertTrue stats, "capacity.tfDefaults." & methodName & ".physical", IsPhysicalLimitState(cap.LimitState)
    AssertClose stats, "capacity.tfDefaults." & methodName & ".lambda", cap.LambdaUltimate, referenceLambda, 0.2
    AssertTrue stats, "capacity.tfDefaults." & methodName & ".notOvershot", cap.LambdaUltimate < 2# * referenceLambda
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestLoadMultiplierSearchMethods(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim bisection As CCapacitySolver
    Dim brent As CCapacitySolver
    Dim secant As CCapacitySolver
    Set bisection = RunSearchMethod("Bisection", mesh, rebars)
    Set brent = RunSearchMethod("Brent", mesh, rebars)
    Set secant = RunSearchMethod("Secant", mesh, rebars)

    AssertTrue stats, "search.bisection.converged", bisection.Converged
    AssertTrue stats, "search.brent.converged", brent.Converged
    AssertTrue stats, "search.secant.converged", secant.Converged
    AssertTrue stats, "search.bisection.branch", InStr(1, bisection.DiagnosticLog, "searchMethod=Bisection", vbTextCompare) > 0
    AssertTrue stats, "search.brent.branch", InStr(1, brent.DiagnosticLog, "searchMethod=Brent", vbTextCompare) > 0
    AssertTrue stats, "search.secant.branch", InStr(1, secant.DiagnosticLog, "searchMethod=Secant", vbTextCompare) > 0
    AssertTrue stats, "search.brent.independent", InStr(1, brent.StopReason, "Brent", vbTextCompare) > 0
    AssertTrue stats, "search.secant.independent", InStr(1, secant.StopReason, "Secant", vbTextCompare) > 0
    AssertClose stats, "search.brent.lambda", brent.LambdaUltimate, bisection.LambdaUltimate, 0.02
    AssertClose stats, "search.secant.lambda", secant.LambdaUltimate, bisection.LambdaUltimate, 0.02
    AssertEquilibrium stats, "search.brent", brent.LastSolver, -300000#, brent.MxUltimate, 0#
    AssertEquilibrium stats, "search.secant", secant.LastSolver, -300000#, secant.MxUltimate, 0#
End Sub

' Проверяет все пользовательские траектории CapacityLoadPath на всех
' доступных способах поиска несущей способности. Тест намеренно работает
' на уровне CCapacitySolver: batch уже переводит пользовательские строки
' lambda*Mx/lambda*My/... в универсальную форму Offset + lambda*Base.
Private Sub TestCapacityLoadPathMethodMatrix(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    CheckCapacityLoadPathMethods stats, "mx", mesh, rebars, -120000#, 0#, 0#, -2000000#, 0#, 0#
    CheckCapacityLoadPathMethods stats, "my", mesh, rebars, -120000#, 0#, 0#, 0#, 0#, -2000000#
    CheckCapacityLoadPathMethods stats, "mxy", mesh, rebars, -120000#, 0#, 0#, -2000000#, 0#, -1200000#
    CheckCapacityLoadPathMethods stats, "n", mesh, rebars, 0#, -50000#, 0#, 0#, 0#, 0#
    CheckCapacityLoadPathMethods stats, "nmxy", mesh, rebars, 0#, -50000#, 0#, -1200000#, 0#, -800000#
End Sub

' Проверяет вырожденные, но допустимые lambda-траектории: в Base-векторе
' могут быть нулевые компоненты, если хотя бы одна компонента нагрузки реально
' масштабируется. Это защищает общий контракт solver-а Offset + lambda*Base:
' lambda*Mx не обязан иметь N, lambda*Mxy может содержать только один момент, а lambda*NMxy
' может фактически свестись к чистому N, чистому Mx или чистому My.
Private Sub TestCapacityLoadPathZeroComponentMatrix(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    CheckCapacityLoadPathMethods stats, "zero.mxNoN", mesh, rebars, 0#, 0#, 0#, -2000000#, 0#, 0#
    CheckCapacityLoadPathMethods stats, "zero.myNoN", mesh, rebars, 0#, 0#, 0#, 0#, 0#, -2000000#
    CheckCapacityLoadPathMethods stats, "zero.mxyOnlyMx", mesh, rebars, 0#, 0#, 0#, -2000000#, 0#, 0#
    CheckCapacityLoadPathMethods stats, "zero.mxyOnlyMy", mesh, rebars, 0#, 0#, 0#, 0#, 0#, -2000000#
    CheckCapacityLoadPathMethods stats, "zero.nmxyOnlyN", mesh, rebars, 0#, -50000#, 0#, 0#, 0#, 0#
    CheckCapacityLoadPathMethods stats, "zero.nmxyOnlyMx", mesh, rebars, 0#, 0#, 0#, -2000000#, 0#, 0#
    CheckCapacityLoadPathMethods stats, "zero.nmxyOnlyMy", mesh, rebars, 0#, 0#, 0#, 0#, 0#, -2000000#
End Sub

' Проверяет пользовательское Г-сечение именно на осевой траектории lambda*N.
' В этой задаче важно, что N приложена в бетонном центре тяжести: после
' переноса к координатам расчетных элементов внутри solver-а появляются
' связанные Mx/My, хотя пользователь масштабирует только продольную силу.
Private Sub TestLShapeCapacityLoadPathSmoke(ByRef stats As TCapacityTestStats)
    Dim section As CSectionModel
    Set section = LShapeCapacitySection()

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete section

    CheckLShapeCapacityPathMethods stats, "lshape.n.tension", section, 0#, 200# * 9806.65, _
        0#, 200# * 9806.65 * props.CentroidY, 0#, 200# * 9806.65 * props.CentroidX, True
    CheckLShapeCapacityPathMethods stats, "lshape.n.compression", section, 0#, -200# * 9806.65, _
        0#, -200# * 9806.65 * props.CentroidY, 0#, -200# * 9806.65 * props.CentroidX, True

    ' Пользовательский сценарий из книги: Г-сечение, N задана относительно
    ' бетонного центра тяжести, а предельная способность ищется по lambda*Mx.
    ' Здесь обязана включаться быстрая моментная постановка
    ' UltimateStrain: N постоянна, направление Mx/My сохраняется.
    CheckLShapeMomentUltimatePath stats, "lshape.moment.mx.userCase", section, _
        -200# * 9806.65, 50# * 9806.65 * 1000#, 0#, props.CentroidX, props.CentroidY
    CheckLShapeMomentUltimatePath stats, "lshape.moment.my.userCase", section, _
        -200# * 9806.65, 0#, 50# * 9806.65 * 1000#, props.CentroidX, props.CentroidY
    CheckLShapeMomentUltimatePath stats, "lshape.moment.mxy.userCase", section, _
        -200# * 9806.65, 50# * 9806.65 * 1000#, 25# * 9806.65 * 1000#, props.CentroidX, props.CentroidY
End Sub

' Проверяет моментную ветку UltimateStrain на несимметричном Г-сечении.
' Внутри solver-а момент от N добавляется как постоянный offset, а
' пользовательский момент масштабируется через lambda. Такой тест защищает
' рабочую моментную постановку от случайного ухода в общий load-path residual.
Private Sub CheckLShapeMomentUltimatePath(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal section As CSectionModel, ByVal nValue As Double, _
        ByVal userMxBase As Double, ByVal userMyBase As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double)
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap

    Dim mxOffset As Double
    Dim myOffset As Double
    mxOffset = nValue * referenceY
    myOffset = nValue * referenceX

    cap.SolveByUltimateLoadPath section, ProvisionalConcrete(), ProvisionalSteel(), _
        nValue, 0#, mxOffset, userMxBase, myOffset, userMyBase

    AppendLine stats, "INFO: " & prefix & _
        "; status=" & cap.LimitState & _
        "; converged=" & CStr(cap.Converged) & _
        "; solutionMethod=" & cap.SolutionMethod & _
        "; lambda=" & FormatNumberInvariant(cap.LambdaUltimate)
    AssertTrue stats, prefix & ".converged", cap.Converged
    AssertEquals stats, prefix & ".solutionMethod", cap.SolutionMethod, "UltimateStrain"
    AssertTrue stats, prefix & ".lambda", cap.LambdaUltimate > 0#
    AssertTrue stats, prefix & ".physical", IsPhysicalLimitState(cap.LimitState)
    AssertLoadComponent stats, prefix & ".equilibrium.N", cap.LastSolver.Nint, nValue, 10#, 0.0001
    AssertLoadComponent stats, prefix & ".equilibrium.Mx", cap.LastSolver.Mxint, _
        mxOffset + cap.LambdaUltimate * userMxBase, 10000#, 0.0002
    AssertLoadComponent stats, prefix & ".equilibrium.My", cap.LastSolver.Myint, _
        myOffset + cap.LambdaUltimate * userMyBase, 10000#, 0.0002
    AssertClose stats, prefix & ".mxUltimate.user", cap.MxUltimate, cap.LambdaUltimate * userMxBase, 20000#
    AssertClose stats, prefix & ".myUltimate.user", cap.MyUltimate, cap.LambdaUltimate * userMyBase, 20000#
End Sub

Private Sub CheckLShapeCapacityPathMethods(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal section As CSectionModel, ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        Optional ByVal allowForcePathFallback As Boolean = False)
    CheckLShapeCapacityPathMethod stats, prefix & ".auto", "Auto", "", section, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowForcePathFallback
    CheckLShapeCapacityPathMethod stats, prefix & ".ultimate", "UltimateStrain", "", section, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowForcePathFallback
    CheckLShapeCapacityPathMethod stats, prefix & ".bisection", "LoadMultiplier", "Bisection", section, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowForcePathFallback
    CheckLShapeCapacityPathMethod stats, prefix & ".brent", "LoadMultiplier", "Brent", section, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowForcePathFallback
    CheckLShapeCapacityPathMethod stats, prefix & ".secant", "LoadMultiplier", "Secant", section, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowForcePathFallback
End Sub

Private Sub CheckLShapeCapacityPathMethod(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal methodName As String, ByVal searchMethod As String, ByVal section As CSectionModel, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        Optional ByVal allowForcePathFallback As Boolean = False)
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    If Len(searchMethod) > 0 Then cap.SearchMethod = searchMethod

    Select Case methodName
        Case "Auto"
            If allowForcePathFallback Then
                cap.SolveByLoadPathMultiplier section, ProvisionalConcrete(), ProvisionalSteel(), _
                    nOffset, nBase, mxOffset, mxBase, myOffset, myBase, True
            Else
                cap.SolveByAutoLoadPath section, ProvisionalConcrete(), ProvisionalSteel(), _
                    nOffset, nBase, mxOffset, mxBase, myOffset, myBase
            End If
        Case "UltimateStrain"
            If allowForcePathFallback Then
                cap.SolveByLoadPathMultiplier section, ProvisionalConcrete(), ProvisionalSteel(), _
                    nOffset, nBase, mxOffset, mxBase, myOffset, myBase, True
            Else
                cap.SolveByUltimateLoadPath section, ProvisionalConcrete(), ProvisionalSteel(), _
                    nOffset, nBase, mxOffset, mxBase, myOffset, myBase
            End If
        Case Else
            cap.SolveByLoadPathMultiplier section, ProvisionalConcrete(), ProvisionalSteel(), _
                nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowForcePathFallback
    End Select

    AppendLine stats, "INFO: " & prefix & _
        "; status=" & cap.LimitState & _
        "; converged=" & CStr(cap.Converged) & _
        "; solutionMethod=" & cap.SolutionMethod & _
        "; lambda=" & FormatNumberInvariant(cap.LambdaUltimate)
    AssertTrue stats, prefix & ".converged", cap.Converged
    AssertTrue stats, prefix & ".lambda", cap.LambdaUltimate > 0#
    AssertTrue stats, prefix & ".physical", IsPhysicalLimitState(cap.LimitState)
    AssertLoadPathResult stats, prefix, cap, nOffset, nBase, mxOffset, mxBase, myOffset, myBase
End Sub

' Прогоняет чистый Nult через разные значения BaseLoadSteps. Настройка не
' является физическим параметром, но для осевой траектории на плато диаграммы
' она может заметно влиять на устойчивость отдельных probe-точек.
Private Sub TestNultBaseLoadStepsSensitivity(ByRef stats As TCapacityTestStats)
    Dim section As CSectionModel
    Set section = LShapeCapacitySection()

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete section

    CheckNultBaseLoadSteps stats, "tension", section, 200# * 9806.65, _
        200# * 9806.65 * props.CentroidY, 200# * 9806.65 * props.CentroidX
    CheckNultBaseLoadSteps stats, "compression", section, -200# * 9806.65, _
        -200# * 9806.65 * props.CentroidY, -200# * 9806.65 * props.CentroidX
End Sub

Private Sub CheckNultBaseLoadSteps(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal section As CSectionModel, ByVal nBase As Double, _
        ByVal mxBase As Double, ByVal myBase As Double)
    Dim steps As Variant
    steps = Array(1, 2, 4, 8)

    Dim i As Long
    For i = LBound(steps) To UBound(steps)
        Dim cap As CCapacitySolver
        Set cap = New CCapacitySolver
        ConfigureCapacity cap
        cap.SolverBaseLoadSteps = CLng(steps(i))
        cap.SolveByLoadPathMultiplier section, ProvisionalConcrete(), ProvisionalSteel(), _
            0#, nBase, 0#, mxBase, 0#, myBase, True

        AppendLine stats, "INFO: capacity.nult.baseLoadSteps." & prefix & "." & CStr(steps(i)) & _
            "; converged=" & CStr(cap.Converged) & _
            "; state=" & cap.LimitState & _
            "; lambda=" & FormatNumberInvariant(cap.LambdaUltimate) & _
            "; iterations=" & CStr(cap.Iterations) & _
            "; retries=" & CStr(cap.RetryCount)
        If CLng(steps(i)) > 1 Then
            AssertTrue stats, "capacity.nult.baseLoadSteps." & prefix & "." & CStr(steps(i)) & ".converged", cap.Converged
            AssertTrue stats, "capacity.nult.baseLoadSteps." & prefix & "." & CStr(steps(i)) & ".lambda", cap.LambdaUltimate > 0#
        End If
    Next i
End Sub

Private Sub CheckCapacityLoadPathMethods(ByRef stats As TCapacityTestStats, ByVal pathName As String, _
        ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double)
    CheckCapacityLoadPathMethod stats, pathName & ".auto", "Auto", "", mesh, rebars, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase
    CheckCapacityLoadPathMethod stats, pathName & ".ultimate", "UltimateStrain", "", mesh, rebars, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase
    CheckCapacityLoadPathMethod stats, pathName & ".bisection", "LoadMultiplier", "Bisection", mesh, rebars, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase
    CheckCapacityLoadPathMethod stats, pathName & ".brent", "LoadMultiplier", "Brent", mesh, rebars, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase
    CheckCapacityLoadPathMethod stats, pathName & ".secant", "LoadMultiplier", "Secant", mesh, rebars, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase
End Sub

Private Sub CheckCapacityLoadPathMethod(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal methodName As String, ByVal searchMethod As String, _
        ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double)
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    If Len(searchMethod) > 0 Then cap.SearchMethod = searchMethod

    Select Case methodName
        Case "Auto"
            cap.SolveByAutoLoadPath BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), _
                nOffset, nBase, mxOffset, mxBase, myOffset, myBase
        Case "UltimateStrain"
            cap.SolveByUltimateLoadPath BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), _
                nOffset, nBase, mxOffset, mxBase, myOffset, myBase
        Case Else
            cap.SolveByLoadPathMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), _
                nOffset, nBase, mxOffset, mxBase, myOffset, myBase
    End Select

    AssertTrue stats, "capacity.pathMatrix." & prefix & ".converged", cap.Converged
    AssertTrue stats, "capacity.pathMatrix." & prefix & ".lambda", cap.LambdaUltimate > 0#
    AssertTrue stats, "capacity.pathMatrix." & prefix & ".physical", IsPhysicalLimitState(cap.LimitState)
    If Not cap.Converged Then
        AppendLine stats, "DIAG: capacity.pathMatrix." & prefix & _
            "; limitState=" & cap.LimitState & _
            "; stopReason=" & cap.StopReason & _
            "; log=" & Replace(cap.DiagnosticLog, vbCrLf, " | ")
    End If
    AssertCapacitySolutionMethod stats, "capacity.pathMatrix." & prefix & ".solutionMethod", cap, methodName, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase
    If cap.Converged Then
        AssertLoadPathResult stats, "capacity.pathMatrix." & prefix, cap, _
            nOffset, nBase, mxOffset, mxBase, myOffset, myBase
    End If
End Sub

Private Function LShapeCapacitySection() As CSectionModel
    Dim geom As CGeometryLShape
    Set geom = New CGeometryLShape
    geom.Initialize 250#, 550#, 600#, 250#, 0#, 0#

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 50#, 50#, 1

    Dim builder As CLShapeRebarLayoutBuilder
    Set builder = New CLShapeRebarLayoutBuilder

    Dim rebars As CRebarLayout
    Set rebars = builder.Build(250#, 550#, 600#, 250#, 0#, 0#, _
        LShapeCapacityFaceSettings(5, 5), _
        LShapeCapacityFaceSettings(2, 2), _
        LShapeCapacityFaceSettings(2, 2), _
        LShapeCapacityFaceSettings(5, 5), _
        "A400")

    Set LShapeCapacitySection = BuildGeneratedSectionModel(mesh, rebars, "LShapeCapacityTest")
End Function

Private Function LShapeCapacityFaceSettings(ByVal count1 As Long, ByVal count2 As Long) As Variant
    LShapeCapacityFaceSettings = Array(40#, 40#, 32#, 32#, count1, count2, 80#, 80#, 80#, 80#, _
        0#, 0#, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")
End Function

' Проверяет диагностическое имя фактической ветки capacity.
' Solver показывает пользователю именно публичный метод настройки:
' UltimateStrain или LoadMultiplier. Внутреннее имя универсальной процедуры
' SolveByLoadPathMultiplier наружу не выводится.
Private Sub AssertCapacitySolutionMethod(ByRef stats As TCapacityTestStats, ByVal name As String, _
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
        AssertEquals stats, name, cap.SolutionMethod, "LoadMultiplier"
    ElseIf methodName = "UltimateStrain" Then
        If isPureAxial Then
            AssertEquals stats, name, cap.SolutionMethod, "LoadMultiplier"
        Else
            AssertEquals stats, name, cap.SolutionMethod, "UltimateStrain"
        End If
    Else
        AssertTrue stats, name, cap.SolutionMethod = "UltimateStrain" Or cap.SolutionMethod = "LoadMultiplier"
    End If
End Sub

Private Sub AssertLoadPathResult(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
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

    AssertLoadComponent stats, prefix & ".equilibrium.N", cap.LastSolver.Nint, expectedN, 10#, 0.0001
    AssertLoadComponent stats, prefix & ".equilibrium.Mx", cap.LastSolver.Mxint, expectedMx, 10000#, 0.0002
    AssertLoadComponent stats, prefix & ".equilibrium.My", cap.LastSolver.Myint, expectedMy, 10000#, 0.0002
    If Abs(nBase) > 0.000000001 Then
        AssertClose stats, prefix & ".nUltimate", cap.NUltimate, expectedN, 20#
    Else
        AssertClose stats, prefix & ".nUltimateBlank", cap.NUltimate, 0#, 0#
    End If
    AssertClose stats, prefix & ".mxUltimate", cap.MxUltimate, expectedMx, 20000#
    AssertClose stats, prefix & ".myUltimate", cap.MyUltimate, expectedMy, 20000#
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestSearchMethodInputErrors(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim emptyMethod As CCapacitySolver
    Set emptyMethod = New CCapacitySolver
    ConfigureCapacity emptyMethod
    emptyMethod.SearchMethod = ""
    emptyMethod.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, -10000000#, 0#

    Dim invalidMethod As CCapacitySolver
    Set invalidMethod = New CCapacitySolver
    ConfigureCapacity invalidMethod
    invalidMethod.SearchMethod = "FalseMethod"
    invalidMethod.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, -10000000#, 0#

    Dim secantFail As CCapacitySolver
    Set secantFail = New CCapacitySolver
    ConfigureCapacity secantFail
    secantFail.SearchMethod = "Secant"
    secantFail.SolverMaxIterations = 1
    secantFail.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, -10000000#, 0#

    AssertTrue stats, "search.empty.inputError", Not emptyMethod.Converged And emptyMethod.LimitState = "InvalidInput"
    AssertTrue stats, "search.invalid.inputError", Not invalidMethod.Converged And invalidMethod.LimitState = "InvalidInput"
    AssertTrue stats, "search.secant.controlledFailure", Not secantFail.Converged And _
        (secantFail.LimitState = "NumericalFailure" Or secantFail.LimitState = "InvalidInput")
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestSearchMethodPerformanceComparison(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    AppendSearchPerformance stats, "Bisection", mesh, rebars
    AppendSearchPerformance stats, "Brent", mesh, rebars
    AppendSearchPerformance stats, "Secant", mesh, rebars
End Sub

' Запускает связанный набор операций и возвращает пользователю итоговый статус выполнения.
Private Function RunSearchMethod(ByVal methodName As String, ByVal mesh As CFiberMeshBuilder, _
        ByVal rebars As CRebarLayout) As CCapacitySolver
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SearchMethod = methodName
    cap.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, -10000000#, 0#
    Set RunSearchMethod = cap
End Function

Private Sub AppendSearchPerformance(ByRef stats As TCapacityTestStats, ByVal methodName As String, _
        ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout)
    Dim runs As Long
    runs = 3
    Dim totalSec As Double
    Dim cap As CCapacitySolver
    Dim i As Long
    For i = 1 To runs
        Dim t0 As Double
        t0 = Timer
        Set cap = RunSearchMethod(methodName, mesh, rebars)
        totalSec = totalSec + (Timer - t0)
    Next i

    AssertTrue stats, "search.perf." & methodName & ".converged", cap.Converged
    AppendLine stats, "PERF_CAPACITY_SEARCH|" & methodName & _
        "|lambda=" & FormatNumberInvariant(cap.LambdaUltimate) & _
        "|iterations=" & CStr(cap.Iterations) & _
        "|functionEvaluations=" & CStr(cap.Iterations) & _
        "|avgSec=" & FormatNumberInvariant(totalSec / runs) & _
        "|residual=" & FormatNumberInvariant(cap.MomentEquilibriumResidual) & _
        "|converged=" & CStr(cap.Converged) & _
        "|status=" & cap.LimitState
End Sub

Private Sub AppendComparison(ByRef stats As TCapacityTestStats, ByVal name As String, _
        ByVal cap As CCapacitySolver, ByVal elapsedSec As Double)
    AppendLine stats, "COMPARE|" & name & _
        "|converged=" & CStr(cap.Converged) & _
        "|lambda=" & FormatNumberInvariant(cap.LambdaUltimate) & _
        "|limitState=" & cap.LimitState & _
        "|criticalStrain=" & FormatNumberInvariant(cap.CriticalStrain) & _
        "|criticalElement=" & cap.CriticalElement & _
        "|solverCalls=" & CStr(cap.Iterations) & _
        "|elapsedSec=" & FormatNumberInvariant(elapsedSec)
End Sub

Private Sub ConfigureCapacity(ByVal cap As CCapacitySolver)
    cap.ConcreteCompressionLimit = -0.0015
    cap.SteelStrainLimit = 0.00175
    cap.LambdaTolerance = 0.01
    cap.MaxLambda = 64#
    cap.InitialLambdaStep = 1#
    cap.MaxRetries = 4
    cap.SolverBaseLoadSteps = 8
    cap.SolverMaxIterations = 60
End Sub

Private Sub ConfigureSectionSolver(ByVal solver As CSectionSolver)
    solver.LoadSteps = 8
    solver.MaxIterations = 80
    solver.ToleranceN = 5#
    solver.ToleranceMx = 5000#
    solver.ToleranceMy = 5000#
End Sub

Private Function ProvisionalConcrete() As CMaterialDiagram
    Dim concrete As CMaterialDiagram
    Set concrete = New CMaterialDiagram
    concrete.Initialize -0.0015, -15.5, -0.0035, -15.5
    Set ProvisionalConcrete = concrete
End Function

Private Function ProvisionalConcreteWithTension() As CMaterialDiagram
    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.InitializeFromParameters TestConcreteParameters(), TestSteelParameters(), "TwoLine", "UseDiagram"
    Set ProvisionalConcreteWithTension = provider.ConcreteMaterial(cpStrength)
End Function

Private Function TestConcreteParameters() As CConcreteMaterialParameters
    Dim parameters As CConcreteMaterialParameters
    Set parameters = New CConcreteMaterialParameters
    parameters.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    Set TestConcreteParameters = parameters
End Function

Private Function TestSteelParameters() As CSteelMaterialParameters
    Dim parameters As CSteelMaterialParameters
    Set parameters = New CSteelMaterialParameters
    parameters.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#, "Ribbed"
    Set TestSteelParameters = parameters
End Function

Private Function ProvisionalSteel() As CMaterialDiagram
    Dim steel As CMaterialDiagram
    Set steel = New CMaterialDiagram
    steel.Initialize 0.00175, 350#, 0.025
    Set ProvisionalSteel = steel
End Function

Private Function LinearConcrete() As CLinearConcreteMaterial
    Dim concrete As CLinearConcreteMaterial
    Set concrete = New CLinearConcreteMaterial
    concrete.Initialize 32500#
    Set LinearConcrete = concrete
End Function

Private Function LinearSteel() As CLinearSteelMaterial
    Dim steel As CLinearSteelMaterial
    Set steel = New CLinearSteelMaterial
    steel.Initialize 200000#
    Set LinearSteel = steel
End Function

Private Sub PrepareSymmetricSection(ByVal width As Double, ByVal height As Double, ByVal stepSize As Double, _
        ByVal xAbs As Double, ByVal yAbs As Double, ByRef mesh As CFiberMeshBuilder, ByRef rebars As CRebarLayout)
    Dim geom As CGeometryRoundedRectangle
    Set geom = RectangleGeometry(width, height)
    Set mesh = BuildMesh(geom, stepSize)
    Set rebars = SymmetricRebars(geom, xAbs, yAbs)
End Sub

Private Function SymmetricRebars(ByVal geom As CGeometryRoundedRectangle, ByVal xAbs As Double, ByVal yAbs As Double) As CRebarLayout
    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -xAbs, -yAbs, 20#, 0#, "A400", "", geom
    rebars.AddBar "B2", xAbs, -yAbs, 20#, 0#, "A400", "", geom
    rebars.AddBar "B3", -xAbs, yAbs, 20#, 0#, "A400", "", geom
    rebars.AddBar "B4", xAbs, yAbs, 20#, 0#, "A400", "", geom
    Set SymmetricRebars = rebars
End Function

Private Function RectangleGeometry(ByVal width As Double, ByVal height As Double) As CGeometryRoundedRectangle
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize width, height, 0#, 0#, 0#, 0#
    Set RectangleGeometry = geom
End Function

' Создает расчетный или интерфейсный объект из нормализованных исходных данных и локальных настроек.
Private Function BuildMesh(ByVal geom As CGeometryRoundedRectangle, ByVal stepSize As Double) As CFiberMeshBuilder
    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, stepSize, stepSize, 1
    Set BuildMesh = mesh
End Function

Private Function IsPhysicalLimitState(ByVal state As String) As Boolean
    IsPhysicalLimitState = (state = "ConcreteStrainLimit" Or state = "ConcreteTensionStrainLimit" Or _
        state = "SteelStrainLimit")
End Function

Private Sub AssertEquilibrium(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal solver As CSectionSolver, ByVal n As Double, ByVal mx As Double, ByVal my As Double)
    AssertLoadComponent stats, prefix & ".N", solver.Nint, n, 10#, 0.0001
    AssertLoadComponent stats, prefix & ".Mx", solver.Mxint, mx, 10000#, 0.0001
    AssertLoadComponent stats, prefix & ".My", solver.Myint, my, 10000#, 0.0001
End Sub

Private Sub AssertLoadComponent(ByRef stats As TCapacityTestStats, ByVal name As String, _
        ByVal actual As Double, ByVal expected As Double, ByVal zeroTolerance As Double, _
        ByVal relTolerance As Double)
    If Abs(expected) <= zeroTolerance Then
        AssertClose stats, name, actual, expected, zeroTolerance
    Else
        AssertRelative stats, name, actual, expected, relTolerance
    End If
End Sub

Private Sub AssertTrue(ByRef stats As TCapacityTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name
    End If
End Sub

Private Sub AssertEquals(ByRef stats As TCapacityTestStats, ByVal name As String, ByVal actual As String, ByVal expected As String)
    If actual = expected Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name & "; actual=" & actual & "; expected=" & expected
    End If
End Sub

Private Sub AssertClose(ByRef stats As TCapacityTestStats, ByVal name As String, ByVal actual As Double, _
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

Private Sub AssertRelative(ByRef stats As TCapacityTestStats, ByVal name As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal relTolerance As Double)
    Dim relDiff As Double
    If Abs(expected) <= 0.000000001 Then
        relDiff = Abs(actual - expected)
    Else
        relDiff = Abs((actual - expected) / expected)
    End If

    If relDiff <= relTolerance Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name & "; actual=" & FormatNumberInvariant(actual) & _
            "; expected=" & FormatNumberInvariant(expected) & "; relDiff=" & FormatNumberInvariant(relDiff)
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name & "; actual=" & FormatNumberInvariant(actual) & _
            "; expected=" & FormatNumberInvariant(expected) & "; relDiff=" & FormatNumberInvariant(relDiff)
    End If
End Sub

Private Sub AppendLine(ByRef stats As TCapacityTestStats, ByVal text As String)
    stats.Report = stats.Report & text & vbCrLf
End Sub

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function















