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

' Выполняет регрессии физических пределов, знаков и общих поисковых методов.
' Отчет различает проверенный limit point, численный отказ и неверный вход.
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
    AppendLine stats, "RUN: TestLimitSearchSecantFinalizesCheckedRoot"
    TestLimitSearchSecantFinalizesCheckedRoot stats
    AppendLine stats, "RUN: TestLimitSearchBisectionIterationLimitFails"
    TestLimitSearchBisectionIterationLimitFails stats
    AppendLine stats, "RUN: TestAudit02GenericUltimateSearch"
    TestAudit02GenericUltimateSearch stats
    AppendLine stats, "RUN: TestAudit02GenericLoadMultiplierMatrix"
    TestAudit02GenericLoadMultiplierMatrix stats
    TestAudit03SearchArithmetic stats
    TestAudit03RealCapacityPrecision stats
    TestAudit03CapacityTypedFailures stats
    TestAudit03CapacitySearchLifecycle stats
    TestAudit03UltimateGuards stats
    TestAudit03MultiplierTypedFaults stats
    AppendLine stats, "RUN: TestCapacityLoadPathMethodMatrix"
    TestCapacityLoadPathMethodMatrix stats
    AppendLine stats, "RUN: TestCapacityLoadPathZeroComponentMatrix"
    TestCapacityLoadPathZeroComponentMatrix stats
    AppendLine stats, "RUN: TestRectSetCapacityLoadPathSmoke"
    TestRectSetCapacityLoadPathSmoke stats
    AppendLine stats, "RUN: TestNultBaseLoadStepsSensitivity"
    TestNultBaseLoadStepsSensitivity stats
    AppendLine stats, "RUN: TestSearchMethodInputErrors"
    TestSearchMethodInputErrors stats
    AppendLine stats, "RUN: TestInitialLambdaFailureStatusMapping"
    TestInitialLambdaFailureStatusMapping stats
    AppendLine stats, "RUN: TestSearchMethodPerformanceComparison"
    TestSearchMethodPerformanceComparison stats

    AppendLine stats, "RUN: TestAudit02CapacitySearchBoundary"
    TestAudit02CapacitySearchBoundary stats
    AppendLine stats, "RUN: TestAudit02AsymmetricSteelLimits"
    TestAudit02AsymmetricSteelLimits stats
    AppendLine stats, "RUN: TestAudit02UnconvergedProbeIsNumerical"
    TestAudit02UnconvergedProbeIsNumerical stats
    TestAudit02PositiveUnconvergedProbe stats
    TestAudit02InitialOffsetBoundary stats

    AppendLine stats, "TOTAL_CAPACITY: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
    RunCapacitySolverTests = stats.Report
    Exit Function

Failed:
    RunCapacitySolverTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

' Решает два направления Mx при постоянной сжимающей N: проверяет знаки
' предельных моментов, физический критерий и равновесие каждой конечной точки.
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

' Проверяет симметрию несущей круглого сечения по X/Y и сохранение направления
' косого изгиба. Допуск симметрии учитывает дискретизацию исходной тестовой сетки.
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

' Сравнивает Ignore/UseDiagram для растянутого бетона при одинаковых нагрузках.
' Активная ветвь должна менять равновесную плоскость в одно-/двухосных задачах.
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


' Решает один вектор усилий с отключенным и активным растянутым бетоном;
' проверяет сходимость обеих моделей и значимое различие параметров плоскости.
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

' Проверяет все четыре квадранта Mx/My при одинаковой постоянной N.
' Предельные моменты сохраняют знаки и отношение компонентов заданного пути.
Private Sub TestMxySignedCombinations(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    CheckMxyCombination stats, "capacity.mxy.pp", mesh, rebars, -250000#, 6000000#, 4000000#
    CheckMxyCombination stats, "capacity.mxy.np", mesh, rebars, -250000#, -6000000#, 4000000#
    CheckMxyCombination stats, "capacity.mxy.pn", mesh, rebars, -250000#, 6000000#, -4000000#
    CheckMxyCombination stats, "capacity.mxy.nn", mesh, rebars, -250000#, -6000000#, -4000000#
End Sub

' Проверяет одну косую траекторию: положительную lambda, направление моментов
' и равновесие конечной точки. Не заменяет косую задачу независимыми осями.
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

' Решает оба направления My при постоянной N и проверяет знаки конечных
' моментов, физический управляющий предел и компонентное равновесие.
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
' Удваивает предварительно найденный предельный момент и ожидает lambda<1.
' Успех Search при этом не должен маскировать инженерный FAIL заданного LC.
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
    CheckAudit02SearchVsCapacity stats, cap, BuildGeneratedSectionModel(mesh, rebars), _
        -300000#, reference.MxUltimate * 2#, 0#
End Sub

' Повторно ищет предел от момента уже найденной опорной точки;
' lambda должна быть около 1 в исходном допуске, без изменения критерия.
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

' Проверяет чистый изгиб при N=0 на линейных материалах с заданными пределами.
' Нулевая продольная сила не должна нарушать поиск или равновесие точки.
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

' Делает сжатие бетона управляющим ограничением, оставляя большой предел
' арматуры; поиск должен завершиться именно ConcreteStrainLimit.
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

' Включает активную tensile-ветвь и малый растягивающий предел бетона.
' Управляющий код должен относиться к растяжению, а не сжатию или арматуре.
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

' Отключает растянутый бетон при таком же малом tensile-пределе;
' неактивная ветвь не ограничивает несущую, в этом fixture управляет арматура.
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

' Делает деформацию арматуры управляющей на линейных материалах;
' ожидается SteelStrainLimit при сохраненном поиске равновесия.
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

' Проверяет старт чистого изгиба из подготовленной StateGuessBuilder плоскости.
' При жестком бюджете итераций должен находиться физический предел, а журнал
' подтверждать использование начального приближения вместо нулевого излома.
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

    If Not cap.Converged Then AppendLine stats, "DIAGNOSTIC pureBendingGuess: " & cap.StopReason & vbCrLf & cap.DiagnosticLog

    AssertTrue stats, "capacity.pureBendingGuess.converged", cap.Converged
    AssertTrue stats, "capacity.pureBendingGuess.physical", IsPhysicalLimitState(cap.LimitState)
    AssertTrue stats, "capacity.pureBendingGuess.usedGuess", InStr(1, cap.DiagnosticLog, "для lambda-точки чистого изгиба применена стартовая плоскость", vbTextCompare) > 0
End Sub

' Проверяет связность кривизн несимметричного сечения с ненулевым Ixy:
' даже при My=0 равновесная KappaY может быть ненулевой.
Private Sub TestAsymmetricCoupledCurvatures(ByRef stats As TCapacityTestStats)
    Dim geom As CGeometryRectSet
    Set geom = CapacityAsymmetricGeometry()

    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 10#)

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete BuildGeneratedSectionModel(mesh, Nothing)

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", 50#, 50#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B2", 550#, 50#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B3", 50#, 700#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B4", 200#, 700#, 20#, 0#, "A400", "", geom

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -260000#, -8000000#, 0#

    AssertTrue stats, "capacity.asym.Ixy.nonzero", Abs(props.Ixyc) > 1000000#
    AssertTrue stats, "capacity.asym.converged", cap.Converged
    AssertTrue stats, "capacity.asym.coupledKappaY", Abs(cap.LastSolver.KappaY) > 0.000000001
    AssertEquilibrium stats, "capacity.asym", cap.LastSolver, -260000#, cap.MxUltimate, 0#
End Sub

' Решает косой изгиб несимметричного RectSet и проверяет обе кривизны,
' отношение предельных моментов и равновесие относительно исходных осей.
Private Sub TestAsymmetricMxy(ByRef stats As TCapacityTestStats)
    Dim geom As CGeometryRectSet
    Set geom = CapacityAsymmetricGeometry()

    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 10#)

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete BuildGeneratedSectionModel(mesh, Nothing)

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", 50#, 50#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B2", 550#, 50#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B3", 50#, 700#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B4", 200#, 700#, 20#, 0#, "A400", "", geom

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

' Передает нулевую масштабируемую часть именно моментного пути.
' Наличие постоянной N не делает такой путь допустимым: ожидается InvalidInput.
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

' Сопоставляет LoadMultiplier и UltimateStrain на прямоугольнике/круге,
' включая управляющие деформации и отдельные контракты численного отказа.
Private Sub TestCapacitySolutionStrategyComparisons(ByRef stats As TCapacityTestStats)
    TestMethodPureCompression stats
    CompareCapacitySolutionStrategys stats, "method.n_plus_mx", -300000#, -10000000#, 0#, False
    CompareCapacitySolutionStrategys stats, "method.n_plus_my", -300000#, 0#, -10000000#, False
    CompareCapacitySolutionStrategys stats, "method.biaxial", -250000#, -6000000#, -4000000#, False
    CompareCircleCapacitySolutionStrategys stats
    TestMethodStrainLimitState stats
    TestUltimateStrainNumericalFailure stats
    TestUltimateStrainDoesNotApplyMaxLambda stats
End Sub

' Проверяет, что моментные entrypoint-ы обоих методов отклоняют чистую N
' без базового момента; допустимая осевая lambda*N тестируется отдельным путем.
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

' Решает одинаковую моментную траекторию двумя стратегиями и сравнивает
' lambda/управляющий предел; время сохраняется как наблюдение, не порог PASS.
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

' Сопоставляет стратегии на симметричном круге с повернутыми X/Y задачами;
' проверяет конечные моменты и определение критического элемента.
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

' Сравнивает обе стратегии при специально заданном бетонном strain-пределе.
' UltimateStrain должен подтвердить тот же управляющий критерий и utilization.
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

' Ограничивает UltimateStrain одной итерацией и проверяет честный численный
' отказ: неподтвержденная проба не становится физической верхней границей.
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

' Проверяет, что MaxLambda не отбрасывает уже найденное UltimateStrain-решение.
' Эта настройка ограничивает только расширение скобки LoadMultiplier.
Private Sub TestUltimateStrainDoesNotApplyMaxLambda(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.MaxLambda = 0.5
    cap.SolveByUltimateStrain BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, -1000000#, 0#

    AssertTrue stats, "method.maxLambdaIgnored.converged", cap.Converged
    AssertTrue stats, "method.maxLambdaIgnored.lambdaAboveLimit", cap.LambdaUltimate > 0.5
    AssertTrue stats, "method.maxLambdaIgnored.momentStored", Abs(cap.MomentUltimate) > 0#
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

    AssertEquals stats, "capacity.tfDefaults.geometry", settings.GetString("Geometry.Type", ""), "RectSet"
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
    Dim loadState As CSectionLoadState
    Set loadState = New CSectionLoadState
    loadState.Initialize nValue, mxValue, myValue, props.CentroidX, props.CentroidY
    mxOffset = loadState.AxialMxAboutPoint(0#)
    myOffset = loadState.AxialMyAboutPoint(0#)

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

' Проверяет один метод LoadMultiplier на настройках, прочитанных в tf/tf*m.
' Сравнение с независимой веткой UltimateStrain сохраняет заданный допуск lambda.
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

' Сравнивает Bisection/Brent/Secant на одной моментной траектории;
' каждый метод должен выполнить свою ветвь и подтвердить близкую lambda.
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

' Проверяет дефект T01: Secant обязан финализировать именно проверенную точку,
' где целевая функция стала нулевой, а не старую нижнюю границу скобки.
Private Sub TestLimitSearchSecantFinalizesCheckedRoot(ByRef stats As TCapacityTestStats)
    On Error GoTo Failed
    Dim stage As String

    stage = "configure"
    Dim problem As CTestLimitSearchProblem
    Set problem = New CTestLimitSearchProblem
    problem.Configure "Secant", 0.75, 0.000000001, 8, 1#, 16#, rkCapacity

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
    AssertTrue stats, "limitSearch.secant.resultObject", Not result Is Nothing
    AssertTrue stats, "limitSearch.secant.converged", problem.Converged
    AssertEquals stats, "limitSearch.secant.method", problem.FinalMethod, "Secant"
    AssertClose stats, "limitSearch.secant.root", problem.FinalLambda, 0.75, 0.000000001
    Exit Sub

Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: limitSearch.secant.runtime; stage=" & stage & _
        "; error=" & CStr(Err.Number) & "; description=" & Err.Description
End Sub

' Проверяет дефект T02: если bisection остановлен лимитом итераций до достижения
' допуска, общий search не должен выдавать недоточненную границу как найденный предел.
Private Sub TestLimitSearchBisectionIterationLimitFails(ByRef stats As TCapacityTestStats)
    On Error GoTo Failed
    Dim stage As String

    stage = "configure"
    Dim problem As CTestLimitSearchProblem
    Set problem = New CTestLimitSearchProblem
    problem.Configure "Bisection", 0.2, 0.001, 1, 1#, 16#, rkCrackFormation

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
    AssertTrue stats, "limitSearch.bisection.resultObject", Not result Is Nothing
    AssertTrue stats, "limitSearch.bisection.iterationLimitHit", problem.IterationLimitHit
    AssertTrue stats, "limitSearch.bisection.notConverged", Not problem.Converged
    AssertEquals stats, "limitSearch.bisection.limitState", problem.LimitState, "NumericalFailure"
    AssertTrue stats, "limitSearch.bisection.noFakeLimit", Abs(problem.FinalLambda) <= 0.000000001
    Exit Sub

Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: limitSearch.bisection.runtime; stage=" & stage & _
        "; error=" & CStr(Err.Number) & "; description=" & Err.Description
End Sub

' Проверяет все пользовательские траектории LoadPath на всех
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
Private Sub TestRectSetCapacityLoadPathSmoke(ByRef stats As TCapacityTestStats)
    Dim section As CSectionModel
    Set section = RectSetCapacitySection()

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete section

    Dim tensionLoad As CSectionLoadState
    Set tensionLoad = New CSectionLoadState
    tensionLoad.Initialize 200# * 9806.65, 0#, 0#, props.CentroidX, props.CentroidY
    CheckRectSetCapacityPathMethods stats, "rectset.n.tension", section, 0#, tensionLoad.N, _
        0#, tensionLoad.AxialMxAboutPoint(0#), 0#, tensionLoad.AxialMyAboutPoint(0#), True

    Dim compressionLoad As CSectionLoadState
    Set compressionLoad = New CSectionLoadState
    compressionLoad.Initialize -200# * 9806.65, 0#, 0#, props.CentroidX, props.CentroidY
    CheckRectSetCapacityPathMethods stats, "rectset.n.compression", section, 0#, compressionLoad.N, _
        0#, compressionLoad.AxialMxAboutPoint(0#), 0#, compressionLoad.AxialMyAboutPoint(0#), True

    ' Пользовательский сценарий из книги: Г-сечение, N задана относительно
    ' бетонного центра тяжести, а предельная способность ищется по lambda*Mx.
    ' Здесь обязана включаться быстрая моментная постановка
    ' UltimateStrain: N постоянна, направление Mx/My сохраняется.
    CheckRectSetMomentUltimatePath stats, "rectset.moment.mx.userCase", section, _
        -200# * 9806.65, 50# * 9806.65 * 1000#, 0#, props.CentroidX, props.CentroidY
    CheckRectSetMomentUltimatePath stats, "rectset.moment.my.userCase", section, _
        -200# * 9806.65, 0#, 50# * 9806.65 * 1000#, props.CentroidX, props.CentroidY
    CheckRectSetMomentUltimatePath stats, "rectset.moment.mxy.userCase", section, _
        -200# * 9806.65, 50# * 9806.65 * 1000#, 25# * 9806.65 * 1000#, props.CentroidX, props.CentroidY
End Sub

' Проверяет моментную ветку UltimateStrain на несимметричном Г-сечении.
' Внутри solver-а момент от N добавляется как постоянный offset, а
' пользовательский момент масштабируется через lambda. Такой тест защищает
' рабочую моментную постановку от случайного ухода в общий load-path residual.
Private Sub CheckRectSetMomentUltimatePath(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal section As CSectionModel, ByVal nValue As Double, _
        ByVal userMxBase As Double, ByVal userMyBase As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double)
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap

    Dim mxOffset As Double
    Dim myOffset As Double
    Dim loadState As CSectionLoadState
    Set loadState = New CSectionLoadState
    loadState.Initialize nValue, userMxBase, userMyBase, referenceX, referenceY
    mxOffset = loadState.AxialMxAboutPoint(0#)
    myOffset = loadState.AxialMyAboutPoint(0#)

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

' Запускает одинаковый RectSet load path всеми разрешенными стратегиями поиска.
' Для силовой траектории явно передает принятый переход к LoadMultiplier.
Private Sub CheckRectSetCapacityPathMethods(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal section As CSectionModel, ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        Optional ByVal allowForcePathFallback As Boolean = False)
    CheckRectSetCapacityPathMethod stats, prefix & ".auto", "Auto", "", section, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowForcePathFallback
    CheckRectSetCapacityPathMethod stats, prefix & ".ultimate", "UltimateStrain", "", section, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowForcePathFallback
    CheckRectSetCapacityPathMethod stats, prefix & ".bisection", "LoadMultiplier", "Bisection", section, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowForcePathFallback
    CheckRectSetCapacityPathMethod stats, prefix & ".brent", "LoadMultiplier", "Brent", section, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowForcePathFallback
    CheckRectSetCapacityPathMethod stats, prefix & ".secant", "LoadMultiplier", "Secant", section, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, allowForcePathFallback
End Sub

' Выполняет один выбранный поиск на готовом RectSet и проверяет физический предел.
' Усилия конечного State дополнительно проверяются на исходной lambda-траектории.
Private Sub CheckRectSetCapacityPathMethod(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
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
    Set section = RectSetCapacitySection()

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete section

    Dim tensionLoad As CSectionLoadState
    Set tensionLoad = New CSectionLoadState
    tensionLoad.Initialize 200# * 9806.65, 0#, 0#, props.CentroidX, props.CentroidY
    CheckNultBaseLoadSteps stats, "tension", section, tensionLoad.N, _
        tensionLoad.AxialMxAboutPoint(0#), tensionLoad.AxialMyAboutPoint(0#)

    Dim compressionLoad As CSectionLoadState
    Set compressionLoad = New CSectionLoadState
    compressionLoad.Initialize -200# * 9806.65, 0#, 0#, props.CentroidX, props.CentroidY
    CheckNultBaseLoadSteps stats, "compression", section, compressionLoad.N, _
        compressionLoad.AxialMxAboutPoint(0#), compressionLoad.AxialMyAboutPoint(0#)
End Sub

' Повторяет осевой путь с 1/2/4/8 ступенями внутренних решений равновесия.
' Один шаг диагностируется отдельно; успешность обязательна для остальных вариантов.
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

' Сравнивает стратегии поиска на одной сетке, арматуре и Offset/Base-траектории.
' Численные методы LoadMultiplier проверяются каждый отдельным запуском.
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

' Собирает CSectionModel и запускает только запрошенный метод поиска предела.
' При отказе сохраняет диагностику; при успехе проверяет равновесие и путь нагрузки.
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

' Строит фиксированное двухпрямоугольное сечение с арматурой по четырем граням.
' Геометрия, шаг сетки и диаметры независимы от пользовательского Config.
Private Function RectSetCapacitySection() As CSectionModel
    Dim geom As CGeometryRectSet
    Set geom = New CGeometryRectSet
    geom.Initialize 250#, 550#, 600#, 250#, 0#, 0#

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 50#, 50#, 1

    Dim builder As CRectSetRebarLayoutBuilder
    Set builder = New CRectSetRebarLayoutBuilder

    Dim rebars As CRebarLayout
    Set rebars = builder.Build(250#, 550#, 600#, 250#, 0#, 0#, _
        RectSetCapacityFaceSettings(5, 5), _
        RectSetCapacityFaceSettings(2, 2), _
        RectSetCapacityFaceSettings(2, 2), _
        RectSetCapacityFaceSettings(5, 5), _
        "A400")

    Set RectSetCapacitySection = BuildGeneratedSectionModel(mesh, rebars, "RectSetCapacityTest")
End Function

' Возвращает полный контракт грани с заданным количеством стержней первого ряда.
' Дополнительные ряды отключены нулевыми диаметрами, отступы и режимы фиксированы.
Private Function RectSetCapacityFaceSettings(ByVal count1 As Long, ByVal count2 As Long) As Variant
    RectSetCapacityFaceSettings = Array(40#, 40#, 32#, 32#, count1, count2, 80#, 80#, 80#, 80#, _
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

' Независимо восстанавливает N/Mx/My по Offset + lambda*Base и сверяет State.
' Выводимый Nult пуст по контракту, если продольная сила пути не масштабируется.
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

' Проверяет пустое/неизвестное имя Search и контролируемый отказ Secant
' при малом бюджете: ни один сценарий не переключается молча на другой метод.
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
    Set numericalResult = Audit02CapacitySnapshot(numericalCap, "LoadMultiplier")
    AssertEquals stats, "capacity.initialLambda.numerical.external", _
        policy.ExternalStatus(numericalResult.Meta), "NumFail"

    Dim physicalCap As CCapacitySolver
    Set physicalCap = New CCapacitySolver
    physicalCap.LimitSearchHandleCapacityInitialFailure "ConcreteStrainLimit"

    Dim physicalResult As CLimitSearchResult
    Set physicalResult = Audit02CapacitySnapshot(physicalCap, "LoadMultiplier")
    AssertEquals stats, "capacity.initialLambda.physical.external", _
        policy.ExternalStatus(physicalResult.Meta), "BaseFail"
End Sub

' Сохраняет три повторных измерения каждого Search-метода на общем fixture.
' Эти строки предназначены для сравнения, а не для утверждения ускорения по одному запуску.
Private Sub TestSearchMethodPerformanceComparison(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    AppendSearchPerformance stats, "Bisection", mesh, rebars
    AppendSearchPerformance stats, "Brent", mesh, rebars
    AppendSearchPerformance stats, "Secant", mesh, rebars
End Sub

' Решает общий контрольный моментный путь указанным Bisection/Brent/Secant
' и возвращает solver с фактическими счетчиками и конечной диагностикой.
Private Function RunSearchMethod(ByVal methodName As String, ByVal mesh As CFiberMeshBuilder, _
        ByVal rebars As CRebarLayout) As CCapacitySolver
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SearchMethod = methodName
    cap.SolveByLoadMultiplier BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, -10000000#, 0#
    Set RunSearchMethod = cap
End Function

' Выполняет повторные измерения метода на свежем solver-е и добавляет
' среднее время вместе с результатом; не изменяет настройки ради быстродействия.
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

' Сохраняет результаты и стоимость поиска для сравнения стратегий без пересчета.
' Указывает найденную lambda, управляющий элемент и число обращений к решателю.
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

' Задает фиксированные пределы и поисковый бюджет контрольных задач Capacity.
' Это fixture, а не нормативный default или скрытая настройка рабочего расчета.
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

' Настраивает независимую проверку равновесия во внутренних единицах N и N*mm.
' Значения не читаются из Config и не меняют допуски проверяемого поискового метода.
Private Sub ConfigureSectionSolver(ByVal solver As CSectionSolver)
    solver.LoadSteps = 8
    solver.MaxIterations = 80
    solver.ToleranceN = 5#
    solver.ToleranceMx = 5000#
    solver.ToleranceMy = 5000#
End Sub

' Создает фиксированный сжатый бетон без растянутой ветви и расширения.
Private Function ProvisionalConcrete() As CMaterialDiagram
    Dim concrete As CMaterialDiagram
    Set concrete = New CMaterialDiagram
    concrete.Initialize -0.0015, -15.5, -0.0035, -15.5
    Set ProvisionalConcrete = concrete
End Function

' Создает физическую TwoLine-диаграмму с растяжением через обычный material provider.
' Параметры берутся из fixtures ниже, а не из листа Config.
Private Function ProvisionalConcreteWithTension() As CMaterialDiagram
    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.InitializeFromParameters TestConcreteParameters(), TestSteelParameters(), "TwoLine", "UseDiagram"
    Set ProvisionalConcreteWithTension = provider.ConcreteMaterial(cpStrength)
End Function

' Задает оба расчетных набора бетона и сопротивление продольным трещинам.
' Значения служат воспроизводимому fixture, не нормативной трассировке.
Private Function TestConcreteParameters() As CConcreteMaterialParameters
    Dim parameters As CConcreteMaterialParameters
    Set parameters = New CConcreteMaterialParameters
    parameters.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    Set TestConcreteParameters = parameters
End Function

' Задает фиксированные модули и сопротивления арматуры обоих расчетных наборов.
Private Function TestSteelParameters() As CSteelMaterialParameters
    Dim parameters As CSteelMaterialParameters
    Set parameters = New CSteelMaterialParameters
    parameters.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#
    Set TestSteelParameters = parameters
End Function

' Возвращает симметричную физическую диаграмму арматуры с постоянным плато.
Private Function ProvisionalSteel() As CMaterialDiagram
    Dim steel As CMaterialDiagram
    Set steel = New CMaterialDiagram
    steel.Initialize 0.00175, 350#, 0.025
    Set ProvisionalSteel = steel
End Function

' Возвращает линейный бетон для аналитически проверяемых задач равновесия.
Private Function LinearConcrete() As CLinearConcreteMaterial
    Dim concrete As CLinearConcreteMaterial
    Set concrete = New CLinearConcreteMaterial
    concrete.Initialize 32500#
    Set LinearConcrete = concrete
End Function

' Возвращает линейную арматуру; предел для поиска задается отдельно в Capacity.
Private Function LinearSteel() As CLinearSteelMaterial
    Dim steel As CLinearSteelMaterial
    Set steel = New CLinearSteelMaterial
    steel.Initialize 200000#
    Set LinearSteel = steel
End Function

' Готовит симметричный прямоугольный fixture и возвращает его сетку и арматуру.
' Центр находится в начале координат, четыре одинаковых стержня зеркальны по осям.
Private Sub PrepareSymmetricSection(ByVal width As Double, ByVal height As Double, ByVal stepSize As Double, _
        ByVal xAbs As Double, ByVal yAbs As Double, ByRef mesh As CFiberMeshBuilder, ByRef rebars As CRebarLayout)
    Dim geom As CGeometryRoundedRectangle
    Set geom = RectangleGeometry(width, height)
    Set mesh = BuildMesh(geom, stepSize)
    Set rebars = SymmetricRebars(geom, xAbs, yAbs)
End Sub

' Расставляет четыре стержня диаметром 20 мм в зеркальных точках (+/-x, +/-y).
' Каждый стержень проходит обычную проверку принадлежности геометрии.
Private Function SymmetricRebars(ByVal geom As CGeometryRoundedRectangle, ByVal xAbs As Double, ByVal yAbs As Double) As CRebarLayout
    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -xAbs, -yAbs, 20#, 0#, "A400", "", geom
    rebars.AddBar "B2", xAbs, -yAbs, 20#, 0#, "A400", "", geom
    rebars.AddBar "B3", -xAbs, yAbs, 20#, 0#, "A400", "", geom
    rebars.AddBar "B4", xAbs, yAbs, 20#, 0#, "A400", "", geom
    Set SymmetricRebars = rebars
End Function

' Возвращает прямоугольник без скруглений, поворота и смещения для fixtures.
Private Function RectangleGeometry(ByVal width As Double, ByVal height As Double) As CGeometryRoundedRectangle
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize width, height, 0#, 0#, 0#, 0#
    Set RectangleGeometry = geom
End Function

' Возвращает тестовую несимметричную геометрию с заметным Ixy для проверки,
' что capacity-решатель учитывает связанную кривизну в общей постановке N+Mx+My.
Private Function CapacityAsymmetricGeometry() As CGeometryRectSet
    Dim geom As CGeometryRectSet
    Set geom = New CGeometryRectSet
    geom.Initialize 250#, 550#, 600#, 250#, 0#, 0#, 0#, "LSection"
    Set CapacityAsymmetricGeometry = geom
End Function

' Строит тестовую сетку любого поддерживаемого ISectionGeometry с одинаковым
' шагом по осям; затем сетка и арматура объединяются в CSectionModel.
Private Function BuildMesh(ByVal geom As ISectionGeometry, ByVal stepSize As Double) As CFiberMeshBuilder
    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, stepSize, stepSize, 1
    Set BuildMesh = mesh
End Function

' Отличает физические деформационные критерии от технического исхода поиска.
' Это проверка диагностического поля теста, не назначение пользовательского статуса.
Private Function IsPhysicalLimitState(ByVal state As String) As Boolean
    IsPhysicalLimitState = (state = "ConcreteStrainLimit" Or state = "ConcreteTensionStrainLimit" Or _
        state = "SteelStrainLimit")
End Function

' Сверяет все компоненты внутренних усилий с независимо заданной нагрузкой.
' Нулевая сила и нулевой момент проверяются по своим абсолютным допускам.
Private Sub AssertEquilibrium(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal solver As CSectionSolver, ByVal n As Double, ByVal mx As Double, ByVal my As Double)
    AssertLoadComponent stats, prefix & ".N", solver.Nint, n, 10#, 0.0001
    AssertLoadComponent stats, prefix & ".Mx", solver.Mxint, mx, 10000#, 0.0001
    AssertLoadComponent stats, prefix & ".My", solver.Myint, my, 10000#, 0.0001
End Sub

' Не делит на почти нулевой эталон: выбирает абсолютную или относительную ошибку.
Private Sub AssertLoadComponent(ByRef stats As TCapacityTestStats, ByVal name As String, _
        ByVal actual As Double, ByVal expected As Double, ByVal zeroTolerance As Double, _
        ByVal relTolerance As Double)
    If Abs(expected) <= zeroTolerance Then
        AssertClose stats, name, actual, expected, zeroTolerance
    Else
        AssertRelative stats, name, actual, expected, relTolerance
    End If
End Sub

' Группа assertions ведет счет и протоколирует условия и фактические отклонения.
' Допуски численных сравнений задаются самим контрольным случаем.
Private Sub AssertTrue(ByRef stats As TCapacityTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name
    End If
End Sub

' Сравнивает строковый результат без нормализации; при отказе сохраняет обе
' строки, чтобы отличить неверный статус или код от численной ошибки теста.
Private Sub AssertEquals(ByRef stats As TCapacityTestStats, ByVal name As String, ByVal actual As String, ByVal expected As String)
    If actual = expected Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name & "; actual=" & actual & "; expected=" & expected
    End If
End Sub

' Сравнивает по абсолютному допуску с записью actual/expected и величины ошибки.
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

' Проверяет относительное отклонение; для почти нулевого эталона избегает деления.
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

' Дополняет тестовый протокол; числовые поля ниже используют точку в любой локали.
Private Sub AppendLine(ByRef stats As TCapacityTestStats, ByVal text As String)
    stats.Report = stats.Report & text & vbCrLf
End Sub

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function

' ============================== ДЛЯ ТЕСТОВ ==============================

' Проверяет предел в последнем интервале удвоения, точно на MaxLambda и
' за MaxLambda. Последняя ситуация не должна выдавать диагностическую нижнюю
' точку как найденную несущую способность.
Private Sub TestAudit02CapacitySearchBoundary(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars
    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, rebars)
    Dim stiffness As Double
    stiffness = Audit02AxialStiffness(section)

    Dim roots As Variant
    roots = Array(3#, 5.25, 6#, 8#)
    Dim root As Variant
    For Each root In roots
        Dim cap As CCapacitySolver
        Set cap = New CCapacitySolver
        ConfigureCapacity cap
        cap.ConcreteCompressionLimit = -1#
        cap.SteelStrainLimit = 0.001#
        cap.MaxLambda = 6#
        cap.LambdaTolerance = 0.000001
        cap.SolverToleranceN = 0.01
        cap.SolveByLoadPathMultiplier section, LinearConcrete(), LinearSteel(), _
            0#, stiffness * 0.001# / CDbl(root), 0#, 0#, 0#, 0#
        Dim prefix As String
        prefix = "audit02.boundary." & CStr(root)
        If CDbl(root) <= 6# Then
            AssertTrue stats, prefix & ".converged", cap.Converged
            AssertClose stats, prefix & ".lambda", cap.LambdaUltimate, CDbl(root), 0.00001
        Else
            AssertTrue stats, prefix & ".noLimit", Not cap.Converged
            AssertTrue stats, prefix & ".boundFlag", cap.SearchBoundReached
            AssertClose stats, prefix & ".notCapacity", cap.LambdaUltimate, 0#, 0#
            Dim searchResult As CLimitSearchResult
            Set searchResult = Audit02CapacitySnapshot(cap, "LoadMultiplier")
            AssertTrue stats, prefix & ".code", searchResult.Meta.ResultCode = rcSearchBoundReached
        End If
    Next root
End Sub

' Проверяет независимые пределы стали двух знаков на линейном материале.
' Числа заданы аналитически из общей осевой жесткости и физических деформаций.
Private Sub TestAudit02AsymmetricSteelLimits(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars
    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, rebars)
    Dim methodIndex As Long
    For methodIndex = 1 To 2
    Dim direction As Long
    For direction = -1 To 1 Step 2
        Dim cap As CCapacitySolver
        Set cap = New CCapacitySolver
        ConfigureCapacity cap
        cap.ConcreteCompressionLimit = -1#
        cap.SteelCompressionLimit = 0.0003
        cap.SteelTensionLimit = 0.0018
        cap.LambdaTolerance = 0.000001
        cap.StrainTolerance = 0.000000001
        cap.SolverToleranceN = 0.01
        If methodIndex = 1 Then
            cap.SolveByLoadPathMultiplier section, LinearConcrete(), LinearSteel(), _
                0#, direction * Audit02AxialStiffness(section) * 0.0001, 0#, 0#, 0#, 0#
        Else
            cap.SolveByUltimateLoadPath section, LinearConcrete(), LinearSteel(), _
                0#, direction * Audit02AxialStiffness(section) * 0.0001, 0#, 0#, 0#, 0#, False
        End If
        Dim expectedLambda As Double
        If direction < 0 Then expectedLambda = 3# Else expectedLambda = 18#
        Dim prefix As String
        prefix = "audit02.steelSign." & CStr(methodIndex) & "." & CStr(direction)
        AssertTrue stats, prefix & ".converged", cap.Converged
        AssertClose stats, prefix & ".lambda", cap.LambdaUltimate, expectedLambda, 0.00001
        AssertEquals stats, prefix & ".criterion", cap.LimitState, "SteelStrainLimit"
    Next direction
    Next methodIndex
End Sub

' Проверяет реальную неудачу равновесия после нескольких итераций.
' Большие деформации пробной плоскости не должны создавать физический предел
' и BaseFail; основанием остаются машинные коды численной неудачи solver-а.
Private Sub TestAudit02UnconvergedProbeIsNumerical(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars
    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, rebars)
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.MaxRetries = 0
    cap.SolverBaseLoadSteps = 1
    cap.SolverMaxIterations = 3
    cap.ConcreteCompressionLimit = -0.000000001
    Dim path As CLoadPathVector
    Set path = New CLoadPathVector
    path.Initialize -1000000000#, 0#, 0#, 1000000#, 0#, 0#
    Dim functionValue As Double
    Dim probeState As String
    cap.LimitSearchEvaluateCapacityLoadMultiplier section, ProvisionalConcrete(), _
        ProvisionalSteel(), path, 0#, functionValue, probeState
    AssertTrue stats, "audit02.failedProbe.solverExists", Not cap.LastSolver Is Nothing
    If cap.LastSolver Is Nothing Then Exit Sub
    AssertTrue stats, "audit02.failedProbe.unconverged", Not cap.LastSolver.Converged
    AssertTrue stats, "audit02.failedProbe.multipleIterations", cap.LastSolver.Iterations > 1
    AssertTrue stats, "audit02.failedProbe.strainBeyondLimit", cap.LastSolver.MinConcreteStrain < -0.000000001
    AssertTrue stats, "audit02.failedProbe.numerical", cap.LimitSearchCapacityStateIsNumericalFailure(probeState)
    cap.LimitSearchHandleCapacityInitialFailure probeState
    Dim result As CLimitSearchResult
    Set result = Audit02CapacitySnapshot(cap, "LoadMultiplier")
    AssertTrue stats, "audit02.failedProbe.noBaseFail", result.Meta.InternalStatus = rsNumericalFailure
End Sub

' ДЛЯ ТЕСТОВ
' Несошедшаяся положительная lambda не подтверждает физическую верхнюю
' границу, даже когда последняя итерационная плоскость превысила предел.
Private Sub TestAudit02PositiveUnconvergedProbe(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars
    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, rebars)
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.MaxRetries = 0
    cap.SolverBaseLoadSteps = 1
    cap.SolverMaxIterations = 3
    cap.ConcreteCompressionLimit = -0.000000001
    Dim path As CLoadPathVector
    Set path = New CLoadPathVector
    path.Initialize 0#, -1000000000#, 0#, 1000000#, 0#, 0#
    Dim residual As Double
    Dim state As String
    cap.LimitSearchEvaluateCapacityLoadMultiplier section, ProvisionalConcrete(), _
        ProvisionalSteel(), path, 1#, residual, state
    AssertTrue stats, "audit02.failedPositive.noEquilibrium", Not cap.LastSolver.Converged
    AssertTrue stats, "audit02.failedPositive.multiple", cap.LastSolver.Iterations > 1
    AssertTrue stats, "audit02.failedPositive.exceeded", cap.LastSolver.MinConcreteStrain < -0.000000001
    AssertTrue stats, "audit02.failedPositive.numerical", cap.LimitSearchCapacityStateIsNumericalFailure(state)
    AssertTrue stats, "audit02.failedPositive.notLimit", Not cap.LimitSearchCapacityStateIsPhysicalLimit(state)
End Sub

' ДЛЯ ТЕСТОВ
' Проверяет Offset ниже, точно на и выше физического критерия на линейном
' материале. Точная начальная предельная точка уже исчерпала Capacity-путь,
' но сама оценка плоскости остается физической, без технического продолжения.
Private Sub TestAudit02InitialOffsetBoundary(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars
    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, rebars)
    Dim ratio As Variant
    For Each ratio In Array(0.5, 1#, 1.1)
        Dim cap As CCapacitySolver
        Set cap = New CCapacitySolver
        ConfigureCapacity cap
        cap.ConcreteCompressionLimit = -1#
        cap.SteelStrainLimit = 0.001
        cap.SolverToleranceN = 0.01
        Dim path As CLoadPathVector
        Set path = New CLoadPathVector
        path.Initialize CDbl(ratio) * Audit02AxialStiffness(section) * 0.001, _
            Audit02AxialStiffness(section) * 0.0001, 0#, 0#, 0#, 0#
        Dim residual As Double
        Dim state As String
        cap.LimitSearchEvaluateCapacityLoadMultiplier section, LinearConcrete(), LinearSteel(), _
            path, 0#, residual, state
        Dim prefix As String
        prefix = "audit02.offset." & CStr(ratio)
        AssertTrue stats, prefix & ".equilibrium", cap.LastSolver.Converged
        If CDbl(ratio) < 1# Then
            AssertTrue stats, prefix & ".acceptable", cap.LimitSearchCapacityStateIsAcceptable(state)
        ElseIf CDbl(ratio) = 1# Then
            ' Равновесие решается с прежним абсолютным допуском усилий.
            ' Уточненная плоскость может оказаться по любую сторону точной
            ' аналитической границы; классификация должна отражать именно ее,
            ' а не объявлять превышение по одному значению входной нагрузки.
            AssertClose stats, prefix & ".exactUtilization", cap.CriticalStrainUtilization, 1#, 0.000000000001
            AssertTrue stats, prefix & ".classification", _
                cap.LimitSearchCapacityStateIsPhysicalLimit(state) = (cap.CriticalStrainUtilization >= 1#)
            AssertTrue stats, prefix & ".notNumerical", Not cap.LimitSearchCapacityStateIsNumericalFailure(state)
        Else
            AssertTrue stats, prefix & ".limit", cap.LimitSearchCapacityStateIsPhysicalLimit(state)
            cap.LimitSearchHandleCapacityInitialFailure state
            Dim result As CLimitSearchResult
            Set result = Audit02CapacitySnapshot(cap, "LoadMultiplier")
            AssertTrue stats, prefix & ".code", result.Meta.ResultCode = rcInitialStateBeyondLimit
            Dim policy As CResultStatusPolicy
            Set policy = New CResultStatusPolicy
            AssertEquals stats, prefix & ".display", policy.ExternalStatus(result.Meta), "BaseFail"
        End If
    Next ratio
End Sub

' ДЛЯ ТЕСТОВ: общий Newton принимает только ILimitSearchProblem и не требует
' живых Capacity/Crack-калькуляторов. Аналитический корень проверяется отдельно
' от физических regression-моделей, их expected values не меняются.
Private Sub TestAudit02GenericUltimateSearch(ByRef stats As TCapacityTestStats)
    Dim problem As CTestLimitSearchProblem
    Set problem = New CTestLimitSearchProblem
    problem.Configure "Bisection", 1.25, 0.000000001, 20
    Dim loadPath As CLoadPathVector
    Set loadPath = New CLoadPathVector
    loadPath.Initialize 0#, 1#, 0#, 0#, 0#, 0#
    Dim search As CUltimateStrainSearch
    Set search = New CUltimateStrainSearch
    AssertTrue stats, "audit02.genericUltimate.succeeded", _
        search.RunNewton(problem, Nothing, Nothing, Nothing, loadPath, Nothing, "Тестовый поиск не сошелся.")
    AssertClose stats, "audit02.genericUltimate.lambda", problem.FinalLambda, 1.25, 0.000000001
    AssertEquals stats, "audit02.genericUltimate.method", problem.FinalMethod, "UltimateStrain"
End Sub

' Возвращает точную осевую жесткость тестового линейного сечения, Н.
' Бетонная сетка включает площадь стержней, поэтому арматура добавляет
' разность Es-Eb, как и интегрирование напряжений в CSectionSolver.
Private Function Audit02AxialStiffness(ByVal section As CSectionModel) As Double
    Dim i As Long
    For i = 1 To section.ConcreteCount
        Audit02AxialStiffness = Audit02AxialStiffness + 32500# * section.ConcreteArea(i)
    Next i
    For i = 1 To section.RebarCount
        Audit02AxialStiffness = Audit02AxialStiffness + (200000# - 32500#) * section.RebarArea(i)
    Next i
End Function

' Проверяет, что найденный предел является успехом Search, а непрохождение
' текущего сочетания определяет только Capacity. Числа старого сценария и
' его допуски не изменяются; отдельно проверяется изоляция выданной meta.
Private Sub CheckAudit02SearchVsCapacity(ByRef stats As TCapacityTestStats, _
        ByVal cap As CCapacitySolver, ByVal section As CSectionModel, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double)
    Dim spec As CMaterialModelSpec
    Set spec = New CMaterialModelSpec
    spec.Initialize "ULS(I)", "TwoLine", "Ignore", "TwoLine"
    Dim search As CLimitSearchResult
    Set search = Audit02CapacitySnapshot(cap, "LoadMultiplier", section, _
        ProvisionalConcrete(), ProvisionalSteel(), spec)
    AssertTrue stats, "audit02.searchEngineering.searchSuccess", search.Meta.InternalStatus = rsSuccess
    AssertTrue stats, "audit02.searchEngineering.searchCode", search.Meta.ResultCode = rcCheckPassed
    AssertTrue stats, "audit02.searchEngineering.searchSucceeded", search.Succeeded
    Dim loadState As CSectionLoadState
    Set loadState = New CSectionLoadState
    loadState.Initialize targetN, targetMx, targetMy, 0#, 0#
    Dim path As CLoadPathDescriptor
    Set path = New CLoadPathDescriptor
    path.InitializeFromLoadState "lambda*Mxy", loadState
    Dim result As CCapacityResult
    Set result = New CCapacityResult
    result.InitializeFromSearch search, path, False, 0#, 0#, Nothing, 300#, 200#
    AssertTrue stats, "audit02.searchEngineering.capacityFailed", result.ResultMeta.InternalStatus = rsCheckFailed
    AssertEquals stats, "audit02.searchEngineering.capacityDisplay", result.Status, "FAIL"
    AssertClose stats, "audit02.searchEngineering.lambdaUnchanged", result.LambdaCapacity, cap.LambdaUltimate, 0#
    Dim detachedMeta As CResultMeta
    Set detachedMeta = search.Meta
    detachedMeta.SetResult rsInternalError, rcInternalError, rkCapacity, "Изменение копии в тесте."
    AssertTrue stats, "audit02.searchEngineering.metaSnapshot", search.Meta.InternalStatus = rsSuccess
    AssertTrue stats, "audit02.searchEngineering.pointState", Not search.PointState Is Nothing
    If Not search.PointState Is Nothing Then
        AssertTrue stats, "audit02.searchEngineering.sharedState", result.StateResult Is search.PointState
        Dim savedN As Double, savedEps As Double
        savedN = search.PointState.Nint
        savedEps = search.PointState.Epsilon0
        cap.LastSolver.EvaluateStrainPlane section, ProvisionalConcrete(), ProvisionalSteel(), 0#, 0#, 0#
        AssertClose stats, "audit02.searchEngineering.solverMutation.N", search.NUltimate, savedN, 0#
        AssertClose stats, "audit02.searchEngineering.solverMutation.plane", search.PointState.Epsilon0, savedEps, 0#
    End If
    Dim detachedSearch As CLimitSearchResult
    Set detachedSearch = result.SearchResult
    detachedSearch.InitializeInvalidConfiguration "Unknown", "Изменение выданной копии Search."
    AssertTrue stats, "audit03.searchEngineering.outputSnapshot", result.SearchResult.HasLimitPoint
    AssertClose stats, "audit03.searchEngineering.outputLambda", result.SearchResult.LambdaUltimate, cap.LambdaUltimate, 0#
    search.InitializeInvalidConfiguration "Unknown", "Повторное заполнение исходного Search."
    AssertTrue stats, "audit03.searchEngineering.inputSnapshot", result.SearchResult.HasLimitPoint
    AssertTrue stats, "audit03.searchEngineering.inputMeta", result.SearchResult.Meta.InternalStatus = rsSuccess
    AssertClose stats, "audit03.searchEngineering.inputLambda", result.SearchResult.LambdaUltimate, cap.LambdaUltimate, 0#
    AssertEquals stats, "audit03.searchEngineering.capacityStatusPreserved", result.Status, "FAIL"
End Sub

' ДЛЯ ТЕСТОВ: собирает capacity-снимок через тот же доменный адаптер, что
' production Search. Сам общий result не принимает инженерные калькуляторы.
Private Function Audit02CapacitySnapshot(ByVal cap As CCapacitySolver, ByVal strategy As String, _
        Optional ByVal section As CSectionModel = Nothing, _
        Optional ByVal concrete As Object = Nothing, Optional ByVal steel As Object = Nothing, _
        Optional ByVal spec As CMaterialModelSpec = Nothing) As CLimitSearchResult
    Dim problem As CCapacityLimitSearchProblem
    Set problem = New CCapacityLimitSearchProblem
    problem.Initialize cap
    Set Audit02CapacitySnapshot = problem.BuildSnapshot(strategy, section, concrete, steel, spec)
End Function

' ДЛЯ ТЕСТОВ: один generic bracket/recovery работает с обоими видами задачи
' без железобетонных калькуляторов. Проверяются все одномерные методы,
' край MaxLambda, отсутствие точки за границей и повторное заполнение result.
Private Sub TestAudit02GenericLoadMultiplierMatrix(ByRef stats As TCapacityTestStats)
    Dim kind As Variant, methodName As Variant, root As Variant
    For Each kind In Array(rkCapacity, rkCrackFormation)
        For Each methodName In Array("Bisection", "Brent", "Secant")
            For Each root In Array(0.75, 5.25, 6#, 8#)
                Dim problem As CTestLimitSearchProblem
                Set problem = New CTestLimitSearchProblem
                problem.Configure CStr(methodName), CDbl(root), 0.000001, 80, 1#, 6#, CLng(kind)
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
                    AssertTrue stats, prefix & ".success", result.Succeeded
                    AssertTrue stats, prefix & ".point", result.HasLimitPoint
                    AssertClose stats, prefix & ".lambda", result.LambdaUltimate, CDbl(root), 0.000001
                    AssertClose stats, prefix & ".loads", result.NUltimate, result.LambdaUltimate, 0#
                    AssertEquals stats, prefix & ".method", result.ActualMethod, CStr(methodName)
                Else
                    AssertTrue stats, prefix & ".boundCode", result.Meta.ResultCode = rcSearchBoundReached
                    AssertTrue stats, prefix & ".noPoint", Not result.HasLimitPoint
                    AssertClose stats, prefix & ".noCapacity", result.LambdaUltimate, 0#, 0#
                    AssertClose stats, prefix & ".maxProbed", problem.ProbeMaxLambda, 6#, 0#
                End If
                result.InitializeInvalidConfiguration "Unknown", "Проверка очистки", CLng(kind)
                AssertTrue stats, prefix & ".resetPoint", Not result.HasLimitPoint
                AssertTrue stats, prefix & ".resetState", result.PointState Is Nothing
                AssertTrue stats, prefix & ".resetExecution", Not result.SearchExecuted
                AssertClose stats, prefix & ".resetLoads", result.NUltimate, 0#, 0#
            Next root
        Next methodName
        Set problem = New CTestLimitSearchProblem
        problem.Configure "Bisection", 1.1, 0.000001, 80, 1#, 6#, CLng(kind)
        problem.FailureAbove = 1.2
        Set request = New CLimitSearchRequest
        request.InitializeWithProblem problem, CLng(kind), "LoadMultiplier", 0#, 1#, 0#, 0#, 0#, 0#
        Set result = search.Execute(request)
        AssertTrue stats, "audit02.genericMultiplier.recovery." & CStr(kind) & ".point", result.HasLimitPoint
        AssertClose stats, "audit02.genericMultiplier.recovery." & CStr(kind) & ".lambda", result.LambdaUltimate, 1.1, 0.000001
        AssertTrue stats, "audit02.genericMultiplier.recovery." & CStr(kind) & ".upperFailed", problem.ProbeMaxLambda > 1.2
    Next kind

    Set request = New CLimitSearchRequest
    Set result = search.Execute(request)
    AssertTrue stats, "audit02.genericMultiplier.missingProblem.internal", result.Meta.InternalStatus = rsInternalError
    AssertTrue stats, "audit02.genericMultiplier.missingProblem.noPoint", Not result.HasLimitPoint
End Sub

' ============================== ДЛЯ ТЕСТОВ ==============================
' Отдельный entrypoint для watchdog: соседние Double и нулевой bisection-budget
' на неисправном алгоритме зависают, а не возвращают правдоподобный предел.
Public Function RunAudit03SearchStagnation() As String
    Dim stats As TCapacityTestStats
    TestAudit03SearchArithmeticCase stats, "Bisection", rkCapacity, 1
    AppendLine stats, "TOTAL_AUDIT03_STAGNATION: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03SearchStagnation = stats.Report
End Function

' ДЛЯ ТЕСТОВ: отдельный полный прогон арифметических краев generic Search.
Public Function RunAudit03SearchTests() As String
    Dim stats As TCapacityTestStats
    TestAudit03SearchArithmetic stats
    AppendLine stats, "TOTAL_AUDIT03_SEARCH: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03SearchTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ: проверяет настоящий Capacity с неограниченным числом делений
' скобки и terminal typed-errors через production адаптер, не ручную meta.
Public Function RunAudit03CapacityContracts() As String
    Dim stats As TCapacityTestStats
    TestAudit03RealCapacityPrecision stats
    TestAudit03CapacityTypedFailures stats
    AppendLine stats, "TOTAL_AUDIT03_CAPACITY_CONTRACTS: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03CapacityContracts = stats.Report
End Function

' ДЛЯ ТЕСТОВ: линейные материалы исключают неудачу внутреннего равновесия.
' При допуске меньше шага Double поиск обязан закончиться с честной причиной,
' а локальный cache не должен объединять разные lambda и создавать ложный предел.
Private Sub TestAudit03RealCapacityPrecision(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder, rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars
    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, rebars)
    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SearchMethod = "Bisection"
    cap.LambdaTolerance = 0.000000000000000001
    cap.MaxRetries = 0
    cap.DiagnosticsEnabled = True
    SaveAudit03SearchProgress stats, "START: audit03.capacity.realPrecision"
    cap.SolveByLoadPathMultiplier section, LinearConcrete(), LinearSteel(), 0#, 0#, 0#, 10000000#, 0#, 0#
    Dim result As CLimitSearchResult
    Set result = Audit02CapacitySnapshot(cap, "LoadMultiplier")
    AssertTrue stats, "audit03.capacity.realPrecision.numerical", result.Meta.InternalStatus = rsNumericalFailure
    AssertTrue stats, "audit03.capacity.realPrecision.code", result.Meta.ResultCode = rcNumericalFailure
    AssertTrue stats, "audit03.capacity.realPrecision.noPoint", Not result.HasLimitPoint
    AssertTrue stats, "audit03.capacity.realPrecision.notBound", Not cap.SearchBoundReached
    AssertTrue stats, "audit03.capacity.realPrecision.finiteWork", cap.Iterations > 40 And cap.Iterations < 100
    AssertTrue stats, "audit03.capacity.realPrecision.calculated", result.Meta.Calculated
    AssertTrue stats, "audit03.capacity.realPrecision.reason", Len(result.Meta.ResultComment) > 0
    AppendLine stats, "COMMENT: audit03.capacity.realPrecision; " & result.Meta.ResultComment
    AppendLine stats, "DIAGNOSTIC: " & cap.DiagnosticLog
    SaveAudit03SearchProgress stats, "DONE: audit03.capacity.realPrecision"
End Sub

' ДЛЯ ТЕСТОВ: configuration/input/internal/singular причины проходят от
' реального Capacity solve в Search-result без retries и назначения BaseFail
' программной ошибке. Повторный запуск не наследует meta предыдущей задачи.
Private Sub TestAudit03CapacityTypedFailures(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder, rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars
    Dim validSection As CSectionModel
    Set validSection = BuildGeneratedSectionModel(mesh, rebars)
    Dim scenario As Long
    For scenario = 1 To 9
        Dim cap As CCapacitySolver
        Set cap = New CCapacitySolver
        ConfigureCapacity cap
        cap.SearchMethod = "Bisection"
        cap.DiagnosticsEnabled = True
        Dim section As CSectionModel
        Set section = validSection
        Dim concrete As Object, steel As Object
        Set concrete = LinearConcrete()
        Set steel = LinearSteel()
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
                If scenario = 7 Then cap.SolverMinLineSearchAlpha = 0#
                If scenario = 8 Then cap.SolverDampingInitial = 0#
                If scenario = 9 Then cap.SolverMinLineSearchAlpha = 2#
                expectedStatus = rsInvalidConfiguration: expectedCode = rcInvalidConfiguration
                expectedFailure = sfcInvalidConfiguration
        End Select
        Dim prefix As String
        prefix = "audit03.capacity.typed." & CStr(scenario)
        SaveAudit03SearchProgress stats, "START: " & prefix
        If scenario <= 6 Then
            cap.SolveByLoadPathMultiplier section, concrete, steel, nOffset, 0#, 0#, 10000000#, 0#, 0#
        Else
            cap.SolveByUltimateLoadPath section, concrete, steel, nOffset, 0#, 0#, 10000000#, 0#, 0#
        End If
        Dim result As CLimitSearchResult
        Set result = Audit02CapacitySnapshot(cap, "LoadMultiplier")
        AssertTrue stats, prefix & ".status", result.Meta.InternalStatus = expectedStatus
        AssertTrue stats, prefix & ".code", result.Meta.ResultCode = expectedCode
        AssertTrue stats, prefix & ".failure", cap.FailureCode = expectedFailure
        AssertTrue stats, prefix & ".noRetry", cap.RetryCount = 0
        AssertTrue stats, prefix & ".noPoint", Not result.HasLimitPoint
        AssertTrue stats, prefix & ".notBaseFail", result.Meta.ResultCode <> rcInitialStateBeyondLimit
        AssertTrue stats, prefix & ".reason", Len(result.Meta.ResultComment) > 0
        If scenario <> 6 Then AssertTrue stats, prefix & ".notCalculated", Not result.Meta.Calculated
        AppendLine stats, "COMMENT: " & prefix & "; " & result.Meta.ResultComment
        SaveAudit03SearchProgress stats, "DONE: " & prefix
    Next scenario
End Sub

' ДЛЯ ТЕСТОВ: отдельный watchdog воспроизводит бесконечный line search
' прежнего общего Newton при нулевом минимальном alpha и неизменной норме.
Public Function RunAudit03UltimateStagnation() As String
    Dim stats As TCapacityTestStats
    TestAudit03UltimateGuardCase stats, rkCapacity, 1
    AppendLine stats, "TOTAL_AUDIT03_ULTIMATE_STAGNATION: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03UltimateStagnation = stats.Report
End Function

' ДЛЯ ТЕСТОВ: конечность, typed failures и реальные Capacity-маршруты.
Public Function RunAudit03UltimateContracts() As String
    Dim stats As TCapacityTestStats
    TestAudit03UltimateGuards stats
    TestAudit03CapacityTypedFailures stats
    AppendLine stats, "TOTAL_AUDIT03_ULTIMATE_CONTRACTS: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03UltimateContracts = stats.Report
End Function

' ДЛЯ ТЕСТОВ: оба domain-kind используют один общий Newton, без выдуманных
' материалов. Критерий и исходная точность линейной задачи не ослабляются.
Private Sub TestAudit03UltimateGuards(ByRef stats As TCapacityTestStats)
    Dim kind As Variant, scenario As Long
    For Each kind In Array(rkCapacity, rkCrackFormation)
        For scenario = 1 To 13
            TestAudit03UltimateGuardCase stats, CLng(kind), scenario
        Next scenario
    Next kind
    TestAudit02GenericUltimateSearch stats
End Sub

' ДЛЯ ТЕСТОВ: проверяет реальный callback call path, отсутствие фиктивной
' точки, точную причину и конечное число проб. Nothing context - нарушение API.
Private Sub TestAudit03UltimateGuardCase(ByRef stats As TCapacityTestStats, _
        ByVal kind As EResultKind, ByVal scenario As Long)
    On Error GoTo Failed
    Dim prefix As String
    prefix = "audit03.ultimate." & CStr(kind) & "." & CStr(scenario)
    SaveAudit03SearchProgress stats, "START: " & prefix
    Dim problem As CTestLimitSearchProblem
    Set problem = New CTestLimitSearchProblem
    problem.ConfigureUltimateGuardCase scenario, kind
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
        AssertTrue stats, prefix & ".controlledContractError", errorNumber = vbObjectError + 4212
        SaveAudit03SearchProgress stats, "DONE: " & prefix
        Exit Sub
    End If
    If scenario = 13 Then Set path = Nothing
    AssertTrue stats, prefix & ".notAccepted", _
        Not search.RunNewton(problem, Nothing, Nothing, Nothing, path, Nothing, "Тестовый поиск не сошелся.")
    Dim expectedFailure As ESolverFailureCode
    expectedFailure = sfcInvalidConfiguration
    Select Case scenario
        Case 6, 11: expectedFailure = sfcNumericalFailure
        Case 8, 10, 13: expectedFailure = sfcInternalError
        Case 9: expectedFailure = sfcSingularTangent
    End Select
    AssertTrue stats, prefix & ".typedCause", problem.FailureCode = expectedFailure
    AssertTrue stats, prefix & ".noPoint", Not problem.Converged
    AssertTrue stats, prefix & ".finiteProbes", problem.UltimateCalls <= 4
    If scenario <= 5 Or scenario = 13 Then AssertTrue stats, prefix & ".noNumericalAttempt", problem.UltimateCalls = 0
    If scenario = 6 Then AssertTrue stats, prefix & ".stagnationReason", InStr(problem.DiagnosticLog, "представимого шага Double") > 0
    If scenario >= 7 And scenario <= 11 Then AssertTrue stats, prefix & ".failedProbeNotOverwritten", problem.UltimateCalls = 2
    Dim callback As ILimitSearchProblem
    Set callback = problem
    Dim request As CLimitSearchRequest
    Set request = New CLimitSearchRequest
    request.InitializeWithProblem problem, kind, "UltimateStrain", 0#, 1#, 0#, 0#, 0#, 0#
    Dim result As CLimitSearchResult
    Set result = callback.BuildResult(request, False)
    Dim expectedMeta As CResultMeta
    Set expectedMeta = New CResultMeta
    expectedMeta.SetSolverFailure expectedFailure, kind, "Контроль typed-причины.", (problem.UltimateCalls > 0)
    AssertTrue stats, prefix & ".status", result.Meta.InternalStatus = expectedMeta.InternalStatus
    AssertTrue stats, prefix & ".code", result.Meta.ResultCode = expectedMeta.ResultCode
    AssertTrue stats, prefix & ".resultKind", result.Meta.ResultKind = kind
    AssertTrue stats, prefix & ".noPhysicalPoint", Not result.HasLimitPoint
    AssertTrue stats, prefix & ".lifecycle", result.Meta.Calculated = expectedMeta.Calculated
    AssertTrue stats, prefix & ".reason", Len(result.Meta.ResultComment) > 0
    AppendLine stats, "COMMENT: " & prefix & "; calls=" & CStr(problem.UltimateCalls) & "; " & result.Meta.ResultComment
    SaveAudit03SearchProgress stats, "DONE: " & prefix
    Exit Sub
Failed:
    AssertTrue stats, prefix & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
    SaveAudit03SearchProgress stats, "FAILED: " & prefix
End Sub

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

' Реальная generic-задача различает невозможную точность, техническую границу,
' точный предел на MaxLambda и ошибочный recovery без фальшивого State.
Private Sub TestAudit03SearchArithmeticCase(ByRef stats As TCapacityTestStats, _
        ByVal methodName As String, ByVal kind As EResultKind, ByVal scenario As Long)
    On Error GoTo Failed
    Dim problem As CTestLimitSearchProblem
    Set problem = New CTestLimitSearchProblem
    Dim prefix As String
    prefix = "audit03.search." & CStr(kind) & "." & methodName & "." & CStr(scenario)
    SaveAudit03SearchProgress stats, "START: " & prefix
    Dim adjacent As Double
    adjacent = 1# + 2# ^ (-52)
    Select Case scenario
        Case 1
            AssertTrue stats, prefix & ".adjacentDistinct", adjacent > 1# And adjacent - 1# < 0.000000000000001
            problem.Configure methodName, adjacent, 0.000000000000000001, 80, 1#, adjacent, kind
            If methodName = "Bisection" Then problem.Configure methodName, adjacent, 0.000000000000000001, 0, 1#, adjacent, kind
            problem.StepCriterion = True
        Case 2
            problem.Configure methodName, 1.2E+308, 1E+294, 120, 1E+308, 1.4E+308, kind
        Case 3
            problem.Configure methodName, 3#, 0.000001, 80, 1#, 6#, kind
            problem.FailureAbove = 0.5
            problem.RecoveryWithoutProgress = True
        Case 4
            problem.Configure methodName, 6#, 0.000001, 80, 1#, 6#, kind
    End Select
    Dim request As CLimitSearchRequest
    Set request = New CLimitSearchRequest
    request.InitializeWithProblem problem, kind, "LoadMultiplier", 0#, 1#, 0#, 0#, 0#, 0#
    Dim search As CLoadMultiplierSearch
    Set search = New CLoadMultiplierSearch
    Dim result As CLimitSearchResult
    Set result = search.Execute(request)
    If scenario = 1 Or scenario = 3 Then
        AssertTrue stats, prefix & ".numerical", result.Meta.InternalStatus = rsNumericalFailure
        AssertTrue stats, prefix & ".notBound", result.Meta.ResultCode <> rcSearchBoundReached
        AssertTrue stats, prefix & ".noPoint", Not result.HasLimitPoint
        AssertTrue stats, prefix & ".noFakeLimit", result.LambdaUltimate = 0#
        AssertTrue stats, prefix & ".finiteWork", problem.ProbeCalls <= 120
        AssertTrue stats, prefix & ".reason", Len(result.DiagnosticLog) > 0
    Else
        AssertTrue stats, prefix & ".success", result.Succeeded And result.HasLimitPoint
        If scenario = 2 Then
            AssertTrue stats, prefix & ".accuracy", Abs(result.LambdaUltimate / 1.2E+308 - 1#) <= 0.00000000000002
            AssertTrue stats, prefix & ".bounded", problem.ProbeMaxLambda <= 1.4E+308
        Else
            AssertClose stats, prefix & ".maxLambda", result.LambdaUltimate, 6#, 0.000001
            AssertClose stats, prefix & ".lastProbe", problem.ProbeMaxLambda, 6#, 0#
        End If
    End If
    AppendLine stats, "COMMENT: " & prefix & "; " & result.Meta.ResultComment & "; " & result.DiagnosticLog
    SaveAudit03SearchProgress stats, "DONE: " & prefix
    Exit Sub
Failed:
    AssertTrue stats, prefix & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
    SaveAudit03SearchProgress stats, "FAILED: " & prefix
End Sub

' ДЛЯ ТЕСТОВ: сохраняет уже выполненные assertions до следующего COM-шага.
' При модальной ошибке или watchdog виден конкретный случай, а не только suite.
Private Sub SaveAudit03SearchProgress(ByRef stats As TCapacityTestStats, ByVal stage As String)
    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open ThisWorkbook.Path & "\Audit03_Search_Progress.txt" For Output As #fileNumber
    Print #fileNumber, stats.Report
    Print #fileNumber, stage
    Close #fileNumber
End Sub

' ============================== ДЛЯ ТЕСТОВ ==============================
' ДЛЯ ТЕСТОВ: проверяет реальный lifecycle Capacity/Search, в том числе снимок
' до запуска и reset одного solver-а после успешной численной попытки.
Public Function RunAudit03CapacityLifecycle() As String
    On Error GoTo Failed
    Dim stats As TCapacityTestStats
    TestAudit03CapacitySearchLifecycle stats
    AppendLine stats, "TOTAL_AUDIT03_CAPACITY_LIFECYCLE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03CapacityLifecycle = stats.Report
    Exit Function
Failed:
    RunAudit03CapacityLifecycle = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' ДЛЯ ТЕСТОВ: оба численных метода действительно ищут предел, а невалидные
' допуски/бюджет и отсутствующая зависимость останавливаются до Search.
' При повторе старый результат должен сохранять независимый факт выполнения.
Private Sub TestAudit03CapacitySearchLifecycle(ByRef stats As TCapacityTestStats)
    Dim problem As CCapacityLimitSearchProblem, result As CLimitSearchResult
    Set problem = New CCapacityLimitSearchProblem
    Set result = problem.BuildSnapshot("Auto")
    AssertTrue stats, "audit03.capacity.lifecycle.missing.notExecuted", Not result.SearchExecuted
    AssertTrue stats, "audit03.capacity.lifecycle.missing.notCalculated", Not result.Meta.Calculated
    Dim mesh As CFiberMeshBuilder, rebars As CRebarLayout, section As CSectionModel
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars
    Set section = BuildGeneratedSectionModel(mesh, rebars)
    Dim cap As CCapacitySolver, retained As CLimitSearchResult, strategy As Variant
    Set cap = New CCapacitySolver
    Set result = Audit02CapacitySnapshot(cap, "Auto")
    AssertTrue stats, "audit03.capacity.lifecycle.empty.notExecuted", Not result.SearchExecuted
    AssertTrue stats, "audit03.capacity.lifecycle.empty.notCalculated", Not result.Meta.Calculated
    For Each strategy In Array("LoadMultiplier", "UltimateStrain")
        ConfigureCapacity cap
        If CStr(strategy) = "LoadMultiplier" Then
            cap.SolveByLoadPathMultiplier section, LinearConcrete(), LinearSteel(), 0#, 0#, 0#, 10000000#, 0#, 0#
        Else
            cap.SolveByUltimateLoadPath section, LinearConcrete(), LinearSteel(), 0#, 0#, 0#, 10000000#, 0#, 0#
        End If
        Set result = Audit02CapacitySnapshot(cap, CStr(strategy))
        Set retained = result.Clone
        Dim prefix As String
        prefix = "audit03.capacity.lifecycle." & CStr(strategy)
        AssertTrue stats, prefix & ".attempted", result.SearchExecuted
        AssertTrue stats, prefix & ".calculated", result.Meta.Calculated
        AssertTrue stats, prefix & ".point", result.HasLimitPoint
        If CStr(strategy) = "LoadMultiplier" Then
            cap.LambdaTolerance = 0#
            cap.SolveByLoadPathMultiplier section, LinearConcrete(), LinearSteel(), 0#, 0#, 0#, 10000000#, 0#, 0#
        Else
            cap.SolverMaxIterations = 0
            cap.SolveByUltimateLoadPath section, LinearConcrete(), LinearSteel(), 0#, 0#, 0#, 10000000#, 0#, 0#
        End If
        Set result = Audit02CapacitySnapshot(cap, CStr(strategy))
        AssertTrue stats, prefix & ".invalidBeforeSearch", Not result.SearchExecuted
        AssertTrue stats, prefix & ".invalidNotCalculated", Not result.Meta.Calculated
        AssertTrue stats, prefix & ".invalidNoPoint", Not result.HasLimitPoint
        AssertTrue stats, prefix & ".priorSnapshot", retained.SearchExecuted And retained.HasLimitPoint
        cap.SolveByLoadPathMultiplier Nothing, LinearConcrete(), LinearSteel(), 0#, 0#, 0#, 10000000#, 0#, 0#
        Set result = Audit02CapacitySnapshot(cap, CStr(strategy))
        AssertTrue stats, prefix & ".missingReset", Not result.SearchExecuted
        AssertTrue stats, prefix & ".missingNoPoint", Not result.HasLimitPoint
    Next strategy
    ' UltimateStrain начинает реальную попытку, а настройки резервного
    ' LoadMultiplier ошибочны. Очистка второй ветви не должна стирать этот факт.
    ConfigureCapacity cap
    cap.SolverMaxIterations = 1
    cap.LambdaTolerance = 0#
    Dim request As CLimitSearchRequest, coordinator As CLimitSearchCoordinator
    Set request = New CLimitSearchRequest: Set coordinator = New CLimitSearchCoordinator
    request.InitializeCapacity cap, section, LinearConcrete(), LinearSteel(), _
        "Auto", -20000#, 0#, 0#, 10000000#, 0#, 0#, False, False
    Set result = coordinator.ExecuteCapacity(request)
    AssertTrue stats, "audit03.capacity.lifecycle.fallback.firstAttemptKept", result.SearchExecuted
    AssertTrue stats, "audit03.capacity.lifecycle.fallback.finalConfig", result.Meta.InternalStatus = rsInvalidConfiguration
    AssertTrue stats, "audit03.capacity.lifecycle.fallback.noPoint", Not result.HasLimitPoint
    AssertTrue stats, "audit03.capacity.lifecycle.fallback.finalNoProbe", cap.Iterations = 0
    ' Прямой содержательный API должен сохранять тот же lifecycle, что coordinator.
    ConfigureCapacity cap
    cap.SolverMaxIterations = 1
    cap.LambdaTolerance = 0#
    cap.SolveByAutoLoadPath section, LinearConcrete(), LinearSteel(), _
        -20000#, 0#, 0#, 10000000#, 0#, 0#, False, False
    Set result = Audit02CapacitySnapshot(cap, "Auto")
    AssertTrue stats, "audit03.capacity.lifecycle.directAuto.firstAttemptKept", result.SearchExecuted
    AssertTrue stats, "audit03.capacity.lifecycle.directAuto.finalConfig", result.Meta.InternalStatus = rsInvalidConfiguration
    AssertTrue stats, "audit03.capacity.lifecycle.directAuto.noPoint", Not result.HasLimitPoint
    AssertTrue stats, "audit03.capacity.lifecycle.directAuto.finalNoProbe", cap.Iterations = 0
    ' Текущее сочетание осевое относительно расчетного центра, но содержит
    ' взаимно компенсирующиеся пользовательский момент и момент от N.
    ' Во всех трех Auto-путях остаются моментные компоненты траектории,
    ' хотя при lambda=1 они компенсированы. Бюджет одной итерации дает
    ' честную численную неудачу; итог обязан помнить выполненный поиск.
    ConfigureCapacity cap
    cap.SolverMaxIterations = 1
    cap.LambdaTolerance = 0#
    Dim autoLoad As CSectionLoadState
    Set autoLoad = New CSectionLoadState
    autoLoad.Initialize -20000#, 10000000#, 0#, 0#, 500#
    request.InitializeCapacity cap, section, LinearConcrete(), LinearSteel(), _
        "UltimateStrain", 0#, 0#, 0#, 1#, 0#, 0#, False, False
    request.ConfigureLoadPathSelection "Auto", autoLoad, True
    Set result = coordinator.ExecuteCapacity(request)
    AppendLine stats, "AUTO_PATH_LIFECYCLE: path=" & request.SelectedLoadPath.Key & _
        "; status=" & CStr(result.Meta.InternalStatus) & "; code=" & CStr(result.Meta.ResultCode) & _
        "; searched=" & CStr(result.SearchExecuted) & "; comment=" & result.Meta.ResultComment
    AppendLine stats, "AUTO_PATH_LIFECYCLE_DIAGNOSTICS: " & result.DiagnosticLog
    AssertTrue stats, "audit03.capacity.lifecycle.autoPaths.firstAttemptKept", result.SearchExecuted
    AssertTrue stats, "audit03.capacity.lifecycle.autoPaths.finalNumerical", result.Meta.InternalStatus = rsNumericalFailure
    AssertTrue stats, "audit03.capacity.lifecycle.autoPaths.finalPathNMxy", request.SelectedLoadPath.Key = "LambdaNMxy"
    AssertTrue stats, "audit03.capacity.lifecycle.autoPaths.noPoint", Not result.HasLimitPoint
    Dim fake As CTestLimitSearchProblem, multiplier As CLoadMultiplierSearch
    Set fake = New CTestLimitSearchProblem: Set multiplier = New CLoadMultiplierSearch
    fake.Configure "Bisection", 1.25, 0.000001, 80
    request.InitializeWithProblem fake, rkCapacity, "LoadMultiplier", 0#, 1#, 0#, 0#, 0#, 0#
    Set result = multiplier.Execute(request)
    AssertTrue stats, "audit03.capacity.lifecycle.generic.executed", result.SearchExecuted
    fake.Configure "Bisection", 1.25, 0#, 80
    Set result = multiplier.Execute(request)
    AssertTrue stats, "audit03.capacity.lifecycle.generic.invalidNotExecuted", Not result.SearchExecuted
    AssertTrue stats, "audit03.capacity.lifecycle.generic.invalidNoProbe", fake.ProbeCalls = 0
    Set result = New CLimitSearchResult
    result.Initialize Nothing, "Auto", vbNullString, False, 0#, 0#, 0#, 0#, Nothing, vbNullString, vbNullString
    AssertTrue stats, "audit03.capacity.lifecycle.default.notExecuted", Not result.SearchExecuted
    AssertTrue stats, "audit03.capacity.lifecycle.default.notCalculated", Not result.Meta.Calculated
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
    AssertTrue stats, "audit03.capacity.lifecycle.report.reason", InStr(1, earlyReport, capacityResult.ResultMeta.ResultComment, vbBinaryCompare) > 0
    AssertTrue stats, "audit03.capacity.lifecycle.report.russian", InStr(1, earlyReport, "Поиск несущей способности", vbBinaryCompare) = 1
    AssertTrue stats, "audit03.capacity.lifecycle.report.noRawStatus", InStr(1, earlyReport, "InvalidInput", vbTextCompare) = 0
    TestAudit03CapacityConfigurationMessages stats, section
End Sub

' ДЛЯ ТЕСТОВ: каждый невалидный поисковый параметр должен сохранить typed
' причину, назвать свою строку Config и завершиться до первой численной пробы.
Private Sub TestAudit03CapacityConfigurationMessages(ByRef stats As TCapacityTestStats, _
        ByVal section As CSectionModel)
    Dim key As Variant, cap As CCapacitySolver, result As CLimitSearchResult
    Dim prefix As String, strategy As String
    For Each key In Array("Capacity.InitialLambda", "Capacity.ToleranceLambda", _
            "Capacity.MaxLambda", "Capacity.MaxRetries", "Capacity.BaseLoadSteps", _
            "Capacity.ToleranceStrain")
        Set cap = New CCapacitySolver
        ConfigureCapacity cap
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
            cap.SolveByUltimateLoadPath section, LinearConcrete(), LinearSteel(), 0#, 0#, 0#, 10000000#, 0#, 0#
        Else
            cap.SolveByLoadPathMultiplier section, LinearConcrete(), LinearSteel(), 0#, 0#, 0#, 10000000#, 0#, 0#
        End If
        Set result = Audit02CapacitySnapshot(cap, strategy)
        prefix = "audit03.capacity.configMessage." & CStr(key)
        AssertTrue stats, prefix & ".status", result.Meta.InternalStatus = rsInvalidConfiguration
        AssertTrue stats, prefix & ".code", result.Meta.ResultCode = rcInvalidConfiguration
        AssertTrue stats, prefix & ".sheet", InStr(1, result.Meta.ResultComment, "Config", vbBinaryCompare) > 0
        AssertTrue stats, prefix & ".key", InStr(1, result.Meta.ResultComment, CStr(key), vbBinaryCompare) > 0
        AssertTrue stats, prefix & ".noSearch", Not result.SearchExecuted And cap.Iterations = 0
        AppendLine stats, "CAPACITY_CONFIG_MESSAGE: " & CStr(key) & "|" & result.Meta.ResultComment
    Next key
End Sub

' ============================== ДЛЯ ТЕСТОВ ==============================
' Проверяет фактический общий LoadMultiplier для обоих инженерных видов и
' всех трех методов. Исключение callback-а не должно выходить из Execute,
' а терминальная причина запрещает последующий recovery/finalization.
Public Function RunAudit03MultiplierTypedFaults() As String
    On Error GoTo Failed
    Dim stats As TCapacityTestStats
    TestAudit03MultiplierTypedFaults stats
    AppendLine stats, "TOTAL_AUDIT03_MULTIPLIER_TYPED: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03MultiplierTypedFaults = stats.Report
    Exit Function
Failed:
    RunAudit03MultiplierTypedFaults = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' ДЛЯ ТЕСТОВ: каждый направленный отказ имеет отдельный expected enum,
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
                fake.ConfigureMultiplierFaultCase scenario, CStr(method), CLng(kind)
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
                Set result = Audit03ExecuteMultiplierFault(search, request, errorNumber, errorText)
                AssertTrue stats, prefix & ".noUnhandled", errorNumber = 0
                AssertTrue stats, prefix & ".resultExists", Not result Is Nothing
                If Not result Is Nothing Then
                    AssertTrue stats, prefix & ".status", result.Meta.InternalStatus = expectedStatus
                    AssertTrue stats, prefix & ".code", result.Meta.ResultCode = expectedCode
                    AssertTrue stats, prefix & ".display", policy.ExternalStatus(result.Meta) = expectedDisplay
                    AssertTrue stats, prefix & ".noPoint", Not result.HasLimitPoint And Not fake.Converged
                    AssertTrue stats, prefix & ".typedCause", fake.FailureCode = expectedFailure
                    AssertTrue stats, prefix & ".comment", InStr(1, result.Meta.ResultComment, _
                        "Контрольный отказ LoadMultiplier, сценарий " & CStr(scenario), vbBinaryCompare) > 0
                    If scenario = 13 Or scenario = 14 Then
                        AssertTrue stats, prefix & ".notCalculated", Not result.Meta.Calculated And Not result.SearchExecuted
                    ElseIf expectedStatus = rsInvalidInput Or expectedStatus = rsInvalidConfiguration Then
                        AssertTrue stats, prefix & ".validationLifecycle", Not result.Meta.Calculated And result.SearchExecuted
                    Else
                        AssertTrue stats, prefix & ".calculated", result.Meta.Calculated And result.SearchExecuted
                    End If
                Else
                    AppendLine stats, "FAULT_EXCEPTION: " & prefix & "; " & CStr(errorNumber) & "; " & errorText
                End If
                If expectedStatus <> rsNumericalFailure And scenario <> 15 Then
                    AssertTrue stats, prefix & ".noTerminalFallback", fake.FallbackCalls = 0
                End If
                AssertTrue stats, prefix & ".finiteWork", fake.ProbeCalls < 100
                fake.Configure CStr(method), 1.25, 0.000001, 80, resultKind:=CLng(kind)
                Set result = search.Execute(request)
                AssertTrue stats, prefix & ".recoverySuccess", result.Meta.InternalStatus = rsSuccess And result.HasLimitPoint
                AssertTrue stats, prefix & ".recoveryCauseReset", fake.FailureCode = sfcNone
                AssertTrue stats, prefix & ".recoveryCommentClean", InStr(1, result.Meta.ResultComment, "Контрольный отказ", vbBinaryCompare) = 0
            Next scenario
        Next method
    Next kind
End Sub

' ДЛЯ ТЕСТОВ: перехватывает только исключение проверяемого публичного Execute,
' чтобы отрицательный reproducer дошел до всех сценариев и записал каждый
' необработанный отказ. Такой отказ никогда не считается успешным result.
Private Function Audit03ExecuteMultiplierFault(ByVal search As CLoadMultiplierSearch, _
        ByVal request As CLimitSearchRequest, ByRef errorNumber As Long, ByRef errorText As String) As CLimitSearchResult
    On Error GoTo Failed
    Set Audit03ExecuteMultiplierFault = search.Execute(request)
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description
End Function















