Attribute VB_Name = "modTestCapacitySolver"
Option Explicit

Private Type TCapacityTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

Public Function RunCapacitySolverTests() As String
    On Error GoTo Failed

    Dim stats As TCapacityTestStats
    Dim t0 As Double
    t0 = Timer

    TestMxPositiveAndNegative stats
    TestMyPositiveAndNegative stats
    TestMxySignedCombinations stats
    TestConcreteTensionModeAffectsSolverAndCapacity stats
    TestCircleCapacitySymmetry stats
    TestLambdaLessThanOne stats
    TestLambdaNearOne stats
    TestZeroAxialForce stats
    TestConcreteLimitState stats
    TestSteelLimitState stats
    TestNumericalFailureNotPhysicalBoundary stats
    TestAsymmetricCoupledCurvatures stats
    TestAsymmetricMxy stats
    TestResultWriter stats
    TestInvalidBaseMoment stats

    AppendLine stats, "TOTAL_CAPACITY: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
    RunCapacitySolverTests = stats.Report
    Exit Function

Failed:
    RunCapacitySolverTests = "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

Private Sub TestMxPositiveAndNegative(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim capNeg As CCapacitySolver
    Set capNeg = New CCapacitySolver
    ConfigureCapacity capNeg
    capNeg.SolveByLoadMultiplier mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), -300000#, -10000000#, 0#

    Dim capPos As CCapacitySolver
    Set capPos = New CCapacitySolver
    ConfigureCapacity capPos
    capPos.SolveByLoadMultiplier mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), -300000#, 10000000#, 0#

    AssertTrue stats, "capacity.mx.negative.converged", capNeg.Converged
    AssertTrue stats, "capacity.mx.positive.converged", capPos.Converged
    AssertTrue stats, "capacity.mx.negative.sign", capNeg.MxUltimate < 0#
    AssertTrue stats, "capacity.mx.positive.sign", capPos.MxUltimate > 0#
    AssertTrue stats, "capacity.mx.limit.physical", IsPhysicalLimitState(capNeg.LimitState) And IsPhysicalLimitState(capPos.LimitState)
    AssertEquilibrium stats, "capacity.mx.negative", capNeg.LastSolver, -300000#, capNeg.MxUltimate, 0#
    AssertEquilibrium stats, "capacity.mx.positive", capPos.LastSolver, -300000#, capPos.MxUltimate, 0#
End Sub

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
    capMx.SolveByLoadMultiplier mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), -220000#, -6000000#, 0#

    Dim capMy As CCapacitySolver
    Set capMy = New CCapacitySolver
    ConfigureCapacity capMy
    capMy.SolveByLoadMultiplier mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), -220000#, 0#, -6000000#

    Dim capMxy As CCapacitySolver
    Set capMxy = New CCapacitySolver
    ConfigureCapacity capMxy
    capMxy.SolveByLoadMultiplier mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), -220000#, -5000000#, -5000000#

    AssertTrue stats, "circle.capacity.mx.converged", capMx.Converged
    AssertTrue stats, "circle.capacity.my.converged", capMy.Converged
    AssertTrue stats, "circle.capacity.mxy.converged", capMxy.Converged
    AssertRelative stats, "circle.capacity.mx.my.symmetry", Abs(capMx.MomentUltimate), Abs(capMy.MomentUltimate), 0.03
    AssertTrue stats, "circle.capacity.mxy.lambda.positive", capMxy.LambdaUltimate > 0#
    AssertClose stats, "circle.capacity.mxy.direction", capMxy.MxUltimate * -5000000# - capMxy.MyUltimate * -5000000#, 0#, 1000#
End Sub

Private Sub TestConcreteTensionModeAffectsSolverAndCapacity(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim ignoreConcrete As CConcreteDiagramMaterial
    Set ignoreConcrete = ProvisionalConcrete()
    ignoreConcrete.TensionMode = "Ignore"

    Dim useConcrete As CConcreteDiagramMaterial
    Set useConcrete = ProvisionalConcrete()
    useConcrete.TensionMode = "UseDiagram"
    useConcrete.TensionElasticModulus = 32500#
    useConcrete.TensionStressLimit = 1.1

    Dim solverIgnore As CSectionSolver
    Set solverIgnore = New CSectionSolver
    ConfigureSectionSolver solverIgnore
    solverIgnore.Solve mesh, rebars, ignoreConcrete, ProvisionalSteel(), -50000#, -5000000#, -3000000#

    Dim solverUse As CSectionSolver
    Set solverUse = New CSectionSolver
    ConfigureSectionSolver solverUse
    solverUse.Solve mesh, rebars, useConcrete, ProvisionalSteel(), -50000#, -5000000#, -3000000#

    AssertTrue stats, "tensionMode.solver.ignore.converged", solverIgnore.Converged
    AssertTrue stats, "tensionMode.solver.use.converged", solverUse.Converged
    AssertTrue stats, "tensionMode.solver.affectsState", Abs(solverIgnore.KappaX - solverUse.KappaX) > 0.000000001 Or Abs(solverIgnore.KappaY - solverUse.KappaY) > 0.000000001

    AssertTensionModeAffectsSolverMode stats, "tensionMode.strength.mx", mesh, rebars, -50000#, -5000000#, 0#
    AssertTensionModeAffectsSolverMode stats, "tensionMode.strength.my", mesh, rebars, -50000#, 0#, -5000000#
    AssertTensionModeAffectsSolverMode stats, "tensionMode.strength.mxy", mesh, rebars, -50000#, -5000000#, -3000000#
End Sub

Private Sub AssertTensionModeAffectsSolverMode(ByRef stats As TCapacityTestStats, ByVal prefix As String, _
        ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double)
    Dim ignoreConcrete As CConcreteDiagramMaterial
    Set ignoreConcrete = ProvisionalConcrete()
    ignoreConcrete.TensionMode = "Ignore"

    Dim useConcrete As CConcreteDiagramMaterial
    Set useConcrete = ProvisionalConcrete()
    useConcrete.TensionMode = "UseDiagram"
    useConcrete.TensionElasticModulus = 32500#
    useConcrete.TensionStressLimit = 1.1

    Dim solverIgnore As CSectionSolver
    Set solverIgnore = New CSectionSolver
    ConfigureSectionSolver solverIgnore
    solverIgnore.Solve mesh, rebars, ignoreConcrete, ProvisionalSteel(), nValue, mxValue, myValue

    Dim solverUse As CSectionSolver
    Set solverUse = New CSectionSolver
    ConfigureSectionSolver solverUse
    solverUse.Solve mesh, rebars, useConcrete, ProvisionalSteel(), nValue, mxValue, myValue

    AssertTrue stats, prefix & ".ignore.converged", solverIgnore.Converged
    AssertTrue stats, prefix & ".use.converged", solverUse.Converged
    AssertTrue stats, prefix & ".stateDiffers", Abs(solverIgnore.KappaX - solverUse.KappaX) > 0.000000001 Or _
        Abs(solverIgnore.KappaY - solverUse.KappaY) > 0.000000001 Or Abs(solverIgnore.Epsilon0 - solverUse.Epsilon0) > 0.000000001
End Sub

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
    cap.SolveByLoadMultiplier mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), nValue, mxBase, myBase

    AssertTrue stats, prefix & ".converged", cap.Converged
    AssertTrue stats, prefix & ".lambda.positive", cap.LambdaUltimate > 0#
    AssertTrue stats, prefix & ".mx.sign", Sgn(cap.MxUltimate) = Sgn(mxBase)
    AssertTrue stats, prefix & ".my.sign", Sgn(cap.MyUltimate) = Sgn(myBase)
    AssertClose stats, prefix & ".direction", cap.MxUltimate * myBase - cap.MyUltimate * mxBase, 0#, 1000#
    AssertEquilibrium stats, prefix, cap.LastSolver, nValue, cap.MxUltimate, cap.MyUltimate
End Sub

Private Sub TestMyPositiveAndNegative(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim capNeg As CCapacitySolver
    Set capNeg = New CCapacitySolver
    ConfigureCapacity capNeg
    capNeg.SolveByLoadMultiplier mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), -300000#, 0#, -10000000#

    Dim capPos As CCapacitySolver
    Set capPos = New CCapacitySolver
    ConfigureCapacity capPos
    capPos.SolveByLoadMultiplier mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), -300000#, 0#, 10000000#

    AssertTrue stats, "capacity.my.negative.converged", capNeg.Converged
    AssertTrue stats, "capacity.my.positive.converged", capPos.Converged
    AssertTrue stats, "capacity.my.negative.sign", capNeg.MyUltimate < 0#
    AssertTrue stats, "capacity.my.positive.sign", capPos.MyUltimate > 0#
    AssertTrue stats, "capacity.my.limit.physical", IsPhysicalLimitState(capNeg.LimitState) And IsPhysicalLimitState(capPos.LimitState)
    AssertEquilibrium stats, "capacity.my.negative", capNeg.LastSolver, -300000#, 0#, capNeg.MyUltimate
    AssertEquilibrium stats, "capacity.my.positive", capPos.LastSolver, -300000#, 0#, capPos.MyUltimate
End Sub
Private Sub TestLambdaLessThanOne(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim reference As CCapacitySolver
    Set reference = New CCapacitySolver
    ConfigureCapacity reference
    reference.SolveByLoadMultiplier mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), -300000#, -10000000#, 0#

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SolveByLoadMultiplier mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), -300000#, reference.MxUltimate * 2#, 0#

    AssertTrue stats, "capacity.lambda.lessThanOne.converged", cap.Converged
    AssertTrue stats, "capacity.lambda.lessThanOne.value", cap.LambdaUltimate < 1#
End Sub

Private Sub TestLambdaNearOne(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim reference As CCapacitySolver
    Set reference = New CCapacitySolver
    ConfigureCapacity reference
    reference.SolveByLoadMultiplier mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), -300000#, -10000000#, 0#

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SolveByLoadMultiplier mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), -300000#, reference.MxUltimate, 0#

    AssertTrue stats, "capacity.lambda.nearOne.converged", cap.Converged
    AssertClose stats, "capacity.lambda.nearOne.value", cap.LambdaUltimate, 1#, 0.03
End Sub

Private Sub TestZeroAxialForce(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.ConcreteCompressionLimit = -0.0006
    cap.SteelStrainLimit = 0.001#
    cap.SolveByLoadMultiplier mesh, rebars, LinearConcrete(), LinearSteel(), 0#, -5000000#, 0#

    AssertTrue stats, "capacity.zeroN.converged", cap.Converged
    AssertEquilibrium stats, "capacity.zeroN", cap.LastSolver, 0#, cap.MxUltimate, 0#
End Sub

Private Sub TestConcreteLimitState(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.ConcreteCompressionLimit = -0.0006
    cap.SteelStrainLimit = 1#
    cap.SolveByLoadMultiplier mesh, rebars, LinearConcrete(), LinearSteel(), 0#, -10000000#, 0#

    AssertTrue stats, "capacity.concreteLimit.converged", cap.Converged
    AssertEquals stats, "capacity.concreteLimit.state", cap.LimitState, "ConcreteStrainLimit"
End Sub

Private Sub TestSteelLimitState(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.ConcreteCompressionLimit = -1#
    cap.SteelStrainLimit = 0.0005
    cap.SolveByLoadMultiplier mesh, rebars, LinearConcrete(), LinearSteel(), 0#, -10000000#, 0#

    AssertTrue stats, "capacity.steelLimit.converged", cap.Converged
    AssertEquals stats, "capacity.steelLimit.state", cap.LimitState, "SteelStrainLimit"
End Sub

Private Sub TestNumericalFailureNotPhysicalBoundary(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SolverMaxIterations = 1
    cap.MaxRetries = 1
    cap.SolveByLoadMultiplier mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), 0#, -10000000#, 0#

    AssertTrue stats, "capacity.numericalFailure.notConverged", Not cap.Converged
    AssertTrue stats, "capacity.numericalFailure.state", cap.LimitState = "NumericalFailure" Or cap.LimitState = "SingularTangent"
    AssertTrue stats, "capacity.numericalFailure.noPhysicalUpper", Not IsPhysicalLimitState(cap.LimitState)
    AssertTrue stats, "capacity.numericalFailure.retried", cap.RetryCount > 0
End Sub

Private Sub TestAsymmetricCoupledCurvatures(ByRef stats As TCapacityTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 320#, 220#, 0#, 70#, 20#, 0#

    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 10#)

    Dim props As CGeometryPropertiesCalculator
    Set props = New CGeometryPropertiesCalculator
    props.CalculateFromMesh mesh

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -110#, -70#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B2", 95#, -65#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B3", -90#, 65#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B4", 80#, 70#, 20#, 0#, "A400", "", geom

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SolveByLoadMultiplier mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), -260000#, -8000000#, 0#

    AssertTrue stats, "capacity.asym.Ixy.nonzero", Abs(props.Ixyc) > 1000000#
    AssertTrue stats, "capacity.asym.converged", cap.Converged
    AssertTrue stats, "capacity.asym.coupledKappaY", Abs(cap.LastSolver.KappaY) > 0.000000001
    AssertEquilibrium stats, "capacity.asym", cap.LastSolver, -260000#, cap.MxUltimate, 0#
End Sub

Private Sub TestAsymmetricMxy(ByRef stats As TCapacityTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 320#, 220#, 0#, 70#, 20#, 0#

    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 10#)

    Dim props As CGeometryPropertiesCalculator
    Set props = New CGeometryPropertiesCalculator
    props.CalculateFromMesh mesh

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -110#, -70#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B2", 95#, -65#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B3", -90#, 65#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B4", 80#, 70#, 20#, 0#, "A400", "", geom

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SolveByLoadMultiplier mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), -260000#, -6000000#, -4000000#

    AssertTrue stats, "capacity.mxy.asym.Ixy.nonzero", Abs(props.Ixyc) > 1000000#
    AssertTrue stats, "capacity.mxy.asym.converged", cap.Converged
    AssertTrue stats, "capacity.mxy.asym.kappaX.nonzero", Abs(cap.LastSolver.KappaX) > 0.000000001
    AssertTrue stats, "capacity.mxy.asym.kappaY.nonzero", Abs(cap.LastSolver.KappaY) > 0.000000001
    AssertClose stats, "capacity.mxy.asym.direction", cap.MxUltimate * -4000000# - cap.MyUltimate * -6000000#, 0#, 1000#
    AssertEquilibrium stats, "capacity.mxy.asym", cap.LastSolver, -260000#, cap.MxUltimate, cap.MyUltimate
End Sub

Private Sub TestResultWriter(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 300#, 200#, 20#, 90#, 60#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SolveByLoadMultiplier mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), -250000#, -6000000#, -4000000#

    Dim writer As CCapacityResultWriter
    Set writer = New CCapacityResultWriter
    writer.WriteCapacityResult ThisWorkbook, cap

    AssertEquals stats, "capacity.writer.mode", CStr(ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(2, 5).Value2), "N+Mx+My"
    AssertTrue stats, "capacity.writer.lambda", CDbl(ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(5, 5).Value2) > 0#
    AssertTrue stats, "capacity.writer.mx", CDbl(ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(6, 5).Value2) <> 0#
    AssertTrue stats, "capacity.writer.my", CDbl(ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(7, 5).Value2) <> 0#
End Sub
Private Sub TestInvalidBaseMoment(ByRef stats As TCapacityTestStats)
    Dim mesh As CFiberMeshBuilder
    Dim rebars As CRebarLayout
    PrepareSymmetricSection 200#, 100#, 20#, 60#, 30#, mesh, rebars

    Dim cap As CCapacitySolver
    Set cap = New CCapacitySolver
    ConfigureCapacity cap
    cap.SolveByLoadMultiplier mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), -100000#, 0#, 0#

    AssertTrue stats, "capacity.invalidBaseMoment.notConverged", Not cap.Converged
    AssertEquals stats, "capacity.invalidBaseMoment.state", cap.LimitState, "InvalidInput"
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

Private Function ProvisionalConcrete() As CConcreteDiagramMaterial
    Dim concrete As CConcreteDiagramMaterial
    Set concrete = New CConcreteDiagramMaterial
    concrete.Initialize -0.0015, -15.5, -0.0035, -15.5
    Set ProvisionalConcrete = concrete
End Function

Private Function ProvisionalSteel() As CSteelDiagramMaterial
    Dim steel As CSteelDiagramMaterial
    Set steel = New CSteelDiagramMaterial
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

Private Function BuildMesh(ByVal geom As CGeometryRoundedRectangle, ByVal stepSize As Double) As CFiberMeshBuilder
    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, stepSize, stepSize, 1
    Set BuildMesh = mesh
End Function

Private Function IsPhysicalLimitState(ByVal state As String) As Boolean
    IsPhysicalLimitState = (state = "ConcreteStrainLimit" Or state = "SteelStrainLimit")
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











