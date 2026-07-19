Attribute VB_Name = "modTestCrackWidth"
Option Explicit

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

    TestConcreteTensionMode stats
    TestCrackMx stats
    TestCrackMy stats
    TestCrackMxy stats
    TestCircleCrackMxy stats
    TestCrackWriter stats
    TestNoTensionRebar stats

    AppendLine stats, "TOTAL_CRACK: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
    RunCrackWidthTests = stats.Report
    Exit Function

Failed:
    RunCrackWidthTests = "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

Private Sub TestConcreteTensionMode(ByRef stats As TCrackTestStats)
    Dim concrete As CConcreteDiagramMaterial
    Set concrete = ProvisionalConcrete()
    AssertClose stats, "crack.tensionMode.defaultStress", concrete.GetStress(0.0001), 0#, 0.000000000001
    AssertClose stats, "crack.tensionMode.defaultTangent", concrete.GetTangentModulus(0.0001), 0#, 0.000000000001

    concrete.TensionMode = "UseDiagram"
    concrete.TensionElasticModulus = 30000#
    concrete.TensionStressLimit = 1.5
    AssertClose stats, "crack.tensionMode.useStress", concrete.GetStress(0.00002), 0.6, 0.000000000001
    AssertClose stats, "crack.tensionMode.useLimit", concrete.GetStress(0.001), 1.5, 0.000000000001
    AssertClose stats, "crack.tensionMode.useTangent", concrete.GetTangentModulus(0.00002), 30000#, 0.000000000001
End Sub

Private Sub TestCircleCrackMxy(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim rebars As CRebarLayout
    Set solver = SolveCircleServiceState(rebars, -10000#, -8000000#, -8000000#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, rebars)
    AssertCrackCommon stats, "crack.circle.mxy", crack
End Sub

Private Sub TestCrackMx(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim rebars As CRebarLayout
    Set solver = SolveServiceState(rebars, -80000#, -5000000#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, rebars)
    AssertCrackCommon stats, "crack.mx", crack
End Sub

Private Sub TestCrackMy(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim rebars As CRebarLayout
    Set solver = SolveServiceState(rebars, -10000#, 0#, 5000000#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, rebars)
    AssertCrackCommon stats, "crack.my", crack
End Sub

Private Sub TestCrackMxy(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim rebars As CRebarLayout
    Set solver = SolveServiceState(rebars, -80000#, -3500000#, -2500000#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, rebars)
    AssertCrackCommon stats, "crack.mxy", crack
End Sub

Private Sub TestCrackWriter(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim rebars As CRebarLayout
    Set solver = SolveServiceState(rebars, -80000#, -3500000#, -2500000#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, rebars)
    Dim writer As CCapacityResultWriter
    Set writer = New CCapacityResultWriter
    writer.WriteCrackResult ThisWorkbook, crack

    AssertClose stats, "crack.writer.width", CDbl(ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(17, 5).Value2), crack.CrackWidth, 0.000000001
    AssertClose stats, "crack.writer.allowable", CDbl(ThisWorkbook.Names.Item("rngResultSection").RefersToRange.Cells.Item(18, 5).Value2), crack.AllowableCrackWidth, 0.000000001
End Sub

Private Sub TestNoTensionRebar(ByRef stats As TCrackTestStats)
    Dim solver As CSectionSolver
    Dim rebars As CRebarLayout
    Set solver = SolveServiceState(rebars, -100000#, 0#, 0#)

    Dim crack As CCrackWidthCalculator
    Set crack = CalculateCrack(solver, rebars)
    AssertTrue stats, "crack.noTension.converged", crack.Converged
    AssertClose stats, "crack.noTension.width", crack.CrackWidth, 0#, 0.000000000001
    AssertTrue stats, "crack.noTension.count", crack.TensionRebarCount = 0
End Sub

Private Function SolveServiceState(ByRef rebars As CRebarLayout, ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double) As CSectionSolver
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 300#, 200#, 0#, 0#, 0#, 0#

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 20#, 20#, 1

    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -90#, -60#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B2", 90#, -60#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B3", -90#, 60#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B4", 90#, 60#, 20#, 0#, "A400", "", geom

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    solver.LoadSteps = 8
    solver.MaxIterations = 80
    solver.ToleranceN = 5#
    solver.ToleranceMx = 5000#
    solver.ToleranceMy = 5000#
    solver.Solve mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), nValue, mxValue, myValue
    If Not solver.Converged Then Err.Raise vbObjectError + 3800, "modTestCrackWidth", "Service state did not converge: " & solver.StopReason
    Set SolveServiceState = solver
End Function

Private Function SolveCircleServiceState(ByRef rebars As CRebarLayout, ByVal nValue As Double, ByVal mxValue As Double, ByVal myValue As Double) As CSectionSolver
    Dim geom As CGeometryCircle
    Set geom = New CGeometryCircle
    geom.InitializeByDiameter 300#

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 15#, 15#, 1

    Set rebars = New CRebarLayout
    rebars.AddBar "B1", 0#, 90#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B2", 90#, 0#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B3", 0#, -90#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B4", -90#, 0#, 20#, 0#, "A400", "", geom

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    solver.LoadSteps = 8
    solver.MaxIterations = 80
    solver.ToleranceN = 5#
    solver.ToleranceMx = 5000#
    solver.ToleranceMy = 5000#
    solver.Solve mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), nValue, mxValue, myValue
    If Not solver.Converged Then Err.Raise vbObjectError + 3801, "modTestCrackWidth", "Circle service state did not converge: " & solver.StopReason
    Set SolveCircleServiceState = solver
End Function

Private Function CalculateCrack(ByVal solver As CSectionSolver, ByVal rebars As CRebarLayout) As CCrackWidthCalculator
    Dim crack As CCrackWidthCalculator
    Set crack = New CCrackWidthCalculator
    crack.AllowableCrackWidth = 0.3
    crack.CrackSpacing = 200#
    crack.StrainFactor = 1#
    crack.DurationFactor = 1#
    crack.Calculate solver, rebars, ProvisionalSteel()
    Set CalculateCrack = crack
End Function

Private Function ProvisionalConcrete() As CConcreteDiagramMaterial
    Dim concrete As CConcreteDiagramMaterial
    Set concrete = New CConcreteDiagramMaterial
    concrete.Initialize -0.0015, -15.5, -0.0035, -15.5
    concrete.TensionMode = "Ignore"
    Set ProvisionalConcrete = concrete
End Function

Private Function ProvisionalSteel() As CSteelDiagramMaterial
    Dim steel As CSteelDiagramMaterial
    Set steel = New CSteelDiagramMaterial
    steel.Initialize 0.00175, 350#, 0.025
    Set ProvisionalSteel = steel
End Function

Private Sub AssertCrackCommon(ByRef stats As TCrackTestStats, ByVal prefix As String, ByVal crack As CCrackWidthCalculator)
    AssertTrue stats, prefix & ".converged", crack.Converged
    AssertTrue stats, prefix & ".tensionBars", crack.TensionRebarCount > 0
    AssertTrue stats, prefix & ".strainPositive", crack.MaxSteelStrain > 0#
    AssertTrue stats, prefix & ".stressPositive", crack.MaxSteelStress > 0#
    AssertClose stats, prefix & ".widthFormula", crack.CrackWidth, crack.MaxSteelStrain * 200#, 0.000000001
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








