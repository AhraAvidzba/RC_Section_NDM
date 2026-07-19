Attribute VB_Name = "modTestLinearCore"
Option Explicit

Private Type TLinearTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

Public Function RunLinearCoreTests() As String
    On Error GoTo Failed

    Dim stats As TLinearTestStats
    Dim t0 As Double
    t0 = Timer

    TestLinearSystem stats
    TestCentralCompression stats
    TestBendingX stats
    TestBendingY stats
    TestCombinedLoading stats
    TestAsymmetricGeometryCoupling stats
    TestRebarAxialStiffness stats
    TestRebarValidation stats
    TestPerformanceLinear stats

    AppendLine stats, "TOTAL_LINEAR: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
    RunLinearCoreTests = stats.Report
    Exit Function

Failed:
    RunLinearCoreTests = "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

Private Sub TestLinearSystem(ByRef stats As TLinearTestStats)
    Dim solver As CLinearSystem3x3
    Set solver = New CLinearSystem3x3
    AssertTrue stats, "linsys.solve", solver.Solve(4#, 1#, 2#, 1#, 3#, 0#, 2#, 0#, 5#, 7#, 8#, 9#)
    AssertClose stats, "linsys.x1", solver.X1, 0.255813953488372, 0.000000000001
    AssertClose stats, "linsys.x2", solver.X2, 2.58139534883721, 0.000000000001
    AssertClose stats, "linsys.x3", solver.X3, 1.69767441860465, 0.000000000001
    AssertTrue stats, "linsys.residual", solver.RelativeResidual < 0.000000001

    Dim singular As CLinearSystem3x3
    Set singular = New CLinearSystem3x3
    AssertTrue stats, "linsys.singular", Not singular.Solve(1#, 2#, 3#, 2#, 4#, 6#, 3#, 6#, 9#, 1#, 2#, 3#)
End Sub

Private Sub TestCentralCompression(ByRef stats As TLinearTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = RectangleGeometry(200#, 100#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 10#)

    Dim props As CGeometryPropertiesCalculator
    Set props = MeshProps(mesh)

    Dim eb As Double
    eb = 30000#
    Dim targetN As Double
    targetN = 600000#

    Dim solver As CLinearSectionSolver
    Set solver = SolveLinear(mesh, Nothing, eb, 200000#, targetN, 0#, 0#)

    AssertTrue stats, "central.converged", solver.Converged
    AssertRelative stats, "central.eps0", solver.Epsilon0, targetN / (eb * props.Area), 0.000000001
    AssertClose stats, "central.kappaX", solver.KappaX, 0#, 0.000000000001
    AssertClose stats, "central.kappaY", solver.KappaY, 0#, 0.000000000001
    AssertEquilibrium stats, "central", solver, targetN, 0#, 0#
End Sub

Private Sub TestBendingX(ByRef stats As TLinearTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = RectangleGeometry(200#, 100#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 5#)
    Dim props As CGeometryPropertiesCalculator
    Set props = MeshProps(mesh)

    Dim eb As Double
    eb = 30000#
    Dim targetMx As Double
    targetMx = 250000000#

    Dim solver As CLinearSectionSolver
    Set solver = SolveLinear(mesh, Nothing, eb, 200000#, 0#, targetMx, 0#)

    AssertTrue stats, "bendX.converged", solver.Converged
    AssertRelative stats, "bendX.kappaX", solver.KappaX, targetMx / (eb * props.Ixc), 0.000000001
    AssertClose stats, "bendX.kappaY", solver.KappaY, 0#, 0.000000000001
    AssertEquilibrium stats, "bendX", solver, 0#, targetMx, 0#
End Sub

Private Sub TestBendingY(ByRef stats As TLinearTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = RectangleGeometry(200#, 100#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 5#)
    Dim props As CGeometryPropertiesCalculator
    Set props = MeshProps(mesh)

    Dim eb As Double
    eb = 30000#
    Dim targetMy As Double
    targetMy = 350000000#

    Dim solver As CLinearSectionSolver
    Set solver = SolveLinear(mesh, Nothing, eb, 200000#, 0#, 0#, targetMy)

    AssertTrue stats, "bendY.converged", solver.Converged
    AssertClose stats, "bendY.kappaX", solver.KappaX, 0#, 0.000000000001
    AssertRelative stats, "bendY.kappaY", solver.KappaY, targetMy / (eb * props.Iyc), 0.000000001
    AssertEquilibrium stats, "bendY", solver, 0#, 0#, targetMy
End Sub

Private Sub TestCombinedLoading(ByRef stats As TLinearTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = RectangleGeometry(200#, 100#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 5#)
    Dim props As CGeometryPropertiesCalculator
    Set props = MeshProps(mesh)

    Dim eb As Double
    eb = 30000#
    Dim n As Double
    Dim mx As Double
    Dim my As Double
    n = 400000#
    mx = 150000000#
    my = -120000000#

    Dim solver As CLinearSectionSolver
    Set solver = SolveLinear(mesh, Nothing, eb, 200000#, n, mx, my)

    AssertTrue stats, "combined.converged", solver.Converged
    AssertRelative stats, "combined.eps0", solver.Epsilon0, n / (eb * props.Area), 0.000000001
    AssertRelative stats, "combined.kappaX", solver.KappaX, mx / (eb * props.Ixc), 0.000000001
    AssertRelative stats, "combined.kappaY", solver.KappaY, my / (eb * props.Iyc), 0.000000001
    AssertEquilibrium stats, "combined", solver, n, mx, my
End Sub

Private Sub TestAsymmetricGeometryCoupling(ByRef stats As TLinearTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 300#, 180#, 10#, 35#, 20#, 50#

    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 12.5)

    Dim eb As Double
    eb = 30000#
    Dim targetMx As Double
    targetMx = 200000000#

    Dim solver As CLinearSectionSolver
    Set solver = SolveLinear(mesh, Nothing, eb, 200000#, 0#, targetMx, 0#)

    AssertTrue stats, "asymLinear.converged", solver.Converged
    AssertTrue stats, "asymLinear.kappaY.nonzero", Abs(solver.KappaY) > 0.000000000001
    AssertEquilibrium stats, "asymLinear", solver, 0#, targetMx, 0#
End Sub

Private Sub TestRebarAxialStiffness(ByRef stats As TLinearTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = RectangleGeometry(200#, 100#)

    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 10#)
    Dim props As CGeometryPropertiesCalculator
    Set props = MeshProps(mesh)

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -50#, 0#, 20#, 0#, "A", "auto area", geom
    rebars.AddBar "B2", 50#, 0#, 20#, 0#, "A", "auto area", geom

    Dim eb As Double
    Dim es As Double
    Dim n As Double
    eb = 30000#
    es = 200000#
    n = 700000#

    Dim solver As CLinearSectionSolver
    Set solver = SolveLinear(mesh, rebars, eb, es, n, 0#, 0#)

    Dim transformedStiffness As Double
    transformedStiffness = eb * props.Area + (es - eb) * (rebars.Area(1) + rebars.Area(2))

    AssertTrue stats, "rebar.converged", solver.Converged
    AssertRelative stats, "rebar.autoArea", rebars.Area(1), GEOM_PI * 20# * 20# / 4#, 0.000000001
    AssertRelative stats, "rebar.eps0", solver.Epsilon0, n / transformedStiffness, 0.000000001
    AssertRelative stats, "rebar.K11", solver.Stiffness(1, 1), transformedStiffness, 0.000000001
    AssertEquilibrium stats, "rebar", solver, n, 0#, 0#
End Sub

Private Sub TestRebarValidation(ByRef stats As TLinearTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = RectangleGeometry(100#, 100#)

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    AssertRebarError stats, "rebar.invalid.diameter", rebars, geom, "B1", 0#, 0#, 0#, 0#
    AssertRebarError stats, "rebar.outside", rebars, geom, "B1", 1000#, 0#, 10#, 0#

    rebars.Clear
    rebars.AddBar "B1", 0#, 0#, 10#, 0#, "A", "", geom
    AssertRebarError stats, "rebar.duplicate.id", rebars, geom, "B1", 10#, 0#, 10#, 0#
    AssertRebarError stats, "rebar.duplicate.coords", rebars, geom, "B2", 0#, 0#, 10#, 0#
End Sub

Private Sub TestPerformanceLinear(ByRef stats As TLinearTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = New CGeometryRoundedRectangle
    geom.Initialize 800#, 500#, 40#, 40#, 40#, 40#

    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 10#)

    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -250#, -150#, 25#, 0#, "A", "", geom
    rebars.AddBar "B2", 250#, -150#, 25#, 0#, "A", "", geom
    rebars.AddBar "B3", -250#, 150#, 25#, 0#, "A", "", geom
    rebars.AddBar "B4", 250#, 150#, 25#, 0#, "A", "", geom

    Dim solver As CLinearSectionSolver
    Set solver = SolveLinear(mesh, rebars, 30000#, 200000#, 1000000#, 250000000#, -150000000#)

    AppendLine stats, "PERFORMANCE_LINEAR: fibers=" & CStr(mesh.FiberCount) & _
        "; rebars=" & CStr(rebars.Count) & _
        "; matrixSec=" & FormatNumberInvariant(solver.MatrixBuildSeconds) & _
        "; solveSec=" & FormatNumberInvariant(solver.SolveSeconds) & _
        "; backCalcSec=" & FormatNumberInvariant(solver.BackCalcSeconds)
    AssertTrue stats, "perfLinear.converged", solver.Converged
End Sub

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

Private Function MeshProps(ByVal mesh As CFiberMeshBuilder) As CGeometryPropertiesCalculator
    Dim props As CGeometryPropertiesCalculator
    Set props = New CGeometryPropertiesCalculator
    props.CalculateFromMesh mesh
    Set MeshProps = props
End Function

Private Function SolveLinear(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal eb As Double, ByVal es As Double, ByVal n As Double, ByVal mx As Double, ByVal my As Double) As CLinearSectionSolver
    Dim solver As CLinearSectionSolver
    Set solver = New CLinearSectionSolver
    solver.Solve mesh, rebars, eb, es, n, mx, my
    Set SolveLinear = solver
End Function

Private Sub AssertRebarError(ByRef stats As TLinearTestStats, ByVal name As String, _
        ByVal rebars As CRebarLayout, ByVal geom As CGeometryRoundedRectangle, _
        ByVal barID As String, ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal diameter As Double, ByVal area As Double)
    On Error GoTo GotError
    rebars.AddBar barID, xCoord, yCoord, diameter, area, "A", "", geom
    On Error GoTo 0
    AssertTrue stats, name, False
    Exit Sub

GotError:
    On Error GoTo 0
    AssertTrue stats, name, True
End Sub

Private Sub AssertEquilibrium(ByRef stats As TLinearTestStats, ByVal prefix As String, _
        ByVal solver As CLinearSectionSolver, ByVal n As Double, ByVal mx As Double, ByVal my As Double)
    AssertLoadComponent stats, prefix & ".N", solver.Nint, n, 0.000001, 0.000000001
    AssertLoadComponent stats, prefix & ".Mx", solver.Mxint, mx, 0.0001, 0.000000001
    AssertLoadComponent stats, prefix & ".My", solver.Myint, my, 0.0001, 0.000000001
End Sub

Private Sub AssertLoadComponent(ByRef stats As TLinearTestStats, ByVal name As String, _
        ByVal actual As Double, ByVal expected As Double, ByVal zeroTolerance As Double, _
        ByVal relTolerance As Double)
    If Abs(expected) <= zeroTolerance Then
        AssertClose stats, name, actual, expected, zeroTolerance
    Else
        AssertRelative stats, name, actual, expected, relTolerance
    End If
End Sub

Private Sub AssertTrue(ByRef stats As TLinearTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name
    End If
End Sub

Private Sub AssertClose(ByRef stats As TLinearTestStats, ByVal name As String, ByVal actual As Double, _
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

Private Sub AssertRelative(ByRef stats As TLinearTestStats, ByVal name As String, ByVal actual As Double, _
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

Private Sub AppendLine(ByRef stats As TLinearTestStats, ByVal text As String)
    stats.Report = stats.Report & text & vbCrLf
End Sub

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function







