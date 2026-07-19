Attribute VB_Name = "modTestSectionSolver"
Option Explicit

Private Type TSectionSolverTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

Public Function RunSectionSolverTests() As String
    On Error GoTo Failed

    Dim stats As TSectionSolverTestStats
    Dim t0 As Double
    t0 = Timer

    TestSystemSettingsReader stats
    TestSystemSettingsCatalog stats
    TestLinearMaterialAgainstLinearSolver stats
    TestLinearWithRebarReplacement stats
    TestDiagramConcreteCentralCompression stats
    TestDiagramConcreteWithRebar stats
    TestIncrementLimitsAndDiagnostics stats

    AppendLine stats, "TOTAL_SECTION_SOLVER: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
    RunSectionSolverTests = stats.Report
    Exit Function

Failed:
    RunSectionSolverTests = "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

Private Sub TestSystemSettingsReader(ByRef stats As TSectionSolverTestStats)
    Dim reader As CSystemSettingsReader
    Set reader = New CSystemSettingsReader
    reader.LoadFromWorkbook ThisWorkbook

    AssertClose stats, "settings.concrete.Eb", reader.GetDouble("Concrete.Eb", 0#), 32500#, 0.000000001
    AssertClose stats, "settings.steel.Es", reader.GetDouble("Steel.Es", 0#), 200000#, 0.000000001
    AssertClose stats, "settings.steel.EpsY", reader.GetDouble("Steel.Point1.Eps", 0#), 0.00175, 0.000000000001
End Sub

Private Sub TestSystemSettingsCatalog(ByRef stats As TSectionSolverTestStats)
    Dim reader As CSystemSettingsReader
    Set reader = New CSystemSettingsReader
    reader.LoadFromWorkbook ThisWorkbook

    AssertTrue stats, "settings.catalog.loaded", reader.KeyCount > 0
    AssertTrue stats, "settings.catalog.noDuplicateKeys", reader.DuplicateCount = 0

    Dim requiredKeys As Variant
    requiredKeys = Array( _
        "Geometry.Type", "Circle.Diameter", "Circle.CenterX", "Circle.CenterY", _
        "Mesh.StepX", "Mesh.StepY", "Mesh.BoundarySubdivisions", _
        "Concrete.Eb", "Concrete.TensionMode", "Concrete.Point1.Eps", _
        "Concrete.Point1.Stress", "Concrete.Point2.Eps", "Concrete.Point2.Stress", _
        "Concrete.Point3.Eps", "Concrete.Point3.Stress", _
        "Steel.Es", "Steel.Point1.Eps", "Steel.Point1.Stress", "Steel.Point2.Eps", _
        "Steel.Point2.Stress", "Steel.Point3.Eps", "Steel.Point3.Stress", _
        "Calculation.Mode", "Solver.MaxIterations", "Solver.LoadSteps", _
        "Solver.ToleranceN", "Solver.ToleranceMx", "Solver.ToleranceMy", _
        "Solver.LineSearchEnabled", "Solver.DampingInitial", "Solver.MinLineSearchAlpha", _
        "Solver.MaxDeltaEpsilon0", "Solver.MaxDeltaKappa", "Solver.DiagnosticsEnabled", _
        "Capacity.InitialLambda", "Capacity.MaxLambda", "Capacity.ToleranceLambda", _
        "Capacity.MaxRetries", "Capacity.BaseLoadSteps", "Capacity.SolverMaxIterations", _
        "Capacity.ConcreteCompressionLimit", "Capacity.SteelStrainLimit", _
        "CrackWidth.Enabled", "CrackWidth.Allowable", "CrackWidth.CrackSpacing", _
        "CrackWidth.StrainFactor", "CrackWidth.DurationFactor")

    Dim i As Long
    For i = LBound(requiredKeys) To UBound(requiredKeys)
        AssertTrue stats, "settings.key." & CStr(requiredKeys(i)), reader.HasKey(CStr(requiredKeys(i)))
    Next i

    Dim removedKeys As Variant
    removedKeys = Array("Capacity.Enabled", "Capacity.CalculateMx", "Capacity.CalculateMy", _
        "Capacity.CalculateMxy", "Mesh.BoundaryMode", "Circle.Radius", _
        "Batch.MaxCombinations", "Batch.Diagnostics", "Materials.SourceStatus", _
        "Concrete.Diagram", "Steel.Diagram")
    For i = LBound(removedKeys) To UBound(removedKeys)
        AssertTrue stats, "settings.removed." & CStr(removedKeys(i)), Not reader.HasKey(CStr(removedKeys(i)))
    Next i

    Dim rangeNames As Variant
    rangeNames = Array("SolverSettings", "CapacitySettings", "ConcreteDiagram", _
        "SteelDiagram", "GeometrySettings", "OutputSettings", "AutoCADSettings")
    For i = LBound(rangeNames) To UBound(rangeNames)
        AssertTrue stats, "settings.range." & CStr(rangeNames(i)), NamedRangeExists(CStr(rangeNames(i)))
    Next i
End Sub

Private Sub TestLinearMaterialAgainstLinearSolver(ByRef stats As TSectionSolverTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = RectangleGeometry(200#, 100#)
    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 10#)

    Dim concrete As CLinearConcreteMaterial
    Set concrete = New CLinearConcreteMaterial
    concrete.Initialize 30000#
    Dim steel As CLinearSteelMaterial
    Set steel = New CLinearSteelMaterial
    steel.Initialize 200000#

    Dim n As Double
    Dim mx As Double
    Dim my As Double
    n = 250000#
    mx = 120000000#
    my = -80000000#

    Dim linear As CLinearSectionSolver
    Set linear = New CLinearSectionSolver
    linear.Solve mesh, Nothing, 30000#, 200000#, n, mx, my

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    ConfigureStrictSolver solver
    solver.Solve mesh, Nothing, concrete, steel, n, mx, my

    AssertTrue stats, "section.linear.converged", solver.Converged
    AssertRelative stats, "section.linear.eps0", solver.Epsilon0, linear.Epsilon0, 0.000000001
    AssertRelative stats, "section.linear.kappaX", solver.KappaX, linear.KappaX, 0.000000001
    AssertRelative stats, "section.linear.kappaY", solver.KappaY, linear.KappaY, 0.000000001
    AssertEquilibrium stats, "section.linear", solver, n, mx, my
End Sub

Private Sub TestLinearWithRebarReplacement(ByRef stats As TSectionSolverTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = RectangleGeometry(200#, 100#)
    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 10#)
    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -60#, -30#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B2", 60#, -30#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B3", -60#, 30#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B4", 60#, 30#, 20#, 0#, "A400", "", geom

    Dim concrete As CLinearConcreteMaterial
    Set concrete = New CLinearConcreteMaterial
    concrete.Initialize 30000#
    Dim steel As CLinearSteelMaterial
    Set steel = New CLinearSteelMaterial
    steel.Initialize 200000#

    Dim n As Double
    Dim mx As Double
    Dim my As Double
    n = 500000#
    mx = 90000000#
    my = 70000000#

    Dim linear As CLinearSectionSolver
    Set linear = New CLinearSectionSolver
    linear.Solve mesh, rebars, 30000#, 200000#, n, mx, my

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    ConfigureStrictSolver solver
    solver.Solve mesh, rebars, concrete, steel, n, mx, my

    AssertTrue stats, "section.rebar.converged", solver.Converged
    AssertRelative stats, "section.rebar.eps0", solver.Epsilon0, linear.Epsilon0, 0.000000001
    AssertRelative stats, "section.rebar.kappaX", solver.KappaX, linear.KappaX, 0.000000001
    AssertRelative stats, "section.rebar.kappaY", solver.KappaY, linear.KappaY, 0.000000001
    AssertEquilibrium stats, "section.rebar", solver, n, mx, my
End Sub

Private Sub TestDiagramConcreteCentralCompression(ByRef stats As TSectionSolverTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = RectangleGeometry(200#, 100#)
    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 10#)

    Dim concrete As CConcreteDiagramMaterial
    Set concrete = ProvisionalConcrete()
    Dim steel As CSteelDiagramMaterial
    Set steel = ProvisionalSteel()

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    ConfigureProvisionalSolver solver
    solver.Solve mesh, Nothing, concrete, steel, -100000#, 0#, 0#

    AssertTrue stats, "section.diagramCompression.converged", solver.Converged
    AssertTrue stats, "section.diagramCompression.epsNegative", solver.Epsilon0 < 0#
    AssertClose stats, "section.diagramCompression.kappaX", solver.KappaX, 0#, 0.000000000001
    AssertClose stats, "section.diagramCompression.kappaY", solver.KappaY, 0#, 0.000000000001
    AssertEquilibrium stats, "section.diagramCompression", solver, -100000#, 0#, 0#
End Sub

Private Sub TestDiagramConcreteWithRebar(ByRef stats As TSectionSolverTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = RectangleGeometry(300#, 200#)
    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 20#)
    Dim rebars As CRebarLayout
    Set rebars = New CRebarLayout
    rebars.AddBar "B1", -90#, -60#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B2", 90#, -60#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B3", -90#, 60#, 20#, 0#, "A400", "", geom
    rebars.AddBar "B4", 90#, 60#, 20#, 0#, "A400", "", geom

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    ConfigureProvisionalSolver solver
    solver.Solve mesh, rebars, ProvisionalConcrete(), ProvisionalSteel(), -300000#, -20000000#, 0#

    AssertTrue stats, "section.diagramRebar.converged", solver.Converged
    AssertEquilibrium stats, "section.diagramRebar", solver, -300000#, -20000000#, 0#
    AssertTrue stats, "section.diagramRebar.steps", solver.LoadStepsCompleted = 5
End Sub

Private Sub TestIncrementLimitsAndDiagnostics(ByRef stats As TSectionSolverTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = RectangleGeometry(200#, 100#)
    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 10#)

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    ConfigureProvisionalSolver solver
    solver.MaxDeltaEpsilon0 = 0.00002
    solver.Solve mesh, Nothing, ProvisionalConcrete(), ProvisionalSteel(), -80000#, 0#, 0#

    AssertTrue stats, "section.diagnostics.converged", solver.Converged
    AssertTrue stats, "section.diagnostics.iterations", solver.Iterations > solver.LoadStepsCompleted
    AssertTrue stats, "section.diagnostics.log", InStr(1, solver.DiagnosticLog, "iter=", vbTextCompare) > 0
End Sub

Private Sub ConfigureStrictSolver(ByVal solver As CSectionSolver)
    solver.LoadSteps = 1
    solver.MaxIterations = 12
    solver.ToleranceN = 0.000001
    solver.ToleranceMx = 0.001
    solver.ToleranceMy = 0.001
    solver.LineSearchEnabled = True
    solver.MaxDeltaEpsilon0 = 0#
    solver.MaxDeltaKappa = 0#
End Sub

Private Sub ConfigureProvisionalSolver(ByVal solver As CSectionSolver)
    solver.LoadSteps = 5
    solver.MaxIterations = 40
    solver.ToleranceN = 1#
    solver.ToleranceMx = 1000#
    solver.ToleranceMy = 1000#
    solver.LineSearchEnabled = True
    solver.MaxDeltaEpsilon0 = 0.0005
    solver.MaxDeltaKappa = 0.00001
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

Private Sub AssertEquilibrium(ByRef stats As TSectionSolverTestStats, ByVal prefix As String, _
        ByVal solver As CSectionSolver, ByVal n As Double, ByVal mx As Double, ByVal my As Double)
    AssertLoadComponent stats, prefix & ".N", solver.Nint, n, 1#, 0.000001
    AssertLoadComponent stats, prefix & ".Mx", solver.Mxint, mx, 1000#, 0.000001
    AssertLoadComponent stats, prefix & ".My", solver.Myint, my, 1000#, 0.000001
End Sub

Private Function NamedRangeExists(ByVal rangeName As String) As Boolean
    On Error GoTo Missing
    Dim target As Object
    Set target = ThisWorkbook.Names.Item(rangeName).RefersToRange
    NamedRangeExists = Not target Is Nothing
    Exit Function
Missing:
    NamedRangeExists = False
End Function

Private Sub AssertLoadComponent(ByRef stats As TSectionSolverTestStats, ByVal name As String, _
        ByVal actual As Double, ByVal expected As Double, ByVal zeroTolerance As Double, _
        ByVal relTolerance As Double)
    If Abs(expected) <= zeroTolerance Then
        AssertClose stats, name, actual, expected, zeroTolerance
    Else
        AssertRelative stats, name, actual, expected, relTolerance
    End If
End Sub

Private Sub AssertTrue(ByRef stats As TSectionSolverTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name
    End If
End Sub

Private Sub AssertClose(ByRef stats As TSectionSolverTestStats, ByVal name As String, ByVal actual As Double, _
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

Private Sub AssertRelative(ByRef stats As TSectionSolverTestStats, ByVal name As String, ByVal actual As Double, _
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

Private Sub AppendLine(ByRef stats As TSectionSolverTestStats, ByVal text As String)
    stats.Report = stats.Report & text & vbCrLf
End Sub

Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function




