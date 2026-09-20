Attribute VB_Name = "modTestSectionSolver"
Option Explicit

' ==========================================================================
' Тесты решателя равновесия CSectionSolver
' ==========================================================================
' Проверяется поиск epsilon0/kappaX/kappaY в единой постановке N + Mx + My,
' включая одноосные частные случаи и перенос начала координат.

Private Type TSectionSolverTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

' Запускает связанный набор операций и возвращает пользователю итоговый статус выполнения.
Public Function RunSectionSolverTests() As String
    On Error GoTo Failed

    Dim stats As TSectionSolverTestStats
    Dim t0 As Double
    t0 = Timer

    TestSystemSettingsReader stats
    TestSystemSettingsCatalog stats
    TestUnitSystemConversions stats
    TestLinearSystem3x3 stats
    TestLinearMaterialEquilibrium stats
    TestLinearMaterialWithRebarReplacement stats
    TestDiagramConcreteCentralCompression stats
    TestDiagramConcreteWithRebar stats
    TestIncrementLimitsAndDiagnostics stats
    TestSecantIndependentBranch stats
    TestSecantComparativeTasks stats
    TestSolverMethodInputErrors stats
    TestSolverMethodFromSystem stats

    AppendLine stats, "TOTAL_SECTION_SOLVER: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
    RunSectionSolverTests = stats.Report
    Exit Function

Failed:
    RunSectionSolverTests = "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestUnitSystemConversions(ByRef stats As TSectionSolverTestStats)
    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.InitializeDefaults

    AssertClose stats, "units.length.mm.cm", units.InternalLengthToOutput(units.InputLengthToInternal(25#)), 25#, 0.000000000001

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook

    Dim configured As CUnitSystem
    Set configured = New CUnitSystem
    configured.LoadFromSettings settings

    AssertClose stats, "units.force.tf.toN.compression", configured.InputForceToInternal(1#), -9806.65, 0.000001
    AssertClose stats, "units.force.N.toUser.compression", configured.InternalForceToOutput(-9806.65), 1#, 0.000001
    AssertClose stats, "units.moment.tfm.toNmm", configured.InputMomentMxToInternal(1#), 9806650#, 0.0001
    AssertClose stats, "units.moment.Nmm.toUser", configured.InternalMomentMxToOutput(9806650#), 1#, 0.000001
    AssertClose stats, "units.momentPerLength.defaultZero", configured.InputMomentPerLengthToInternal(0.0005), 4903.325, 0.000001

    AssertClose stats, "units.loadcase.N.example", configured.InputForceToInternal(30#), -294199.5, 0.0001
    AssertClose stats, "units.loadcase.Mx.example", configured.InputMomentMxToInternal(150#), 1470997500#, 0.1
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestLinearSystem3x3(ByRef stats As TSectionSolverTestStats)
    Dim system As CLinearSystem3x3
    Set system = New CLinearSystem3x3
    AssertTrue stats, "linsys.solve", system.Solve(3#, 2#, -1#, 2#, -2#, 4#, -1#, 0.5, -1#, 1#, -2#, 0#)
    AssertClose stats, "linsys.x1", system.X1, 1#, 0.000000000001
    AssertClose stats, "linsys.x2", system.X2, -2#, 0.000000000001
    AssertClose stats, "linsys.x3", system.X3, -2#, 0.000000000001
    AssertTrue stats, "linsys.residual", system.AbsoluteResidual < 0.000000001

    Dim singular As CLinearSystem3x3
    Set singular = New CLinearSystem3x3
    AssertTrue stats, "linsys.singular", Not singular.Solve(1#, 2#, 3#, 2#, 4#, 6#, 3#, 6#, 9#, 1#, 2#, 3#)
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestSystemSettingsReader(ByRef stats As TSectionSolverTestStats)
    Dim reader As CSystemSettingsReader
    Set reader = New CSystemSettingsReader
    reader.LoadFromWorkbook ThisWorkbook

    AssertClose stats, "settings.concrete.Eb", reader.GetDouble("Concrete.Eb", 0#), 32500#, 0.000000001
    AssertClose stats, "settings.steel.Es", reader.GetDouble("Steel.Es", 0#), 200000#, 0.000000001
    AssertTrue stats, "settings.profiles.present", ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange.Rows.Count > 1

    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.Initialize reader
    AssertClose stats, "settings.material.strengthTensionIgnored", provider.ConcreteMaterial(cpStrength).GetStress(0.0001), 0#, 0.000000000001
    AssertTrue stats, "settings.material.mcrcTensionEnabled", provider.ConcreteMaterial(cpMcrc).GetStress(0.0001) > 0#
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestSecantComparativeTasks(ByRef stats As TSectionSolverTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = RectangleGeometry(300#, 200#)
    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 20#)

    AssertNewtonSecantCase stats, "secant.compare.compression", mesh, Nothing, -100000#, 0#, 0#
    AssertNewtonSecantCase stats, "secant.compare.n_mx", mesh, Nothing, -150000#, -8000000#, 0#
    AssertNewtonSecantCase stats, "secant.compare.n_my", mesh, Nothing, -150000#, 0#, -6000000#
    AssertNewtonSecantCase stats, "secant.compare.n_mx_my", mesh, Nothing, -150000#, -8000000#, -6000000#

    Dim circleGeom As CGeometryCircle
    Set circleGeom = New CGeometryCircle
    circleGeom.InitializeByDiameter 300#
    Dim circleMesh As CFiberMeshBuilder
    Set circleMesh = New CFiberMeshBuilder
    circleMesh.BuildMesh circleGeom, 20#, 20#, 1
    AssertNewtonSecantCase stats, "secant.compare.circle", circleMesh, Nothing, -120000#, -5000000#, 0#
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestSystemSettingsCatalog(ByRef stats As TSectionSolverTestStats)
    Dim reader As CSystemSettingsReader
    Set reader = New CSystemSettingsReader
    reader.LoadFromWorkbook ThisWorkbook

    AssertTrue stats, "settings.catalog.loaded", reader.KeyCount > 0
    AssertTrue stats, "settings.catalog.noDuplicateKeys", reader.DuplicateCount = 0

    Dim requiredKeys As Variant
    requiredKeys = Array( _
        "Units.Length.Input", "Units.Length.Internal", "Units.Length.Output", _
        "Units.Area.Input", "Units.Area.Internal", "Units.Area.Output", _
        "Units.Force.Input", "Units.Force.Internal", "Units.Force.Output", _
        "Units.Moment.Input", "Units.Moment.Internal", "Units.Moment.Output", _
        "Units.Stress.Input", "Units.Stress.Internal", "Units.Stress.Output", _
        "Units.Curvature.Input", "Units.Curvature.Internal", "Units.Curvature.Output", _
        "Sign.N.User", "Sign.N.Internal", "Sign.Mx.User", "Sign.Mx.Internal", "Sign.My.User", "Sign.My.Internal")
    AssertRequiredKeys stats, requiredKeys

    requiredKeys = Array( _
        "General.ExecutionReportEnabled", "General.NonCriticalMessagesEnabled", "Calculation.ZeroMomentPerDepth", _
        "Geometry.Source", "Geometry.Type", "RectSet.B1", "RectSet.H1", "RectSet.B2", "RectSet.H2", _
        "Mesh.StepX", "Mesh.StepY", "Mesh.BoundarySubdivisions", _
        "Load.ReferenceOffsetX", "Load.ReferenceOffsetY", _
        "RectSet.H1.as_1", "RectSet.H1.as_2", "RectSet.H1.d_1", "RectSet.H1.d_2", _
        "RectSet.H1.n_1", "RectSet.H1.n_2", _
        "RectSet.H1.StartOffset1", "RectSet.H1.EndOffset1", _
        "RectSet.H1.StartOffset2", "RectSet.H1.EndOffset2", _
        "RectSet.H1.d_2row_1", "RectSet.H1.d_2row_2", "RectSet.H1.d_3row_1", "RectSet.H1.d_3row_2", "RectSet.H1.loc_2row", "RectSet.H1.loc_3row", _
        "RectSet.H2.as_1", "RectSet.H2.as_2", "RectSet.H2.d_1", "RectSet.H2.d_2", _
        "RectSet.H2.n_1", "RectSet.H2.n_2", _
        "RectSet.H2.StartOffset1", "RectSet.H2.EndOffset1", _
        "RectSet.H2.StartOffset2", "RectSet.H2.EndOffset2", _
        "RectSet.H2.d_2row_1", "RectSet.H2.d_2row_2", "RectSet.H2.d_3row_1", "RectSet.H2.d_3row_2", "RectSet.H2.loc_2row", "RectSet.H2.loc_3row", _
        "RectSet.B1.as_1", "RectSet.B1.as_2", "RectSet.B1.d_1", "RectSet.B1.d_2", _
        "RectSet.B1.n_1", "RectSet.B1.n_2", _
        "RectSet.B1.StartOffset1", "RectSet.B1.EndOffset1", _
        "RectSet.B1.StartOffset2", "RectSet.B1.EndOffset2", _
        "RectSet.B1.d_2row_1", "RectSet.B1.d_2row_2", "RectSet.B1.d_3row_1", "RectSet.B1.d_3row_2", "RectSet.B1.loc_2row", "RectSet.B1.loc_3row", _
        "RectSet.B2.as_1", "RectSet.B2.as_2", "RectSet.B2.d_1", "RectSet.B2.d_2", _
        "RectSet.B2.n_1", "RectSet.B2.n_2", _
        "RectSet.B2.StartOffset1", "RectSet.B2.EndOffset1", _
        "RectSet.B2.StartOffset2", "RectSet.B2.EndOffset2", _
        "RectSet.B2.d_2row_1", "RectSet.B2.d_2row_2", "RectSet.B2.d_3row_1", "RectSet.B2.d_3row_2", "RectSet.B2.loc_2row", "RectSet.B2.loc_3row")
    AssertRequiredKeys stats, requiredKeys

    requiredKeys = Array( _
        "Concrete.Rb.ULS", "Concrete.Rbt.ULS", "Concrete.Rb.SLS", "Concrete.Rbt.SLS", "Concrete.Rb.mc2", _
        "Concrete.Eb", "Concrete.Ebt", "Concrete.Eb1Red", "Concrete.Ebt1Red", _
        "Concrete.Eb0", "Concrete.Ebt0", "Concrete.Eb2", "Concrete.Ebt2", _
        "Steel.Rsc.ULS", "Steel.Rs.ULS", "Steel.Rsc.SLS", "Steel.Rs.SLS", _
        "Steel.Esc", "Steel.Es", "Steel.TwoLine.Esc2", "Steel.TwoLine.Es2", _
        "Steel.ThreeLine.Esc2", "Steel.ThreeLine.Es2", _
        "Solver.Method", "Solver.MaxIterations", "Solver.LoadSteps", _
        "Solver.DirectState.DiagramExtension", _
        "Solver.ToleranceN", "Solver.ToleranceMx", "Solver.ToleranceMy", _
        "Solver.LineSearchEnabled", "Solver.DampingInitial", "Solver.MinLineSearchAlpha", _
        "Solver.MaxDeltaEpsilon0", "Solver.MaxDeltaKappa", _
        "Solver.SecantMaxRestarts", "Solver.SecantMinStepNorm", _
        "Capacity.SolutionStrategy", "Capacity.SearchMethod", "Capacity.InitialLambda", "Capacity.MaxLambda", "Capacity.ToleranceLambda", _
        "Capacity.ToleranceStrain", _
        "Capacity.MaxRetries", "Capacity.BaseLoadSteps", "Capacity.SolverMaxIterations", _
        "SLS.Crack.Allowable", "SLS.Crack.InitiationLoadPath", _
        "SLS.Crack.Phi1", "SLS.Crack.Phi2", "SLS.Crack.Phi3Mode", "SLS.Crack.Phi3", _
        "SLS.Crack.PsiMode", "SLS.Crack.PsiS", "SLS.Crack.TensionZoneMode", _
        "SLS.Crack.CoverDistanceMode")
    AssertRequiredKeys stats, requiredKeys

    requiredKeys = Array( _
        "AutoCAD.Export.CombinationID", "AutoCAD.Export.NeutralLineEnabled", _
        "AutoCAD.Export.PrincipalAxesMode", "AutoCAD.Export.LoadPointEnabled", _
        "AutoCAD.Export.ContourEnabled", "AutoCAD.Export.LabelMode", _
        "AutoCAD.Layer.Concrete", "AutoCAD.Layer.Rebar", "AutoCAD.Layer.Contour", _
        "AutoCAD.Layer.ConcreteTension", "AutoCAD.Layer.ConcreteCompression", _
        "AutoCAD.Layer.RebarTension", "AutoCAD.Layer.RebarCompression", _
        "AutoCAD.Color.ConcreteTension", "AutoCAD.Color.ConcreteCompression", _
        "AutoCAD.Color.RebarTension", "AutoCAD.Color.RebarCompression", _
        "AutoCAD.Color.Neutral", _
        "AutoCAD.Import.ConcreteLayer", "AutoCAD.Import.RebarLayer", "AutoCAD.Import.MinArea")
    AssertRequiredKeys stats, requiredKeys

    requiredKeys = Array( _
        "Plot.Enabled", "Plot.AutoUpdateAfterCalculation", "Plot.LoadCase", _
        "Plot.ResultGradient", "Plot.ResultLabelsEnabled", "Plot.ResultLabelSpacing", "Plot.ResultPrecision", _
        "Plot.NeutralLineEnabled", "Plot.PrincipalAxesMode", "Plot.LoadApplicationPointEnabled", _
        "Plot.AxisLabelsEnabled", "Plot.AxisLabelsFontSize", _
        "Plot.ContourEnabled", "Plot.LegendEnabled", _
        "Plot.RebarLabels.Enabled", "Plot.Dimensions.Enabled", _
        "Plot.RebarLabels.Placement", "Plot.Dimensions.Placement", _
        "Plot.RebarLabels.Offset", "Plot.Dimensions.Offset", _
        "Plot.RebarLabels.TextUnits", "Plot.Dimensions.TextUnits", _
        "Plot.RebarLabels.TextHeight", "Plot.Dimensions.TextHeight", _
        "Plot.RebarLabels.TextGap", "Plot.Dimensions.TextGap", _
        "Plot.RebarLabels.LineEnabled", "Plot.Dimensions.ArrowType", "Plot.Dimensions.ArrowSize")
    AssertRequiredKeys stats, requiredKeys

    Dim i As Long
    Dim removedKeys As Variant
    removedKeys = Array("Capacity.Enabled", "Capacity.CalculateMx", "Capacity.CalculateMy", _
        "Capacity.CalculateMxy", "Mesh.BoundaryMode", "Mesh.Step", "Circle.Radius", _
        "Batch.MaxCombinations", "Batch.Diagnostics", "Materials.SourceStatus", _
        "Concrete.Diagram", "Steel.Diagram", "Concrete.Point1.Eps", _
        "Concrete.Point1.Stress", "Concrete.Point2.Eps", "Concrete.Point2.Stress", _
        "Concrete.Point3.Eps", "Concrete.Point3.Stress", _
        "Steel.Point1.Eps", "Steel.Point1.Stress", "Steel.Point2.Eps", _
        "Steel.Point2.Stress", "Steel.Point3.Eps", "Steel.Point3.Stress", _
        "Solver.DiagnosticsEnabled", "Circle.CenterX", "Circle.CenterY", _
        "RectSet.OriginX", "RectSet.OriginY", "Plot.DimensionsEnabled", "Plot.RebarLabelsEnabled", _
        "Plot.PrincipalAxesEnabled", "Plot.CentroidEnabled", "AutoCAD.Export.PrincipalAxesEnabled", _
        "Concrete.TensionMode", "Capacity.ConcreteCompressionLimit", "Capacity.ConcreteTensionLimit", _
        "Capacity.SteelStrainLimit", "Concrete.Class", _
        "Steel.RebarProfile", _
        "Calculation.Mode", "Capacity.CalculationScope", "SLS.Crack.Enabled", _
        "AutoCAD.Export.ResultType", "Plot.ResultType", _
        "Diagram.Strength.Concrete", "Diagram.Strength.ConcreteTension", "Diagram.Strength.Steel", _
        "Diagram.Mcrc.Concrete", "Diagram.Mcrc.ConcreteTension", "Diagram.Mcrc.Steel", _
        "Diagram.CrackedNDS.Concrete", "Diagram.CrackedNDS.ConcreteTension", "Diagram.CrackedNDS.Steel")
    For i = LBound(removedKeys) To UBound(removedKeys)
        AssertTrue stats, "settings.removed." & CStr(removedKeys(i)), Not reader.HasKey(CStr(removedKeys(i)))
    Next i

    Dim rangeNames As Variant
    rangeNames = Array("rngUnitSettings", "rngSignConventionSettings", _
        "rngConcreteMaterialParameters", "rngSteelMaterialParameters", "rngCalculationProfiles", _
        "rngCircleGeometry", "rngRoundedRectangleGeometry", "rngHollowRectangleGeometry", "rngRectSetGeometry", _
        "rngNDMSectionProperties", "rngNDMSectionAnnotations", "rngNDMMaterialDiagrams")
    For i = LBound(rangeNames) To UBound(rangeNames)
        AssertTrue stats, "settings.range." & CStr(rangeNames(i)), NamedRangeExists(CStr(rangeNames(i)))
    Next i

    Dim removedRanges As Variant
    removedRanges = Array("SolverSettings", "CapacitySettings", "ConcreteDiagram", _
        "SteelDiagram", "GeometrySettings", "OutputSettings", "AutoCADSettings", _
        "rngMainInput", "rngRebarInput", "rngConcreteDiagramPoints", "rngSteelDiagramPoints", _
        "rngCalculationDiagramSettings")
    For i = LBound(removedRanges) To UBound(removedRanges)
        AssertTrue stats, "settings.range.removed." & CStr(removedRanges(i)), Not NamedRangeExists(CStr(removedRanges(i)))
    Next i
End Sub

Private Sub AssertRequiredKeys(ByRef stats As TSectionSolverTestStats, ByVal requiredKeys As Variant)
    Dim reader As CSystemSettingsReader
    Set reader = New CSystemSettingsReader
    reader.LoadFromWorkbook ThisWorkbook

    Dim i As Long
    For i = LBound(requiredKeys) To UBound(requiredKeys)
        AssertTrue stats, "settings.key." & CStr(requiredKeys(i)), reader.HasKey(CStr(requiredKeys(i)))
    Next i
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestLinearMaterialEquilibrium(ByRef stats As TSectionSolverTestStats)
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

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    ConfigureStrictSolver solver
    solver.Solve BuildGeneratedSectionModel(mesh, Nothing), concrete, steel, n, mx, my

    AssertTrue stats, "section.linear.converged", solver.Converged
    AssertRelative stats, "section.linear.eps0", solver.Epsilon0, 0.000416666667, 0.00000001
    AssertRelative stats, "section.linear.kappaX", solver.KappaX, 0.000242424242, 0.00000001
    AssertRelative stats, "section.linear.kappaY", solver.KappaY, -0.000040100251, 0.00000001
    AssertEquilibrium stats, "section.linear", solver, n, mx, my
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestLinearMaterialWithRebarReplacement(ByRef stats As TSectionSolverTestStats)
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

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    ConfigureStrictSolver solver
    solver.Solve BuildGeneratedSectionModel(mesh, rebars), concrete, steel, n, mx, my

    AssertTrue stats, "section.rebar.converged", solver.Converged
    AssertRelative stats, "section.rebar.eps0", solver.Epsilon0, 0.00061453123, 0.00000001
    AssertRelative stats, "section.rebar.kappaX", solver.KappaX, 0.000130953764, 0.00000001
    AssertRelative stats, "section.rebar.kappaY", solver.KappaY, 0.000025325048, 0.00000001
    AssertEquilibrium stats, "section.rebar", solver, n, mx, my
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestDiagramConcreteCentralCompression(ByRef stats As TSectionSolverTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = RectangleGeometry(200#, 100#)
    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 10#)

    Dim concrete As CMaterialDiagram
    Set concrete = ProvisionalConcrete()
    Dim steel As CMaterialDiagram
    Set steel = ProvisionalSteel()

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    ConfigureProvisionalSolver solver
    solver.Solve BuildGeneratedSectionModel(mesh, Nothing), concrete, steel, -100000#, 0#, 0#

    AssertTrue stats, "section.diagramCompression.converged", solver.Converged
    AssertTrue stats, "section.diagramCompression.epsNegative", solver.Epsilon0 < 0#
    AssertClose stats, "section.diagramCompression.kappaX", solver.KappaX, 0#, 0.000000000001
    AssertClose stats, "section.diagramCompression.kappaY", solver.KappaY, 0#, 0.000000000001
    AssertEquilibrium stats, "section.diagramCompression", solver, -100000#, 0#, 0#
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
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
    solver.Solve BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, -20000000#, 0#

    AssertTrue stats, "section.diagramRebar.converged", solver.Converged
    AssertEquilibrium stats, "section.diagramRebar", solver, -300000#, -20000000#, 0#
    AssertTrue stats, "section.diagramRebar.steps", solver.LoadStepsCompleted = 5
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestIncrementLimitsAndDiagnostics(ByRef stats As TSectionSolverTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = RectangleGeometry(200#, 100#)
    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 10#)

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    ConfigureProvisionalSolver solver
    solver.MaxDeltaEpsilon0 = 0.00002
    solver.Solve BuildGeneratedSectionModel(mesh, Nothing), ProvisionalConcrete(), ProvisionalSteel(), -80000#, 0#, 0#

    AssertTrue stats, "section.diagnostics.converged", solver.Converged
    AssertTrue stats, "section.diagnostics.iterations", solver.Iterations > solver.LoadStepsCompleted
    AssertTrue stats, "section.diagnostics.log", InStr(1, solver.DiagnosticLog, "iter=", vbTextCompare) > 0
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestSecantIndependentBranch(ByRef stats As TSectionSolverTestStats)
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

    Dim newton As CSectionSolver
    Set newton = New CSectionSolver
    ConfigureProvisionalSolver newton
    newton.SolverMethod = "Newton"
    newton.Solve BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, -20000000#, -12000000#

    Dim secant As CSectionSolver
    Set secant = New CSectionSolver
    ConfigureProvisionalSolver secant
    secant.SolverMethod = "Secant"
    secant.SecantMaxRestarts = 4
    secant.Solve BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), -300000#, -20000000#, -12000000#

    AssertTrue stats, "section.secant.branch", StrComp(secant.SolverMethod, "Secant", vbTextCompare) = 0
    AssertTrue stats, "section.secant.converged", secant.Converged
    AssertTrue stats, "section.secant.noNewtonCalls", secant.InternalNewtonCallCount = 0
    AssertTrue stats, "section.newton.calls", newton.InternalNewtonCallCount > 0
    AssertEquilibrium stats, "section.secant.equilibrium", secant, -300000#, -20000000#, -12000000#
    AssertRelative stats, "section.secant.eps0MatchesNewton", secant.Epsilon0, newton.Epsilon0, 0.0001
    AssertRelative stats, "section.secant.kappaXMatchesNewton", secant.KappaX, newton.KappaX, 0.0001
    AssertRelative stats, "section.secant.kappaYMatchesNewton", secant.KappaY, newton.KappaY, 0.0001
    AssertTrue stats, "section.secant.differentIterations", secant.Iterations <> newton.Iterations
    AssertTrue stats, "section.secant.diagnostics", InStr(1, secant.DiagnosticLog, "method=Secant", vbTextCompare) > 0
    AssertTrue stats, "section.secant.broydenDiagnostics", secant.InternalForceEvaluationCount > secant.Iterations
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestSolverMethodInputErrors(ByRef stats As TSectionSolverTestStats)
    Dim geom As CGeometryRoundedRectangle
    Set geom = RectangleGeometry(200#, 100#)
    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 20#)

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    ConfigureProvisionalSolver solver
    solver.SolverMethod = "Bogus"
    solver.Solve BuildGeneratedSectionModel(mesh, Nothing), ProvisionalConcrete(), ProvisionalSteel(), -100000#, 0#, 0#
    AssertTrue stats, "section.method.invalidInput", Not solver.Converged
    AssertTrue stats, "section.method.invalidMessage", InStr(1, solver.StopReason, "InputError", vbTextCompare) > 0

    Set solver = New CSectionSolver
    ConfigureProvisionalSolver solver
    solver.SolverMethod = vbNullString
    solver.Solve BuildGeneratedSectionModel(mesh, Nothing), ProvisionalConcrete(), ProvisionalSteel(), -100000#, 0#, 0#
    AssertTrue stats, "section.method.emptyInput", Not solver.Converged
    AssertTrue stats, "section.method.emptyMessage", InStr(1, solver.StopReason, "InputError", vbTextCompare) > 0
End Sub

' Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией.
Private Sub TestSolverMethodFromSystem(ByRef stats As TSectionSolverTestStats)
    Dim reader As CSystemSettingsReader
    Set reader = New CSystemSettingsReader
    reader.LoadFromWorkbook ThisWorkbook
    AssertTrue stats, "settings.solverMethod.defaultNewton", StrComp(reader.GetRawString("Solver.Method", vbNullString), "Newton", vbTextCompare) = 0

    Dim geom As CGeometryRoundedRectangle
    Set geom = RectangleGeometry(200#, 100#)
    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildMesh(geom, 20#)

    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    solver.ApplySettings reader
    solver.Solve BuildGeneratedSectionModel(mesh, Nothing), ProvisionalConcrete(), ProvisionalSteel(), -100000#, 0#, 0#
    AssertTrue stats, "section.systemMethod.newtonBranch", solver.Converged And solver.InternalNewtonCallCount > 0
End Sub

Private Sub AssertNewtonSecantCase(ByRef stats As TSectionSolverTestStats, ByVal prefix As String, _
        ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, _
        ByVal n As Double, ByVal mx As Double, ByVal my As Double)
    Dim newton As CSectionSolver
    Set newton = New CSectionSolver
    ConfigureProvisionalSolver newton
    newton.SolverMethod = "Newton"
    newton.MaxIterations = 80
    newton.Solve BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), n, mx, my

    Dim secant As CSectionSolver
    Set secant = New CSectionSolver
    ConfigureProvisionalSolver secant
    secant.SolverMethod = "Secant"
    secant.MaxIterations = 80
    secant.SecantMaxRestarts = 6
    secant.Solve BuildGeneratedSectionModel(mesh, rebars), ProvisionalConcrete(), ProvisionalSteel(), n, mx, my

    AssertTrue stats, prefix & ".newton.converged", newton.Converged
    AssertTrue stats, prefix & ".secant.converged", secant.Converged
    AssertTrue stats, prefix & ".secant.noNewtonCalls", secant.InternalNewtonCallCount = 0
    If newton.Converged And secant.Converged Then
        AssertRelative stats, prefix & ".eps0", secant.Epsilon0, newton.Epsilon0, 0.001
        AssertLoadComponent stats, prefix & ".N", secant.Nint, n, 1#, 0.000001
        AssertLoadComponent stats, prefix & ".Mx", secant.Mxint, mx, 1000#, 0.000001
        AssertLoadComponent stats, prefix & ".My", secant.Myint, my, 1000#, 0.000001
    End If
    AppendLine stats, "COMPARE_SOLVER|" & prefix & "|NewtonIterations=" & CStr(newton.Iterations) & _
        "|SecantIterations=" & CStr(secant.Iterations) & _
        "|SecantForceEvaluations=" & CStr(secant.InternalForceEvaluationCount) & _
        "|SecantNewtonCalls=" & CStr(secant.InternalNewtonCallCount) & _
        "|SecantRestarts=" & CStr(secant.MatrixRestartCount)
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

Private Function ProvisionalConcrete() As CMaterialDiagram
    Dim concrete As CMaterialDiagram
    Set concrete = New CMaterialDiagram
    concrete.Initialize -0.0015, -15.5, -0.0035, -15.5
    Set ProvisionalConcrete = concrete
End Function

Private Function ProvisionalSteel() As CMaterialDiagram
    Dim steel As CMaterialDiagram
    Set steel = New CMaterialDiagram
    steel.Initialize 0.00175, 350#, 0.025
    Set ProvisionalSteel = steel
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







