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

' Выполняет регрессии равновесия, единиц, численных методов и State-контрактов.
' Возвращает все assertions и итог; ошибка suite остается отдельной диагностикой.
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
    TestAudit02EvaluatedPlaneRequiresEquilibrium stats
    TestAudit02DiagramExtensionReaderMigration stats
    TestAudit02StateSnapshotIsolation stats
    TestAudit02AbsentStressSign stats
    TestAudit02LoadPathResidualScaling stats
    TestAudit02ExtendedInitialGuessPhysicalFinal stats
    TestAudit03TypedStateFailures stats
    TestAudit03RetryAttemptSession stats

    AppendLine stats, "TOTAL_SECTION_SOLVER: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & _
        "; elapsedSec=" & FormatNumberInvariant(Timer - t0)
    RunSectionSolverTests = stats.Report
    Exit Function

Failed:
    RunSectionSolverTests = "RUNTIME ERROR: " & CStr(Err.Number) & _
        "; source=" & Err.Source & "; description=" & Err.Description
End Function

' Проверяет пересчет пользовательских tf/tf*m и знака сжатия во внутренние
' Н/Н*мм, а также обратный вывод и размерность порога малых моментов.
Private Sub TestUnitSystemConversions(ByRef stats As TSectionSolverTestStats)
    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.InitializeDefaults

    AssertClose stats, "units.length.mm.m", units.InternalLengthToOutput(units.InputLengthToInternal(25#)), 0.025, 0.000000000001

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

' Сверяет решение невырожденной системы 3x3 с известным вектором и невязкой.
' Вырожденная система должна вернуть отказ, а не фиктивную поправку Newton.
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

' Читает фактические именованные диапазоны и проверяет модули материалов,
' наличие профилей и различие активной tensile-ветви Strength/CrackInitiation.
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

' Сравнивает независимый Secant с Newton для сжатия и одно-/двухосного изгиба
' прямоугольника и круга; обе ветви должны подтвердить компонентное равновесие.
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

' Проверяет наличие обязательных ключей и диапазонов действующей схемы Config,
' отсутствие дубликатов и удаленных полей. Это структурная, не поведенческая приемка.
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
        "General.DiagramExtension", _
        "Solver.ToleranceN", "Solver.ToleranceMx", "Solver.ToleranceMy", _
        "Solver.LineSearchEnabled", "Solver.DampingInitial", "Solver.MinLineSearchAlpha", _
        "Solver.MaxDeltaEpsilon0", "Solver.MaxDeltaKappa", _
        "Solver.SecantMaxRestarts", "Solver.SecantMinStepNorm", _
        "Capacity.SolutionStrategy", "Capacity.SearchMethod", "Capacity.InitialLambda", "Capacity.MaxLambda", "Capacity.ToleranceLambda", _
        "Capacity.ToleranceStrain", _
        "Capacity.MaxRetries", "Capacity.BaseLoadSteps", "Capacity.SolverMaxIterations", _
        "SLS.Crack.Allowable", "SLS.Crack.InitiationLoadPath", _
        "SLS.Crack.Phi1", "SLS.Crack.Phi2", "SLS.Crack.Phi3Mode", "SLS.Crack.Phi3", _
        "SLS.Crack.PsiMode", "SLS.Crack.SigmaSCrcAveragingMode", "SLS.Crack.PsiS", "SLS.Crack.TensionZoneMode", _
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
        "Plot.ResultGradient", "Plot.ResultLabelsEnabled", "Plot.ResultLabelSpacing", _
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

' Проверяет каждый переданный ключ в настоящем reader-е книги и сохраняет
' отдельный assertion, чтобы отсутствие настройки не скрывалось общим итогом.
Private Sub AssertRequiredKeys(ByRef stats As TSectionSolverTestStats, ByVal requiredKeys As Variant)
    Dim reader As CSystemSettingsReader
    Set reader = New CSystemSettingsReader
    reader.LoadFromWorkbook ThisWorkbook

    Dim i As Long
    For i = LBound(requiredKeys) To UBound(requiredKeys)
        AssertTrue stats, "settings.key." & CStr(requiredKeys(i)), reader.HasKey(CStr(requiredKeys(i)))
    Next i
End Sub

' Сверяет плоскость деформаций линейного бетонного прямоугольника с заранее
' рассчитанными epsilon0/kappaX/kappaY и независимо проверяет N/Mx/My.
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

' Проверяет линейное армированное сечение с замещением бетона площадью стержней:
' эталонные параметры плоскости и равновесие должны учитывать именно эту модель.
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

' Проверяет центральное сжатие по нелинейной диаграмме: отрицательную осевую
' деформацию, нулевые кривизны и равновесие без искусственного изгиба.
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

' Решает внецентренное сжатие армированного прямоугольника по диаграммам;
' проверяет равновесие и прохождение всех пяти ступеней заданного нагружения.
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

' Ограничивает приращение epsilon0 и проверяет сходимость с дополнительными
' итерациями, а также наличие реальных итерационных строк в диагностике.
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

' Доказывает, что Secant не вызывает Newton: сравнивает плоскости и равновесие,
' проверяет счетчики вычислений, рестарты и собственную диагностику метода.
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

' Передает неизвестный и пустой Solver.Method; оба значения должны дать
' ошибку ввода, а не незаявленное переключение на метод по умолчанию.
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

' Применяет настройки из Config к реальному solver-у и проверяет, что
' выбранный Newton действительно выполняет свою ветвь и находит равновесие.
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

' Решает одну и ту же задачу двумя методами с одинаковыми материалами;
' сохраняет сравнительные iterations/evaluations и проверяет независимый Secant.
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

' Строит бетонную сетку тестового прямоугольника с одинаковым шагом по осям.
' Арматура добавляется отдельно, а ядру затем передается собранный CSectionModel.
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

' ДЛЯ ТЕСТОВ
' Запускает направленные проверки typed failure через настоящий provider,
' runner и solver, отдельно от полного набора задач равновесия.
Public Function RunAudit03StateTests() As String
    Dim stats As TSectionSolverTestStats
    TestAudit03TypedStateFailures stats
    AppendLine stats, "TOTAL_AUDIT03_STATE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03StateTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ
' Ошибка метода, вырожденная геометрия, бюджет итераций и отсутствующий
' контекст не должны терять свою причину при формировании named-state.
Private Sub TestAudit03TypedStateFailures(ByRef stats As TSectionSolverTestStats)
    Dim scenario As Long
    For scenario = 1 To 11
        TestAudit03TypedStateFailureCase stats, scenario
    Next scenario
End Sub

' ДЛЯ ТЕСТОВ
' Создает изолированный реальный маршрут одного отказа. Неуспешный snapshot
' сохраняется для диагностики, но не становится reusable и допускает retry.
Private Sub TestAudit03TypedStateFailureCase(ByRef stats As TSectionSolverTestStats, ByVal scenario As Long)
    On Error GoTo Failed
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Dim materials As CMaterialModelProvider
    Set materials = New CMaterialModelProvider
    materials.Initialize settings
    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(BuildMesh(RectangleGeometry(200#, 100#), 20#), Nothing)
    If scenario = 2 Then
        Set section = New CSectionModel
        section.AddConcreteElement 0#, 0#, 10000#
    ElseIf scenario = 5 Then
        Set section = New CSectionModel
    End If
    Dim repository As CStateRepository
    Set repository = New CStateRepository
    Dim provider As CStateProvider
    Set provider = New CStateProvider
    If scenario <> 4 Then provider.Initialize section, materials, repository
    If scenario = 1 Then provider.SolverMethod = "UnknownMethod"
    If scenario = 3 Then provider.MaxIterations = 1
    If scenario = 6 Then provider.MaxIterations = 0
    If scenario = 7 Then provider.MinLineSearchAlpha = 0#
    If scenario = 8 Then provider.DampingInitial = 0#
    If scenario = 9 Then provider.MinLineSearchAlpha = 2#
    If scenario = 10 Then provider.MaxDeltaEpsilon0 = -1#
    If scenario = 11 Then provider.MaxDeltaKappa = -1#
    Dim spec As CMaterialModelSpec
    Set spec = New CMaterialModelSpec
    spec.Initialize "ULS(I)", "ThreeLine", "Ignore", "TwoLine"
    Dim request As CStateRequest
    Set request = New CStateRequest
    request.Initialize sstStrengthState, cpStrength, spec, -100000#, 10000000#, 2000000#, False, True
    Dim state As CSectionStateResult
    Set state = provider.GetOrSolve(request)
    Dim prefix As String
    prefix = "audit03.state.failure." & CStr(scenario)
    AssertTrue stats, prefix & ".exists", Not state Is Nothing
    If state Is Nothing Then Exit Sub
    Dim expectedStatus As EResultInternalStatus
    Dim expectedCode As EResultCode
    Select Case scenario
        Case 1, 6 To 11: expectedStatus = rsInvalidConfiguration: expectedCode = rcInvalidConfiguration
        Case 2: expectedStatus = rsNumericalFailure: expectedCode = rcSingularTangent
        Case 3: expectedStatus = rsNumericalFailure: expectedCode = rcNumericalFailure
        Case 4: expectedStatus = rsInternalError: expectedCode = rcInternalError
        Case 5: expectedStatus = rsInvalidInput: expectedCode = rcInvalidInput
    End Select
    AssertTrue stats, prefix & ".status", state.InternalStatus = expectedStatus
    AssertTrue stats, prefix & ".code", state.ResultCode = expectedCode
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    Dim expectedExternal As String
    expectedExternal = "InputErr"
    If scenario = 2 Or scenario = 3 Then expectedExternal = "NumFail"
    If scenario = 4 Then expectedExternal = "CalcErr"
    AssertTrue stats, prefix & ".display", policy.ExternalStatus(state.ResultMeta) = expectedExternal
    AssertTrue stats, prefix & ".comment", Len(Trim$(state.ResultComment)) > 0
    AppendLine stats, "COMMENT: " & prefix & "; " & state.ResultComment
    AssertTrue stats, prefix & ".notConverged", Not state.Converged
    AssertTrue stats, prefix & ".notReusable", repository.FindEquivalent(request) Is Nothing
    If scenario = 1 Or scenario = 5 Or scenario >= 6 Then
        AssertTrue stats, prefix & ".oneAttempt", state.SolverCallCount = 1
        AssertTrue stats, prefix & ".notCalculated", Not state.ResultMeta.Calculated
    ElseIf scenario = 4 Then
        AssertTrue stats, prefix & ".noAttempt", state.SolverCallCount = 0
        AssertTrue stats, prefix & ".notCalculated", Not state.ResultMeta.Calculated
    Else
        AssertTrue stats, prefix & ".attempted", state.SolverCallCount >= 1
        AssertTrue stats, prefix & ".calculated", state.ResultMeta.Calculated
    End If
    If scenario <> 4 Then
        AssertTrue stats, prefix & ".stored", repository.StateCount = 1
        AssertTrue stats, prefix & ".snapshotCode", repository.StateAt(1).ResultCode = expectedCode
        AssertTrue stats, prefix & ".snapshotComment", repository.StateAt(1).ResultComment = state.ResultComment
    End If
    If scenario = 1 Then
        provider.SolverMethod = "Newton"
        request.Initialize sstStrengthState, cpStrength, spec, -50000#, 0#, 0#, False, True
        Set state = provider.GetOrSolve(request)
        AssertTrue stats, prefix & ".newAttemptSuccess", state.Converged
        AssertTrue stats, prefix & ".notReused", Not provider.LastStateWasReused
        Set state = provider.GetOrSolve(request)
        AssertTrue stats, prefix & ".successReusable", provider.LastStateWasReused
        AssertTrue stats, prefix & ".successSnapshot", Not provider.SolverSnapshot(state) Is Nothing
    End If
    Exit Sub
Failed:
    AssertTrue stats, "audit03.state.failure." & CStr(scenario) & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' ============================== ДЛЯ ТЕСТОВ ==============================

' Проверяет различие оценки плоскости и решения равновесия. Подтверждение
' по целевым усилиям не должно менять плоскость или запускать новый solve.
Private Sub TestAudit02EvaluatedPlaneRequiresEquilibrium(ByRef stats As TSectionSolverTestStats)
    Dim concrete As CLinearConcreteMaterial
    Set concrete = New CLinearConcreteMaterial
    concrete.Initialize 32500#
    Dim steel As CLinearSteelMaterial
    Set steel = New CLinearSteelMaterial
    steel.Initialize 200000#
    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(BuildMesh(RectangleGeometry(200#, 100#), 20#), Nothing)
    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    solver.ToleranceN = 1#
    solver.ToleranceMx = 1#
    solver.ToleranceMy = 1#
    Dim solveCount As Long
    solveCount = SectionEquilibriumSolveCount()
    solver.EvaluateStrainPlane section, concrete, steel, -0.0001, 0#, 0#
    AssertTrue stats, "audit02.plane.evaluationNotSolve", Not solver.Converged
    AssertTrue stats, "audit02.plane.wrongTargetRejected", _
        Not solver.ConfirmEquilibrium(solver.Nint + 100#, solver.Mxint, solver.Myint)
    AssertTrue stats, "audit02.plane.actualTargetConfirmed", _
        solver.ConfirmEquilibrium(solver.Nint, solver.Mxint, solver.Myint)
    AssertClose stats, "audit02.plane.sameStrain", solver.Epsilon0, -0.0001, 0#
    AssertTrue stats, "audit02.plane.noHeavySolve", SectionEquilibriumSolveCount() = solveCount
End Sub

' Проверяет миграцию на входной границе reader-а: старое No не перекрывается
' default, новый ключ имеет приоритет, конфликт виден, ошибочный ввод отклонен.
' Runtime-контекст после чтения содержит только канонический ключ.
Private Sub TestAudit02DiagramExtensionReaderMigration(ByRef stats As TSectionSolverTestStats)
    On Error GoTo Failed
    Dim testBook As Object
    Set testBook = ThisWorkbook.Application.Workbooks.Add(-4167)
    Dim testRange As Object
    Set testRange = testBook.Worksheets.Item(1).Range("A1:C5")
    Dim scenario As Variant
    Dim data(1 To 5, 1 To 3) As Variant
    Dim reader As CSystemSettingsReader
    Dim errorNumber As Long
    Dim testKey As String
    For Each scenario In Array( _
            Array("legacyNo", True, "No", False, "", "No", False, False), _
            Array("canonicalNo", False, "", True, "No", "No", False, False), _
            Array("bothConflict", True, "No", True, "Yes", "Yes", True, False), _
            Array("bothEquivalent", True, "1", True, "Yes", "Yes", False, False), _
            Array("missing", False, "", False, "", "Yes", False, False), _
            Array("invalidCanonical", True, "Yes", True, "invalid", "", False, True), _
            Array("emptyCanonical", True, "Yes", True, "", "", False, True), _
            Array("invalidLegacy", True, "invalid", False, "", "", False, True))
        Erase data
        data(1, 1) = "Параметр": data(1, 2) = "Значение": data(1, 3) = "Default"
        If CBool(scenario(1)) Then
            data(2, 1) = "Solver.DirectState.DiagramExtension": data(2, 2) = scenario(2)
        End If
        If CBool(scenario(3)) Then
            data(3, 1) = "General.DiagramExtension": data(3, 2) = scenario(4)
        End If
        data(4, 1) = "Solver.MaxIterations": data(4, 2) = "37"
        testRange.Value2 = data
        testKey = "audit02.migration.reader." & CStr(scenario(0))
        errorNumber = 0
        Set reader = Audit02LoadMigrationRange(testRange, errorNumber)
        If CBool(scenario(7)) Then
            AssertTrue stats, testKey & ".inputError", errorNumber = vbObjectError + 4310
        Else
            AssertTrue stats, testKey & ".loaded", errorNumber = 0
            If errorNumber = 0 Then
                AssertTrue stats, testKey & ".value", reader.GetRawString("General.DiagramExtension") = CStr(scenario(5))
                AssertTrue stats, testKey & ".legacyRemoved", Not reader.HasKey("Solver.DirectState.DiagramExtension")
                AssertTrue stats, testKey & ".warning", (Len(reader.MigrationWarning) > 0) = CBool(scenario(6))
                AssertTrue stats, testKey & ".neighbor", reader.GetRequiredDouble("Solver.MaxIterations") = 37#
            End If
        End If
    Next scenario
    testBook.Close False
    Exit Sub
Failed:
    Dim reason As String
    reason = Err.Description
    If Not testBook Is Nothing Then testBook.Close False
    AssertTrue stats, "audit02.migration.reader.runtime: " & reason, False
End Sub

' Перехватывает только ожидаемую ошибку ввода отдельного migration-сценария.
' Ошибка возвращается числом, без обратного определения статуса по ее тексту.
Private Function Audit02LoadMigrationRange(ByVal testRange As Object, ByRef errorNumber As Long) As CSystemSettingsReader
    On Error GoTo Failed
    Dim reader As CSystemSettingsReader
    Set reader = New CSystemSettingsReader
    reader.LoadFromRange testRange
    Set Audit02LoadMigrationRange = reader
    Exit Function
Failed:
    errorNumber = Err.Number
End Function

' Проверяет независимость чисел, метаданных и ключа запроса от дальнейшего
' использования рабочего solver-а и изменения выдаваемых spec/meta-копий.
Private Sub TestAudit02StateSnapshotIsolation(ByRef stats As TSectionSolverTestStats)
    Dim concrete As CLinearConcreteMaterial
    Set concrete = New CLinearConcreteMaterial
    concrete.Initialize 32500#
    Dim steel As CLinearSteelMaterial
    Set steel = New CLinearSteelMaterial
    steel.Initialize 200000#
    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(BuildMesh(RectangleGeometry(200#, 100#), 20#), Nothing)
    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    solver.EvaluateStrainPlane section, concrete, steel, -0.0001, 0#, 0#
    AssertTrue stats, "audit02.stateSnapshot.equilibrium", _
        solver.ConfirmEquilibrium(solver.Nint, solver.Mxint, solver.Myint)

    Dim spec As CMaterialModelSpec
    Set spec = New CMaterialModelSpec
    spec.Initialize "ULS(I)", "TwoLine", "Ignore", "TwoLine"
    Dim state As CSectionStateResult
    Set state = New CSectionStateResult
    state.InitializeFromSolver sstStrengthState, cpStrength, spec, solver, False, True
    Dim request As CStateRequest
    Set request = New CStateRequest
    request.Initialize sstStrengthState, cpStrength, spec, solver.Nint, solver.Mxint, solver.Myint, False
    Dim originalKey As String
    originalKey = request.EquivalenceKey
    Dim originalN As Double
    originalN = state.Nint

    spec.Initialize "SLS(II)", "ThreeLine", "UseDiagram", "ThreeLine"
    AssertTrue stats, "audit02.stateSnapshot.inputSpec", state.MaterialSpec.SpecKey = "ULS(I)|TwoLine|Ignore|TwoLine"
    Dim copy As CMaterialModelSpec
    Set copy = state.MaterialSpec
    copy.Clear
    AssertTrue stats, "audit02.stateSnapshot.returnedSpec", state.MaterialSpec.IsComplete
    Set copy = request.MaterialSpec
    copy.Clear
    AssertTrue stats, "audit02.stateSnapshot.requestKey", request.EquivalenceKey = originalKey

    Dim meta As CResultMeta
    Set meta = state.ResultMeta
    meta.SetResult rsNumericalFailure, rcNumericalFailure, rkDirectState, "Изменение внешней копии."
    AssertTrue stats, "audit02.stateSnapshot.returnedMeta", state.InternalStatus = rsSuccess
    Dim solveCount As Long
    solveCount = SectionEquilibriumSolveCount()
    solver.EvaluateStrainPlane section, concrete, steel, -0.001, 0#, 0#
    AssertClose stats, "audit02.stateSnapshot.strain", state.Epsilon0, -0.0001, 0#
    AssertClose stats, "audit02.stateSnapshot.force", state.Nint, originalN, 0#
    AssertClose stats, "audit02.stateSnapshot.extremum", state.MinConcreteStrain, -0.0001, 0#
    AssertTrue stats, "audit02.stateSnapshot.noSolve", SectionEquilibriumSolveCount() = solveCount
    Set state = New CSectionStateResult
    AssertTrue stats, "audit02.stateSnapshot.emptyStatus", state.InternalStatus = rsInternalError
    AssertTrue stats, "audit02.stateSnapshot.emptyCode", state.ResultCode = rcInternalError
End Sub

' Проверяет отсутствие фиктивных экстремумов отсутствующего знака напряжений:
' полностью растянутое или сжатое сечение не возвращает sentinel в snapshot.
Private Sub TestAudit02AbsentStressSign(ByRef stats As TSectionSolverTestStats)
    Dim concrete As CLinearConcreteMaterial
    Set concrete = New CLinearConcreteMaterial
    concrete.Initialize 32500#
    Dim steel As CLinearSteelMaterial
    Set steel = New CLinearSteelMaterial
    steel.Initialize 200000#
    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(BuildMesh(RectangleGeometry(200#, 100#), 20#), Nothing)
    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    solver.EvaluateStrainPlane section, concrete, steel, 0.0001, 0#, 0#
    AssertClose stats, "audit02.stressSign.noCompression", solver.MinConcreteStress, 0#, 0#
    AssertClose stats, "audit02.stressSign.tension", solver.MaxConcreteStress, 3.25, 0.000000001
    AssertClose stats, "audit02.stressSign.noSteelCompression", solver.MinSteelStress, 0#, 0#
    AssertClose stats, "audit02.stressSign.noSteelTension", solver.MaxSteelStress, 0#, 0#
    solver.EvaluateStrainPlane section, concrete, steel, -0.0001, 0#, 0#
    AssertClose stats, "audit02.stressSign.noTension", solver.MaxConcreteStress, 0#, 0#
    AssertClose stats, "audit02.stressSign.compression", solver.MinConcreteStress, -3.25, 0.000000001
End Sub

' Проверяет линейность невязки постоянной нулевой компоненты траектории.
' Рост пробного момента не должен менять допуск и обнулять его производную.
Private Sub TestAudit02LoadPathResidualScaling(ByRef stats As TSectionSolverTestStats)
    Dim math As CLoadPathMath
    Set math = New CLoadPathMath
    math.Configure 0.0001, 0.0001, 0.0001
    Dim path As CLoadPathVector
    Set path = New CLoadPathVector
    path.Initialize 0#, 100000#, 0#, 0#, 0#, 0#
    Dim firstMx As Double, firstMy As Double
    Dim secondMx As Double, secondMy As Double
    math.BuildResiduals 100000#, 1000000#, -1000000#, path, 1#, firstMx, firstMy
    math.BuildResiduals 100000#, 2000000#, -2000000#, path, 1#, secondMx, secondMy
    AssertClose stats, "audit02.pathResidual.zeroMx", firstMx, 10000000000#, 0.001
    AssertClose stats, "audit02.pathResidual.zeroMy", firstMy, -10000000000#, 0.001
    AssertClose stats, "audit02.pathResidual.mxLinear", secondMx, firstMx * 2#, 0.001
    AssertClose stats, "audit02.pathResidual.myLinear", secondMy, firstMy * 2#, 0.001
End Sub

' ДЛЯ ТЕСТОВ
' Начальная плоскость намеренно лежит в численном продолжении, но конечная
' нагрузка мала. Финальные признаки должны описывать найденный физический
' State, а не исходное приближение и не промежуточные Newton-итерации.
Private Sub TestAudit02ExtendedInitialGuessPhysicalFinal(ByRef stats As TSectionSolverTestStats)
    Dim concreteParameters As CConcreteMaterialParameters
    Set concreteParameters = New CConcreteMaterialParameters
    concreteParameters.Initialize 15.5, 1.08, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    Dim steelParameters As CSteelMaterialParameters
    Set steelParameters = New CSteelMaterialParameters
    steelParameters.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#
    Dim provider As CMaterialModelProvider
    Set provider = New CMaterialModelProvider
    provider.InitializeFromParameters concreteParameters, steelParameters, diagramExtensionEnabled:=True
    Dim spec As CMaterialModelSpec
    Set spec = New CMaterialModelSpec
    spec.Initialize "ULS(I)", "ThreeLine", "Ignore", "TwoLine"
    Dim concrete As CMaterialDiagram
    Set concrete = provider.ConcreteMaterialForEquilibriumFromSpec(spec)
    Dim steel As CMaterialDiagram
    Set steel = provider.SteelMaterialForEquilibriumFromSpec(spec)
    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(BuildMesh(RectangleGeometry(200#, 100#), 20#), Nothing)
    Dim initial As CSectionSolver
    Set initial = New CSectionSolver
    initial.EvaluateStrainPlane section, concrete, steel, -0.01, 0#, 0#
    AssertTrue stats, "audit02.finalFlags.initialEquilibrium", _
        initial.ConfirmEquilibrium(initial.Nint, initial.Mxint, initial.Myint)
    Dim runner As CStateSolutionRunner
    Set runner = New CStateSolutionRunner
    AssertTrue stats, "audit02.finalFlags.initialExtended", runner.StateUsesExtension(section, initial, concrete, steel)
    runner.SolveWithInitialSolver section, concrete, steel, -50000#, 0#, 0#, initial, True
    AssertTrue stats, "audit02.finalFlags.converged", runner.Converged
    AssertTrue stats, "audit02.finalFlags.physical", runner.WithinPhysicalRange
    AssertTrue stats, "audit02.finalFlags.notExtended", Not runner.ExtensionUsed
    AssertClose stats, "audit02.finalFlags.force", runner.ResultSolver.Nint, -50000#, 5#
End Sub









' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Проверяет память только идентичных численных retries. Иные старты/options,
' новое сечение/материалы, новая session и успех не блокируются прежней неудачей.
Public Function RunAudit03RetryAttemptSessionTests() As String
    On Error GoTo Failed
    Dim stats As TSectionSolverTestStats
    TestAudit03RetryAttemptSession stats
    AppendLine stats, "TOTAL_RETRY_SESSION: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03RetryAttemptSessionTests = stats.Report
    Exit Function
Failed:
    RunAudit03RetryAttemptSessionTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' Одной Newton-итерации недостаточно для финального подтверждения линейной
' задачи: это управляемая численная неудача без изменения эталонной физики.
Private Sub TestAudit03RetryAttemptSession(ByRef stats As TSectionSolverTestStats)
    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(BuildMesh(RectangleGeometry(200#, 100#), 20#), Nothing)
    Dim concrete As CLinearConcreteMaterial, steel As CLinearSteelMaterial
    Set concrete = New CLinearConcreteMaterial
    concrete.Initialize 32500#
    Set steel = New CLinearSteelMaterial
    steel.Initialize 200000#
    Dim runner As CStateSolutionRunner
    Set runner = New CStateSolutionRunner
    runner.MaxIterations = 1
    runner.DiagnosticsEnabled = False
    runner.BeginRetryAttemptSession section, concrete, steel
    AssertTrue stats, "audit03.retry.firstFailure", Not runner.TestRunRetryAttempt(section, concrete, steel, -100000#)
    AssertTrue stats, "audit03.retry.firstSolve", runner.SolverCallCount = 1
    AssertTrue stats, "audit03.retry.sameFailure", Not runner.TestRunRetryAttempt(section, concrete, steel, -100000#)
    AssertTrue stats, "audit03.retry.sameNoSolve", runner.SolverCallCount = 1
    AssertTrue stats, "audit03.retry.noDisabledDiagnostic", Len(runner.DiagnosticLog) = 0
    runner.TestRunRetryAttempt section, concrete, steel, -100000#, -0.000001
    AssertTrue stats, "audit03.retry.newPlane", runner.SolverCallCount = 2
    runner.TestRunRetryAttempt section, concrete, steel, -110000#
    AssertTrue stats, "audit03.retry.newTarget", runner.SolverCallCount = 3
    runner.ToleranceN = 0.1
    runner.TestRunRetryAttempt section, concrete, steel, -100000#
    AssertTrue stats, "audit03.retry.newTolerance", runner.SolverCallCount = 4
    runner.TestRunRetryAttempt section, concrete, steel, -100000#, 0#, 2
    AssertTrue stats, "audit03.retry.newSteps", runner.SolverCallCount = 5
    runner.TestRunRetryAttempt section, concrete, steel, -100000#, 0#, 1, True
    AssertTrue stats, "audit03.retry.newLineSearch", runner.SolverCallCount = 6
    runner.BeginRetryAttemptSession section, concrete, steel
    runner.TestRunRetryAttempt section, concrete, steel, -100000#
    AssertTrue stats, "audit03.retry.newSession", runner.SolverCallCount = 7
    runner.MaxIterations = 2
    AssertTrue stats, "audit03.retry.canRecover", runner.TestRunRetryAttempt(section, concrete, steel, -100000#)
    AssertTrue stats, "audit03.retry.recoverySolve", runner.SolverCallCount = 8
    AssertTrue stats, "audit03.retry.successNotFailureCache", runner.TestRunRetryAttempt(section, concrete, steel, -100000#)
    AssertTrue stats, "audit03.retry.successSolvedAgain", runner.SolverCallCount = 9
    runner.MaxIterations = 0
    runner.TestRunRetryAttempt section, concrete, steel, -100000#
    runner.TestRunRetryAttempt section, concrete, steel, -100000#
    AssertTrue stats, "audit03.retry.configurationNotCached", runner.SolverCallCount = 11
    runner.MaxIterations = 1
    Dim otherSection As CSectionModel
    Set otherSection = BuildGeneratedSectionModel(BuildMesh(RectangleGeometry(200#, 100#), 20#), Nothing)
    runner.TestRunRetryAttempt otherSection, concrete, steel, -100000#
    runner.TestRunRetryAttempt otherSection, concrete, steel, -100000#
    AssertTrue stats, "audit03.retry.otherContextClosesSession", runner.SolverCallCount = 13
    runner.BeginRetryAttemptSession section, concrete, steel
    runner.TestRunRetryAttempt section, concrete, steel, -100000#
    Dim otherConcrete As CLinearConcreteMaterial
    Set otherConcrete = New CLinearConcreteMaterial
    otherConcrete.Initialize 30000#
    runner.TestRunRetryAttempt section, otherConcrete, steel, -100000#
    runner.TestRunRetryAttempt section, otherConcrete, steel, -100000#
    AssertTrue stats, "audit03.retry.otherMaterialClosesSession", runner.SolverCallCount = 16
    runner.BeginRetryAttemptSession section, concrete, steel
    runner.TestRunRetryAttempt section, concrete, steel, -100000#
    section.Clear
    section.AddConcreteElement 0#, 0#, 20000#
    runner.TestRunRetryAttempt section, concrete, steel, -100000#
    runner.TestRunRetryAttempt section, concrete, steel, -100000#
    AssertTrue stats, "audit03.retry.revisionClosesSession", runner.SolverCallCount = 19
    runner.BeginRetryAttemptSession section, concrete, steel
    Dim targetIndex As Long
    For targetIndex = 1 To 80
        runner.TestRunRetryAttempt section, concrete, steel, -100000# - targetIndex
    Next targetIndex
    AssertTrue stats, "audit03.retry.boundedScalarMemory", runner.TestRetryAttemptCount = 64
    Dim calls As Long
    calls = runner.SolverCallCount
    runner.TestRunRetryAttempt section, concrete, steel, -100080#
    AssertTrue stats, "audit03.retry.lastAttemptRetained", runner.SolverCallCount = calls
    runner.TestRunRetryAttempt section, concrete, steel, -100001#
    AssertTrue stats, "audit03.retry.evictedAttemptCanRun", runner.SolverCallCount = calls + 1
End Sub
