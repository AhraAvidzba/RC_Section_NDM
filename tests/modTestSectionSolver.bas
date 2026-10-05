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
    TestAudit03ExtremeLinearSystem stats
    TestAudit03LinearFailureCodes stats
    TestAudit03LargeTangentState stats
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
    TestAudit03SmallLoadPathComponents stats
    TestAudit02ExtendedInitialGuessPhysicalFinal stats
    TestAudit03TypedStateFailures stats
    TestAudit03RetryAttemptSession stats
    TestAudit03ExtremeStateInputs stats
    TestAudit03ExtremeRunnerStates stats
    TestAudit03SolverSettingEffects stats
    TestAudit03InputCurvatureBinding stats

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
        "RectSet.H1.d_2row_1", "RectSet.H1.d_2row_2", "RectSet.H1.d_3row_1", "RectSet.H1.d_3row_2", "RectSet.H1.loc_2row_1", "RectSet.H1.loc_3row_1", _
        "RectSet.H2.as_1", "RectSet.H2.as_2", "RectSet.H2.d_1", "RectSet.H2.d_2", _
        "RectSet.H2.n_1", "RectSet.H2.n_2", _
        "RectSet.H2.StartOffset1", "RectSet.H2.EndOffset1", _
        "RectSet.H2.StartOffset2", "RectSet.H2.EndOffset2", _
        "RectSet.H2.d_2row_1", "RectSet.H2.d_2row_2", "RectSet.H2.d_3row_1", "RectSet.H2.d_3row_2", "RectSet.H2.loc_2row_1", "RectSet.H2.loc_3row_1", _
        "RectSet.B1.as_1", "RectSet.B1.as_2", "RectSet.B1.d_1", "RectSet.B1.d_2", _
        "RectSet.B1.n_1", "RectSet.B1.n_2", _
        "RectSet.B1.StartOffset1", "RectSet.B1.EndOffset1", _
        "RectSet.B1.StartOffset2", "RectSet.B1.EndOffset2", _
        "RectSet.B1.d_2row_1", "RectSet.B1.d_2row_2", "RectSet.B1.d_3row_1", "RectSet.B1.d_3row_2", "RectSet.B1.loc_2row_1", "RectSet.B1.loc_3row_1", _
        "RectSet.B2.as_1", "RectSet.B2.as_2", "RectSet.B2.d_1", "RectSet.B2.d_2", _
        "RectSet.B2.n_1", "RectSet.B2.n_2", _
        "RectSet.B2.StartOffset1", "RectSet.B2.EndOffset1", _
        "RectSet.B2.StartOffset2", "RectSet.B2.EndOffset2", _
        "RectSet.B2.d_2row_1", "RectSet.B2.d_2row_2", "RectSet.B2.d_3row_1", "RectSet.B2.d_3row_2", "RectSet.B2.loc_2row_1", "RectSet.B2.loc_3row_1")
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
        "SLS.Crack.Allowable", "SLS.Crack.InitiationSolutionStrategy", _
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
    AssertTrue stats, "section.method.invalidMessage", InStr(1, solver.StopReason, "Ошибка настройки Solver.Method:", vbTextCompare) > 0
    AssertTrue stats, "section.method.invalidCode", solver.FailureCode = sfcInvalidConfiguration

    Set solver = New CSectionSolver
    ConfigureProvisionalSolver solver
    solver.SolverMethod = vbNullString
    solver.Solve BuildGeneratedSectionModel(mesh, Nothing), ProvisionalConcrete(), ProvisionalSteel(), -100000#, 0#, 0#
    AssertTrue stats, "section.method.emptyInput", Not solver.Converged
    AssertTrue stats, "section.method.emptyMessage", InStr(1, solver.StopReason, "Ошибка настройки Solver.Method:", vbTextCompare) > 0
    AssertTrue stats, "section.method.emptyCode", solver.FailureCode = sfcInvalidConfiguration
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

' Задает жесткий допуск и короткий бюджет для воспроизведения отказов сходимости.
' Ограничения приращений отключены явно, чтобы тест проверял сам решатель.
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

' Настраивает контрольные задачи равновесия независимо от текущего Config.
' Допуски заданы во внутренних единицах, шаг нагрузки и ограничения фиксированы.
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

' Создает фиксированную диаграмму бетона для численных контрольных задач.
' Она не использует материал из книги и не включает расширение диаграммы.
Private Function ProvisionalConcrete() As CMaterialDiagram
    Dim concrete As CMaterialDiagram
    Set concrete = New CMaterialDiagram
    concrete.Initialize -0.0015, -15.5, -0.0035, -15.5
    Set ProvisionalConcrete = concrete
End Function

' Создает фиксированную симметричную диаграмму арматуры для тех же задач.
' Пределы и модуль не зависят от пользовательских настроек материала.
Private Function ProvisionalSteel() As CMaterialDiagram
    Dim steel As CMaterialDiagram
    Set steel = New CMaterialDiagram
    steel.Initialize 0.00175, 350#, 0.025
    Set ProvisionalSteel = steel
End Function

' Возвращает прямоугольник без скруглений и смещения для контрольного сечения.
' Сетка и арматура затем создаются отдельно, как в обычном расчетном маршруте.
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

' Проверяет все три компоненты равновесия, а не только признак Converged.
' Для нулевой компоненты действует абсолютный допуск, для ненулевой относительный.
Private Sub AssertEquilibrium(ByRef stats As TSectionSolverTestStats, ByVal prefix As String, _
        ByVal solver As CSectionSolver, ByVal n As Double, ByVal mx As Double, ByVal my As Double)
    AssertLoadComponent stats, prefix & ".N", solver.Nint, n, 1#, 0.000001
    AssertLoadComponent stats, prefix & ".Mx", solver.Mxint, mx, 1000#, 0.000001
    AssertLoadComponent stats, prefix & ".My", solver.Myint, my, 1000#, 0.000001
End Sub

' Проверяет наличие диапазона перед тестом, которому требуется схема книги.
' Отсутствие имени означает недоступность fixture, а не отказ численного решателя.
Private Function NamedRangeExists(ByVal rangeName As String) As Boolean
    On Error GoTo Missing
    Dim target As Object
    Set target = ThisWorkbook.Names.Item(rangeName).RefersToRange
    NamedRangeExists = Not target Is Nothing
    Exit Function
Missing:
    NamedRangeExists = False
End Function

' Выбирает допустимую меру ошибки для одной компоненты нагрузки.
' Около нуля не делит на ожидаемое значение и использует его абсолютный допуск.
Private Sub AssertLoadComponent(ByRef stats As TSectionSolverTestStats, ByVal name As String, _
        ByVal actual As Double, ByVal expected As Double, ByVal zeroTolerance As Double, _
        ByVal relTolerance As Double)
    If Abs(expected) <= zeroTolerance Then
        AssertClose stats, name, actual, expected, zeroTolerance
    Else
        AssertRelative stats, name, actual, expected, relTolerance
    End If
End Sub

' Группа проверок ведет общий счет и записывает фактические отклонения в отчет.
' Численные сравнения ниже сохраняют отдельные абсолютные и относительные допуски.
Private Sub AssertTrue(ByRef stats As TSectionSolverTestStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        AppendLine stats, "OK: " & name
    Else
        stats.Failed = stats.Failed + 1
        AppendLine stats, "FAIL: " & name
    End If
End Sub

' Сравнивает значения по абсолютному допуску и сохраняет обе величины и разность.
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

' Сравнивает относительную ошибку; для почти нулевого эталона избегает деления.
' Используемый допуск передается тестом и не подменяется настройками Config.
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

' Добавляет строку в протокол; формат чисел ниже одинаков при любой локали Excel.
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

' ДЛЯ ТЕСТОВ: отдельный запуск связывает выбор INPUT-кривизны с реально
' активным ограничителем шага Newton/Secant. Исходные ячейки восстанавливаются.
Public Function RunAudit03InputCurvatureBindingTests() As String
    Dim stats As TSectionSolverTestStats
    TestAudit03InputCurvatureBinding stats
    AppendLine stats, "TOTAL_AUDIT03_INPUT_CURVATURE_BINDING: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03InputCurvatureBindingTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ: одинаковый физический clamp задается как 1e-7 1/мм либо
' 1e-4 1/м. Сравниваем не только найденное равновесие, но и число итераций:
' без пересчета единиц второй clamp перестает ограничивать шаг. Линейный
' прямоугольник дает независимые аналитические значения обеих кривизн.
Private Sub TestAudit03InputCurvatureBinding(ByRef stats As TSectionSolverTestStats)
    Dim unitRange As Range, savedUnits As Variant, sheet As Worksheet
    Dim oldAlerts As Boolean, row As Long, column As Long, curvatureRow As Long
    Dim section As CSectionModel, concrete As CLinearConcreteMaterial, steel As CLinearSteelMaterial
    Dim settings As CSystemSettingsReader, units As CUnitSystem
    Dim freeSolver As CSectionSolver, referenceSolver As CSectionSolver, solver As CSectionSolver
    Dim disconnectedSolver As CSectionSolver, method As Variant, axis As Long, unitIndex As Long
    Dim mx As Double, my As Double, expectedKx As Double, expectedKy As Double, rawClamp As Double
    Dim prefix As String, unitName As String
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    Set unitRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    savedUnits = unitRange.Formula
    For row = 2 To unitRange.Rows.Count
        Select Case CStr(unitRange.Cells(row, 1).Value2)
            Case "Force": unitRange.Cells(row, 2).Value2 = "N"
            Case "Moment": unitRange.Cells(row, 2).Value2 = "N*mm"
            Case "Curvature": curvatureRow = row
        End Select
    Next row
    If curvatureRow = 0 Then Err.Raise vbObjectError + 4601, "TestAudit03InputCurvatureBinding", "Не найдена INPUT-единица кривизны."
    Set sheet = ThisWorkbook.Worksheets.Add
    sheet.Name = "__Audit03Curvature"
    Set section = BuildGeneratedSectionModel(BuildMesh(RectangleGeometry(200#, 100#), 10#), Nothing)
    Set concrete = New CLinearConcreteMaterial
    concrete.Initialize 30000#
    Set steel = New CLinearSteelMaterial
    steel.Initialize 200000#

    For Each method In Array("Newton", "Secant")
        For axis = 1 To 2
            mx = 0#: my = 0#: expectedKx = 0#: expectedKy = 0#
            ' Моменты инерции сетки считаются по центрам квадратов 10x10 мм,
            ' как в линейном равновесии: B*H*(H^2-step^2)/12 и аналогично по Y.
            If axis = 1 Then
                mx = 1000000#
                expectedKx = mx / (30000# * 200# * 100# * (100# ^ 2 - 10# ^ 2) / 12#)
            Else
                my = -1000000#
                expectedKy = my / (30000# * 200# * 100# * (200# ^ 2 - 10# ^ 2) / 12#)
            End If
            unitRange.Cells(curvatureRow, 2).Value2 = "1/mm"
            Set settings = New CSystemSettingsReader
            settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem
            units.LoadFromSettings settings
            Set freeSolver = Audit03ConfiguredSolver(sheet, "Solver.MaxDeltaKappa", 0#, CStr(method), False, units)
            Set referenceSolver = Audit03ConfiguredSolver(sheet, "Solver.MaxDeltaKappa", 0.0000001, CStr(method), False, units)
            freeSolver.Solve section, concrete, steel, 0#, mx, my
            referenceSolver.Solve section, concrete, steel, 0#, mx, my
            prefix = "audit03.inputCurvature." & CStr(method) & ".axis" & CStr(axis)
            AssertTrue stats, prefix & ".referenceConverged", freeSolver.Converged And referenceSolver.Converged
            AssertTrue stats, prefix & ".binding", referenceSolver.Iterations > freeSolver.Iterations

            For unitIndex = 0 To 1
                If unitIndex = 0 Then
                    unitName = "1/mm": rawClamp = 0.0000001
                Else
                    unitName = "1/m": rawClamp = 0.0001
                End If
                unitRange.Cells(curvatureRow, 2).Value2 = unitName
                Set settings = New CSystemSettingsReader
                settings.LoadFromWorkbook ThisWorkbook
                Set units = New CUnitSystem
                units.LoadFromSettings settings
                Set solver = Audit03ConfiguredSolver(sheet, "Solver.MaxDeltaKappa", rawClamp, CStr(method), False, units)
                solver.Solve section, concrete, steel, 0#, mx, my
                prefix = "audit03.inputCurvature." & CStr(method) & ".axis" & CStr(axis) & ".unit" & CStr(unitIndex)
                AssertClose stats, prefix & ".physicalClamp", units.InputCurvatureToInternal(rawClamp), 0.0000001, 0.000000000000000001
                AssertTrue stats, prefix & ".converged", solver.Converged
                AssertTrue stats, prefix & ".iterations", solver.Iterations = referenceSolver.Iterations
                AssertTrue stats, prefix & ".moreWork", solver.Iterations > freeSolver.Iterations
                AssertClose stats, prefix & ".N", solver.Nint, 0#, 0.001
                AssertClose stats, prefix & ".Mx", solver.Mxint, mx, 0.001
                AssertClose stats, prefix & ".My", solver.Myint, my, 0.001
                AssertClose stats, prefix & ".epsilon0", solver.Epsilon0, 0#, 0.000000000001
                AssertClose stats, prefix & ".kappaX", solver.KappaX, expectedKx, 0.000000000001
                AssertClose stats, prefix & ".kappaY", solver.KappaY, expectedKy, 0.000000000001
                AppendLine stats, "CURVATURE_BINDING: method=" & CStr(method) & "|axis=" & CStr(axis) & _
                    "|input=" & unitName & "|raw=" & FormatNumberInvariant(rawClamp) & _
                    "|iterations=" & CStr(solver.Iterations) & "|freeIterations=" & CStr(freeSolver.Iterations)
                If unitIndex = 1 Then
                    ' Отдельный контроль чувствительности отключает адаптер только
                    ' в тестовом вызове. Production код и физический итог не меняются.
                    Set disconnectedSolver = Audit03ConfiguredSolver(sheet, "Solver.MaxDeltaKappa", rawClamp, CStr(method))
                    disconnectedSolver.Solve section, concrete, steel, 0#, mx, my
                    AssertTrue stats, prefix & ".disconnectedConverged", disconnectedSolver.Converged
                    AssertTrue stats, prefix & ".detectDisconnectedUnits", disconnectedSolver.Iterations < solver.Iterations
                    AssertClose stats, prefix & ".disconnectedKx", disconnectedSolver.KappaX, expectedKx, 0.000000000001
                    AssertClose stats, prefix & ".disconnectedKy", disconnectedSolver.KappaY, expectedKy, 0.000000000001
                    AppendLine stats, "CURVATURE_DISCONNECTED: method=" & CStr(method) & "|axis=" & CStr(axis) & _
                        "|iterations=" & CStr(disconnectedSolver.Iterations) & "|correctIterations=" & CStr(solver.Iterations)
                End If
            Next unitIndex
        Next axis
    Next method
    GoTo Cleanup
Failed:
    AssertTrue stats, "audit03.inputCurvature.runtime." & CStr(Err.Number) & "." & Err.Description, False
Cleanup:
    On Error Resume Next
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
    If Not unitRange Is Nothing And IsArray(savedUnits) Then
        For row = 1 To unitRange.Rows.Count
            For column = 1 To unitRange.Columns.Count
                AssertTrue stats, "audit03.inputCurvature.restore." & CStr(row) & "." & CStr(column), _
                    CStr(unitRange.Cells(row, column).Formula) = CStr(savedUnits(row, column))
            Next column
        Next row
    End If
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
' Проверяет общую lambda-математику отдельно от solver tolerance: маленькая
' ненулевая Base-компонента остается частью пути. Ошибка отдельного случая
' сохраняется в отчете, чтобы проверить все знаки и последующее восстановление.
Public Function RunAudit03SmallLoadPathTests() As String
    Dim stats As TSectionSolverTestStats
    TestAudit03SmallLoadPathComponents stats
    AppendLine stats, "TOTAL_AUDIT03_SMALL_LOAD_PATH: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03SmallLoadPathTests = stats.Report
End Function

' Проверяет N/Mx/My разных масштабов при одних допусках, затем повторяет
' обычный путь на том же helper-е. Полностью нулевой Base остается ошибкой API.
Private Sub TestAudit03SmallLoadPathComponents(ByRef stats As TSectionSolverTestStats)
    Dim math As CLoadPathMath, magnitude As Variant, signValue As Variant, component As Long
    Set math = New CLoadPathMath
    math.Configure 5#, 5000#, 5000#
    For Each magnitude In Array(1000000#, 0.001, 0.000001, 0.000000001, 0.000000000001)
        For Each signValue In Array(-1#, 1#)
            For component = 1 To 3
                CheckAudit03SmallLoadPath stats, math, component, CDbl(magnitude) * CDbl(signValue), _
                    "audit03.smallPath." & CStr(component) & "." & FormatNumberInvariant(CDbl(magnitude)) & "." & CStr(signValue)
            Next component
        Next signValue
    Next magnitude
    CheckAudit03SmallLoadPath stats, math, 2, 1000000#, "audit03.smallPath.recovery"
    Dim path As CLoadPathVector, errorNumber As Long, value As Double
    Set path = New CLoadPathVector
    path.Initialize 0#, 0#, 0#, 0#, 0#, 0#
    On Error Resume Next
    value = math.Lambda(0#, 0#, 0#, path)
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertTrue stats, "audit03.smallPath.zeroBaseControlled", errorNumber = vbObjectError + 3841
End Sub

' Удвоенные усилия лежат точно на lambda=2; независимые невязки должны быть
' нулевыми, а выбор опоры не может перейти к отсутствующей N-компоненте.
Private Sub CheckAudit03SmallLoadPath(ByRef stats As TSectionSolverTestStats, ByVal math As CLoadPathMath, _
        ByVal component As Long, ByVal baseValue As Double, ByVal prefix As String)
    On Error GoTo Failed
    Dim path As CLoadPathVector, nBase As Double, mxBase As Double, myBase As Double
    Dim lambdaValue As Double, r1 As Double, r2 As Double
    Select Case component
        Case 1: nBase = baseValue
        Case 2: mxBase = baseValue
        Case 3: myBase = baseValue
    End Select
    Set path = New CLoadPathVector
    path.Initialize 0#, nBase, 0#, mxBase, 0#, myBase
    AssertTrue stats, prefix & ".anchor", math.AnchorComponent(path) = component
    lambdaValue = math.Lambda(2# * nBase, 2# * mxBase, 2# * myBase, path)
    AssertClose stats, prefix & ".lambda", lambdaValue, 2#, 0#
    math.BuildResiduals 2# * nBase, 2# * mxBase, 2# * myBase, path, lambdaValue, r1, r2
    AssertClose stats, prefix & ".r1", r1, 0#, 0#
    AssertClose stats, prefix & ".r2", r2, 0#, 0#
    Exit Sub
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: " & prefix & ".runtime; " & CStr(Err.Number) & "; " & Err.Description
End Sub

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

' ============================== ДЛЯ ТЕСТОВ: AUDIT03 EXTREME INPUTS ==============================

' Проверяет конечность и typed-исход прямого НДС при представимых нагрузках
' около верхней границы Double. Это отдельный gate, а не обычный overload 1e6.
Public Function RunAudit03ExtremeStateInputTests() As String
    Dim stats As TSectionSolverTestStats
    TestAudit03ExtremeStateInputs stats
    AppendLine stats, "TOTAL_AUDIT03_EXTREME_STATE: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03ExtremeStateInputTests = stats.Report
End Function

' Для Newton/Secant и Extension Off/On проверяет центрированную и смещенную
' модель. Смещение усиливает риск переполнения при исключении связанных
' компонент N/Mx/My, не меняя допустимость самого представимого входного вектора.
Private Sub TestAudit03ExtremeStateInputs(ByRef stats As TSectionSolverTestStats)
    Dim method As Variant, shifted As Variant, extended As Variant
    For Each method In Array("Newton", "Secant")
        For Each shifted In Array(False, True)
            For Each extended In Array(False, True)
                Audit03ExtremeStateCase stats, CStr(method), CBool(shifted), CBool(extended)
            Next extended
        Next shifted
    Next method
End Sub

' Выполняет реальный solver и строит named-state из его результата. Сверхбольшой
' корректный вектор не обязан сходиться, но численный предел не должен стать
' программным CalcErr или физически допустимым state. Следующий обычный solve
' на том же объекте должен восстанавливаться без утечки предыдущего отказа.
Private Sub Audit03ExtremeStateCase(ByRef stats As TSectionSolverTestStats, _
        ByVal method As String, ByVal shifted As Boolean, ByVal extended As Boolean)
    On Error GoTo Failed
    Dim prefix As String
    prefix = "audit03.extreme." & method & ".shift=" & CStr(shifted) & ".extension=" & CStr(extended)
    Dim geometry As CGeometryRoundedRectangle
    Set geometry = New CGeometryRoundedRectangle
    If shifted Then
        geometry.Initialize 200#, 100#, 0#, 0#, 0#, 0#, 125#, 80#
    Else
        geometry.Initialize 200#, 100#, 0#, 0#, 0#, 0#
    End If
    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(BuildMesh(geometry, 20#), Nothing)
    Dim concreteParameters As CConcreteMaterialParameters, steelParameters As CSteelMaterialParameters
    Set concreteParameters = New CConcreteMaterialParameters
    concreteParameters.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    Set steelParameters = New CSteelMaterialParameters
    steelParameters.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#
    Dim materials As CMaterialModelProvider
    Set materials = New CMaterialModelProvider
    materials.InitializeFromParameters concreteParameters, steelParameters, diagramExtensionEnabled:=extended
    Dim spec As CMaterialModelSpec
    Set spec = New CMaterialModelSpec
    spec.Initialize "ULS(I)", "ThreeLine", "Ignore", "TwoLine"
    Dim concrete As CMaterialDiagram, steel As CMaterialDiagram
    Set concrete = materials.ConcreteMaterialForEquilibriumFromSpec(spec)
    Set steel = materials.SteelMaterialForEquilibriumFromSpec(spec)
    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    solver.SolverMethod = method
    solver.LoadSteps = 1
    solver.MaxIterations = 3
    solver.DiagnosticsEnabled = False
    solver.ToleranceN = 1#: solver.ToleranceMx = 1000#: solver.ToleranceMy = 1000#
    solver.Solve section, concrete, steel, -1E+308, 1E+308, -1E+308
    AssertTrue stats, prefix & ".notConverged", Not solver.Converged
    AssertTrue stats, prefix & ".numericCause", solver.FailureCode = sfcNumericalFailure Or solver.FailureCode = sfcSingularTangent
    AssertTrue stats, prefix & ".reason", Len(Trim$(solver.StopReason)) > 0
    AppendLine stats, "EXTREME_CASE: " & prefix & "|N=-1e308|Mx=1e308|My=-1e308|failureCode=" & _
        CStr(solver.FailureCode) & "|comment=" & solver.StopReason
    Dim state As CSectionStateResult
    Set state = New CSectionStateResult
    state.InitializeFromSolver sstStrengthState, cpStrength, spec, solver, False, False, 1
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    AssertTrue stats, prefix & ".stateNumeric", state.InternalStatus = rsNumericalFailure
    AssertTrue stats, prefix & ".displayNumFail", policy.ExternalStatus(state.ResultMeta) = policy.NumFail
    AssertTrue stats, prefix & ".notPhysical", Not state.WithinPhysicalRange
    solver.MaxIterations = 80
    If shifted Then
        solver.Solve section, concrete, steel, -50000#, -4000000#, -6250000#
    Else
        solver.Solve section, concrete, steel, -50000#, 0#, 0#
    End If
    AssertTrue stats, prefix & ".normalRecovers", solver.Converged And solver.FailureCode = sfcNone
    Exit Sub
Failed:
    AssertTrue stats, prefix & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' ============================== ДЛЯ ТЕСТОВ: AUDIT03 SOLVER SETTINGS ==============================

' Проверяет активное действие численных настроек через настоящий reader и Solve.
' Изменение входа должно менять ограничения/ветвь, а не только прочитанное поле.
Public Function RunAudit03SolverSettingEffectTests() As String
    On Error GoTo Failed
    Dim stats As TSectionSolverTestStats
    TestAudit03SolverSettingEffects stats
    AppendLine stats, "TOTAL_AUDIT03_SOLVER_EFFECTS: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03SolverSettingEffectTests = stats.Report
    Exit Function
Failed:
    RunAudit03SolverSettingEffectTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function

' Изолирует каждый ключ на временном листе, затем восстанавливает окружение
' даже после ошибки. Линейная задача дает независимый oracle ограничений;
' кусочно-линейный материал отдельно активирует дробление шага Newton.
Private Sub TestAudit03SolverSettingEffects(ByRef stats As TSectionSolverTestStats)
    On Error GoTo Failed
    Dim oldAlerts As Boolean
    oldAlerts = Application.DisplayAlerts
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Worksheets.Add
    sheet.Name = "__Audit03SolverEffects"
    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(BuildMesh(RectangleGeometry(200#, 100#), 10#), Nothing)
    Dim concrete As CLinearConcreteMaterial, steel As CLinearSteelMaterial
    Set concrete = New CLinearConcreteMaterial
    concrete.Initialize 30000#
    Set steel = New CLinearSteelMaterial
    steel.Initialize 200000#
    Dim first As CSectionSolver, second As CSectionSolver
    Dim key As Variant, n As Double, mx As Double, my As Double

    Set first = Audit03ConfiguredSolver(sheet, "Solver.LoadSteps", 1)
    Set second = Audit03ConfiguredSolver(sheet, "Solver.LoadSteps", 4)
    first.Solve section, concrete, steel, -100000#, 0#, 0#
    second.Solve section, concrete, steel, -100000#, 0#, 0#
    AssertTrue stats, "audit03.effect.Solver.LoadSteps.one", first.Converged And first.LoadStepsCompleted = 1
    AssertTrue stats, "audit03.effect.Solver.LoadSteps.four", second.Converged And second.LoadStepsCompleted = 4
    AssertTrue stats, "audit03.effect.Solver.LoadSteps.work", second.InternalNewtonCallCount = 4 And first.InternalNewtonCallCount = 1

    Set first = Audit03ConfiguredSolver(sheet, "Solver.MaxIterations", 1)
    Set second = Audit03ConfiguredSolver(sheet, "Solver.MaxIterations", 4)
    first.Solve section, concrete, steel, -100000#, 0#, 0#
    second.Solve section, concrete, steel, -100000#, 0#, 0#
    AssertTrue stats, "audit03.effect.Solver.MaxIterations.one", Not first.Converged And first.Iterations = 1
    AssertTrue stats, "audit03.effect.Solver.MaxIterations.four", second.Converged And second.Iterations <= 4

    For Each key In Array("Solver.ToleranceN", "Solver.ToleranceMx", "Solver.ToleranceMy")
        n = 0#: mx = 0#: my = 0#
        Select Case CStr(key)
            Case "Solver.ToleranceN": n = -100000#
            Case "Solver.ToleranceMx": mx = 1000000#
            Case "Solver.ToleranceMy": my = -1000000#
        End Select
        Set first = Audit03ConfiguredSolver(sheet, CStr(key), 2000000#)
        Set second = Audit03ConfiguredSolver(sheet, CStr(key), 0.001)
        first.Solve section, concrete, steel, n, mx, my
        second.Solve section, concrete, steel, n, mx, my
        AssertTrue stats, "audit03.effect." & CStr(key) & ".loose", first.Converged And first.Iterations = 1
        AssertTrue stats, "audit03.effect." & CStr(key) & ".strict", second.Converged And second.Iterations > first.Iterations
        AssertClose stats, "audit03.effect." & CStr(key) & ".looseZero", first.Epsilon0 + first.KappaX + first.KappaY, 0#, 0#
        AssertClose stats, "audit03.effect." & CStr(key) & ".strictN", second.Nint, n, 0.001
        AssertClose stats, "audit03.effect." & CStr(key) & ".strictMx", second.Mxint, mx, 0.001
        AssertClose stats, "audit03.effect." & CStr(key) & ".strictMy", second.Myint, my, 0.001
    Next key

    For Each key In Array("Solver.MaxDeltaEpsilon0", "Solver.MaxDeltaKappa", "Solver.DampingInitial")
        n = -100000#: mx = 0#: my = 0#
        Dim restricted As Double
        Select Case CStr(key)
            Case "Solver.MaxDeltaEpsilon0": restricted = 0.00001
            Case "Solver.MaxDeltaKappa": restricted = 0.0000001: n = 0#: mx = 1000000#
            Case "Solver.DampingInitial": restricted = 0.5
        End Select
        If CStr(key) = "Solver.DampingInitial" Then
            Set first = Audit03ConfiguredSolver(sheet, CStr(key), 1#)
        Else
            Set first = Audit03ConfiguredSolver(sheet, CStr(key), 0#)
        End If
        Set second = Audit03ConfiguredSolver(sheet, CStr(key), restricted)
        first.Solve section, concrete, steel, n, mx, my
        second.Solve section, concrete, steel, n, mx, my
        AssertTrue stats, "audit03.effect." & CStr(key) & ".converged", first.Converged And second.Converged
        AssertTrue stats, "audit03.effect." & CStr(key) & ".moreWork", second.Iterations > first.Iterations
        AssertClose stats, "audit03.effect." & CStr(key) & ".N", second.Nint, n, 0.001
        AssertClose stats, "audit03.effect." & CStr(key) & ".Mx", second.Mxint, mx, 0.001
    Next key

    Set first = Audit03ConfiguredSolver(sheet, "Solver.Method", "Newton")
    Set second = Audit03ConfiguredSolver(sheet, "Solver.Method", "Secant")
    first.Solve section, concrete, steel, -100000#, 0#, 0#
    second.Solve section, concrete, steel, -100000#, 0#, 0#
    AssertTrue stats, "audit03.effect.Solver.Method.Newton", first.Converged And first.InternalNewtonCallCount = 1
    AssertTrue stats, "audit03.effect.Solver.Method.Secant", second.Converged And second.InternalNewtonCallCount = 0

    Set first = Audit03ConfiguredSolver(sheet, "Solver.SecantMinStepNorm", 0.000000000001, "Secant")
    Set second = Audit03ConfiguredSolver(sheet, "Solver.SecantMinStepNorm", 1#, "Secant")
    first.Solve section, concrete, steel, -100000#, 0#, 0#
    second.Solve section, concrete, steel, -100000#, 0#, 0#
    AssertTrue stats, "audit03.effect.Solver.SecantMinStepNorm.small", first.Converged
    AssertTrue stats, "audit03.effect.Solver.SecantMinStepNorm.large", Not second.Converged And second.MatrixRestartCount = 2

    Set first = Audit03ConfiguredSolver(sheet, "Solver.SecantMaxRestarts", 0, "Secant", True)
    Set second = Audit03ConfiguredSolver(sheet, "Solver.SecantMaxRestarts", 3, "Secant", True)
    first.Solve section, concrete, steel, -100000#, 0#, 0#
    second.Solve section, concrete, steel, -100000#, 0#, 0#
    AssertTrue stats, "audit03.effect.Solver.SecantMaxRestarts.zero", Not first.Converged And first.MatrixRestartCount = 0
    AssertTrue stats, "audit03.effect.Solver.SecantMaxRestarts.three", Not second.Converged And second.MatrixRestartCount = 3

    Set first = Audit03ConfiguredSolver(sheet, "Solver.SecantMinStepNorm", 1#, "Newton", True)
    first.Solve section, concrete, steel, -100000#, 0#, 0#
    AssertTrue stats, "audit03.effect.SecantOptions.inactiveInNewton", first.Converged And first.MatrixRestartCount = 0
    Audit03InvalidSolverSettingEffects stats, sheet, section, concrete, steel
    Audit03LineSearchSettingEffects stats, sheet, section, steel
    Audit03NegativeUnsignedUnitSettings stats, sheet, section, concrete, steel
    GoTo Cleanup
Failed:
    AssertTrue stats, "audit03.effect.runtime." & CStr(Err.Number) & "." & Err.Description, False
Cleanup:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
End Sub

' Формирует полный численный fixture и заменяет ровно один key до чтения Range.
' Значения внутренних допусков фиксированы; пользовательская книга не меняется.
Private Function Audit03ConfiguredSolver(ByVal sheet As Worksheet, ByVal key As String, _
        ByVal value As Variant, Optional ByVal method As String = "Newton", _
        Optional ByVal forceSmallSecantStep As Boolean = False, _
        Optional ByVal units As CUnitSystem = Nothing) As CSectionSolver
    Dim data(1 To 15, 1 To 3) As Variant
    Dim keys As Variant, values As Variant, i As Long
    keys = Array("Solver.Method", "Solver.MaxIterations", "Solver.LoadSteps", _
        "Solver.ToleranceN", "Solver.ToleranceMx", "Solver.ToleranceMy", _
        "Solver.LineSearchEnabled", "Solver.DampingInitial", "Solver.MinLineSearchAlpha", _
        "Solver.MaxDeltaEpsilon0", "Solver.MaxDeltaKappa", "Solver.SecantMaxRestarts", _
        "Solver.SecantMinStepNorm", "General.DiagramExtension")
    values = Array(method, 80, 1, 0.001, 0.001, 0.001, "Yes", 1#, 0.03125, 0#, 0#, 2, 0.000000000001, "No")
    data(1, 1) = "Параметр": data(1, 2) = "Значение": data(1, 3) = "Ед."
    For i = 0 To UBound(keys)
        data(i + 2, 1) = keys(i): data(i + 2, 2) = values(i): data(i + 2, 3) = "-"
        If forceSmallSecantStep And CStr(keys(i)) = "Solver.SecantMinStepNorm" Then data(i + 2, 2) = 1#
        If CStr(keys(i)) = key Then data(i + 2, 2) = value
    Next i
    sheet.Range("A1:C15").Value2 = data
    Dim reader As CSystemSettingsReader
    Set reader = New CSystemSettingsReader
    reader.LoadFromRange sheet.Range("A1:C15")
    Dim solver As CSectionSolver
    Set solver = New CSectionSolver
    solver.ApplySettings reader, units
    Set Audit03ConfiguredSolver = solver
End Function

' ДЛЯ ТЕСТОВ: пользовательский знак силы не относится к знаку допуска.
' Перевод tf/tf*m и кривизны не должен скрывать отрицательный численный ввод.
' Проверяем отказ до итераций и последующий обычный solve с положительным вводом.
Private Sub Audit03NegativeUnsignedUnitSettings(ByRef stats As TSectionSolverTestStats, _
        ByVal sheet As Worksheet, ByVal section As CSectionModel, _
        ByVal concrete As CLinearConcreteMaterial, ByVal steel As CLinearSteelMaterial)
    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.InitializeDefaults
    Dim key As Variant, solver As CSectionSolver
    For Each key In Array("Solver.ToleranceN", "Solver.ToleranceMx", "Solver.ToleranceMy", "Solver.MaxDeltaKappa")
        Set solver = Audit03ConfiguredSolver(sheet, CStr(key), -1#, "Newton", False, units)
        solver.Solve section, concrete, steel, -100000#, 0#, 0#
        AssertTrue stats, "audit03.effect.units." & CStr(key) & ".negativeTyped", _
            Not solver.Converged And solver.FailureCode = sfcInvalidConfiguration
        AssertTrue stats, "audit03.effect.units." & CStr(key) & ".noIterations", solver.Iterations = 0
        AssertTrue stats, "audit03.effect.units." & CStr(key) & ".reason", _
            InStr(1, solver.StopReason, "Solver.", vbBinaryCompare) > 0
        AppendLine stats, "SIGNED_SETTING: key=" & CStr(key) & "|raw=-1|converged=" & CStr(solver.Converged) & _
            "|failureCode=" & CStr(solver.FailureCode) & "|iterations=" & CStr(solver.Iterations) & _
            "|reason=" & solver.StopReason
        Set solver = Audit03ConfiguredSolver(sheet, CStr(key), 1#, "Newton", False, units)
        solver.Solve section, concrete, steel, -100000#, 0#, 0#
        AssertTrue stats, "audit03.effect.units." & CStr(key) & ".positiveRecovery", solver.Converged
    Next key
End Sub

' Проверяет диапазоны на публичном ApplySettings -> Solve маршруте. Ошибочная
' конфигурация не должна начинать итерации или превращаться в NumFail.
Private Sub Audit03InvalidSolverSettingEffects(ByRef stats As TSectionSolverTestStats, _
        ByVal sheet As Worksheet, ByVal section As CSectionModel, _
        ByVal concrete As CLinearConcreteMaterial, ByVal steel As CLinearSteelMaterial)
    Dim keys As Variant, values As Variant, i As Long
    keys = Array("Solver.Method", "Solver.MaxIterations", "Solver.LoadSteps", _
        "Solver.ToleranceN", "Solver.ToleranceMx", "Solver.ToleranceMy", _
        "Solver.DampingInitial", "Solver.MinLineSearchAlpha", "Solver.MaxDeltaEpsilon0", _
        "Solver.MaxDeltaKappa", "Solver.SecantMaxRestarts", "Solver.SecantMinStepNorm")
    values = Array("Unknown", 0, 0, 0#, 0#, 0#, 1.1, 1.1, -0.00001, -0.00001, -1, 0#)
    For i = 0 To UBound(keys)
        Dim solver As CSectionSolver
        Set solver = Audit03ConfiguredSolver(sheet, CStr(keys(i)), values(i))
        solver.Solve section, concrete, steel, -100000#, 0#, 0#
        AssertTrue stats, "audit03.effect.invalid." & CStr(keys(i)), _
            Not solver.Converged And solver.FailureCode = sfcInvalidConfiguration
        AssertTrue stats, "audit03.effect.invalid.noIterations." & CStr(keys(i)), solver.Iterations = 0
        AssertTrue stats, "audit03.effect.invalid.reason." & CStr(keys(i)), Len(solver.StopReason) > 0
    Next i
End Sub

' Старт на пологом участке требует дробления Newton: полный шаг ухудшает
' невязку. Проверяется реальная разница On/Off и нижней границы alpha, а не getter.
Private Sub Audit03LineSearchSettingEffects(ByRef stats As TSectionSolverTestStats, _
        ByVal sheet As Worksheet, ByVal section As CSectionModel, ByVal steel As CLinearSteelMaterial)
    Dim strains(1 To 6) As Double, stresses(1 To 6) As Double
    strains(1) = -0.01: stresses(1) = -100#
    strains(2) = -0.001: stresses(2) = -10#
    strains(3) = 0#: stresses(3) = 0#
    strains(4) = 0.001: stresses(4) = 10#
    strains(5) = 0.002: stresses(5) = 11#
    strains(6) = 0.01: stresses(6) = 19#
    Dim concrete As CMaterialDiagram
    Set concrete = New CMaterialDiagram
    concrete.InitializeFromArrays strains, stresses, 6
    Dim first As CSectionSolver, second As CSectionSolver
    Set first = Audit03ConfiguredSolver(sheet, "Solver.LineSearchEnabled", "No")
    Set second = Audit03ConfiguredSolver(sheet, "Solver.LineSearchEnabled", "Yes")
    first.SetInitialState 0.002, 0#, 0#
    second.SetInitialState 0.002, 0#, 0#
    first.Solve section, concrete, steel, 100000#, 0#, 0#
    second.Solve section, concrete, steel, 100000#, 0#, 0#
    AssertTrue stats, "audit03.effect.Solver.LineSearchEnabled.off", first.Converged
    AssertTrue stats, "audit03.effect.Solver.LineSearchEnabled.on", second.Converged
    AssertTrue stats, "audit03.effect.Solver.LineSearchEnabled.offNoReduction", _
        InStr(1, first.DiagnosticLog, "lineSearchUsed=True", vbBinaryCompare) = 0
    AssertTrue stats, "audit03.effect.Solver.LineSearchEnabled.onReduction", _
        InStr(1, second.DiagnosticLog, "lineSearchUsed=True", vbBinaryCompare) > 0
    AppendLine stats, "SETTING_EFFECT: Solver.LineSearchEnabled|offIterations=" & CStr(first.Iterations) & _
        "|onIterations=" & CStr(second.Iterations) & "|offEvaluations=" & CStr(first.InternalForceEvaluationCount) & _
        "|onEvaluations=" & CStr(second.InternalForceEvaluationCount)
    Set first = Audit03ConfiguredSolver(sheet, "Solver.MinLineSearchAlpha", 1#)
    first.SetInitialState 0.002, 0#, 0#
    first.Solve section, concrete, steel, 100000#, 0#, 0#
    AssertTrue stats, "audit03.effect.Solver.MinLineSearchAlpha.one", Not first.Converged And first.FailureCode = sfcNumericalFailure
    AssertTrue stats, "audit03.effect.Solver.MinLineSearchAlpha.small", second.Converged
End Sub

' ==================== ДЛЯ ТЕСТОВ: КРАЙНИЕ МАСШТАБЫ СИСТЕМЫ 3x3 ====================

' Проверяет конечные большие коэффициенты отдельно от нелинейного НДС.
' Численные expected и допуски штатных инженерных регрессий не меняются.
Public Function RunAudit03ExtremeLinearSystemTests() As String
    Dim stats As TSectionSolverTestStats
    TestLinearSystem3x3 stats
    TestAudit03ExtremeLinearSystem stats
    TestAudit03LinearFailureCodes stats
    TestAudit03LargeTangentState stats
    AppendLine stats, "TOTAL_AUDIT03_EXTREME_LINEAR: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    RunAudit03ExtremeLinearSystemTests = stats.Report
End Function

' Проверяет причины настоящего линейного solve, а не вручную созданную meta.
' Большая сумма RHS проходит масштабированную проверку, но непредставимая
' поправка остается численной ошибкой; recovery очищает прошлый FailureCode.
Private Sub TestAudit03LinearFailureCodes(ByRef stats As TSectionSolverTestStats)
    Dim system As CLinearSystem3x3
    Set system = New CLinearSystem3x3
    AssertTrue stats, "audit03.linear.typed.large", system.Solve(2#, 1#, 0#, 1#, 2#, 0#, 0#, 0#, 1#, 1E+308, 1E+308, 1E+308)
    AssertClose stats, "audit03.linear.typed.large.x1", system.X1 / 1E+308, 1# / 3#, 0.000000000000001
    AssertClose stats, "audit03.linear.typed.large.x2", system.X2 / 1E+308, 1# / 3#, 0.000000000000001
    AssertTrue stats, "audit03.linear.typed.large.residual", system.RelativeResidual >= 0# And system.RelativeResidual <= 1E-8
    AssertTrue stats, "audit03.linear.typed.large.code", system.FailureCode = sfcNone
    AssertTrue stats, "audit03.linear.typed.singular", Not system.Solve(1#, 2#, 3#, 2#, 4#, 6#, 3#, 6#, 9#, 1#, 2#, 3#)
    AssertTrue stats, "audit03.linear.typed.singular.code", system.FailureCode = sfcSingularTangent
    AssertTrue stats, "audit03.linear.typed.overflow", Not system.Solve(0.00000000000000001, 0#, 0#, 0#, 1#, 0#, 0#, 0#, 1#, 1E+308, 1#, 1#)
    AssertTrue stats, "audit03.linear.typed.overflow.code", system.FailureCode = sfcNumericalFailure
    AssertTrue stats, "audit03.linear.typed.overflow.vector", system.X1 = 0# And system.X2 = 0# And system.X3 = 0# And Not system.Solved
    AssertTrue stats, "audit03.linear.typed.recovery", system.Solve(1#, 0#, 0#, 0#, 1#, 0#, 0#, 0#, 1#, 1#, 2#, 3#)
    AssertTrue stats, "audit03.linear.typed.recovery.code", system.FailureCode = sfcNone
End Sub

' Проверяет реальный Newton-потребитель с большой, но конечной жесткостью.
' Модуль синтетического линейного материала не является нормативным вводом;
' нагрузки обычного масштаба, проверяется исходное компонентное равновесие.
Private Sub TestAudit03LargeTangentState(ByRef stats As TSectionSolverTestStats)
    Dim geom As CGeometryRoundedRectangle, mesh As CFiberMeshBuilder, section As CSectionModel
    Dim concrete As CLinearConcreteMaterial, steel As CLinearSteelMaterial, solver As CSectionSolver
    Dim modulus As Variant, code As Long, reason As String, prefix As String
    Set geom = RectangleGeometry(200#, 100#)
    Set mesh = BuildMesh(geom, 10#)
    Set section = BuildGeneratedSectionModel(mesh, Nothing)
    Set concrete = New CLinearConcreteMaterial
    Set steel = New CLinearSteelMaterial: steel.Initialize 200000#
    For Each modulus In Array(1E+110, 1E+200)
        concrete.Initialize CDbl(modulus)
        Set solver = New CSectionSolver
        ConfigureStrictSolver solver
        solver.SolverMethod = "Newton"
        solver.SetInitialState 0#, 0#, 0#
        On Error Resume Next
        Err.Clear: solver.Solve section, concrete, steel, 250000#, 120000000#, -80000000#
        code = Err.Number: reason = Err.Description
        On Error GoTo 0
        prefix = "audit03.linear.largeTangent." & CStr(modulus)
        AssertTrue stats, prefix & ".noRuntimeError", code = 0
        AssertTrue stats, prefix & ".converged", solver.Converged
        AssertTrue stats, prefix & ".code", solver.FailureCode = sfcNone
        If solver.Converged Then AssertEquilibrium stats, prefix, solver, 250000#, 120000000#, -80000000#
        AppendLine stats, "LARGE_TANGENT_STATE: " & prefix & "|iterations=" & CStr(solver.Iterations) & "|" & solver.StopReason
    Next modulus
End Sub

' Диагональные и переставленные системы имеют точно известный вектор.
' Большой determinant/норма RHS не должны мешать решению. Непредставимая
' поправка и арифметическое переполнение возвращают управляемый отказ;
' последующий обычный solve того же объекта не наследует неудачу.
Private Sub TestAudit03ExtremeLinearSystem(ByRef stats As TSectionSolverTestStats)
    Dim system As CLinearSystem3x3, coefficientScale As Variant, direction As Variant
    Dim value As Double, solved As Boolean, code As Long, reason As String, prefix As Variant
    Set system = New CLinearSystem3x3
    For Each coefficientScale In Array(1E+110, 1E+200, 1E+307, 1E+308)
        For Each direction In Array(-1#, 1#)
            value = CDbl(coefficientScale) * CDbl(direction)
            For Each prefix In Array("diagonal", "permuted")
                On Error Resume Next
                Err.Clear
                If prefix = "diagonal" Then
                    solved = system.Solve(value, 0#, 0#, 0#, value, 0#, 0#, 0#, value, value, -value, value)
                Else
                    solved = system.Solve(0#, 0#, value, value, 0#, 0#, 0#, value, 0#, value, value, -value)
                End If
                code = Err.Number: reason = Err.Description
                On Error GoTo 0
                Dim label As String
                label = "audit03.linear." & CStr(prefix) & "." & CStr(coefficientScale) & "." & CStr(direction)
                AssertTrue stats, label & ".noRuntimeError", code = 0
                AssertTrue stats, label & ".solved", solved And system.Solved
                If code = 0 And solved Then
                    AssertClose stats, label & ".x1", system.X1, 1#, 0#
                    AssertClose stats, label & ".x2", system.X2, -1#, 0#
                    AssertClose stats, label & ".x3", system.X3, 1#, 0#
                    AssertClose stats, label & ".residual", system.RelativeResidual, 0#, 0#
                End If
                AppendLine stats, "LINEAR_SCALE: " & label & "|solved=" & CStr(solved) & "|error=" & CStr(code) & "|" & reason
            Next prefix
        Next direction
    Next coefficientScale
    ' Здесь determinant равен 1: воспроизводим отдельно переполнение нормы RHS.
    For Each direction In Array(-1#, 1#)
        value = 1E+308 * CDbl(direction)
        On Error Resume Next
        Err.Clear
        solved = system.Solve(1#, 0#, 0#, 0#, 1#, 0#, 0#, 0#, 1#, value, value, value)
        code = Err.Number: reason = Err.Description
        On Error GoTo 0
        label = "audit03.linear.rhsNorm." & CStr(direction)
        AssertTrue stats, label & ".noRuntimeError", code = 0
        AssertTrue stats, label & ".solved", solved And system.Solved
        If code = 0 And solved Then
            AssertTrue stats, label & ".vector", system.X1 = value And system.X2 = value And system.X3 = value
            AssertClose stats, label & ".residual", system.RelativeResidual, 0#, 0#
        End If
        AppendLine stats, "LINEAR_SCALE: " & label & "|solved=" & CStr(solved) & "|error=" & CStr(code) & "|" & reason
    Next direction
    On Error Resume Next
    Err.Clear
    solved = system.Solve(0.00000000000000001, 0#, 0#, 0#, 1#, 0#, 0#, 0#, 1#, 1E+308, 1#, 1#)
    code = Err.Number: reason = Err.Description
    On Error GoTo 0
    AssertTrue stats, "audit03.linear.unrepresentable.noRuntimeError", code = 0
    AssertTrue stats, "audit03.linear.unrepresentable.failed", Not solved And Not system.Solved
    AssertTrue stats, "audit03.linear.unrepresentable.reason", Len(system.StopReason) > 0
    AppendLine stats, "LINEAR_OVERFLOW: division|error=" & CStr(code) & "|" & reason & "|" & system.StopReason
    On Error Resume Next
    Err.Clear
    solved = system.Solve(1E+308, 1E+308, 0#, -1E+308, 1E+308, 0#, 0#, 0#, 1#, 0#, 1E+308, 1#)
    code = Err.Number: reason = Err.Description
    On Error GoTo 0
    AssertTrue stats, "audit03.linear.elimination.noRuntimeError", code = 0
    AssertTrue stats, "audit03.linear.elimination.failed", Not solved And Not system.Solved
    AssertTrue stats, "audit03.linear.elimination.reason", Len(system.StopReason) > 0
    AppendLine stats, "LINEAR_OVERFLOW: elimination|error=" & CStr(code) & "|" & reason & "|" & system.StopReason
    AssertTrue stats, "audit03.linear.recovery", system.Solve(3#, 2#, -1#, 2#, -2#, 4#, -1#, 0.5, -1#, 1#, -2#, 0#) And system.Solved
    AssertClose stats, "audit03.linear.recovery.x1", system.X1, 1#, 0.000000000001
    AppendLine stats, "LINEAR_CASES: validSystems=18; arithmeticFailures=2; nonlinearCases=0"
End Sub

' Измеряет только общий линейный шаг на одинаковой известной системе.
' Пять повторов позволяют сравнить старый и исправленный helper отдельно
' от Excel IO и создания геометрии; это не итоговый benchmark всего НДС.
Public Function RunAudit03LinearSystemBenchmark() As String
    Dim system As CLinearSystem3x3, repetition As Long, i As Long
    Dim started As Double, elapsed As Double, checksum As Double, report As String
    Set system = New CLinearSystem3x3
    For repetition = 1 To 5
        checksum = 0#: started = Timer
        For i = 1 To 100000
            If Not system.Solve(3#, 2#, -1#, 2#, -2#, 4#, -1#, 0.5, -1#, 1#, -2#, 0#) Then
                RunAudit03LinearSystemBenchmark = "RUNTIME ERROR: linear benchmark solve failed; " & system.StopReason
                Exit Function
            End If
            checksum = checksum + system.X1 + system.X2 + system.X3
        Next i
        elapsed = Timer - started
        If elapsed < 0# Then elapsed = elapsed + 86400#
        report = report & "BENCHMARK_LINEAR: sample=" & CStr(repetition) & "; iterations=100000; elapsedSec=" & _
            FormatNumberInvariant(elapsed) & "; checksum=" & FormatNumberInvariant(checksum) & vbCrLf
        If Abs(checksum + 300000#) > 0.000001 Then
            RunAudit03LinearSystemBenchmark = report & "RUNTIME ERROR: linear benchmark checksum differs."
            Exit Function
        End If
    Next repetition
    RunAudit03LinearSystemBenchmark = report & "TOTAL_AUDIT03_LINEAR_BENCHMARK: passed=5; failed=0"
End Function

' ============================== ДЛЯ ТЕСТОВ: БОЛЬШИЕ МОМЕНТЫ RUNNER ==============================
' ДЛЯ ТЕСТОВ: проходит обычный pipeline прямого НДС при N=0, где начальная
' плоскость раньше требовала вычислить Mx^2 + My^2. Конечные нагрузки вне
' возможностей диаграммы должны дать typed численный отказ, а не overflow VBA.
Public Function RunAudit03ExtremeRunnerTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TSectionSolverTestStats
    TestAudit03ExtremeRunnerStates stats
    AppendLine stats, "TOTAL_AUDIT03_EXTREME_RUNNER: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed)
    passed = stats.Passed: failed = stats.Failed
    RunAudit03ExtremeRunnerTests = stats.Report
End Function

' ДЛЯ ТЕСТОВ: обе схемы равновесия, Off/On и шесть направлений чистого изгиба.
' Ошибка отдельной попытки сохраняется и не останавливает остальные варианты;
' после большого запроса тот же runner обязан решить обычное сжатие.
Private Sub TestAudit03ExtremeRunnerStates(ByRef stats As TSectionSolverTestStats)
    Dim section As CSectionModel, concreteParameters As CConcreteMaterialParameters
    Dim steelParameters As CSteelMaterialParameters, materials As CMaterialModelProvider
    Dim spec As CMaterialModelSpec, concrete As CMaterialDiagram, steel As CMaterialDiagram
    Dim runner As CStateSolutionRunner, solver As CSectionSolver, state As CSectionStateResult
    Dim extended As Variant, method As Variant, magnitude As Variant, direction As Long
    Dim mx As Double, my As Double, errorCode As Long, reason As String, prefix As String, magnitudeName As String
    On Error GoTo Failed
    Set section = BuildGeneratedSectionModel(BuildMesh(RectangleGeometry(200#, 100#), 20#), Nothing)
    Set concreteParameters = New CConcreteMaterialParameters
    concreteParameters.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    Set steelParameters = New CSteelMaterialParameters
    steelParameters.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#
    Set spec = New CMaterialModelSpec
    spec.Initialize "ULS(I)", "ThreeLine", "Ignore", "TwoLine"
    For Each extended In Array(False, True)
        Set materials = New CMaterialModelProvider
        materials.InitializeFromParameters concreteParameters, steelParameters, diagramExtensionEnabled:=CBool(extended)
        Set concrete = materials.ConcreteMaterialForEquilibriumFromSpec(spec)
        Set steel = materials.SteelMaterialForEquilibriumFromSpec(spec)
        For Each method In Array("Newton", "Secant")
            For Each magnitude In Array(1E+160, 1E+308)
                If CDbl(magnitude) = 1E+160 Then magnitudeName = "1e160" Else magnitudeName = "1e308"
                For direction = 0 To 5
                    mx = 0#: my = 0#
                    Select Case direction
                        Case 0: mx = CDbl(magnitude)
                        Case 1: mx = -CDbl(magnitude)
                        Case 2: my = CDbl(magnitude)
                        Case 3: my = -CDbl(magnitude)
                        Case 4: mx = CDbl(magnitude): my = -CDbl(magnitude)
                        Case 5: mx = -CDbl(magnitude): my = CDbl(magnitude)
                    End Select
                    prefix = "audit03.extremeRunner." & CStr(extended) & "." & CStr(method) & "." & _
                        magnitudeName & "." & CStr(direction)
                    Set runner = New CStateSolutionRunner
                    runner.SolverMethod = CStr(method): runner.LoadSteps = 1
                    runner.MaxIterations = 3: runner.DiagnosticsEnabled = False
                    On Error Resume Next
                    Err.Clear
                    runner.Solve section, concrete, steel, 0#, mx, my, CBool(extended)
                    errorCode = Err.Number: reason = Err.Description
                    On Error GoTo Failed
                    AssertTrue stats, prefix & ".noRuntimeError", errorCode = 0
                    Set solver = runner.ResultSolver
                    If errorCode = 0 Then
                        AssertTrue stats, prefix & ".hasAttempt", Not solver Is Nothing
                        If Not solver Is Nothing Then
                            AssertTrue stats, prefix & ".notConverged", Not solver.Converged
                            Set state = New CSectionStateResult
                            state.InitializeFromSolver sstStrengthState, cpStrength, spec, solver, False, False, runner.SolverCallCount
                            AssertTrue stats, prefix & ".numericFailure", state.InternalStatus = rsNumericalFailure
                            AssertTrue stats, prefix & ".noPhysicalState", Not state.WithinPhysicalRange
                            AssertTrue stats, prefix & ".reasonPresent", Len(state.ResultMeta.ResultComment) > 0
                        End If
                    End If
                    AppendLine stats, "EXTREME_RUNNER: " & prefix & "|error=" & CStr(errorCode) & "|" & reason
                    runner.MaxIterations = 80
                    runner.Solve section, concrete, steel, -50000#, 0#, 0#, CBool(extended)
                    AssertTrue stats, prefix & ".recovers", runner.Converged
                Next direction
            Next magnitude
        Next method
    Next extended
    Exit Sub
Failed:
    stats.Failed = stats.Failed + 1
    AppendLine stats, "FAIL: audit03.extremeRunner.runtime; " & CStr(Err.Number) & "; " & Err.Description
End Sub
