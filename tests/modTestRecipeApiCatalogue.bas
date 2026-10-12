Attribute VB_Name = "modTestRecipeApiCatalogue"
Option Explicit

' ДЛЯ ТЕСТОВ: воспроизводимые контексты недостающих примеров API-справочника.
Private mPassed As Long, mFailed As Long, mReport As String

Public Function RunAdvancedGuideExamples() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim definition As CSectionDefinition, builder As CSectionBuilder, model As CSectionModel
    Dim path As CGeometryPath, mesh As TSectionMeshOptions, rules As CRebarRuleExecutor
    Dim useMaxStep As Variant, expected As Long
    mesh.StepX = 100#: mesh.StepY = 100#: mesh.BoundarySubdivisions = 4: mesh.MaterialID = 1
    For Each useMaxStep In Array(False, True)
        Set definition = AdvancedGuideDefinition(CBool(useMaxStep))
        Set path = definition.Path("Side")
        Check "advanced.chain.segments." & CStr(useMaxStep), path.SegmentCount = 5
        Check "advanced.chain.namedNose." & CStr(useMaxStep), path.SourceSegmentID(3) = "NoseArc"
        Set builder = New CSectionBuilder: Set model = builder.Build(definition, mesh): Set rules = builder.Rules
        expected = 26: If useMaxStep Then expected = 24
        Check "advanced.physical.count." & CStr(useMaxStep), model.RebarCount = expected
        Check "advanced.projection.skips." & CStr(useMaxStep), rules.SkippedProjections = 2
        Check "advanced.noOverlap." & CStr(useMaxStep), rules.BarOverlaps = 0
        Near "advanced.verticalExtent." & CStr(useMaxStep), model.Annotations.Value(1), 800#
        expected = 8: If useMaxStep Then expected = 6
        Check "advanced.actualLabel." & CStr(useMaxStep), InStr(model.Annotations.Text(2), CStr(expected)) > 0
        Check "advanced.dependentTrace." & CStr(useMaxStep), InStr(rules.RuleTrace("Extra"), "Created=2") > 0
    Next useMaxStep
    GoTo Finished
Failed:
    Check "advanced.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunAdvancedGuideExamples = mReport & "TOTAL ADVANCED GUIDE EXAMPLES: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: полный контекст коротких фрагментов, без нового production-типа.
Public Function AdvancedGuideDefinition(ByVal useMaxStep As Boolean) As CSectionDefinition
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim a As Double, b As Double, c As Double
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "GuideFragments"
    Set path = facade.RectanglePath("Concrete", -300#, -400#, 1000#, 800#)
    facade.Outer "Outer", "Concrete"
    Set path = facade.CirclePath("Hole", 0#, -250#, 50#)
    facade.Opening "Opening", "Hole", "Outer"
    ' Боковой путь: дуга + наклонная + дуга носа + наклонная + дуга.
    a = 100# / Sqr(2#): b = 300# + 50# / Sqr(2#): c = 50# / Sqr(2#)
    Set path = New CGeometryPath: path.Initialize "Nose"
    path.AddArc "HeelLow", 0#, 0#, 100#, -GEOM_PI / 2#, GEOM_PI / 4#
    path.AddLine "InclinedLow", a, -a, b, -c
    path.AddArc "NoseArc", 300#, 0#, 50#, -GEOM_PI / 4#, GEOM_PI / 2#
    path.AddLine "InclinedHigh", b, c, a, a
    path.AddArc "HeelHigh", 0#, 0#, 100#, GEOM_PI / 4#, GEOM_PI / 4#
    path.Seal: facade.AddPath path
    Set path = facade.Chain("Side", "Nose", "HeelLow", 0#, "HeelHigh", 1#)
    facade.RebarAtStations "SideBars", "Side", Array(50#, 200#, 400#, 600#), 10#
    ' Одна группа, два независимых пути: количество/шаг применяется PerPath.
    Set path = facade.LinePath("RowA", -150#, 250#, 50#, 250#)
    Set path = facade.LinePath("RowB", 250#, 200#, 500#, 200#)
    facade.PathGroup "Independent", Array("RowA", "RowB")
    Set options = New CRebarRuleOptions
    options.DistributionScope = PerPath: options.StartSetback = 20#: options.EndSetback = 30#
    options.AnnotationGroup = "RowsLabel"
    options.PlacementCheck = CheckFullDisk: options.OverlapPolicy = RejectOverlap
    If useMaxStep Then
        facade.RebarByMaxStep "Rows", "Independent", 100#, 10#, options
    Else
        facade.RebarByCount "Rows", "Independent", 4, 10#, options
    End If
    ' Внутренний ряд — наружу от CCW-грани отверстия, без ручных координат баров.
    Set path = facade.Offset("HoleAxis", "Hole", -30#)
    Set options = New CRebarRuleOptions: options.EndMode = ClosedUniform
    options.PlacementCheck = CheckFullDisk: options.OverlapPolicy = RejectOverlap
    facade.RebarByCount "Inner", "HoleAxis", 6, 8#, options
    Set path = facade.LinePath("BasePath", -150#, -350#, 150#, -350#)
    facade.RebarByCount "Base", "BasePath", 4, 10#
    ' Два из четырёх лучей попадают на конечную круговую ось; пропуски явные.
    Set options = New CRebarRuleOptions: options.EndMode = ClosedUniform
    options.DirectionY = 1#: options.SkipMissingProjection = True
    options.PlacementCheck = CheckFullDisk: options.OverlapPolicy = RejectOverlap
    facade.RebarProjected "Projected", "HoleAxis", "Base", 6#, options
    ' Связанные дополнительные позиции по локальной нормали, каждая вторая.
    Set options = New CRebarRuleOptions: options.NormalOffset = 10#: options.EverySecond = True
    options.SourceSuffix = "_Extra": facade.RebarDependent "Extra", "Base", 10#, options
    facade.VerticalExtent "Height", "Concrete", "H", 740#
    facade.RebarRowLabel "RowsLabel", "Rows", 0#, 1#, 50#
    Set AdvancedGuideDefinition = facade.Finish
End Function

Public Function RunApiGeometryReadExamples() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim facade As CSectionDefinitionBuilder, definition As CSectionDefinition, builder As CSectionBuilder
    Dim path As CGeometryPath, diagnostic As CGeometryDiagnostic, query As CSectionGeometryQuery
    Dim region As CConcreteRegion, opening As CConcreteRegion, frame As CConcreteRegion
    Dim options As CRebarRuleOptions, rules As CRebarRuleExecutor, model As CSectionModel, mesh As CFiberMeshBuilder
    Dim meshOptions As TSectionMeshOptions, ids As Variant, probes As Variant, x As Double, y As Double
    Dim x2 As Double, y2 As Double, cx As Double, cy As Double, sweep As Double, loopID As Long, edge As TRegionEdge
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "ApiRead", 2
    Set path = facade.AdjacentRectangles("Concrete", Array(Array(-200#, -200#, 400#, 400#)))
    facade.Outer "Outer", "Concrete"
    Set path = facade.LinePath("Row", -150#, 100#, 150#, 100#)
    Set path = facade.LinePath("Row2", -150#, -100#, 150#, -100#)
    facade.PathGroup "Rows", Array("Row", "Row2")
    Set options = New CRebarRuleOptions
    Check "defaults.remainder", options.RemainderPlacement = RemainderCenter
    Check "defaults.tangent", options.TangentOffset = 0#
    Check "defaults.projectionDirection", options.DirectionX = 0# And options.DirectionY = 0#
    options.RemainderPlacement = RemainderStart: options.TangentOffset = 0#: options.DirectionX = 1#
    options.StartSetback = 20#: options.EndSetback = 30#: options.Comment = "Guide": options.SourceReference = "ProgramInputs.Row"
    facade.RebarByCount "Bars", "Row", 3, 20#, options
    facade.VerticalExtent "Height", "Concrete", "H", 250#
    Set definition = facade.Finish
    Check "definition.type", definition.TypeID = "ApiRead"
    Check "definition.schema", definition.SchemaVersion = 2
    Check "definition.annotationCount", definition.AnnotationCount = 1
    ids = definition.TargetPaths("Rows"): Check "group.independent", UBound(ids) - LBound(ids) + 1 = 2
    Set diagnostic = definition.LastDiagnostic: Check "definition.successDiagnostic", diagnostic Is Nothing
    meshOptions.StepX = 100#: meshOptions.StepY = 100#: meshOptions.BoundarySubdivisions = 4: meshOptions.MaterialID = 1
    Set builder = New CSectionBuilder: Set model = builder.Build(definition, meshOptions)
    Set mesh = builder.Mesh: Check "builder.preparedMesh", Not mesh Is Nothing
    Set rules = builder.Rules
    Check "rules.positionCount", rules.PositionCount = 3
    Check "rules.diagnosticsText", InStr(rules.DiagnosticsText, "Created=3") > 0
    Set query = New CSectionGeometryQuery: Set region = definition.ConcreteDomain
    Set opening = query.CircleRegion(0#, 0#, 20#): Set frame = query.SubtractContainedRegion(region, opening)
    Check "query.loopRole.outer", Not query.LoopIsOpening(frame, 1)
    Check "query.loopRole.opening", query.LoopIsOpening(frame, 2)
    Near "query.subtractedArea", frame.Area, 160000# - 400# * GEOM_PI
    probes = query.SampleMaterialIntervals(region, 1#, 0#, 0#, -10#, 10#, 3)
    Check "query.samples.shape", UBound(probes, 1) = 3 And UBound(probes, 2) = 5
    Check "query.samples.hit", probes(2, 1)
    Near "query.samples.firstX", probes(2, 2), -200#: Near "query.samples.lastX", probes(2, 4), 200#
    region.GetSegment 1, x, y, x2, y2, cx, cy, sweep, loopID
    Check "region.scalarSegment", loopID = 1 And sweep = 0#
    x = 1#: y = 2#: query.TransformCoordinates x, y, 10#, 20#, GEOM_PI / 2#, True
    Near "isometry.reflectThenRotate.x", x, 12#: Near "isometry.reflectThenRotate.y", y, 21#
    Set path = definition.Path("Row")
    Set diagnostic = path.LastDiagnostic: Check "path.successDiagnostic", diagnostic Is Nothing
    On Error Resume Next
    path.CopySegment 0, edge
    Check "path.invalidIndexCode", Err.Number = vbObjectError + 5520: Err.Clear
    On Error GoTo Failed
    Set diagnostic = path.LastDiagnostic: Check "path.reason", Len(diagnostic.Reason) > 0
    On Error Resume Next
    Set path = definition.Path("Missing")
    Check "definition.missingPathCode", Err.Number = vbObjectError + 5530: Err.Clear
    On Error GoTo Failed
    Set diagnostic = definition.LastDiagnostic: Check "definition.errorDiagnostic", diagnostic.Code = "MISSING_PATH"
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "ApiFailure"
    Set path = facade.CirclePath("Small", 0#, 0#, 10#): facade.Source "Small", "ProgramInputs.Radius"
    On Error Resume Next
    Set path = facade.Offset("Lost", "Small", 20#)
    Check "offset.failure", Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    Set diagnostic = facade.LastGeometryDiagnostic
    Check "facade.offsetDiagnostic", Not diagnostic Is Nothing
    Check "facade.offsetReason", Len(diagnostic.Reason) > 0
    GoTo Finished
Failed:
    Check "geometry.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunApiGeometryReadExamples = mReport & "TOTAL API GEOMETRY EXAMPLES: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Public Function RunApiUnitExamples() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim units As CUnitSystem
    Set units = New CUnitSystem: units.InitializeDefaults
    Near "units.inputLength", units.InputLengthToInternal(20#), 20#
    Near "units.outputLength", units.InternalLengthToOutput(2000#), 2#
    Near "units.savedLength", units.OutputLengthToInternal(2#, "cm"), 20#
    Near "units.inputArea", units.InputAreaToInternal(100#), 100#
    Near "units.outputArea", units.InternalAreaToOutput(1000000#), 1#
    Near "units.savedArea", units.OutputAreaToInternal(2#, "cm2"), 200#
    Near "units.inputFourth", units.InputFourthPowerLengthToInternal(100#), 100#
    Near "units.outputFourth", units.InternalFourthPowerLengthToOutput(1000000000000#), 1#
    Near "units.savedFourth", units.OutputFourthPowerLengthToInternal(2#, "cm"), 20000#
    Near "units.force.input", units.InputForceToInternal(123#), 123#
    Near "units.force.output", units.InternalForceToOutput(123#), 123#
    Near "units.stiffness", units.InternalForceLengthSquaredToOutput(1000000#), 1#
    Near "units.moment.inputX", units.InputMomentMxToInternal(12#), 12#
    Near "units.moment.inputY", units.InputMomentMyToInternal(-12#), -12#
    Near "units.momentPerLength", units.InputMomentPerLengthToInternal(12#), 12#
    Near "units.moment.outputX", units.InternalMomentMxToOutput(12#), 12#
    Near "units.moment.outputY", units.InternalMomentMyToOutput(-12#), -12#
    Near "units.moment.magnitude", units.InternalMomentMagnitudeToOutput(12#), 12#
    Near "units.stress.input", units.InputStressToInternal(12#), 12#
    Near "units.stress.output", units.InternalStressToOutput(12#), 12#
    Near "units.stress.saved", units.OutputStressToInternal(12#, "MPa"), 12#
    Near "units.curvature.input", units.InputCurvatureToInternal(0.01), 0.01
    Near "units.curvature.output", units.InternalCurvatureToOutput(0.001, "1/m"), 1#
    Near "units.curvature.inverse", units.OutputCurvatureToInternal(0.001), 0.001
    Check "units.selector.length", units.OutputLengthUnit = "m"
    Check "units.selector.area", units.OutputAreaUnit = "m2"
    Check "units.selector.force", units.OutputForceUnit = "N"
    Check "units.selector.moment", units.OutputMomentUnit = "N*mm"
    Check "units.selector.momentPerLength", units.InputMomentPerLengthUnit = "N*mm/mm"
    Check "units.selector.stiffness", units.OutputForceLengthSquaredUnit = "N*m2"
    Check "units.selector.stress", units.OutputStressUnit = "MPa"
    Check "units.selector.curvature", units.OutputCurvatureUnit = "1/mm"
    Check "units.sign.N", units.OutputSignN = "Tension"
    Check "units.sign.Mx", units.OutputSignMx = "+Y tension"
    Check "units.sign.My", units.OutputSignMy = "+X tension"
    GoTo Finished
Failed:
    Check "units.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunApiUnitExamples = mReport & "TOTAL API UNIT EXAMPLES: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Public Function RunApiInputAndDataExamples() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim table As CSectionInputTable, context As CSectionRecipeContext, units As CUnitSystem, values(1 To 3, 1 To 2) As Variant
    Dim registry As CSectionTypeRegistry, recipe As ISectionRecipe
    Set units = New CUnitSystem: units.InitializeDefaults
    values(1, 1) = "Key": values(1, 2) = "Value": values(2, 1) = "R": values(2, 2) = 100#
    values(3, 1) = "Caption": values(3, 2) = "Guide"
    Set table = New CSectionInputTable: table.Initialize "ProgramInputs", values
    Check "table.name", table.Name = "ProgramInputs"
    Check "table.text", table.Text(3, 2) = "Guide"
    Set context = New CSectionRecipeContext: context.Initialize units: context.AddTable table: context.Seal
    Set units = context.Units: Near "context.normalizedUnit", units.InputLengthToInternal(table.Number(2, 2)), 100#
    Set registry = New CSectionTypeRegistry: Set recipe = registry.CreateRecipe("Circle")
    Check "registry.recipe", Not recipe Is Nothing
    Dim loopRole As ESectionLoopRole, distribution As ERebarDistribution, scope As ERebarDistributionScope, remainder As ERemainderPlacement
    loopRole = LoopOuter: distribution = ByCount: scope = PerPath: remainder = RemainderCenter
    Check "typed.selectors", loopRole = 1 And distribution = 1 And scope = 1 And remainder = 3
    Dim loopValue As TSectionLoop, ruleValue As TRebarRule, annotationValue As TSectionAnnotation, positionValue As TRebarPosition
    loopValue.PathID = "Concrete": loopValue.Role = LoopOuter
    ruleValue.RuleID = "Bars": ruleValue.Mode = ByCount: ruleValue.DistributionScope = PerPath
    annotationValue.AnnotationID = "Height": annotationValue.Kind = "VerticalExtent"
    positionValue.X = 20#: positionValue.Station = 100#: positionValue.PathID = "Row"
    Check "raw.loop", loopValue.PathID = "Concrete" And loopValue.Role = LoopOuter
    Check "raw.rule", ruleValue.RuleID = "Bars" And ruleValue.Mode = ByCount
    Check "raw.annotation", annotationValue.AnnotationID = "Height"
    Check "raw.position", positionValue.X = 20# And positionValue.Station = 100#
    Dim point As TPoint2D, fiber As TFiber, properties As TGeometryProperties
    point.X = 1#: point.Y = 2#: fiber.X = point.X: fiber.Y = point.Y: fiber.Area = 100#: properties.Area = fiber.Area
    Check "raw.point", point.X = 1# And point.Y = 2#
    Check "raw.fiber", fiber.Area = 100#
    Check "raw.properties", properties.Area = 100#
    Check "limit.counter", GEOM_MAX_REBAR_COUNT > 2147483645# And GEOM_MAX_REBAR_COUNT < 2147483647#
    Near "helper.min", GeomMin(1#, 2#), 1#
    Near "saved.sweep.strict", ReadContourArcSweep("1.25", "ProgramInputs.Arc"), 1.25
    On Error Resume Next
    SectionRecipeError "ApiGuide", "Invalid count", "ProgramInputs.Count"
    Check "error.address", Err.Number = vbObjectError + 5530 And InStr(Err.Description, "ProgramInputs.Count") > 0: Err.Clear
    Err.Raise vbObjectError + 5520, "CGeometryPath", "Broken arc"
    SectionRecipeGeometryError "ApiGuide", "ProgramInputs.Arc"
    Check "error.geometryAddress", Err.Number = vbObjectError + 5530 And InStr(Err.Description, "ProgramInputs.Arc") > 0: Err.Clear
    On Error GoTo Failed
    GoTo Finished
Failed:
    Check "input.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunApiInputAndDataExamples = mReport & "TOTAL API INPUT EXAMPLES: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Private Sub Near(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    Check name, Abs(actual - expected) <= 0.000001 * GeomMax(1#, Abs(expected))
End Sub

Private Sub Check(ByVal name As String, ByVal passed As Boolean)
    If passed Then
        mPassed = mPassed + 1: mReport = mReport & "PASS: " & name & vbCrLf
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
    End If
End Sub
