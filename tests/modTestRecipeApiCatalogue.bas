Attribute VB_Name = "modTestRecipeApiCatalogue"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: воспроизводимые контексты недостающих примеров API-справочника.
' ==========================================================================
Private mPassed As Long, mFailed As Long, mReport As String

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет пример API с составной стороной, независимыми рядами, проекцией и дополнительной арматурой при двух способах задания шага.
Public Function TestRunAdvancedGuideExamples() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim definition As CSectionDefinition, builder As CSectionBuilder, model As CSectionModel
    Dim path As CGeometryPath, mesh As TSectionMeshOptions, rules As CRebarRuleExecutor
    Dim useMaxStep As Variant, expected As Long
    mesh.StepX = 100#: mesh.StepY = 100#: mesh.BoundarySubdivisions = 4: mesh.MaterialID = 1
    For Each useMaxStep In Array(False, True)
        Set definition = TestAdvancedGuideDefinition(CBool(useMaxStep))
        Set path = definition.Path("Side")
        TestCheck "advanced.chain.segments." & CStr(useMaxStep), path.SegmentCount = 5
        TestCheck "advanced.chain.namedNose." & CStr(useMaxStep), path.SourceSegmentID(3) = "NoseArc"
        Set builder = New CSectionBuilder: Set model = builder.Build(definition, mesh): Set rules = builder.Rules
        expected = 26: If useMaxStep Then expected = 24
        TestCheck "advanced.physical.count." & CStr(useMaxStep), model.RebarCount = expected
        TestCheck "advanced.projection.skips." & CStr(useMaxStep), rules.SkippedProjections = 2
        TestCheck "advanced.noOverlap." & CStr(useMaxStep), rules.BarOverlaps = 0
        TestNear "advanced.verticalExtent." & CStr(useMaxStep), model.Annotations.Value(1), 800#
        expected = 8: If useMaxStep Then expected = 6
        TestCheck "advanced.actualLabel." & CStr(useMaxStep), InStr(model.Annotations.Text(2), CStr(expected)) > 0
        TestCheck "advanced.dependentTrace." & CStr(useMaxStep), InStr(rules.RuleTrace("Extra"), "Created=2") > 0
    Next useMaxStep
    GoTo Finished
Failed:
    TestCheck "advanced.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    TestRunAdvancedGuideExamples = mReport & "TOTAL ADVANCED GUIDE EXAMPLES: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' задаёт все данные для коротких примеров API без нового рабочего типа сечения.
Public Function TestAdvancedGuideDefinition(ByVal useMaxStep As Boolean) As CSectionDefinition
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim a As Double, b As Double, c As Double
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "GuideFragments"
    Set path = facade.RectanglePath("Concrete", -300#, -400#, 1000#, 800#)
    facade.Outer "Outer", "Concrete"
    Set path = facade.CirclePath("Hole", 0#, -250#, 50#)
    facade.Opening "Opening", "Hole", "Outer"
    ' Боковая линия состоит из дуги, наклонного отрезка, дуги вершины, наклонного отрезка и
    ' дуги.
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
    ' Одна группа объединяет две независимые линии. Количество и шаг задаются для каждой линии
    ' отдельно.
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
    ' Внутренний ряд смещается от границы отверстия в бетон без ручного задания координат
    ' стержней.
    Set path = facade.Offset("HoleAxis", "Hole", -30#)
    Set options = New CRebarRuleOptions: options.EndMode = IncludeStart
    options.PlacementCheck = CheckFullDisk: options.OverlapPolicy = RejectOverlap
    facade.RebarByCount "Inner", "HoleAxis", 6, 8#, options
    Set path = facade.LinePath("BasePath", -150#, -350#, 150#, -350#)
    facade.RebarByCount "Base", "BasePath", 4, 10#
    Set options = New CRebarRuleOptions: options.EndMode = IncludeStart ' Два из четырёх лучей попадают на конечную круговую ось; пропуски явные.
    options.DirectionY = 1#: options.SkipMissingProjection = True
    options.PlacementCheck = CheckFullDisk: options.OverlapPolicy = RejectOverlap
    facade.RebarProjected "Projected", "HoleAxis", "Base", 6#, options
    ' Дополнительные стержни связываются с каждой второй исходной позицией и смещаются по
    ' нормали.
    Set options = New CRebarRuleOptions: options.NormalOffset = 10#: options.EverySecond = True
    options.SourceSuffix = "_Extra": facade.RebarDependent "Extra", "Base", 10#, options
    facade.VerticalExtent "Height", "Concrete", "H", 740#
    facade.RebarRowLabel "RowsLabel", "Rows", 0#, 1#
    Set TestAdvancedGuideDefinition = facade.Finish
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет чтение описания, источников и областей, геометрические преобразования и структурированную диагностику ошибок API.
Public Function TestRunApiGeometryReadExamples() As String
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
    TestCheck "defaults.remainder", options.RemainderPlacement = RemainderCenter
    TestCheck "defaults.tangent", options.TangentOffset = 0#
    TestCheck "defaults.projectionDirection", options.DirectionX = 0# And options.DirectionY = 0#
    options.RemainderPlacement = RemainderStart: options.TangentOffset = 0#: options.DirectionX = 1#
    options.StartSetback = 20#: options.EndSetback = 30#: options.Comment = "Guide": options.SourceReference = "ProgramInputs.Row"
    facade.RebarByCount "Bars", "Row", 3, 20#, options
    facade.VerticalExtent "Height", "Concrete", "H", 250#
    Set definition = facade.Finish
    TestCheck "definition.type", definition.TypeID = "ApiRead"
    TestCheck "definition.schema", definition.SchemaVersion = 2
    TestCheck "definition.annotationCount", definition.AnnotationCount = 1
    ids = definition.TargetPaths("Rows"): TestCheck "group.independent", UBound(ids) - LBound(ids) + 1 = 2
    Set diagnostic = definition.LastDiagnostic: TestCheck "definition.successDiagnostic", diagnostic Is Nothing
    meshOptions.StepX = 100#: meshOptions.StepY = 100#: meshOptions.BoundarySubdivisions = 4: meshOptions.MaterialID = 1
    Set builder = New CSectionBuilder: Set model = builder.Build(definition, meshOptions)
    Set mesh = builder.Mesh: TestCheck "builder.preparedMesh", Not mesh Is Nothing
    Set rules = builder.Rules
    TestCheck "rules.positionCount", rules.PositionCount = 3
    TestCheck "rules.diagnosticsText", InStr(rules.DiagnosticsText, "Created=3") > 0
    Set query = New CSectionGeometryQuery: Set region = definition.ConcreteDomain
    Set opening = query.CircleRegion(0#, 0#, 20#): Set frame = query.SubtractContainedRegion(region, opening)
    TestCheck "query.loopRole.outer", Not query.LoopIsOpening(frame, 1)
    TestCheck "query.loopRole.opening", query.LoopIsOpening(frame, 2)
    TestNear "query.subtractedArea", frame.Area, 160000# - 400# * GEOM_PI
    probes = query.SampleMaterialIntervals(region, 1#, 0#, 0#, -10#, 10#, 3)
    TestCheck "query.samples.shape", UBound(probes, 1) = 3 And UBound(probes, 2) = 5
    TestCheck "query.samples.hit", probes(2, 1)
    TestNear "query.samples.firstX", probes(2, 2), -200#: TestNear "query.samples.lastX", probes(2, 4), 200#
    region.GetSegment 1, x, y, x2, y2, cx, cy, sweep, loopID
    TestCheck "region.scalarSegment", loopID = 1 And sweep = 0#
    x = 1#: y = 2#: query.TransformCoordinates x, y, 10#, 20#, GEOM_PI / 2#, True
    TestNear "isometry.reflectThenRotate.x", x, 12#: TestNear "isometry.reflectThenRotate.y", y, 21#
    Set path = definition.Path("Row")
    Set diagnostic = path.LastDiagnostic: TestCheck "path.successDiagnostic", diagnostic Is Nothing
    On Error Resume Next
    path.CopySegment 0, edge
    TestCheck "path.invalidIndexCode", Err.Number = vbObjectError + 5520: Err.Clear
    On Error GoTo Failed
    Set diagnostic = path.LastDiagnostic: TestCheck "path.reason", Len(diagnostic.Reason) > 0
    On Error Resume Next
    Set path = definition.Path("Missing")
    TestCheck "definition.missingPathCode", Err.Number = vbObjectError + 5530: Err.Clear
    On Error GoTo Failed
    Set diagnostic = definition.LastDiagnostic: TestCheck "definition.errorDiagnostic", diagnostic.Code = "MISSING_PATH"
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "ApiFailure"
    Set path = facade.CirclePath("Small", 0#, 0#, 10#): facade.Source "Small", "ProgramInputs.Radius"
    On Error Resume Next
    Set path = facade.Offset("Lost", "Small", 20#)
    TestCheck "offset.failure", Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    Set diagnostic = facade.LastGeometryDiagnostic
    TestCheck "facade.offsetDiagnostic", Not diagnostic Is Nothing
    If Not diagnostic Is Nothing Then
        TestCheck "facade.offsetReason", Len(diagnostic.Reason) > 0
        TestCheck "facade.offsetCode", diagnostic.Code = "ROW_OFFSET_COLLAPSED"
        TestCheck "facade.offsetEntity", diagnostic.EntityID = "Small"
        TestCheck "facade.offsetSource", diagnostic.SourceReference = "ProgramInputs.Radius"
        TestNear "facade.offsetDistance", diagnostic.Detail1, 20#
        TestNear "facade.offsetRadius", diagnostic.Detail2, 10#
    End If
    GoTo Finished
Failed:
    TestCheck "geometry.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    TestRunApiGeometryReadExamples = mReport & "TOTAL API GEOMETRY EXAMPLES: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет преобразования длины, площади, инерции, усилий, напряжений и кривизны и явные обозначения единиц и знаков.
Public Function TestRunApiUnitExamples() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim units As CUnitSystem
    Set units = New CUnitSystem: units.InitializeDefaults
    TestNear "units.inputLength", units.InputLengthToInternal(20#), 20#
    TestNear "units.outputLength", units.InternalLengthToOutput(2000#), 2#
    TestNear "units.savedLength", units.OutputLengthToInternal(2#, "cm"), 20#
    TestNear "units.inputArea", units.InputAreaToInternal(100#), 100#
    TestNear "units.outputArea", units.InternalAreaToOutput(1000000#), 1#
    TestNear "units.savedArea", units.OutputAreaToInternal(2#, "cm2"), 200#
    TestNear "units.inputFourth", units.InputFourthPowerLengthToInternal(100#), 100#
    TestNear "units.outputFourth", units.InternalFourthPowerLengthToOutput(1000000000000#), 1#
    TestNear "units.savedFourth", units.OutputFourthPowerLengthToInternal(2#, "cm"), 20000#
    TestNear "units.force.input", units.InputForceToInternal(123#), 123#
    TestNear "units.force.output", units.InternalForceToOutput(123#), 123#
    TestNear "units.stiffness", units.InternalForceLengthSquaredToOutput(1000000#), 1#
    TestNear "units.moment.inputX", units.InputMomentMxToInternal(12#), 12#
    TestNear "units.moment.inputY", units.InputMomentMyToInternal(-12#), -12#
    TestNear "units.momentPerLength", units.InputMomentPerLengthToInternal(12#), 12#
    TestNear "units.moment.outputX", units.InternalMomentMxToOutput(12#), 12#
    TestNear "units.moment.outputY", units.InternalMomentMyToOutput(-12#), -12#
    TestNear "units.moment.magnitude", units.InternalMomentMagnitudeToOutput(12#), 12#
    TestNear "units.stress.input", units.InputStressToInternal(12#), 12#
    TestNear "units.stress.output", units.InternalStressToOutput(12#), 12#
    TestNear "units.stress.saved", units.OutputStressToInternal(12#, "MPa"), 12#
    TestNear "units.curvature.input", units.InputCurvatureToInternal(0.01), 0.01
    TestNear "units.curvature.output", units.InternalCurvatureToOutput(0.001, "1/m"), 1#
    TestNear "units.curvature.inverse", units.OutputCurvatureToInternal(0.001), 0.001
    TestCheck "units.selector.length", units.OutputLengthUnit = "m"
    TestCheck "units.selector.area", units.OutputAreaUnit = "m2"
    TestCheck "units.selector.force", units.OutputForceUnit = "N"
    TestCheck "units.selector.moment", units.OutputMomentUnit = "N*mm"
    TestCheck "units.selector.momentPerLength", units.InputMomentPerLengthUnit = "N*mm/mm"
    TestCheck "units.selector.stiffness", units.OutputForceLengthSquaredUnit = "N*m2"
    TestCheck "units.selector.stress", units.OutputStressUnit = "MPa"
    TestCheck "units.selector.curvature", units.OutputCurvatureUnit = "1/mm"
    TestCheck "units.sign.N", units.OutputSignN = "Tension"
    TestCheck "units.sign.Mx", units.OutputSignMx = "+Y tension"
    TestCheck "units.sign.My", units.OutputSignMy = "+X tension"
    GoTo Finished
Failed:
    TestCheck "units.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    TestRunApiUnitExamples = mReport & "TOTAL API UNIT EXAMPLES: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет примеры таблиц, контекста рецепта, реестра, типизированных структур и адресных ошибок API.
Public Function TestRunApiInputAndDataExamples() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim table As CSectionInputTable, context As CSectionRecipeContext, units As CUnitSystem, values(1 To 3, 1 To 2) As Variant
    Dim registry As CSectionTypeRegistry, recipe As ISectionRecipe
    Set units = New CUnitSystem: units.InitializeDefaults
    values(1, 1) = "Key": values(1, 2) = "Value": values(2, 1) = "R": values(2, 2) = 100#
    values(3, 1) = "Caption": values(3, 2) = "Guide"
    Set table = New CSectionInputTable: table.Initialize "ProgramInputs", values
    TestCheck "table.name", table.Name = "ProgramInputs"
    TestCheck "table.text", table.Text(3, 2) = "Guide"
    Set context = New CSectionRecipeContext: context.Initialize units: context.AddTable table: context.Seal
    Set units = context.Units: TestNear "context.normalizedUnit", units.InputLengthToInternal(table.Number(2, 2)), 100#
    Set registry = New CSectionTypeRegistry: Set recipe = registry.CreateRecipe("Circle")
    TestCheck "registry.recipe", Not recipe Is Nothing
    Dim loopRole As ESectionLoopRole, distribution As ERebarDistribution, scope As ERebarDistributionScope, remainder As ERemainderPlacement
    loopRole = LoopOuter: distribution = ByCount: scope = PerPath: remainder = RemainderCenter
    TestCheck "typed.selectors", loopRole = 1 And distribution = 1 And scope = 1 And remainder = 3
    Dim loopValue As TSectionLoop, ruleValue As TRebarRule, annotationValue As TSectionAnnotation, positionValue As TRebarPosition
    loopValue.PathID = "Concrete": loopValue.Role = LoopOuter
    ruleValue.RuleID = "Bars": ruleValue.Mode = ByCount: ruleValue.DistributionScope = PerPath
    annotationValue.AnnotationID = "Height": annotationValue.Kind = "VerticalExtent"
    positionValue.X = 20#: positionValue.Station = 100#: positionValue.PathID = "Row"
    TestCheck "raw.loop", loopValue.PathID = "Concrete" And loopValue.Role = LoopOuter
    TestCheck "raw.rule", ruleValue.RuleID = "Bars" And ruleValue.Mode = ByCount
    TestCheck "raw.annotation", annotationValue.AnnotationID = "Height"
    TestCheck "raw.position", positionValue.X = 20# And positionValue.Station = 100#
    Dim point As TPoint2D, fiber As TFiber, properties As TGeometryProperties
    point.X = 1#: point.Y = 2#: fiber.X = point.X: fiber.Y = point.Y: fiber.Area = 100#: properties.Area = fiber.Area
    TestCheck "raw.point", point.X = 1# And point.Y = 2#
    TestCheck "raw.fiber", fiber.Area = 100#
    TestCheck "raw.properties", properties.Area = 100#
    TestCheck "limit.counter", GEOM_MAX_REBAR_COUNT > 2147483645# And GEOM_MAX_REBAR_COUNT < 2147483647#
    TestNear "helper.min", GeomMin(1#, 2#), 1#
    TestNear "saved.sweep.strict", ReadContourArcSweep("1.25", "ProgramInputs.Arc"), 1.25
    On Error Resume Next
    SectionRecipeError "ApiGuide", "Invalid count", "ProgramInputs.Count"
    TestCheck "error.address", Err.Number = vbObjectError + 5530 And InStr(Err.Description, "ProgramInputs.Count") > 0: Err.Clear
    Err.Raise vbObjectError + 5520, "CGeometryPath", "Broken arc"
    SectionRecipeGeometryError "ApiGuide", "ProgramInputs.Arc"
    TestCheck "error.geometryAddress", Err.Number = vbObjectError + 5530 And InStr(Err.Description, "ProgramInputs.Arc") > 0: Err.Clear
    On Error GoTo Failed
    GoTo Finished
Failed:
    TestCheck "input.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    TestRunApiInputAndDataExamples = mReport & "TOTAL API INPUT EXAMPLES: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает фактическую и ожидаемую величины с заданным в этом наборе допуском, сохраняя числа в протоколе.
Private Sub TestNear(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    TestCheck name, Abs(actual - expected) <= 0.000001 * GeomMax(1#, Abs(expected))
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Учитывает отдельное условие проверки и сохраняет его имя в итоговом протоколе.
Private Sub TestCheck(ByVal name As String, ByVal passed As Boolean)
    If passed Then
        mPassed = mPassed + 1: mReport = mReport & "PASS: " & name & vbCrLf
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
    End If
End Sub
