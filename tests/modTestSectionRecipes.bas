Attribute VB_Name = "modTestSectionRecipes"
Option Explicit

Private mPassed As Long, mFailed As Long, mReport As String

' ДЛЯ ТЕСТОВ: штатный вход читает только выбранный рецепт и сохраняет snapshot.
Public Function RunWorkbookRecipeInputTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, fixture As Variant, rootName As String, names As Variant, name As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, registry As CSectionTypeRegistry, typeID As String
    Dim reader As CSectionInputSnapshotReader, context As CSectionRecipeContext, table As CSectionInputTable
    Dim rootRange As Object, changedName As Object, originalName As String, initialCell As Variant
    Dim definition As CSectionDefinition, builder As CSectionBuilder, model As CSectionModel, rebuilt As CSectionModel
    Dim mesh As TSectionMeshOptions, number As Long, reason As String, i As Long
    Dim calculation As Long, events As Boolean, updating As Boolean, diameterCell As Object, coverCell As Object
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    Set registry = New CSectionTypeRegistry
    For Each fixture In Array("Circle", "Rectangle", "RoundedMixed", "HollowOffset")
        config.Formula = saved
        modTestShapeArchitecture.ConfigureShapeFixture "Circle"
        modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        typeID = registry.SelectedRecipeType(settings): names = registry.RecipeTableNames(typeID): rootName = CStr(names(0))
        For Each name In Array("Circle.Diameter", "RectSet.B1", "RoundedRectangle.B", "HollowRectangle.B")
            Check "workbook.reader.noShapeKey." & CStr(fixture) & "." & CStr(name), Not settings.HasKey(CStr(name))
        Next name
        Set changedName = ThisWorkbook.Names(IIf(rootName = "rngCircleGeometry", "rngHollowRectangleGeometry", "rngCircleGeometry"))
        originalName = changedName.RefersTo: changedName.RefersTo = "=#REF!"
        Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
        Check "workbook.inactiveName." & CStr(fixture), model.RebarCount > 0 And model.ConcreteCount > 0
        changedName.RefersTo = originalName: Set changedName = Nothing
        Set rootRange = ThisWorkbook.Names(rootName).RefersToRange
        Set reader = New CSectionInputSnapshotReader: Set context = reader.Capture(ThisWorkbook, names, units)
        Set table = context.Table(rootName): initialCell = table.CellValue(1, 1)
        rootRange.Cells(1, 1).Value2 = "CHANGED_AFTER_CAPTURE"
        Check "workbook.table.immutable." & CStr(fixture), CStr(table.CellValue(1, 1)) = CStr(initialCell)
        Set definition = registry.CreateDefinition(typeID, context)
        Check "workbook.definition.sealed." & CStr(fixture), definition.IsSealed
        rootRange.ClearContents
        mesh = ReadSectionMeshOptions(settings, units)
        Set rebuilt = builder.Build(definition, mesh)
        Check "workbook.definition.meshCount." & CStr(fixture), rebuilt.ConcreteCount = model.ConcreteCount
        Check "workbook.definition.barCount." & CStr(fixture), rebuilt.RebarCount = model.RebarCount
        For i = 1 To model.RebarCount
            Near "workbook.definition.x." & CStr(fixture) & "." & CStr(i), rebuilt.RebarX(i), model.RebarX(i)
            Near "workbook.definition.y." & CStr(fixture) & "." & CStr(i), rebuilt.RebarY(i), model.RebarY(i)
        Next i
        config.Formula = saved
        Set changedName = ThisWorkbook.Names(rootName): originalName = changedName.RefersTo: changedName.RefersTo = "=#REF!"
        number = 0: reason = ""
        On Error Resume Next
        Set context = reader.Capture(ThisWorkbook, names, units)
        number = Err.Number: reason = Err.Description: Err.Clear
        On Error GoTo Failed
        changedName.RefersTo = originalName: Set changedName = Nothing
        Check "workbook.activeName.reject." & CStr(fixture), number <> 0 And InStr(1, reason, rootName, vbBinaryCompare) > 0
    Next fixture
    config.Formula = saved
    ThisWorkbook.Names("rngUnitSettings").RefersToRange.Cells(2, 2).Value2 = "mm"
    modTestShapeArchitecture.ConfigureShapeFixture "Circle"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set rootRange = ThisWorkbook.Names("rngCircleGeometry").RefersToRange
    Set context = reader.Capture(ThisWorkbook, Array("rngCircleGeometry"), units): Set table = context.Table("rngCircleGeometry")
    Set diameterCell = rootRange.Cells(table.RowByKey("Circle.Diameter"), 2)
    Set coverCell = rootRange.Cells(table.RowByKey("Rebar.AxisDistance"), 2)
    Application.Calculation = xlCalculationManual
    diameterCell.Value2 = 300#: coverCell.Formula = "=" & diameterCell.Address & "/4": rootRange.Calculate
    diameterCell.Value2 = 320#
    Near "workbook.manual.stale.fixture", CDbl(coverCell.Value2), 75#
    Set context = reader.Capture(ThisWorkbook, Array("rngCircleGeometry"), units): Set table = context.Table("rngCircleGeometry")
    Near "workbook.manual.capture.refreshed", table.Length(table.RowByKey("Rebar.AxisDistance"), 2, units), 80#
    names = registry.RecipeTableNames("Circle"): names(0) = "CHANGED_COPY"
    names = registry.RecipeTableNames("Circle")
    Check "workbook.registry.metadata.copy", CStr(names(0)) = "rngCircleGeometry"
    registry.RegisterRecipe "FixedMetadata", New CSectionRecipeCircle, Array()
    i = 0: names = registry.RecipeTableNames("FixedMetadata")
    For Each name In names: i = i + 1: Next name
    Check "workbook.registry.emptyMetadata", i = 0
    SetSystemInput "Geometry.Type", "FixedMetadata"
    settings.LoadFromWorkbook ThisWorkbook
    Check "workbook.registry.dynamicChoice", registry.SelectedRecipeType(settings) = "FixedMetadata"
    GoTo Finished
Failed:
    Check "workbook.recipe.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    On Error Resume Next
    If Not changedName Is Nothing Then changedName.RefersTo = originalName
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    On Error GoTo 0
    RunWorkbookRecipeInputTests = mReport & "TOTAL WORKBOOK RECIPE INPUT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: шум координаты касания не меняет дробление соседней ячейки.
Public Function RunRecipeMeshBoundaryTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, definition As CSectionDefinition
    Dim geometry As CGeometryRegion, mesh As CFiberMeshBuilder, reference As CFiberMeshBuilder
    Dim delta As Variant, i As Long, found As Boolean
    For Each delta In Array(0#, 0.000000000001, -0.000000000001)
        Set facade = New CSectionDefinitionBuilder: facade.Initialize "MeshTouch"
        Set path = facade.RectanglePath("Outer", 0#, 0#, 160#, 120#): facade.Outer "OUTER", path.ID
        Set path = facade.RectanglePath("Hole", 50#, 30#, 20#, 10# + CDbl(delta)): facade.Opening "HOLE", path.ID, "OUTER"
        Set definition = facade.Finish
        Set geometry = New CGeometryRegion: geometry.Initialize definition.ConcreteDomain
        Set mesh = New CFiberMeshBuilder: mesh.BuildMesh geometry, 40#, 20#, 1, 2
        found = False
        For i = 1 To mesh.FiberCount
            If Abs(mesh.FiberX(i) - 60#) <= GEOM_TOLERANCE And Abs(mesh.FiberY(i) - 50#) <= GEOM_TOLERANCE Then
                Near "mesh.touch.full.area." & CStr(delta), mesh.FiberArea(i), 800#
                found = True
            End If
        Next i
        Check "mesh.touch.full.cell." & CStr(delta), found
        If reference Is Nothing Then
            Set reference = mesh
        Else
            Check "mesh.touch.count." & CStr(delta), mesh.FiberCount = reference.FiberCount
            If mesh.FiberCount = reference.FiberCount Then
                For i = 1 To mesh.FiberCount
                    Near "mesh.touch.x." & CStr(i), mesh.FiberX(i), reference.FiberX(i)
                    Near "mesh.touch.y." & CStr(i), mesh.FiberY(i), reference.FiberY(i)
                    Near "mesh.touch.area." & CStr(i), mesh.FiberArea(i), reference.FiberArea(i)
                Next i
            End If
        End If
    Next delta
    GoTo Finished
Failed:
    Check "mesh.touch.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRecipeMeshBoundaryTests = mReport & "TOTAL RECIPE MESH BOUNDARY: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: общие правила без таблиц типов или старых раскладчиков.
Public Function RunRecipeRuleTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim facade As CSectionDefinitionBuilder, options As CRebarRuleOptions, model As CSectionModel
    Dim builder As CSectionBuilder, path As CGeometryPath, mode As Long, scenario As Long, expected As Variant, i As Long
    For mode = IncludeBoth To IncludeEnd
        Set facade = RuleFacade: Set options = New CRebarRuleOptions
        options.EndMode = mode: options.StartSetback = 10#: options.EndSetback = 20#
        facade.RebarByCount "Row.Main", "Row", 3, 10#, options
        options.StartSetback = 999# ' Копия options в definition уже независима.
        Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
        Select Case mode
            Case IncludeBoth: expected = Array(10#, 45#, 80#)
            Case ExcludeBoth: expected = Array(27.5, 45#, 62.5)
            Case IncludeStart: expected = Array(10#, 10# + 70# / 3#, 10# + 140# / 3#)
            Case IncludeEnd: expected = Array(10# + 70# / 3#, 10# + 140# / 3#, 80#)
        End Select
        Check "ends.count." & CStr(mode), model.RebarCount = 3
        For i = 1 To 3: Near "ends.station." & CStr(mode) & "." & CStr(i), model.RebarX(i), CDbl(expected(i - 1)): Next i
    Next mode
    Set facade = RuleFacade: Set options = New CRebarRuleOptions
    options.StartSetback = 10#: options.EndSetback = 20#
    facade.RebarByCount "One", "Row", 1, 10#, options
    Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
    Near "one.midpoint", model.RebarX(1), 45#
    Set facade = RuleFacade: Set options = New CRebarRuleOptions
    options.EndMode = ClosedUniform
    Set path = facade.CirclePath("Closed", 0#, 0#, 100#)
    facade.RebarByCount "Closed.Row", "Closed", 4, 10#, options
    Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
    Check "closed.count", model.RebarCount = 4
    Near "closed.start.x", model.RebarX(1), 100#: Near "closed.start.y", model.RebarY(1), 0#
    Near "closed.last.x", model.RebarX(4), 0#: Near "closed.last.y", model.RebarY(4), -100#
    Set facade = RuleFacade: Set options = New CRebarRuleOptions
    options.StartSetback = 10#: options.EndSetback = 20#
    facade.RebarByMaxStep "Step", "Row", 30#, 10#, options
    Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
    Check "maxstep.count", model.RebarCount = 4
    Near "maxstep.first", model.RebarX(1), 10#: Near "maxstep.last", model.RebarX(4), 80#
    Near "maxstep.actual", model.RebarX(2) - model.RebarX(1), 70# / 3#
    Set facade = RuleFacade: Set options = New CRebarRuleOptions
    options.EndMode = ExcludeBoth
    facade.RebarByFixedStep "Fixed", "Row", 30#, 10#, options
    Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
    Check "fixed.count", model.RebarCount = 4
    Near "fixed.first", model.RebarX(1), 5#: Near "fixed.last", model.RebarX(4), 95#
    For i = 2 To 4: Near "fixed.step." & CStr(i), model.RebarX(i) - model.RebarX(i - 1), 30#: Next i
    Set facade = RuleFacade
    facade.RebarAtStations "Explicit", "Row", Array(10#, 30#, 90#), 10#
    Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
    Check "explicit.count", model.RebarCount = 3: Near "explicit.last", model.RebarX(3), 90#
    Set facade = RuleFacade: Set options = New CRebarRuleOptions
    options.SourcePrefix = "P"
    facade.RebarByCount "Base", "Row", 3, 10#, options
    Set options = New CRebarRuleOptions
    options.NormalOffset = 10#: options.EverySecond = True: options.SourceSuffix = "_Extra"
    facade.RebarDependent "Extra", "Base", 10#, options
    Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
    Check "dependent.count", model.RebarCount = 5
    Near "dependent.first.x", model.RebarX(4), 0#: Near "dependent.last.x", model.RebarX(5), 100#
    Near "dependent.normal", model.RebarY(5), 10#
    Check "dependent.source", model.RebarSourceName(5) = "P3_Extra"
    Set facade = RuleFacade: Set options = New CRebarRuleOptions
    options.SourcePrefix = "P": facade.RebarByCount "Base", "Row", 3, 10#, options
    Set path = facade.LinePath("Target", 0#, 100#, 100#, 100#)
    Set options = New CRebarRuleOptions
    options.DirectionY = 1#: options.SkipMissingProjection = True
    options.StartSetback = 10#: options.EndSetback = 10#: options.SourceSuffix = "_Projected"
    facade.RebarProjected "Projection", "Target", "Base", 10#, options
    Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
    Check "projection.count", model.RebarCount = 4: Check "projection.skips", builder.Rules.SkippedProjections = 2
    Near "projection.x", model.RebarX(4), 50#: Near "projection.y", model.RebarY(4), 100#
    Check "projection.source", model.RebarSourceName(4) = "P2_Projected"
    ' Первое пересечение исключено setback; второе на том же пути допустимо.
    Set facade = RuleFacade
    facade.RebarByCount "Base", "Row", 1, 10#
    Set path = New CGeometryPath: path.Initialize "TwoHits"
    path.AddLine "First", 0#, 50#, 100#, 50#
    path.AddLine "Side", 100#, 50#, 100#, 150#
    path.AddLine "Second", 100#, 150#, 0#, 150#
    path.Seal: facade.AddPath path
    Set options = New CRebarRuleOptions
    options.DirectionY = 1#: options.StartSetback = 110#: options.SourceSuffix = "_Hit"
    facade.RebarProjected "SecondHit", "TwoHits", "Base", 10#, options
    Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
    Check "projection.secondhit.count", model.RebarCount = 2
    Near "projection.secondhit.x", model.RebarX(2), 50#: Near "projection.secondhit.y", model.RebarY(2), 150#
    Set facade = RuleFacade
    Set path = facade.LinePath("Target", 0#, 100#, 100#, 100#)
    facade.PathGroup "Both", Array("Row", "Target")
    facade.RebarByCount "BothRows", "Both", 2, 10#
    Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
    Check "group.perpath", model.RebarCount = 4: Near "group.secondpath", model.RebarY(3), 100#
    Dim query As CSectionGeometryQuery, region As CConcreteRegion
    Set query = New CSectionGeometryQuery: Set region = query.CircleRegion(10#, -20#, 100#)
    Near "distance.center", query.BoundaryDistance(region, 10#, -20#), 100#
    Near "distance.inside", query.BoundaryDistance(region, 90#, -20#), 20#
    Near "distance.outside", query.BoundaryDistance(region, 150#, -20#), 40#
    For scenario = 0 To 7: Check "rule.invalid." & CStr(scenario), InvalidRuleRejected(scenario): Next scenario
    TestPublicationAndCoincidence
    TestLayoutIndex
    TestAdjacentRectangles
    TestAnnotationMeasurements
    TestRuleJointTangents
    TestProjectionNormalsAndOrdinals
    GoTo Finished
Failed:
    Check "rules.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRecipeRuleTests = mReport & "TOTAL RECIPE RULES: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Private Sub TestProjectionNormalsAndOrdinals()
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim builder As CSectionBuilder, model As CSectionModel, mode As Long, i As Long
    Set facade = RuleFacade
    Set path = facade.ArcPath("SourceCurve", 0#, 0#, 100#, 0#, GEOM_PI / 2#)
    Set options = New CRebarRuleOptions: options.AnnotationGroup = "CurveBars"
    facade.RebarByCount "CurveMain", path.ID, 3, 10#, options
    Set path = facade.ArcPath("TargetCurve", 0#, 0#, 200#, 0#, GEOM_PI / 2#)
    Set options = New CRebarRuleOptions: options.SourceNormalSign = -1#: options.NewPublicationRoot = True
    facade.RebarProjected "Radial", path.ID, "CurveMain", 10#, options
    options.SourceNormalSign = 1#
    facade.RebarRowExtremeLabel "CurveBars", "CurveMain", 0#, 1#, 0#, 1#
    Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
    Check "projection.normal.count", model.RebarCount = 6
    For i = 1 To 3
        Near "projection.normal.x." & CStr(i), model.RebarX(i + 3), 2# * model.RebarX(i)
        Near "projection.normal.y." & CStr(i), model.RebarY(i + 3), 2# * model.RebarY(i)
    Next i
    Near "extreme.anchor.start", builder.Rebars.AnnotationStartX(1), 0#
    Near "extreme.anchor.end", builder.Rebars.AnnotationEndX(1), 0#
    Near "extreme.anchor.y", builder.Rebars.AnnotationStartY(1), 100#
    For mode = 0 To 1
        Set facade = RuleFacade
        facade.RebarByCount "Base", "Row", 4, 10#
        Set path = facade.LinePath("LimitedTarget", 20#, 100#, 80#, 100#)
        Set options = New CRebarRuleOptions: options.DirectionY = 1#: options.SkipMissingProjection = True
        options.SuccessfulProjectionOrdinal = (mode = 1): options.NewPublicationRoot = True
        options.Publication = PublishPerSourcePosition
        facade.RebarProjected "Projected", path.ID, "Base", 10#, options
        options.SuccessfulProjectionOrdinal = Not options.SuccessfulProjectionOrdinal
        Set options = New CRebarRuleOptions: options.EverySecond = True: options.NormalOffset = 10#
        options.Publication = PublishPerSourcePosition
        facade.RebarDependent "Extra", "Projected", 10#, options
        Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
        Check "projection.ordinal.count." & CStr(mode), model.RebarCount = 7 And builder.Rules.SkippedProjections = 2
        Near "projection.publication.first." & CStr(mode), model.RebarY(4), 0#
        If mode = 1 Then
            Near "projection.ordinal.success.x", model.RebarX(6), 100# / 3#
            Near "projection.ordinal.success.y", model.RebarY(6), 110#
        Else
            Near "projection.ordinal.source.x", model.RebarX(7), 200# / 3#
            Near "projection.ordinal.source.y", model.RebarY(7), 110#
        End If
    Next mode
End Sub

Private Sub TestRuleJointTangents()
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim builder As CSectionBuilder, model As CSectionModel, incoming As Long
    For incoming = 0 To 1
        Set facade = RuleFacade
        Set path = New CGeometryPath: path.Initialize "Joint"
        path.AddLine "Incoming", 0#, 0#, 100#, 0#
        path.AddLine "Outgoing", 100#, 0#, 100#, 100#
        path.Seal: facade.AddPath path
        Set options = New CRebarRuleOptions: options.IncomingTangentAtJoint = (incoming = 1)
        facade.RebarAtStations "JointMain", path.ID, Array(100#), 10#, options
        options.IncomingTangentAtJoint = Not options.IncomingTangentAtJoint
        Set options = New CRebarRuleOptions: options.NormalOffset = 10#
        facade.RebarDependent "JointExtra", "JointMain", 10#, options
        Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
        If incoming = 1 Then
            Near "joint.rule.incoming.x", model.RebarX(2), 100#: Near "joint.rule.incoming.y", model.RebarY(2), 10#
        Else
            Near "joint.rule.default.x", model.RebarX(2), 90#: Near "joint.rule.default.y", model.RebarY(2), 0#
        End If
    Next incoming
End Sub

Private Sub TestAnnotationMeasurements()
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim builder As CSectionBuilder, model As CSectionModel, annotations As CSectionAnnotations
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Set facade = RuleFacade
    Set path = facade.ArcPath("Curve", 0#, 0#, 100#, 0#, GEOM_PI / 2#)
    Set options = New CRebarRuleOptions: options.AnnotationGroup = "CurveBars": options.EndMode = ExcludeBoth
    facade.RebarByCount "CurveMain", path.ID, 3, 10#, options
    facade.RebarRowExtentLabel "CurveBars", "CurveMain", -1#, 0#, 20#, True
    facade.HorizontalDistance "REF", "Reference", 0#, 123.6, 200#, 1#, "", 1#
    facade.HorizontalDistance "NEG", "Negative", 0#, -123.6, 200#, 1#, "", 1#
    facade.HorizontalDistance "EXACT", "Exact", 0#, 123.6, 200#
    Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
    Check "measurements.rule.bounds", builder.Rules.RuleBounds("CurveMain", minX, maxX, minY, maxY)
    Near "measurements.min.x", minX, 100# * Cos(3# * GEOM_PI / 8#)
    Near "measurements.max.x", maxX, 100# * Cos(GEOM_PI / 8#)
    Near "measurements.min.y", minY, 100# * Sin(GEOM_PI / 8#)
    Near "measurements.max.y", maxY, 100# * Sin(3# * GEOM_PI / 8#)
    Near "measurements.anchor.x", builder.Rebars.AnnotationStartX(1), minX
    Near "measurements.anchor.start", builder.Rebars.AnnotationStartY(1), maxY
    Near "measurements.anchor.end", builder.Rebars.AnnotationEndY(1), minY
    Set annotations = model.Annotations
    Near "measurements.round.positive", annotations.Value(1), 124#
    Near "measurements.round.negative", annotations.Value(2), -124#
    Near "measurements.round.exact", annotations.Value(3), 123.6
    Near "measurements.coordinates.unchanged", annotations.EndX(1), 123.6
    Check "measurements.count", AnnotatedBarCount(annotations) = 3
End Sub

Private Sub TestPublicationAndCoincidence()
    Dim facade As CSectionDefinitionBuilder, options As CRebarRuleOptions, builder As CSectionBuilder
    Dim model As CSectionModel, expected As Variant, i As Long, memberships As Variant
    Set facade = RuleFacade: Set options = New CRebarRuleOptions
    options.PositionOrder = FromPathMiddle: options.Publication = PublishPerSourcePosition
    options.StartSetback = 10#: options.EndSetback = 20#: options.AnnotationGroup = "Base"
    facade.RebarByCount "Base", "Row", 5, 10#, options
    Set options = New CRebarRuleOptions
    options.Publication = PublishPerSourcePosition: options.EverySecond = True: options.NormalOffset = 10#
    facade.RebarDependent "Extra", "Base", 10#, options
    facade.RebarRowLabel "Base", "Base", 0#, -1#
    Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
    expected = Array(45#, 45#, 62.5, 27.5, 80#, 80#, 10#, 10#)
    Check "middle.count", model.RebarCount = 8
    For i = 1 To 8: Near "middle.order." & CStr(i), model.RebarX(i), CDbl(expected(i - 1)): Next i
    Near "middle.first.extra", model.RebarY(2), 10#
    Near "middle.anchor.first", builder.Rebars.AnnotationStartX(1), 10#
    Near "middle.anchor.last", builder.Rebars.AnnotationEndX(1), 80#
    Set facade = RuleFacade: Set options = New CRebarRuleOptions
    options.AnnotationGroup = "First": facade.RebarByCount "First", "Row", 2, 10#, options
    Set options = New CRebarRuleOptions
    options.AnnotationGroup = "Second": options.Coincidence = MergeIdentical
    facade.RebarByCount "Second", "Row", 2, 10#, options
    facade.RebarRowLabel "First", "First", 0#, -1#
    facade.RebarRowLabel "Second", "Second", 0#, 1#
    Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
    Check "coincidence.physical.count", model.RebarCount = 2 And builder.Rules.CreatedPositions = 2
    Check "coincidence.requested", builder.Rules.RequestedPositions = 4 And builder.Rules.MergedPositions = 2
    memberships = builder.Rebars.BarMemberships(1)
    Check "coincidence.memberships", UBound(memberships) = 1 And memberships(0) = "First" And memberships(1) = "Second"
    Check "coincidence.annotation.owner", builder.Rebars.BarAnnotationGroupName(1) = "First"
    Check "coincidence.labels", model.Annotations.Count = 1 And model.Annotations.Text(1) = "2" & ChrW$(216) & "10"
    Check "coincidence.labels.total", AnnotatedBarCount(model.Annotations) = model.RebarCount
    For i = 0 To 3: Check "coincidence.invalid." & CStr(i), InvalidCoincidenceRejected(i): Next i
    Set facade = RuleFacade
    facade.RebarByCount "First", "Row", 2, 10#
    Set options = New CRebarRuleOptions: options.Coincidence = KeepPublishedOwner
    facade.RebarByCount "Second", "Row", 2, 20#, options
    Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
    Check "coincidence.explicit.owner", model.RebarCount = 2 And builder.Rules.Conflicts = 2
    Near "coincidence.owner.diameter", model.RebarDiameter(1), 10#
    TestAnnotationOwnership
End Sub

' Одна общая позиция, полностью совпавшая третья группа и разные диаметры
' не должны повторно увеличивать количество стержней в подписях.
Private Sub TestAnnotationOwnership()
    Dim facade As CSectionDefinitionBuilder, options As CRebarRuleOptions, builder As CSectionBuilder
    Dim model As CSectionModel, path As CGeometryPath
    Dim annotations As CSectionAnnotations, labels As CRebarGroupAnnotationBuilder
    Set facade = RuleFacade: Set options = New CRebarRuleOptions
    options.AnnotationGroup = "First": facade.RebarByCount "First", "Row", 2, 10#, options
    Set path = facade.LinePath("SecondPath", 100#, 0#, 200#, 0#)
    Set options = New CRebarRuleOptions: options.Coincidence = KeepPublishedOwner
    options.AnnotationGroup = "Second": facade.RebarByCount "Second", path.ID, 3, 20#, options
    Set path = facade.LinePath("ThirdPath", 100#, 0#, 150#, 0#)
    options.AnnotationGroup = "Third": facade.RebarByCount "Third", path.ID, 2, 30#, options
    Set path = facade.LinePath("OtherDiameter", 300#, 0#, 400#, 0#)
    options.AnnotationGroup = "First": facade.RebarByCount "Extra", path.ID, 1, 12#, options
    facade.RebarRowLabel "First", "First", 0#, -1#
    facade.RebarRowLabel "Second", "Second", 0#, 1#
    facade.RebarRowLabel "Third", "Third", 0#, 1#
    Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
    Check "annotations.partial.physical", model.RebarCount = 5 And builder.Rules.MergedPositions = 3
    Check "annotations.partial.groups", model.Annotations.Count = 2
    Check "annotations.partial.first", model.Annotations.Text(1) = "2" & ChrW$(216) & "10 + 1" & ChrW$(216) & "12"
    Check "annotations.partial.second", model.Annotations.Text(2) = "2" & ChrW$(216) & "20"
    Check "annotations.partial.total", AnnotatedBarCount(model.Annotations) = model.RebarCount
    builder.Rebars.AddAnnotationAnchor "first", 300#, 0#, 400#, 0#, 0#, -1#
    Set annotations = New CSectionAnnotations: Set labels = New CRebarGroupAnnotationBuilder
    labels.AddRebarLabels annotations, builder.Rebars
    Check "annotations.repeated.anchor", annotations.Count = 2
    Check "annotations.repeated.total", AnnotatedBarCount(annotations) = model.RebarCount
End Sub

' Независимый подсчет по готовому пользовательскому тексту, а не по layout.
Private Function AnnotatedBarCount(ByVal annotations As CSectionAnnotations) As Long
    Dim i As Long, part As Variant, position As Long
    For i = 1 To annotations.Count
        If annotations.AnnotationType(i) = "REBAR_ANNOTATION" Then
            For Each part In Split(annotations.Text(i), "+")
                position = InStr(1, CStr(part), ChrW$(216), vbBinaryCompare)
                If position = 0 Then Err.Raise vbObjectError + 5532, "modTestSectionRecipes", "Неверный формат групповой подписи: " & CStr(part)
                AnnotatedBarCount = AnnotatedBarCount + CLng(Trim$(Left$(CStr(part), position - 1)))
            Next part
        End If
    Next i
End Function

Private Function InvalidCoincidenceRejected(ByVal scenario As Long) As Boolean
    On Error GoTo Rejected
    Dim facade As CSectionDefinitionBuilder, options As CRebarRuleOptions, builder As CSectionBuilder, model As CSectionModel
    Set facade = RuleFacade
    facade.RebarByCount "First", "Row", 2, 10#
    Set options = New CRebarRuleOptions
    Select Case scenario
        Case 0: facade.RebarByCount "Second", "Row", 2, 10#, options
        Case 1: options.Coincidence = MergeIdentical: facade.RebarByCount "Second", "Row", 2, 20#, options
        Case 2: options.Coincidence = MergeIdentical: facade.RebarByCount "Second", "Row", 2, 10#, options, "OtherSteel"
        Case 3
            facade.RebarByCount "Huge", "Row", 2147483647, 10#
            options.NormalOffset = 10#: facade.RebarDependent "Overflow", "Huge", 10#, options
    End Select
    Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
    Exit Function
Rejected:
    InvalidCoincidenceRejected = (Err.Number = vbObjectError + 5530)
End Function

' Независимый линейный oracle проверяет индекс на границах bucket,
' отрицательных координатах, далеко от нуля и после Clear.
Private Sub TestLayoutIndex()
    Dim layout As CRebarLayout, coordinates As Variant, value As Variant, x As Double, y As Double
    Dim i As Long, j As Long, expected As Long, queryX As Double, delta As Variant
    Set layout = New CRebarLayout
    coordinates = Array(-1000000#, -10# * GEOM_TOLERANCE, -4# * GEOM_TOLERANCE, _
        -GEOM_TOLERANCE / 2#, 4# * GEOM_TOLERANCE, 10# * GEOM_TOLERANCE, 1000000#)
    For Each value In coordinates
        i = i + 1: x = CDbl(value): y = x / 2#
        layout.AddBar "I" & CStr(i), x, y, 0.01, 0#, "Rebar"
    Next value
    For i = 1 To layout.Count
        For Each delta In Array(-1.01, -1#, -0.9, 0#, 0.9, 1#, 1.01)
            queryX = layout.X(i) + CDbl(delta) * GEOM_TOLERANCE: y = layout.Y(i)
            expected = 0
            For j = 1 To layout.Count
                If Abs(layout.X(j) - queryX) <= GEOM_TOLERANCE And Abs(layout.Y(j) - y) <= GEOM_TOLERANCE Then
                    expected = j: Exit For
                End If
            Next j
            Check "index.oracle." & CStr(i) & "." & CStr(delta), layout.FindBarAt(queryX, y) = expected
        Next delta
    Next i
    layout.Clear
    Check "index.clear", layout.Count = 0 And Not layout.HasBarAt(1000000#, 500000#)
    layout.AddBar "I1", 1000000#, 500000#, 10#, 0#, "Rebar"
    Check "index.reuse", layout.FindBarAt(1000000#, 500000#) = 1
End Sub

Private Function RuleFacade() As CSectionDefinitionBuilder
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "RuleFixture"
    Set path = facade.CirclePath("Outer", 0#, 0#, 1000#): facade.Outer "OUTER", path.ID
    Set path = facade.LinePath("Row", 0#, 0#, 100#, 0#)
    Set RuleFacade = facade
End Function

Private Function BuildFixture(ByVal facade As CSectionDefinitionBuilder, ByVal builder As CSectionBuilder) As CSectionModel
    Dim mesh As TSectionMeshOptions
    mesh.StepX = 1000#: mesh.StepY = 1000#: mesh.BoundarySubdivisions = 1: mesh.MaterialID = 1
    Set BuildFixture = builder.Build(facade.Finish, mesh)
End Function

Private Function InvalidRuleRejected(ByVal scenario As Long) As Boolean
    On Error GoTo Rejected
    Dim facade As CSectionDefinitionBuilder, options As CRebarRuleOptions, builder As CSectionBuilder, model As CSectionModel, path As CGeometryPath
    Set facade = RuleFacade: Set options = New CRebarRuleOptions
    Select Case scenario
        Case 0: options.StartSetback = -1#: facade.RebarByCount "Bad", "Row", 2, 10#, options
        Case 1: options.StartSetback = 60#: options.EndSetback = 60#: facade.RebarByCount "Bad", "Row", 2, 10#, options
        Case 2: facade.RebarByFixedStep "Bad", "Row", 30#, 10#, options
        Case 3: facade.RebarByMaxStep "Bad", "Row", 0#, 10#, options
        Case 4
            Set path = facade.CirclePath("Closed", 0#, 0#, 100#)
            options.EndMode = ClosedUniform: options.StartSetback = 1#
            facade.RebarByCount "Bad", "Closed", 2, 10#, options
        Case 5
            facade.RebarDependent "A", "B", 10#, options
            facade.RebarDependent "B", "A", 10#, options
        Case 6: facade.RebarByCount "Bad", "Missing", 2, 10#, options
        Case 7
            Set path = facade.LinePath("Target", 0#, 100#, 100#, 100#)
            facade.RebarByCount "Base", "Row", 2, 10#
            options.DirectionY = -1#
            facade.RebarProjected "Bad", "Target", "Base", 10#, options
    End Select
    Set builder = New CSectionBuilder: Set model = BuildFixture(facade, builder)
    Exit Function
Rejected:
    InvalidRuleRejected = (Err.Number = vbObjectError + 5530)
End Function

Private Sub Check(ByVal name As String, ByVal passed As Boolean)
    If passed Then mPassed = mPassed + 1 Else mFailed = mFailed + 1
    mReport = mReport & IIf(passed, "PASS: ", "FAIL: ") & name & vbCrLf
End Sub

Private Sub Near(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    Check name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) <= 0.0000001
End Sub

' ДЛЯ ТЕСТОВ: роли не зависят от ID; несколько Opening и независимость
' опубликованной области от изменяемых копий, выдаваемых потребителю.
Public Function RunRecipeDefinitionTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, definition As CSectionDefinition
    Dim region As CConcreteRegion, copy As CConcreteRegion, query As CSectionGeometryQuery, contours As CSectionContours
    Set facade = RuleFacade
    Set path = facade.CirclePath("VoidA", -100#, 0#, 20#): facade.Opening "A", path.ID, "OUTER"
    Set path = facade.CirclePath("VoidB", 100#, 0#, 20#): facade.Opening "B", path.ID, "OUTER"
    Set definition = facade.Finish: Set region = definition.ConcreteDomain
    Set query = New CSectionGeometryQuery
    Check "definition.sealed", definition.IsSealed
    Check "definition.loops", definition.LoopCount = 3 And region.LoopCount = 3
    Check "definition.area", Abs(region.Area - GEOM_PI * (1000000# - 800#)) < 0.00001
    Check "definition.concrete", query.ContainsPoint(region, 0#, 0#)
    Check "definition.hole.a", Not query.ContainsPoint(region, -100#, 0#)
    Check "definition.hole.b", Not query.ContainsPoint(region, 100#, 0#)
    Set contours = definition.MaterialContours
    Check "definition.roles", contours.LoopRole(1) = "Outer" And contours.LoopRole(2) = "Opening" And contours.LoopRole(3) = "Opening"
    Check "definition.ids", contours.LoopID(2) = "A" And contours.LoopID(3) = "B"
    contours.Clear
    Check "definition.contours.copy", definition.MaterialContours.Count = 3
    Set copy = definition.ConcreteDomain: copy.SetOwner "Changed", "SP35", "G1"
    Check "definition.region.copy", Len(definition.ConcreteDomain.OwnerID) = 0
    Dim scenario As Long
    For scenario = 0 To 4: Check "definition.invalid." & CStr(scenario), InvalidDefinitionRejected(scenario): Next scenario
    GoTo Finished
Failed:
    Check "definition.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRecipeDefinitionTests = mReport & "TOTAL RECIPE DEFINITION: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Private Function InvalidDefinitionRejected(ByVal scenario As Long) As Boolean
    On Error GoTo Rejected
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, definition As CSectionDefinition
    Set facade = RuleFacade
    Select Case scenario
        Case 0: Set path = facade.LinePath("Row", 0#, 0#, 10#, 0#)
        Case 1: facade.Outer "Other", "Row"
        Case 2
            Set path = facade.CirclePath("Void", 0#, 0#, 20#)
            facade.Opening "VoidLoop", path.ID, "Missing"
        Case 3
            Set definition = facade.Finish
            Set path = facade.LinePath("AfterSeal", 0#, 0#, 10#, 0#)
        Case 4
            Set definition = facade.Finish: Set path = definition.Path("Row")
            path.AddLine "AfterSeal", 100#, 0#, 200#, 0#
    End Select
    Set definition = facade.Finish
    Exit Function
Rejected:
    InvalidDefinitionRejected = (Err.Number = vbObjectError + 5530 Or Err.Number = vbObjectError + 5520)
End Function

' ДЛЯ ТЕСТОВ: адресность ошибок и неактивные поля Circle.
Public Function RunCircleRecipeInputTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, calculation As Long, events As Boolean, updating As Boolean, scenario As Long
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    For scenario = 0 To 4: Check "circle.input." & CStr(scenario), CircleInputCase(scenario): Next scenario
    GoTo Finished
Failed:
    Check "circle.input.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    RunCircleRecipeInputTests = mReport & "TOTAL CIRCLE INPUT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Private Function CircleInputCase(ByVal scenario As Long) As Boolean
    Dim expectedReference As String, key As String, settings As CSystemSettingsReader, units As CUnitSystem
    Dim reader As CSectionInputSnapshotReader, context As CSectionRecipeContext, recipe As ISectionRecipe
    Dim definition As CSectionDefinition, table As CSectionInputTable, builder As CSectionBuilder, mesh As TSectionMeshOptions, model As CSectionModel
    On Error GoTo Rejected
    modTestShapeArchitecture.ConfigureShapeFixture "Circle"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Select Case scenario
        Case 0: key = "Rebar.Count": SetCircleInput key, 1
        Case 1: key = "Rebar.AxisDistance": SetCircleInput key, 10# / units.InputLengthToInternal(1#)
        Case 2: key = "Rebar.Diameter2": SetCircleInput key, -1#
        Case 3
            SetCircleInput "Rebar.Count", 0: SetCircleInput "Rebar.AxisDistance", "bad"
            SetCircleInput "Rebar.Diameter2", "bad": SetCircleInput "Rebar.Loc2row", "bad"
        Case 4
            key = "Rebar.Diameter2": SetCircleInput key, 400# / units.InputLengthToInternal(1#)
            SetCircleInput "Rebar.Loc2row", "SideBySide"
    End Select
    Set reader = New CSectionInputSnapshotReader: Set context = reader.Capture(ThisWorkbook, Array("rngCircleGeometry"), units)
    Set table = context.Table("rngCircleGeometry")
    If Len(key) > 0 Then expectedReference = table.SourceReference(table.RowByKey(key), 2)
    Set recipe = New CSectionRecipeCircle: Set definition = recipe.CreateDefinition(context)
    If scenario = 3 Then CircleInputCase = (definition.RuleCount = 0): Exit Function
    mesh.StepX = 40#: mesh.StepY = 30#: mesh.BoundarySubdivisions = 2: mesh.MaterialID = 1
    Set builder = New CSectionBuilder: Set model = builder.Build(definition, mesh)
    Exit Function
Rejected:
    CircleInputCase = (Len(expectedReference) > 0 And InStr(1, Err.Description, expectedReference, vbBinaryCompare) > 0)
End Function

' ДЛЯ ТЕСТОВ: активные/неактивные поля и адресность RoundedRectangle.
Public Function RunRoundedRecipeInputTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, scenario As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    For scenario = 0 To 13
        config.Formula = saved
        Check "rounded.input." & CStr(scenario), RoundedInputCase(scenario)
    Next scenario
    GoTo Finished
Failed:
    Check "rounded.input.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.EnableEvents = events: Application.ScreenUpdating = updating
    RunRoundedRecipeInputTests = mReport & "TOTAL ROUNDED INPUT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Private Function RoundedInputCase(ByVal scenario As Long) As Boolean
    Dim range As Object, row As Long, column As Long, valid As Boolean, reference As String
    Dim settings As CSystemSettingsReader, units As CUnitSystem, reader As CSectionInputSnapshotReader
    Dim context As CSectionRecipeContext, table As CSectionInputTable, recipe As ISectionRecipe, definition As CSectionDefinition
    Dim builder As CSectionBuilder, model As CSectionModel, mesh As TSectionMeshOptions
    On Error GoTo Rejected
    modTestShapeArchitecture.ConfigureShapeFixture "Circle"
    modTestShapeArchitecture.ConfigureShapeFixture "RoundedSimple"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set range = ThisWorkbook.Names("rngRoundedRectangleGeometry").RefersToRange
    Select Case scenario
        Case 0: row = 13: column = 4: range.Cells(row, column).Value2 = 1.5
        Case 1: row = 13: column = 3: range.Cells(row, column).Value2 = -1#
        Case 2: row = 13: column = 2: range.Cells(row, column).Value2 = CDbl(range.Cells(row, 3).Value2) / 2#
        Case 3: row = 19: column = 2: range.Cells(row, column).Value2 = -1#
        Case 4: row = 19: column = 3: range.Cells(19, 2).Value2 = 10#: range.Cells(row, column).Value2 = "bad"
        Case 5: row = 19: column = 4: range.Cells(19, 2).Value2 = 10#: range.Cells(row, column).Value2 = "bad"
        Case 6
            valid = True: range.Cells(8, 2).Value2 = "inactive": range.Cells(10, 2).Value2 = "inactive"
        Case 7
            valid = True: range.Cells(13, 4).Value2 = 0: range.Cells(13, 2).Value2 = "inactive"
            range.Cells(19, 2).Value2 = "inactive": range.Cells(19, 3).Value2 = "inactive"
        Case 8
            valid = True: range.Cells(13, 3).Value2 = 0: range.Cells(13, 2).Value2 = "inactive"
            range.Cells(19, 2).Value2 = "inactive": range.Cells(19, 3).Value2 = "inactive"
        Case 9
            valid = True: range.Cells(19, 2).Value2 = 0: range.Cells(19, 5).Value2 = 0
            range.Cells(19, 3).Value2 = "inactive": range.Cells(19, 4).Value2 = "inactive"
        Case 10: row = 9: column = 2: range.Cells(row, column).Value2 = 1000#
        Case 11
            valid = True: range.Cells(7, 2).Value2 = "Tapered": range.Cells(8, 2).Value2 = 100#
            range.Cells(9, 2).Value2 = 220#: range.Cells(10, 2).Value2 = 0#
        Case 12: valid = True: range.Cells(9, 2).Value2 = 200#: range.Cells(9, 3).Value2 = 200#
        Case 13: row = 13: column = 2: range.Cells(row, column).Value2 = 380#
    End Select
    Set reader = New CSectionInputSnapshotReader: Set context = reader.Capture(ThisWorkbook, Array("rngRoundedRectangleGeometry"), units)
    Set table = context.Table("rngRoundedRectangleGeometry")
    If Not valid Then reference = table.SourceReference(row, column)
    Set recipe = New CSectionRecipeRoundedRectangle: Set definition = recipe.CreateDefinition(context)
    mesh.StepX = 40#: mesh.StepY = 30#: mesh.BoundarySubdivisions = 2: mesh.MaterialID = 1
    Set builder = New CSectionBuilder: Set model = builder.Build(definition, mesh)
    RoundedInputCase = valid
    Exit Function
Rejected:
    RoundedInputCase = (Not valid And Len(reference) > 0 And InStr(1, Err.Description, reference, vbBinaryCompare) > 0)
    If Not RoundedInputCase Then mReport = mReport & "ROUNDED INPUT DETAIL: " & CStr(scenario) & "; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: адресность размеров/проекционных ограничителей и неактивные поля.
Public Function RunHollowRecipeInputTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, scenario As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    For scenario = 0 To 13
        config.Formula = saved
        Check "hollow.input." & CStr(scenario), HollowInputCase(scenario)
    Next scenario
    GoTo Finished
Failed:
    Check "hollow.input.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.EnableEvents = events: Application.ScreenUpdating = updating
    RunHollowRecipeInputTests = mReport & "TOTAL HOLLOW INPUT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Private Function HollowInputCase(ByVal scenario As Long) As Boolean
    Dim range As Object, row As Long, column As Long, valid As Boolean, reference As String, factor As Double
    Dim settings As CSystemSettingsReader, units As CUnitSystem, reader As CSectionInputSnapshotReader
    Dim context As CSectionRecipeContext, table As CSectionInputTable, recipe As ISectionRecipe, definition As CSectionDefinition
    Dim builder As CSectionBuilder, model As CSectionModel, mesh As TSectionMeshOptions
    On Error GoTo Rejected
    modTestShapeArchitecture.ConfigureShapeFixture "Circle"
    modTestShapeArchitecture.ConfigureShapeFixture "HollowCentered"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    Set range = ThisWorkbook.Names("rngHollowRectangleGeometry").RefersToRange
    Select Case scenario
        Case 0: row = 13: column = 4: range.Cells(row, column).Value2 = 1.5
        Case 1: row = 13: column = 3: range.Cells(row, column).Value2 = -1#
        Case 2: row = 13: column = 2: range.Cells(row, column).Value2 = CDbl(range.Cells(row, 3).Value2) / 2#
        Case 3: row = 15: column = 2: range.Cells(15, 3).Value2 = 0#: range.Cells(row, column).Value2 = -1#
        Case 4: row = 7: column = 6: range.Cells(row, column).Value2 = -1#
        Case 5: row = 3: column = 2: range.Cells(row, column).Value2 = 1000# / factor
        Case 6
            valid = True: range.Cells(13, 4).Value2 = 0: range.Cells(13, 2).Value2 = "inactive"
            range.Cells(23, 2).Value2 = "inactive": range.Cells(23, 5).Value2 = "inactive"
        Case 7
            valid = True: range.Cells(15, 3).Value2 = 0#: range.Cells(25, 2).Value2 = "inactive"
            range.Cells(25, 5).Value2 = "inactive"
        Case 8: row = 15: column = 2: range.Cells(7, 6).Value2 = 0#: range.Cells(15, 3).Value2 = 0#: range.Cells(row, column).Value2 = "bad"
        Case 9
            valid = True: range.Cells(23, 2).Value2 = 0#: range.Cells(23, 5).Value2 = 0#
            range.Cells(23, 3).Value2 = "inactive": range.Cells(23, 4).Value2 = "inactive"
        Case 10: row = 7: column = 6: range.Cells(row, column).Value2 = 1000# / factor
        Case 11: valid = True: range.Cells(7, 3).Value2 = 500# / factor
        Case 12: row = 23: column = 4: range.Cells(23, 2).Value2 = 16# / factor: range.Cells(row, column).Value2 = "bad"
        Case 13: row = 13: column = 4: range.Cells(13, 3).Value2 = 0#: range.Cells(row, column).Value2 = 1.5
    End Select
    Set reader = New CSectionInputSnapshotReader: Set context = reader.Capture(ThisWorkbook, Array("rngHollowRectangleGeometry"), units)
    Set table = context.Table("rngHollowRectangleGeometry")
    If Not valid Then reference = table.SourceReference(row, column)
    Set recipe = New CSectionRecipeHollowRectangle: Set definition = recipe.CreateDefinition(context)
    mesh.StepX = 40#: mesh.StepY = 30#: mesh.BoundarySubdivisions = 2: mesh.MaterialID = 1
    Set builder = New CSectionBuilder: Set model = builder.Build(definition, mesh)
    HollowInputCase = valid
    Exit Function
Rejected:
    HollowInputCase = (Not valid And Len(reference) > 0 And InStr(1, Err.Description, reference, vbBinaryCompare) > 0)
    If Not HollowInputCase Then mReport = mReport & "HOLLOW INPUT DETAIL: " & CStr(scenario) & "; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: сквозная сборка рецепта без старой geometry/layout/annotation.
Public Function DumpCircleRecipeMatrix() As String
    Dim config As Object, saved As Variant, result As String, counts As Variant, number As Variant, scenario As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, factor As Double, fixture As String
    Dim calculation As Long, events As Boolean, updating As Boolean, builder As CSectionBuilder, model As CSectionModel
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    modTestShapeArchitecture.ConfigureShapeFixture "Circle"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    counts = Array(2, 3, 12, 13)
    For Each number In counts
        For scenario = 0 To 6
            SetCircleInput "Rebar.Count", CLng(number)
            SetCircleInput "Rebar.Diameter2", IIf(scenario = 0 Or scenario = 2, 0#, 16# / factor)
            SetCircleInput "Rebar.Diameter3", IIf(scenario = 0 Or scenario = 1, 0#, 12# / factor)
            SetCircleInput "Rebar.Loc2row", IIf(scenario = 5 Or scenario = 6, "SideBySide", "Stacked")
            SetCircleInput "Rebar.Loc3row", IIf(scenario = 4 Or scenario = 6, "SideBySide", "Stacked")
            ThisWorkbook.Worksheets("Config").Calculate
            settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = BuildCircleRecipe(settings, units, builder)
            fixture = "CircleMatrix.n" & CStr(number) & ".rows" & CStr(scenario)
            result = result & modTestShapeArchitecture.PreparedShapeSnapshot(fixture, model, builder.Mesh, builder.Rebars)
        Next scenario
    Next number
    GoTo Finished
Failed:
    result = result & "FAIL: Circle recipe matrix; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    DumpCircleRecipeMatrix = result
End Function

Private Function BuildCircleRecipe(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, ByVal builder As CSectionBuilder) As CSectionModel
    Set BuildCircleRecipe = BuildRecipe("Circle", settings, units, builder)
End Function

Private Function BuildRecipe(ByVal typeID As String, ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, ByVal builder As CSectionBuilder) As CSectionModel
    Dim reader As CSectionInputSnapshotReader, context As CSectionRecipeContext
    Dim definition As CSectionDefinition, mesh As TSectionMeshOptions
    Dim registry As CSectionTypeRegistry
    Set registry = New CSectionTypeRegistry
    Set reader = New CSectionInputSnapshotReader: Set context = reader.Capture(ThisWorkbook, registry.RecipeTableNames(typeID), units)
    Set definition = registry.CreateDefinition(typeID, context)
    mesh = ReadSectionMeshOptions(settings, units)
    Set BuildRecipe = builder.Build(definition, mesh)
End Function

' ДЛЯ ТЕСТОВ: три формы RectSet проходят один рецепт и общий сборщик.
Public Function DumpRectSetRecipeGeometry() As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    modTestShapeArchitecture.ConfigureShapeFixture "Circle"
    For Each fixture In Array("Rectangle", "LSection", "TwoRectangles")
        modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set builder = New CSectionBuilder: Set model = BuildRecipe("RectSet", settings, units, builder)
        result = result & modTestShapeArchitecture.PreparedShapeSnapshot(CStr(fixture), model, builder.Mesh, builder.Rebars)
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: RectSet recipe geometry; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    DumpRectSetRecipeGeometry = result
End Function

' ДЛЯ ТЕСТОВ: тот же ввод расширенной матрицы, только новый pipeline.
Public Function DumpRectSetRecipeMatrix() As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant, scenario As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    For Each fixture In Array("Rectangle", "LSection", "TwoRectangles")
        For scenario = 0 To 5
            config.Formula = saved
            modTestShapeArchitecture.ConfigureShapeFixture "Circle"
            modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
            modTestShapeArchitecture.ConfigureRectSetScenario CStr(fixture), scenario
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = BuildRecipe("RectSet", settings, units, builder)
            result = result & modTestShapeArchitecture.PreparedShapeSnapshot("RectMatrix." & CStr(fixture) & ".case" & CStr(scenario), model, builder.Mesh, builder.Rebars)
        Next scenario
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: RectSet recipe matrix; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    DumpRectSetRecipeMatrix = result
End Function

' ДЛЯ ТЕСТОВ: точные пути общего builder на тех же 24 исходных fixtures.
Public Function DumpRoundedRecipeMatrix() As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant, scenario As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    For Each fixture In Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft")
        For scenario = 0 To 5
            config.Formula = saved
            modTestShapeArchitecture.ConfigureShapeFixture "Circle"
            modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
            modTestShapeArchitecture.ConfigureRoundedScenario scenario
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = BuildRecipe("RoundedRectangle", settings, units, builder)
            result = result & modTestShapeArchitecture.PreparedShapeSnapshot("RoundedMatrix." & CStr(fixture) & ".case" & CStr(scenario), model, builder.Mesh, builder.Rebars)
        Next scenario
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: Rounded recipe matrix; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    DumpRoundedRecipeMatrix = result
End Function

' ДЛЯ ТЕСТОВ: точные дуги и проекции для 16 исходных Hollow fixtures.
Public Function DumpHollowRecipeMatrix() As String
    DumpHollowRecipeMatrix = HollowRecipeMatrix(0, 7)
End Function

' ДЛЯ ТЕСТОВ: отсутствующие прямые участки и sharp frame тем же рецептом.
Public Function DumpHollowRecipeEdgeMatrix() As String
    DumpHollowRecipeEdgeMatrix = HollowRecipeMatrix(8, 10)
End Function

Private Function HollowRecipeMatrix(ByVal firstScenario As Long, ByVal lastScenario As Long) As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant, scenario As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    For Each fixture In Array("HollowCentered", "HollowOffset")
        For scenario = firstScenario To lastScenario
            config.Formula = saved
            modTestShapeArchitecture.ConfigureShapeFixture "Circle"
            modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
            modTestShapeArchitecture.ConfigureHollowScenario scenario
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = BuildRecipe("HollowRectangle", settings, units, builder)
            result = result & modTestShapeArchitecture.PreparedShapeSnapshot("HollowMatrix." & CStr(fixture) & ".case" & CStr(scenario), model, builder.Mesh, builder.Rebars)
        Next scenario
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: Hollow recipe matrix; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    HollowRecipeMatrix = result
End Function

' ДЛЯ ТЕСТОВ: сумма готовых подписей в полной матрице перенесенных рецептов.
Public Function RunRecipeAnnotationCountTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, fixture As Variant, scenario As Long, number As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean, factor As Double
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    modTestShapeArchitecture.ConfigureShapeFixture "Circle"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    For Each number In Array(2, 3, 12, 13)
        For scenario = 0 To 6
            SetCircleInput "Rebar.Count", CLng(number)
            SetCircleInput "Rebar.Diameter2", IIf(scenario = 0 Or scenario = 2, 0#, 16# / factor)
            SetCircleInput "Rebar.Diameter3", IIf(scenario = 0 Or scenario = 1, 0#, 12# / factor)
            SetCircleInput "Rebar.Loc2row", IIf(scenario = 5 Or scenario = 6, "SideBySide", "Stacked")
            SetCircleInput "Rebar.Loc3row", IIf(scenario = 4 Or scenario = 6, "SideBySide", "Stacked")
            ThisWorkbook.Worksheets("Config").Calculate
            settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = BuildRecipe("Circle", settings, units, builder)
            Check "annotations.total.Circle.n" & CStr(number) & ".case" & CStr(scenario), AnnotatedBarCount(model.Annotations) = model.RebarCount
        Next scenario
    Next number
    For Each fixture In Array("Rectangle", "LSection", "TwoRectangles")
        For scenario = 0 To 5
            config.Formula = saved
            modTestShapeArchitecture.ConfigureShapeFixture "Circle"
            modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
            modTestShapeArchitecture.ConfigureRectSetScenario CStr(fixture), scenario
            settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = BuildRecipe("RectSet", settings, units, builder)
            Check "annotations.total." & CStr(fixture) & ".case" & CStr(scenario), AnnotatedBarCount(model.Annotations) = model.RebarCount
        Next scenario
    Next fixture
    For Each fixture In Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft")
        For scenario = 0 To 5
            config.Formula = saved
            modTestShapeArchitecture.ConfigureShapeFixture "Circle"
            modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
            modTestShapeArchitecture.ConfigureRoundedScenario scenario
            settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = BuildRecipe("RoundedRectangle", settings, units, builder)
            Check "annotations.total." & CStr(fixture) & ".case" & CStr(scenario), AnnotatedBarCount(model.Annotations) = model.RebarCount
        Next scenario
    Next fixture
    For Each fixture In Array("HollowCentered", "HollowOffset")
        For scenario = 0 To 10
            config.Formula = saved
            modTestShapeArchitecture.ConfigureShapeFixture "Circle"
            modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
            modTestShapeArchitecture.ConfigureHollowScenario scenario
            settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = BuildRecipe("HollowRectangle", settings, units, builder)
            Check "annotations.total." & CStr(fixture) & ".case" & CStr(scenario), AnnotatedBarCount(model.Annotations) = model.RebarCount
        Next scenario
    Next fixture
    GoTo Finished
Failed:
    Check "annotations.total.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    RunRecipeAnnotationCountTests = mReport & "TOTAL RECIPE ANNOTATIONS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Private Sub TestAdjacentRectangles()
    Dim query As CSectionGeometryQuery, path As CGeometryPath, region As CConcreteRegion
    Set query = New CSectionGeometryQuery
    Set path = query.AdjacentRectangles("Pair", Array(Array(0#, 0#, 600#, 200#), Array(-50#, 200#, 300#, 400#)))
    Set region = query.RegionFromPath(path)
    Near "adjacent.area", region.Area, 240000#
    Check "adjacent.closed", path.IsClosed And region.LoopCount = 1
    Check "adjacent.lower", query.ContainsPoint(region, 500#, 100#)
    Check "adjacent.upper", query.ContainsPoint(region, -25#, 300#)
    Check "adjacent.concavity", Not query.ContainsPoint(region, 500#, 300#)
    Set path = query.AdjacentRectangles("Three", Array(Array(0#, 0#, 100#, 100#), Array(100#, 0#, 100#, 100#), Array(0#, 100#, 100#, 100#)))
    Set region = query.RegionFromPath(path)
    Near "adjacent.three.area", region.Area, 30000#
    Check "adjacent.three.concavity", Not query.ContainsPoint(region, 150#, 150#)
    Check "adjacent.overlap.rejected", InvalidAdjacentRejected(0)
    Check "adjacent.disjoint.rejected", InvalidAdjacentRejected(1)
    Check "adjacent.corner.rejected", InvalidAdjacentRejected(2)
End Sub

Private Function InvalidAdjacentRejected(ByVal scenario As Long) As Boolean
    On Error GoTo Rejected
    Dim query As CSectionGeometryQuery, path As CGeometryPath, rect As Variant
    Set query = New CSectionGeometryQuery
    Select Case scenario
        Case 0: rect = Array(50#, 50#, 100#, 100#)
        Case 1: rect = Array(200#, 0#, 100#, 100#)
        Case 2: rect = Array(100#, 100#, 100#, 100#)
    End Select
    Set path = query.AdjacentRectangles("Bad", Array(Array(0#, 0#, 100#, 100#), rect))
    Exit Function
Rejected:
    InvalidAdjacentRejected = (InStr(1, Err.Source, "CSectionGeometryQuery", vbBinaryCompare) > 0)
End Function

Private Sub SetCircleInput(ByVal key As String, ByVal value As Variant)
    Dim table As Object, i As Long
    Set table = ThisWorkbook.Names("rngCircleGeometry").RefersToRange
    For i = 1 To table.Rows.Count
        If CStr(table.Cells(i, 1).Value2) = key Then table.Cells(i, 2).Value2 = value: Exit Sub
    Next i
    Err.Raise vbObjectError + 5531, "modTestSectionRecipes", "Не найден ключ тестового ввода '" & key & "'."
End Sub

' ДЛЯ ТЕСТОВ: НДС/несущая/трещины для фактической модели нового Circle.
Public Function DumpCircleRecipeCalculations(Optional ByVal crackCode As String = "SP35") As String
    DumpCircleRecipeCalculations = RecipeCalculations("Circle", Array("Circle"), crackCode)
End Function

Public Function DumpRectSetRecipeCalculations(Optional ByVal crackCode As String = "SP35") As String
    DumpRectSetRecipeCalculations = RecipeCalculations("RectSet", Array("Rectangle", "LSection", "TwoRectangles"), crackCode)
End Function

Public Function DumpRoundedRecipeCalculations(Optional ByVal crackCode As String = "SP35") As String
    DumpRoundedRecipeCalculations = RecipeCalculations("RoundedRectangle", _
        Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft"), crackCode)
End Function

Public Function DumpRoundedRecipeArcCalculations(Optional ByVal crackCode As String = "SP35") As String
    DumpRoundedRecipeArcCalculations = RecipeCalculations("RoundedRectangle", _
        Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft"), crackCode, 1)
End Function

' ДЛЯ ТЕСТОВ: независимое суммирование N/Mx/My вне solver для новых дуг.
' Площадь стали замещает бетон; вычитается напряжение бетона в той же точке.
Public Function RunRoundedRecipeEquilibriumTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, fixture As Variant, events As Boolean, updating As Boolean
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    On Error GoTo Failed
    events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    modTestShapeArchitecture.ConfigureShapeFixture "Circle"
    For Each fixture In Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft")
        modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
        modTestShapeArchitecture.ConfigureRoundedScenario 1
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set builder = New CSectionBuilder: Set model = BuildRecipe("RoundedRectangle", settings, units, builder)
        CheckModelEquilibrium CStr(fixture), model, settings, units
    Next fixture
    GoTo Finished
Failed:
    Check "equilibrium.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.EnableEvents = events: Application.ScreenUpdating = updating
    RunRoundedRecipeEquilibriumTests = mReport & "TOTAL ROUNDED EQUILIBRIUM: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Private Sub CheckModelEquilibrium(ByVal fixture As String, ByVal model As CSectionModel, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim props As CSectionPropertiesCalculator, provider As CMaterialModelProvider, solver As CSectionSolver
    Dim purpose As Variant, loadIndex As Long, i As Long, concrete As Object, steel As Object
    Dim n As Double, mx As Double, my As Double, minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim sn As Double, smx As Double, smy As Double, x As Double, y As Double, strain As Double, force As Double, prefix As String
    Set props = New CSectionPropertiesCalculator: props.CalculateConcrete model
    props.CalculateProjection model, 1#, 0#, False, minX, maxX
    props.CalculateProjection model, 0#, 1#, False, minY, maxY
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    For Each purpose In Array(cpStrength, cpMcrc, cpCrackedNDS)
        Set concrete = provider.ConcreteMaterial(CLng(purpose)): Set steel = provider.SteelMaterial(CLng(purpose))
        For loadIndex = 0 To 3
            prefix = "equilibrium." & fixture & "." & CStr(purpose) & "." & CStr(loadIndex)
            n = -2# * props.Area: mx = n * props.CentroidY: my = n * props.CentroidX
            If loadIndex = 1 Or loadIndex = 3 Then mx = mx + props.Area * (maxY - minY) * 0.2
            If loadIndex = 2 Or loadIndex = 3 Then my = my - props.Area * (maxX - minX) * 0.15
            Set solver = New CSectionSolver: solver.ApplySettings settings, units
            solver.Solve model, concrete, steel, n, mx, my
            Check prefix & ".converged", solver.Converged
            If Not solver.Converged Then GoTo NextLoad
            sn = 0#: smx = 0#: smy = 0#
            For i = 1 To model.ConcreteCount
                x = model.ConcreteX(i): y = model.ConcreteY(i)
                strain = solver.Epsilon0 + solver.KappaX * y + solver.KappaY * x
                force = concrete.StressAtStrain(strain) * model.ConcreteArea(i)
                sn = sn + force: smx = smx + force * y: smy = smy + force * x
            Next i
            For i = 1 To model.RebarCount
                x = model.RebarX(i): y = model.RebarY(i)
                strain = solver.Epsilon0 + solver.KappaX * y + solver.KappaY * x
                force = (steel.StressAtStrain(strain) - concrete.StressAtStrain(strain)) * model.RebarArea(i)
                sn = sn + force: smx = smx + force * y: smy = smy + force * x
            Next i
            Check prefix & ".sum.N", Abs(sn - solver.Nint) <= 0.0000001 + Abs(sn) * 0.000000000001
            Check prefix & ".sum.Mx", Abs(smx - solver.Mxint) <= 0.0001 + Abs(smx) * 0.000000000001
            Check prefix & ".sum.My", Abs(smy - solver.Myint) <= 0.0001 + Abs(smy) * 0.000000000001
            ' Независимые суммы проверены выше в размерном допуске. Невязка
            ' обязана соответствовать опубликованному усилию, а не иной сумме,
            ' где вычитание больших близких чисел усиливает округление Double.
            Near prefix & ".residual.N", solver.Nint - n, solver.ResidualN
            Near prefix & ".residual.Mx", solver.Mxint - mx, solver.ResidualMx
            Near prefix & ".residual.My", solver.Myint - my, solver.ResidualMy
NextLoad:
        Next loadIndex
    Next purpose
End Sub

' ДЛЯ ТЕСТОВ: стандартные, дуговые, Ro=0 и смещенные проекции.
Public Function DumpHollowRecipeCalculations(Optional ByVal crackCode As String = "SP35", Optional ByVal scenario As Long = 0) As String
    DumpHollowRecipeCalculations = RecipeCalculations("HollowRectangle", Array("HollowCentered", "HollowOffset"), crackCode, 0, scenario)
End Function

' ДЛЯ ТЕСТОВ: суммирование сил отдельно от solver на всех четырех Hollow-сценариях.
Public Function RunHollowRecipeEquilibriumTests() As String
    RunHollowRecipeEquilibriumTests = HollowRecipeEquilibrium(0, 3)
End Function

' ДЛЯ ТЕСТОВ: исчезнувшие прямые грани, полный круг и острые углы.
Public Function RunHollowRecipeEdgeEquilibriumTests() As String
    RunHollowRecipeEdgeEquilibriumTests = HollowRecipeEquilibrium(8, 10)
End Function

Private Function HollowRecipeEquilibrium(ByVal firstScenario As Long, ByVal lastScenario As Long) As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, fixture As Variant, scenario As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    For Each fixture In Array("HollowCentered", "HollowOffset")
        For scenario = firstScenario To lastScenario
            config.Formula = saved
            modTestShapeArchitecture.ConfigureShapeFixture "Circle"
            modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
            modTestShapeArchitecture.ConfigureHollowScenario scenario
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = BuildRecipe("HollowRectangle", settings, units, builder)
            CheckModelEquilibrium CStr(fixture) & ".case" & CStr(scenario), model, settings, units
        Next scenario
    Next fixture
    GoTo Finished
Failed:
    Check "hollow.equilibrium.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    HollowRecipeEquilibrium = mReport & "TOTAL HOLLOW EQUILIBRIUM: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Private Function RecipeCalculations(ByVal typeID As String, ByVal fixtures As Variant, ByVal crackCode As String, _
        Optional ByVal roundedScenario As Long = 0, Optional ByVal hollowScenario As Long = 0) As String
    Dim config As Object, saved As Variant, result As String, settings As CSystemSettingsReader, units As CUnitSystem
    Dim profiles As Object, i As Long, key As String, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean, fixture As Variant
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    Set profiles = ThisWorkbook.Names("rngCalculationProfiles").RefersToRange
    For i = 2 To profiles.Rows.Count
        key = CStr(profiles.Cells(i, 2).Value2)
        Select Case key
            Case "Calculation.Strength.DirectState", "Calculation.Strength.Capacity", "Calculation.Crack.Width"
                profiles.Cells(i, 3).Value2 = "Yes"
            Case "Calculation.Stability.Enabled", "Calculation.Crack.Longitudinal"
                profiles.Cells(i, 3).Value2 = "No"
        End Select
    Next i
    SetSystemInput "SLS.Crack.Code", crackCode
    modTestShapeArchitecture.ConfigureShapeFixture "Circle"
    For Each fixture In fixtures
        modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
        If roundedScenario > 0 Then modTestShapeArchitecture.ConfigureRoundedScenario roundedScenario
        If hollowScenario > 0 Then modTestShapeArchitecture.ConfigureHollowScenario hollowScenario
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set builder = New CSectionBuilder: Set model = BuildRecipe(typeID, settings, units, builder)
        Dim label As String
        label = CStr(fixture): If roundedScenario > 0 Then label = label & ".Arc"
        If hollowScenario > 0 Then label = label & ".case" & CStr(hollowScenario)
        result = result & modTestShapeArchitecture.PreparedCalculationSnapshot(label, model)
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: recipe calculations; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    RecipeCalculations = result
End Function

' ДЛЯ ТЕСТОВ: настоящий Excel PNG и структурный снимок схемы.
Public Function DumpCircleRecipePreview(ByVal outputDirectory As String) As String
    DumpCircleRecipePreview = RecipePreview("Circle", Array("Circle"), outputDirectory)
End Function

Public Function DumpRectSetRecipePreview(ByVal outputDirectory As String) As String
    DumpRectSetRecipePreview = RecipePreview("RectSet", Array("Rectangle", "LSection", "TwoRectangles"), outputDirectory)
End Function

Public Function DumpRoundedRecipePreview(ByVal outputDirectory As String) As String
    DumpRoundedRecipePreview = RecipePreview("RoundedRectangle", _
        Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft"), outputDirectory)
End Function

Public Function DumpRoundedRecipeArcPreview(ByVal outputDirectory As String) As String
    DumpRoundedRecipeArcPreview = RecipePreview("RoundedRectangle", _
        Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft"), outputDirectory, 1)
End Function

' ДЛЯ ТЕСТОВ: настоящий Excel preview стандартных/измененных Hollow-сценариев.
Public Function DumpHollowRecipePreview(ByVal outputDirectory As String, Optional ByVal scenario As Long = 0) As String
    DumpHollowRecipePreview = RecipePreview("HollowRectangle", Array("HollowCentered", "HollowOffset"), outputDirectory, 0, scenario)
End Function

Private Function RecipePreview(ByVal typeID As String, ByVal fixtures As Variant, ByVal outputDirectory As String, _
        Optional ByVal roundedScenario As Long = 0, Optional ByVal hollowScenario As Long = 0) As String
    Dim config As Object, saved As Variant, result As String, settings As CSystemSettingsReader, units As CUnitSystem
    Dim builder As CSectionBuilder, model As CSectionModel, calculation As Long, events As Boolean, updating As Boolean, fixture As Variant
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = True
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    modTestShapeArchitecture.ConfigureShapeFixture "Circle"
    For Each fixture In fixtures
        modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
        If roundedScenario > 0 Then modTestShapeArchitecture.ConfigureRoundedScenario roundedScenario
        If hollowScenario > 0 Then modTestShapeArchitecture.ConfigureHollowScenario hollowScenario
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set builder = New CSectionBuilder: Set model = BuildRecipe(typeID, settings, units, builder)
        Dim label As String
        label = CStr(fixture): If roundedScenario > 0 Then label = label & ".Arc"
        If hollowScenario > 0 Then label = label & ".case" & CStr(hollowScenario)
        result = result & modTestShapeArchitecture.PreparedPreviewSnapshot(label, model, outputDirectory)
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: recipe preview; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    RecipePreview = result
End Function

Private Sub SetSystemInput(ByVal key As String, ByVal value As Variant)
    Dim table As Object, i As Long
    Set table = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    For i = 1 To table.Rows.Count
        If CStr(table.Cells(i, 1).Value2) = key Then table.Cells(i, 2).Value2 = value: Exit Sub
    Next i
    Err.Raise vbObjectError + 5531, "modTestSectionRecipes", "Не найден системный ключ '" & key & "'."
End Sub
