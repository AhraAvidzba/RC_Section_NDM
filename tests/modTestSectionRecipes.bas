Attribute VB_Name = "modTestSectionRecipes"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: чтение рецептов, общие правила армирования, проекции и сохранение точных геометрических и расчётных снимков.
' ==========================================================================

Private mPassed As Long, mFailed As Long, mReport As String

' -------------------- ДЛЯ ТЕСТОВ --------------------
' обычное чтение входных данных использует только таблицы выбранного рецепта.
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
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    Set registry = New CSectionTypeRegistry
    For Each fixture In Array("Circle", "Rectangle", "RoundedMixed", "HollowOffset")
        TestFixtureRestoreInputs config, saved
        modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
        modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        typeID = registry.SelectedRecipeType(settings): names = registry.RecipeTableNames(typeID): rootName = CStr(names(0))
        For Each name In Array("Circle.Diameter", "Rectangles.B1", "PointedSection.B", "HollowRectangle.B")
            TestCheck "workbook.reader.noShapeKey." & CStr(fixture) & "." & CStr(name), Not settings.HasKey(CStr(name))
        Next name
        Set changedName = ThisWorkbook.Names(IIf(rootName = "rngCircleGeometry", "rngHollowRectangleGeometry", "rngCircleGeometry"))
        originalName = changedName.RefersTo: changedName.RefersTo = "=#REF!"
        Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
        TestCheck "workbook.inactiveName." & CStr(fixture), model.RebarCount > 0 And model.ConcreteCount > 0
        changedName.RefersTo = originalName: Set changedName = Nothing
        Set rootRange = ThisWorkbook.Names(rootName).RefersToRange
        Set reader = New CSectionInputSnapshotReader: Set context = reader.Capture(ThisWorkbook, names, units)
        Set table = context.Table(rootName): initialCell = table.CellValue(1, 1)
        rootRange.Cells(1, 1).Value2 = "CHANGED_AFTER_CAPTURE"
        TestCheck "workbook.table.immutable." & CStr(fixture), CStr(table.CellValue(1, 1)) = CStr(initialCell)
        Set definition = registry.CreateDefinition(typeID, context)
        TestCheck "workbook.definition.sealed." & CStr(fixture), definition.IsSealed
        rootRange.ClearContents
        mesh = ReadSectionMeshOptions(settings, units)
        Set rebuilt = builder.Build(definition, mesh)
        TestCheck "workbook.definition.meshCount." & CStr(fixture), rebuilt.ConcreteCount = model.ConcreteCount
        TestCheck "workbook.definition.barCount." & CStr(fixture), rebuilt.RebarCount = model.RebarCount
        For i = 1 To model.RebarCount
            TestNear "workbook.definition.x." & CStr(fixture) & "." & CStr(i), rebuilt.RebarX(i), model.RebarX(i)
            TestNear "workbook.definition.y." & CStr(fixture) & "." & CStr(i), rebuilt.RebarY(i), model.RebarY(i)
        Next i
        TestFixtureRestoreInputs config, saved
        Set changedName = ThisWorkbook.Names(rootName): originalName = changedName.RefersTo: changedName.RefersTo = "=#REF!"
        number = 0: reason = ""
        On Error Resume Next
        Set context = reader.Capture(ThisWorkbook, names, units)
        number = Err.Number: reason = Err.Description: Err.Clear
        On Error GoTo Failed
        changedName.RefersTo = originalName: Set changedName = Nothing
        TestCheck "workbook.activeName.reject." & CStr(fixture), number <> 0 And InStr(1, reason, rootName, vbBinaryCompare) > 0
    Next fixture
    TestFixtureRestoreInputs config, saved
    ThisWorkbook.Names("rngUnitSettings").RefersToRange.Cells(2, 2).Value2 = "mm"
    modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set rootRange = ThisWorkbook.Names("rngCircleGeometry").RefersToRange
    Set context = reader.Capture(ThisWorkbook, Array("rngCircleGeometry"), units): Set table = context.Table("rngCircleGeometry")
    Set diameterCell = TestFixtureCircleFieldCell(ThisWorkbook, "Circle.Diameter")
    Set coverCell = TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.AxisDistance")
    Application.Calculation = xlCalculationManual
    diameterCell.Value2 = 300#: coverCell.Formula = "=" & diameterCell.Address & "/4": coverCell.Calculate
    diameterCell.Value2 = 320#
    TestNear "workbook.manual.stale.fixture", CDbl(coverCell.Value2), 75#
    Set context = reader.Capture(ThisWorkbook, Array("rngCircleGeometry"), units): Set table = context.Table("rngCircleGeometry")
    TestNear "workbook.manual.capture.noHiddenCalculate", context.Reader("rngCircleGeometry").Length("Main", "Circle", "as"), 75#
    TestNear "workbook.manual.capture.cacheUnchanged", CDbl(coverCell.Value2), 75#
    TestCheck "workbook.manual.capture.modeUnchanged", Application.Calculation = xlCalculationManual
    Application.Calculate ' Низкоуровневый API снимает готовый кэш; штатный Form/Calculate актуализирует весь граф до чтения.
    Set context = reader.Capture(ThisWorkbook, Array("rngCircleGeometry"), units)
    TestNear "workbook.manual.capture.afterExplicitCalculate", context.Reader("rngCircleGeometry").Length("Main", "Circle", "as"), 80#
    names = registry.RecipeTableNames("Circle"): names(0) = "CHANGED_COPY"
    names = registry.RecipeTableNames("Circle")
    TestCheck "workbook.registry.metadata.copy", CStr(names(0)) = "rngCircleGeometry"
    registry.RegisterRecipe "FixedMetadata", New CSectionRecipeCircle, Array()
    i = 0: names = registry.RecipeTableNames("FixedMetadata")
    For Each name In names: i = i + 1: Next name
    TestCheck "workbook.registry.emptyMetadata", i = 0
    TestSetSystemInput "Geometry.Type", "FixedMetadata"
    settings.LoadFromWorkbook ThisWorkbook
    TestCheck "workbook.registry.dynamicChoice", registry.SelectedRecipeType(settings) = "FixedMetadata"
    TestInputTableCaptionSearch
    GoTo Finished
Failed:
    TestCheck "workbook.recipe.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    On Error Resume Next
    If Not changedName Is Nothing Then changedName.RefersTo = originalName
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    On Error GoTo 0
    RunWorkbookRecipeInputTests = mReport & "TOTAL WORKBOOK RECIPE INPUT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' поиск заголовка не потребляет неактивное ошибочное значение.
' При прямом чтении этого же поля ошибка остаётся адресной и не маскируется.
Private Sub TestInputTableCaptionSearch()
    Dim values(1 To 3, 1 To 2) As Variant, addresses(1 To 3, 1 To 2) As String
    Dim table As CSectionInputTable, row As Long, column As Long, value As Variant, number As Long, reason As String
    values(1, 1) = "Header": values(1, 2) = "as"
    values(2, 1) = "Disabled": values(2, 2) = CVErr(xlErrValue)
    values(3, 1) = "Active": values(3, 2) = 50#
    For row = 1 To 3
        For column = 1 To 2: addresses(row, column) = "Moved!" & IIf(column = 1, "A", "B") & CStr(row): Next column
    Next row
    Set table = New CSectionInputTable: table.Initialize "CaptionFixture", values, addresses
    TestCheck "inputTable.caption.ignoresInactiveError", table.RowByKey("as", 2) = 1
    On Error Resume Next
    value = table.CellValue(2, 2): number = Err.Number: reason = Err.Description
    Err.Clear: On Error GoTo 0
    TestCheck "inputTable.caption.activeReadRejects", number <> 0 And InStr(1, reason, "Moved!B2", vbBinaryCompare) > 0
    TestNear "inputTable.caption.otherValuePreserved", table.Number(3, 2), 50#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' шум координаты касания не меняет дробление соседней ячейки.
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
                TestNear "mesh.touch.full.area." & CStr(delta), mesh.FiberArea(i), 800#
                found = True
            End If
        Next i
        TestCheck "mesh.touch.full.cell." & CStr(delta), found
        If reference Is Nothing Then
            Set reference = mesh
        Else
            TestCheck "mesh.touch.count." & CStr(delta), mesh.FiberCount = reference.FiberCount
            If mesh.FiberCount = reference.FiberCount Then
                For i = 1 To mesh.FiberCount
                    TestNear "mesh.touch.x." & CStr(i), mesh.FiberX(i), reference.FiberX(i)
                    TestNear "mesh.touch.y." & CStr(i), mesh.FiberY(i), reference.FiberY(i)
                    TestNear "mesh.touch.area." & CStr(i), mesh.FiberArea(i), reference.FiberArea(i)
                Next i
            End If
        End If
    Next delta
    GoTo Finished
Failed:
    TestCheck "mesh.touch.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRecipeMeshBoundaryTests = mReport & "TOTAL RECIPE MESH BOUNDARY: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' общие правила без таблиц типов или старых раскладчиков.
Public Function RunRecipeRuleTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim facade As CSectionDefinitionBuilder, options As CRebarRuleOptions, model As CSectionModel
    Dim builder As CSectionBuilder, path As CGeometryPath, mode As Long, scenario As Long, expected As Variant, i As Long
    For mode = IncludeBoth To IncludeEnd
        Set facade = TestRuleFacade: Set options = New CRebarRuleOptions
        options.EndMode = mode: options.StartSetback = 10#: options.EndSetback = 20#
        facade.RebarByCount "Row.Main", "Row", 3, 10#, options
        options.StartSetback = 999# ' Параметры арматуры уже скопированы в описание сечения и не зависят от исходного объекта.
        Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
        Select Case mode
            Case IncludeBoth: expected = Array(10#, 45#, 80#)
            Case ExcludeBoth: expected = Array(27.5, 45#, 62.5)
            Case IncludeStart: expected = Array(10#, 10# + 70# / 3#, 10# + 140# / 3#)
            Case IncludeEnd: expected = Array(10# + 70# / 3#, 10# + 140# / 3#, 80#)
        End Select
        TestCheck "ends.count." & CStr(mode), model.RebarCount = 3
        For i = 1 To 3: TestNear "ends.station." & CStr(mode) & "." & CStr(i), model.RebarX(i), CDbl(expected(i - 1)): Next i
    Next mode
    Set facade = TestRuleFacade: Set options = New CRebarRuleOptions
    options.StartSetback = 10#: options.EndSetback = 20#
    facade.RebarByCount "One", "Row", 1, 10#, options
    Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
    TestNear "one.midpoint", model.RebarX(1), 45#
    Set facade = TestRuleFacade: Set options = New CRebarRuleOptions
    options.EndMode = IncludeStart
    Set path = facade.CirclePath("Closed", 0#, 0#, 100#)
    facade.RebarByCount "Closed.Row", "Closed", 4, 10#, options
    Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
    TestCheck "closed.count", model.RebarCount = 4
    TestNear "closed.start.x", model.RebarX(1), 100#: TestNear "closed.start.y", model.RebarY(1), 0#
    TestNear "closed.last.x", model.RebarX(4), 0#: TestNear "closed.last.y", model.RebarY(4), -100#
    Set facade = TestRuleFacade: Set options = New CRebarRuleOptions
    options.StartSetback = 10#: options.EndSetback = 20#
    facade.RebarByMaxStep "Step", "Row", 30#, 10#, options
    Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
    TestCheck "maxstep.count", model.RebarCount = 4
    TestNear "maxstep.first", model.RebarX(1), 10#: TestNear "maxstep.last", model.RebarX(4), 80#
    TestNear "maxstep.actual", model.RebarX(2) - model.RebarX(1), 70# / 3#
    Set facade = TestRuleFacade: Set options = New CRebarRuleOptions
    options.EndMode = ExcludeBoth
    facade.RebarByFixedStep "Fixed", "Row", 30#, 10#, options
    Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
    TestCheck "fixed.count", model.RebarCount = 4
    TestNear "fixed.first", model.RebarX(1), 5#: TestNear "fixed.last", model.RebarX(4), 95#
    For i = 2 To 4: TestNear "fixed.step." & CStr(i), model.RebarX(i) - model.RebarX(i - 1), 30#: Next i
    Set facade = TestRuleFacade
    facade.RebarAtStations "Explicit", "Row", Array(10#, 30#, 90#), 10#
    Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
    TestCheck "explicit.count", model.RebarCount = 3: TestNear "explicit.last", model.RebarX(3), 90#
    Set facade = TestRuleFacade: Set options = New CRebarRuleOptions
    options.SourcePrefix = "P"
    facade.RebarByCount "Base", "Row", 3, 10#, options
    Set options = New CRebarRuleOptions
    options.NormalOffset = 10#: options.EverySecond = True: options.SourceSuffix = "_Extra"
    facade.RebarDependent "Extra", "Base", 10#, options
    Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
    TestCheck "dependent.count", model.RebarCount = 5
    TestNear "dependent.first.x", model.RebarX(4), 0#: TestNear "dependent.last.x", model.RebarX(5), 100#
    TestNear "dependent.normal", model.RebarY(5), 10#
    TestCheck "dependent.source", model.RebarSourceName(5) = "P3_Extra"
    Set facade = TestRuleFacade: Set options = New CRebarRuleOptions
    options.SourcePrefix = "P": facade.RebarByCount "Base", "Row", 3, 10#, options
    Set path = facade.LinePath("Target", 0#, 100#, 100#, 100#)
    Set options = New CRebarRuleOptions
    options.DirectionY = 1#: options.SkipMissingProjection = True
    options.StartSetback = 10#: options.EndSetback = 10#: options.SourceSuffix = "_Projected"
    facade.RebarProjected "Projection", "Target", "Base", 10#, options
    Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
    TestCheck "projection.count", model.RebarCount = 4: TestCheck "projection.skips", builder.Rules.SkippedProjections = 2
    TestNear "projection.x", model.RebarX(4), 50#: TestNear "projection.y", model.RebarY(4), 100#
    TestCheck "projection.source", model.RebarSourceName(4) = "P2_Projected"
    Set facade = TestRuleFacade ' Первое пересечение исключено setback; второе на том же пути допустимо.
    facade.RebarByCount "Base", "Row", 1, 10#
    Set path = New CGeometryPath: path.Initialize "TwoHits"
    path.AddLine "First", 0#, 50#, 100#, 50#
    path.AddLine "Side", 100#, 50#, 100#, 150#
    path.AddLine "Second", 100#, 150#, 0#, 150#
    path.Seal: facade.AddPath path
    Set options = New CRebarRuleOptions
    options.DirectionY = 1#: options.StartSetback = 110#: options.SourceSuffix = "_Hit"
    facade.RebarProjected "SecondHit", "TwoHits", "Base", 10#, options
    Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
    TestCheck "projection.secondhit.count", model.RebarCount = 2
    TestNear "projection.secondhit.x", model.RebarX(2), 50#: TestNear "projection.secondhit.y", model.RebarY(2), 150#
    Set facade = TestRuleFacade
    Set path = facade.LinePath("Target", 0#, 100#, 100#, 100#)
    facade.PathGroup "Both", Array("Row", "Target")
    facade.RebarByCount "BothRows", "Both", 2, 10#
    Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
    TestCheck "group.perpath", model.RebarCount = 4: TestNear "group.secondpath", model.RebarY(3), 100#
    Dim query As CSectionGeometryQuery, region As CConcreteRegion
    Set query = New CSectionGeometryQuery: Set region = query.CircleRegion(10#, -20#, 100#)
    TestNear "distance.center", query.BoundaryDistance(region, 10#, -20#), 100#
    TestNear "distance.inside", query.BoundaryDistance(region, 90#, -20#), 20#
    TestNear "distance.outside", query.BoundaryDistance(region, 150#, -20#), 40#
    For scenario = 0 To 7: TestCheck "rule.invalid." & CStr(scenario), TestInvalidRuleRejected(scenario): Next scenario
    TestPublicationAndCoincidence
    TestLocalRowBoundaryAnchor
    TestLayoutIndex
    TestAdjacentRectangles
    TestAnnotationMeasurements
    TestRuleJointTangents
    TestProjectionNormalsAndOrdinals
    TestMinimumRebarArea
    GoTo Finished
Failed:
    TestCheck "rules.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRecipeRuleTests = mReport & "TOTAL RECIPE RULES: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет проекцию по нормали дуги и нумерацию успешно попавших стержней отдельно от номеров исходного ряда.
Private Sub TestProjectionNormalsAndOrdinals()
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim builder As CSectionBuilder, model As CSectionModel, mode As Long, i As Long
    Set facade = TestRuleFacade
    Set path = facade.ArcPath("SourceCurve", 0#, 0#, 100#, 0#, GEOM_PI / 2#)
    Set options = New CRebarRuleOptions: options.AnnotationGroup = "CurveBars"
    facade.RebarByCount "CurveMain", path.ID, 3, 10#, options
    Set path = facade.ArcPath("TargetCurve", 0#, 0#, 200#, 0#, GEOM_PI / 2#)
    Set options = New CRebarRuleOptions: options.SourceNormalSign = -1#: options.NewPublicationRoot = True
    facade.RebarProjected "Radial", path.ID, "CurveMain", 10#, options
    options.SourceNormalSign = 1#
    facade.RebarRowExtremeLabel "CurveBars", "CurveMain", 0#, 1#, 0#, 1#
    Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
    TestCheck "projection.normal.count", model.RebarCount = 6
    For i = 1 To 3
        TestNear "projection.normal.x." & CStr(i), model.RebarX(i + 3), 2# * model.RebarX(i)
        TestNear "projection.normal.y." & CStr(i), model.RebarY(i + 3), 2# * model.RebarY(i)
    Next i
    TestNear "extreme.anchor.start", model.Annotations.StartX(1), 0#
    TestNear "extreme.anchor.end", model.Annotations.EndX(1), 100#
    TestNear "extreme.anchor.y", model.Annotations.StartY(1), 100#
    For mode = 0 To 1
        Set facade = TestRuleFacade
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
        Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
        TestCheck "projection.ordinal.count." & CStr(mode), model.RebarCount = 7 And builder.Rules.SkippedProjections = 2
        TestNear "projection.publication.first." & CStr(mode), model.RebarY(4), 0#
        If mode = 1 Then
            TestNear "projection.ordinal.success.x", model.RebarX(6), 100# / 3#
            TestNear "projection.ordinal.success.y", model.RebarY(6), 110#
        Else
            TestNear "projection.ordinal.source.x", model.RebarX(7), 200# / 3#
            TestNear "projection.ordinal.source.y", model.RebarY(7), 110#
        End If
    Next mode
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет выбор входящей или выходящей касательной в углу и направление смещения связанного стержня.
Private Sub TestRuleJointTangents()
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim builder As CSectionBuilder, model As CSectionModel, incoming As Long
    For incoming = 0 To 1
        Set facade = TestRuleFacade
        Set path = New CGeometryPath: path.Initialize "Joint"
        path.AddLine "Incoming", 0#, 0#, 100#, 0#
        path.AddLine "Outgoing", 100#, 0#, 100#, 100#
        path.Seal: facade.AddPath path
        Set options = New CRebarRuleOptions: options.IncomingTangentAtJoint = (incoming = 1)
        facade.RebarAtStations "JointMain", path.ID, Array(100#), 10#, options
        options.IncomingTangentAtJoint = Not options.IncomingTangentAtJoint
        Set options = New CRebarRuleOptions: options.NormalOffset = 10#
        facade.RebarDependent "JointExtra", "JointMain", 10#, options
        Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
        If incoming = 1 Then
            TestNear "joint.rule.incoming.x", model.RebarX(2), 100#: TestNear "joint.rule.incoming.y", model.RebarY(2), 10#
        Else
            TestNear "joint.rule.default.x", model.RebarX(2), 90#: TestNear "joint.rule.default.y", model.RebarY(2), 0#
        End If
    Next incoming
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет привязку подписи к фактическим границам ряда и округление только отображаемого размера без изменения координат.
Private Sub TestAnnotationMeasurements()
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim builder As CSectionBuilder, model As CSectionModel, annotations As CSectionAnnotations
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Set facade = TestRuleFacade
    Set path = facade.ArcPath("Curve", 0#, 0#, 100#, 0#, GEOM_PI / 2#)
    Set options = New CRebarRuleOptions: options.AnnotationGroup = "CurveBars": options.EndMode = ExcludeBoth
    facade.RebarByCount "CurveMain", path.ID, 3, 10#, options
    facade.RebarRowExtentLabel "CurveBars", "CurveMain", -1#, 0#, True
    facade.HorizontalDistance "REF", "Reference", 0#, 123.6, 200#, 1#, "", 1#
    facade.HorizontalDistance "NEG", "Negative", 0#, -123.6, 200#, 1#, "", 1#
    facade.HorizontalDistance "EXACT", "Exact", 0#, 123.6, 200#
    Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
    TestCheck "measurements.rule.bounds", builder.Rules.RuleBounds("CurveMain", minX, maxX, minY, maxY)
    TestNear "measurements.min.x", minX, 100# * Cos(3# * GEOM_PI / 8#)
    TestNear "measurements.max.x", maxX, 100# * Cos(GEOM_PI / 8#)
    TestNear "measurements.min.y", minY, 100# * Sin(GEOM_PI / 8#)
    TestNear "measurements.max.y", maxY, 100# * Sin(3# * GEOM_PI / 8#)
    TestNear "measurements.anchor.x", model.Annotations.StartX(4), minX
    TestNear "measurements.anchor.start", model.Annotations.StartY(4), maxY
    TestNear "measurements.anchor.end", model.Annotations.EndY(4), minY
    Set annotations = model.Annotations
    TestNear "measurements.round.positive", annotations.Value(1), 124#
    TestNear "measurements.round.negative", annotations.Value(2), -124#
    TestNear "measurements.round.exact", annotations.Value(3), 123.6
    TestNear "measurements.coordinates.unchanged", annotations.EndX(1), 123.6
    TestCheck "measurements.count", TestAnnotatedBarCount(annotations) = 3
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет порядок публикации основных и связанных стержней, объединение совпадений и единственного владельца подписи.
Private Sub TestPublicationAndCoincidence()
    Dim facade As CSectionDefinitionBuilder, options As CRebarRuleOptions, builder As CSectionBuilder
    Dim model As CSectionModel, expected As Variant, i As Long, memberships As Variant
    Set facade = TestRuleFacade: Set options = New CRebarRuleOptions
    options.PositionOrder = FromPathMiddle: options.Publication = PublishPerSourcePosition
    options.StartSetback = 10#: options.EndSetback = 20#: options.AnnotationGroup = "Base"
    facade.RebarByCount "Base", "Row", 5, 10#, options
    Set options = New CRebarRuleOptions
    options.Publication = PublishPerSourcePosition: options.EverySecond = True: options.NormalOffset = 10#
    facade.RebarDependent "Extra", "Base", 10#, options
    facade.RebarRowLabel "Base", "Base", 0#, -1#
    Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
    expected = Array(45#, 45#, 62.5, 27.5, 80#, 80#, 10#, 10#)
    TestCheck "middle.count", model.RebarCount = 8
    For i = 1 To 8: TestNear "middle.order." & CStr(i), model.RebarX(i), CDbl(expected(i - 1)): Next i
    TestNear "middle.first.extra", model.RebarY(2), 10#
    TestNear "middle.anchor.first", model.Annotations.StartX(1), 10#
    TestNear "middle.anchor.last", model.Annotations.EndX(1), 80#
    Set facade = TestRuleFacade: Set options = New CRebarRuleOptions
    options.AnnotationGroup = "First": facade.RebarByCount "First", "Row", 2, 10#, options
    Set options = New CRebarRuleOptions
    options.AnnotationGroup = "Second": options.Coincidence = MergeIdentical
    facade.RebarByCount "Second", "Row", 2, 10#, options
    facade.RebarRowLabel "First", "First", 0#, -1#
    facade.RebarRowLabel "Second", "Second", 0#, 1#
    Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
    TestCheck "coincidence.physical.count", model.RebarCount = 2 And builder.Rules.CreatedPositions = 2
    TestCheck "coincidence.requested", builder.Rules.RequestedPositions = 4 And builder.Rules.MergedPositions = 2
    memberships = builder.Rebars.BarMemberships(1)
    TestCheck "coincidence.memberships", UBound(memberships) = 1 And memberships(0) = "First" And memberships(1) = "Second"
    TestCheck "coincidence.annotation.owner", builder.Rebars.BarAnnotationGroupName(1) = "First"
    TestCheck "coincidence.labels", model.Annotations.Count = 1 And model.Annotations.Text(1) = "2" & ChrW$(216) & "10"
    TestCheck "coincidence.labels.total", TestAnnotatedBarCount(model.Annotations) = model.RebarCount
    For i = 0 To 3: TestCheck "coincidence.invalid." & CStr(i), TestInvalidCoincidenceRejected(i): Next i
    Set facade = TestRuleFacade
    facade.RebarByCount "First", "Row", 2, 10#
    Set options = New CRebarRuleOptions: options.Coincidence = KeepPublishedOwner
    facade.RebarByCount "Second", "Row", 2, 20#, options
    Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
    TestCheck "coincidence.explicit.owner", model.RebarCount = 2 And builder.Rules.Conflicts = 2
    TestNear "coincidence.owner.diameter", model.RebarDiameter(1), 10#
    TestAnnotationOwnership
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Одна общая позиция, полностью совпавшая третья группа и разные диаметры
' не должны повторно увеличивать количество стержней в подписях.
Private Sub TestAnnotationOwnership()
    Dim facade As CSectionDefinitionBuilder, options As CRebarRuleOptions, builder As CSectionBuilder
    Dim model As CSectionModel, path As CGeometryPath
    Dim annotations As CSectionAnnotations, labels As CRebarGroupAnnotationBuilder
    Set facade = TestRuleFacade: Set options = New CRebarRuleOptions
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
    Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
    TestCheck "annotations.partial.physical", model.RebarCount = 5 And builder.Rules.MergedPositions = 3
    TestCheck "annotations.partial.groups", model.Annotations.Count = 2
    TestCheck "annotations.partial.first", model.Annotations.Text(1) = "2" & ChrW$(216) & "10 + 1" & ChrW$(216) & "12"
    TestCheck "annotations.partial.second", model.Annotations.Text(2) = "2" & ChrW$(216) & "20"
    TestCheck "annotations.partial.total", TestAnnotatedBarCount(model.Annotations) = model.RebarCount
    Dim anchors(1 To 4) As TSectionAnnotation
    anchors(1).Name = "First": anchors(1).EndX = 100#: anchors(1).NormalY = -1#
    anchors(2).Name = "Second": anchors(2).StartX = 100#: anchors(2).EndX = 200#: anchors(2).NormalY = 1#
    anchors(3).Name = "Third": anchors(3).StartX = 100#: anchors(3).EndX = 150#: anchors(3).NormalY = 1#
    anchors(4).Name = "first": anchors(4).StartX = 300#: anchors(4).EndX = 400#: anchors(4).NormalY = -1#
    Set annotations = New CSectionAnnotations: Set labels = New CRebarGroupAnnotationBuilder
    labels.AddRebarLabels annotations, builder.Rebars, anchors, 4
    TestCheck "annotations.repeated.anchor", annotations.Count = 2
    TestCheck "annotations.repeated.total", TestAnnotatedBarCount(annotations) = model.RebarCount
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимо подсчитывает стержни по готовому тексту подписи без чтения раскладки.
Private Function TestAnnotatedBarCount(ByVal annotations As CSectionAnnotations) As Long
    Dim i As Long, part As Variant, position As Long
    For i = 1 To annotations.Count
        If annotations.AnnotationType(i) = "REBAR_ANNOTATION" Or annotations.AnnotationType(i) = "REBAR_LOCAL_ANNOTATION" Then
            For Each part In Split(annotations.Text(i), "+")
                position = InStr(1, CStr(part), ChrW$(216), vbBinaryCompare)
                If position = 0 Then Err.Raise vbObjectError + 5532, "modTestSectionRecipes", "Неверный формат групповой подписи: " & CStr(part)
                TestAnnotatedBarCount = TestAnnotatedBarCount + CLng(Trim$(Left$(CStr(part), position - 1)))
            Next part
        End If
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет отказ при запрещённом совпадении, разных диаметрах или стали и непредставимом числе связанных позиций.
Private Function TestInvalidCoincidenceRejected(ByVal scenario As Long) As Boolean
    On Error GoTo Rejected
    Dim facade As CSectionDefinitionBuilder, options As CRebarRuleOptions, builder As CSectionBuilder, model As CSectionModel
    Set facade = TestRuleFacade
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
    Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
    Exit Function
Rejected:
    TestInvalidCoincidenceRejected = (Err.Number = vbObjectError + 5530)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимый последовательный поиск проверяет ускоренный поиск совпавших
' стержней: на границах условных ячеек, при отрицательных и удалённых координатах,
' а также после очистки раскладки.
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
            TestCheck "index.oracle." & CStr(i) & "." & CStr(delta), layout.FindBarAt(queryX, y) = expected
        Next delta
    Next i
    layout.Clear
    TestCheck "index.clear", layout.Count = 0 And Not layout.HasBarAt(1000000#, 500000#)
    layout.AddBar "I1", 1000000#, 500000#, 10#, 0#, "Rebar"
    TestCheck "index.reuse", layout.FindBarAt(1000000#, 500000#) = 1
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт большой бетонный круг и прямой путь 100 мм для независимой проверки правил армирования.
Private Function TestRuleFacade() As CSectionDefinitionBuilder
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "RuleFixture"
    Set path = facade.CirclePath("Outer", 0#, 0#, 1000#): facade.Outer "OUTER", path.ID
    Set path = facade.LinePath("Row", 0#, 0#, 100#, 0#)
    Set TestRuleFacade = facade
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Собирает контрольное описание штатным сборщиком с крупной фиксированной сеткой без отдельной логики рецепта.
Private Function TestBuildFixture(ByVal facade As CSectionDefinitionBuilder, ByVal builder As CSectionBuilder) As CSectionModel
    Dim mesh As TSectionMeshOptions
    mesh.StepX = 1000#: mesh.StepY = 1000#: mesh.BoundarySubdivisions = 1: mesh.MaterialID = 1
    Set TestBuildFixture = builder.Build(facade.Finish, mesh)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Внутренняя грань Upper смежных прямоугольников не входит в общую оболочку: подпись остаётся у y=250, стержни — у y=290.
Private Sub TestLocalRowBoundaryAnchor()
    Dim facade As CSectionDefinitionBuilder: Set facade = New CSectionDefinitionBuilder: facade.Initialize "TestLocalRowBoundary"
    Dim path As CGeometryPath
    Set path = facade.RectanglePath("Upper", 0#, 250#, 250#, 550#)
    Set path = facade.RectanglePath("Lower", 0#, 0#, 600#, 250#)
    Set path = facade.AdjacentRectangles("Boundary", Array(Array(0#, 0#, 600#, 250#), Array(0#, 250#, 250#, 550#)))
    facade.Outer "OUTER", path.ID
    Set path = facade.LinePath("UpperBottom", 200#, 290#, 50#, 290#)
    Dim options As CRebarRuleOptions: Set options = New CRebarRuleOptions: options.AnnotationGroup = "UpperBottom"
    facade.RebarByCount "UpperBottom.Main", path.ID, 2, 20#, options
    facade.RebarContour "UpperBottom.Main", "OUTER", "Upper", 40#, "UpperBottom.Row2", "UpperBottom.Row3", normalSign:=-1#
    facade.RebarRowLabel "UpperBottom", "UpperBottom.Main", 0#, -1#
    Dim mesh As TSectionMeshOptions: mesh.StepX = 100#: mesh.StepY = 100#: mesh.MaterialID = 1: mesh.CellGeometry = "Contour"
    Dim builder As CSectionBuilder: Set builder = New CSectionBuilder
    Dim model As CSectionModel: Set model = builder.Build(facade.Finish, mesh)
    TestCheck "localLabel.single", model.Annotations.Count = 1
    TestCheck "localLabel.resolvedType", model.Annotations.AnnotationType(1) = "REBAR_LOCAL_ANNOTATION"
    TestNear "localLabel.upperBoundary.start", model.Annotations.StartY(1), 250#
    TestNear "localLabel.upperBoundary.end", model.Annotations.EndY(1), 250#
    TestNear "localLabel.normal", model.Annotations.NormalY(1), -1#
    TestNear "localLabel.barUnchanged.first", model.RebarY(1), 290#
    TestNear "localLabel.barUnchanged.second", model.RebarY(2), 290#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет недопустимые отступы, шаг, режим замыкания, циклическую зависимость и отсутствие допустимой проекции.
Private Function TestInvalidRuleRejected(ByVal scenario As Long) As Boolean
    On Error GoTo Rejected
    Dim facade As CSectionDefinitionBuilder, options As CRebarRuleOptions, builder As CSectionBuilder, model As CSectionModel, path As CGeometryPath
    Set facade = TestRuleFacade: Set options = New CRebarRuleOptions
    Select Case scenario
        Case 0: options.StartSetback = -1#: facade.RebarByCount "Bad", "Row", 2, 10#, options
        Case 1: options.StartSetback = 60#: options.EndSetback = 60#: facade.RebarByCount "Bad", "Row", 2, 10#, options
        Case 2: facade.RebarByFixedStep "Bad", "Row", 30#, 10#, options
        Case 3: facade.RebarByMaxStep "Bad", "Row", 0#, 10#, options
        Case 4
            Set path = facade.CirclePath("Closed", 0#, 0#, 100#)
            options.EndMode = IncludeStart: options.StartSetback = 1#
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
    Set builder = New CSectionBuilder: Set model = TestBuildFixture(facade, builder)
    Exit Function
Rejected:
    TestInvalidRuleRejected = (Err.Number = vbObjectError + 5530)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Учитывает отдельное условие проверки и сохраняет его имя в итоговом протоколе.
Private Sub TestCheck(ByVal name As String, ByVal passed As Boolean)
    If passed Then mPassed = mPassed + 1 Else mFailed = mFailed + 1
    mReport = mReport & IIf(passed, "PASS: ", "FAIL: ") & name & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает фактическую и ожидаемую величины с заданным в этом наборе допуском, сохраняя числа в протоколе.
Private Sub TestNear(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    TestCheck name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) <= 0.0000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' роли контуров не зависят от их имён. Проверяет несколько отверстий
' и независимость готовой области от изменяемых копий, выданных для чтения.
Public Function RunRecipeDefinitionTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    TestRectanglesDimensionsAndRadius
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, definition As CSectionDefinition
    Dim region As CConcreteRegion, copy As CConcreteRegion, query As CSectionGeometryQuery, contours As CSectionContours
    Set facade = TestRuleFacade
    Set path = facade.CirclePath("VoidA", -100#, 0#, 20#): facade.Opening "A", path.ID, "OUTER"
    Set path = facade.CirclePath("VoidB", 100#, 0#, 20#): facade.Opening "B", path.ID, "OUTER"
    Set definition = facade.Finish: Set region = definition.ConcreteDomain
    Set query = New CSectionGeometryQuery
    TestCheck "definition.sealed", definition.IsSealed
    TestCheck "definition.loops", definition.LoopCount = 3 And region.LoopCount = 3
    TestCheck "definition.area", Abs(region.Area - GEOM_PI * (1000000# - 800#)) < 0.00001
    TestCheck "definition.concrete", query.ContainsPoint(region, 0#, 0#)
    TestCheck "definition.hole.a", Not query.ContainsPoint(region, -100#, 0#)
    TestCheck "definition.hole.b", Not query.ContainsPoint(region, 100#, 0#)
    Set contours = definition.MaterialContours
    TestCheck "definition.roles", contours.LoopRole(1) = "Outer" And contours.LoopRole(2) = "Opening" And contours.LoopRole(3) = "Opening"
    TestCheck "definition.ids", contours.LoopID(2) = "A" And contours.LoopID(3) = "B"
    contours.Clear
    TestCheck "definition.contours.copy", definition.MaterialContours.Count = 3
    Set copy = definition.ConcreteDomain: copy.SetOwner "Changed", "SP35", "G1"
    TestCheck "definition.region.copy", Len(definition.ConcreteDomain.OwnerID) = 0
    Dim scenario As Long
    For scenario = 0 To 4: TestCheck "definition.invalid." & CStr(scenario), TestInvalidDefinitionRejected(scenario): Next scenario
    GoTo Finished
Failed:
    TestCheck "definition.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRecipeDefinitionTests = mReport & "TOTAL RECIPE DEFINITION: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет целостность описания: уникальные пути, замкнутый наружный контур, владельца отверстия и запрет изменений после фиксации.
Private Function TestInvalidDefinitionRejected(ByVal scenario As Long) As Boolean
    On Error GoTo Rejected
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, definition As CSectionDefinition
    Set facade = TestRuleFacade
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
    TestInvalidDefinitionRejected = (Err.Number = vbObjectError + 5530 Or Err.Number = vbObjectError + 5520)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' адресность ошибок и неактивные поля Circle.
Public Function RunCircleRecipeInputTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, calculation As Long, events As Boolean, updating As Boolean, scenario As Long
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    For scenario = 0 To 4: TestCheck "circle.input." & CStr(scenario), TestCircleInputCase(scenario): Next scenario
    GoTo Finished
Failed:
    TestCheck "circle.input.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    RunCircleRecipeInputTests = mReport & "TOTAL CIRCLE INPUT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет ввод круга через настоящую таблицу: адресные ошибки активных рядов и игнорирование настроек отключённого ряда.
Private Function TestCircleInputCase(ByVal scenario As Long) As Boolean
    Dim expectedReference As String, key As String, settings As CSystemSettingsReader, units As CUnitSystem
    Dim reader As CSectionInputSnapshotReader, context As CSectionRecipeContext, recipe As ISectionRecipe
    Dim definition As CSectionDefinition, table As CSectionInputTable, builder As CSectionBuilder, mesh As TSectionMeshOptions, model As CSectionModel
    On Error GoTo Rejected
    modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Select Case scenario
        Case 0: key = "Rebar.Count": TestSetCircleInput key, 1.5 ' Проверяем недопустимое дробное количество; n=1 покрыто положительными случаями EndMode.
        Case 1: key = "Rebar.AxisDistance": TestSetCircleInput key, 10# / units.InputLengthToInternal(1#)
        Case 2: key = "Rebar.Diameter2": TestSetCircleInput key, -1#
        Case 3
            TestSetCircleInput "Rebar.Count", 0: TestSetCircleInput "Rebar.AxisDistance", "bad"
            TestSetCircleInput "Rebar.Diameter2", "bad": TestSetCircleInput "Rebar.Loc2row", "bad"
        Case 4
            key = "Rebar.Diameter2": TestSetCircleInput key, 400# / units.InputLengthToInternal(1#)
            TestSetCircleInput "Rebar.Loc2row", "SideBySide"
    End Select
    Set reader = New CSectionInputSnapshotReader: Set context = reader.Capture(ThisWorkbook, Array("rngCircleGeometry"), units)
    Set table = context.Table("rngCircleGeometry")
    If Len(key) > 0 Then expectedReference = TestFixtureCircleFieldCell(ThisWorkbook, key).Worksheet.Name & "!" & TestFixtureCircleFieldCell(ThisWorkbook, key).Address(False, False)
    Set recipe = New CSectionRecipeCircle: Set definition = recipe.CreateDefinition(context)
    If scenario = 3 Then TestCircleInputCase = (definition.RuleCount = 0): Exit Function
    mesh.StepX = 40#: mesh.StepY = 30#: mesh.BoundarySubdivisions = 2: mesh.MaterialID = 1
    Set builder = New CSectionBuilder: Set model = builder.Build(definition, mesh)
    Exit Function
Rejected:
    TestCircleInputCase = (Len(expectedReference) > 0 And InStr(1, Err.Description, expectedReference, vbBinaryCompare) > 0)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' активные/неактивные поля и адресность PointedSection.
Public Function RunRoundedRecipeInputTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, scenario As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    For scenario = 0 To 16
        TestFixtureRestoreInputs config, saved
        TestCheck "rounded.input." & CStr(scenario), TestRoundedInputCase(scenario)
    Next scenario
    GoTo Finished
Failed:
    TestCheck "rounded.input.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.EnableEvents = events: Application.ScreenUpdating = updating
    RunRoundedRecipeInputTests = mReport & "TOTAL ROUNDED INPUT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет ввод PointedSection: параметры скругления, армирование, неактивные поля и фактический адрес ошибочной ячейки.
Private Function TestRoundedInputCase(ByVal scenario As Long) As Boolean
    Dim range As Object, invalid As Object, valid As Boolean, reference As String
    Dim settings As CSystemSettingsReader, units As CUnitSystem, reader As CSectionInputSnapshotReader
    Dim context As CSectionRecipeContext, table As CSectionInputTable, recipe As ISectionRecipe, definition As CSectionDefinition
    Dim builder As CSectionBuilder, model As CSectionModel, mesh As TSectionMeshOptions
    On Error GoTo Rejected
    modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
    modTestShapeArchitecture.TestConfigureShapeFixture "RoundedSimple"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set range = ThisWorkbook.Names("rngPointedSectionGeometry").RefersToRange
    Select Case scenario
        Case 0: Set invalid = TestFixtureInputCell(range, "Main", "Left", "n"): invalid.Value2 = 1.5
        Case 1: Set invalid = TestFixtureInputCell(range, "Main", "Left", "d"): invalid.Value2 = -1#
        Case 2: Set invalid = TestFixtureInputCell(range, "Main", "Left", "as"): invalid.Value2 = CDbl(invalid.Offset(0, 1).Value2) / 2#
        Case 3: Set invalid = TestFixtureInputCell(range, "Extra", "Left", "d2"): invalid.Value2 = -1#
        Case 4: Set invalid = TestFixtureInputCell(range, "Extra", "Left", "pos2"): TestFixtureInputCell(range, "Extra", "Left", "d2").Value2 = 10#: invalid.Value2 = "bad"
        Case 5: Set invalid = TestFixtureInputCell(range, "Extra", "Left", "bind2"): TestFixtureInputCell(range, "Extra", "Left", "d2").Value2 = 10#: invalid.Value2 = "bad"
        Case 6
            valid = True: TestFixtureInputCell(range, "Detail", "R2", "left").Value2 = "inactive"
        Case 7
            TestFixtureInputCell(range, "Main", "Left", "n").Value2 = 0: Set invalid = TestFixtureInputCell(range, "Main", "Left", "as"): invalid.Value2 = "inactive"
            TestFixtureInputCell(range, "Extra", "Left", "d2").Value2 = "inactive": TestFixtureInputCell(range, "Extra", "Left", "pos2").Value2 = "inactive"
        Case 8
            TestFixtureInputCell(range, "Main", "Left", "d").Value2 = 0: Set invalid = TestFixtureInputCell(range, "Main", "Left", "as"): invalid.Value2 = "inactive"
            TestFixtureInputCell(range, "Extra", "Left", "d2").Value2 = "inactive": TestFixtureInputCell(range, "Extra", "Left", "pos2").Value2 = "inactive"
        Case 9
            valid = True: TestFixtureInputCell(range, "Extra", "Left", "d2").Value2 = 0: TestFixtureInputCell(range, "Extra", "Left", "d3").Value2 = 0
            TestFixtureInputCell(range, "Extra", "Left", "pos2").Value2 = "inactive": TestFixtureInputCell(range, "Extra", "Left", "bind2").Value2 = "inactive"
        Case 10: Set invalid = TestFixtureInputCell(range, "Detail", "R1", "left"): invalid.Value2 = 1000#
        Case 11
            valid = True: TestFixtureInputCell(range, "Detail", "W", "left").Value2 = 100#
            TestFixtureInputCell(range, "Detail", "R1", "left").Value2 = 220#: TestFixtureInputCell(range, "Detail", "R2", "left").Value2 = 0#
        Case 12: valid = True: TestFixtureInputCell(range, "Detail", "R1", "left").Value2 = 200#: TestFixtureInputCell(range, "Detail", "R1", "right").Value2 = 200#
        Case 13: Set invalid = TestFixtureInputCell(range, "Main", "Left", "as"): invalid.Value2 = 600# ' Сторона действительно исчезает с учётом as противоположной стороны; прежние 380 мм оставляют допустимую ширину.
        Case 14
            valid = True: TestFixtureInputCell(range, "Main", "Left", "n").Value2 = 0: TestFixtureInputCell(range, "Main", "Left", "as").ClearContents
            TestFixtureInputCell(range, "Extra", "Left", "d2").Value2 = "inactive"
        Case 15
            valid = True: TestFixtureInputCell(range, "Main", "Left", "d").Value2 = 0: TestFixtureInputCell(range, "Main", "Left", "as").ClearContents
            TestFixtureInputCell(range, "Extra", "Left", "d2").Value2 = "inactive"
        Case 16: valid = True: TestFixtureInputCell(range, "Main", "Left", "as").Value2 = 380# ' Большой as одной стороны не превращается в такой же отступ остальных сторон.
    End Select
    Set reader = New CSectionInputSnapshotReader: Set context = reader.Capture(ThisWorkbook, Array("rngPointedSectionGeometry"), units)
    Set table = context.Table("rngPointedSectionGeometry")
    If Not valid Then reference = invalid.Worksheet.Name & "!" & invalid.Address(False, False)
    Set recipe = New CSectionRecipePointedSection: Set definition = recipe.CreateDefinition(context)
    mesh.StepX = 40#: mesh.StepY = 30#: mesh.BoundarySubdivisions = 2: mesh.MaterialID = 1
    Set builder = New CSectionBuilder: Set model = builder.Build(definition, mesh)
    TestRoundedInputCase = valid
    Exit Function
Rejected:
    TestRoundedInputCase = (Not valid And Len(reference) > 0 And InStr(1, Err.Description, reference, vbBinaryCompare) > 0)
    If Not TestRoundedInputCase Then mReport = mReport & "ROUNDED INPUT DETAIL: " & CStr(scenario) & "; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' адресность размеров/проекционных ограничителей и неактивные поля.
Public Function RunHollowRecipeInputTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, scenario As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    For scenario = 0 To 21
        TestFixtureRestoreInputs config, saved
        TestCheck "hollow.input." & CStr(scenario), TestHollowInputCase(scenario)
    Next scenario
    GoTo Finished
Failed:
    TestCheck "hollow.input.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.EnableEvents = events: Application.ScreenUpdating = updating
    RunHollowRecipeInputTests = mReport & "TOTAL HOLLOW INPUT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет ввод HollowRectangle, включая положение отверстия, обязательные отступы его осевого контура и отключённые ряды.
Private Function TestHollowInputCase(ByVal scenario As Long) As Boolean
    Dim range As Object, invalid As Object, valid As Boolean, reference As String, factor As Double
    Dim settings As CSystemSettingsReader, units As CUnitSystem, reader As CSectionInputSnapshotReader
    Dim context As CSectionRecipeContext, table As CSectionInputTable, recipe As ISectionRecipe, definition As CSectionDefinition
    Dim builder As CSectionBuilder, model As CSectionModel, mesh As TSectionMeshOptions
    On Error GoTo Rejected
    modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
    modTestShapeArchitecture.TestConfigureShapeFixture "HollowCentered"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    Set range = ThisWorkbook.Names("rngHollowRectangleGeometry").RefersToRange
    Select Case scenario
        Case 0: Set invalid = TestFixtureInputCell(range, "Main", "Top", "n"): invalid.Value2 = 1.5
        Case 1: Set invalid = TestFixtureInputCell(range, "Main", "Top", "d"): invalid.Value2 = -1#
        Case 2: Set invalid = TestFixtureInputCell(range, "Main", "Top", "as"): invalid.Value2 = CDbl(invalid.Offset(0, 1).Value2) / 2#
        Case 3: Set invalid = TestFixtureInputCell(range, "Main", "OpeningLeft", "as"): TestFixtureInputCell(range, "Main", "OpeningLeft", "d").Value2 = 0#: invalid.Value2 = -1#
        Case 4: Set invalid = TestFixtureInputCell(range, "Detail", "Ro", "val"): invalid.Value2 = -1#
        Case 5: Set invalid = TestFixtureInputCell(range, "Detail", "Xo", "val"): invalid.Value2 = 1000# / factor
        Case 6
            TestFixtureInputCell(range, "Main", "Top", "n").Value2 = 0: Set invalid = TestFixtureInputCell(range, "Main", "Top", "as"): invalid.Value2 = "inactive"
            TestFixtureInputCell(range, "Extra", "Top", "d2").Value2 = "inactive": TestFixtureInputCell(range, "Extra", "Top", "d3").Value2 = "inactive"
        Case 7
            valid = True: TestFixtureInputCell(range, "Main", "OpeningLeft", "d").Value2 = 0#: TestFixtureInputCell(range, "Extra", "OpeningLeft", "d2").Value2 = "inactive"
            TestFixtureInputCell(range, "Extra", "OpeningLeft", "d3").Value2 = "inactive"
        Case 8: Set invalid = TestFixtureInputCell(range, "Main", "OpeningLeft", "as"): TestFixtureInputCell(range, "Detail", "Ro", "val").Value2 = 0#: TestFixtureInputCell(range, "Main", "OpeningLeft", "d").Value2 = 0#: invalid.Value2 = "bad"
        Case 9
            valid = True: TestFixtureInputCell(range, "Extra", "Top", "d2").Value2 = 0#: TestFixtureInputCell(range, "Extra", "Top", "d3").Value2 = 0#
            TestFixtureInputCell(range, "Extra", "Top", "pos2").Value2 = "inactive": TestFixtureInputCell(range, "Extra", "Top", "bind2").Value2 = "inactive"
        Case 10: Set invalid = TestFixtureInputCell(range, "Detail", "Ro", "val"): invalid.Value2 = 1000# / factor
        Case 11: valid = True: TestFixtureInputCell(range, "General", "R", "val").Value2 = 500# / factor
        Case 12: Set invalid = TestFixtureInputCell(range, "Extra", "Top", "bind2"): TestFixtureInputCell(range, "Extra", "Top", "d2").Value2 = 16# / factor: invalid.Value2 = "bad"
        Case 13: Set invalid = TestFixtureInputCell(range, "Main", "Top", "n"): TestFixtureInputCell(range, "Main", "Top", "d").Value2 = 0#: invalid.Value2 = 1.5
        Case 14: Set invalid = TestFixtureInputCell(range, "Main", "Top", "d"): invalid.Value2 = "-"
        Case 15: Set invalid = TestFixtureInputCell(range, "Extra", "Top", "d2"): invalid.Value2 = "-"
        Case 16: Set invalid = TestFixtureInputCell(range, "Extra", "Top", "d3"): invalid.Value2 = "-"
        Case 17: Set invalid = TestFixtureInputCell(range, "Extra", "Top", "d2"): invalid.Value2 = "TODO"
        Case 18: valid = True: TestFixtureInputCell(range, "Extra", "Top", "d2").ClearContents: TestFixtureInputCell(range, "Extra", "Top", "d3").ClearContents
        Case 19: valid = True: TestFixtureInputCell(range, "Main", "Top", "d").ClearContents
        Case 20
            valid = True: TestFixtureInputCell(range, "Main", "Top", "n").Value2 = 0: TestFixtureInputCell(range, "Main", "Top", "as").ClearContents
            TestFixtureInputCell(range, "Extra", "Top", "d2").Value2 = "inactive"
        Case 21
            valid = True: TestFixtureInputCell(range, "Main", "OpeningLeft", "d").Value2 = 0#: TestFixtureInputCell(range, "Main", "OpeningLeft", "as").ClearContents
            TestFixtureInputCell(range, "Extra", "OpeningLeft", "d2").Value2 = "inactive"
    End Select
    Set reader = New CSectionInputSnapshotReader: Set context = reader.Capture(ThisWorkbook, Array("rngHollowRectangleGeometry"), units)
    Set table = context.Table("rngHollowRectangleGeometry")
    If Not valid Then reference = invalid.Worksheet.Name & "!" & invalid.Address(False, False)
    Set recipe = New CSectionRecipeHollowRectangle: Set definition = recipe.CreateDefinition(context)
    mesh.StepX = 40#: mesh.StepY = 30#: mesh.BoundarySubdivisions = 2: mesh.MaterialID = 1
    Set builder = New CSectionBuilder: Set model = builder.Build(definition, mesh)
    TestHollowInputCase = valid
    Exit Function
Rejected:
    TestHollowInputCase = (Not valid And Len(reference) > 0 And InStr(1, Err.Description, reference, vbBinaryCompare) > 0)
    If Not TestHollowInputCase Then mReport = mReport & "HOLLOW INPUT DETAIL: " & CStr(scenario) & "; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' полная сборка рецепта общими классами без старых способов геометрии и размещения
' арматуры.
Public Function TestDumpCircleRecipeMatrix() As String
    Dim config As Object, saved As Variant, result As String, counts As Variant, number As Variant, scenario As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, factor As Double, fixture As String
    Dim calculation As Long, events As Boolean, updating As Boolean, builder As CSectionBuilder, model As CSectionModel
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    counts = Array(2, 3, 12, 13)
    For Each number In counts
        For scenario = 0 To 6
            TestSetCircleInput "Rebar.Count", CLng(number)
            TestSetCircleInput "Rebar.Diameter2", IIf(scenario = 0 Or scenario = 2, 0#, 16# / factor)
            TestSetCircleInput "Rebar.Diameter3", IIf(scenario = 0 Or scenario = 1, 0#, 12# / factor)
            TestSetCircleInput "Rebar.Loc2row", IIf(scenario = 5 Or scenario = 6, "SideBySide", "Stacked")
            TestSetCircleInput "Rebar.Loc3row", IIf(scenario = 4 Or scenario = 6, "SideBySide", "Stacked")
            ThisWorkbook.Worksheets("Config").Calculate
            settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = TestBuildCircleRecipe(settings, units, builder)
            fixture = "CircleMatrix.n" & CStr(number) & ".rows" & CStr(scenario)
            result = result & modTestShapeArchitecture.TestPreparedShapeSnapshot(fixture, model, builder.Mesh, builder.Rebars)
        Next scenario
    Next number
    GoTo Finished
Failed:
    result = result & "FAIL: Circle recipe matrix; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    TestDumpCircleRecipeMatrix = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Собирает круг через общий реестр и снимок входных таблиц без отдельного тестового алгоритма геометрии.
Private Function TestBuildCircleRecipe(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, ByVal builder As CSectionBuilder) As CSectionModel
    Set TestBuildCircleRecipe = TestBuildRecipe("Circle", settings, units, builder)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает только зарегистрированные таблицы типа, создаёт его описание и передаёт общему сборщику с текущими настройками сетки.
Private Function TestBuildRecipe(ByVal typeID As String, ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, ByVal builder As CSectionBuilder) As CSectionModel
    Dim reader As CSectionInputSnapshotReader, context As CSectionRecipeContext
    Dim definition As CSectionDefinition, mesh As TSectionMeshOptions
    Dim registry As CSectionTypeRegistry
    Set registry = New CSectionTypeRegistry
    Set reader = New CSectionInputSnapshotReader: Set context = reader.Capture(ThisWorkbook, registry.RecipeTableNames(typeID), units)
    Set definition = registry.CreateDefinition(typeID, context)
    mesh = ReadSectionMeshOptions(settings, units)
    Set TestBuildRecipe = builder.Build(definition, mesh)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' три формы Rectangles проходят один рецепт и общий сборщик.
Public Function TestDumpRectanglesRecipeGeometry() As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
    For Each fixture In Array("Rectangle", "LSection", "TwoRectangles")
        modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set builder = New CSectionBuilder: Set model = TestBuildRecipe("Rectangles", settings, units, builder)
        result = result & modTestShapeArchitecture.TestPreparedShapeSnapshot(CStr(fixture), model, builder.Mesh, builder.Rebars)
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: Rectangles recipe geometry; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    TestDumpRectanglesRecipeGeometry = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' тот же ввод расширенного набора проверок обрабатывается новым общим способом
' сборки.
Public Function TestDumpRectanglesRecipeMatrix() As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant, scenario As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    For Each fixture In Array("Rectangle", "LSection", "TwoRectangles")
        For scenario = 0 To 5
            TestFixtureRestoreInputs config, saved
            modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
            modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
            modTestShapeArchitecture.TestConfigureRectanglesScenario CStr(fixture), scenario
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = TestBuildRecipe("Rectangles", settings, units, builder)
            result = result & modTestShapeArchitecture.TestPreparedShapeSnapshot("RectMatrix." & CStr(fixture) & ".case" & CStr(scenario), model, builder.Mesh, builder.Rebars)
        Next scenario
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: Rectangles recipe matrix; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    TestDumpRectanglesRecipeMatrix = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' точные пути общего builder на тех же 24 исходных fixtures.
Public Function TestDumpRoundedRecipeMatrix() As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant, scenario As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    For Each fixture In Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft")
        For scenario = 0 To 5
            TestFixtureRestoreInputs config, saved
            modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
            modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
            modTestShapeArchitecture.TestConfigureRoundedScenario scenario
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = TestBuildRecipe("PointedSection", settings, units, builder)
            result = result & modTestShapeArchitecture.TestPreparedShapeSnapshot("RoundedMatrix." & CStr(fixture) & ".case" & CStr(scenario), model, builder.Mesh, builder.Rebars)
        Next scenario
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: Rounded recipe matrix; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    TestDumpRoundedRecipeMatrix = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' точные дуги и проекции для 16 исходных Hollow fixtures.
Public Function TestDumpHollowRecipeMatrix() As String
    TestDumpHollowRecipeMatrix = TestHollowRecipeMatrix(0, 7)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отсутствующие прямые участки и sharp frame тем же рецептом.
Public Function TestDumpHollowRecipeEdgeMatrix() As String
    TestDumpHollowRecipeEdgeMatrix = TestHollowRecipeMatrix(8, 10)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает точные снимки диапазона сценариев HollowRectangle с центральным и смещённым отверстием, восстанавливая входы.
Private Function TestHollowRecipeMatrix(ByVal firstScenario As Long, ByVal lastScenario As Long) As String
    Dim config As Object, saved As Variant, result As String, fixture As Variant, scenario As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    For Each fixture In Array("HollowCentered", "HollowOffset")
        For scenario = firstScenario To lastScenario
            TestFixtureRestoreInputs config, saved
            modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
            modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
            modTestShapeArchitecture.TestConfigureHollowScenario scenario
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = TestBuildRecipe("HollowRectangle", settings, units, builder)
            result = result & modTestShapeArchitecture.TestPreparedShapeSnapshot("HollowMatrix." & CStr(fixture) & ".case" & CStr(scenario), model, builder.Mesh, builder.Rebars)
        Next scenario
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: Hollow recipe matrix; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    TestHollowRecipeMatrix = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сумма готовых подписей в полной матрице перенесенных рецептов.
Public Function RunRecipeAnnotationCountTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, fixture As Variant, scenario As Long, number As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean, factor As Double
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    For Each number In Array(2, 3, 12, 13)
        For scenario = 0 To 6
            TestSetCircleInput "Rebar.Count", CLng(number)
            TestSetCircleInput "Rebar.Diameter2", IIf(scenario = 0 Or scenario = 2, 0#, 16# / factor)
            TestSetCircleInput "Rebar.Diameter3", IIf(scenario = 0 Or scenario = 1, 0#, 12# / factor)
            TestSetCircleInput "Rebar.Loc2row", IIf(scenario = 5 Or scenario = 6, "SideBySide", "Stacked")
            TestSetCircleInput "Rebar.Loc3row", IIf(scenario = 4 Or scenario = 6, "SideBySide", "Stacked")
            ThisWorkbook.Worksheets("Config").Calculate
            settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = TestBuildRecipe("Circle", settings, units, builder)
            TestCheck "annotations.total.Circle.n" & CStr(number) & ".case" & CStr(scenario), TestAnnotatedBarCount(model.Annotations) = model.RebarCount
        Next scenario
    Next number
    For Each fixture In Array("Rectangle", "LSection", "TwoRectangles")
        For scenario = 0 To 5
            TestFixtureRestoreInputs config, saved
            modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
            modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
            modTestShapeArchitecture.TestConfigureRectanglesScenario CStr(fixture), scenario
            settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = TestBuildRecipe("Rectangles", settings, units, builder)
            TestCheck "annotations.total." & CStr(fixture) & ".case" & CStr(scenario), TestAnnotatedBarCount(model.Annotations) = model.RebarCount
        Next scenario
    Next fixture
    For Each fixture In Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft")
        For scenario = 0 To 5
            TestFixtureRestoreInputs config, saved
            modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
            modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
            modTestShapeArchitecture.TestConfigureRoundedScenario scenario
            settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = TestBuildRecipe("PointedSection", settings, units, builder)
            TestCheck "annotations.total." & CStr(fixture) & ".case" & CStr(scenario), TestAnnotatedBarCount(model.Annotations) = model.RebarCount
        Next scenario
    Next fixture
    For Each fixture In Array("HollowCentered", "HollowOffset")
        For scenario = 0 To 10
            TestFixtureRestoreInputs config, saved
            modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
            modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
            modTestShapeArchitecture.TestConfigureHollowScenario scenario
            settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = TestBuildRecipe("HollowRectangle", settings, units, builder)
            TestCheck "annotations.total." & CStr(fixture) & ".case" & CStr(scenario), TestAnnotatedBarCount(model.Annotations) = model.RebarCount
        Next scenario
    Next fixture
    GoTo Finished
Failed:
    TestCheck "annotations.total.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    RunRecipeAnnotationCountTests = mReport & "TOTAL RECIPE ANNOTATIONS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет объединение двух и трёх смежных прямоугольников: площадь, вогнутые участки и отказ для недопустимого стыка.
Private Sub TestAdjacentRectangles()
    Dim query As CSectionGeometryQuery, path As CGeometryPath, region As CConcreteRegion
    Set query = New CSectionGeometryQuery
    Set path = query.AdjacentRectangles("Pair", Array(Array(0#, 0#, 600#, 200#), Array(-50#, 200#, 300#, 400#)))
    Set region = query.RegionFromPath(path)
    TestNear "adjacent.area", region.Area, 240000#
    TestCheck "adjacent.closed", path.IsClosed And region.LoopCount = 1
    TestCheck "adjacent.lower", query.ContainsPoint(region, 500#, 100#)
    TestCheck "adjacent.upper", query.ContainsPoint(region, -25#, 300#)
    TestCheck "adjacent.concavity", Not query.ContainsPoint(region, 500#, 300#)
    Set path = query.AdjacentRectangles("Three", Array(Array(0#, 0#, 100#, 100#), Array(100#, 0#, 100#, 100#), Array(0#, 100#, 100#, 100#)))
    Set region = query.RegionFromPath(path)
    TestNear "adjacent.three.area", region.Area, 30000#
    TestCheck "adjacent.three.concavity", Not query.ContainsPoint(region, 150#, 150#)
    TestCheck "adjacent.overlap.rejected", TestInvalidAdjacentRejected(0)
    TestCheck "adjacent.disjoint.rejected", TestInvalidAdjacentRejected(1)
    TestCheck "adjacent.corner.rejected", TestInvalidAdjacentRejected(2)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет отказ при наложении прямоугольников, разрыве между ними и касании только общей вершиной.
Private Function TestInvalidAdjacentRejected(ByVal scenario As Long) As Boolean
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
    TestInvalidAdjacentRejected = (InStr(1, Err.Source, "CSectionGeometryQuery", vbBinaryCompare) > 0)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Изменяет поле круга через общий поиск ячейки по ключу, без постоянного адреса в листе.
Private Sub TestSetCircleInput(ByVal key As String, ByVal value As Variant)
    Dim cell As Object
    Set cell = TestFixtureCircleFieldCell(ThisWorkbook, key)
    cell.Value2 = value
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' обе независимые привязки круга, нечётные n, направления и реальные подписи.
Public Function RunCircleBindingTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim number As Variant, loc2 As Variant, loc3 As Variant, bind2 As Variant, bind3 As Variant
    Dim definition As CSectionDefinition, reference As CRebarLayout, actual As CRebarLayout
    Dim before As CRebarRuleExecutor, after As CRebarRuleExecutor, annotations As CSectionAnnotations
    Dim ordinal As Long, row As Long, position As Long, original As Long, owner As Long, oldOwner As Long, previous As Long
    Dim expected As Long, selected As Boolean, id As String, prefix As String, binding As String
    On Error GoTo Failed
    For Each number In Array(2, 3, 8, 9)
        For Each loc2 In Array("Stacked", "SideBySide")
            For Each loc3 In Array("Stacked", "SideBySide")
                Set definition = TestFixtureCircleDefinition(Array(1000#, 50#, number, 20#, 12#, 10#, loc2, loc3, "EachBar", "EachBar"))
                Set before = New CRebarRuleExecutor: Set reference = TestFixtureAnnotatedLayout(definition, before)
                For Each bind2 In Array("EachBar", "EverySecondBar")
                    For Each bind3 In Array("EachBar", "EverySecondBar")
                        prefix = "circle.bind.n" & CStr(number) & "." & loc2 & "." & loc3 & "." & bind2 & "." & bind3
                        Set definition = TestFixtureCircleDefinition(Array(1000#, 50#, number, 20#, 12#, 10#, loc2, loc3, bind2, bind3))
                        Set after = New CRebarRuleExecutor: Set actual = TestFixtureAnnotatedLayout(definition, after, annotations)
                        expected = CLng(number) + IIf(bind2 = "EachBar", number, (CLng(number) + 1) \ 2) + IIf(bind3 = "EachBar", number, (CLng(number) + 1) \ 2)
                        TestCheck prefix & ".count", actual.Count = expected
                        TestCheck prefix & ".annotation", TestAnnotatedBarCount(annotations) = actual.Count
                        previous = 0
                        For ordinal = 1 To CLng(number)
                            For row = 1 To 3
                                id = "B" & CStr(ordinal): binding = "EachBar"
                                If row = 2 Then id = id & "_R2": binding = CStr(bind2)
                                If row = 3 Then id = id & "_R3": binding = CStr(bind3)
                                selected = (binding = "EachBar" Or ordinal Mod 2 = 1)
                                position = TestCircleSourcePosition(after, id)
                                TestCheck prefix & ".selected." & id, (position > 0) = selected
                                If selected And position > 0 Then
                                    original = TestCircleSourcePosition(before, id)
                                    owner = after.PositionPhysicalOwner(position): oldOwner = before.PositionPhysicalOwner(original)
                                    TestNear prefix & ".x." & id, actual.X(owner), reference.X(oldOwner)
                                    TestNear prefix & ".y." & id, actual.Y(owner), reference.Y(oldOwner)
                                    TestNear prefix & ".d." & id, actual.Diameter(owner), reference.Diameter(oldOwner)
                                    TestCheck prefix & ".order." & id, owner > previous
                                    previous = owner
                                End If
                            Next row
                        Next ordinal
                    Next bind3
                Next bind2
            Next loc3
        Next loc2
    Next number
    TestCheck "circle.bind.invalid.active2", TestCircleBindingInputCase(12#, 0#, "bad", "bad", False)
    TestCheck "circle.bind.invalid.active3", TestCircleBindingInputCase(0#, 10#, "bad", "bad", False)
    TestCheck "circle.bind.blank.active2", TestCircleBindingInputCase(12#, 0#, "", "EachBar", False)
    TestCheck "circle.bind.inactive.rows", TestCircleBindingInputCase(0#, 0#, "bad", "bad", True)
    TestCheck "circle.bind.inactive.row2", TestCircleBindingInputCase(0#, 10#, "bad", "EverySecondBar", True)
    TestCheck "circle.bind.inactive.row3", TestCircleBindingInputCase(12#, 0#, "EverySecondBar", "bad", True)
    GoTo Finished
Failed:
    TestCheck "circle.bind.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunCircleBindingTests = mReport & "TOTAL CIRCLE BINDINGS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ищет исходное имя стержня; порядковый номер записи может отличаться от исходного номера.
Private Function TestCircleSourcePosition(ByVal executor As CRebarRuleExecutor, ByVal id As String) As Long
    Dim i As Long
    For i = 1 To executor.PositionCount
        If executor.PositionSourceID(i) = id Then TestCircleSourcePosition = i: Exit Function
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет обязательный выбор только для существующего дополнительного ряда.
Private Function TestCircleBindingInputCase(ByVal diameter2 As Double, ByVal diameter3 As Double, _
        ByVal binding2 As String, ByVal binding3 As String, ByVal valid As Boolean) As Boolean
    On Error GoTo Rejected
    Dim definition As CSectionDefinition
    Set definition = TestFixtureCircleDefinition(Array(1000#, 50#, 8, 20#, diameter2, diameter3, "Stacked", "Stacked", binding2, binding3))
    TestCircleBindingInputCase = valid
    Exit Function
Rejected:
    TestCircleBindingInputCase = Not valid And InStr(1, Err.Description, "rngCircleGeometry", vbBinaryCompare) > 0
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' два реальных выпадающих списка, независимость и адрес ошибки Config.
Public Function RunCircleBindingWorkbookTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim table As Object, saved As Variant, index As Long
    Dim keys As Variant, choices As Variant, formula As String, optionValue As Variant, cell As Object
    On Error GoTo Failed
    Set table = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(table)
    modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
    keys = Array("Rebar.Bind2row", "Rebar.Bind3row")
    TestCheck "circle.bind.cells.independent", TestFixtureCircleFieldCell(ThisWorkbook, CStr(keys(0))).Address <> TestFixtureCircleFieldCell(ThisWorkbook, CStr(keys(1))).Address
    For index = 0 To 1
        Set cell = TestFixtureCircleFieldCell(ThisWorkbook, CStr(keys(index)))
        formula = cell.Validation.Formula1
        choices = Split(formula, Application.International(xlListSeparator))
        TestCheck "circle.bind.list.inline." & CStr(index), Left$(formula, 1) <> "=" And cell.Validation.Type = xlValidateList
        TestCheck "circle.bind.list.size." & CStr(index), UBound(choices) = 1
        TestCheck "circle.bind.list.values." & CStr(index), choices(0) = "EachBar" And choices(1) = "EverySecondBar"
        TestCheck "circle.bind.list.center." & CStr(index), cell.HorizontalAlignment = xlCenter
        For Each optionValue In Array("EachBar", "EverySecondBar")
            cell.Value2 = optionValue
            TestCheck "circle.bind.list.accept." & CStr(index) & "." & optionValue, cell.Validation.Value
        Next optionValue
        cell.Value2 = "bad": TestCheck "circle.bind.list.reject." & CStr(index), Not cell.Validation.Value
    Next index
    TestSetCircleInput "Rebar.Bind2row", "EverySecondBar": TestSetCircleInput "Rebar.Bind3row", "EachBar"
    modSectionRecipeSetup.UpdateSectionRecipeTypes
    TestCheck "circle.bind.update.preserves2", TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.Bind2row").Value2 = "EverySecondBar"
    TestCheck "circle.bind.update.preserves3", TestFixtureCircleFieldCell(ThisWorkbook, "Rebar.Bind3row").Value2 = "EachBar"
    TestCheck "circle.bind.workbook.address2", TestCircleBindingWorkbookRejected("Rebar.Bind2row", "Rebar.Diameter2")
    TestCheck "circle.bind.workbook.address3", TestCircleBindingWorkbookRejected("Rebar.Bind3row", "Rebar.Diameter3")
    GoTo Restore
Failed:
    TestCheck "circle.bind.workbook.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Restore:
    If Not table Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs table, saved
    RunCircleBindingWorkbookTests = mReport & "TOTAL CIRCLE BINDING WORKBOOK: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ошибка активной привязки должна содержать действительный адрес её пользовательской ячейки.
Private Function TestCircleBindingWorkbookRejected(ByVal key As String, ByVal diameterKey As String) As Boolean
    On Error GoTo Rejected
    Dim settings As CSystemSettingsReader, units As CUnitSystem, context As CSectionRecipeContext, reader As CSectionInputSnapshotReader
    Dim recipe As ISectionRecipe, definition As CSectionDefinition, factor As Double
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    TestSetCircleInput "Circle.Diameter", 1000# / factor: TestSetCircleInput "Rebar.AxisDistance", 50# / factor
    TestSetCircleInput "Rebar.Count", 8: TestSetCircleInput "Rebar.Diameter", 20# / factor
    TestSetCircleInput "Rebar.Diameter2", 0#: TestSetCircleInput "Rebar.Diameter3", 0#
    TestSetCircleInput diameterKey, 12# / factor: TestSetCircleInput "Rebar.Loc2row", "Stacked": TestSetCircleInput "Rebar.Loc3row", "Stacked"
    TestSetCircleInput key, "bad"
    Set reader = New CSectionInputSnapshotReader: Set context = reader.Capture(ThisWorkbook, Array("rngCircleGeometry"), units)
    Set recipe = New CSectionRecipeCircle: Set definition = recipe.CreateDefinition(context)
    Exit Function
Rejected:
    Dim reason As String
    reason = Err.Description ' Чтение нового helper не должно сбросить исходную ошибку до её проверки.
    TestCircleBindingWorkbookRejected = InStr(1, reason, TestFixtureCircleFieldCell(ThisWorkbook, key).Worksheet.Name & "!" & TestFixtureCircleFieldCell(ThisWorkbook, key).Address(False, False), vbBinaryCompare) > 0
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' НДС/несущая/трещины для фактической модели нового Circle.
Public Function TestDumpCircleRecipeCalculations(Optional ByVal crackCode As String = "SP35") As String
    TestDumpCircleRecipeCalculations = TestRecipeCalculations("Circle", Array("Circle"), crackCode)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет расчётные снимки прямоугольника и двух вариантов составного сечения при выбранной методике трещин.
Public Function TestDumpRectanglesRecipeCalculations(Optional ByVal crackCode As String = "SP35") As String
    TestDumpRectanglesRecipeCalculations = TestRecipeCalculations("Rectangles", Array("Rectangle", "LSection", "TwoRectangles"), crackCode)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет расчётные снимки четырёх вариантов PointedSection при выбранной методике трещин.
Public Function TestDumpRoundedRecipeCalculations(Optional ByVal crackCode As String = "SP35") As String
    TestDumpRoundedRecipeCalculations = TestRecipeCalculations("PointedSection", _
        Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft"), crackCode)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет расчётные снимки PointedSection с включёнными дуговыми участками арматуры.
Public Function TestDumpRoundedRecipeArcCalculations(Optional ByVal crackCode As String = "SP35") As String
    TestDumpRoundedRecipeArcCalculations = TestRecipeCalculations("PointedSection", _
        Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft"), crackCode, 1)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' независимо суммирует N, Mx и My для новой геометрии с дугами.
' Из вклада стали вычитается вклад замещённого бетона в той же точке.
Public Function RunRoundedRecipeEquilibriumTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, fixture As Variant, events As Boolean, updating As Boolean
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    On Error GoTo Failed
    events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
    For Each fixture In Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft")
        modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
        modTestShapeArchitecture.TestConfigureRoundedScenario 1
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set builder = New CSectionBuilder: Set model = TestBuildRecipe("PointedSection", settings, units, builder)
        TestCheckModelEquilibrium CStr(fixture), model, settings, units
    Next fixture
    GoTo Finished
Failed:
    TestCheck "equilibrium.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.EnableEvents = events: Application.ScreenUpdating = updating
    RunRoundedRecipeEquilibriumTests = mReport & "TOTAL ROUNDED EQUILIBRIUM: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' независимо суммирует усилия для переданной модели с контурами без смены исходных
' данных.
Public Function TestReviewModelEquilibrium(ByVal fixture As String, ByVal model As CSectionModel, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, ByRef passed As Long, ByRef failed As Long) As String
    mPassed = 0: mFailed = 0: mReport = ""
    TestCheckModelEquilibrium fixture, model, settings, units
    passed = mPassed: failed = mFailed: TestReviewModelEquilibrium = mReport
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимо суммирует усилия по бетонным элементам и арматуре для трёх режимов материалов и четырёх нагрузок.
Private Sub TestCheckModelEquilibrium(ByVal fixture As String, ByVal model As CSectionModel, _
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
            TestCheck prefix & ".converged", solver.Converged
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
                force = (steel.StressAtStrain(strain) - concrete.StressAtStrain(strain)) * model.RebarArea(i) ' Из бетона вычитается площадь, занятая арматурой; иначе независимый контроль удвоил бы её вклад.
                sn = sn + force: smx = smx + force * y: smy = smy + force * x
            Next i
            TestCheck prefix & ".sum.N", Abs(sn - solver.Nint) <= 0.0000001 + Abs(sn) * 0.000000000001
            TestCheck prefix & ".sum.Mx", Abs(smx - solver.Mxint) <= 0.0001 + Abs(smx) * 0.000000000001
            TestCheck prefix & ".sum.My", Abs(smy - solver.Myint) <= 0.0001 + Abs(smy) * 0.000000000001
            ' Независимые суммы уже проверены с допуском в единицах усилий. Невязка должна
            ' соответствовать усилию, возвращённому решателем; другой порядок вычитания
            ' близких больших чисел может усилить округление Double.
            TestNear prefix & ".residual.N", solver.Nint - n, solver.ResidualN
            TestNear prefix & ".residual.Mx", solver.Mxint - mx, solver.ResidualMx
            TestNear prefix & ".residual.My", solver.Myint - my, solver.ResidualMy
NextLoad:
        Next loadIndex
    Next purpose
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' стандартные, дуговые, Ro=0 и смещенные проекции.
Public Function TestDumpHollowRecipeCalculations(Optional ByVal crackCode As String = "SP35", Optional ByVal scenario As Long = 0) As String
    TestDumpHollowRecipeCalculations = TestRecipeCalculations("HollowRectangle", Array("HollowCentered", "HollowOffset"), crackCode, 0, scenario)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' независимо суммирует усилия для четырёх вариантов пустотелого прямоугольника.
Public Function RunHollowRecipeEquilibriumTests() As String
    RunHollowRecipeEquilibriumTests = TestHollowRecipeEquilibrium(0, 3)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' исчезнувшие прямые грани, полный круг и острые углы.
Public Function RunHollowRecipeEdgeEquilibriumTests() As String
    RunHollowRecipeEdgeEquilibriumTests = TestHollowRecipeEquilibrium(8, 10)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет поэлементное равновесие заданного диапазона вариантов HollowRectangle и возвращает входные данные после проверки.
Private Function TestHollowRecipeEquilibrium(ByVal firstScenario As Long, ByVal lastScenario As Long) As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, fixture As Variant, scenario As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    For Each fixture In Array("HollowCentered", "HollowOffset")
        For scenario = firstScenario To lastScenario
            TestFixtureRestoreInputs config, saved
            modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
            modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
            modTestShapeArchitecture.TestConfigureHollowScenario scenario
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings
            Set builder = New CSectionBuilder: Set model = TestBuildRecipe("HollowRectangle", settings, units, builder)
            TestCheckModelEquilibrium CStr(fixture) & ".case" & CStr(scenario), model, settings, units
        Next scenario
    Next fixture
    GoTo Finished
Failed:
    TestCheck "hollow.equilibrium.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    TestHollowRecipeEquilibrium = mReport & "TOTAL HOLLOW EQUILIBRIUM: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Настраивает реальные профили прочности и трещин, строит выбранные рецепты и сохраняет их расчётные снимки с восстановлением входов.
Private Function TestRecipeCalculations(ByVal typeID As String, ByVal fixtures As Variant, ByVal crackCode As String, _
        Optional ByVal roundedScenario As Long = 0, Optional ByVal hollowScenario As Long = 0) As String
    Dim config As Object, saved As Variant, result As String, settings As CSystemSettingsReader, units As CUnitSystem
    Dim profiles As Object, i As Long, key As String, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean, fixture As Variant
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
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
    TestSetSystemInput "SLS.Crack.Code", crackCode
    modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
    For Each fixture In fixtures
        modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
        If roundedScenario > 0 Then modTestShapeArchitecture.TestConfigureRoundedScenario roundedScenario
        If hollowScenario > 0 Then modTestShapeArchitecture.TestConfigureHollowScenario hollowScenario
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set builder = New CSectionBuilder: Set model = TestBuildRecipe(typeID, settings, units, builder)
        Dim label As String
        label = CStr(fixture): If roundedScenario > 0 Then label = label & ".Arc"
        If hollowScenario > 0 Then label = label & ".case" & CStr(hollowScenario)
        result = result & modTestShapeArchitecture.TestPreparedCalculationSnapshot(label, model)
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: recipe calculations; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    TestRecipeCalculations = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' настоящий Excel PNG и структурный снимок схемы.
Public Function TestDumpCircleRecipePreview(ByVal outputDirectory As String) As String
    TestDumpCircleRecipePreview = TestRecipePreview("Circle", Array("Circle"), outputDirectory)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет схемы прямоугольника и составных сечений, построенных штатными рецептами.
Public Function TestDumpRectanglesRecipePreview(ByVal outputDirectory As String) As String
    TestDumpRectanglesRecipePreview = TestRecipePreview("Rectangles", Array("Rectangle", "LSection", "TwoRectangles"), outputDirectory)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет схемы четырёх контрольных вариантов PointedSection.
Public Function TestDumpRoundedRecipePreview(ByVal outputDirectory As String) As String
    TestDumpRoundedRecipePreview = TestRecipePreview("PointedSection", _
        Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft"), outputDirectory)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет схемы PointedSection с арматурой на дуговых участках.
Public Function TestDumpRoundedRecipeArcPreview(ByVal outputDirectory As String) As String
    TestDumpRoundedRecipeArcPreview = TestRecipePreview("PointedSection", _
        Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft"), outputDirectory, 1)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' настоящий Excel preview стандартных/измененных Hollow-сценариев.
Public Function TestDumpHollowRecipePreview(ByVal outputDirectory As String, Optional ByVal scenario As Long = 0) As String
    TestDumpHollowRecipePreview = TestRecipePreview("HollowRectangle", Array("HollowCentered", "HollowOffset"), outputDirectory, 0, scenario)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит варианты рецепта, сохраняет схему и её координаты, затем восстанавливает исходные таблицы и режим Excel.
Private Function TestRecipePreview(ByVal typeID As String, ByVal fixtures As Variant, ByVal outputDirectory As String, _
        Optional ByVal roundedScenario As Long = 0, Optional ByVal hollowScenario As Long = 0) As String
    Dim config As Object, saved As Variant, result As String, settings As CSystemSettingsReader, units As CUnitSystem
    Dim builder As CSectionBuilder, model As CSectionModel, calculation As Long, events As Boolean, updating As Boolean, fixture As Variant
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = True
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    modTestShapeArchitecture.TestConfigureShapeFixture "Circle"
    For Each fixture In fixtures
        modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
        If roundedScenario > 0 Then modTestShapeArchitecture.TestConfigureRoundedScenario roundedScenario
        If hollowScenario > 0 Then modTestShapeArchitecture.TestConfigureHollowScenario hollowScenario
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set builder = New CSectionBuilder: Set model = TestBuildRecipe(typeID, settings, units, builder)
        Dim label As String
        label = CStr(fixture): If roundedScenario > 0 Then label = label & ".Arc"
        If hollowScenario > 0 Then label = label & ".case" & CStr(hollowScenario)
        result = result & modTestShapeArchitecture.TestPreparedPreviewSnapshot(label, model, outputDirectory)
    Next fixture
    GoTo Finished
Failed:
    result = result & "FAIL: recipe preview; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    TestRecipePreview = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет системное значение по ключу именованной таблицы; отсутствие ключа является ошибкой подготовки проверки.
Private Sub TestSetSystemInput(ByVal key As String, ByVal value As Variant)
    Dim table As Object, i As Long
    Set table = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    For i = 1 To table.Rows.Count
        If CStr(table.Cells(i, 1).Value2) = key Then table.Cells(i, 2).Value2 = value: Exit Sub
    Next i
    Err.Raise vbObjectError + 5531, "modTestSectionRecipes", "Не найден системный ключ '" & key & "'."
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' рецепты используют названия сторон из Config, а не внутренние имена и прежние варианты.
Public Function RunRecipeCaptionTests() As String
    Dim config As Object, saved As Variant, fixture As Variant, table As Object, reader As CSectionInputReader
    Dim settings As CSystemSettingsReader, units As CUnitSystem, typeID As String, tableName As String
    Dim block As Variant, key As Variant, keys As Variant, row As Long, caption As Variant, count As Long
    Dim builder As CSectionBuilder, model As CSectionModel, number As Long, reason As String
    Dim events As Boolean, updating As Boolean
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = ""
    events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    For Each fixture In Array("Circle", "LSection", "RoundedSimple", "HollowCentered")
        TestFixtureRestoreInputs config, saved: modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
        Select Case CStr(fixture)
            Case "Circle": typeID = "Circle"
            Case "LSection": typeID = "Rectangles"
            Case "RoundedSimple": typeID = "PointedSection"
            Case Else: typeID = "HollowRectangle"
        End Select
        tableName = "rng" & typeID & "Geometry"
        Set table = ThisWorkbook.Names(tableName).RefersToRange: Set reader = TestFixtureInput(table.Value2)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set builder = New CSectionBuilder: Set model = TestBuildRecipe(typeID, settings, units, builder): count = model.RebarCount
        For Each block In Array("Main", "Extra")
            keys = reader.RowKeys(CStr(block))
            For Each key In keys
                row = reader.RowIndex(CStr(block), CStr(key))
                caption = table.Cells(row, reader.ColumnIndex(CStr(block), "name")).Value2
                table.Cells(row, reader.ColumnIndex(CStr(block), "name")).Value2 = "Свободное название стороны"
                Set builder = New CSectionBuilder: Set model = TestBuildRecipe(typeID, settings, units, builder)
                TestCheck typeID & ".visible.caption." & CStr(block) & CStr(key), model.RebarCount = count
                table.Cells(row, reader.ColumnIndex(CStr(block), "name")).Value2 = caption
                table.Cells(row, 1).Value2 = "Missing"
                number = 0: reason = ""
                On Error Resume Next
                Set builder = New CSectionBuilder: Set model = TestBuildRecipe(typeID, settings, units, builder)
                number = Err.Number: reason = Err.Description: Err.Clear
                On Error GoTo Failed
                TestCheck typeID & ".missing.key." & CStr(block) & CStr(key), number <> 0 And InStr(1, reason, CStr(key), vbTextCompare) > 0
                table.Cells(row, 1).Value2 = key
            Next key
        Next block
    Next fixture
    GoTo Finished
Failed:
    TestCheck "recipe.captions.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.EnableEvents = events: Application.ScreenUpdating = updating
    RunRecipeCaptionTests = mReport & "TOTAL RECIPE CAPTIONS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отсутствие утверждённой подписи должно дать понятную ошибку с её названием.
Private Function TestRecipeCaptionRejected(ByVal typeID As String, ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, ByVal caption As String) As Boolean
    Dim builder As CSectionBuilder, model As CSectionModel
    On Error GoTo Rejected
    Set builder = New CSectionBuilder: Set model = TestBuildRecipe(typeID, settings, units, builder)
    Exit Function
Rejected:
    TestRecipeCaptionRejected = InStr(1, Err.Description, caption, vbBinaryCompare) > 0
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ограниченный пакет общего порога: равенство, отключение, разные диаметры и зависимости.
' Использует штатные правила и сборщик без Excel/AutoCAD и без расчёта НДС.
Public Function RunRecipeMinimumAreaTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    TestMinimumRebarArea
    GoTo Finish
Failed:
    TestCheck "minimumArea.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finish:
    RunRecipeMinimumAreaTests = mReport & "TOTAL_RECIPE_MINIMUM_AREA: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Мелкие исходные стержни остаются геометрической опорой зависимых рядов и проекций,
' но не получают physical owner/ID. Подписи и trace содержат только принятые стержни.
Private Sub TestMinimumRebarArea()
    Dim facade As CSectionDefinitionBuilder: Set facade = TestRuleFacade
    Dim options As CRebarRuleOptions: Set options = New CRebarRuleOptions
    options.AnnotationGroup = "Small": options.SourcePrefix = "Seed"
    facade.RebarByCount "Small", "Row", 3, 8#, options
    Set options = New CRebarRuleOptions: options.AnnotationGroup = "Equal": options.NormalOffset = 20#
    facade.RebarDependent "Equal", "Small", 10#, options
    Dim path As CGeometryPath: Set path = facade.LinePath("ProjectionTarget", 0#, 100#, 100#, 100#)
    Set options = New CRebarRuleOptions: options.AnnotationGroup = "Large": options.DirectionY = 1#
    facade.RebarProjected "Large", path.ID, "Small", 12#, options
    facade.RebarRowLabel "Small", "Small", 0#, -1#
    facade.RebarRowExtentLabel "Equal", "Equal", 0#, 1#
    facade.RebarRowExtremeLabel "Large", "Large", 0#, 1#, 0#, 1#
    Dim definition As CSectionDefinition: Set definition = facade.Finish
    Dim mesh As TSectionMeshOptions
    mesh.StepX = 1000#: mesh.StepY = 1000#: mesh.BoundarySubdivisions = 0: mesh.MaterialID = 1: mesh.CellGeometry = "Contour"
    mesh.SourceReference = "Mesh.MinElementArea: Config!B2"
    Dim zeroBuilder As CSectionBuilder: Set zeroBuilder = New CSectionBuilder
    Dim baseline As CSectionModel: Set baseline = zeroBuilder.Build(definition, mesh)
    TestCheck "minimumArea.contourUnusedDepth", mesh.BoundarySubdivisions = 0 And baseline.ConcreteCount > 0
    TestCheck "minimumArea.zero.count", baseline.RebarCount = 9 And zeroBuilder.Rules.FilteredByArea = 0
    Dim builder As CSectionBuilder: Set builder = New CSectionBuilder
    Dim model As CSectionModel, i As Long, owner As Long, accepted As Long, filtered As Long
    mesh.MinArea = GEOM_PI * 10# * 10# / 4#
    Set model = builder.Build(definition, mesh)
    TestCheck "minimumArea.filtered.count", model.RebarCount = 6 And builder.Rules.FilteredByArea = 3
    TestCheck "minimumArea.requestedNotDoubled", builder.Rules.RequestedPositions = 9 And builder.Rules.CreatedPositions = 6 And builder.Rules.MergedPositions = 0
    For i = 1 To builder.Rules.PositionCount
        owner = builder.Rules.PositionPhysicalOwner(i)
        If builder.Rules.PositionRuleID(i) = "Small" Then
            filtered = filtered + 1
            TestCheck "minimumArea.filtered.ownerZero." & CStr(i), owner = 0
        Else
            accepted = accepted + 1
            TestCheck "minimumArea.accepted.owner." & CStr(i), owner >= 1 And owner <= model.RebarCount
        End If
    Next i
    TestCheck "minimumArea.positionCounts", filtered = 3 And accepted = 6
    For i = 1 To model.RebarCount
        TestCheck "minimumArea.contiguousIDs." & CStr(i), model.RebarID(i) = "R" & CStr(i)
        TestCheck "minimumArea.noSmallDiameter." & CStr(i), model.RebarDiameter(i) >= 10#
        TestCheck "minimumArea.acceptedTrace." & CStr(i), Len(model.RebarRecipeTrace(i)) > 0
    Next i
    For i = 1 To 3
        TestNear "minimumArea.dependent.x." & CStr(i), model.RebarX(i), baseline.RebarX(i + 3)
        TestNear "minimumArea.dependent.y." & CStr(i), model.RebarY(i), 20#
        TestNear "minimumArea.equalThreshold." & CStr(i), model.RebarArea(i), mesh.MinArea
        TestNear "minimumArea.projected.x." & CStr(i), model.RebarX(i + 3), baseline.RebarX(i + 6)
        TestNear "minimumArea.projected.y." & CStr(i), model.RebarY(i + 3), 100#
    Next i
    TestCheck "minimumArea.annotations.acceptedOnly", TestAnnotatedBarCount(model.Annotations) = 6 And model.Annotations.Count = 2
    Dim lowX As Double, highX As Double, lowY As Double, highY As Double
    TestCheck "minimumArea.filteredBoundsAbsent", Not builder.Rules.RuleBounds("Small", lowX, highX, lowY, highY)
    Dim trace As String: trace = builder.Rules.RuleTrace("Small")
    TestCheck "minimumArea.ruleCounters", InStr(trace, "Requested=3;") > 0 And InStr(trace, "Skipped=3;") > 0 And InStr(trace, "FilteredArea=3;") > 0
    TestCheck "minimumArea.totalDiagnostics", InStr(builder.Rules.DiagnosticsText, "Requested=9;") > 0 And _
        InStr(builder.Rules.DiagnosticsText, "Skipped=3;") > 0 And InStr(builder.Rules.GeometryWarningText, "Mesh.MinElementArea") > 0
    TestCheck "minimumArea.diagnosticCodes", builder.Rules.DiagnosticCount = 3 And builder.Rules.DiagnosticAt(1).Code = "MIN_ELEMENT_AREA"
    mesh.MinArea = 0#: Set builder = New CSectionBuilder: Set model = builder.Build(definition, mesh)
    TestCheck "minimumArea.zeroRestoresAll", model.RebarCount = baseline.RebarCount And builder.Rules.FilteredByArea = 0
    For i = 1 To baseline.RebarCount
        TestNear "minimumArea.zero.x." & CStr(i), model.RebarX(i), baseline.RebarX(i)
        TestNear "minimumArea.zero.y." & CStr(i), model.RebarY(i), baseline.RebarY(i)
        TestNear "minimumArea.zero.area." & CStr(i), model.RebarArea(i), baseline.RebarArea(i)
    Next i
    TestMinimumAreaCoincidenceAndEmpty mesh
    Dim invalidBuilder As CSectionBuilder: Set invalidBuilder = New CSectionBuilder
    Dim invalidNumber As Long
    mesh.CellGeometry = "Adaptive": mesh.MinArea = 0#
    On Error Resume Next
    Set model = invalidBuilder.Build(definition, mesh)
    invalidNumber = Err.Number: Err.Clear
    On Error GoTo 0
    TestCheck "minimumArea.adaptiveDepthStillRequired", invalidNumber = vbObjectError + 5530
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Отфильтрованный стержень не создаёт ложный конфликт в совпавшей позиции.
' Если исключены все стержни, сборщик сообщает общий порог и переданный адрес настройки.
Private Sub TestMinimumAreaCoincidenceAndEmpty(ByRef mesh As TSectionMeshOptions)
    Dim facade As CSectionDefinitionBuilder: Set facade = TestRuleFacade
    facade.RebarByCount "Small", "Row", 2, 8#
    facade.RebarByCount "Kept", "Row", 2, 10#
    mesh.MinArea = GEOM_PI * 10# * 10# / 4#
    Dim builder As CSectionBuilder: Set builder = New CSectionBuilder
    Dim model As CSectionModel: Set model = builder.Build(facade.Finish, mesh)
    TestCheck "minimumArea.coincidence.filteredBeforeSearch", model.RebarCount = 2 And builder.Rules.FilteredByArea = 2 And builder.Rules.Conflicts = 0
    TestNear "minimumArea.coincidence.keptDiameter", model.RebarDiameter(1), 10#
    Set facade = TestRuleFacade: facade.RebarByCount "OnlySmall", "Row", 2, 8#
    Dim definition As CSectionDefinition: Set definition = facade.Finish
    Dim number As Long, description As String
    Set builder = New CSectionBuilder
    On Error Resume Next
    Set model = builder.Build(definition, mesh)
    number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo 0
    TestCheck "minimumArea.empty.rejected", number = vbObjectError + 5530
    TestCheck "minimumArea.empty.address", InStr(description, "Mesh.MinElementArea") > 0 And InStr(description, "Config!B2") > 0 And InStr(description, "не осталось") > 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет выбор частей размерами, игнорирование неактивного армирования и общие скругления.
Public Function RunRectanglesDimensionTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    TestRectanglesDimensionsAndRadius
    GoTo Finished
Failed:
    TestCheck "rectDimensions.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRectanglesDimensionTests = mReport & "TOTAL_RECTANGLES_DIMENSIONS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Для каждого прямоугольника отдельно проверяет B=0, H=0 и оба нуля при нескольких радиусах.
' CVErr, текст и отрицательные значения во всех полях отсутствующей арматуры не должны читаться.
Private Sub TestRectanglesDimensionsAndRadius()
    Dim values As Variant, changed As Variant, recipe As ISectionRecipe, definition As CSectionDefinition
    Dim layout As CRebarLayout, reference As CRebarLayout, query As CSectionGeometryQuery
    Dim radius As Variant, absent As Long, zeroKind As Long, block As Variant, key As Variant, field As Variant, bad As Variant
    Dim expectedArea As Double, width As Double, height As Double, i As Long, code As Long, reason As String
    Dim table As CSectionInputReader, inputTable As CSectionInputTable, prefix As String, path As CGeometryPath
    Set recipe = New CSectionRecipeRectangles: Set query = New CSectionGeometryQuery
    For Each radius In Array(0#, 50#, 200#)
        For absent = 0 To 2
            For zeroKind = 0 To IIf(absent = 0, 0, 2)
                values = TestRectanglesDimensionsValues(CDbl(radius))
                If absent > 0 Then
                    If zeroKind <> 1 Then TestFixtureSetInput values, "General", "B" & CStr(absent), "val", 0#
                    If zeroKind <> 0 Then TestFixtureSetInput values, "General", "H" & CStr(absent), "val", 0#
                    TestFixtureSetInput values, "General", "XOffset", "val", CVErr(2015)
                End If
                Set definition = TestFixtureRecipeDefinition(recipe, "rngRectanglesGeometry", values)
                Set reference = TestFixtureAnnotatedLayout(definition)
                width = 800#: height = 1000#
                If absent = 1 Then width = 1200#: height = 600#
                expectedArea = width * height - (4# - GEOM_PI) * CDbl(radius) ^ 2
                If absent = 0 Then expectedArea = 800# * 1000# + 1200# * 600# - 2# * (4# - GEOM_PI) * CDbl(radius) ^ 2
                prefix = "rectDimensions." & CStr(radius) & "." & CStr(absent) & "." & CStr(zeroKind)
                TestNear prefix & ".area", definition.ConcreteDomain.Area, expectedArea
                TestCheck prefix & ".rules", definition.RuleCount = IIf(absent = 0, 8, 4)
                For i = 1 To reference.Count
                    TestCheck prefix & ".bar.inside." & CStr(i), query.ContainsPoint(definition.ConcreteDomain, reference.X(i), reference.Y(i))
                Next i
                If absent > 0 Then
                    Set path = definition.Path(IIf(absent = 1, "Lower", "Upper"))
                    path.ProjectionBounds 1#, 0#, width, height
                    TestNear prefix & ".single.minX", width, 0#
                    path.ProjectionBounds 0#, 1#, width, height
                    TestNear prefix & ".single.minY", width, 0#
                    Set table = TestFixtureInput(values)
                    For Each bad In Array("unused text", -1#, CVErr(2015))
                        changed = values
                        For Each block In Array("Main", "Extra")
                            For Each key In table.RowKeys(CStr(block))
                                If Left$(CStr(key), 3) = "R" & CStr(absent) & "." Then
                                    If CStr(block) = "Main" Then
                                        For Each field In Array("as", "d", "n", "s0", "s1", "endMode"): TestFixtureSetInput changed, CStr(block), CStr(key), CStr(field), bad: Next field
                                    Else
                                        For Each field In Array("d2", "a2", "bind2", "pos2", "d3", "a3", "bind3", "pos3"): TestFixtureSetInput changed, CStr(block), CStr(key), CStr(field), bad: Next field
                                    End If
                                End If
                            Next key
                        Next block
                        Set definition = TestFixtureRecipeDefinition(recipe, "rngRectanglesGeometry", changed)
                        Set layout = TestFixtureAnnotatedLayout(definition)
                        TestCheck prefix & ".unused.count", layout.Count = reference.Count
                        For i = 1 To layout.Count
                            TestNear prefix & ".unused.x." & CStr(i), layout.X(i), reference.X(i)
                            TestNear prefix & ".unused.y." & CStr(i), layout.Y(i), reference.Y(i)
                            TestNear prefix & ".unused.d." & CStr(i), layout.Diameter(i), reference.Diameter(i)
                        Next i
                    Next bad
                End If
            Next zeroKind
        Next absent
    Next radius
    For Each key In Array("B1", "H1", "B2", "H2", "R")
        For Each bad In Array(-1#, "invalid", CVErr(2015), Empty)
            values = TestRectanglesDimensionsValues(0#)
            TestFixtureSetInput values, "General", CStr(key), "val", bad
            code = 0: reason = ""
            On Error Resume Next
            Set definition = TestFixtureRecipeDefinition(recipe, "rngRectanglesGeometry", values)
            code = Err.Number: reason = Err.Description: Err.Clear
            On Error GoTo 0
            Set inputTable = New CSectionInputTable: inputTable.Initialize "rngRectanglesGeometry", values
            Set table = New CSectionInputReader: table.Initialize inputTable
            mReport = mReport & "RECTANGLE_INVALID_INPUT: " & CStr(key) & "; " & reason & vbCrLf
            TestCheck "rectDimensions.invalid." & CStr(key), code <> 0 And InStr(reason, table.SourceReference("General", CStr(key))) > 0
        Next bad
    Next key
    For Each radius In Array(0#, 50#)
        values = TestRectanglesDimensionsValues(CDbl(radius))
        TestFixtureSetInput values, "General", "B1", "val", 0#: TestFixtureSetInput values, "General", "H2", "val", 0#
        code = 0: reason = ""
        On Error Resume Next
        Set definition = TestFixtureRecipeDefinition(recipe, "rngRectanglesGeometry", values)
        code = Err.Number: reason = Err.Description: Err.Clear
        On Error GoTo 0
        TestCheck "rectDimensions.both.absent", code <> 0 And InStr(reason, "Не задано ни одного прямоугольника") > 0
    Next radius
    TestRectanglesHollowAxisParity
    TestJoinedRoundedPaths
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Готовит только входные данные; геометрию, арматуру и дуги строят рабочие общие классы.
Private Function TestRectanglesDimensionsValues(ByVal radius As Double) As Variant
    Dim values As Variant, reader As CSectionInputReader, key As Variant
    values = TestFixtureRecipeValues("rngRectanglesGeometry"): Set reader = TestFixtureInput(values)
    TestFixtureSetInput values, "General", "B1", "val", 800#: TestFixtureSetInput values, "General", "H1", "val", 1000#
    TestFixtureSetInput values, "General", "B2", "val", 1200#: TestFixtureSetInput values, "General", "H2", "val", 600#
    TestFixtureSetInput values, "General", "XOffset", "val", 50#: TestFixtureSetInput values, "General", "R", "val", radius
    For Each key In reader.RowKeys("Main")
        TestFixtureSetInput values, "Main", CStr(key), "as", 40#: TestFixtureSetInput values, "Main", CStr(key), "d", 20#
        TestFixtureSetInput values, "Main", CStr(key), "n", 5: TestFixtureSetInput values, "Main", CStr(key), "endMode", "IncludeBoth"
        TestFixtureSetInput values, "Main", CStr(key), "s0", 10#: TestFixtureSetInput values, "Main", CStr(key), "s1", 20#
        TestFixtureSetInput values, "Extra", CStr(key), "d2", 0#: TestFixtureSetInput values, "Extra", CStr(key), "d3", 0#
    Next key
    TestRectanglesDimensionsValues = values
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает все четыре пути с наружной гранью HollowRectangle, включая дуги, направления и разные as.
Private Sub TestRectanglesHollowAxisParity()
    Dim rectValues As Variant, hollowValues As Variant, rect As CSectionDefinition, hollow As CSectionDefinition
    Dim recipe As ISectionRecipe, sides As Variant, rectIDs As Variant, hollowIDs As Variant, offsets As Variant
    Dim i As Long, j As Long, p As CGeometryPath, q As CGeometryPath, x As Double, y As Double, tx As Double, ty As Double
    Dim hx As Double, hy As Double, htx As Double, hty As Double
    rectValues = TestRectanglesDimensionsValues(180#): TestFixtureSetInput rectValues, "General", "B2", "val", 0#
    hollowValues = TestFixtureRecipeValues("rngHollowRectangleGeometry")
    TestFixtureSetInput hollowValues, "General", "B", "val", 800#: TestFixtureSetInput hollowValues, "General", "H", "val", 1000#
    TestFixtureSetInput hollowValues, "General", "R", "val", 180#
    TestFixtureSetInput hollowValues, "General", "Bo", "val", 100#: TestFixtureSetInput hollowValues, "General", "Ho", "val", 200#
    TestFixtureSetInput hollowValues, "General", "Ro", "val", 0#: TestFixtureSetInput hollowValues, "General", "Xo", "val", 0#: TestFixtureSetInput hollowValues, "General", "Yo", "val", 0#
    sides = Array("Top", "Bottom", "Left", "Right"): offsets = Array(40#, 50#, 60#, 70#)
    rectIDs = Array("B1.as_1", "B1.as_2", "H1.as_1", "H1.as_2"): hollowIDs = Array("B.Top", "B.Bottom", "H.Left", "H.Right")
    For i = 0 To 3
        TestFixtureSetInput rectValues, "Main", "R1." & CStr(sides(i)), "as", offsets(i)
        TestFixtureSetInput hollowValues, "Main", CStr(sides(i)), "as", offsets(i)
        TestFixtureSetInput hollowValues, "Main", CStr(sides(i)), "d", 20#: TestFixtureSetInput hollowValues, "Main", CStr(sides(i)), "n", 5
        TestFixtureSetInput hollowValues, "Main", "Opening" & CStr(sides(i)), "d", 0#
        TestFixtureSetInput hollowValues, "Extra", CStr(sides(i)), "d2", 0#: TestFixtureSetInput hollowValues, "Extra", CStr(sides(i)), "d3", 0#
    Next i
    Set recipe = New CSectionRecipeRectangles: Set rect = TestFixtureRecipeDefinition(recipe, "rngRectanglesGeometry", rectValues)
    Set recipe = New CSectionRecipeHollowRectangle: Set hollow = TestFixtureRecipeDefinition(recipe, "rngHollowRectangleGeometry", hollowValues)
    For i = 0 To 3
        Set p = rect.Path(CStr(rectIDs(i))): Set q = hollow.Path(CStr(hollowIDs(i)))
        TestNear "rectDimensions.hollow.length." & CStr(i), p.Length, q.Length
        For j = 0 To 10
            p.AtDistance p.Length * j / 10#, x, y, tx, ty: q.AtDistance q.Length * j / 10#, hx, hy, htx, hty
            TestNear "rectDimensions.hollow.x", x, hx + 400#: TestNear "rectDimensions.hollow.y", y, hy + 500#
            TestNear "rectDimensions.hollow.tx", tx, htx: TestNear "rectDimensions.hollow.ty", ty, hty
        Next j
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет объединение в фасаде и ядре с аналитической площадью и отклонение раздельных частей.
Private Sub TestJoinedRoundedPaths()
    Dim facade As CSectionDefinitionBuilder, first As CGeometryPath, second As CGeometryPath, joined As CGeometryPath
    Dim paths As Collection, query As CSectionGeometryQuery, polygon As CPolygonSpec, code As Long, reason As String
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "TestJoinedPaths"
    Set first = facade.RectanglePath("A", 0#, 0#, 300#, 200#)
    Set second = facade.RectanglePath("B", 100#, 200#, 300#, 200#)
    Set joined = facade.JoinedPaths("AB", Array("A", "B"))
    Set query = New CSectionGeometryQuery
    TestNear "joined.facade.area", query.RegionFromPath(joined).Area, 120000#
    Set paths = New Collection: paths.Add first: paths.Add second
    Set joined = query.JoinedPaths("AB.Query", paths)
    TestNear "joined.query.area", query.RegionFromPath(joined).Area, 120000#
    Set paths = New Collection: paths.Add first
    Set second = facade.RectanglePath("Separated", 500#, 200#, 300#, 200#): paths.Add second
    On Error Resume Next
    Set joined = query.JoinedPaths("Invalid", paths)
    code = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo 0
    TestCheck "joined.disconnected.rejected", code <> 0 And InStr(reason, "связную") > 0
End Sub
