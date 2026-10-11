Attribute VB_Name = "modTestSectionRecipes"
Option Explicit

Private mPassed As Long, mFailed As Long, mReport As String

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
    GoTo Finished
Failed:
    Check "rules.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRecipeRuleTests = mReport & "TOTAL RECIPE RULES: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

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
    Dim reader As CSectionInputSnapshotReader, context As CSectionRecipeContext, recipe As ISectionRecipe
    Dim definition As CSectionDefinition, mesh As TSectionMeshOptions
    Dim registry As CSectionTypeRegistry
    Set registry = New CSectionTypeRegistry
    Set reader = New CSectionInputSnapshotReader: Set context = reader.Capture(ThisWorkbook, registry.RecipeTableNames("Circle"), units)
    Set recipe = registry.CreateRecipe("Circle"): Set definition = recipe.CreateDefinition(context)
    mesh.StepX = units.InputLengthToInternal(settings.GetRequiredDouble("Mesh.StepX"))
    mesh.StepY = units.InputLengthToInternal(settings.GetRequiredDouble("Mesh.StepY"))
    mesh.BoundarySubdivisions = ReadMeshBoundarySubdivisions(settings): mesh.MaterialID = 1
    Set BuildCircleRecipe = builder.Build(definition, mesh)
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
    Dim config As Object, saved As Variant, result As String, settings As CSystemSettingsReader, units As CUnitSystem
    Dim profiles As Object, i As Long, key As String, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean
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
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set builder = New CSectionBuilder: Set model = BuildCircleRecipe(settings, units, builder)
    result = modTestShapeArchitecture.PreparedCalculationSnapshot("Circle", model)
    GoTo Finished
Failed:
    result = result & "FAIL: Circle recipe calculations; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    DumpCircleRecipeCalculations = result
End Function

' ДЛЯ ТЕСТОВ: настоящий Excel PNG и структурный снимок схемы.
Public Function DumpCircleRecipePreview(ByVal outputDirectory As String) As String
    Dim config As Object, saved As Variant, result As String, settings As CSystemSettingsReader, units As CUnitSystem
    Dim builder As CSectionBuilder, model As CSectionModel, calculation As Long, events As Boolean, updating As Boolean
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = True
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    modTestShapeArchitecture.ConfigureShapeFixture "Circle"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set builder = New CSectionBuilder: Set model = BuildCircleRecipe(settings, units, builder)
    result = modTestShapeArchitecture.PreparedPreviewSnapshot("Circle", model, outputDirectory)
    GoTo Finished
Failed:
    result = result & "FAIL: Circle recipe preview; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    DumpCircleRecipePreview = result
End Function

Private Sub SetSystemInput(ByVal key As String, ByVal value As Variant)
    Dim table As Object, i As Long
    Set table = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    For i = 1 To table.Rows.Count
        If CStr(table.Cells(i, 1).Value2) = key Then table.Cells(i, 2).Value2 = value: Exit Sub
    Next i
    Err.Raise vbObjectError + 5531, "modTestSectionRecipes", "Не найден системный ключ '" & key & "'."
End Sub
