Attribute VB_Name = "modTestRecipeSetbacks"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: редактируемые отступы таблиц листа «Сечения» и их влияние на общий
' рецепт, включая фильтр проекций Opening. НДС в этом наборе не решается.
' ==========================================================================
Private mPassed As Long, mFailed As Long, mReport As String

' -------------------- ДЛЯ ТЕСТОВ --------------------
' те же внутренние координаты при вводе отступов в мм/см/м.
Public Function RunRecipeSetbackUnitTests() As String
    Dim values As Variant, baseline As Variant, scaled As Variant, table As CSectionInputReader
    Dim sheet As Object, setting As Object, settings As CSystemSettingsReader, units As CUnitSystem
    Dim recipe As ISectionRecipe, definition As CSectionDefinition, expected As CRebarLayout, actual As CRebarLayout
    Dim main As Long, extra As Long, row As Long, column As Long, i As Long, item As Variant, factor As Double
    Dim alerts As Boolean, itemColumn As Variant
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = "": alerts = Application.DisplayAlerts
    values = TestFixtureRecipeValues("rngPointedSectionGeometry")
    TestFixtureSetInput values, "General", "B", "val", 800#: TestFixtureSetInput values, "General", "H", "val", 250#
    TestFixtureSetInput values, "Detail", "W", "left", 0#: TestFixtureSetInput values, "Detail", "W", "right", 150#
    TestFixtureSetInput values, "Detail", "R1", "left", 40#: TestFixtureSetInput values, "Detail", "R1", "right", 40#
    TestFixtureSetInput values, "Detail", "R2", "left", 0#: TestFixtureSetInput values, "Detail", "R2", "right", 80#
    Dim key As Variant, block As Variant, fields As Variant, field As Variant, rowKeys As Variant
    Set table = TestFixtureInput(values)
    For Each key In table.RowKeys("Main")
        TestFixtureSetInput values, "Main", CStr(key), "as", 40#: TestFixtureSetInput values, "Main", CStr(key), "d", 20#
        TestFixtureSetInput values, "Main", CStr(key), "n", 3
        TestFixtureSetInput values, "Main", CStr(key), "s0", 10#: TestFixtureSetInput values, "Main", CStr(key), "s1", 20#
        TestFixtureSetInput values, "Extra", CStr(key), "d2", 0#: TestFixtureSetInput values, "Extra", CStr(key), "d3", 0#
    Next key
    Set recipe = New CSectionRecipePointedSection
    Set expected = TestFixtureAnnotatedLayout(TestFixtureRecipeDefinition(recipe, "rngPointedSectionGeometry", values))
    baseline = values
    Set sheet = ThisWorkbook.Worksheets.Add: Set setting = sheet.Range("A1:C2")
    setting.Cells(1, 1).Value2 = "Key": setting.Cells(1, 2).Value2 = "Value": setting.Cells(1, 3).Value2 = "Unit"
    setting.Cells(2, 1).Value2 = "Units.Length.Input"
    For Each item In Array(Array("mm", 1#), Array("cm", 10#), Array("m", 1000#))
        factor = CDbl(item(1)): scaled = baseline
        For Each block In Array("General", "Main", "Extra")
            Select Case CStr(block)
                Case "General": fields = Array("val")
                Case "Detail": fields = Array("left", "right")
                Case "Main": fields = Array("as", "d", "s0", "s1")
                Case "Extra": fields = Array("d2", "a2", "d3", "a3")
            End Select
            For Each key In table.RowKeys(CStr(block))
                If CStr(block) <> "Detail" Or CStr(key) <> "Type" Then
                    row = table.RowIndex(CStr(block), CStr(key))
                    For Each field In fields
                        column = table.ColumnIndex(CStr(block), CStr(field))
                        If IsNumeric(baseline(row, column)) Then scaled(row, column) = CDbl(baseline(row, column)) / factor
                    Next field
                End If
            Next key
        Next block
        setting.Cells(2, 2).Value2 = CStr(item(0))
        Set settings = New CSystemSettingsReader: settings.LoadFromRange setting
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set definition = TestFixtureRecipeDefinition(recipe, "rngPointedSectionGeometry", scaled, units)
        Set actual = TestFixtureAnnotatedLayout(definition)
        TestCheck "setback.units.count." & CStr(item(0)), actual.Count = expected.Count
        For i = 1 To expected.Count
            TestNear "setback.units.x." & CStr(item(0)) & "." & CStr(i), actual.X(i), expected.X(i)
            TestNear "setback.units.y." & CStr(item(0)) & "." & CStr(i), actual.Y(i), expected.Y(i)
        Next i
    Next item
    mReport = mReport & "TOTAL_RECIPE_SETBACK_UNITS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
Restore:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    Application.DisplayAlerts = alerts
    On Error GoTo 0
    RunRecipeSetbackUnitTests = mReport
    Exit Function
Failed:
    mFailed = mFailed + 1: mReport = mReport & "FAIL: units.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
    Resume Restore
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Обновление реестра меняет только встроенный список, без служебных ячеек и таблиц геометрии.
Public Function RunRecipeTypeListTests() As String
    Dim config As Object, system As Object, typeKey As Object, typeCell As Object, saved As Variant, expected As Variant
    Dim ids As Variant, choices As String, originalValidation As String, originalCircle As String, brokenCircle As String, originalType As Variant, currentValidation As String
    Dim i As Long, count As Long, code As Long, absent As Boolean, unused As Object
    Dim calculation As Long, events As Boolean, updating As Boolean, alerts As Boolean
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = ""
    calculation = Application.Calculation: events = Application.EnableEvents
    updating = Application.ScreenUpdating: alerts = Application.DisplayAlerts
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    Set system = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    For i = 2 To system.Rows.Count
        If CStr(system.Cells(i, 1).Value2) = "Geometry.Type" Then Set typeKey = system.Cells(i, 1): Set typeCell = system.Cells(i, 2)
    Next i
    If typeCell Is Nothing Then Err.Raise 5, "RunRecipeTypeListTests", "Geometry.Type отсутствует в тестовой книге."
    originalValidation = typeCell.Validation.Formula1: originalType = typeCell.Formula
    originalCircle = ThisWorkbook.Names("rngCircleGeometry").RefersTo
    ids = RegisteredSectionRecipeTypes: count = UBound(ids) - LBound(ids) + 1
    choices = Join(ids, ",")
    Application.Calculation = xlCalculationManual: Application.EnableEvents = False
    Application.ScreenUpdating = False: Application.DisplayAlerts = False
    ThisWorkbook.Names("rngCircleGeometry").RefersTo = "=#REF!" ' Повреждённая таблица геометрии не нужна обновлению списка.
    brokenCircle = ThisWorkbook.Names("rngCircleGeometry").RefersTo
    typeCell.Validation.Delete: typeCell.Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "stale choice"
    expected = config.Formula: modSectionRecipeSetup.UpdateSectionRecipeTypes
    TestCheck "types.config.allValuesPreserved", TestSameConfigOutsideTypeList(config, expected, Nothing)
    TestCheck "types.current.selection", TestSameConfigOutsideTypeList(config, saved(0), Nothing)
    TestCheck "types.sectionInputsPreserved", TestSameConfigOutsideTypeList(ThisWorkbook.Worksheets("Сечения").Range(CStr(saved(1))), saved(2), Nothing)
    TestCheck "types.unrelated.name", ThisWorkbook.Names("rngCircleGeometry").RefersTo = brokenCircle
    TestCheck "types.inline", Left$(typeCell.Validation.Formula1, 1) <> "="
    currentValidation = typeCell.Validation.Formula1 ' Excel может локализовать строку при чтении; отдельные пункты проверяются его собственным Validation.Value.
    TestCheck "types.dropdown", typeCell.Validation.Type = xlValidateList And typeCell.Validation.InCellDropdown
    For i = LBound(ids) To UBound(ids)
        typeCell.Value2 = ids(i)
        TestCheck "types.option.accepted." & CStr(ids(i)), typeCell.Validation.Value ' Отдельные значения обязан принимать сам Excel, а не тестовый разбор строки.
    Next i
    typeCell.Value2 = Join(ids, ";")
    TestCheck "types.joinedValue.rejected", Not typeCell.Validation.Value
    typeCell.Value2 = "TEST_UNKNOWN_GEOMETRY_TYPE"
    TestCheck "types.unknownValue.rejected", Not typeCell.Validation.Value
    typeCell.Formula = originalType
    On Error Resume Next
    Set unused = ThisWorkbook.Names("rngSectionRecipeTypes"): absent = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    TestCheck "types.noHelperName", absent
    TestCheck "types.state.success", Application.Calculation = xlCalculationManual And Not Application.EnableEvents And Not Application.ScreenUpdating And Not Application.DisplayAlerts
    expected = config.Formula: modSectionRecipeSetup.UpdateSectionRecipeTypes
    TestCheck "types.idempotent", TestSameConfigOutsideTypeList(config, expected, Nothing) And typeCell.Validation.Formula1 = currentValidation
    typeKey.Value2 = "Missing.Geometry.Type": expected = config.Formula
    On Error Resume Next
    modSectionRecipeSetup.UpdateSectionRecipeTypes: code = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "types.missing.rejected", code <> 0
    TestCheck "types.missing.data", TestSameConfigOutsideTypeList(config, expected, Nothing)
    TestCheck "types.state.failure", Application.Calculation = xlCalculationManual And Not Application.EnableEvents And Not Application.ScreenUpdating And Not Application.DisplayAlerts
    GoTo Restore
Failed:
    TestCheck "types.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Restore:
    On Error Resume Next
    If Len(originalCircle) > 0 Then ThisWorkbook.Names("rngCircleGeometry").RefersTo = originalCircle
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    If Not typeCell Is Nothing And Len(originalValidation) > 0 Then
        typeCell.Validation.Delete: typeCell.Validation.Add xlValidateList, xlValidAlertStop, xlBetween, originalValidation
        typeCell.Validation.IgnoreBlank = False: typeCell.Validation.InCellDropdown = True
    End If
    Application.Calculation = calculation: Application.EnableEvents = events
    Application.ScreenUpdating = updating: Application.DisplayAlerts = alerts
    On Error GoTo 0
    RunRecipeTypeListTests = mReport & "TOTAL RECIPE TYPE LIST: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает весь Config, кроме явно разрешённого служебного списка.
Private Function TestSameConfigOutsideTypeList(ByVal config As Object, ByVal expected As Variant, ByVal list As Object) As Boolean
    Dim actual As Variant, row As Long, column As Long, sheetRow As Long, sheetColumn As Long, skip As Boolean
    actual = config.Formula
    For row = 1 To UBound(expected, 1)
        For column = 1 To UBound(expected, 2)
            skip = False
            If Not list Is Nothing Then
                sheetRow = config.Row + row - 1: sheetColumn = config.Column + column - 1
                skip = sheetColumn = list.Column And sheetRow >= list.Row And sheetRow < list.Row + list.Rows.Count
            End If
            If Not skip Then
                If VarType(actual(row, column)) <> VarType(expected(row, column)) Then Exit Function
                If CStr(actual(row, column)) <> CStr(expected(row, column)) Then Exit Function
            End If
        Next column
    Next row
    TestSameConfigOutsideTypeList = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет влияние редактируемых концевых отступов PointedSection и HollowRectangle на позиции, проекции, адреса ошибок и неизменность бетона.
Public Function RunRecipeSetbackInputTests() As String
    Dim rounded As Object, hollow As Object, savedRounded As Variant, savedHollow As Variant
    Dim table As CSectionInputReader, row As Long, main As Long, def As CSectionDefinition
    Dim original As CRebarLayout, changed As CRebarLayout, region As CConcreteRegion
    Dim face As Variant, faces As Variant, helper As CRebarRuleExecutor, originalHelper As CRebarRuleExecutor, parameters As Variant
    Dim geometryArea As Double, layoutCount As Long, reference As String, reason As String, number As Long
    Dim config As Object, savedConfig As Variant
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: savedConfig = TestFixtureSaveInputs(config)
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = ""
    modTestShapeArchitecture.TestConfigureShapeFixture "RoundedMixed"
    Set rounded = ThisWorkbook.Names("rngPointedSectionGeometry").RefersToRange
    Set hollow = ThisWorkbook.Names("rngHollowRectangleGeometry").RefersToRange
    savedRounded = rounded.Formula: savedHollow = hollow.Formula
    Set table = TestFixtureInput(rounded.Value2)
    row = table.RowIndex("Main", "Top"): main = row - 1
    TestCheck "rounded.start.header", table.ColumnIndex("Main", "s0") = 6
    TestCheck "rounded.end.header", table.ColumnIndex("Main", "s1") = 7
    TestCheck "rounded.editable.centered", rounded.Cells(row, 6).HorizontalAlignment = xlCenter And rounded.Cells(row, 7).HorizontalAlignment = xlCenter
    face = Array(40#, 20#, 3, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")
    faces = Array(face, face, Array(40#, 20#, 5, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), face)
    parameters = Array(800#, 250#, 0#, 150#, 40#, 40#, 0#, 80#)
    rounded.Cells(main + 1, 6).Resize(4, 2).Value2 = 0#
    Set def = TestFixtureRoundedDefinition(parameters, faces, True): Set original = TestFixtureAnnotatedLayout(def)
    geometryArea = def.ConcreteDomain.Area: layoutCount = original.Count
    TestNear "rounded.zero.first", original.X(1), -360#
    rounded.Cells(row, 6).Value2 = 10#: rounded.Cells(row, 7).Value2 = 20#
    Set def = TestFixtureRoundedDefinition(parameters, faces, True): Set changed = TestFixtureAnnotatedLayout(def)
    TestNear "rounded.start.actual", changed.X(1), original.X(1) + 10#
    TestNear "rounded.end.actual", changed.X(5), original.X(5) - 20#
    TestNear "rounded.concrete.unchanged", def.ConcreteDomain.Area, geometryArea
    TestCheck "rounded.count.unchanged", changed.Count = layoutCount
    TestNear "rounded.opposite.row.unchanged", changed.Y(6), original.Y(6)
    rounded.Cells(row, 6).ClearContents
    Set changed = Nothing: number = 0
    On Error Resume Next
    Set changed = TestFixtureAnnotatedLayout(TestFixtureRoundedDefinition(parameters, faces, True))
    number = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    TestCheck "rounded.active.empty.isZero", number = 0 And Not changed Is Nothing
    If Not changed Is Nothing Then TestNear "rounded.active.empty.coordinate", changed.X(1), original.X(1)
    rounded.Cells(row, 6).Value2 = -1#
    On Error Resume Next
    Set changed = TestFixtureAnnotatedLayout(TestFixtureRoundedDefinition(parameters, faces, True))
    number = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    TestCheck "rounded.active.negative.rejected", number <> 0 And InStr(1, reason, "[" & CStr(row) & ",6]", vbBinaryCompare) > 0
    rounded.Cells(row, 6).Value2 = CVErr(xlErrValue)
    faces(2) = Array(40#, 0#, 0, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")
    Set changed = TestFixtureAnnotatedLayout(TestFixtureRoundedDefinition(parameters, faces, True))
    TestCheck "rounded.inactive.error.ignored", changed.Count = layoutCount - 5
    rounded.Formula = savedRounded

    modTestShapeArchitecture.TestConfigureShapeFixture "HollowCentered"
    Set table = TestFixtureInput(hollow.Value2)
    main = table.RowIndex("Main", "Top") - 1
    row = table.RowIndex("Main", "OpeningTop")
    TestCheck "hollow.start.header", table.ColumnIndex("Main", "s0") = 6
    TestCheck "hollow.end.header", table.ColumnIndex("Main", "s1") = 7
    hollow.Cells(main + 1, 6).Resize(8, 2).Value2 = 0#
    face = Array(40#, 20#, 6, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")
    faces = Array(face, face, Array(40#, 20#, 8, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), _
        Array(40#, 20#, 8, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), face, face, face, face)
    parameters = Array(500#, 800#, 0#, 200#, 500#, 0#)
    Set originalHelper = New CRebarRuleExecutor
    Set def = TestFixtureHollowDefinition(parameters, faces, True): Set original = TestFixtureAnnotatedLayout(def, originalHelper)
    geometryArea = def.ConcreteDomain.Area
    TestCheck "hollow.zero.projected.count", TestGroupCount(original, "Opening.B.Top") = 6
    hollow.Cells(row, 6).Value2 = 75#: hollow.Cells(row, 7).Value2 = 15#
    Set helper = New CRebarRuleExecutor
    Set def = TestFixtureHollowDefinition(parameters, faces, True): Set changed = TestFixtureAnnotatedLayout(def, helper)
    TestCheck "hollow.projection.filtered", TestGroupCount(changed, "Opening.B.Top") = 4
    TestCheck "hollow.skips.diagnostic", helper.SkippedProjections - originalHelper.SkippedProjections = 2
    TestCheck "hollow.positions.not.redistributed", TestProjectedSubset(changed, original, "Opening.B.Top")
    TestCheck "hollow.other.projections.unchanged", TestGroupCount(changed, "Opening.B.Bottom") = TestGroupCount(original, "Opening.B.Bottom")
    TestNear "hollow.concrete.unchanged", def.ConcreteDomain.Area, geometryArea
    hollow.Formula = savedHollow
    TestUnifiedRebarPaths
    mReport = mReport & "TOTAL_RECIPE_SETBACK_INPUT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
Restore:
    On Error Resume Next
    If Not rounded Is Nothing Then rounded.Formula = savedRounded
    If Not hollow Is Nothing Then hollow.Formula = savedHollow
    If Not config Is Nothing Then TestFixtureRestoreInputs config, savedConfig
    On Error GoTo 0
    RunRecipeSetbackInputTests = mReport
    Exit Function
Failed:
    mFailed = mFailed + 1: mReport = mReport & "FAIL: runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
    Resume Restore
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет общий путь по аналитическим координатам; источник — реальные таблицы рецептов, без расчёта НДС.
Private Sub TestUnifiedRebarPaths()
    Dim values As Variant, table As CSectionInputReader, key As Variant, recipe As ISectionRecipe, definition As CSectionDefinition
    Dim path As CGeometryPath, x As Double, y As Double, tx As Double, ty As Double, reference As String
    Dim facade As CSectionDefinitionBuilder, contour As CGeometryPath, options As CRebarRuleOptions, layout As CRebarLayout
    Dim i As Long, minX As Double, maxX As Double, code As Long, reason As String
    values = TestFixtureRecipeValues("rngRectanglesGeometry")
    Set table = TestFixtureInput(values)
    TestFixtureSetInput values, "General", "Shape", "val", "OneRect"
    TestFixtureSetInput values, "General", "B1", "val", 1000#: TestFixtureSetInput values, "General", "H1", "val", 600#
    For Each key In table.RowKeys("Main")
        TestFixtureSetInput values, "Main", CStr(key), "n", 0: TestFixtureSetInput values, "Main", CStr(key), "d", 0#
        TestFixtureSetInput values, "Main", CStr(key), "as", Empty
        TestFixtureSetInput values, "Main", CStr(key), "s0", Empty: TestFixtureSetInput values, "Main", CStr(key), "s1", Empty
        TestFixtureSetInput values, "Extra", CStr(key), "d2", 0#: TestFixtureSetInput values, "Extra", CStr(key), "d3", 0#
    Next key
    TestFixtureSetInput values, "Main", "R1.Left", "as", 30#: TestFixtureSetInput values, "Main", "R1.Right", "as", 70#
    TestFixtureSetInput values, "Main", "R1.Bottom", "as", 90#: TestFixtureSetInput values, "Main", "R1.Top", "as", 50#
    TestFixtureSetInput values, "Main", "R1.Top", "n", 3: TestFixtureSetInput values, "Main", "R1.Top", "d", 20#
    TestFixtureSetInput values, "Main", "R1.Top", "s0", 10#: TestFixtureSetInput values, "Main", "R1.Top", "s1", 20#
    Set recipe = New CSectionRecipeRectangles
    Set definition = TestFixtureRecipeDefinition(recipe, "rngRectanglesGeometry", values)
    Set path = definition.Path("B1.as_1")
    path.AtDistance 0#, x, y, tx, ty: TestNear "unified.rectangle.start.x", x, 30#: TestNear "unified.rectangle.start.y", y, 550#
    path.AtDistance path.Length, x, y, tx, ty: TestNear "unified.rectangle.end.x", x, 930#
    Set layout = TestFixtureAnnotatedLayout(definition): minX = 1E+30: maxX = -1E+30
    For i = 1 To layout.Count
        minX = GeomMin(minX, layout.X(i)): maxX = GeomMax(maxX, layout.X(i))
        TestNear "unified.rectangle.bar.y." & CStr(i), layout.Y(i), 550#
    Next i
    TestCheck "unified.rectangle.count", layout.Count = 3
    TestNear "unified.rectangle.s0", minX, 40#: TestNear "unified.rectangle.s1", maxX, 910#
    TestFixtureSetInput values, "Main", "R1.Right", "as", Empty
    Set definition = TestFixtureRecipeDefinition(recipe, "rngRectanglesGeometry", values)
    Set path = definition.Path("B1.as_1"): path.AtDistance path.Length, x, y, tx, ty
    TestNear "unified.disabled.blank.as.zero", x, 1000#
    TestFixtureSetInput values, "Main", "R1.Top", "s0", Empty: TestFixtureSetInput values, "Main", "R1.Top", "s1", Empty
    Set layout = TestFixtureAnnotatedLayout(TestFixtureRecipeDefinition(recipe, "rngRectanglesGeometry", values))
    TestCheck "unified.rectangle.blank.s.accepted", layout.Count = 3
    For Each key In Array(-1#, "not a length", CVErr(xlErrValue))
        TestFixtureSetInput values, "Main", "R1.Top", "s0", key
        code = 0: reason = "": Set definition = Nothing
        On Error Resume Next
        Set definition = TestFixtureRecipeDefinition(recipe, "rngRectanglesGeometry", values)
        code = Err.Number: reason = Err.Description: Err.Clear
        On Error GoTo 0
        TestCheck "unified.invalid.s.rejected." & CStr(VarType(key)), code <> 0 And definition Is Nothing And InStr(reason, "rngRectanglesGeometry[") > 0
    Next key

    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Test.Unified"
    Set contour = facade.RectanglePath("Source", 0#, 0#, 1000#, 600#, "Bottom", "Right", "Top", "Left")
    Set contour = facade.OffsetContour("Axis", "Source", Array("Left", "Right", "Top", "Bottom"), Array(30#, 70#, 50#, 90#), Array("left.as", "right.as", "top.as", "bottom.as"))
    Set path = facade.Chain("Top", "Axis", "Top", 1#, "Top", 0#, False)
    facade.Outer "OUTER", "Source"
    Set options = New CRebarRuleOptions: options.EndMode = ExcludeBoth: options.StartSetback = 10#: options.EndSetback = 20#
    facade.RebarByCount "Test.Main", path.ID, 2, 20#, options
    Set layout = TestFixtureAnnotatedLayout(facade.Finish)
    TestNear "unified.excludeBoth.first", layout.X(1), 330#: TestNear "unified.excludeBoth.last", layout.X(2), 620#

    Dim faces As Variant, dimensions As Variant, edge As TRegionEdge, arcFound As Boolean
    faces = Array(Array(20#, 20#, 3, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), _
        Array(80#, 20#, 3, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), _
        Array(40#, 20#, 3, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), _
        Array(60#, 20#, 3, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"))
    dimensions = Array(1000#, 600#, 0#, 0#, 100#, 100#, 0#, 0#)
    Set definition = TestFixtureRoundedDefinition(dimensions, faces)
    Set path = definition.Path("H.Left")
    TestNear "unified.pointed.left.length", path.Length, 460# + 80# * GEOM_PI
    path.AtDistance 20#, x, y, tx, ty
    TestNear "unified.pointed.arc.start.x", x, -400#: TestNear "unified.pointed.arc.start.y", y, 280#
    path.AtDistance 20# + 40# * GEOM_PI, x, y, tx, ty
    TestNear "unified.pointed.arc.end.x", x, -480#: TestNear "unified.pointed.arc.end.y", y, 200#
    Set layout = TestFixtureAnnotatedLayout(definition)
    TestCheck "unified.pointed.variable.as.layout", layout.Count > 0
    faces(0) = Array(20#, 20#, 3, 12#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")
    Set layout = TestFixtureAnnotatedLayout(TestFixtureRoundedDefinition(dimensions, faces))
    TestCheck "unified.pointed.variable.as.dependent", layout.Count > 12
    Set path = definition.Path("B.Top"): path.AtDistance 0#, x, y, tx, ty
    TestNear "unified.pointed.top.start.x", x, -400#: TestNear "unified.pointed.top.start.y", y, 260#

    faces = Array(faces(0), faces(1), faces(2), faces(3), _
        Array(30#, 20#, 3, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), _
        Array(70#, 20#, 3, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), _
        Array(50#, 20#, 3, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), _
        Array(90#, 20#, 3, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"))
    Set definition = TestFixtureHollowDefinition(Array(1000#, 600#, 0#, 400#, 200#, 0#), faces)
    Set path = definition.Path("B.Top"): path.AtDistance 0#, x, y, tx, ty
    TestNear "unified.hollow.outer.start.x", x, -480#: TestNear "unified.hollow.outer.start.y", y, 260#
    path.AtDistance path.Length, x, y, tx, ty: TestNear "unified.hollow.outer.end.x", x, 420#
    Set path = definition.Path("Opening.B.Top"): path.AtDistance 0#, x, y, tx, ty
    TestNear "unified.hollow.opening.start.x", x, -230#: TestNear "unified.hollow.opening.start.y", y, 150#
    path.AtDistance path.Length, x, y, tx, ty: TestNear "unified.hollow.opening.end.x", x, 270#
    Set layout = TestFixtureAnnotatedLayout(definition)
    TestCheck "unified.hollow.variable.as.projections", TestGroupCount(layout, "Opening.B.Top") > 0

    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Test.Collapsed"
    Set contour = facade.RectanglePath("Source", 0#, 0#, 100#, 100#, "Bottom", "Right", "Top", "Left")
    code = 0: reason = ""
    On Error Resume Next
    Set contour = facade.OffsetContour("Axis", "Source", Array("Left", "Right", "Top", "Bottom"), Array(60#, 60#, 10#, 10#), Array("left.as", "right.as", "top.as", "bottom.as"))
    code = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo 0
    TestCheck "unified.disappearing.side.rejected", code <> 0 And InStr(reason, "исчез") > 0 And InStr(reason, "left.as") > 0 And InStr(reason, "right.as") > 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Считает физические стержни, которые состоят в указанной группе, включая дополнительные принадлежности.
Private Function TestGroupCount(ByVal layout As CRebarLayout, ByVal group As String) As Long
    Dim i As Long
    For i = 1 To layout.Count
        If TestInGroup(layout, i, group) Then TestGroupCount = TestGroupCount + 1
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что отфильтрованные проекции сохранили исходные координаты и не были равномерно перераспределены заново.
Private Function TestProjectedSubset(ByVal actual As CRebarLayout, ByVal original As CRebarLayout, ByVal group As String) As Boolean
    Dim i As Long, j As Long, found As Boolean
    For i = 1 To actual.Count
        If TestInGroup(actual, i, group) Then
            found = False
            For j = 1 To original.Count
                If TestInGroup(original, j, group) Then
                    If Abs(actual.X(i) - original.X(j)) < 0.000001 And Abs(actual.Y(i) - original.Y(j)) < 0.000001 Then found = True
                End If
            Next j
            If Not found Then Exit Function
        End If
    Next i
    TestProjectedSubset = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет все принадлежности физического стержня, а не только первую группу публикации.
Private Function TestInGroup(ByVal layout As CRebarLayout, ByVal index As Long, ByVal group As String) As Boolean
    Dim membership As Variant
    For Each membership In layout.BarMemberships(index)
        If CStr(membership) = group Then TestInGroup = True: Exit Function
    Next membership
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Учитывает отдельное условие проверки и сохраняет его имя в итоговом протоколе.
Private Sub TestCheck(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        mPassed = mPassed + 1: mReport = mReport & "OK: " & name & vbCrLf
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает фактическую и ожидаемую величины с заданным в этом наборе допуском, сохраняя числа в протоколе.
Private Sub TestNear(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    TestCheck name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) <= 0.000001
End Sub
