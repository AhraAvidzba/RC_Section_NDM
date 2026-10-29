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
    TestUnifiedContourSourceParameters
    TestExplicitEndModes
End Sub


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет исходные параметры после двух смещений, переименования и обратного обхода.
' Аналитические координаты задаются прямоугольником 1000x600, первым отступом 10 и независимыми as=30/70/50/90.
Private Sub TestUnifiedContourSourceParameters()
    Dim definition As CSectionDefinition, contour As CGeometryPath, path As CGeometryPath, transferred As CGeometryPath
    Dim query As CSectionGeometryQuery, layout As CRebarLayout, i As Long, code As Long, reason As String, faces As Variant, joins As Long
    Set definition = TestProvenanceOffsetDefinition(False)
    Set contour = definition.Path("Axis"): i = contour.SegmentIndex("Renamed.Bottom")
    TestCheck "unified.provenance.root.path", contour.SourcePathID(i) = "Original"
    TestCheck "unified.provenance.root.segment", contour.SourceSegmentID(i) = "Bottom"
    TestNear "unified.provenance.bottom.first", contour.SourceStartParameter(i), 0.04
    TestNear "unified.provenance.bottom.last", contour.SourceEndParameter(i), 0.92
    Set path = definition.Path("Top")
    TestNear "unified.provenance.reversed.first", path.SourceStartParameter(1), 0.96
    TestNear "unified.provenance.reversed.last", path.SourceEndParameter(1), 0.08
    Set query = New CSectionGeometryQuery: Set transferred = query.OffsetPath("Transferred", contour, 10#)
    i = transferred.SegmentIndex("Renamed.Bottom")
    TestCheck "unified.provenance.transfer.root", transferred.SourcePathID(i) = "Original" And transferred.SourceSegmentID(i) = "Bottom"
    TestNear "unified.provenance.transfer.first", transferred.SourceStartParameter(i), 0.05
    TestNear "unified.provenance.transfer.last", transferred.SourceEndParameter(i), 0.91
    Set layout = TestFixtureAnnotatedLayout(definition)
    TestCheck "unified.provenance.layout.count", layout.Count = 6
    TestCheck "unified.provenance.layout.main.count", TestGroupCount(layout, "Main") = 3
    TestCheck "unified.provenance.layout.row2.count", TestGroupCount(layout, "R2") = 3
    TestNear "unified.provenance.layout.first.x", layout.X(1), 40#
    TestNear "unified.provenance.layout.last.x", layout.X(3), 920#
    For i = 1 To layout.Count
        If TestInGroup(layout, i, "Main") Then
            TestNear "unified.provenance.layout.y." & CStr(i), layout.Y(i), 540#
        Else
            TestNear "unified.provenance.layout.row2.y." & CStr(i), layout.Y(i), 520#
        End If
    Next i
    code = 0: reason = "": Set layout = Nothing
    On Error Resume Next
    Set layout = TestFixtureAnnotatedLayout(TestProvenanceOffsetDefinition(True))
    code = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo 0
    TestCheck "unified.provenance.wrongSource.rejected", code <> 0 And InStr(reason, "не принадлежит заданной группе") > 0

    faces = Array(Array(20#, 20#, 3, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), _
        Array(80#, 20#, 3, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), _
        Array(40#, 20#, 3, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"), _
        Array(60#, 20#, 3, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar"))
    Set definition = TestFixtureRoundedDefinition(Array(1000#, 600#, 0#, 0#, 100#, 100#, 0#, 0#), faces)
    Set contour = definition.Path("Rebar.Boundary")
    For i = 1 To contour.SegmentCount
        If Left$(contour.SegmentID(i), 12) = "Offset.Join." Then
            joins = joins + 1
            TestCheck "unified.provenance.join.path." & CStr(i), contour.SourcePathID(i) = contour.ID
            TestCheck "unified.provenance.join.segment." & CStr(i), contour.SourceSegmentID(i) = contour.SegmentID(i)
            TestNear "unified.provenance.join.first." & CStr(i), contour.SourceStartParameter(i), 0#
            TestNear "unified.provenance.join.last." & CStr(i), contour.SourceEndParameter(i), 1#
        End If
    Next i
    TestCheck "unified.provenance.joins.present", joins > 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт производный контур с отличающимися локальными именами и сохранёнными исходными параметрами.
' Ненулевой отступ R2 включает проверку связей контура; без него этот контракт штатно не используется.
' Отрицательный вариант намеренно связывает верхний Main с нижней исходной гранью; общий исполнитель обязан отклонить такую связь.
Private Function TestProvenanceOffsetDefinition(ByVal wrongSource As Boolean) As CSectionDefinition
    Dim facade As CSectionDefinitionBuilder, original As CGeometryPath, inner As CGeometryPath, renamed As CGeometryPath
    Dim query As CSectionGeometryQuery, axis As CGeometryPath, path As CGeometryPath, options As CRebarRuleOptions
    Dim edge As TRegionEdge, i As Long, selected As String
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Test.Provenance"
    Set original = facade.RectanglePath("Original", 0#, 0#, 1000#, 600#, "Bottom", "Right", "Top", "Left")
    Set query = New CSectionGeometryQuery: Set inner = query.OffsetPath("Inner", original, 10#)
    Set renamed = New CGeometryPath: renamed.Initialize "Renamed"
    For i = 1 To inner.SegmentCount
        inner.CopySegment i, edge
        renamed.Append "Renamed." & inner.SegmentID(i), edge, inner.SourcePathID(i), inner.SourceSegmentID(i), inner.SourceStartParameter(i), inner.SourceEndParameter(i)
    Next i
    renamed.Seal True: facade.AddPath renamed
    Set axis = facade.OffsetContour("Axis", renamed.ID, Array("Renamed.Left", "Renamed.Right", "Renamed.Top", "Renamed.Bottom"), _
        Array(30#, 70#, 50#, 90#), Array("left.as", "right.as", "top.as", "bottom.as"))
    facade.Outer "OUTER", renamed.ID
    Set path = facade.Chain("Top", axis.ID, "Renamed.Top", 1#, "Renamed.Top", 0#, False)
    Set options = New CRebarRuleOptions: options.SourceReference = "TestProvenance.As": options.AnnotationGroup = "Main"
    facade.RebarByCount "Main", path.ID, 3, 20#, options
    Set options = New CRebarRuleOptions: options.SourceReference = "TestProvenance.R2": options.AnnotationGroup = "R2"
    options.NormalOffset = -20#: options.ValidateRowOffset = True
    facade.RebarDependent "R2", "Main", 20#, options
    selected = "Renamed.Top": If wrongSource Then selected = "Renamed.Bottom"
    facade.RebarContour "Main", "OUTER", renamed.ID, 50#, "R2", "", sourceSegments:=Array(selected), normalSign:=-1#
    Set TestProvenanceOffsetDefinition = facade.Finish
End Function

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

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет координаты четырёх открытых режимов, один стержень, замкнутый обход и адреса несовместимых настроек.
Private Sub TestExplicitEndModes()
    Dim modes As Variant, starts As Variant, ends As Variant, singles As Variant, mode As Long, count As Long
    Dim layout As CRebarLayout, i As Long, minX As Double, maxX As Double, values As Variant, key As Variant
    Dim recipe As ISectionRecipe, definition As CSectionDefinition, code As Long, reason As String, reader As CSectionInputReader
    modes = Array(IncludeBoth, ExcludeBoth, IncludeStart, IncludeEnd)
    starts = Array(120#, 180#, 120#, 200#): ends = Array(360#, 300#, 280#, 360#)
    singles = Array(240#, 240#, 120#, 360#)
    Dim modeNames As Variant: modeNames = Array("IncludeBoth", "ExcludeBoth", "IncludeStart", "IncludeEnd")
    values = TestFixtureRecipeValues("rngCircleGeometry")
    For mode = 0 To 3
        TestFixtureSetInput values, "Main", "Circle", "endMode", CStr(modeNames(mode))
        Set reader = TestFixtureInput(values)
        TestCheck "endMode.reader.typed." & CStr(mode), reader.RebarEndMode("Main", "Circle") = CLng(modes(mode))
        For count = 1 To 3 Step 2
            Set layout = TestEndModeLayout(CLng(modes(mode)), count)
            TestCheck "endMode.open.count." & CStr(mode) & "." & CStr(count), layout.Count = count
            minX = 1E+30: maxX = -1E+30
            For i = 1 To layout.Count
                minX = GeomMin(minX, layout.X(i)): maxX = GeomMax(maxX, layout.X(i))
                TestNear "endMode.open.y." & CStr(mode) & "." & CStr(count) & "." & CStr(i), layout.Y(i), 100#
            Next i
            If count = 1 Then
                TestNear "endMode.single." & CStr(mode), minX, CDbl(singles(mode))
            Else
                TestNear "endMode.first." & CStr(mode), minX, CDbl(starts(mode))
                TestNear "endMode.last." & CStr(mode), maxX, CDbl(ends(mode))
            End If
        Next count
    Next mode
    Set layout = TestEndModeLayout(IncludeStart, 4, True)
    TestCheck "endMode.closed.count", layout.Count = 4
    For i = 1 To 4
        TestNear "endMode.closed.radius." & CStr(i), Sqr((layout.X(i) - 250#) ^ 2 + (layout.Y(i) - 250#) ^ 2), 100#
    Next i
    TestNear "endMode.closed.unique.spacing", Sqr((layout.X(1) - layout.X(4)) ^ 2 + (layout.Y(1) - layout.Y(4)) ^ 2), Sqr(2#) * 100#

    values = TestFixtureRecipeValues("rngRectanglesGeometry")
    TestFixtureSetInput values, "General", "Shape", "val", "OneRect"
    For Each key In Array("R1.Top", "R1.Bottom", "R1.Left", "R1.Right")
        TestFixtureSetInput values, "Main", CStr(key), "n", 0
        TestFixtureSetInput values, "Main", CStr(key), "as", 40#
    Next key
    TestFixtureSetInput values, "Main", "R1.Top", "n", 3
    TestFixtureSetInput values, "Main", "R1.Top", "d", 20#
    Set recipe = New CSectionRecipeRectangles
    For Each key In Array("ClosedUniform", "Auto", "unknown", "IncludeBoth,ExcludeBoth", "IncludeBoth;ExcludeBoth", "", CVErr(xlErrValue))
        TestFixtureSetInput values, "Main", "R1.Top", "endMode", key
        code = 0: reason = "": Set definition = Nothing
        On Error Resume Next
        Set definition = TestFixtureRecipeDefinition(recipe, "rngRectanglesGeometry", values)
        code = Err.Number: reason = Err.Description: Err.Clear
        On Error GoTo 0
        TestCheck "endMode.invalid.address." & CStr(VarType(key)) & "." & CStr(Len(reason)), code <> 0 And definition Is Nothing And InStr(reason, "rngRectanglesGeometry[") > 0
    Next key
    values = TestFixtureRecipeValues("rngHollowRectangleGeometry")
    TestFixtureSetInput values, "Main", "OpeningTop", "endMode", "IncludeBoth"
    Set recipe = New CSectionRecipeHollowRectangle: code = 0: reason = "": Set definition = Nothing
    On Error Resume Next
    Set definition = TestFixtureRecipeDefinition(recipe, "rngHollowRectangleGeometry", values)
    code = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo 0
    TestCheck "endMode.projection.inapplicable.address", code <> 0 And InStr(reason, "проекц") > 0 And InStr(reason, "rngHollowRectangleGeometry[") > 0
    TestMainCoincidenceLabels
    TestCircleSingleEndModes
    TestCircleClosedEndModes
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт независимый путь: открытый x100..400 с отступами20/40 либо полный круг радиуса100 внутри бетона.
Private Function TestEndModeLayout(ByVal mode As ERebarEndMode, ByVal count As Long, Optional ByVal closed As Boolean = False) As CRebarLayout
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Test.EndMode"
    Set path = facade.RectanglePath("Concrete", 0#, 0#, 600#, 600#): facade.Outer "OUTER", path.ID
    If closed Then
        Set path = facade.CirclePath("Axis", 250#, 250#, 100#)
    Else
        Set path = facade.LinePath("Axis", 100#, 100#, 400#, 100#)
    End If
    Set options = New CRebarRuleOptions: options.EndMode = mode
    If Not closed Then options.StartSetback = 20#: options.EndSetback = 40#
    facade.RebarByCount "Main", path.ID, count, 20#, options
    Set TestEndModeLayout = TestFixtureAnnotatedLayout(facade.Finish)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет общий угол двух Main: первый стержень сохраняется, второй пропускается с предупреждением, подписи считают только физические стержни.
Private Sub TestMainCoincidenceLabels()
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions, copied As CRebarRuleOptions
    Dim layout As CRebarLayout, executor As CRebarRuleExecutor, labels As CSectionAnnotations, definition As CSectionDefinition
    Dim diameter As Variant, i As Long, owner As Long, firstFound As Boolean, secondFound As Boolean
    For Each diameter In Array(20#, 16#)
        Set facade = New CSectionDefinitionBuilder: facade.Initialize "Test.MainCoincidence"
        Set path = facade.RectanglePath("Concrete", 0#, 0#, 600#, 600#): facade.Outer "OUTER", path.ID
        Set path = facade.LinePath("AxisA", 100#, 100#, 300#, 100#)
        Set path = facade.LinePath("AxisB", 300#, 100#, 300#, 300#)
        Set options = New CRebarRuleOptions: options.Coincidence = KeepPublishedOwner: options.WarnOnCoincidence = True
        Set copied = options.Clone: options.WarnOnCoincidence = False
        TestCheck "main.coincidence.options.clone", copied.WarnOnCoincidence
        copied.AnnotationGroup = "First": copied.SourceReference = "First.EndMode"
        facade.RebarByCount "First.Main", "AxisA", 3, 20#, copied, "A400"
        copied.AnnotationGroup = "Second": copied.SourceReference = "Second.EndMode"
        facade.RebarByCount "Second.Main", "AxisB", 3, CDbl(diameter), copied, "A400"
        copied.WarnOnCoincidence = False ' Проверяем, что объявленные правила сохранили независимую копию параметров.
        facade.RebarLabel "First", "AxisA", 0#, -1#: facade.RebarLabel "Second", "AxisB", 1#, 0#
        Set definition = facade.Finish: Set executor = New CRebarRuleExecutor
        Set layout = TestFixtureAnnotatedLayout(definition, executor, labels)
        TestCheck "main.coincidence.physical.count." & CStr(diameter), layout.Count = 5 And executor.RequestedPositions = 6 And executor.MergedPositions = 1
        owner = layout.FindBarAt(300#, 100#)
        TestCheck "main.coincidence.first.owner." & CStr(diameter), owner > 0
        If owner > 0 Then
            TestNear "main.coincidence.first.diameter." & CStr(diameter), layout.Diameter(owner), 20#
            TestCheck "main.coincidence.first.material." & CStr(diameter), layout.SteelClass(owner) = "A400"
            TestCheck "main.coincidence.first.label.group." & CStr(diameter), layout.BarAnnotationGroupName(owner) = "First"
        End If
        TestCheck "main.coincidence.warning.typed." & CStr(diameter), executor.HasGeometryWarnings And InStr(executor.GeometryWarningText, "пропущено совпавших стержней: 1") > 0
        TestCheck "main.coincidence.warning.detail." & CStr(diameter), executor.DiagnosticCount = 1 And executor.DiagnosticAt(1).Code = "COINCIDENCE_SKIPPED" And executor.DiagnosticAt(1).SourceReference = "Second.EndMode"
        TestCheck "main.coincidence.warning.saved.text." & CStr(diameter), InStr(executor.DiagnosticsText, "Second.EndMode") > 0 And InStr(executor.DiagnosticsText, "пропущен") > 0
        firstFound = False: secondFound = False
        For i = 1 To labels.Count
            If labels.AnnotationID(i) = "REBAR_First" Then firstFound = labels.Text(i) = "3" & ChrW$(216) & "20"
            If labels.AnnotationID(i) = "REBAR_Second" Then secondFound = labels.Text(i) = "2" & ChrW$(216) & CStr(diameter)
        Next i
        TestCheck "main.coincidence.labels.actual." & CStr(diameter), firstFound And secondFound And labels.Count = 2
    Next diameter
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет один реальный стержень Circle: четыре режима дуги270 градусов и полный IncludeStart, без старого ограничения n>=2.
Private Sub TestCircleSingleEndModes()
    Dim values As Variant, recipe As ISectionRecipe, layout As CRebarLayout, mode As Long
    Dim modes As Variant, expectedX As Variant, expectedY As Variant
    modes = Array("IncludeBoth", "ExcludeBoth", "IncludeStart", "IncludeEnd", "IncludeStart")
    expectedX = Array(100# / Sqr(2#), 100# / Sqr(2#), 0#, -100#, 0#)
    expectedY = Array(-100# / Sqr(2#), -100# / Sqr(2#), 100#, 0#, 100#)
    For mode = 0 To 4
        values = TestFixtureRecipeValues("rngCircleGeometry")
        TestFixtureSetInput values, "General", "D", "val", 300#
        TestFixtureSetInput values, "Main", "Circle", "as", 50#
        TestFixtureSetInput values, "Main", "Circle", "d", 20#
        TestFixtureSetInput values, "Main", "Circle", "n", 1
        TestFixtureSetInput values, "Main", "Circle", "endMode", modes(mode)
        If mode < 4 Then TestFixtureSetInput values, "Main", "Circle", "s1", 50# * GEOM_PI
        TestFixtureSetInput values, "Extra", "Circle", "d2", 0#
        TestFixtureSetInput values, "Extra", "Circle", "d3", 0#
        Set recipe = New CSectionRecipeCircle
        Set layout = TestFixtureAnnotatedLayout(TestFixtureRecipeDefinition(recipe, "rngCircleGeometry", values))
        TestCheck "endMode.circle.single.count." & CStr(mode), layout.Count = 1
        If layout.Count > 0 Then
            TestNear "endMode.circle.single.x." & CStr(mode), layout.X(1), CDbl(expectedX(mode))
            TestNear "endMode.circle.single.y." & CStr(mode), layout.Y(1), CDbl(expectedY(mode))
        End If
    Next mode
End Sub


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет все четыре режима на полном круге, один стержень, фактические количества и отсутствие повторных дочерних стержней на замыкании.
Private Sub TestCircleClosedEndModes()
    Dim values As Variant, recipe As ISectionRecipe, layout As CRebarLayout, executor As CRebarRuleExecutor
    Dim labels As CSectionAnnotations, mode As Long, count As Variant, i As Long, wanted As Long, ordinal As Long
    Dim modes As Variant, fraction As Double, angle As Double, sum As Double, kind As Long, code As Long, reason As String
    Dim mainCount As Long, row2Count As Long, row3Count As Long
    Dim labelParts As Variant, part As Variant, labelCount As Long, labelMatch As Boolean
    Dim diagnosticCount As Long, diagnostic As CGeometryDiagnostic
    modes = Array("IncludeBoth", "ExcludeBoth", "IncludeStart", "IncludeEnd")
    For Each count In Array(1, 4, 5)
        For mode = 0 To 3
            values = TestCircleClosedValues(CStr(modes(mode)), CLng(count))
            Set recipe = New CSectionRecipeCircle: Set executor = New CRebarRuleExecutor
            Set layout = TestFixtureAnnotatedLayout(TestFixtureRecipeDefinition(recipe, "rngCircleGeometry", values), executor)
            wanted = CLng(count): If mode = 0 And wanted > 1 Then wanted = wanted - 1
            TestCheck "endMode.full.circle.count." & CStr(mode) & "." & CStr(count), layout.Count = wanted
            For i = 1 To layout.Count
                If CLng(count) = 1 Then
                    fraction = 0#: If mode < 2 Then fraction = 0.5
                    If mode = 3 Then fraction = 1#
                Else
                    Select Case mode
                        Case 0: fraction = CDbl(i - 1) / (CDbl(count) - 1#)
                        Case 1: fraction = CDbl(i) / (CDbl(count) + 1#)
                        Case 2: fraction = CDbl(i - 1) / CDbl(count)
                        Case 3: fraction = CDbl(i) / CDbl(count)
                    End Select
                End If
                angle = GEOM_PI / 2# - 2# * GEOM_PI * fraction
                TestNear "endMode.full.circle.x." & CStr(mode) & "." & CStr(count) & "." & CStr(i), layout.X(i), 100# * Cos(angle)
                TestNear "endMode.full.circle.y." & CStr(mode) & "." & CStr(count) & "." & CStr(i), layout.Y(i), 100# * Sin(angle)
            Next i
            TestCheck "endMode.full.circle.warning." & CStr(mode) & "." & CStr(count), executor.HasGeometryWarnings = (mode = 0 And CLng(count) > 1)
        Next mode
    Next count
    values = TestCircleClosedValues("IncludeBoth", 5)
    TestFixtureSetInput values, "Extra", "Circle", "d2", 12#
    TestFixtureSetInput values, "Extra", "Circle", "d3", 10#
    TestFixtureSetInput values, "Extra", "Circle", "bind2", "EachBar"
    TestFixtureSetInput values, "Extra", "Circle", "bind3", "EverySecondBar"
    TestFixtureSetInput values, "Extra", "Circle", "pos2", "Stacked"
    TestFixtureSetInput values, "Extra", "Circle", "pos3", "Stacked"
    TestFixtureSetInput values, "Extra", "Circle", "a2", 0#
    TestFixtureSetInput values, "Extra", "Circle", "a3", 0#
    Set executor = New CRebarRuleExecutor: Set recipe = New CSectionRecipeCircle
    Set layout = TestFixtureAnnotatedLayout(TestFixtureRecipeDefinition(recipe, "rngCircleGeometry", values), executor, labels)
    For i = 1 To layout.Count
        Select Case layout.Diameter(i)
            Case 20#: mainCount = mainCount + 1
            Case 12#: row2Count = row2Count + 1
            Case 10#: row3Count = row3Count + 1
        End Select
    Next i
    TestCheck "endMode.full.circle.dependent.physical", layout.Count = 10 And mainCount = 4 And row2Count = 4 And row3Count = 2
    TestCheck "endMode.full.circle.dependent.accounting", executor.RequestedPositions = 13 And executor.CreatedPositions = 10 And executor.MergedPositions = 3
    TestCheck "endMode.full.circle.dependent.warning", executor.HasGeometryWarnings And InStr(executor.GeometryWarningText, "пропущено совпавших стержней: 1") > 0
    For i = 1 To labels.Count
        If labels.AnnotationID(i) = "REBAR_Circle" Then
            labelCount = labelCount + 1
            labelMatch = (labels.Text(i) = "4" & ChrW$(216) & "20 + 4" & ChrW$(216) & "12 + 2" & ChrW$(216) & "10")
            labelParts = Split(labels.Text(i), " + ") ' Количества находятся в тексте; Value предназначен для числовых размеров.
            For Each part In labelParts
                If InStr(CStr(part), ChrW$(216)) > 1 Then sum = sum + CLng(Left$(CStr(part), InStr(CStr(part), ChrW$(216)) - 1))
            Next part
        End If
    Next i
    TestCheck "endMode.full.circle.dependent.label.text", labelCount = 1 And labelMatch
    TestNear "endMode.full.circle.dependent.label.total", sum, 10#
    For i = 1 To executor.DiagnosticCount
        Set diagnostic = executor.DiagnosticAt(i)
        If diagnostic.Code = "DEPENDENT_SOURCE_DUPLICATE_SKIPPED" Then
            diagnosticCount = diagnosticCount + 1
            TestCheck "endMode.full.circle.dependent.diagnostic.source." & CStr(i), InStr(diagnostic.SourceReference, "rngCircleGeometry") > 0 And Len(diagnostic.RelatedEntityID) > 0
        End If
    Next i
    TestCheck "endMode.full.circle.dependent.diagnostics", diagnosticCount = 2
    values = TestCircleClosedValues("ClosedUniform", 5): code = 0: reason = ""
    On Error Resume Next
    Set layout = TestFixtureAnnotatedLayout(TestFixtureRecipeDefinition(recipe, "rngCircleGeometry", values))
    code = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo 0
    TestCheck "endMode.full.circle.removed.mode.rejected", code <> 0 And InStr(reason, "rngCircleGeometry[") > 0
    For kind = 0 To 2
        code = 0: Set layout = Nothing
        On Error Resume Next
        Set layout = TestCircleFixedStepLayout(kind)
        code = Err.Number: Err.Clear
        On Error GoTo 0
        If kind = 1 Then
            TestCheck "endMode.full.circle.exact.step.incompatible", code = vbObjectError + 5530 And layout Is Nothing
        Else
            TestCheck "endMode.full.circle.exact.step.compatible." & CStr(kind), code = 0 And Not layout Is Nothing
            If Not layout Is Nothing Then TestCheck "endMode.full.circle.exact.step.count." & CStr(kind), layout.Count = 4
        End If
    Next kind
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Готовит явные исходные значения полной окружности радиуса100; дополнительные ряды выключены до специальных проверок.
Private Function TestCircleClosedValues(ByVal mode As String, ByVal count As Long) As Variant
    Dim values As Variant
    values = TestFixtureRecipeValues("rngCircleGeometry")
    TestFixtureSetInput values, "General", "D", "val", 300#
    TestFixtureSetInput values, "Main", "Circle", "as", 50#
    TestFixtureSetInput values, "Main", "Circle", "d", 20#
    TestFixtureSetInput values, "Main", "Circle", "n", count
    TestFixtureSetInput values, "Main", "Circle", "s0", 0#
    TestFixtureSetInput values, "Main", "Circle", "s1", 0#
    TestFixtureSetInput values, "Main", "Circle", "endMode", mode
    TestFixtureSetInput values, "Extra", "Circle", "d2", 0#
    TestFixtureSetInput values, "Extra", "Circle", "d3", 0#
    TestCircleClosedValues = values
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Полная окружность радиуса100: IncludeStart/IncludeEnd сохраняют ровно четыре точных шага; дробное число шагов отклоняется до размещения.
Private Function TestCircleFixedStepLayout(ByVal kind As Long) As CRebarLayout
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions, step As Double
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Test.ClosedStep"
    Set path = facade.RectanglePath("Concrete", 0#, 0#, 500#, 500#): facade.Outer "OUTER", path.ID
    Set path = facade.CirclePath("Axis", 250#, 250#, 100#)
    Set options = New CRebarRuleOptions: options.EndMode = IncludeStart
    If kind = 2 Then options.EndMode = IncludeEnd
    step = path.Length / 4#: If kind = 1 Then step = path.Length / 4.5
    facade.RebarByFixedStep "Main", path.ID, step, 10#, options
    Set TestCircleFixedStepLayout = TestFixtureAnnotatedLayout(facade.Finish)
End Function
