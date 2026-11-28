Attribute VB_Name = "modTestOptionalOpenings"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: круглое отверстие, радиальные проекции и отключение отверстий.
' ==========================================================================
Private mPassed As Long, mFailed As Long, mReport As String

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет реальные рецепты по копиям входных таблиц без изменения пользовательского ввода.
Public Function RunOptionalOpeningTests() As String
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = ""
    Dim values As Variant, recipe As ISectionRecipe, definition As CSectionDefinition, model As CSectionModel
    Dim builder As CSectionBuilder, scenario As Long, i As Long, radius As Double, x As Double, y As Double
    Dim dx As Double, dy As Double, ray As Double, expected As Long, j As Long
    Dim found As Boolean, region As CConcreteRegion, query As CSectionGeometryQuery
    Set recipe = New CSectionRecipeCircle
    For scenario = 0 To 3
        values = TestCircleValues()
        If scenario = 1 Then
            TestPut values, "General", "Xo", "val", 100#: TestPut values, "General", "Yo", "val", -60#
        ElseIf scenario = 2 Then
            TestPut values, "General", "Do", "val", 100#: TestPut values, "General", "Xo", "val", 250#
        ElseIf scenario = 3 Then
            TestPut values, "General", "Do", "val", 0#
            TestPut values, "General", "Xo", "val", CVErr(xlErrValue): TestPut values, "General", "Yo", "val", "bad"
            TestPut values, "Main", "OpeningCircle", "as", "bad": TestPut values, "Main", "OpeningCircle", "d", "bad"
        End If
        Set definition = TestFixtureRecipeDefinition(recipe, "rngCircleGeometry", values)
        Set builder = New CSectionBuilder: Set model = TestBuild(definition, builder)
        Set region = definition.ConcreteDomain: Set query = New CSectionGeometryQuery
        TestCheck "circle.loops." & scenario, region.LoopCount = IIf(scenario = 3, 1, 2)
        If scenario = 3 Then
            TestNear "solid.area", region.Area, GEOM_PI * 500# ^ 2
            TestCheck "solid.bars", model.RebarCount = 8
            TestCheck "solid.no.hole.dimension", TestAnnotation(model, "DIM_CIRCLE_DO") = 0
            TestCheck "solid.no.hole.label", TestAnnotation(model, "REBAR_Opening.Circle") = 0
        Else
            radius = 150#: x = 0#: y = 0#
            If scenario = 1 Then x = 100#: y = -60#
            If scenario = 2 Then radius = 50#: x = 250#
            TestNear "circle.area." & scenario, region.Area, GEOM_PI * (500# ^ 2 - radius ^ 2)
            TestCheck "circle.void." & scenario, Not query.ContainsPoint(region, x, y)
            expected = 0
            For i = 1 To 8
                dx = 450# * Sin((i - 1) * GEOM_PI / 4#) - x
                dy = 450# * Cos((i - 1) * GEOM_PI / 4#) - y
                ray = Sqr(dx ^ 2 + dy ^ 2)
                dx = x + (radius + 40#) * dx / ray: dy = y + (radius + 40#) * dy / ray ' Независимая аналитическая точка на радиусе отверстия, без вызова RayHit.
                expected = expected + 1: found = False
                For j = 1 To builder.Rebars.Count
                    If builder.Rebars.BarAnnotationGroupName(j) = "Opening.Circle" Then
                        If Abs(builder.Rebars.X(j) - dx) < 0.000001 And Abs(builder.Rebars.Y(j) - dy) < 0.000001 Then found = True
                    End If
                Next j
                TestCheck "circle.projection.to.hole." & scenario & "." & i & "; x=" & dx & "; y=" & dy, found
            Next i
            TestCheck "circle.projected.count." & scenario, model.RebarCount = 8 + expected
            If scenario = 2 Then TestCheck "circle.offset.all.rays.hit", builder.Rules.SkippedProjections = 0 And expected = 8
            i = TestAnnotation(model, "DIM_CIRCLE_DO")
            TestCheck "circle.dimension.exists." & scenario, i > 0
            If i > 0 Then
                TestNear "circle.dimension.value." & scenario, model.Annotations.Value(i), 2# * radius
                TestNear "circle.dimension.right." & scenario, model.Annotations.StartX(i), 500#
                TestNear "circle.dimension.bottom." & scenario, model.Annotations.StartY(i), y - radius
                TestCheck "circle.dimension.direction." & scenario, model.Annotations.NormalX(i) = 1#
            End If
            i = TestAnnotation(model, "REBAR_Opening.Circle")
            TestCheck "circle.label.exists." & scenario, i > 0
            If i > 0 Then TestCheck "circle.label.into.opening." & scenario, model.Annotations.NormalY(i) = -1#
        End If
    Next scenario
    values = TestCircleValues(): TestPut values, "General", "Do", "val", -1#
    TestCheck "circle.negative.rejected", TestRejected(recipe, "rngCircleGeometry", values, "Do")
    values = TestCircleValues(): TestPut values, "General", "Do", "val", 1100#
    TestCheck "circle.oversized.rejected", TestRejected(recipe, "rngCircleGeometry", values)
    values = TestCircleValues(): TestPut values, "General", "Xo", "val", 350#
    TestCheck "circle.tangent.rejected", TestRejected(recipe, "rngCircleGeometry", values)
    values = TestCircleValues(): TestPut values, "General", "Xo", "val", 400#
    TestCheck "circle.intersection.rejected", TestRejected(recipe, "rngCircleGeometry", values)
    values = TestCircleValues(): TestPut values, "Main", "OpeningCircle", "d", 0#: TestPut values, "Main", "OpeningCircle", "as", "bad"
    Set definition = TestFixtureRecipeDefinition(recipe, "rngCircleGeometry", values)
    Set builder = New CSectionBuilder: Set model = TestBuild(definition, builder)
    TestCheck "circle.no.inner.row", model.RebarCount = 8 And TestAnnotation(model, "REBAR_Opening.Circle") = 0
    TestCheck "point.zero.direction.rejected", TestPointProjectionRejected(False)
    TestCheck "point.normal.conflict.rejected", TestPointProjectionRejected(True)
    TestCircleExtra
    TestHollowDisabled
    GoTo Finished
Failed:
    TestCheck "unexpected " & Err.Source & ": " & Err.Description, False
Finished:
    RunOptionalOpeningTests = mReport & "TOTAL OPTIONAL OPENINGS: passed=" & mPassed & "; failed=" & mFailed & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ноль любого размера отключает все поля отверстия, сохраняя наружную обработку углов.
Private Sub TestHollowDisabled()
    Dim values As Variant, recipe As ISectionRecipe, definition As CSectionDefinition, scenario As Long, corner As Variant
    Dim key As Variant, builder As CSectionBuilder, model As CSectionModel, expectedArea As Double
    Set recipe = New CSectionRecipeHollowRectangle
    For Each corner In Array("Скругление", "Фаска")
        For scenario = 0 To 4
            values = TestFixtureRecipeValues("rngHollowRectangleGeometry")
            TestPut values, "General", "H", "val", 800#: TestPut values, "General", "B", "val", 600#
            TestPut values, "General", "R", "val", 50#: TestPut values, "General", "CornerMode", "val", corner
            TestPut values, "General", "Ho", "val", IIf(scenario = 1, "bad", 0#)
            TestPut values, "General", "Bo", "val", IIf(scenario = 0, "bad", 0#)
            If scenario = 3 Then TestPut values, "General", "Ho", "val", CVErr(xlErrValue)
            If scenario = 4 Then TestPut values, "General", "Bo", "val", CVErr(xlErrValue)
            For Each key In Array("Ro", "Xo", "Yo"): TestPut values, "General", CStr(key), "val", CVErr(xlErrValue): Next key
            For Each key In Array("OpeningTop", "OpeningBottom", "OpeningLeft", "OpeningRight")
                TestPut values, "Main", CStr(key), "as", "bad": TestPut values, "Main", CStr(key), "d", "bad"
                TestPut values, "Extra", CStr(key), "d2", "bad": TestPut values, "Extra", CStr(key), "d3", CVErr(xlErrValue)
            Next key
            For Each key In Array("Top", "Bottom", "Left", "Right")
                TestPut values, "Main", CStr(key), "as", 50#: TestPut values, "Main", CStr(key), "d", 16#: TestPut values, "Main", CStr(key), "n", 3
                TestPut values, "Main", CStr(key), "endMode", "IncludeBoth"
                TestPut values, "Extra", CStr(key), "d2", 0#: TestPut values, "Extra", CStr(key), "d3", 0#
            Next key
            Set definition = TestFixtureRecipeDefinition(recipe, "rngHollowRectangleGeometry", values)
            TestCheck "hollow.one.loop." & corner & scenario, definition.LoopCount = 1
            expectedArea = 600# * 800# - (4# - GEOM_PI) * 50# ^ 2
            If corner = "Фаска" Then expectedArea = 600# * 800# - 2# * 50# ^ 2
            TestNear "hollow.solid.area." & corner & scenario, definition.ConcreteDomain.Area, expectedArea
            Set builder = New CSectionBuilder: Set model = TestBuild(definition, builder)
            TestCheck "hollow.outer.bars." & corner & scenario, model.RebarCount > 0
            TestCheck "hollow.no.dimensions." & corner & scenario, TestAnnotation(model, "DIM_H_OPENING") = 0 And TestAnnotation(model, "DIM_B_OPENING") = 0
        Next scenario
    Next corner
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет обе привязки и оба положения внутренних дополнительных рядов.
Private Sub TestCircleExtra()
    Dim values As Variant, definition As CSectionDefinition, recipe As ISectionRecipe, builder As CSectionBuilder, model As CSectionModel
    Dim binding As Variant, location As Variant, i As Long, count As Long, owner As Long, radius As Double, expected As Long
    Set recipe = New CSectionRecipeCircle
    Dim options As CRebarRuleOptions, cloned As CRebarRuleOptions
    Set options = New CRebarRuleOptions: options.ProjectTowardPoint = True
    options.ProjectionPointX = 100#: options.ProjectionPointY = -60#
    Set cloned = options.Clone
    TestCheck "point.options.clone", cloned.ProjectTowardPoint And cloned.ProjectionPointX = 100# And cloned.ProjectionPointY = -60#
    For Each binding In Array("EachBar", "EverySecondBar")
        For Each location In Array("Stacked", "SideBySide")
            values = TestCircleValues()
            TestPut values, "Extra", "OpeningCircle", "d2", 12#: TestPut values, "Extra", "OpeningCircle", "d3", 10#
            TestPut values, "Extra", "OpeningCircle", "bind2", binding: TestPut values, "Extra", "OpeningCircle", "bind3", "EachBar"
            TestPut values, "Extra", "OpeningCircle", "pos2", location: TestPut values, "Extra", "OpeningCircle", "pos3", "Stacked"
            Set definition = TestFixtureRecipeDefinition(recipe, "rngCircleGeometry", values)
            Set builder = New CSectionBuilder: Set model = TestBuild(definition, builder)
            expected = IIf(binding = "EverySecondBar", 4, 8): count = 0
            For i = 1 To builder.Rules.PositionCount
                If builder.Rules.PositionRuleID(i) = "Opening.Circle.Row2" Then
                    count = count + 1: owner = builder.Rules.PositionPhysicalOwner(i)
                    radius = Sqr(builder.Rebars.X(owner) ^ 2 + builder.Rebars.Y(owner) ^ 2)
                    TestCheck "circle.extra.in.concrete." & binding & location & i, radius > 190# And radius < 450#
                    TestNear "circle.extra.diameter." & binding & location & i, builder.Rebars.Diameter(owner), 12#
                End If
            Next i
            TestCheck "circle.extra.binding." & binding & location, count = expected And model.RebarCount = 24 + expected
            TestCheck "circle.extra.label." & binding & location, InStr(model.Annotations.Text(TestAnnotation(model, "REBAR_Opening.Circle")), CStr(expected) & ChrW$(216) & "12") > 0
        Next location
    Next binding
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет нулевой луч и противоречивый выбор точки вместе с местной нормалью общего правила.
Private Function TestPointProjectionRejected(ByVal conflict As Boolean) As Boolean
    On Error GoTo Rejected
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim definition As CSectionDefinition, builder As CSectionBuilder, model As CSectionModel
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "PointProjectionTest"
    Set path = facade.CirclePath("Outer", 0#, 0#, 500#): facade.Outer "OUTER", path.ID
    Set path = facade.ArcPath("Source", 0#, 0#, 100#, GEOM_PI / 2#, -2# * GEOM_PI)
    Set options = New CRebarRuleOptions: options.EndMode = IncludeStart
    facade.RebarByCount "Main", path.ID, 4, 10#, options
    Set path = facade.CirclePath("Target", 0#, 100#, 20#)
    Set options = New CRebarRuleOptions: options.ProjectTowardPoint = True: options.ProjectionPointY = 100#
    If conflict Then options.SourceNormalSign = 1#
    facade.RebarProjected "Projection", path.ID, "Main", 10#, options
    Set definition = facade.Finish: Set builder = New CSectionBuilder: Set model = TestBuild(definition, builder)
    Exit Function
Rejected:
    If conflict Then
        TestPointProjectionRejected = InStr(Err.Description, "SourceNormalSign") > 0
    Else
        TestPointProjectionRejected = InStr(Err.Description, "направление луча не определено") > 0
    End If
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает независимый числовой ввод круга в миллиметрах, с одним наружным и внутренним рядом.
Private Function TestCircleValues() As Variant
    Dim values As Variant
    values = TestFixtureRecipeValues("rngCircleGeometry")
    TestPut values, "General", "D", "val", 1000#: TestPut values, "General", "Do", "val", 300#
    TestPut values, "General", "Xo", "val", 0#: TestPut values, "General", "Yo", "val", 0#
    TestPut values, "Main", "Circle", "as", 50#: TestPut values, "Main", "Circle", "d", 20#: TestPut values, "Main", "Circle", "n", 8
    TestPut values, "Main", "Circle", "endMode", "IncludeStart"
    TestPut values, "Main", "OpeningCircle", "as", 40#: TestPut values, "Main", "OpeningCircle", "d", 16#
    TestPut values, "Extra", "Circle", "d2", 0#: TestPut values, "Extra", "Circle", "d3", 0#
    TestPut values, "Extra", "OpeningCircle", "d2", 0#: TestPut values, "Extra", "OpeningCircle", "d3", 0#
    TestCircleValues = values
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет одну ячейку копии по структурным ключам, не обращаясь к листу Excel.
Private Sub TestPut(ByRef values As Variant, ByVal block As String, ByVal row As String, ByVal column As String, ByVal value As Variant)
    Dim reader As CSectionInputReader: Set reader = TestFixtureInput(values)
    values(reader.RowIndex(block, row), reader.ColumnIndex(block, column)) = value
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит точные границы ячеек штатным сборщиком.
Private Function TestBuild(ByVal definition As CSectionDefinition, ByVal builder As CSectionBuilder) As CSectionModel
    Dim mesh As TSectionMeshOptions
    mesh.StepX = 100#: mesh.StepY = 100#: mesh.MaterialID = 1: mesh.CellGeometry = "Contour"
    Set TestBuild = builder.Build(definition, mesh)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ищет конкретный размер или подпись, чтобы отсутствие отверстия не скрывалось общим количеством аннотаций.
Private Function TestAnnotation(ByVal model As CSectionModel, ByVal id As String) As Long
    Dim i As Long
    For i = 1 To model.Annotations.Count
        If model.Annotations.AnnotationID(i) = id Then TestAnnotation = i: Exit Function
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет отказ настоящего рецепта и, при необходимости, адресный ключ ошибочного поля.
Private Function TestRejected(ByVal recipe As ISectionRecipe, ByVal name As String, ByRef values As Variant, Optional ByVal key As String = "") As Boolean
    On Error GoTo Rejected
    Dim definition As CSectionDefinition: Set definition = TestFixtureRecipeDefinition(recipe, name, values)
    Exit Function
Rejected:
    TestRejected = Len(key) = 0 Or InStr(1, Err.Description, key, vbTextCompare) > 0
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает один результат проверки без остановки остальных независимых случаев.
Private Sub TestCheck(ByVal name As String, ByVal valid As Boolean)
    If valid Then mPassed = mPassed + 1: mReport = mReport & "PASS " & name & vbCrLf Else mFailed = mFailed + 1: mReport = mReport & "FAIL " & name & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает точные геометрические величины с малым абсолютным и относительным допуском.
Private Sub TestNear(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    TestCheck name & "; actual=" & actual & "; expected=" & expected, Abs(actual - expected) <= 0.000001 + 0.000000001 * Abs(expected)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет действующие ссылки «Расчета», единицы и восстановление круглого отверстия из Excel и файла.
Public Function RunOptionalOpeningWorkbookTests() As String
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, calc As Object, savedCalc(0 To 6) As Variant, calcFields As Variant, events As Boolean
    Dim unitName As Variant, settings As CSystemSettingsReader, units As CUnitSystem, factor As Double
    Dim values As Variant, recipe As ISectionRecipe, definition As CSectionDefinition, builder As CSectionBuilder
    Dim model As CSectionModel, restored As CSectionModel, properties As CSectionPropertiesCalculator
    Dim writer As CNDMResultsWriter, snapshot As CCalculationSnapshot, loaded As CCalculationSnapshot, store As CCalculationSnapshotStore
    Dim key As Variant, column As Variant, inputReader As CSectionInputReader, i As Long, a As Variant, b As Variant
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    Set calc = ThisWorkbook.Worksheets("Расчет")
    calcFields = Array("E20", "G28", "G29", "G30", "G31", "G40", "H40")
    For i = 0 To 6: savedCalc(i) = calc.Range(CStr(calcFields(i))).Formula: Next i
    events = Application.EnableEvents: Application.EnableEvents = False
    calc.Range("E20").Value2 = "Circle"
    calc.Range("G28").Value2 = 1000#: calc.Range("G29").Value2 = 300#
    calc.Range("G30").Value2 = 100#: calc.Range("G31").Value2 = -60#
    calc.Range("G40").Value2 = 40#: calc.Range("H40").Value2 = 16#
    Application.Calculate
    TestCheck "workbook.calc.Do", calc.Range("B29").Value2 = "Do"
    TestCheck "workbook.calc.Xo", calc.Range("B30").Value2 = "Xo"
    TestCheck "workbook.calc.Yo", calc.Range("B31").Value2 = "Yo"
    TestCheck "workbook.calc.inner.main", calc.Range("B40").Value2 = "OpeningCircle"
    TestCheck "workbook.calc.inner.extra", calc.Range("B52").Value2 = "OpeningCircle"
    Set inputReader = TestFixtureInput(ThisWorkbook.Names("rngCircleGeometry").RefersToRange.Value2)
    TestNear "workbook.link.Do", inputReader.Number("General", "Do"), 300#
    TestNear "workbook.link.Xo", inputReader.Number("General", "Xo"), 100#
    TestNear "workbook.link.Yo", inputReader.Number("General", "Yo"), -60#
    TestNear "workbook.link.inner.as", inputReader.Number("Main", "OpeningCircle", "as"), 40#
    TestNear "workbook.link.inner.d", inputReader.Number("Main", "OpeningCircle", "d"), 16#
    Set recipe = New CSectionRecipeCircle
    For Each unitName In Array("mm", "cm", "m")
        ThisWorkbook.Names("rngUnitSettings").RefersToRange.Cells(2, 2).Value2 = unitName
        ThisWorkbook.Names("rngUnitSettings").RefersToRange.Cells(2, 3).Value2 = unitName
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
        values = TestCircleValues(): TestPut values, "General", "Xo", "val", 100#: TestPut values, "General", "Yo", "val", -60#
        Set inputReader = TestFixtureInput(values)
        For Each key In Array("D", "Do", "Xo", "Yo")
            values(inputReader.RowIndex("General", CStr(key)), inputReader.ColumnIndex("General", "val")) = CDbl(inputReader.Number("General", CStr(key))) / factor
        Next key
        For Each key In Array("Circle", "OpeningCircle")
            For Each column In Array("as", "d")
                values(inputReader.RowIndex("Main", CStr(key)), inputReader.ColumnIndex("Main", CStr(column))) = inputReader.Number("Main", CStr(key), CStr(column)) / factor
            Next column
        Next key
        Set definition = TestFixtureRecipeDefinition(recipe, "rngCircleGeometry", values, units)
        TestNear "units.area." & unitName, definition.ConcreteDomain.Area, GEOM_PI * (500# ^ 2 - 150# ^ 2)
        Set builder = New CSectionBuilder: Set model = TestBuild(definition, builder)
        Set properties = New CSectionPropertiesCalculator: properties.PrepareSnapshot model
        Set writer = New CNDMResultsWriter
        Set snapshot = writer.PrepareGeometryPreviewSnapshot(ThisWorkbook, model, properties, units)
        TestSetSystem "Results.SnapshotStorage", "Excel"
        writer.RestoreSnapshotBase ThisWorkbook, snapshot, False
        Set restored = ReadFormedSectionFromResults(ThisWorkbook, True)
        TestCheck "excel.roundtrip.bars." & unitName, restored.RebarCount = model.RebarCount
        i = TestAnnotation(restored, "DIM_CIRCLE_DO")
        TestCheck "excel.roundtrip.dimension." & unitName, i > 0
        If i > 0 Then TestNear "excel.roundtrip.Do." & unitName, restored.Annotations.Value(i), 300#
        TestSetSystem "Results.SnapshotStorage", "File": TestSetSystem "Results.SnapshotName", "CircleOpeningPreview_" & unitName
        Set store = New CCalculationSnapshotStore: store.SaveFile ThisWorkbook, snapshot
        Set loaded = store.LoadFile(ThisWorkbook)
        For Each key In snapshot.Keys
            a = snapshot.Table(CStr(key)): b = loaded.Table(CStr(key))
            TestCheck "file.shape." & unitName & "." & key, UBound(a, 1) = UBound(b, 1) And UBound(a, 2) = UBound(b, 2)
        Next key
        writer.RestoreSnapshotBase ThisWorkbook, loaded, False, False
        Set restored = ReadFormedSectionFromResults(ThisWorkbook, True)
        TestCheck "file.roundtrip.contours." & unitName, restored.Contours.Count = model.Contours.Count
        TestCheck "file.roundtrip.bars." & unitName, restored.RebarCount = model.RebarCount
        For i = 1 To model.RebarCount
            TestNear "file.bar.x." & unitName & i, restored.RebarX(i), model.RebarX(i)
            TestNear "file.bar.y." & unitName & i, restored.RebarY(i), model.RebarY(i)
        Next i
        i = TestAnnotation(restored, "DIM_CIRCLE_DO")
        TestCheck "file.roundtrip.dimension." & unitName, i > 0
        If i > 0 Then TestNear "file.roundtrip.Do." & unitName, restored.Annotations.Value(i), 300#
    Next unitName
    GoTo Finished
Failed:
    TestCheck "unexpected workbook " & Err.Source & ": " & Err.Description, False
Finished:
    If Not config Is Nothing Then If Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    If Not calc Is Nothing Then
        If IsArray(calcFields) Then
            For i = 0 To 6: calc.Range(CStr(calcFields(i))).Formula = savedCalc(i): Next i
        End If
    End If
    Application.EnableEvents = events
    RunOptionalOpeningWorkbookTests = mReport & "TOTAL OPTIONAL OPENING WORKBOOK: passed=" & mPassed & "; failed=" & mFailed & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет только существующий ключ системной таблицы частной книги.
Private Sub TestSetSystem(ByVal key As String, ByVal value As Variant)
    Dim table As Object, i As Long: Set table = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    For i = 2 To table.Rows.Count
        If table.Cells(i, 1).Value2 = key Then table.Cells(i, 2).Value2 = value: Exit Sub
    Next i
    Err.Raise 5, "modTestOptionalOpenings", "Не найден тестовый ключ " & key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет настоящую Excel-схему с эксцентрическим отверстием и проецируемой арматурой.
Public Function TestOptionalOpeningPreview(ByVal directory As String) As String
    Dim values As Variant, recipe As ISectionRecipe, definition As CSectionDefinition, builder As CSectionBuilder, model As CSectionModel
    values = TestCircleValues(): TestPut values, "General", "Xo", "val", 100#: TestPut values, "General", "Yo", "val", -60#
    TestPut values, "Extra", "OpeningCircle", "d2", 12#: TestPut values, "Extra", "OpeningCircle", "pos2", "Stacked"
    TestPut values, "Extra", "OpeningCircle", "bind2", "EachBar"
    Set recipe = New CSectionRecipeCircle: Set definition = TestFixtureRecipeDefinition(recipe, "rngCircleGeometry", values)
    Set builder = New CSectionBuilder: Set model = TestBuild(definition, builder)
    TestOptionalOpeningPreview = modTestShapeArchitecture.TestPreparedPreviewSnapshot("CircleOpening", model, directory)
End Function
