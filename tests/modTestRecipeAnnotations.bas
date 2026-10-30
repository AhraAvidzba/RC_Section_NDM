Attribute VB_Name = "modTestRecipeAnnotations"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: измерения общего API и сохраненные подписи без повторной сборки.
' ==========================================================================
Private mPassed As Long, mFailed As Long, mReport As String

' -------------------- ДЛЯ ТЕСТОВ --------------------
' точные привязки при наклонных нормалях и сохранность границы отверстия.
Public Function RunArbitraryAnnotationGeometryTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim degrees As Variant, value As Variant, shape As Variant, angle As Double, nx As Double, ny As Double
    Dim facade As CSectionDefinitionBuilder, definition As CSectionDefinition, builder As CSectionBuilder, model As CSectionModel
    Dim path As CGeometryPath, polygon As CPolygonSpec, options As CRebarRuleOptions, support As Double, mx As Double, my As Double
    Dim query As CSectionGeometryQuery, region As CConcreteRegion, distance As Double, prefix As String
    Dim annotationBuilder As CSectionAnnotationBuilder, repeated As CSectionAnnotations, count As Long
    degrees = Array(0#, 30#, 45#, 90#, 135#, 225#, 315#)
    For Each shape In Array("Circle", "Rounded", "HollowRectangle")
        For Each value In degrees
            angle = CDbl(value) * GEOM_PI / 180#: nx = Cos(angle): ny = Sin(angle)
            prefix = CStr(shape) & "." & CStr(value)
            Set facade = New CSectionDefinitionBuilder: facade.Initialize "AngleFixture"
            If shape = "HollowRectangle" Then
                Set polygon = TestAngleRoundedPolygon(500#, 600#, 160#)
                Set path = facade.RoundedPolygon("OuterBase", polygon)
                Set path = facade.Transform("Outer", path.ID, 0#, 0#, angle)
            Else
                Set path = facade.CirclePath("Outer", 0#, 0#, 800#)
            End If
            facade.Outer "OUTER", path.ID
            Set options = New CRebarRuleOptions: options.AnnotationGroup = "G"
            If shape = "Circle" Then
                Set path = facade.CirclePath("Axis", 0#, 0#, 150#)
                options.EndMode = IncludeStart: support = 150#
                facade.RebarByCount "Bars", path.ID, 8, 10#, options
            ElseIf shape = "Rounded" Then
                Set polygon = TestAngleRoundedPolygon(200#, 300#, 40#)
                Set path = facade.RoundedPolygon("Axis", polygon)
                options.EndMode = IncludeStart
                support = 160# * Abs(nx) + 260# * Abs(ny) + 40#
                facade.RebarByCount "Bars", path.ID, 12, 10#, options
            Else
                Set polygon = TestAngleRoundedPolygon(200#, 300#, 80#)
                Set path = facade.RoundedPolygon("HoleBase", polygon)
                Set path = facade.Transform("Hole", path.ID, 40#, 20#, angle)
                facade.Opening "OPENING", path.ID, "OUTER"
                Set path = facade.LinePath("AxisBase", -240#, -100#, -240#, 100#)
                Set path = facade.Transform("Axis", path.ID, 40#, 20#, angle)
                facade.RebarByCount "Bars", path.ID, 2, 10#, options
            End If
            facade.RebarLabel "G", path.ID, 2# * nx, 2# * ny
            Set definition = facade.Finish: Set builder = New CSectionBuilder: Set model = TestBuildDefinition(definition, builder)
            TestCheck prefix & ".one.group", model.Annotations.Count = 1
            mx = 0.5 * (model.Annotations.StartX(1) + model.Annotations.EndX(1))
            my = 0.5 * (model.Annotations.StartY(1) + model.Annotations.EndY(1))
            If shape = "HollowRectangle" Then
                TestNear prefix & ".anchor.x", mx, 40# - 240# * nx
                TestNear prefix & ".anchor.y", my, 20# - 240# * ny
                Set query = New CSectionGeometryQuery: Set region = query.RegionFromContours(model.Contours)
                TestCheck prefix & ".hit", query.LabelBoundaryDistance(region, mx, my, nx, ny, distance)
                TestNear prefix & ".opening.not.outer", distance, 40#
                TestCheck prefix & ".opening.role", model.Contours.LoopRole(model.Contours.Count) = "Opening"
            Else
                TestNear prefix & ".support.x", mx, support * nx
                TestNear prefix & ".support.y", my, support * ny
            End If
            TestNear prefix & ".line.perpendicular", (model.Annotations.EndX(1) - model.Annotations.StartX(1)) * nx + _
                (model.Annotations.EndY(1) - model.Annotations.StartY(1)) * ny, 0#
            count = model.RebarCount
            Set annotationBuilder = New CSectionAnnotationBuilder: Set repeated = New CSectionAnnotations
            annotationBuilder.Build definition, builder.Rebars, repeated, builder.Rules
            TestCheck prefix & ".repeat", repeated.Count = 1 And builder.Rebars.Count = count
            TestNear prefix & ".repeat.x", repeated.StartX(1), model.Annotations.StartX(1)
            TestNear prefix & ".repeat.y", repeated.StartY(1), model.Annotations.StartY(1)
        Next value
    Next shape
    GoTo Finished
Failed:
    TestCheck "arbitrary.geometry.runtime." & CStr(Err.Number) & "." & Err.Description, False
Finished:
    RunArbitraryAnnotationGeometryTests = mReport & "TOTAL ARBITRARY ANNOTATION GEOMETRY: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' симметричный скруглённый прямоугольник с известными проекциями.
Private Function TestAngleRoundedPolygon(ByVal halfWidth As Double, ByVal halfHeight As Double, ByVal radius As Double) As CPolygonSpec
    Dim polygon As CPolygonSpec
    Set polygon = New CPolygonSpec
    polygon.AddVertex "BL", -halfWidth, -halfHeight, radius
    polygon.AddVertex "BR", halfWidth, -halfHeight, radius
    polygon.AddVertex "TR", halfWidth, halfHeight, radius
    polygon.AddVertex "TL", -halfWidth, halfHeight, radius
    polygon.Seal
    Set TestAngleRoundedPolygon = polygon
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет размеры, радиусы, длину пути, текст и подписи фактически созданной арматуры, а также недопустимые привязки.
Public Function RunRecipeAnnotationTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim facade As CSectionDefinitionBuilder, definition As CSectionDefinition, builder As CSectionBuilder
    Dim model As CSectionModel, annotations As CSectionAnnotations, x As Double, y As Double, i As Long
    Set facade = TestAnnotationFacade
    facade.Source "Diagonal", "ProgramInputs.PointB"
    Set definition = facade.Finish
    TestCheck "source.scalar", definition.SourceReference("Diagonal") = "ProgramInputs.PointB"
    definition.AnchorCoordinates "B", x, y: x = 999#: y = 999#
    definition.AnchorCoordinates "B", x, y
    TestNear "anchor.immutable.x", x, 30#: TestNear "anchor.immutable.y", y, 40#
    Set builder = New CSectionBuilder: Set model = TestBuildDefinition(definition, builder): Set annotations = model.Annotations
    TestCheck "annotations.count", annotations.Count = 11
    For i = 1 To 9: TestCheck "annotation.id." & CStr(i), Len(annotations.AnnotationID(i)) > 0: Next i
    TestNear "horizontal.projection", annotations.Value(1), 30#
    TestNear "horizontal.placement", annotations.EndY(1), 0#
    TestNear "vertical.projection", annotations.Value(2), 40#
    TestNear "vertical.placement", annotations.EndX(2), 0#
    TestNear "diagonal.distance", annotations.Value(3), 50#
    TestNear "diagonal.normal.x", annotations.NormalX(3), -0.8
    TestNear "diagonal.normal.y", annotations.NormalY(3), 0.6
    TestNear "vertical.explicit", annotations.Value(4), 100#
    TestNear "extent.full.height", annotations.Value(5), 400#
    TestNear "extent.full.width", annotations.Value(6), 400#
    TestNear "radius", annotations.Value(7), 200#
    TestNear "diameter", annotations.Value(8), 400#
    TestNear "radius.endpoint.x", annotations.EndX(7), 200#
    TestNear "radius.endpoint.y", annotations.EndY(7), 0#
    TestNear "diameter.opposite.x", annotations.StartX(8), -200#
    TestNear "length.analytic", annotations.Value(9), 400# * GEOM_PI
    TestCheck "length.distinct.extent", annotations.Value(9) > annotations.Value(5)
    TestCheck "text.unicode", annotations.Text(10) = "Проверка привязки"
    TestCheck "text.kind", annotations.AnnotationType(10) = "TEXT_ANNOTATION"
    TestCheck "length.kind.unit", annotations.AnnotationType(9) = "MEASUREMENT_LABEL" And annotations.Unit(9) = "Length"
    TestCheck "physical.labels", InStr(annotations.Text(11), "2") > 0 And model.RebarCount = 2
    For i = 0 To 4: TestCheck "invalid." & CStr(i), TestInvalidAnnotation(i): Next i
    TestPhysicalLabelOwnership
    GoTo Finished
Failed:
    TestCheck "annotations.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRecipeAnnotationTests = mReport & "TOTAL RECIPE ANNOTATIONS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет запись и чтение аннотаций в мм, см и м; последующая смена Config не должна менять сохранённые измерения.
Public Function RunRecipeAnnotationSnapshotTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, events As Boolean, updating As Boolean, calculation As Long
    On Error GoTo Failed
    events = Application.EnableEvents: updating = Application.ScreenUpdating: calculation = Application.Calculation
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    Dim facade As CSectionDefinitionBuilder, builder As CSectionBuilder, definition As CSectionDefinition, model As CSectionModel
    Dim properties As CSectionPropertiesCalculator, units As CUnitSystem, settings As CSystemSettingsReader
    Dim writer As CNDMResultsWriter, reader As CSectionPlotDataReader, outputUnit As Variant, i As Long, found As Long
    Set facade = TestAnnotationFacade: Set definition = facade.Finish
    Set builder = New CSectionBuilder: Set model = TestBuildDefinition(definition, builder)
    Set properties = New CSectionPropertiesCalculator: properties.PrepareSnapshot model
    For Each outputUnit In Array("mm", "cm", "m")
        TestSetUnitSetting "Units.Length.Output", CStr(outputUnit)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, model, properties, units
        Dim geometryTable As Variant, propertyTable As Variant, row As Long, foundDiagnostics As Boolean
        geometryTable = ThisWorkbook.Names("rngNDMSectionGeometry").RefersToRange.Resize(1 + model.ConcreteCount + model.RebarCount, 16).Value2
        TestCheck "snapshot.provenance." & CStr(outputUnit), InStr(SectionOptionalText(ReadSectionOptionalJson(CStr(geometryTable(UBound(geometryTable, 1), 16))), "RecipeTrace"), "RuleId=TopBars|PathId=Row|SourcePositionId=TopBars.2") > 0
        propertyTable = ThisWorkbook.Names("rngNDMSectionProperties").RefersToRange.Resize(46, 6).Value2: foundDiagnostics = False
        For row = 2 To UBound(propertyTable, 1)
            If CStr(propertyTable(row, 3)) = "Geometry.RecipeDiagnostics" Then
                foundDiagnostics = True
                TestCheck "snapshot.diagnostics.counts." & CStr(outputUnit), InStr(CStr(propertyTable(row, 4)), "Requested=2; Created=2") > 0
            End If
        Next row
        TestCheck "snapshot.diagnostics.present." & CStr(outputUnit), foundDiagnostics
        TestSetUnitSetting "Units.Length.Input", "m" ' Изменение input и output после сохранения не меняет измерение Results.
        TestSetUnitSetting "Units.Length.Output", "mm"
        settings.LoadFromWorkbook ThisWorkbook
        Set reader = New CSectionPlotDataReader: reader.LoadGeometryPreviewFromWorkbook ThisWorkbook, settings
        found = 0
        For i = 1 To reader.AnnotationCount
            If reader.AnnotationType(i) = "MEASUREMENT_LABEL" Then
                found = found + 1
                TestNear "snapshot.length.anchor.x." & CStr(outputUnit), reader.AnnotationStartX(i), 0#
                TestNear "snapshot.length.anchor.y." & CStr(outputUnit), reader.AnnotationStartY(i), -100#
                TestCheck "snapshot.length.text." & CStr(outputUnit) & "; " & reader.AnnotationText(i), InStr(reader.AnnotationText(i), "L=") > 0 And InStr(reader.AnnotationText(i), CStr(outputUnit)) > 0
            ElseIf reader.AnnotationType(i) = "TEXT_ANNOTATION" Then
                found = found + 1
                TestCheck "snapshot.text." & CStr(outputUnit), reader.AnnotationText(i) = "Проверка привязки"
            End If
        Next i
        TestCheck "snapshot.labels." & CStr(outputUnit), found = 2
    Next outputUnit
    GoTo Finished
Failed:
    TestCheck "snapshot.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    On Error Resume Next
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    Application.EnableEvents = events: Application.ScreenUpdating = updating: Application.Calculation = calculation
    On Error GoTo 0
    RunRecipeAnnotationSnapshotTests = mReport & "TOTAL RECIPE ANNOTATION SNAPSHOT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет только INPUT или OUTPUT длины по ключу строки именованной таблицы единиц.
Private Sub TestSetUnitSetting(ByVal key As String, ByVal value As String)
    Dim range As Object, table As CSectionInputTable
    Set range = ThisWorkbook.Names("rngUnitSettings").RefersToRange
    Set table = New CSectionInputTable: table.Initialize "rngUnitSettings", range.Value2
    If key = "Units.Length.Output" Then
        range.Cells(table.RowByKey("Length"), 4).Value2 = value
    Else
        range.Cells(table.RowByKey("Length"), 2).Value2 = value
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Включает размеры и подписи программного примера, сохраняет схему и возвращает исходные настройки аннотаций.
Public Function TestDumpRecipeAnnotationPreview(ByVal outputDirectory As String) As String
    Dim facade As CSectionDefinitionBuilder, builder As CSectionBuilder, model As CSectionModel, definition As CSectionDefinition
    Dim range As Object, saved As Variant, table As CSectionInputTable, row As Long, result As String, number As Long, reason As String
    On Error GoTo Failed
    Set range = ThisWorkbook.Names("rngPlotAnnotationSettings").RefersToRange: saved = range.Formula
    Set table = New CSectionInputTable: table.Initialize "rngPlotAnnotationSettings", range.Value2
    row = table.RowByKey("Enabled"): range.Cells(row, 2).Value2 = "Yes": range.Cells(row, 3).Value2 = "Yes"
    Set facade = TestAnnotationFacade(True): Set definition = facade.Finish: Set builder = New CSectionBuilder
    Set model = TestBuildDefinition(definition, builder)
    result = modTestShapeArchitecture.TestPreparedPreviewSnapshot("RecipeAnnotations", model, outputDirectory)
    GoTo Finished
Failed:
    number = Err.Number: reason = Err.Description
Finished:
    On Error Resume Next
    If Not range Is Nothing And Not IsEmpty(saved) Then range.Formula = saved
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "modTestRecipeAnnotations", reason
    TestDumpRecipeAnnotationPreview = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт круг, два стержня и набор независимых измерений; отдельный режим разводит их для наглядной схемы.
Private Function TestAnnotationFacade(Optional ByVal previewPlacement As Boolean = False) As CSectionDefinitionBuilder
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "AnnotationFixture"
    Set path = facade.CirclePath("Boundary", 0#, 0#, 200#): facade.Outer "Concrete", path.ID
    Set path = facade.LinePath("Row", -50#, 100#, 50#, 100#)
    Set options = New CRebarRuleOptions: options.AnnotationGroup = "Top"
    facade.RebarByCount "TopBars", path.ID, 2, 10#, options
    If previewPlacement Then
        facade.Anchor "A", -100#, -75#: facade.Anchor "B", -70#, -35#
        facade.Anchor "LengthNote", -60#, -150#: facade.Anchor "TextNote", 50#, -50#
    Else
        facade.Anchor "A", 0#, 0#: facade.Anchor "B", 30#, 40#: facade.Anchor "LengthNote", 0#, -100#
        facade.Anchor "TextNote", 200#, 0#
    End If
    facade.PathAnchor "East", "Boundary", "Boundary", 0#
    facade.Dimension "Horizontal", "A", "B", "dx", "Horizontal"
    facade.Dimension "Vertical", "A", "B", "dy", "Vertical"
    facade.Dimension "Diagonal", "A", "B", "s"
    If previewPlacement Then facade.VerticalDistance "VerticalExplicit", "v", -100#, 0#, 150#, 1# Else facade.VerticalDistance "VerticalExplicit", "v", -100#, 0#, -50#
    facade.VerticalExtent "Height", "Boundary", "H", -200#
    If previewPlacement Then
        facade.HorizontalExtent "Width", "Boundary", "B", -200#, -1#
        facade.Radius "Radius", "Boundary", "Boundary", 0.75
        facade.Diameter "Diameter", "Boundary", "Boundary", 0.25
    Else
        facade.HorizontalExtent "Width", "Boundary", "B", -200#
        facade.Radius "Radius", "Boundary", "Boundary", 0#
        facade.Diameter "Diameter", "Boundary", "Boundary", 0#
    End If
    facade.PathLength "Length", "Boundary", "LengthNote"
    facade.Text "Note", "TextNote", "Проверка привязки"
    facade.RebarRowLabel "Top", "TopBars", 0#, 1#
    Set TestAnnotationFacade = facade
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит модель с фиксированной контрольной сеткой, используя переданный общий сборщик.
Private Function TestBuildDefinition(ByVal definition As CSectionDefinition, ByVal builder As CSectionBuilder) As CSectionModel
    Dim mesh As TSectionMeshOptions
    mesh.StepX = 100#: mesh.StepY = 100#: mesh.MaterialID = 1: mesh.BoundarySubdivisions = 2
    Set TestBuildDefinition = builder.Build(definition, mesh)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет отказ при неизвестной привязке, неподдерживаемом измерении и попытке задать радиус прямой.
Private Function TestInvalidAnnotation(ByVal scenario As Long) As Boolean
    On Error GoTo Rejected
    Dim facade As CSectionDefinitionBuilder, definition As CSectionDefinition, builder As CSectionBuilder, model As CSectionModel
    Set facade = TestAnnotationFacade
    Select Case scenario
        Case 0: facade.Dimension "Missing", "Unknown", "B", "s"
        Case 1: facade.Dimension "UnknownMode", "A", "B", "s", "Unknown"
        Case 2: facade.Radius "LineRadius", "Row", "Row"
        Case 3: facade.Diameter "InvalidParameter", "Boundary", "Boundary", 2#
        Case 4: facade.PathAnchor "MissingSegment", "Boundary", "RemovedArc"
    End Select
    Set definition = facade.Finish: Set builder = New CSectionBuilder: Set model = TestBuildDefinition(definition, builder)
    Exit Function
Rejected:
    TestInvalidAnnotation = (Err.Number = vbObjectError + 5530 Or Err.Number = vbObjectError + 5520): Err.Clear
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Учитывает отдельное условие проверки и сохраняет его имя в итоговом протоколе.
Private Sub TestCheck(ByVal name As String, ByVal value As Boolean)
    If value Then mPassed = mPassed + 1 Else mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает фактическую и ожидаемую величины с заданным в этом наборе допуском, сохраняя числа в протоколе.
Private Sub TestNear(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    TestCheck name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) <= 0.000000001 * (1# + Abs(expected))
End Sub


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Полностью совпавший ряд не получает подпись ни при одной привязке; частичный ряд показывает только созданные стержни.
Private Sub TestPhysicalLabelOwnership()
    Dim kind As Long, partial As Long, facade As CSectionDefinitionBuilder, path As CGeometryPath
    Dim options As CRebarRuleOptions, definition As CSectionDefinition, executor As CRebarRuleExecutor
    Dim layout As CRebarLayout, labels As CSectionAnnotations, repeated As CSectionAnnotations, builder As CSectionAnnotationBuilder
    Dim i As Long, firstText As String, secondText As String, prefix As String
    For kind = 0 To 3
        For partial = 0 To 1
            prefix = "physical.label.kind" & CStr(kind) & ".partial" & CStr(partial)
            Set facade = New CSectionDefinitionBuilder: facade.Initialize "Test.PhysicalLabels"
            Set path = facade.RectanglePath("Concrete", 0#, 0#, 600#, 600#): facade.Outer "OUTER", path.ID
            Set path = facade.LinePath("AxisA", 100#, 100#, 300#, 100#)
            Set path = facade.LinePath("AxisB", 100# + 200# * partial, 100#, 300# + 200# * partial, 100#)
            Set options = New CRebarRuleOptions
            options.Coincidence = KeepPublishedOwner: options.WarnOnCoincidence = True: options.EndMode = IncludeBoth
            options.AnnotationGroup = "First": facade.RebarByCount "First.Main", "AxisA", 3, 20#, options
            options.AnnotationGroup = "Second": facade.RebarByCount "Second.Main", "AxisB", 3, 16#, options
            facade.RebarLabel "First", "AxisA", 0#, -1#
            Select Case kind
                Case 0: facade.RebarLabel "Second", "AxisB", 0#, -1#
                Case 1: facade.RebarRowLabel "Second", "Second.Main", 0#, -1#
                Case 2: facade.RebarRowExtentLabel "Second", "Second.Main", 0#, -1#
                Case 3: facade.RebarRowExtremeLabel "Second", "Second.Main", 0#, -1#, 0#, -1#
            End Select
            Set definition = facade.Finish: Set executor = New CRebarRuleExecutor
            Set layout = TestFixtureAnnotatedLayout(definition, executor, labels)
            TestCheck prefix & ".physical", layout.Count = 3 + 2 * partial
            TestCheck prefix & ".merged", executor.MergedPositions = 3 - 2 * partial
            TestCheck prefix & ".labels", labels.Count = 1 + partial
            firstText = "": secondText = ""
            For i = 1 To labels.Count
                If labels.AnnotationID(i) = "REBAR_First" Then firstText = labels.Text(i)
                If labels.AnnotationID(i) = "REBAR_Second" Then secondText = labels.Text(i)
            Next i
            TestCheck prefix & ".first.owner", firstText = "3" & ChrW$(216) & "20"
            If partial = 0 Then
                TestCheck prefix & ".empty.hidden", secondText = ""
            Else
                TestCheck prefix & ".created.only", secondText = "2" & ChrW$(216) & "16"
            End If
            Set builder = New CSectionAnnotationBuilder: Set repeated = New CSectionAnnotations
            builder.Build definition, layout, repeated, executor
            TestCheck prefix & ".repeat", repeated.Count = labels.Count And layout.Count = 3 + 2 * partial
        Next partial
    Next kind
End Sub
