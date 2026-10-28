Attribute VB_Name = "modTestGeometryReview"
Option Explicit
' ==========================================================================
' ДЛЯ ТЕСТОВ: воспроизведение F01/F02 и проверка подписей отключённых граней.
' ==========================================================================
Private mPassed As Long, mFailed As Long, mReport As String

' -------------------- ДЛЯ ТЕСТОВ --------------------
' бетонная граница, сторона отступа и полная подпись независимых рядов.
Public Function RunRebarLabelBoundaryTests(Optional ByVal outputDirectory As String = "") As String
    mPassed = 0: mFailed = 0: mReport = ""
    TestLabelBoundaryGeometry
    TestLabelBoundaryLayout
    TestIndependentRowLabels 2, False, outputDirectory
    TestIndependentRowLabels 3, False, outputDirectory
    TestIndependentRowLabels 2, True, outputDirectory
    TestIndependentRowLabels 3, True, outputDirectory
    RunRebarLabelBoundaryTests = mReport & "TOTAL_REBAR_LABEL_BOUNDARY: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Эталон задаётся аналитически, включая отверстие и точку вне бетона.
Private Sub TestLabelBoundaryGeometry()
    On Error GoTo Failed
    Dim query As CSectionGeometryQuery, region As CConcreteRegion, opening As CConcreteRegion
    Dim distance As Double, found As Boolean, errorNumber As Long, coordinates(1 To 4, 1 To 2) As Double
    Set query = New CSectionGeometryQuery
    coordinates(1, 1) = -300#: coordinates(1, 2) = -600#: coordinates(2, 1) = 300#: coordinates(2, 2) = -600#
    coordinates(3, 1) = 300#: coordinates(3, 2) = 600#: coordinates(4, 1) = -300#: coordinates(4, 2) = 600#
    Set region = query.PolygonRegion(coordinates)
    found = query.LabelBoundaryDistance(region, 0#, 550#, 0#, 2#, distance)
    TestCheck "label.rectangle.as50", found And Abs(distance - 50#) < 0.000001
    found = query.LabelBoundaryDistance(region, 0#, -550#, 0#, -1#, distance)
    TestCheck "label.rectangle.bottom", found And Abs(distance - 50#) < 0.000001
    found = query.LabelBoundaryDistance(region, 250#, 0#, 1#, 0#, distance)
    TestCheck "label.rectangle.right", found And Abs(distance - 50#) < 0.000001
    found = query.LabelBoundaryDistance(region, -250#, 0#, -1#, 0#, distance)
    TestCheck "label.rectangle.left", found And Abs(distance - 50#) < 0.000001
    found = query.LabelBoundaryDistance(region, 0#, 605#, 0#, 1#, distance)
    TestCheck "label.import.outsideAnchor", found And Abs(distance + 5#) < 0.000001
    found = query.LabelBoundaryDistance(region, 0#, 600#, 0#, 1#, distance)
    TestCheck "label.onBoundary", found And Abs(distance) < 0.000001
    coordinates(1, 1) = -100#: coordinates(1, 2) = -400#: coordinates(2, 1) = 100#: coordinates(2, 2) = -400#
    coordinates(3, 1) = 100#: coordinates(3, 2) = 400#: coordinates(4, 1) = -100#: coordinates(4, 2) = 400#
    Set opening = query.PolygonRegion(coordinates)
    Set region = query.SubtractContainedRegion(region, opening)
    found = query.LabelBoundaryDistance(region, 150#, 0#, -1#, 0#, distance)
    TestCheck "label.opening.face", found And Abs(distance - 50#) < 0.000001
    found = query.LabelBoundaryDistance(region, 150#, 0#, 1#, 0#, distance)
    TestCheck "label.outer.sameWall", found And Abs(distance - 150#) < 0.000001
    Set region = query.CircleRegion(0#, 0#, 300#)
    found = query.LabelBoundaryDistance(region, 0#, 250#, 0#, 1#, distance)
    TestCheck "label.circle.radius", found And Abs(distance - 50#) < 0.000001
    found = query.LabelBoundaryDistance(region, 100#, 200#, 0#, 1#, distance)
    TestCheck "label.arc.exactRay", found And Abs(distance - (Sqr(80000#) - 200#)) < 0.000001
    found = query.LabelBoundaryDistance(region, 0#, 305#, 1#, 0#, distance)
    TestCheck "label.noIntersection", Not found
    On Error Resume Next
    found = query.LabelBoundaryDistance(region, 0#, 250#, 0#, 0#, distance)
    errorNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "label.zeroNormal.rejected", errorNumber <> 0
    Exit Sub
Failed:
    TestCheck "label.geometry.runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимые координаты линии подписи: граница y=600, Offset=40.
' Для Inside сначала находим границу, затем отступаем внутрь; as не меняет знак.
Private Sub TestLabelBoundaryLayout()
    Dim config As Object, saved As Variant, table As Object, settings As CSystemSettingsReader
    Dim layout As CPlotAnnotationLayout, row As Long, placement As Variant, unit As Variant, factor As Double, expected As Double
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    Set layout = New CPlotAnnotationLayout
    layout.Initialize -400#, 400#, -800#, 800#, 0#, 0#, 800#, 1600#, "mm"
    layout.AddRebarLabel -100#, 550#, 100#, 550#, 0#, 1#, "DEFAULT", 50#
    TestCheck "label.layout.default30", Abs(layout.Y1(1) - 170#) < 0.000001
    Set table = ThisWorkbook.Names("rngPlotAnnotationSettings").RefersToRange
    For row = 2 To table.Rows.Count
        Select Case CStr(table.Cells(row, 1).Value2)
            Case "Enabled", "LineEnabled": table.Cells(row, 2).Value2 = "Yes"
            Case "Offset": table.Cells(row, 2).Value2 = 40#
        End Select
    Next row
    For Each placement In Array("Outside", "Inside")
        For row = 2 To table.Rows.Count
            If CStr(table.Cells(row, 1).Value2) = "Placement" Then table.Cells(row, 2).Value2 = CStr(placement)
        Next row
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        For Each unit In Array("mm", "cm", "m")
            Select Case CStr(unit): Case "mm": factor = 1#: Case "cm": factor = 10#: Case "m": factor = 1000#: End Select
            Set layout = New CPlotAnnotationLayout: layout.ConfigureFromSettings settings
            layout.Initialize -400# / factor, 400# / factor, -800# / factor, 800# / factor, 0#, 0#, 800#, 1600#, CStr(unit)
            layout.AddRebarLabel -100# / factor, 550# / factor, 100# / factor, 550# / factor, 0#, 2#, "TOTAL", 50# / factor
            If CStr(placement) = "Outside" Then expected = 160# Else expected = 240#
            TestCheck "label.layout." & CStr(placement) & "." & CStr(unit), layout.Count = 2 And Abs(layout.Y1(1) - expected) < 0.000001 And Abs(layout.Y2(1) - expected) < 0.000001
        Next unit
    Next placement
    GoTo Restore
Failed:
    TestCheck "label.layout.runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    If IsArray(saved) Then TestFixtureRestoreInputs config, saved
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Каждый ряд задан собственным независимым правилом и своим as: 50/100/150.
' Повторные привязки стороны должны дать одну подпись всех её физических стержней.
Private Sub TestIndependentRowLabels(ByVal rows As Long, ByVal mixedDiameter As Boolean, ByVal outputDirectory As String)
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim builder As CSectionBuilder, model As CSectionModel, mesh As TSectionMeshOptions, i As Long, diameter As Double
    Dim expected As String, prefix As String, count As Long, settings As CSystemSettingsReader, units As CUnitSystem
    Dim properties As CSectionPropertiesCalculator, writer As CNDMResultsWriter, reader As CSectionPlotDataReader
    Dim config As Object, saved As Variant, table As Object, row As Long, outputUnit As Variant, labelIndex As Long, labelCount As Long
    Dim annotations As CSectionAnnotations, labels As CRebarGroupAnnotationBuilder
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    prefix = "rows." & CStr(rows) & ".mixed" & CStr(mixedDiameter)
    Set facade = New CSectionDefinitionBuilder: facade.Initialize prefix
    Set path = facade.RectanglePath("Outer", -300#, -600#, 600#, 1200#): facade.Outer "OUTER", path.ID
    For i = 1 To rows
        Set path = facade.LinePath("Row" & CStr(i), -200#, 600# - 50# * i, 200#, 600# - 50# * i)
        Set options = New CRebarRuleOptions: options.AnnotationGroup = "Top"
        diameter = 20#: If mixedDiameter And i = rows Then diameter = 16#
        facade.RebarByCount "Rule" & CStr(i), path.ID, i + 2, diameter, options
        count = count + i + 2
    Next i
    facade.RebarRowLabel "Top", "Rule1", 0#, 1#
    mesh.StepX = 200#: mesh.StepY = 200#: mesh.BoundarySubdivisions = 2: mesh.MaterialID = 1
    Set builder = New CSectionBuilder: Set model = builder.Build(facade.Finish, mesh)
    TestCheck prefix & ".physicalTotal", model.RebarCount = count
    TestCheck prefix & ".singleLabel", model.Annotations.Count = 1
    If mixedDiameter Then
        expected = CStr(count - rows - 2) & ChrW$(216) & "20 + " & CStr(rows + 2) & ChrW$(216) & "16"
    Else
        expected = CStr(count) & ChrW$(216) & "20"
    End If
    TestCheck prefix & ".allDiameters", model.Annotations.Text(1) = expected
    Dim anchors(1 To 2) As TSectionAnnotation
    anchors(1).Name = "Top": anchors(1).StartX = -200#: anchors(1).EndX = 200#: anchors(1).StartY = 550#: anchors(1).EndY = 550#: anchors(1).NormalY = 1#
    anchors(2) = anchors(1): anchors(2).StartY = 500#: anchors(2).EndY = 500#
    Set annotations = New CSectionAnnotations: Set labels = New CRebarGroupAnnotationBuilder
    labels.AddRebarLabels annotations, builder.Rebars, anchors, 2
    TestCheck prefix & ".duplicateAnchor", annotations.Count = 1 And annotations.Text(1) = expected
    Set properties = New CSectionPropertiesCalculator: properties.PrepareSnapshot model
    Set table = ThisWorkbook.Names("rngUnitSettings").RefersToRange
    row = TestHeaderRow(table, "Length", 1)
    Set reader = New CSectionPlotDataReader
    For Each outputUnit In Array("mm", "cm", "m")
        table.Cells(row, 4).Value2 = CStr(outputUnit)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, model, properties, units
        table.Cells(row, 4).Value2 = "cm" ' После записи меняем OUTPUT: читать геометрию надо в единицах снимка.
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        reader.LoadGeometryPreviewFromWorkbook ThisWorkbook, settings
        labelIndex = 0: labelCount = 0
        For i = 1 To reader.AnnotationCount
            If reader.AnnotationType(i) = "REBAR_ANNOTATION" Then labelIndex = i: labelCount = labelCount + 1
        Next i
        TestCheck prefix & ".savedTotal." & CStr(outputUnit), labelCount = 1 And reader.AnnotationText(labelIndex) = expected
        TestCheck prefix & ".savedBoundary." & CStr(outputUnit), Abs(reader.AnnotationBoundaryDistance(labelIndex) - 50#) < 0.000001 And reader.LengthUnit = "mm"
    Next outputUnit
    If Len(outputDirectory) > 0 Then
        TestCheck prefix & ".preview", Len(modTestShapeArchitecture.TestPreparedPreviewSnapshot(Replace$(prefix, ".", "_"), model, outputDirectory)) > 0
    End If
    GoTo Restore
Failed:
    TestCheck prefix & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    If IsArray(saved) Then TestFixtureRestoreInputs config, saved
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' исчезновение дуг и прямых, координаты стержней, схема и штатный расчёт.
Public Function RunRoundedEndpointTests(Optional ByVal outputDirectory As String = "") As String
    mPassed = 0: mFailed = 0: mReport = ""
    TestCollapsedBoundaries
    TestRoundedEndpoints "F02.Asymmetric", 1200#, 50#, 550#, 50#, outputDirectory
    TestRoundedEndpoints "F02.Mirrored", 1200#, 550#, 50#, 50#, outputDirectory
    TestRoundedEndpoints "F02.NearRadius", 1200#, 49.999, 550#, 50#, outputDirectory
    TestRoundedEndpoints "F02.NearOffset", 1200#, 50#, 550#, 49#, outputDirectory
    TestRoundedEndpoints "F02.Square", 1200#, 0#, 0#, 50#, outputDirectory
    TestRoundedEndpoints "F02.Ordinary", 1200#, 100#, 200#, 50#, outputDirectory
    TestRoundedEndpoints "F02.LeftLineGone", 600#, 300#, 50#, 50#, outputDirectory
    TestRoundedEndpoints "F02.RightLineGone", 600#, 50#, 300#, 50#, outputDirectory
    RunRoundedEndpointTests = mReport & "TOTAL_ROUNDED_ENDPOINTS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' общие привязки с произвольными именами, без рецепта сечения.
Private Sub TestCollapsedBoundaries()
    Dim polygonSpec As CPolygonSpec, polygonPoint As Variant
    On Error GoTo Failed
    Dim query As CSectionGeometryQuery, path As CGeometryPath, copy As CGeometryPath, edge As TRegionEdge
    Dim i As Long, x As Double, y As Double, tx As Double, ty As Double, errorNumber As Long, station As Double
    Set query = New CSectionGeometryQuery
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "A", -300#, 600#, 50#
    polygonSpec.AddVertex "B", 300#, 600#, 550#
    polygonSpec.AddVertex "C", 300#, -600#, 550#
    polygonSpec.AddVertex "D", -300#, -600#, 50#
    polygonSpec.Seal
    Set path = query.OffsetRoundedPolygon("Generic", polygonSpec, 50#)
    TestCheck "core.onlyFourPhysicalSegments", path.SegmentCount = 4
    TestCheck "core.noPhantomCorner", Not path.HasSegmentID("Corner.A") And Not path.HasSegmentID("Corner.D")
    TestCheck "core.noPhantomEdge", Not path.HasSegmentID("Edge.A") And Not path.HasSegmentID("Edge.C")
    TestCheck "core.collapsedParameters", path.AnchorStation("Edge.A", 0#) = path.AnchorStation("Edge.A", 1#)
    TestCheck "core.consecutiveCollapsed", path.AnchorStation("Corner.A") = path.AnchorStation("Edge.A")
    station = path.AnchorStation("Corner.D")
    TestCheck "core.secondJoint", Abs(station - (100# + GEOM_PI * 500#)) < 0.000001
    For i = 1 To path.SegmentCount
        path.CopySegment i, edge
        If edge.Sweep <> 0# Then
            TestCheck "core.analyticRadius." & CStr(i), Abs(Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2) - 500#) < 0.000001
            TestCheck "core.analyticSweep." & CStr(i), Abs(edge.Sweep + GEOM_PI / 2#) < 0.000000001
        End If
    Next i
    Set copy = path.WithID("Copy")
    TestCheck "core.copyAlias", copy.AnchorStation("Corner.D") = station
    Set copy = path.Reversed
    TestCheck "core.reverseAlias", Abs(copy.AnchorStation("Corner.D") - 1100#) < 0.000001
    Set copy = path.Transformed(1000000#, -1000000#, GEOM_PI / 2#, True)
    copy.AtDistance copy.AnchorStation("Edge.A"), x, y, tx, ty
    TestCheck "core.transformedAlias", Abs(x - 1000550#) < 0.000001 And Abs(y + 1000250#) < 0.000001
    Set copy = path.Slice(0#, 300#)
    TestCheck "core.sliceAlias", copy.AnchorStation("Corner.A") = 0# And copy.SegmentCount = 1
    Set copy = path.BetweenAnchors("Full", "Edge.D", 0#, "Edge.D", 0#, True, True)
    TestCheck "core.wrapAlias", Abs(copy.AnchorStation("Edge.A") - 1100#) < 0.000001 And Abs(copy.Length - path.Length) < 0.000001
    On Error Resume Next
    Set copy = path.BetweenAnchors("Empty", "Edge.A", 0#, "Edge.A", 1#)
    errorNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "core.emptyStillRejected", errorNumber <> 0
    TestCheck "core.emptyReason", path.LastDiagnostic.Code = "SEGMENT_DISAPPEARED" And InStr(path.LastDiagnostic.Reason, "После скругления") > 0
    On Error Resume Next
    station = path.AnchorStation("Misspelled")
    errorNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "core.unknownStillRejected", errorNumber <> 0 And path.LastDiagnostic.Code = "MISSING_SEGMENT"
    Exit Sub
Failed:
    TestCheck "core.runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' исчезнувшая наружная сторона, сторона отверстия и горизонтальная
' сторона скруглённого сечения. Проверяет причину, источники и отключение ряда.
Public Function RunCollapsedInputMessageTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim scenario As Long, mode As Long, message As String, prefix As String, captions As Variant
    captions = Array("левой наружной", "правой наружной", "левой стороны отверстия", "правой стороны отверстия", "верхняя", "нижняя")
    For scenario = 0 To 5
        For mode = 0 To 4
            prefix = "collapsed." & CStr(scenario) & "." & CStr(mode)
            message = TestCollapsedInputMessage(scenario, mode)
            If mode = 0 Then
                TestCheck prefix & ".inputError", Left$(message, 11) = "INPUT_ERROR"
                TestCheck prefix & ".cause", InStr(message, "не осталось участка для размещения арматуры") > 0
                TestCheck prefix & ".side", InStr(message, CStr(captions(scenario))) > 0
                TestCheck prefix & ".geometrySource", InStr(message, "Geometry[") > 0
                TestCheck prefix & ".rebarSource", InStr(message, "Geometry[") > 0
                TestCheck prefix & ".action", InStr(message, "задайте диаметр 0") > 0 And InStr(message, "fullLoopIfSame") = 0
            Else
                TestCheck prefix & ".accepted", message = "OK"
            End If
        Next mode
    Next scenario
    message = TestCollapsedExtraRowMessage()
    TestCheck "collapsed.extra.inputError", Left$(message, 11) = "INPUT_ERROR"
    TestCheck "collapsed.extra.row", InStr(message, "контур для 2-го ряда арматуры") > 0
    TestCheck "collapsed.extra.cause", InStr(message, "проверочный контур исчез целиком") > 0
    TestCheck "collapsed.extra.values", InStr(message, "радиус первой дуги: 150 мм") > 0 And _
        InStr(message, "отступ внутрь этой дуги: 240 мм") > 0
    TestCheck "collapsed.extra.source", InStr(message, "Geometry[") > 0
    mReport = mReport & "OBSERVED_COLLAPSED_EXTRA: " & Replace(message, vbCrLf, " | ") & vbCrLf
    RunCollapsedInputMessageTests = mReport & "TOTAL_COLLAPSED_INPUT_MESSAGES: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' слишком большое смещение второго ряда на дуге должно сохранить
' номер ряда и источники его параметров в сообщении общего исполнителя.
Private Function TestCollapsedExtraRowMessage() As String
    On Error GoTo Failed
    Dim definition As CSectionDefinition, layout As CRebarLayout
    Set definition = TestFixtureCircleDefinition(Array(300#, 40#, 8, 20#, 16#, 0#, "Stacked", "Stacked", "EachBar", "EachBar", 200#, 0#))
    Set layout = TestFixtureLayout(definition)
    TestCollapsedExtraRowMessage = "UNEXPECTED_SUCCESS"
    Exit Function
Failed:
    If Err.Number = vbObjectError + 5530 Then
        TestCollapsedExtraRowMessage = "INPUT_ERROR: " & Err.Description
    Else
        TestCollapsedExtraRowMessage = "UNEXPECTED: " & CStr(Err.Number) & "; " & Err.Description
    End If
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' mode 0 — включённый ряд; 1 — нулевое количество; 2/3 — нулевой/пустой
' диаметр; 4 — включённый ряд при радиусе, оставляющем прямой участок ненулевой длины.
Private Function TestCollapsedInputMessage(ByVal scenario As Long, ByVal mode As Long) As String
    On Error GoTo Failed
    Dim definition As CSectionDefinition, active As Variant, selected As Variant, faces As Variant, radius As Double
    active = Array(40#, 20#, 3, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")
    selected = active: radius = 200#
    If mode = 1 Then selected(2) = 0
    If mode = 2 Then selected(1) = 0#
    If mode = 3 Then selected(1) = Empty
    If mode = 4 Then radius = 199#
    If scenario < 4 Then
        faces = Array(Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty)
        If scenario < 2 Then
            faces(scenario) = selected
            Set definition = TestFixtureHollowDefinition(Array(600#, 400#, radius, 100#, 100#, 20#), faces)
        Else
            faces(scenario - 2) = active
            ' У отверстия количество определяется проекцией; выключаем исходный ряд,
            ' если проверяется нулевое количество, а диаметр отверстия оставляем заданным.
            If mode = 1 Then faces(scenario - 2) = selected
            If mode = 4 Then radius = 99# Else radius = 100#
            faces(scenario + 2) = selected
            Set definition = TestFixtureHollowDefinition(Array(1000#, 1000#, 50#, 400#, 200#, radius), faces)
        End If
    Else
        faces = Array(Empty, Empty, Empty, Empty): faces(scenario - 2) = selected
        Set definition = TestFixtureRoundedDefinition(Array(400#, 1000#, 0#, 0#, radius, radius, 0#, 0#), faces)
    End If
    TestCollapsedInputMessage = "OK"
    Exit Function
Failed:
    If Err.Number = vbObjectError + 5530 Then
        TestCollapsedInputMessage = "INPUT_ERROR: " & Err.Description
    Else
        TestCollapsedInputMessage = "UNEXPECTED: " & CStr(Err.Number) & "; " & Err.Description
    End If
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' читает настоящие именованные диапазоны в двух новых местах листа.
' Ошибка должна ссылаться на геометрию и арматуру нужного контура, без концевых отступов.
Public Function RunCollapsedInputAddressTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim originalGeometry As String, originalRebar As String, config As Object, saved As Variant
    Dim probe As Object, geometry As Object, rebar As Object, geometryValues As Variant, rebarValues As Variant
    Dim units As CUnitSystem, location As Long, opening As Long, key As Variant, field As Variant, message As String
    Dim expected As String, prefix As String, alerts As Boolean
    originalGeometry = ThisWorkbook.Names("rngHollowRectangleGeometry").RefersTo
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    alerts = Application.DisplayAlerts
    TestFixtureSelectRecipeType ThisWorkbook, "HollowRectangle"
    Application.Calculate
    geometryValues = ThisWorkbook.Names("rngHollowRectangleGeometry").RefersToRange.Value2
    Set probe = ThisWorkbook.Worksheets.Add
    Set units = New CUnitSystem: units.InitializeDefaults
    For location = 0 To 1
        Set geometry = probe.Cells(40 + 160 * location, 2 + 28 * location).Resize(UBound(geometryValues, 1), UBound(geometryValues, 2))
        geometry.Value2 = geometryValues
        ThisWorkbook.Names("rngHollowRectangleGeometry").RefersTo = "='" & probe.Name & "'!" & geometry.Address
        For Each key In Array("Top", "Bottom", "Left", "Right", "OpeningTop", "OpeningBottom", "OpeningLeft", "OpeningRight")
            TestFixtureInputCell(geometry, "Main", CStr(key), "as").Value2 = 40#
            For Each field In Array("d", "n", "s0", "s1")
                TestFixtureInputCell(geometry, "Main", CStr(key), CStr(field)).Value2 = 0#
            Next field
            For Each field In Array("d2", "d3")
                TestFixtureInputCell(geometry, "Extra", CStr(key), CStr(field)).Value2 = 0#
            Next field
        Next key
        TestFixtureInputCell(geometry, "Main", "Left", "d").Value2 = 20#
        TestFixtureInputCell(geometry, "Main", "Left", "n").Value2 = 3
        TestFixtureInputCell(geometry, "General", "Xo").Value2 = 0#
        TestFixtureInputCell(geometry, "General", "Yo").Value2 = 0#
        For opening = 0 To 1
            prefix = "collapsed.address." & CStr(location) & "." & CStr(opening)
            TestFixtureInputCell(geometry, "General", "B").Value2 = IIf(opening = 0, 600#, 1000#)
            TestFixtureInputCell(geometry, "General", "H").Value2 = IIf(opening = 0, 400#, 1000#)
            TestFixtureInputCell(geometry, "General", "R").Value2 = IIf(opening = 0, 200#, 50#)
            TestFixtureInputCell(geometry, "General", "Bo").Value2 = IIf(opening = 0, 100#, 400#)
            TestFixtureInputCell(geometry, "General", "Ho").Value2 = IIf(opening = 0, 100#, 200#)
            TestFixtureInputCell(geometry, "General", "Ro").Value2 = IIf(opening = 0, 20#, 100#)
            If opening = 1 Then TestFixtureInputCell(geometry, "Main", "OpeningLeft", "d").Value2 = 16#
            message = TestNativeCollapsedInputMessage(units)
            TestCheck prefix & ".inputError", Left$(message, Len(CStr(vbObjectError + 5530))) = CStr(vbObjectError + 5530)
            TestCheck prefix & ".cause", InStr(message, "не осталось участка для размещения арматуры") > 0
            If opening = 0 Then key = "Left" Else key = "OpeningLeft"
            For Each field In Array("as", "d", "s0", "s1", "n")
                expected = probe.Name & "!" & TestFixtureInputCell(geometry, "Main", CStr(key), CStr(field)).Address(False, False)
                If field = "s0" Or field = "s1" Or (opening = 1 And field = "n") Then
                    TestCheck prefix & ".excluded." & CStr(field), InStr(message, expected & ";") = 0 And InStr(message, expected & ":") = 0
                Else
                    TestCheck prefix & ".included." & CStr(field), InStr(message, expected) > 0
                End If
            Next field
            For Each field In Array("B", "H", "R")
                If opening = 0 Then key = field Else key = CStr(field) & "o"
                expected = probe.Name & "!" & TestFixtureInputCell(geometry, "General", CStr(key)).Address(False, False)
                TestCheck prefix & ".geometry." & CStr(key), InStr(message, expected) > 0
            Next field
        Next opening
    Next location
    GoTo Restore
Failed:
    TestCheck "collapsed.address.runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    On Error Resume Next
    ThisWorkbook.Names("rngHollowRectangleGeometry").RefersTo = originalGeometry
    If IsArray(saved) Then TestFixtureRestoreInputs config, saved
    Application.DisplayAlerts = False
    If Not probe Is Nothing Then probe.Delete
    Application.DisplayAlerts = alerts
    On Error GoTo 0
    RunCollapsedInputAddressTests = mReport & "TOTAL_COLLAPSED_INPUT_ADDRESSES: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' перехватывает только ожидаемую ошибку ввода рабочего рецепта.
Private Function TestNativeCollapsedInputMessage(ByVal units As CUnitSystem) As String
    On Error GoTo Failed
    Dim definition As CSectionDefinition
    Set definition = TestFixtureWorkbookDefinition("HollowRectangle", ThisWorkbook, units)
    TestNativeCollapsedInputMessage = "UNEXPECTED_SUCCESS"
    Exit Function
Failed:
    TestNativeCollapsedInputMessage = CStr(Err.Number) & ": " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' точные снимки сечений без выступов, с двумя выступами и с одним выступом до и после правки.
Public Function TestRunRoundedOrdinarySnapshots() As String
    Dim config As Object, saved As Variant, fixture As Variant, settings As CSystemSettingsReader
    Dim units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel, result As String
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    For Each fixture In Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft")
        modTestShapeArchitecture.TestConfigureShapeFixture CStr(fixture)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
        result = result & modTestShapeArchitecture.TestPreparedShapeSnapshot(CStr(fixture), model, builder.Mesh, builder.Rebars)
    Next fixture
    result = result & "TOTAL_ROUNDED_ORDINARY: passed=4; failed=0" & vbCrLf
    GoTo Restore
Failed:
    result = result & "FAIL: ordinary.runtime." & CStr(Err.Number) & "." & Err.Description & vbCrLf
Restore:
    If IsArray(saved) Then TestFixtureRestoreInputs config, saved
    TestRunRoundedOrdinarySnapshots = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Для стороны без выступа независимый эталон — две четверти окружности и вертикаль.
' Проверяются реальные центры стержней, а не только число и длина пути.
Private Sub TestRoundedEndpoints(ByVal label As String, ByVal height As Double, ByVal leftRadius As Double, _
        ByVal rightRadius As Double, ByVal cover As Double, ByVal outputDirectory As String)
    Dim config As Object, saved As Variant, table As Object, main As Long, extra As Long, row As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, definition As CSectionDefinition
    Dim builder As CSectionBuilder, model As CSectionModel, path As CGeometryPath
    Dim factor As Double, side As Long, k As Long, i As Long, count As Long, group As String
    Dim radius As Double, effectiveRadius As Double, vertical As Double, length As Double, distance As Double
    Dim x As Double, y As Double, expectedX As Double, expectedY As Double, tx As Double, ty As Double
    Dim angle As Double, direction As Double, top As Double, outsideX As Double, startX As Double
    Dim passed As Long, failed As Long, result As String
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    modTestShapeArchitecture.TestConfigureShapeFixture "RoundedSimple"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    Set table = ThisWorkbook.Names("rngPointedSectionGeometry").RefersToRange
    TestFixtureInputCell(table, "General", "B").Value2 = 600# / factor
    TestFixtureInputCell(table, "General", "H").Value2 = height / factor
    TestFixtureInputCell(table, "General", "Left.W").Value2 = 0#
    TestFixtureInputCell(table, "General", "Right.W").Value2 = 0#
    TestFixtureInputCell(table, "General", "Left.R1").Value2 = leftRadius / factor
    TestFixtureInputCell(table, "General", "Right.R1").Value2 = rightRadius / factor
    TestConfigureReviewRebar table, Array("Left", "Right", "Top", "Bottom"), cover, factor, True
    Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
    Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
    TestCheck label & ".tenBars", model.RebarCount = 10
    TestCheck label & ".outerArea", Abs(definition.ConcreteDomain.Area - (600# * height - (2# - GEOM_PI / 2#) * (leftRadius ^ 2 + rightRadius ^ 2))) < 0.00001
    For side = 1 To 2
        direction = IIf(side = 1, -1#, 1#): group = IIf(side = 1, "H.Left", "H.Right")
        radius = IIf(side = 1, leftRadius, rightRadius): effectiveRadius = GeomMax(0#, radius - cover)
        top = height / 2# - cover: outsideX = direction * (300# - cover)
        startX = outsideX - direction * effectiveRadius
        vertical = 2# * top - 2# * effectiveRadius: length = vertical + GEOM_PI * effectiveRadius
        Set path = definition.Path(group)
        TestCheck label & ".length." & group, Abs(path.Length - length) < 0.000001
        path.AtDistance 0#, x, y, tx, ty
        TestCheck label & ".start." & group, Abs(x - startX) < 0.000001 And Abs(y - top) < 0.000001
        path.AtDistance path.Length, x, y, tx, ty
        TestCheck label & ".end." & group, Abs(x - startX) < 0.000001 And Abs(y + top) < 0.000001
        For k = 1 To 5
            distance = length * k / 6#
            If effectiveRadius > 0# And distance < GEOM_PI * effectiveRadius / 2# Then
                angle = distance / effectiveRadius
                expectedX = startX + direction * effectiveRadius * Sin(angle)
                expectedY = top - effectiveRadius + effectiveRadius * Cos(angle)
            ElseIf distance <= GEOM_PI * effectiveRadius / 2# + vertical Or effectiveRadius = 0# Then
                expectedX = outsideX: expectedY = top - effectiveRadius - (distance - GEOM_PI * effectiveRadius / 2#)
            Else
                angle = (distance - GEOM_PI * effectiveRadius / 2# - vertical) / effectiveRadius
                expectedX = startX + direction * effectiveRadius * Cos(angle)
                expectedY = -top + effectiveRadius - effectiveRadius * Sin(angle)
            End If
            path.AtDistance distance, x, y, tx, ty
            TestCheck label & ".position." & group & "." & CStr(k), Abs(x - expectedX) < 0.000001 And Abs(y - expectedY) < 0.000001
            count = 0
            For i = 1 To model.RebarCount
                If Abs(model.RebarX(i) - expectedX) < 0.000001 And Abs(model.RebarY(i) - expectedY) < 0.000001 Then count = count + 1
            Next i
            TestCheck label & ".physical." & group & "." & CStr(k), count = 1
            mReport = mReport & "BAR: " & label & "; " & group & "; " & CStr(k) & "; x=" & CStr(expectedX) & "; y=" & CStr(expectedY) & vbCrLf
        Next k
    Next side
    result = modTestSectionRecipes.TestReviewModelEquilibrium(label, model, settings, units, passed, failed)
    mPassed = mPassed + passed: mFailed = mFailed + failed: mReport = mReport & result
    If Len(outputDirectory) > 0 Then
        result = modTestShapeArchitecture.TestPreparedPreviewSnapshot(Replace$(label, ".", "_"), model, outputDirectory)
        TestCheck label & ".preview", Len(result) > 0
    End If
    TestCheckRoundedWorkbookRun label, units, model
    GoTo Restore
Failed:
    TestCheck label & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    If IsArray(saved) Then TestFixtureRestoreInputs config, saved
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет распределение по независимым путям, предельные радиусы, отключённые стороны и понятные ошибки исчезнувших граней.
Public Function RunGeometryReviewTests(Optional ByVal outputDirectory As String = "") As String
    mPassed = 0: mFailed = 0: mReport = ""
    TestPerPath 3, False: TestPerPath 3, True
    TestPerPath 15, False: TestPerPath 15, True
    TestRoundedLimit 1000#, outputDirectory: TestRoundedLimit 600#, outputDirectory
    TestRoundedLimit 600.001, outputDirectory
    TestDisabledSides "Rectangle", "rngRectanglesGeometry", Array("H1.as_1", "H1.as_2", "B1.as_1", "B1.as_2")
    TestDisabledSides "RoundedSimple", "rngPointedSectionGeometry", Array("H.Left", "H.Right", "B.Top", "B.Bottom")
    TestDisabledSides "HollowCentered", "rngHollowRectangleGeometry", Array("H.Left", "H.Right", "B.Top", "B.Bottom")
    Dim result As String
    result = mReport & "TOTAL_GEOMETRY_REVIEW: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
    result = result & RunCollapsedInputMessageTests()
    RunGeometryReviewTests = result & RunCollapsedInputAddressTests()
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Две независимые линии имеют собственную нумерацию позиций 1, 3, ...
' Смена порядка линий меняет порядок записи стержней, а не их координаты.
Private Sub TestPerPath(ByVal n As Long, ByVal reverse As Boolean)
    On Error GoTo Failed
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim builder As CSectionBuilder, model As CSectionModel, mesh As TSectionMeshOptions
    Dim i As Long, first As Long, second As Long, id As String, ids As Object, label As String
    label = "F01.n" & CStr(n) & ".reverse" & CStr(reverse)
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "ReviewPerPath"
    Set path = facade.CirclePath("Outer", 0#, 0#, 1000#): facade.Outer "OUTER", path.ID
    Set path = facade.LinePath("P1", 0#, 0#, 100#, 0#)
    Set path = facade.LinePath("P2", 0#, 200#, 100#, 200#)
    If reverse Then facade.PathGroup "Sides", Array("P2", "P1") Else facade.PathGroup "Sides", Array("P1", "P2")
    Set options = New CRebarRuleOptions
    facade.RebarByCount "Main", "Sides", n, 10#, options
    Set options = New CRebarRuleOptions: options.NormalOffset = 20#: options.EverySecond = True
    facade.RebarDependent "Extra", "Main", 8#, options
    mesh.StepX = 1000#: mesh.StepY = 1000#: mesh.BoundarySubdivisions = 1: mesh.MaterialID = 1
    Set builder = New CSectionBuilder: Set model = builder.Build(facade.Finish, mesh)
    Set ids = CreateObject("Scripting.Dictionary")
    For i = 1 To builder.Rules.PositionCount
        id = builder.Rules.PositionSourceID(i)
        TestCheck label & ".unique." & CStr(i), Not ids.Exists(id)
        ids.Add id, True
        If builder.Rules.PositionRuleID(i) = "Extra" Then
            If builder.Rules.PositionPathID(i) = "P1" Then first = first + 1 Else second = second + 1
        End If
    Next i
    TestCheck label & ".P1", first = (n + 1) \ 2
    TestCheck label & ".P2", second = (n + 1) \ 2
    TestCheck label & ".physical", model.RebarCount = 2 * n + n + 1
    mReport = mReport & "PER_PATH: " & label & "; extra=" & CStr(first) & "+" & CStr(second) & "; physical=" & CStr(model.RebarCount) & vbCrLf
    Exit Sub
Failed:
    TestCheck label & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Отключает по одной стороне через n=0 и d=0. Подписи остальных сторон
' обязаны совпасть с их фактическими стержнями; размеры сечения сохраняются.
Private Sub TestDisabledSides(ByVal fixture As String, ByVal tableName As String, ByVal groups As Variant)
    Dim config As Object, saved As Variant, table As Object, main As Long, extra As Long, side As Long, mode As Long, row As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim labels As Object, counts As Object, diameters As Object, group As Variant, i As Long, key As String, factor As Double
    Dim baseDimensions As Long, dimensions As Long, label As String, keys As Variant, field As String
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    modTestShapeArchitecture.TestConfigureShapeFixture "Circle": modTestShapeArchitecture.TestConfigureShapeFixture fixture
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    Set table = ThisWorkbook.Names(tableName).RefersToRange
    If tableName = "rngRectanglesGeometry" Then keys = Array("R1.Left", "R1.Right", "R1.Top", "R1.Bottom") Else keys = Array("Left", "Right", "Top", "Bottom")
    TestConfigureReviewRebar table, keys, 50#, factor, False
    Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
    For i = 1 To model.AnnotationCount
        If model.Annotations.AnnotationType(i) = "DIMENSION" Then baseDimensions = baseDimensions + 1
    Next i
    For side = 1 To 4
        For mode = 3 To 4
            label = "disabledSide." & fixture & "." & CStr(groups(side - 1)) & ".column" & CStr(mode)
            If mode = 3 Then field = "d" Else field = "n"
            TestFixtureInputCell(table, "Main", CStr(keys(side - 1)), field).Value2 = 0
            Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
            Set labels = CreateObject("Scripting.Dictionary"): Set counts = CreateObject("Scripting.Dictionary")
            Set diameters = CreateObject("Scripting.Dictionary")
            For i = 1 To model.RebarCount
                key = builder.Rebars.BarAnnotationGroupName(i)
                If Not counts.Exists(key) Then counts.Add key, 0: diameters.Add key, builder.Rebars.Diameter(i)
                counts(key) = CLng(counts(key)) + 1
            Next i
            dimensions = 0
            For i = 1 To model.AnnotationCount
                If model.Annotations.AnnotationType(i) = "REBAR_ANNOTATION" Or model.Annotations.AnnotationType(i) = "REBAR_LOCAL_ANNOTATION" Then
                    key = Mid$(model.Annotations.AnnotationID(i), 7): labels.Add key, model.Annotations.Text(i)
                    TestCheck label & ".noEmptyLabel." & key, counts.Exists(key)
                End If
                If model.Annotations.AnnotationType(i) = "DIMENSION" Then dimensions = dimensions + 1
            Next i
            TestCheck label & ".off", Not labels.Exists(CStr(groups(side - 1))) And Not counts.Exists(CStr(groups(side - 1)))
            TestCheck label & ".dimensions", dimensions = baseDimensions
            For Each group In counts.Keys
                TestCheck label & ".otherLabel." & CStr(group), labels.Exists(CStr(group))
                If labels.Exists(CStr(group)) Then TestCheck label & ".quantity." & CStr(group), _
                    CStr(labels(group)) = CStr(counts(group)) & ChrW$(216) & CStr(diameters(group))
            Next group
            TestFixtureInputCell(table, "Main", CStr(keys(side - 1)), "d").Value2 = 20# / factor
            TestFixtureInputCell(table, "Main", CStr(keys(side - 1)), "n").Value2 = 4
        Next mode
    Next side
    GoTo Restore
Failed:
    TestCheck label & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    If IsArray(saved) Then TestFixtureRestoreInputs config, saved
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Прямые B исчезают при R1=B/2; боковой аналитический путь сохраняется.
' Ожидаемая длина — остаточная вертикаль плюс две четверти окружности R-as.
Private Sub TestRoundedLimit(ByVal height As Double, ByVal outputDirectory As String)
    Dim config As Object, saved As Variant, table As Object, header As Long, row As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder
    Dim model As CSectionModel, label As String, i As Long, expected As Double
    Dim definition As CSectionDefinition, paths As Variant, path As CGeometryPath, factor As Double, extra As Long
    Dim passed As Long, failed As Long, result As String
    On Error GoTo Failed
    label = "F02.H" & CStr(height)
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    modTestShapeArchitecture.TestConfigureShapeFixture "RoundedSimple"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    Set table = ThisWorkbook.Names("rngPointedSectionGeometry").RefersToRange
    TestFixtureInputCell(table, "General", "B").Value2 = 600# / factor
    TestFixtureInputCell(table, "General", "H").Value2 = height / factor
    TestFixtureInputCell(table, "General", "Left.W").Value2 = 0#
    TestFixtureInputCell(table, "General", "Right.W").Value2 = 0#
    TestFixtureInputCell(table, "General", "Left.R1").Value2 = 300# / factor
    TestFixtureInputCell(table, "General", "Right.R1").Value2 = 300# / factor
    TestConfigureReviewRebar table, Array("Left", "Right", "Top", "Bottom"), 50#, factor, True
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
    expected = height - 600# + GEOM_PI * 250#
    For Each paths In Array("H.Left", "H.Right")
        Set path = definition.Path(CStr(paths))
        TestCheck label & ".length." & CStr(paths), Abs(path.Length - expected) < 0.000001
    Next paths
    Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
    TestCheck label & ".tenBars", model.RebarCount = 10
    TestCheck label & ".contour", Abs(definition.ConcreteDomain.Area - (600# * (height - 600#) + GEOM_PI * 90000#)) < 0.00001
    For i = 1 To model.AnnotationCount
        TestCheck label & ".noDisabledLabel." & CStr(i), model.Annotations.AnnotationID(i) <> "REBAR_B.Top" And model.Annotations.AnnotationID(i) <> "REBAR_B.Bottom"
    Next i
    result = modTestSectionRecipes.TestReviewModelEquilibrium(label, model, settings, units, passed, failed)
    mPassed = mPassed + passed: mFailed = mFailed + failed: mReport = mReport & result
    If Len(outputDirectory) > 0 Then
        result = modTestShapeArchitecture.TestPreparedPreviewSnapshot(Replace$(label, ".", "_"), model, outputDirectory)
        TestCheck label & ".preview", Len(result) > 0
    End If
    TestCheckRoundedWorkbookRun label, units, model
    GoTo Restore
Failed:
    TestCheck label & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    If IsArray(saved) Then TestFixtureRestoreInputs config, saved
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет штатный расчёт одного сочетания и проверяет запись НДС при заданной
' нагрузке. Профиль не включает поиск несущей способности и расчёт трещин.
Private Sub TestCheckRoundedWorkbookRun(ByVal label As String, ByVal units As CUnitSystem, ByVal model As CSectionModel)
    Dim loads As Object, profiles As Object, settings As Object, row As Long, key As String, message As String
    Dim properties As CSectionPropertiesCalculator
    Set properties = New CSectionPropertiesCalculator: properties.CalculateConcrete model
    Set settings = ThisWorkbook.Names("rngSystemSettings").RefersToRange
    For row = 2 To settings.Rows.Count
        key = CStr(settings.Cells(row, 1).Value2)
        If key = "General.ExecutionReportEnabled" Or key = "Plot.AutoUpdate" Then settings.Cells(row, 2).Value2 = "No"
    Next row
    Set profiles = ThisWorkbook.Names("rngCalculationProfiles").RefersToRange
    For row = 2 To profiles.Rows.Count
        Select Case CStr(profiles.Cells(row, 2).Value2)
            Case "Calculation.Strength.DirectState": profiles.Cells(row, 3).Value2 = "Yes"
            Case "Calculation.Strength.Capacity", "Calculation.Crack.Width", "Calculation.Stability.Enabled", "Calculation.Crack.Longitudinal"
                profiles.Cells(row, 3).Value2 = "No"
        End Select
    Next row
    Set loads = ThisWorkbook.Names("rngLoadCombinations").RefersToRange
    loads.Offset(1, 0).Resize(loads.Rows.Count - 1, loads.Columns.Count).ClearContents
    loads.Cells(2, 1).Value2 = label: loads.Cells(2, 2).Value2 = -2# * properties.Area / units.InputForceToInternal(1#)
    loads.Cells(2, 3).Value2 = 0#: loads.Cells(2, 4).Value2 = 0#: loads.Cells(2, 5).Value2 = "PR1"
    message = TestFixtureRunSectionWorkflow(ThisWorkbook, False)
    TestCheck label & ".workbook.run", InStr(1, message, "ошиб", vbTextCompare) = 0 And InStr(1, message, "InputErr", vbTextCompare) = 0
    TestCheck label & ".workbook.state", SnapshotAnchoredExtent(ThisWorkbook.Names("rngNDMElementResults").RefersToRange, False) > 1
    mReport = mReport & "WORKBOOK_RUN: " & label & "; " & message & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' находит настройку единиц по ключу строки, без фиксированного адреса.
Private Function TestHeaderRow(ByVal table As Object, ByVal text As String, ByVal column As Long) As Long
    Dim row As Long
    For row = 1 To table.Rows.Count
        If CStr(table.Cells(row, column).Value2) = text Then TestHeaderRow = row: Exit Function
    Next row
    Err.Raise 5, "modTestGeometryReview", "Не найден ключ настройки " & text
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' задаёт параметры сторон по ключам блоков Main и Extra переданной таблицы сечения.
' В проверке скруглений включены только две боковые стороны, в остальных — все четыре.
Private Sub TestConfigureReviewRebar(ByVal table As Object, ByVal keys As Variant, ByVal cover As Double, ByVal factor As Double, ByVal sidesOnly As Boolean)
    Dim index As Long, key As String, field As Variant
    For index = 0 To UBound(keys)
        key = CStr(keys(index))
        TestFixtureInputCell(table, "Main", key, "as").Value2 = cover / factor
        TestFixtureInputCell(table, "Main", key, "d").Value2 = 20# / factor
        If sidesOnly Then
            TestFixtureInputCell(table, "Main", key, "n").Value2 = IIf(index < 2, 5, 0)
        Else
            TestFixtureInputCell(table, "Main", key, "n").Value2 = 4
        End If
        For Each field In Array("s0", "s1"): TestFixtureInputCell(table, "Main", key, CStr(field)).Value2 = 0#: Next field
        For Each field In Array("d2", "d3"): TestFixtureInputCell(table, "Extra", key, CStr(field)).Value2 = 0#: Next field
    Next index
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Учитывает отдельное условие проверки и сохраняет его имя в итоговом протоколе.
Private Sub TestCheck(ByVal name As String, ByVal success As Boolean)
    If success Then mPassed = mPassed + 1: mReport = mReport & "OK: " & name & vbCrLf Else mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
End Sub


' -------------------- ДЛЯ ТЕСТОВ --------------------
' форма стороны определяется только W; единицы и зеркальное отражение
' не меняют допустимость размеров. Ошибки скруглений дают адреса, причину и действие.
Public Function RunPointedWidthInputTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, unitTable As Object, lengthCell As Object
    Dim unitName As Variant, side As Variant, item As Variant, factor As Double, label As String
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    Set unitTable = ThisWorkbook.Names("rngUnitSettings").RefersToRange
    Set lengthCell = unitTable.Cells(TestHeaderRow(unitTable, "Length", 1), 2)
    For Each unitName In Array("mm", "cm", "m")
        lengthCell.Value2 = unitName
        Select Case CStr(unitName)
            Case "mm": factor = 1#
            Case "cm": factor = 10#
            Case "m": factor = 1000#
        End Select
        For Each side In Array("Left", "Right")
            For Each item In Array(Array(0#, 40#, "inactive", True), Array(150#, 40#, 80#, True), _
                    Array(0.001, 0#, 80#, True), Array(0.000000000001, 40#, 80#, False), _
                    Array(0.000001, 0#, 0.000000000001, False), Array(1#, 1000#, 1000#, False), _
                    Array(-1#, 40#, 80#, False), Array("", 40#, 80#, False), Array("bad", 40#, 80#, False))
                label = "pointed.W." & CStr(unitName) & "." & CStr(side) & "." & CStr(item(0)) & ".R2=" & CStr(item(2))
                TestCheck label, TestPointedWidthCase(CStr(side), item, factor, label)
            Next item
        Next side
    Next unitName
    TestRecipeErrorParagraphs
    GoTo Restore
Failed:
    TestCheck "pointed.W.runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    If IsArray(saved) Then TestFixtureRestoreInputs config, saved
    RunPointedWidthInputTests = mReport & "TOTAL_POINTED_WIDTH_INPUT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет штатное чтение таблицы и построение контура, не решая НДС.
Private Function TestPointedWidthCase(ByVal side As String, ByVal item As Variant, ByVal factor As Double, ByVal label As String) As Boolean
    Dim table As Object, key As Variant, widthCell As Object, widthValue As Variant, radiusValue As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, definition As CSectionDefinition
    Dim reason As String, reference As String, valid As Boolean
    On Error GoTo Rejected
    valid = CBool(item(3))
    modTestShapeArchitecture.TestConfigureShapeFixture "RoundedSimple"
    Set table = ThisWorkbook.Names("rngPointedSectionGeometry").RefersToRange
    For Each key In Array("Top", "Bottom", "Left", "Right")
        TestFixtureInputCell(table, "Main", CStr(key), "n").Value2 = 0
    Next key
    Set widthCell = TestFixtureInputCell(table, "General", side & ".W")
    reference = widthCell.Worksheet.Name & "!" & widthCell.Address(False, False)
    widthValue = item(0): If IsNumeric(widthValue) Then widthValue = CDbl(widthValue) / factor
    widthCell.Value2 = widthValue
    TestFixtureInputCell(table, "General", side & ".R1").Value2 = CDbl(item(1)) / factor
    radiusValue = item(2): If IsNumeric(radiusValue) Then radiusValue = CDbl(radiusValue) / factor
    TestFixtureInputCell(table, "General", side & ".R2").Value2 = radiusValue
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
    TestPointedWidthCase = valid And definition.ConcreteDomain.Area > 0#
    If Not TestPointedWidthCase Then mReport = mReport & "UNEXPECTED_ACCEPT: " & label & vbCrLf
    Exit Function
Rejected:
    reason = Err.Description
    TestPointedWidthCase = Not valid And InStr(1, reason, reference, vbBinaryCompare) > 0 And _
        Left$(reason, Len("Адреса ячеек:")) = "Адреса ячеек:" And InStr(1, reason, "Возможное решение:", vbBinaryCompare) > 0
    mReport = mReport & "WIDTH_DIAGNOSTIC: " & label & vbCrLf & reason & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Смысловые абзацы сохраняют исходные координаты, идентификаторы и точные адреса.
Private Sub TestRecipeErrorParagraphs()
    Dim text As String, wrapped As String
    text = SectionRecipeErrorText("Ось стержня 'B.Top.Row2.4' вне бетона. Уменьшите смещение или измените положение ряда.", _
        "Сечения!C123; Сечения!D123", "Пустотелый прямоугольник: арматура у верхней наружной стороны, ряд 2.", "Ряд 'B.Top.Row2':")
    TestCheck "message.addresses.first", Left$(text, Len("Адреса ячеек:" & vbCrLf & "Сечения!C123; Сечения!D123")) = "Адреса ячеек:" & vbCrLf & "Сечения!C123; Сечения!D123"
    TestCheck "message.section.newline", InStr(text, vbCrLf & vbCrLf & "Пустотелый прямоугольник:" & vbCrLf & "Арматура") > 0
    TestCheck "message.bar.newline", InStr(text, "Ряд 'B.Top.Row2':" & vbCrLf & "Ось стержня 'B.Top.Row2.4'") > 0
    TestCheck "message.solution.paragraph", InStr(text, vbCrLf & vbCrLf & "Возможное решение:" & vbCrLf & "Уменьшите смещение") > 0
    wrapped = SectionRecipeErrorText(text)
    TestCheck "message.nested.stable", wrapped = text
    text = WorkbookErrorMessage("Расчёт не выполнен", -2147215974, "CRebarRuleExecutor", text)
    TestCheck "message.operation.paragraph", Left$(text, Len("Расчёт не выполнен:" & vbCrLf & vbCrLf & "Адреса ячеек:")) = "Расчёт не выполнен:" & vbCrLf & vbCrLf & "Адреса ячеек:"
    TestCheck "message.technical.paragraph", InStr(text, vbCrLf & vbCrLf & "Источник ошибки: CRebarRuleExecutor." & vbCrLf & "Код ошибки: -2147215974.") > 0
    mReport = mReport & "MESSAGE_EXAMPLE:" & vbCrLf & text & vbCrLf
    text = WorkbookErrorMessage("Расчёт не выполнен", -2147217194, "CSystemSettingsReader", _
        "Настройка Mesh.StepX (лист Config, ячейка B15): Числовое значение не задано. Введите положительное число.")
    TestCheck "message.setting.address.first", InStr(text, "Адреса ячеек:" & vbCrLf & "Config!B15" & vbCrLf & vbCrLf & "Настройка Mesh.StepX") > 0
    TestCheck "message.setting.solution", InStr(text, "Возможное решение:" & vbCrLf & "Введите положительное число.") > 0
    text = WorkbookErrorMessage("Расчёт не выполнен", -2147217194, "CSystemSettingsReader", _
        "Настройка Mesh.StepX (лист Config, раздел ""геометрия и расчетная сетка""): Строка не найдена. Восстановите строку настройки.")
    TestCheck "message.missing.setting.no.fake.address", InStr(text, "Источник данных:" & vbCrLf & "лист Config, раздел") > 0 And InStr(text, "Адреса ячеек:") = 0
End Sub
