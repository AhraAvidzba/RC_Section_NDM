Attribute VB_Name = "modTestGeometryReview"
Option Explicit
' ДЛЯ ТЕСТОВ: воспроизведение F01/F02 и проверка подписей отключённых граней.
Private mPassed As Long, mFailed As Long, mReport As String

' ДЛЯ ТЕСТОВ: бетонная граница, сторона отступа и полная подпись независимых рядов.
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
    Check "label.rectangle.as50", found And Abs(distance - 50#) < 0.000001
    found = query.LabelBoundaryDistance(region, 0#, -550#, 0#, -1#, distance)
    Check "label.rectangle.bottom", found And Abs(distance - 50#) < 0.000001
    found = query.LabelBoundaryDistance(region, 250#, 0#, 1#, 0#, distance)
    Check "label.rectangle.right", found And Abs(distance - 50#) < 0.000001
    found = query.LabelBoundaryDistance(region, -250#, 0#, -1#, 0#, distance)
    Check "label.rectangle.left", found And Abs(distance - 50#) < 0.000001
    found = query.LabelBoundaryDistance(region, 0#, 605#, 0#, 1#, distance)
    Check "label.import.outsideAnchor", found And Abs(distance + 5#) < 0.000001
    found = query.LabelBoundaryDistance(region, 0#, 600#, 0#, 1#, distance)
    Check "label.onBoundary", found And Abs(distance) < 0.000001
    coordinates(1, 1) = -100#: coordinates(1, 2) = -400#: coordinates(2, 1) = 100#: coordinates(2, 2) = -400#
    coordinates(3, 1) = 100#: coordinates(3, 2) = 400#: coordinates(4, 1) = -100#: coordinates(4, 2) = 400#
    Set opening = query.PolygonRegion(coordinates)
    Set region = query.SubtractContainedRegion(region, opening)
    found = query.LabelBoundaryDistance(region, 150#, 0#, -1#, 0#, distance)
    Check "label.opening.face", found And Abs(distance - 50#) < 0.000001
    found = query.LabelBoundaryDistance(region, 150#, 0#, 1#, 0#, distance)
    Check "label.outer.sameWall", found And Abs(distance - 150#) < 0.000001
    Set region = query.CircleRegion(0#, 0#, 300#)
    found = query.LabelBoundaryDistance(region, 0#, 250#, 0#, 1#, distance)
    Check "label.circle.radius", found And Abs(distance - 50#) < 0.000001
    found = query.LabelBoundaryDistance(region, 100#, 200#, 0#, 1#, distance)
    Check "label.arc.exactRay", found And Abs(distance - (Sqr(80000#) - 200#)) < 0.000001
    found = query.LabelBoundaryDistance(region, 0#, 305#, 1#, 0#, distance)
    Check "label.noIntersection", Not found
    On Error Resume Next
    found = query.LabelBoundaryDistance(region, 0#, 250#, 0#, 0#, distance)
    errorNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "label.zeroNormal.rejected", errorNumber <> 0
    Exit Sub
Failed:
    Check "label.geometry.runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' Независимые координаты линии подписи: граница y=600, Offset=40.
' Для Inside сначала находим границу, затем отступаем внутрь; as не меняет знак.
Private Sub TestLabelBoundaryLayout()
    Dim config As Object, saved As Variant, table As Object, settings As CSystemSettingsReader
    Dim layout As CPlotAnnotationLayout, row As Long, placement As Variant, unit As Variant, factor As Double, expected As Double
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    Set layout = New CPlotAnnotationLayout
    layout.Initialize -400#, 400#, -800#, 800#, 0#, 0#, 800#, 1600#, "mm"
    layout.AddRebarLabel -100#, 550#, 100#, 550#, 0#, 1#, "DEFAULT", 50#
    Check "label.layout.default30", Abs(layout.Y1(1) - 170#) < 0.000001
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
            Check "label.layout." & CStr(placement) & "." & CStr(unit), layout.Count = 2 And Abs(layout.Y1(1) - expected) < 0.000001 And Abs(layout.Y2(1) - expected) < 0.000001
        Next unit
    Next placement
    GoTo Restore
Failed:
    Check "label.layout.runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    If IsArray(saved) Then config.Formula = saved
End Sub

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
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
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
    Check prefix & ".physicalTotal", model.RebarCount = count
    Check prefix & ".singleLabel", model.Annotations.Count = 1
    If mixedDiameter Then
        expected = CStr(count - rows - 2) & ChrW$(216) & "20 + " & CStr(rows + 2) & ChrW$(216) & "16"
    Else
        expected = CStr(count) & ChrW$(216) & "20"
    End If
    Check prefix & ".allDiameters", model.Annotations.Text(1) = expected
    Dim anchors(1 To 2) As TSectionAnnotation
    anchors(1).Name = "Top": anchors(1).StartX = -200#: anchors(1).EndX = 200#: anchors(1).StartY = 550#: anchors(1).EndY = 550#: anchors(1).NormalY = 1#
    anchors(2) = anchors(1): anchors(2).StartY = 500#: anchors(2).EndY = 500#
    Set annotations = New CSectionAnnotations: Set labels = New CRebarGroupAnnotationBuilder
    labels.AddRebarLabels annotations, builder.Rebars, anchors, 2
    Check prefix & ".duplicateAnchor", annotations.Count = 1 And annotations.Text(1) = expected
    Set properties = New CSectionPropertiesCalculator: properties.PrepareSnapshot model
    Set table = ThisWorkbook.Names("rngUnitSettings").RefersToRange
    row = HeaderRow(table, "Length", 1)
    Set reader = New CSectionPlotDataReader
    For Each outputUnit In Array("mm", "cm", "m")
        table.Cells(row, 4).Value2 = CStr(outputUnit)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, model, properties, units
        ' После записи меняем OUTPUT: читать геометрию надо в единицах снимка.
        table.Cells(row, 4).Value2 = "cm"
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        reader.LoadGeometryPreviewFromWorkbook ThisWorkbook, settings
        labelIndex = 0: labelCount = 0
        For i = 1 To reader.AnnotationCount
            If reader.AnnotationType(i) = "REBAR_ANNOTATION" Then labelIndex = i: labelCount = labelCount + 1
        Next i
        Check prefix & ".savedTotal." & CStr(outputUnit), labelCount = 1 And reader.AnnotationText(labelIndex) = expected
        Check prefix & ".savedBoundary." & CStr(outputUnit), Abs(reader.AnnotationBoundaryDistance(labelIndex) - 50#) < 0.000001 And reader.LengthUnit = "mm"
    Next outputUnit
    If Len(outputDirectory) > 0 Then
        Check prefix & ".preview", Len(modTestShapeArchitecture.PreparedPreviewSnapshot(Replace$(prefix, ".", "_"), model, outputDirectory)) > 0
    End If
    GoTo Restore
Failed:
    Check prefix & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    If IsArray(saved) Then config.Formula = saved
End Sub

' ДЛЯ ТЕСТОВ: исчезновение дуг и прямых, координаты стержней, схема и штатный расчёт.
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

' ДЛЯ ТЕСТОВ: общие привязки с произвольными именами, без рецепта сечения.
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
    Check "core.onlyFourPhysicalSegments", path.SegmentCount = 4
    Check "core.noPhantomCorner", Not path.HasSegmentID("Corner.A") And Not path.HasSegmentID("Corner.D")
    Check "core.noPhantomEdge", Not path.HasSegmentID("Edge.A") And Not path.HasSegmentID("Edge.C")
    Check "core.collapsedParameters", path.AnchorStation("Edge.A", 0#) = path.AnchorStation("Edge.A", 1#)
    Check "core.consecutiveCollapsed", path.AnchorStation("Corner.A") = path.AnchorStation("Edge.A")
    station = path.AnchorStation("Corner.D")
    Check "core.secondJoint", Abs(station - (100# + GEOM_PI * 500#)) < 0.000001
    For i = 1 To path.SegmentCount
        path.CopySegment i, edge
        If edge.Sweep <> 0# Then
            Check "core.analyticRadius." & CStr(i), Abs(Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2) - 500#) < 0.000001
            Check "core.analyticSweep." & CStr(i), Abs(edge.Sweep + GEOM_PI / 2#) < 0.000000001
        End If
    Next i
    Set copy = path.WithID("Copy")
    Check "core.copyAlias", copy.AnchorStation("Corner.D") = station
    Set copy = path.Reversed
    Check "core.reverseAlias", Abs(copy.AnchorStation("Corner.D") - 1100#) < 0.000001
    Set copy = path.Transformed(1000000#, -1000000#, GEOM_PI / 2#, True)
    copy.AtDistance copy.AnchorStation("Edge.A"), x, y, tx, ty
    Check "core.transformedAlias", Abs(x - 1000550#) < 0.000001 And Abs(y + 1000250#) < 0.000001
    Set copy = path.Slice(0#, 300#)
    Check "core.sliceAlias", copy.AnchorStation("Corner.A") = 0# And copy.SegmentCount = 1
    Set copy = path.BetweenAnchors("Full", "Edge.D", 0#, "Edge.D", 0#, True, True)
    Check "core.wrapAlias", Abs(copy.AnchorStation("Edge.A") - 1100#) < 0.000001 And Abs(copy.Length - path.Length) < 0.000001
    On Error Resume Next
    Set copy = path.BetweenAnchors("Empty", "Edge.A", 0#, "Edge.A", 1#)
    errorNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "core.emptyStillRejected", errorNumber <> 0
    On Error Resume Next
    station = path.AnchorStation("Misspelled")
    errorNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "core.unknownStillRejected", errorNumber <> 0 And path.LastDiagnostic.Code = "MISSING_SEGMENT"
    Exit Sub
Failed:
    Check "core.runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' ДЛЯ ТЕСТОВ: точные снимки обычных Simple/Tapered/смешанных сечений до и после правки.
Public Function RunRoundedOrdinarySnapshots() As String
    Dim config As Object, saved As Variant, fixture As Variant, settings As CSystemSettingsReader
    Dim units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel, result As String
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    For Each fixture In Array("RoundedSimple", "RoundedTapered", "RoundedMixed", "RoundedMixedLeft")
        modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
        result = result & modTestShapeArchitecture.PreparedShapeSnapshot(CStr(fixture), model, builder.Mesh, builder.Rebars)
    Next fixture
    result = result & "TOTAL_ROUNDED_ORDINARY: passed=4; failed=0" & vbCrLf
    GoTo Restore
Failed:
    result = result & "FAIL: ordinary.runtime." & CStr(Err.Number) & "." & Err.Description & vbCrLf
Restore:
    If IsArray(saved) Then config.Formula = saved
    RunRoundedOrdinarySnapshots = result
End Function

' Для Simple независимый эталон — две четверти окружности и вертикаль.
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
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    modTestShapeArchitecture.ConfigureShapeFixture "RoundedSimple"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    Set table = ThisWorkbook.Names("rngRoundedRectangleGeometry").RefersToRange
    table.Cells(3, 2).Value2 = 600# / factor: table.Cells(3, 3).Value2 = height / factor
    table.Cells(7, 2).Value2 = "Simple": table.Cells(7, 3).Value2 = "Simple"
    table.Cells(9, 2).Value2 = leftRadius / factor: table.Cells(9, 3).Value2 = rightRadius / factor
    main = HeaderRow(table, "as", 2): extra = HeaderRow(table, "d2", 2)
    For row = main + 1 To main + 4
        table.Cells(row, 2).Value2 = cover / factor: table.Cells(row, 3).Value2 = 20# / factor
        table.Cells(row, 4).Value2 = IIf(row < main + 3, 5, 0)
        table.Cells(row, 5).Value2 = 0#: table.Cells(row, 6).Value2 = 0#
    Next row
    For row = extra + 1 To extra + 4: table.Cells(row, 2).Value2 = 0#: table.Cells(row, 5).Value2 = 0#: Next row
    Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
    Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
    Check label & ".tenBars", model.RebarCount = 10
    Check label & ".outerArea", Abs(definition.ConcreteDomain.Area - (600# * height - (2# - GEOM_PI / 2#) * (leftRadius ^ 2 + rightRadius ^ 2))) < 0.00001
    For side = 1 To 2
        direction = IIf(side = 1, -1#, 1#): group = IIf(side = 1, "H.Left", "H.Right")
        radius = IIf(side = 1, leftRadius, rightRadius): effectiveRadius = GeomMax(0#, radius - cover)
        top = height / 2# - cover: outsideX = direction * (300# - cover)
        startX = outsideX - direction * effectiveRadius
        vertical = 2# * top - 2# * effectiveRadius: length = vertical + GEOM_PI * effectiveRadius
        Set path = definition.Path(group)
        Check label & ".length." & group, Abs(path.Length - length) < 0.000001
        path.AtDistance 0#, x, y, tx, ty
        Check label & ".start." & group, Abs(x - startX) < 0.000001 And Abs(y - top) < 0.000001
        path.AtDistance path.Length, x, y, tx, ty
        Check label & ".end." & group, Abs(x - startX) < 0.000001 And Abs(y + top) < 0.000001
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
            Check label & ".position." & group & "." & CStr(k), Abs(x - expectedX) < 0.000001 And Abs(y - expectedY) < 0.000001
            count = 0
            For i = 1 To model.RebarCount
                If Abs(model.RebarX(i) - expectedX) < 0.000001 And Abs(model.RebarY(i) - expectedY) < 0.000001 Then count = count + 1
            Next i
            Check label & ".physical." & group & "." & CStr(k), count = 1
            mReport = mReport & "BAR: " & label & "; " & group & "; " & CStr(k) & "; x=" & CStr(expectedX) & "; y=" & CStr(expectedY) & vbCrLf
        Next k
    Next side
    result = modTestSectionRecipes.ReviewModelEquilibrium(label, model, settings, units, passed, failed)
    mPassed = mPassed + passed: mFailed = mFailed + failed: mReport = mReport & result
    If Len(outputDirectory) > 0 Then
        result = modTestShapeArchitecture.PreparedPreviewSnapshot(Replace$(label, ".", "_"), model, outputDirectory)
        Check label & ".preview", Len(result) > 0
    End If
    CheckRoundedWorkbookRun label, units, model
    GoTo Restore
Failed:
    Check label & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    If IsArray(saved) Then config.Formula = saved
End Sub

Public Function RunGeometryReviewTests(Optional ByVal outputDirectory As String = "") As String
    mPassed = 0: mFailed = 0: mReport = ""
    TestPerPath 3, False: TestPerPath 3, True
    TestPerPath 15, False: TestPerPath 15, True
    TestRoundedLimit 1000#, outputDirectory: TestRoundedLimit 600#, outputDirectory
    TestRoundedLimit 600.001, outputDirectory
    TestDisabledSides "Rectangle", "rngRectSetGeometry", Array("H1.as_1", "H1.as_2", "B1.as_1", "B1.as_2")
    TestDisabledSides "RoundedSimple", "rngRoundedRectangleGeometry", Array("H.Left", "H.Right", "B.Top", "B.Bottom")
    TestDisabledSides "HollowCentered", "rngHollowRectangleGeometry", Array("H.Left", "H.Right", "B.Top", "B.Bottom")
    RunGeometryReviewTests = mReport & "TOTAL_GEOMETRY_REVIEW: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

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
        Check label & ".unique." & CStr(i), Not ids.Exists(id)
        ids.Add id, True
        If builder.Rules.PositionRuleID(i) = "Extra" Then
            If builder.Rules.PositionPathID(i) = "P1" Then first = first + 1 Else second = second + 1
        End If
    Next i
    Check label & ".P1", first = (n + 1) \ 2
    Check label & ".P2", second = (n + 1) \ 2
    Check label & ".physical", model.RebarCount = 2 * n + n + 1
    mReport = mReport & "PER_PATH: " & label & "; extra=" & CStr(first) & "+" & CStr(second) & "; physical=" & CStr(model.RebarCount) & vbCrLf
    Exit Sub
Failed:
    Check label & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
End Sub

' Отключает по одной стороне через n=0 и d=0. Подписи остальных сторон
' обязаны совпасть с их фактическими стержнями; размеры сечения сохраняются.
Private Sub TestDisabledSides(ByVal fixture As String, ByVal tableName As String, ByVal groups As Variant)
    Dim config As Object, saved As Variant, table As Object, main As Long, extra As Long, side As Long, mode As Long, row As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim labels As Object, counts As Object, diameters As Object, group As Variant, i As Long, key As String, factor As Double
    Dim baseDimensions As Long, dimensions As Long, label As String
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    modTestShapeArchitecture.ConfigureShapeFixture "Circle": modTestShapeArchitecture.ConfigureShapeFixture fixture
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    Set table = ThisWorkbook.Names(tableName).RefersToRange
    main = HeaderRow(table, "as", 2): extra = HeaderRow(table, "d2", 2)
    For row = main + 1 To main + 4
        table.Cells(row, 2).Value2 = 50# / factor: table.Cells(row, 3).Value2 = 20# / factor
        table.Cells(row, 4).Value2 = 4: table.Cells(row, 5).Value2 = 0#: table.Cells(row, 6).Value2 = 0#
    Next row
    For row = extra + 1 To extra + 4: table.Cells(row, 2).Value2 = 0#: table.Cells(row, 5).Value2 = 0#: Next row
    Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
    For i = 1 To model.AnnotationCount
        If model.Annotations.AnnotationType(i) = "DIMENSION" Then baseDimensions = baseDimensions + 1
    Next i
    For side = 1 To 4
        For mode = 3 To 4
            label = "disabledSide." & fixture & "." & CStr(groups(side - 1)) & ".column" & CStr(mode)
            table.Cells(main + side, mode).Value2 = 0
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
                If model.Annotations.AnnotationType(i) = "REBAR_ANNOTATION" Then
                    key = Mid$(model.Annotations.AnnotationID(i), 7): labels.Add key, model.Annotations.Text(i)
                    Check label & ".noEmptyLabel." & key, counts.Exists(key)
                End If
                If model.Annotations.AnnotationType(i) = "DIMENSION" Then dimensions = dimensions + 1
            Next i
            Check label & ".off", Not labels.Exists(CStr(groups(side - 1))) And Not counts.Exists(CStr(groups(side - 1)))
            Check label & ".dimensions", dimensions = baseDimensions
            For Each group In counts.Keys
                Check label & ".otherLabel." & CStr(group), labels.Exists(CStr(group))
                If labels.Exists(CStr(group)) Then Check label & ".quantity." & CStr(group), _
                    CStr(labels(group)) = CStr(counts(group)) & ChrW$(216) & CStr(diameters(group))
            Next group
            table.Cells(main + side, 3).Value2 = 20# / factor: table.Cells(main + side, 4).Value2 = 4
        Next mode
    Next side
    GoTo Restore
Failed:
    Check label & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    If IsArray(saved) Then config.Formula = saved
End Sub

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
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    modTestShapeArchitecture.ConfigureShapeFixture "RoundedSimple"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    Set table = ThisWorkbook.Names("rngRoundedRectangleGeometry").RefersToRange
    table.Cells(3, 2).Value2 = 600# / factor: table.Cells(3, 3).Value2 = height / factor
    table.Cells(7, 2).Value2 = "Simple": table.Cells(7, 3).Value2 = "Simple"
    table.Cells(9, 2).Value2 = 300# / factor: table.Cells(9, 3).Value2 = 300# / factor
    header = HeaderRow(table, "as", 2): extra = HeaderRow(table, "d2", 2)
    For row = header + 1 To header + 4
        table.Cells(row, 2).Value2 = 50# / factor: table.Cells(row, 3).Value2 = 20# / factor
        ' В таблице сначала H.Left/H.Right, затем B.Top/B.Bottom.
        table.Cells(row, 4).Value2 = IIf(row < header + 3, 5, 0)
        table.Cells(row, 5).Value2 = 0#: table.Cells(row, 6).Value2 = 0#
    Next row
    For row = extra + 1 To extra + 4: table.Cells(row, 2).Value2 = 0#: table.Cells(row, 5).Value2 = 0#: Next row
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
    expected = height - 600# + GEOM_PI * 250#
    For Each paths In Array("H.Left", "H.Right")
        Set path = definition.Path(CStr(paths))
        Check label & ".length." & CStr(paths), Abs(path.Length - expected) < 0.000001
    Next paths
    Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
    Check label & ".tenBars", model.RebarCount = 10
    Check label & ".contour", Abs(definition.ConcreteDomain.Area - (600# * (height - 600#) + GEOM_PI * 90000#)) < 0.00001
    For i = 1 To model.AnnotationCount
        Check label & ".noDisabledLabel." & CStr(i), model.Annotations.AnnotationID(i) <> "REBAR_B.Top" And model.Annotations.AnnotationID(i) <> "REBAR_B.Bottom"
    Next i
    result = modTestSectionRecipes.ReviewModelEquilibrium(label, model, settings, units, passed, failed)
    mPassed = mPassed + passed: mFailed = mFailed + failed: mReport = mReport & result
    If Len(outputDirectory) > 0 Then
        result = modTestShapeArchitecture.PreparedPreviewSnapshot(Replace$(label, ".", "_"), model, outputDirectory)
        Check label & ".preview", Len(result) > 0
    End If
    CheckRoundedWorkbookRun label, units, model
    GoTo Restore
Failed:
    Check label & ".runtime." & CStr(Err.Number) & "." & Err.Description, False
Restore:
    If IsArray(saved) Then config.Formula = saved
End Sub

' Выполняет штатный расчёт одного сочетания и проверяет запись НДС при заданной
' нагрузке. Профиль не включает поиск несущей способности и расчёт трещин.
Private Sub CheckRoundedWorkbookRun(ByVal label As String, ByVal units As CUnitSystem, ByVal model As CSectionModel)
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
    message = RunSectionCalculationForWorkbook(ThisWorkbook, False)
    Check label & ".workbook.run", InStr(1, message, "ошиб", vbTextCompare) = 0 And InStr(1, message, "InputErr", vbTextCompare) = 0
    Check label & ".workbook.state", SnapshotAnchoredExtent(ThisWorkbook.Names("rngNDMElementResults").RefersToRange, False) > 1
    mReport = mReport & "WORKBOOK_RUN: " & label & "; " & message & vbCrLf
End Sub

' Находит строку заголовка тестовой таблицы; перенос таблицы не должен ломать проверку.
Private Function HeaderRow(ByVal table As Object, ByVal text As String, ByVal column As Long) As Long
    Dim row As Long
    For row = 1 To table.Rows.Count
        If CStr(table.Cells(row, column).Value2) = text Then HeaderRow = row: Exit Function
    Next row
    Err.Raise 5, "modTestGeometryReview", "Не найден заголовок " & text
End Function

Private Sub Check(ByVal name As String, ByVal success As Boolean)
    If success Then mPassed = mPassed + 1: mReport = mReport & "OK: " & name & vbCrLf Else mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
End Sub
