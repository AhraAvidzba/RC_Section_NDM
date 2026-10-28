Attribute VB_Name = "modTestElementGeometryViews"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: представление и обмен фактическими границами бетонных ячеек
' ==========================================================================
' Проверяет рабочий CAD-экспортер на подставном ModelSpace, строгий адаптер
' импорта и настоящий Freeform Excel. Растровые доказательства отверстий
' сохраняются отдельно: проверки узлов сами по себе не доказывают заливку.

Private mPassed As Long, mFailed As Long ' ДЛЯ ТЕСТОВ: итог независимых проверок представления.
Private mReport As String ' ДЛЯ ТЕСТОВ: первичный журнал с путями растровых доказательств.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Запускает обмен аналитическими дугами/отверстиями и создание фигур Excel;
' пользовательские книги и AutoCAD не закрывает, расчет НДС не запускает.
Public Function RunElementGeometryViewTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    TestCurvedElementExport
    TestElementExportFailureCleanup
    TestStrictContourImport
    TestNativeFreeformViews
    TestPlotReaderGeometryOwnership
    GoTo Finish
Failed:
    TestCheck "elementViews.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finish:
    mReport = mReport & "TOTAL_ELEMENT_GEOMETRY_VIEWS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
    RunElementGeometryViewTests = mReport
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет дуги и несвязные outer без отверстий на подставном ModelSpace.
' Метод COM Boolean имеет зарезервированное имя VBA и не имитируется классом;
' площадь, отверстия и очистка после реального Boolean проверяются в native CAD.
Private Sub TestCurvedElementExport()
    Dim query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    Dim disk As CConcreteRegion: Set disk = query.CircleRegion(0#, 0#, 100#)
    Dim ms As CTestFakeAcadContourExport: Set ms = New CTestFakeAcadContourExport
    TestCheck "elementViews.export.disk.oneRegion", TestExportConcreteElementGeometry(ms, disk) = 1
    Dim entity As CTestFakeAcadContourExport, area As Double, arcCount As Long, i As Long, normal As Variant
    For Each entity In ms.TestCreatedEntities
        If entity.ObjectName = "AcDbRegion" And Not entity.TestDeleted Then area = area + entity.Area
        If entity.ObjectName = "AcDbPolyline" Then
            TestCheck "elementViews.export.sourceDeleted", entity.TestDeleted
            normal = entity.Normal
            TestCheck "elementViews.export.WCS", CDbl(normal(0)) = 0# And CDbl(normal(1)) = 0# And CDbl(normal(2)) = 1# And entity.Elevation = 0#
            For i = 0 To (UBound(entity.Coordinates) + 1) \ 2 - 1
                If Abs(entity.TestBulgeAt(i)) > 0# Then arcCount = arcCount + 1
            Next i
        End If
    Next entity
    TestNear "elementViews.export.disk.area", area, GEOM_PI * 10000#, 0.00000001
    TestCheck "elementViews.export.analyticArcs", arcCount = 4
    TestCheck "elementViews.export.noSourceScan", ms.TestForbiddenReads = 0
    Dim parts As Collection: Set parts = New Collection
    parts.Add query.CircleRegion(-150#, 0#, 25#): parts.Add query.CircleRegion(150#, 0#, 25#)
    Dim disconnected As CConcreteRegion: Set disconnected = query.MergeDisjointRegions(parts)
    Set ms = New CTestFakeAcadContourExport
    TestCheck "elementViews.export.disconnected.twoRegions", TestExportConcreteElementGeometry(ms, disconnected) = 2
    area = 0#
    For Each entity In ms.TestCreatedEntities
        If entity.ObjectName = "AcDbRegion" And Not entity.TestDeleted Then area = area + entity.Area
    Next entity
    TestNear "elementViews.export.disconnected.area", area, GEOM_PI * 1250#, 0.00000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Имитирует отказ второго несвязного outer и проверяет только собственную очистку.
' Отказ после вычитания отверстия проверяется реальным AutoCAD, без суррогата Boolean.
Private Sub TestElementExportFailureCleanup()
    Dim query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    Dim parts As Collection: Set parts = New Collection
    parts.Add query.CircleRegion(-150#, 0#, 25#): parts.Add query.CircleRegion(150#, 0#, 25#)
    Dim disconnected As CConcreteRegion: Set disconnected = query.MergeDisjointRegions(parts)
    Dim ms As CTestFakeAcadContourExport: Set ms = New CTestFakeAcadContourExport
    ms.TestFailRegionAt = 2
    Dim number As Long, ignored As Long
    On Error Resume Next
    ignored = TestExportConcreteElementGeometry(ms, disconnected): number = Err.Number
    On Error GoTo 0
    TestCheck "elementViews.export.failurePreserved", number = vbObjectError + 9103
    Dim entity As CTestFakeAcadContourExport
    For Each entity In ms.TestCreatedEntities
        TestCheck "elementViews.export.failureOwnedCleanup", entity.TestDeleted
    Next entity
    TestCheck "elementViews.export.failureNoScan", ms.TestForbiddenReads = 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Contour читает реальную границу и отклоняет spline без подмены;
' Adaptive сохраняет эквивалентное представление с переданными A/I.
Private Sub TestStrictContourImport()
    Dim fixture As Object: Set fixture = Application.Workbooks.Add(-4167)
    Dim number As Long, description As String
    On Error GoTo Failed
    fixture.Worksheets(1).Cells(1, 1).Value2 = "Key": fixture.Worksheets(1).Cells(1, 2).Value2 = "Value"
    fixture.Worksheets(1).Cells(1, 3).Value2 = "Unit"
    fixture.Worksheets(1).Cells(2, 1).Value2 = "Mesh.CellGeometry": fixture.Worksheets(1).Cells(2, 2).Value2 = "Contour"
    fixture.Worksheets(1).Cells(2, 3).Value2 = "-"
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader
    settings.LoadFromRange fixture.Worksheets(1).Range("A1:C2")
    Dim query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    Dim geometry As CConcreteRegion: Set geometry = query.CircleRegion(40#, -20#, 100#)
    Dim properties As TGeometryProperties: query.RegionProperties geometry, properties
    Dim concrete As CTestFakeAcadRegion: Set concrete = New CTestFakeAcadRegion
    concrete.TestInitialize properties.Area, properties.CentroidX, properties.CentroidY, properties.Ixc, properties.Iyc, properties.Ixyc, "C", "C1"
    concrete.TestSetGeometry geometry
    Dim rebar As CTestFakeAcadRegion: Set rebar = New CTestFakeAcadRegion
    rebar.TestInitialize 100#, 40#, -20#, 1000#, 1000#, 0#, "R", "R1"
    Dim entities As Collection: Set entities = New Collection
    entities.Add concrete: entities.Add rebar
    Dim importer As CAutoCADSectionModelImporter: Set importer = New CAutoCADSectionModelImporter
    Dim model As CSectionModel
    Set model = importer.ImportFromModelSpace(entities, "C", "R", "A400", 0.000001, meshSettings:=settings)
    TestCheck "elementViews.import.exactBoundary", model.ConcreteHasBoundary(1)
    TestNear "elementViews.import.area", model.ConcreteGeometry(1).Area, geometry.Area, 0.00000001
    TestCheck "elementViews.import.sourcePreserved", Not concrete.TestDeleted And Not rebar.TestDeleted
    concrete.TestSetUnsupportedCurve "AcDbSpline"
    On Error Resume Next
    Set model = importer.ImportFromModelSpace(entities, "C", "R", "A400", 0.000001, meshSettings:=settings)
    number = Err.Number: description = Err.Description
    On Error GoTo Failed
    TestCheck "elementViews.import.strictNoEquivalent", number = vbObjectError + 4410 And InStr(description, "AcDbSpline") > 0 And InStr(description, "C1") > 0
    Set model = importer.ImportFromModelSpace(entities, "C", "R", "A400", 0.000001)
    TestCheck "elementViews.import.legacyEquivalentRetained", Not model.ConcreteHasBoundary(1) And model.ConcreteCount = 1
    fixture.Close False
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    On Error Resume Next: fixture.Close False: On Error GoTo 0
    Err.Raise number, "TestStrictContourImport", description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создает настоящие Freeform в собственной временной книге и сохраняет PNG
' для независимой проверки заливки отверстия, дуг и несвязных областей.
Private Sub TestNativeFreeformViews()
    Dim fixture As Object: Set fixture = Application.Workbooks.Add(-4167)
    Dim number As Long, description As String
    On Error GoTo Failed
    Dim query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    TestOneNativeFreeform fixture, query.CircleRegion(0#, 0#, 100#), 0#, 0#, "Circle"
    Dim ring As CConcreteRegion: Set ring = query.SubtractContainedRegion(query.CircleRegion(0#, 0#, 100#), query.CircleRegion(0#, 0#, 30#))
    TestOneNativeFreeform fixture, ring, 0#, 0#, "Hole"
    Dim parts As Collection: Set parts = New Collection
    parts.Add query.CircleRegion(-150#, 0#, 25#): parts.Add query.CircleRegion(150#, 0#, 25#)
    TestOneNativeFreeform fixture, query.MergeDisjointRegions(parts), 0#, 0#, "Disconnected"
    Dim coordinates(1 To 6, 1 To 2) As Double
    coordinates(1, 1) = 0#: coordinates(1, 2) = 0#: coordinates(2, 1) = 100#: coordinates(2, 2) = 0#
    coordinates(3, 1) = 100#: coordinates(3, 2) = 30#: coordinates(4, 1) = 30#: coordinates(4, 2) = 30#
    coordinates(5, 1) = 30#: coordinates(5, 2) = 100#: coordinates(6, 1) = 0#: coordinates(6, 2) = 100#
    TestOneNativeFreeform fixture, query.PolygonRegion(coordinates), 40#, 40#, "Concave"
    fixture.Close False
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    On Error Resume Next: fixture.Close False: On Error GoTo 0
    Err.Raise number, "TestNativeFreeformViews", description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет тип и масштаб фигуры, сохраняет первичный растр настоящего Excel.
Private Sub TestOneNativeFreeform(ByVal fixture As Object, ByVal region As CConcreteRegion, _
        ByVal centerX As Double, ByVal centerY As Double, ByVal label As String)
    Dim chartObject As Object: Set chartObject = fixture.Worksheets(1).ChartObjects.Add(0#, 0#, 220#, 220#)
    Dim chart As Object: Set chart = chartObject.Chart
    chart.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Dim plotter As CSectionPlotter: Set plotter = New CSectionPlotter
    Dim shape As Object: Set shape = plotter.TestDrawConcreteRegionShape(chart, region, centerX, centerY)
    TestCheck "elementViews.plot." & label & ".oneShape", chart.Shapes.Count = 1
    TestCheck "elementViews.plot." & label & ".freeform", shape.Type = 5
    TestNear "elementViews.plot." & label & ".width07", shape.Width, 140#, 0.01
    TestNear "elementViews.plot." & label & ".height07", shape.Height, 140#, 0.01
    TestCheck "elementViews.plot." & label & ".lineEnabled", shape.Line.Visible <> 0
    Dim transparency As Double: If region.LoopCount > 1 Then transparency = 1#
    TestNear "elementViews.plot." & label & ".lineTransparency", shape.Line.Transparency, transparency, 0.00001
    TestCheck "elementViews.plot." & label & ".lineMatchesFill", shape.Line.ForeColor.RGB = shape.Fill.ForeColor.RGB
    Dim lineWeight As Double: lineWeight = 0.15
    If region.LoopCount > 1 Then lineWeight = 0.01
    TestNear "elementViews.plot." & label & ".lineWeight", shape.Line.Weight, lineWeight, 0.0001
    Dim path As String: path = ThisWorkbook.Path & "\TestElementGeometry_" & label & ".png"
    TestCheck "elementViews.plot." & label & ".PNG", chart.Export(path, "PNG")
    mReport = mReport & "RASTER_ELEMENT_GEOMETRY: " & label & "; " & path & vbCrLf
    chartObject.Delete
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает точную модель и проверяет независимость областей reader-а.
' Изменяет всю полученную геометрию, включая дуги, отверстие и несвязные части;
' последующее чтение и физические A/c/I должны оставаться согласованными.
Private Sub TestPlotReaderGeometryOwnership()
    Dim query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Dim units As CUnitSystem: Set units = New CUnitSystem: units.LoadFromSettings settings
    Dim regions As Collection: Set regions = New Collection
    regions.Add query.SubtractContainedRegion(query.CircleRegion(0#, 0#, 100#), query.CircleRegion(0#, 0#, 30#))
    Dim parts As Collection: Set parts = New Collection
    parts.Add query.CircleRegion(300#, 0#, 25#): parts.Add query.CircleRegion(380#, 0#, 20#)
    regions.Add query.MergeDisjointRegions(parts)
    Dim compact As CConcreteRegion: Set compact = query.RectangleRegion(600#, 0#, 80#, 40#, 0.15)
    Dim model As CSectionModel: Set model = New CSectionModel: model.SourceType = "Generated"
    Dim i As Long, props As TGeometryProperties, region As CConcreteRegion
    For i = 1 To regions.Count
        Set region = regions(i): query.RegionProperties region, props
        model.AddConcreteElement props.CentroidX, props.CentroidY, props.Area, 1, shapeType:="Contour", _
            localIx:=props.Ixc, localIy:=props.Iyc, localIxy:=props.Ixyc, rotationKnown:=True, exactRegion:=region
    Next i
    model.AddConcreteElement 600#, 0#, 3200#, 1, shapeType:="Rectangle", width:=80#, height:=40#, rotation:=0.15, rotationKnown:=True
    Dim writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
    writer.WriteGeometryPreview ThisWorkbook, model, PrepareSectionSnapshot(model), units, False
    Dim reader As CSectionPlotDataReader: Set reader = New CSectionPlotDataReader
    reader.LoadGeometryPreviewFromWorkbook ThisWorkbook, settings
    TestCheck "elementViews.reader.count", reader.Count = 3
    Dim returned As CConcreteRegion, again As CConcreteRegion, replacement As CConcreteRegion
    Dim edges() As TRegionEdge, count As Long, opening(1 To 1) As Boolean
    Set replacement = query.RectangleRegion(-1000#, -1000#, 2#, 3#)
    replacement.CopyEdges edges, count
    For i = 1 To reader.Count
        Set returned = reader.ConcreteGeometry(i): Set again = reader.ConcreteGeometry(i)
        TestCheck "elementViews.reader.independent.C" & CStr(i), Not returned Is again
        returned.SetGeometry edges, count, 1, replacement.Area, "TestMutation", 0, opening
        returned.SetOwner "TestMutation", "TestMutation", "TestMutation"
        Set again = reader.ConcreteGeometry(i)
        If i <= regions.Count Then
            Set region = regions(i)
            TestCheck "elementViews.reader.boundary.C" & CStr(i), reader.ConcreteHasBoundary(i)
            TestCheck "elementViews.reader.topology.C" & CStr(i), query.RegionsCoincide(region, again) And region.LoopCount = again.LoopCount
        Else
            Set region = compact
            TestCheck "elementViews.reader.compact", Not reader.ConcreteHasBoundary(i) And Not again Is Nothing
            TestCheck "elementViews.reader.compactGeometry", query.RegionsCoincide(compact, again)
        End If
        query.RegionProperties again, props
        TestNear "elementViews.reader.area.C" & CStr(i), props.Area, model.ConcreteArea(i), 0.00000001
        TestNear "elementViews.reader.cx.C" & CStr(i), props.CentroidX, model.ConcreteX(i), 0.00000001
        TestNear "elementViews.reader.cy.C" & CStr(i), props.CentroidY, model.ConcreteY(i), 0.00000001
        TestNear "elementViews.reader.Ix.C" & CStr(i), props.Ixc, model.ConcreteLocalIx(i), 0.001
        TestNear "elementViews.reader.Iy.C" & CStr(i), props.Iyc, model.ConcreteLocalIy(i), 0.001
        TestNear "elementViews.reader.Ixy.C" & CStr(i), props.Ixyc, model.ConcreteLocalIxy(i), 0.001
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает результат одного независимого условия без остановки остальных проверок.
Private Sub TestCheck(ByVal label As String, ByVal passed As Boolean)
    If passed Then mPassed = mPassed + 1 Else mFailed = mFailed + 1
    mReport = mReport & IIf(passed, "PASS ", "FAIL ") & label & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает фактическое значение с аналитическим ожиданием в заданном абсолютном допуске.
Private Sub TestNear(ByVal label As String, ByVal actual As Double, ByVal expected As Double, ByVal tolerance As Double)
    TestCheck label & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) <= tolerance
End Sub
