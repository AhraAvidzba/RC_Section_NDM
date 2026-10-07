Attribute VB_Name = "modTestAutoCADContours"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: достоверные контуры Region и Polyline
' ==========================================================================
' Адресные проверки фактических ребер, отверстий, ориентации, строгих ошибок
' и очистки временных COM-объектов. Проверяют production reader/importer и
' общий geometry-query без solver. Настоящий AutoCAD выполняется отдельным
' entrypoint-ом; fixture-прогон не объявляется живой CAD-интеграцией.

Private mPassed As Long
Private mFailed As Long
Private mReport As String

' ДЛЯ ТЕСТОВ: запускает parity/negative/import проверки без AutoCAD.
Public Function RunAutoCADContourTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    CheckEquivalentSources
    CheckNestedRegion
    CheckInvalidContours
    CheckImportedSnapshot
    CheckMultipleOpenings
    CheckFiveOpenings
    CheckOptionalLayersAndSeparateConcrete
    CheckEmptyContourLayers
    CheckMixedContourSources
    CheckMixedContourTopology
    CheckResultsRoundTrip
    CheckAdditiveContourExport
    CheckImportedContourDimensions
    CheckOpeningWithoutOuter
    GoTo Finish
Failed:
    Check "contour.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finish:
    mReport = mReport & "TOTAL_AUTOCAD_CONTOURS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
    RunAutoCADContourTests = mReport
End Function

' ДЛЯ ТЕСТОВ: отдельная книга проверяет единственный формат контуров и
' save/reopen без импорта, solve и изменений пользовательского Results.
Public Function RunPostAudit03ContourSnapshotTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    Dim fixture As Object, sheet As Object, writer As CNDMResultsWriter, model As CSectionModel, restored As CSectionModel
    Dim contours As CSectionContours, readBack As CSectionContours, query As CSectionGeometryQuery
    Dim props As CSectionPropertiesCalculator, anchor As Object, unitText As Variant, marker As Object
    Dim solves As Long, path As String, value As Variant, number As Long, description As String, i As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, unitSheet As Object
    On Error GoTo Failed
    solves = SectionEquilibriumSolveCount()
    Set fixture = Application.Workbooks.Add(-4167): Set sheet = fixture.Worksheets(1): sheet.Name = "Results"
    CreatePostAudit03SnapshotAnchors fixture, sheet
    Set unitSheet = fixture.Worksheets.Add
    unitSheet.Cells(1, 1).Value2 = "Key": unitSheet.Cells(1, 2).Value2 = "Value"
    unitSheet.Cells(1, 3).Value2 = "Unit": unitSheet.Cells(2, 3).Value2 = "-"
    unitSheet.Cells(2, 1).Value2 = "Units.Length.Output"
    Set writer = New CNDMResultsWriter: Set model = New CSectionModel
    model.AddConcreteElement 0#, 0#, 10000#, 1, , , "Rectangle", 100#, 100#, 0#, , 123456#, 654321#, 12#
    model.AddRebarElement 5#, 5#, 12#, 0#, "A400"
    Set contours = model.Contours
    contours.AddContourArc "OUT_BIG", 100#, 0#, 0#, -100#, 0#, 0#, 100#, 1.5 * GEOM_PI, "Большая дуга", "O1", "Outer", "TEST"
    contours.AddContourArc "OUT_CLOSE", 0#, -100#, 100#, 0#, 0#, 0#, 100#, 0.5 * GEOM_PI, "Замыкание", "O1", "Outer", "TEST"
    contours.AddContourArc "HOLE_BIG", 10#, 0#, 0#, 10#, 0#, 0#, 10#, -1.5 * GEOM_PI, "Отрицательная дуга", "H1", "Opening", "TEST"
    contours.AddContourArc "HOLE_CLOSE", 0#, 10#, 10#, 0#, 0#, 0#, 10#, -0.5 * GEOM_PI, "Замыкание отверстия", "H1", "Opening", "TEST"
    contours.AddContourCircle "OUT_CIRCLE", 250#, 0#, 50#, "Окружность", "O2", "Outer", "TEST"
    contours.AddContourCircle "HOLE_CIRCLE", 250#, 0#, 20#, "Отдельное отверстие", "H2", "Opening", "TEST"
    contours.AddContourLine "LINE1", 400#, 0#, 500#, 0#, "Прямая 1", "O3", "Outer", "TEST"
    contours.AddContourLine "LINE2", 500#, 0#, 500#, 100#, "Прямая 2", "O3", "Outer", "TEST"
    contours.AddContourLine "LINE3", 500#, 100#, 400#, 100#, "Прямая 3", "O3", "Outer", "TEST"
    contours.AddContourLine "LINE4", 400#, 100#, 400#, 0#, "Прямая 4", "O3", "Outer", "TEST"
    Set props = PrepareSectionSnapshot(model)
    writer.WriteGeometryPreview fixture, model, props
    Set marker = sheet.Cells(22, 46): marker.Value2 = "USER-GAP"
    Check "postAudit03.contours.geometry15", sheet.Cells(20, 27).Value2 = vbNullString
    Check "postAudit03.contours.geometryOnlyElements", sheet.Cells(23, 14).Value2 = vbNullString
    For Each unitText In Array("mm", "cm", "m")
        unitSheet.Cells(2, 2).Value2 = CStr(unitText)
        Set settings = New CSystemSettingsReader: settings.LoadFromRange unitSheet.Range("A1:C2")
        Set units = New CUnitSystem: units.LoadFromSettings settings
        writer.WriteSectionContours fixture, contours, sheet.Cells(21, 12).Value2, units
        Set readBack = ReadSavedSectionContours(fixture)
        CheckPostAudit03ContoursEqual "postAudit03.contours." & CStr(unitText), contours, readBack
        Set restored = ReadSectionGeometryFromResults(fixture)
        CheckArea "postAudit03.contours.mechanicalA." & CStr(unitText), restored.ConcreteArea(1), 10000#
        CheckArea "postAudit03.contours.mechanicalIx." & CStr(unitText), restored.ConcreteLocalIx(1), 123456#
        Set query = New CSectionGeometryQuery: query.Initialize restored
        CheckArea "postAudit03.contours.domain." & CStr(unitText), query.ConcreteDomain.Area, 12000# * GEOM_PI + 10000#
        Check "postAudit03.contours.marker." & CStr(unitText), marker.Value2 = "USER-GAP"
    Next unitText
    path = ThisWorkbook.Path & "\PostAudit03_Contour_RoundTrip.xlsm"
    Application.DisplayAlerts = False: fixture.SaveAs path, 52: fixture.Close False
    Set fixture = Application.Workbooks.Open(path): Set sheet = fixture.Worksheets("Results")
    Set readBack = ReadSavedSectionContours(fixture)
    CheckPostAudit03ContoursEqual "postAudit03.contours.reopen", contours, readBack
    Set anchor = fixture.Names.Item("rngNDMSectionContours").RefersToRange
    anchor.Offset(-1, 0).Resize(contours.Count + 2, 17).Cut sheet.Cells(19, 200)
    fixture.Names.Item("rngNDMSectionContours").RefersTo = "=Results!$GR$20"
    Set anchor = fixture.Names.Item("rngNDMSectionContours").RefersToRange
    Set readBack = ReadSavedSectionContours(fixture)
    CheckPostAudit03ContoursEqual "postAudit03.contours.moved", contours, readBack
    Check "postAudit03.contours.canonicalHeader", anchor.Value2 = "RunID"
    anchor.Value2 = "RunID v1"
    On Error Resume Next
    Set readBack = ReadSavedSectionContours(fixture): number = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "postAudit03.contours.obsoleteHeaderRejected", number = vbObjectError + 4376
    anchor.Value2 = "RunID"
    value = anchor.Offset(1, 6).Value2: anchor.Offset(1, 6).Value2 = "bad"
    On Error Resume Next
    Set readBack = ReadSavedSectionContours(fixture): number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo Failed
    anchor.Offset(1, 6).Value2 = value
    Check "postAudit03.contours.dynamicErrorAddress", number <> 0 And InStr(description, anchor.Offset(1, 6).Address(False, False)) > 0
    value = anchor.Value2: anchor.ClearContents
    On Error Resume Next
    Set readBack = ReadSavedSectionContours(fixture): number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo Failed
    anchor.Value2 = value
    Check "postAudit03.contours.blankHeaderRejected", number = vbObjectError + 4376
    fixture.Names.Item("rngNDMSectionContours").RefersTo = "=#REF!"
    On Error Resume Next
    Set readBack = ReadSavedSectionContours(fixture): number = Err.Number: Err.Clear
    On Error GoTo Failed
    fixture.Names.Item("rngNDMSectionContours").RefersTo = "=Results!$GR$20"
    Check "postAudit03.contours.brokenNameRejected", number = vbObjectError + 4376 And _
        fixture.Names.Item("rngNDMSectionProperties").RefersToRange.Column = 48
    Set readBack = New CSectionContours
    readBack.AddContourCircle "ONLY_HOLE", 5#, 5#, 2#, "Только отверстие", "ONLY_HOLE", "Opening"
    writer.WriteSectionContours fixture, readBack, sheet.Cells(21, 12).Value2
    Set readBack = ReadSavedSectionContours(fixture)
    Check "postAudit03.contours.onlyOpening", readBack.Count = 1 And readBack.LoopRole(1) = "Opening"
    Check "postAudit03.contours.largeSmallCleared", Len(CStr(anchor.Offset(2, 0).Value2)) = 0 And Len(CStr(anchor.Offset(10, 15).Value2)) = 0
    readBack.Clear: writer.WriteSectionContours fixture, readBack, "EMPTY"
    Set readBack = ReadSavedSectionContours(fixture)
    Check "postAudit03.contours.emptyAvailable", readBack.Count = 0 And anchor.Value2 = "RunID"
    fixture.Names.Item("rngNDMSectionContours").Delete
    On Error Resume Next
    Set readBack = ReadSavedSectionContours(fixture): number = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "postAudit03.contours.missingTableRejected", number = vbObjectError + 4376
    Check "postAudit03.contours.missingTableDoesNotInsert", fixture.Names.Count = 5 And _
        fixture.Names.Item("rngNDMSectionProperties").RefersToRange.Column = 48
    sheet.Cells(20, 27).Value2 = "EndX, mm"
    On Error Resume Next
    Set restored = ReadSectionGeometryFromResults(fixture): number = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "postAudit03.contours.expandedGeometryRejected", number = vbObjectError + 4355
    fixture.Close False: Set fixture = Nothing
    Check "postAudit03.contours.noSolve", SectionEquilibriumSolveCount() = solves
    GoTo Finished
Failed:
    Check "postAudit03.contours.runtime." & CStr(Err.Number) & "." & Err.Description, False
Finished:
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
    RunPostAudit03ContourSnapshotTests = mReport & "TOTAL_POSTAUDIT03_CONTOURS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed)
End Function

' Создает только тестовые якоря нижнего ряда; ширины не переназначаются.
Private Sub CreatePostAudit03SnapshotAnchors(ByVal workbook As Object, ByVal sheet As Object)
    Dim names As Variant, columns As Variant, i As Long
    names = Array("rngNDMElementResults", "rngNDMSectionGeometry", "rngNDMSectionProperties", "rngNDMMaterialDiagrams", "rngNDMSectionAnnotations")
    columns = Array(1, 12, 48, 56, 69)
    For i = 0 To UBound(names)
        workbook.Names.Add CStr(names(i)), "='" & sheet.Name & "'!" & sheet.Cells(20, CLng(columns(i))).Address
    Next i
    workbook.Names.Add "rngNDMSectionContours", "='" & sheet.Name & "'!$AC$20"
End Sub

' Поля проверяются по собственным ID и аналитической геометрии, не по площади одной фигуры.
Private Sub CheckPostAudit03ContoursEqual(ByVal prefix As String, ByVal expected As CSectionContours, ByVal actual As CSectionContours)
    Check prefix & ".count", expected.Count = actual.Count
    If expected.Count <> actual.Count Then Exit Sub
    Dim i As Long
    For i = 1 To expected.Count
        Check prefix & ".id." & CStr(i), expected.ContourID(i) = actual.ContourID(i) And expected.LoopID(i) = actual.LoopID(i)
        Check prefix & ".context." & CStr(i), expected.LoopRole(i) = actual.LoopRole(i) And expected.SourceID(i) = actual.SourceID(i) And expected.Comment(i) = actual.Comment(i)
        Check prefix & ".type." & CStr(i), expected.SegmentType(i) = actual.SegmentType(i)
        Check prefix & ".geometry." & CStr(i), Abs(expected.StartX(i) - actual.StartX(i)) < 0.00000001 And Abs(expected.StartY(i) - actual.StartY(i)) < 0.00000001 And _
            Abs(expected.EndX(i) - actual.EndX(i)) < 0.00000001 And Abs(expected.EndY(i) - actual.EndY(i)) < 0.00000001 And _
            Abs(expected.Radius(i) - actual.Radius(i)) < 0.00000001 And Abs(expected.SweepAngle(i) - actual.SweepAngle(i)) < 0.000000000001
    Next i
End Sub

' ДЛЯ ТЕСТОВ: отдельный opening допустим при уже пустой или приближенной
' сеточной границе. Точные наружные контуры здесь намеренно отсутствуют.
Private Sub CheckOpeningWithoutOuter()
    Dim space As Collection, layers As Collection, importer As CAutoCADSectionModelImporter
    Dim cell As CFakeAcadRegion, layer As CFakeAcadContour, opening As CFakeAcadContour
    Dim section As CSectionModel, query As CSectionGeometryQuery, region As CConcreteRegion
    Dim mode As Long, expected As Double, writer As CNDMResultsWriter, restored As CSectionModel
    Set importer = New CAutoCADSectionModelImporter: Set query = New CSectionGeometryQuery
    For mode = 0 To 3
        Set space = New Collection: Set layers = New Collection
        Set cell = New CFakeAcadRegion
        cell.Initialize 2000#, 50#, 10#, 66666.6666666667, 1666666.66666667, 0#, "CONCRETE", "BOTTOM", 0#, True: space.Add cell
        Set cell = New CFakeAcadRegion
        cell.Initialize 2000#, 50#, 90#, 66666.6666666667, 1666666.66666667, 0#, "CONCRETE", "TOP", 0#, True: space.Add cell
        Set cell = New CFakeAcadRegion
        cell.Initialize 1200#, 10#, 50#, 360000#, 40000#, 0#, "CONCRETE", "LEFT", 0#, True: space.Add cell
        Set cell = New CFakeAcadRegion
        cell.Initialize 1200#, 90#, 50#, 360000#, 40000#, 0#, "CONCRETE", "RIGHT", 0#, True: space.Add cell
        Set cell = New CFakeAcadRegion
        cell.Initialize GEOM_PI * 4#, 10#, 10#, 10#, 10#, 0#, "REBAR", "BAR": space.Add cell
        Set layer = New CFakeAcadContour: layers.Add layer, "OPENING"
        If mode = 0 Or mode = 3 Then
            Set opening = Contour("AcDbPolyline", RectangleEdges(20#, 20#, 60#, 60#), "OPENING"): expected = 6400#
        ElseIf mode = 1 Then
            Set opening = Contour("AcDbRegion", RectangleEdges(15#, 15#, 70#, 70#), "OPENING"): expected = 5100#
        Else
            Set opening = Contour("AcDbPolyline", RectangleEdges(25#, 25#, 50#, 50#), "OPENING"): expected = 6400#
        End If
        space.Add opening
        If mode = 3 Then
            Set opening = Contour("AcDbPolyline", RectangleEdges(85#, 40#, 10#, 10#), "OPENING")
            space.Add opening: expected = expected - 100#
        End If
        Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
        query.Initialize section: Set region = query.ConcreteDomain
        CheckArea "openingOnly.meshVoid." & CStr(mode), region.Area, expected
        Check "openingOnly.noInventedOuter." & CStr(mode), InStr(region.Source, "AuthoritativeContour") = 0
        Check "openingOnly.actualHole." & CStr(mode), Not query.ContainsPoint(region, 50#, 50#)
        Check "openingOnly.sourceKept." & CStr(mode), Not opening.Deleted And section.ConcreteCount = 4
        Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
        Set restored = ReadSectionGeometryFromResults(ThisWorkbook, "AutoCADImport")
        query.Initialize restored
        CheckArea "openingOnly.savedRoundTrip." & CStr(mode), query.ConcreteDomain.Area, expected
    Next mode
    Set opening = Contour("AcDbPolyline", RectangleEdges(87#, 42#, 5#, 5#), "OPENING"): space.Add opening
    Dim number As Long
    On Error Resume Next
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    number = Err.Number: Err.Clear
    On Error GoTo 0
    Check "openingOnly.overlappingExactOpeningsRejected", number <> 0
End Sub

' ДЛЯ ТЕСТОВ: настоящий exporter только добавляет полилинии. Даже повторный
' запуск не читает ModelSpace.Count/Item, не удаляет прежние объекты, сохраняет
' разные цвета outer/opening/Ar и реальные дуги в миллиметрах.
Private Sub CheckAdditiveContourExport()
    Dim section As CSectionModel, writer As CNDMResultsWriter, space As CFakeAcadContourExport
    Dim region As CConcreteRegion, query As CSectionGeometryQuery, entity As CFakeAcadContourExport
    Dim points As Variant, exported As Long, i As Long
    Set section = New CSectionModel
    section.AddConcreteElement 50#, 50#, 10000#, 1, , , "Rectangle", 100#, 100#
    Set query = New CSectionGeometryQuery
    Set region = query.PolygonRegion(ContourRectanglePoints(0#, 0#, 100#, 100#))
    section.Contours.AddMaterialRegion region, "OUTER", "Наружный контур теста."
    section.Contours.AddContourCircle "CONTOUR_HOLE", 50#, 50#, 10#, "Круглое отверстие теста.", "HOLE", "Opening"
    section.Annotations.AddAnnotation "CRACK_REGION_CIRCLE", "CRACK_REGION_TEST", _
        50#, 50#, 12#, 0#, 0#, 0#, vbNullString, 0#, "mm", "Расчетная область теста."
    Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
    Set space = New CFakeAcadContourExport
    exported = SP35ExportSavedContoursToModelSpaceForTests(space)
    Check "export.append.first", exported = 2 And space.CreatedEntities.Count = 2
    exported = SP35ExportSavedContoursToModelSpaceForTests(space)
    Check "export.append.repeat", exported = 2 And space.CreatedEntities.Count = 4
    For i = 1 To 4
        Set entity = space.CreatedEntities(i)
        Check "export.append.closed." & CStr(i), entity.Closed And Not entity.Deleted And IsNDMContourOutput(entity)
        If entity.Layer = "TEST_OPENING" Then
            Check "export.color.opening." & CStr(i), entity.Color = 4
            CheckArea "export.arc.bulge." & CStr(i), entity.BulgeAt(0), Tan(GEOM_PI / 8#)
        Else
            Check "export.color.outer." & CStr(i), entity.Color = 30
            points = entity.Coordinates
            Check "export.mm.coordinates." & CStr(i), Abs(CDbl(points(2)) - CDbl(points(0))) = 100#
        End If
    Next i
    exported = SP35ExportSavedContoursToModelSpaceForTests(space, True)
    Check "export.append.crack", exported = 1 And space.CreatedEntities.Count = 5
    Set entity = space.CreatedEntities(5)
    Check "export.color.crack", entity.Color = 31 And entity.Closed And Not entity.Deleted
    Check "export.append.noExistingObjectRead", space.ForbiddenReads = 0
End Sub

' ДЛЯ ТЕСТОВ: четыре независимые координаты для явного наружного контура.
Private Function ContourRectanglePoints(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Variant
    Dim points(1 To 4, 1 To 2) As Double
    points(1, 1) = x1: points(1, 2) = y1: points(2, 1) = x2: points(2, 2) = y1
    points(3, 1) = x2: points(3, 2) = y2: points(4, 1) = x1: points(4, 2) = y2
    ContourRectanglePoints = points
End Function

' ДЛЯ ТЕСТОВ: точная окружность и повернутый наружный контур дают габариты
' по аналитической границе, а не по меньшей сетке. Без outer сохраняется
' прежний приблизительный размер. INPUT/OUTPUT не меняют исходные координаты.
Private Sub CheckImportedContourDimensions()
    Dim section As CSectionModel, query As CSectionGeometryQuery, region As CConcreteRegion
    Dim writer As CNDMResultsWriter, annotations As CSectionAnnotations, kind As Long, i As Long
    Dim expectedB As Double, expectedH As Double, raw() As TRegionEdge, count As Long, x As Double, y As Double
    Set query = New CSectionGeometryQuery: Set writer = New CNDMResultsWriter
    For kind = 0 To 2
        Set section = New CSectionModel: section.SourceType = "AutoCADImport"
        section.AddConcreteElement 10#, 20#, 400#, 1, , , "Rectangle", 20#, 20#
        expectedB = 20#: expectedH = 20#
        If kind = 1 Then
            Set region = query.CircleRegion(10#, 20#, 100#)
            section.Contours.AddMaterialRegion region, "EXACT_CIRCLE", "Точный наружный контур."
            expectedB = 200#: expectedH = 200#
        ElseIf kind = 2 Then
            Set region = query.PolygonRegion(ContourRectanglePoints(-100#, -40#, 100#, 40#))
            region.CopyEdges raw, count
            For i = 1 To count
                x = raw(i).X1: y = raw(i).Y1
                raw(i).X1 = x * Cos(0.47) - y * Sin(0.47): raw(i).Y1 = x * Sin(0.47) + y * Cos(0.47)
                x = raw(i).X2: y = raw(i).Y2
                raw(i).X2 = x * Cos(0.47) - y * Sin(0.47): raw(i).Y2 = x * Sin(0.47) + y * Cos(0.47)
            Next i
            Set region = query.BoundaryRegion(raw, count)
            section.Contours.AddMaterialRegion region, "EXACT_ROTATED", "Повернутый наружный контур."
            expectedB = 200# * Cos(0.47) + 80# * Sin(0.47)
            expectedH = 200# * Sin(0.47) + 80# * Cos(0.47)
        End If
        writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
        Set annotations = section.Annotations
        For i = 1 To annotations.Count
            If annotations.AnnotationID(i) = "DIM_AUTO_BOUNDS_B" Then CheckArea "dimension.width." & CStr(kind), annotations.Value(i), expectedB
            If annotations.AnnotationID(i) = "DIM_AUTO_BOUNDS_H" Then CheckArea "dimension.height." & CStr(kind), annotations.Value(i), expectedH
            If annotations.AnnotationType(i) = "DIMENSION" Then
                Check "dimension.approxOnlyWithoutOuter." & CStr(kind) & "." & CStr(i), _
                    (annotations.Text(i) = ChrW$(&H2248)) = (kind = 0)
            End If
        Next i
    Next kind
End Sub

' ДЛЯ ТЕСТОВ: наружная граница и отверстия выбираются независимо по своим
' слоям. Исходник одного типа не должен скрывать собственный экспорт другого;
' запись и чтение Results сохраняют все три контура и прежнюю площадь сетки.
Private Sub CheckMixedContourSources()
    Dim space As Collection, layers As Collection, cell As CFakeAcadRegion, layer As CFakeAcadContour
    Dim outer As CFakeAcadContour, opening As CFakeAcadContour, duplicate As CFakeAcadContour
    Dim importer As CAutoCADSectionModelImporter, section As CSectionModel, restored As CSectionModel
    Dim query As CSectionGeometryQuery, region As CConcreteRegion, writer As CNDMResultsWriter
    Dim mode As Long, i As Long, countBefore As Long, prefix As String, outerCount As Long, openingCount As Long
    Set importer = New CAutoCADSectionModelImporter: Set query = New CSectionGeometryQuery
    Set writer = New CNDMResultsWriter
    For mode = 0 To 4
        Set space = New Collection: Set layers = New Collection
        Set cell = New CFakeAcadRegion
        cell.Initialize 10000#, 50#, 50#, 8333333.33333333, 8333333.33333333, 0#, "CONCRETE", "CELL", 0#, True
        space.Add cell
        Set cell = New CFakeAcadRegion
        cell.Initialize GEOM_PI * 25#, 10#, 10#, 100#, 100#, 0#, "REBAR", "BAR": space.Add cell
        Set outer = Contour("AcDbRegion", RectangleEdges(0#, 0#, 100#, 100#))
        outer.SetOutputOwned (mode = 1 Or mode = 3): space.Add outer
        For i = 0 To 1
            Set opening = Contour("AcDbPolyline", RectangleEdges(25# + 35# * i, 30# + 30# * i, 10# + 10# * i, 10#), "OPENING")
            opening.SetOutputOwned (mode = 0 Or mode = 3): space.Add opening
            If mode = 4 Then
                Set duplicate = Contour("AcDbPolyline", RectangleEdges(25# + 35# * i, 30# + 30# * i, 10# + 10# * i, 10#), "OPENING")
                duplicate.SetOutputOwned True: space.Add duplicate
            End If
        Next i
        If mode = 4 Then
            Set duplicate = Contour("AcDbPolyline", RectangleEdges(0#, 0#, 100#, 100#))
            duplicate.SetOutputOwned True: space.Add duplicate
        End If
        Set layer = New CFakeAcadContour: layers.Add layer, "OUTER"
        Set layer = New CFakeAcadContour: layers.Add layer, "OPENING"
        countBefore = space.Count: prefix = "mixedSources." & CStr(mode)
        Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
        query.Initialize section: Set region = query.ConcreteDomain
        CheckArea prefix & ".area", region.Area, 9700#
        Check prefix & ".authoritativeOuter", InStr(region.Source, "AuthoritativeContour") > 0
        Check prefix & ".threeLoops", region.LoopCount = 3
        Check prefix & ".bothOpeningsExcluded", Not query.ContainsPoint(region, 30#, 35#) And Not query.ContainsPoint(region, 70#, 65#)
        Check prefix & ".sourcesPreserved", space.Count = countBefore And Not outer.Deleted And Not opening.Deleted
        CheckArea prefix & ".mechanicalArea", section.ConcreteArea(1), 10000#
        writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
        Set restored = ReadSectionGeometryFromResults(ThisWorkbook, "AutoCADImport")
        query.Initialize restored: Set region = query.ConcreteDomain
        CheckArea prefix & ".resultsArea", region.Area, 9700#
        Check prefix & ".resultsLoops", region.LoopCount = 3
        outerCount = 0: openingCount = 0
        For i = 1 To restored.Contours.Count
            If restored.Contours.LoopRole(i) = "Opening" Then openingCount = openingCount + 1 Else outerCount = outerCount + 1
        Next i
        Check prefix & ".resultsBothRoles", outerCount = 4 And openingCount = 8
    Next mode
End Sub

' ДЛЯ ТЕСТОВ: смешанные источники сохраняют встроенное отверстие Region
' и бетонную часть внутри проема. Повторная собственная граница не вычитается
' дважды, но исходное отверстие вне материала остается ошибкой, а не outer.
Private Sub CheckMixedContourTopology()
    Dim space As Collection, layers As Collection, cell As CFakeAcadRegion, layer As CFakeAcadContour
    Dim outer As CFakeAcadContour, part As CFakeAcadContour, hole As CFakeAcadContour, edges As Collection, edge As Variant
    Dim importer As CAutoCADSectionModelImporter, section As CSectionModel, query As CSectionGeometryQuery
    Dim region As CConcreteRegion, number As Long, description As String
    Set space = New Collection: Set layers = New Collection
    Set cell = New CFakeAcadRegion
    cell.Initialize 10000#, 50#, 50#, 8333333.33333333, 8333333.33333333, 0#, "CONCRETE", "CELL", 0#, True: space.Add cell
    Set cell = New CFakeAcadRegion
    cell.Initialize GEOM_PI * 25#, 10#, 10#, 100#, 100#, 0#, "REBAR", "BAR": space.Add cell
    Set layer = New CFakeAcadContour: layers.Add layer, "OUTER"
    Set layer = New CFakeAcadContour: layers.Add layer, "OPENING"
    Set edges = RectangleEdges(0#, 0#, 100#, 100#)
    For Each edge In RectangleEdges(25#, 30#, 10#, 10#): edges.Add edge: Next edge
    Set outer = Contour("AcDbRegion", edges): space.Add outer
    Set hole = Contour("AcDbPolyline", RectangleEdges(25#, 30#, 10#, 10#), "OPENING")
    hole.SetOutputOwned True: space.Add hole
    Set hole = Contour("AcDbPolyline", RectangleEdges(60#, 60#, 20#, 10#), "OPENING")
    hole.SetOutputOwned True: space.Add hole
    Set importer = New CAutoCADSectionModelImporter: Set query = New CSectionGeometryQuery
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    query.Initialize section: Set region = query.ConcreteDomain
    CheckArea "mixedTopology.embeddedOpening.area", region.Area, 9700#
    Check "mixedTopology.embeddedOpening.threeLoops", region.LoopCount = 3
    Check "mixedTopology.embeddedOpening.bothExcluded", Not query.ContainsPoint(region, 30#, 35#) And Not query.ContainsPoint(region, 70#, 65#)
    space.Remove 5: space.Remove 4: space.Remove 3
    Set outer = Contour("AcDbPolyline", RectangleEdges(0#, 0#, 100#, 100#)): outer.SetOutputOwned True: space.Add outer
    Set part = Contour("AcDbPolyline", RectangleEdges(40#, 40#, 20#, 20#)): part.SetOutputOwned True: space.Add part
    Set hole = Contour("AcDbRegion", RectangleEdges(20#, 20#, 60#, 60#), "OPENING"): space.Add hole
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    query.Initialize section: Set region = query.ConcreteDomain
    CheckArea "mixedTopology.separateConcrete.area", region.Area, 6800#
    Check "mixedTopology.separateConcrete.threeLoops", region.LoopCount = 3
    Check "mixedTopology.separateConcrete.roles", query.ContainsPoint(region, 50#, 50#) And Not query.ContainsPoint(region, 30#, 30#)
    space.Add Contour("AcDbPolyline", RectangleEdges(120#, 20#, 10#, 10#), "OPENING")
    On Error Resume Next
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo 0
    Check "mixedTopology.outsideOpening.rejected", number <> 0 And InStr(description, "слое") > 0
    Check "mixedTopology.outsideOpening.oldModelKept", section.Contours.Count = 12 And Abs(query.ConcreteDomain.Area - 6800#) < 0.00001
End Sub

' ДЛЯ ТЕСТОВ: отсутствующие общие слои не отменяют сетку; отдельный Region
' внутри уже вырезанного проема возвращает бетон, а не заполняет весь проем.
' Экспортированная копия на том же слое не удваивает исходный контур.
Private Sub CheckOptionalLayersAndSeparateConcrete()
    Dim space As Collection, layers As Collection, cell As CFakeAcadRegion, dummy As CFakeAcadContour
    Dim importer As CAutoCADSectionModelImporter, section As CSectionModel, query As CSectionGeometryQuery
    Dim region As CConcreteRegion, outer As CFakeAcadContour, holeEdges As Collection, edge As Variant
    Dim part As CFakeAcadContour, duplicate As CFakeAcadContour, countBefore As Long
    Set space = New Collection: Set layers = New Collection
    Set cell = New CFakeAcadRegion
    cell.Initialize 10000#, 50#, 50#, 8333333.33333333, 8333333.33333333, 0#, "CONCRETE", "CELL", 0#, True
    space.Add cell
    Set cell = New CFakeAcadRegion
    cell.Initialize GEOM_PI * 25#, 50#, 50#, 100#, 100#, 0#, "REBAR", "BAR": space.Add cell
    Set importer = New CAutoCADSectionModelImporter
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    Check "optional.absentLayers.meshKept", section.ConcreteCount = 1 And section.RebarCount = 1 And section.Contours.Count = 0
    Check "optional.absentLayers.comment", InStr(importer.ImportComment, "OUTER") > 0 And InStr(importer.ImportComment, "OPENING") > 0
    Set holeEdges = RectangleEdges(0#, 0#, 100#, 100#)
    For Each edge In RectangleEdges(20#, 20#, 60#, 60#): holeEdges.Add edge: Next edge
    Set outer = Contour("AcDbRegion", holeEdges): space.Add outer
    Set part = Contour("AcDbRegion", RectangleEdges(40#, 40#, 20#, 20#)): space.Add part
    Set dummy = New CFakeAcadContour: layers.Add dummy, "OUTER"
    Set duplicate = Contour("AcDbPolyline", holeEdges): duplicate.SetOutputOwned True: space.Add duplicate
    countBefore = space.Count
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    Set query = New CSectionGeometryQuery: query.Initialize section: Set region = query.ConcreteDomain
    CheckArea "separateConcrete.area", region.Area, 6800#
    Check "separateConcrete.threeContours", region.LoopCount = 3
    Check "separateConcrete.material", query.ContainsPoint(region, 50#, 50#) And query.ContainsPoint(region, 10#, 10#)
    Check "separateConcrete.void", Not query.ContainsPoint(region, 30#, 30#)
    Check "separateConcrete.sourcesKept", space.Count = countBefore And Not outer.Deleted And Not part.Deleted And Not duplicate.Deleted
    Check "separateConcrete.openingAbsent", InStr(importer.ImportComment, "OPENING") > 0
    ' Новый чертеж содержит только экспорт: наружная граница, отверстие
    ' и отдельная бетонная часть должны восстановить те же три контура.
    space.Remove 5: space.Remove 4: space.Remove 3
    Set outer = Contour("AcDbPolyline", RectangleEdges(0#, 0#, 100#, 100#)): outer.SetOutputOwned True: space.Add outer
    Set part = Contour("AcDbPolyline", RectangleEdges(40#, 40#, 20#, 20#)): part.SetOutputOwned True: space.Add part
    Set duplicate = Contour("AcDbPolyline", RectangleEdges(20#, 20#, 60#, 60#), "OPENING")
    duplicate.SetOutputOwned True: space.Add duplicate
    Set dummy = New CFakeAcadContour: layers.Add dummy, "OPENING"
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    query.Initialize section: Set region = query.ConcreteDomain
    CheckArea "exportOnly.separateConcrete.area", region.Area, 6800#
    Check "exportOnly.separateConcrete.threeContours", region.LoopCount = 3
    Check "exportOnly.separateConcrete.material", query.ContainsPoint(region, 50#, 50#) And Not query.ContainsPoint(region, 30#, 30#)
End Sub

' ДЛЯ ТЕСТОВ: существующие пустые слои не делают контуры обязательными.
' Проверяет пересечение по бетонной сетке, отдельные отверстия без outer
' и сохранение строгой ошибки для реально заданной незамкнутой полилинии.
Private Sub CheckEmptyContourLayers()
    Dim space As Collection, layers As Collection, cell As CFakeAcadRegion, bar As CFakeAcadRegion
    Dim layer As CFakeAcadContour, opening As CFakeAcadContour, invalid As CFakeAcadContour
    Dim importer As CAutoCADSectionModelImporter, section As CSectionModel, query As CSectionGeometryQuery
    Dim domain As CConcreteRegion, clipped As CConcreteRegion, mode As Long, expectedArea As Double
    Dim contourLayer As String, openingLayer As String, prefix As String, number As Long, description As String
    Set importer = New CAutoCADSectionModelImporter
    For mode = 0 To 3
        Set space = New Collection: Set layers = New Collection
        Set cell = New CFakeAcadRegion
        cell.Initialize 10000#, 50#, 50#, 8333333.33333333, 8333333.33333333, 0#, "CONCRETE", "CELL", 0#, True
        space.Add cell
        Set bar = New CFakeAcadRegion
        bar.Initialize GEOM_PI * 25#, 10#, 10#, 100#, 100#, 0#, "REBAR", "BAR": space.Add bar
        contourLayer = "OUTER": openingLayer = "OPENING": expectedArea = 10000#
        Set layer = New CFakeAcadContour: layers.Add layer, "OUTER"
        If mode > 0 Then
            Set layer = New CFakeAcadContour: layers.Add layer, "OPENING"
        End If
        If mode = 2 Then
            Set opening = Contour("AcDbPolyline", RectangleEdges(60#, 60#, 20#, 20#), "OPENING")
            space.Add opening: expectedArea = 9600#
        ElseIf mode = 3 Then
            contourLayer = vbNullString: openingLayer = vbNullString
        End If
        Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, contourLayer, openingLayer, layers)
        Set query = New CSectionGeometryQuery: query.Initialize section: Set domain = query.ConcreteDomain
        prefix = "emptyLayers." & CStr(mode)
        Check prefix & ".meshKept", section.ConcreteCount = 1 And section.RebarCount = 1
        Check prefix & ".noInventedContour", section.Contours.Count = IIf(mode = 2, 4, 0)
        Check prefix & ".meshFallback", InStr(domain.Source, "AuthoritativeContour") = 0
        Check prefix & ".comment", InStr(importer.ImportComment, "бетонной сетке") > 0
        Check prefix & ".sourceKept", Not cell.Deleted And Not bar.Deleted
        CheckArea prefix & ".area", domain.Area, expectedArea
        Set clipped = query.ClipHalfPlane(domain, 1#, 0#, 50#)
        CheckArea prefix & ".meshIntersection", clipped.Area, expectedArea - 5000#
        If mode = 2 Then Check prefix & ".openingKept", Not query.ContainsPoint(domain, 70#, 70#) And Not opening.Deleted
    Next mode
    Set invalid = New CFakeAcadContour
    invalid.InitializeLoop "AcDbPolyline", "OUTER", "OPEN", RectangleEdges(0#, 0#, 100#, 100#), False
    space.Add invalid
    On Error Resume Next
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo 0
    Check "emptyLayers.invalidContourStillRejected", number <> 0 And InStr(description, "замкнутой") > 0
    Check "emptyLayers.invalidSourceKept", Not invalid.Deleted
    Check "emptyLayers.failedImportKeepsPreviousModel", section.ConcreteCount = 1 And section.RebarCount = 1 And section.Contours.Count = 0
End Sub

' ДЛЯ ТЕСТОВ: сохраняет несколько отверстий и аналитические дуги через
' штатный preview writer, меняет текущие единицы и восстанавливает модель.
' Проверяет также сценарий только opening и адрес поврежденного snapshot.
Private Sub CheckResultsRoundTrip()
    Dim unitRange As Object, savedUnits As Variant, settings As CSystemSettingsReader, units As CUnitSystem
    Dim section As CSectionModel, restored As CSectionModel, query As CSectionGeometryQuery, region As CConcreteRegion
    Dim reader As CAutoCADContourReader, writer As CNDMResultsWriter, contourEntity As CFakeAcadContour
    Dim edges As Collection, curve As CFakeAcadContour, mode As Long, row As Long, lengthUnit As String
    Dim number As Long, description As String, anchor As Object, savedValue As Variant
    Dim ix As Double, iy As Double, ixy As Double, failureNumber As Long, failureDescription As String
    On Error GoTo Failed
    Set unitRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange: savedUnits = unitRange.Formula
    Set reader = New CAutoCADContourReader: Set writer = New CNDMResultsWriter
    For mode = 0 To 3
        If mode Mod 2 = 0 Then lengthUnit = "mm" Else lengthUnit = "m"
        For row = 2 To unitRange.Rows.Count
            If CStr(unitRange.Cells(row, 1).Value2) = "Length" Then unitRange.Cells(row, 4).Value2 = lengthUnit
            If CStr(unitRange.Cells(row, 1).Value2) = "Area" Then unitRange.Cells(row, 4).Value2 = lengthUnit & "2"
        Next row
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set section = New CSectionModel: section.SourceType = "AutoCADImport"
        section.AddConcreteElement 50#, 50#, 10000#, 1, , , "Rectangle", 100#, 100#, 0#, , 123456#, 654321#, 12#
        section.AddRebarElement 10#, 10#, 10#, GEOM_PI * 25#, "Rebar"
        If mode < 2 Then
            Set edges = RectangleEdges(0#, 0#, 100#, 100#)
            Set curve = New CFakeAcadContour: curve.InitializeCircle 35#, 35#, 10#: edges.Add curve
            Set curve = New CFakeAcadContour: curve.InitializeCircle 70#, 65#, 5#: edges.Add curve
            Set contourEntity = Contour("AcDbRegion", edges)
            Set region = reader.ReadRegion(contourEntity)
            section.Contours.AddMaterialRegion region, "ROUNDTRIP", "Точный наружный контур с двумя проемами."
        Else
            For row = 0 To 1
                Set curve = New CFakeAcadContour: curve.InitializeCircle 35# + 35# * row, 35# + 30# * row, 10# - 5# * row
                Set region = reader.ReadRegion(curve)
                section.Contours.AddMaterialRegion region, "HOLE_" & CStr(row), "Отдельный проем.", True
            Next row
        End If
        writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section), units
        ' Новые input/output настройки не меняют единицы сохраненного снимка.
        For row = 2 To unitRange.Rows.Count
            If CStr(unitRange.Cells(row, 1).Value2) = "Length" Then
                unitRange.Cells(row, 2).Value2 = "cm"
                If lengthUnit = "m" Then unitRange.Cells(row, 4).Value2 = "mm" Else unitRange.Cells(row, 4).Value2 = "m"
            End If
        Next row
        Set restored = ReadSectionGeometryFromResults(ThisWorkbook, "AutoCADImport")
        Set query = New CSectionGeometryQuery: query.Initialize restored: Set region = query.ConcreteDomain
        CheckArea "snapshot." & CStr(mode) & ".area", region.Area, 10000# - 125# * GEOM_PI
        Check "snapshot." & CStr(mode) & ".threeLoops", region.LoopCount = 3
        Check "snapshot." & CStr(mode) & ".holes", Not query.ContainsPoint(region, 35#, 35#) And Not query.ContainsPoint(region, 70#, 65#)
        Check "snapshot." & CStr(mode) & ".material", query.ContainsPoint(region, 50#, 50#)
        CheckArea "snapshot." & CStr(mode) & ".mechanicalArea", restored.ConcreteArea(1), 10000#
        restored.ConcreteLocalInertiaComponents 1, ix, iy, ixy
        Check "snapshot." & CStr(mode) & ".mechanicalInertia", Abs(ix - 123456#) < 0.00001 And Abs(iy - 654321#) < 0.00001 And Abs(ixy - 12#) < 0.00001
        If mode < 2 Then
            Check "snapshot." & CStr(mode) & ".authoritative", InStr(region.Source, "AuthoritativeContour") > 0
        Else
            Check "snapshot." & CStr(mode) & ".meshFallback", InStr(region.Source, "AuthoritativeContour") = 0
        End If
        Set anchor = ThisWorkbook.Names.Item("rngNDMSectionContours").RefersToRange
        Dim contourRow As Long, annotationAnchor As Object
        contourRow = 1
        Check "snapshot." & CStr(mode) & ".contoursOwnBlock", anchor.Value2 = "RunID" And anchor.Offset(contourRow, 5).Value2 Like "CONTOUR_*"
        Check "snapshot." & CStr(mode) & ".explicitRole", Len(CStr(anchor.Offset(contourRow, 4).Value2)) > 0
        Set annotationAnchor = ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange
        Check "snapshot." & CStr(mode) & ".noSourceAnnotations", Application.CountIf(annotationAnchor.Offset(1, 0).Resize(1000, 1), "CONTOUR_*") = 0
        savedValue = anchor.Offset(contourRow, 6).Value2: anchor.Offset(contourRow, 6).Value2 = "invalid"
        On Error Resume Next
        Set restored = ReadSectionGeometryFromResults(ThisWorkbook, "AutoCADImport")
        number = Err.Number: description = Err.Description: Err.Clear
        On Error GoTo Failed
        anchor.Offset(contourRow, 6).Value2 = savedValue
        Check "snapshot." & CStr(mode) & ".invalidCoordinate", number <> 0 And InStr(description, anchor.Offset(contourRow, 6).Address(False, False)) > 0
    Next mode
    GoTo Restore
Failed:
    failureNumber = Err.Number: failureDescription = Err.Description
Restore:
    On Error Resume Next
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, "CheckResultsRoundTrip", failureDescription
End Sub

' ДЛЯ ТЕСТОВ: сохраняет утверждение и независимый численный эталон.
Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        mPassed = mPassed + 1: mReport = mReport & "OK: " & name & vbCrLf
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
    End If
End Sub

' ДЛЯ ТЕСТОВ: несколько opening читаются отдельными Region/Polyline как с
' достоверным outer, так и при прежней внешней границе сетки. Последующий
' ошибочный импорт не меняет уже полученную пригодную модель.
Private Sub CheckMultipleOpenings()
    Dim space As Collection, layers As Collection, cell As CFakeAcadRegion, layer As CFakeAcadContour
    Dim firstOpening As CFakeAcadContour, secondOpening As CFakeAcadContour, outer As CFakeAcadContour
    Dim importer As CAutoCADSectionModelImporter, section As CSectionModel, query As CSectionGeometryQuery
    Dim region As CConcreteRegion, mode As Long, outerLayer As String, number As Long, description As String, annotationCount As Long
    For mode = 0 To 1
        Set space = New Collection: Set layers = New Collection
        Set cell = New CFakeAcadRegion
        cell.Initialize 10000#, 50#, 50#, 8333333.33333333, 8333333.33333333, 0#, "CONCRETE", "CELL", 0#, True
        space.Add cell
        Set cell = New CFakeAcadRegion
        cell.Initialize GEOM_PI * 25#, 10#, 10#, 100#, 100#, 0#, "REBAR", "BAR": space.Add cell
        outerLayer = vbNullString
        If mode = 1 Then
            Set outer = Contour("AcDbRegion", RectangleEdges(0#, 0#, 100#, 100#))
            space.Add outer: outerLayer = "OUTER"
        End If
        Set firstOpening = Contour("AcDbRegion", RectangleEdges(25#, 30#, 10#, 10#), "OPENING")
        Set secondOpening = Contour("AcDbPolyline", RectangleEdges(60#, 60#, 20#, 10#), "OPENING")
        space.Add firstOpening: space.Add secondOpening
        Set layer = New CFakeAcadContour: layers.Add layer, "OUTER"
        Set layer = New CFakeAcadContour: layers.Add layer, "OPENING"
        Set importer = New CAutoCADSectionModelImporter
        Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, outerLayer, "OPENING", layers)
        Set query = New CSectionGeometryQuery: query.Initialize section: Set region = query.ConcreteDomain
        CheckArea "multiple." & CStr(mode) & ".area", region.Area, 9700#
        Check "multiple." & CStr(mode) & ".threeLoops", region.LoopCount = 3
        Check "multiple." & CStr(mode) & ".bothExcluded", Not query.ContainsPoint(region, 30#, 35#) And Not query.ContainsPoint(region, 70#, 65#)
        Check "multiple." & CStr(mode) & ".material", query.ContainsPoint(region, 50#, 50#)
        Check "multiple." & CStr(mode) & ".sourceKept", Not firstOpening.Deleted And Not secondOpening.Deleted
        If mode = 0 Then Check "multiple.meshOuterNotAuthoritative", InStr(region.Source, "AuthoritativeContour") = 0
        annotationCount = section.Contours.Count
        space.Add Contour("AcDbRegion", RectangleEdges(60#, 60#, 20#, 10#), "OPENING")
        On Error Resume Next
        Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, outerLayer, "OPENING", layers)
        number = Err.Number: description = Err.Description: Err.Clear
        On Error GoTo 0
        Check "multiple." & CStr(mode) & ".duplicateRejected", number <> 0 And Len(description) > 25
        Check "multiple." & CStr(mode) & ".oldModelKept", section.Contours.Count = annotationCount
        Check "multiple." & CStr(mode) & ".oldAreaKept", Abs(query.ConcreteDomain.Area - 9700#) < 0.00001
    Next mode
End Sub

' ДЛЯ ТЕСТОВ: число отдельных opening не ограничено двумя. Пять проемов
' разных поддержанных типов остаются независимыми замкнутыми контурами, вычитаются один
' раз и сохраняются при записи/чтении Results в собственных единицах снимка.
Private Sub CheckFiveOpenings()
    Dim space As Collection, layers As Collection, cell As CFakeAcadRegion, layer As CFakeAcadContour
    Dim importer As CAutoCADSectionModelImporter, section As CSectionModel, restored As CSectionModel
    Dim query As CSectionGeometryQuery, region As CConcreteRegion, writer As CNDMResultsWriter
    Dim i As Long, kind As String, opening As CFakeAcadContour
    Set space = New Collection: Set layers = New Collection
    Set cell = New CFakeAcadRegion
    cell.Initialize 10000#, 50#, 50#, 8333333.33333333, 8333333.33333333, 0#, "CONCRETE", "CELL", 0#, True
    space.Add cell
    Set cell = New CFakeAcadRegion
    cell.Initialize GEOM_PI * 25#, 5#, 5#, 100#, 100#, 0#, "REBAR", "BAR": space.Add cell
    space.Add Contour("AcDbRegion", RectangleEdges(0#, 0#, 100#, 100#))
    Set layer = New CFakeAcadContour: layers.Add layer, "OUTER"
    Set layer = New CFakeAcadContour: layers.Add layer, "OPENING"
    For i = 1 To 5
        kind = "AcDbRegion"
        If i Mod 2 = 0 Then kind = "AcDbPolyline"
        Set opening = Contour(kind, RectangleEdges(10# + 15# * (i - 1), 40#, 5#, 10#), "OPENING")
        space.Add opening
    Next i
    Set importer = New CAutoCADSectionModelImporter
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    Set query = New CSectionGeometryQuery: query.Initialize section: Set region = query.ConcreteDomain
    CheckArea "fiveOpenings.area", region.Area, 9750#
    Check "fiveOpenings.sixLoops", region.LoopCount = 6
    For i = 1 To 5
        Check "fiveOpenings.excluded." & CStr(i), Not query.ContainsPoint(region, 12.5 + 15# * (i - 1), 45#)
    Next i
    Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
    Set restored = ReadSectionGeometryFromResults(ThisWorkbook, "AutoCADImport")
    query.Initialize restored: Set region = query.ConcreteDomain
    CheckArea "fiveOpenings.snapshotArea", region.Area, 9750#
    Check "fiveOpenings.snapshotLoops", region.LoopCount = 6
End Sub

' ДЛЯ ТЕСТОВ: проверяет площадь в мм2 без вызова другой production формулы.
Private Sub CheckArea(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    Check name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) < 0.00001
End Sub

' ДЛЯ ТЕСТОВ: создает неупорядоченный прямоугольник с обращенными гранями.
Private Function RectangleEdges(ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal height As Double, _
        Optional ByVal angle As Double = 0#) As Collection
    Dim points(1 To 4, 1 To 2) As Double, i As Long, j As Long, lx As Double, ly As Double
    Dim edges As Collection, line As CFakeAcadContour, order As Variant
    Set edges = New Collection: order = Array(3, 1, 4, 2)
    For i = 1 To 4
        Select Case i
            Case 1: lx = 0#: ly = 0#
            Case 2: lx = width: ly = 0#
            Case 3: lx = width: ly = height
            Case 4: lx = 0#: ly = height
        End Select
        points(i, 1) = x + lx * Cos(angle) - ly * Sin(angle)
        points(i, 2) = y + lx * Sin(angle) + ly * Cos(angle)
    Next i
    For i = 0 To 3
        j = CLng(order(i)): lx = j + 1: If lx = 5 Then lx = 1
        Set line = New CFakeAcadContour
        If i Mod 2 = 0 Then
            line.InitializeLine points(CLng(lx), 1), points(CLng(lx), 2), points(j, 1), points(j, 2)
        Else
            line.InitializeLine points(j, 1), points(j, 2), points(CLng(lx), 1), points(CLng(lx), 2)
        End If
        edges.Add line
    Next i
    Set RectangleEdges = edges
End Function

' ДЛЯ ТЕСТОВ: оборачивает ребра в Region или Polyline с одним COM-контрактом.
Private Function Contour(ByVal kind As String, ByVal edges As Collection, Optional ByVal layerName As String = "OUTER") As CFakeAcadContour
    Dim result As CFakeAcadContour
    Set result = New CFakeAcadContour
    result.InitializeLoop kind, layerName, "C_TEST", edges
    Set Contour = result
End Function

' ДЛЯ ТЕСТОВ: Region и обе 2D-полилинии дают одну геометрию после поворота.
Private Sub CheckEquivalentSources()
    Dim reader As CAutoCADContourReader, query As CSectionGeometryQuery, entity As CFakeAcadContour
    Dim region As CConcreteRegion, kind As Variant, curve As CFakeAcadContour
    Set reader = New CAutoCADContourReader: Set query = New CSectionGeometryQuery
    For Each kind In Array("AcDbRegion", "AcDbPolyline", "AcDb2dPolyline")
        Set entity = Contour(CStr(kind), RectangleEdges(137#, -73#, 100#, 80#, 0.37))
        Set region = reader.ReadRegion(entity)
        CheckArea CStr(kind) & ".area", region.Area, 8000#
        Check CStr(kind) & ".actualEdges", region.SegmentCount = 4 And region.LoopCount = 1
        Check CStr(kind) & ".outer", Not query.LoopIsOpening(region, 1)
        Check CStr(kind) & ".sourceKept", Not entity.Deleted
        Check CStr(kind) & ".copyDeleted", entity.LastCopy.Deleted
        For Each curve In entity.LastCopy.ExplodedPieces: Check CStr(kind) & ".pieceDeleted", curve.Deleted: Next curve
    Next kind
End Sub

' ДЛЯ ТЕСТОВ: сохраняет внутреннюю окружность и вложенный остров материала.
' Начальные направления Explode не определяют семантику outer/opening.
Private Sub CheckNestedRegion()
    Dim edges As Collection, island As Collection, curve As CFakeAcadContour, entity As CFakeAcadContour
    Dim reader As CAutoCADContourReader, query As CSectionGeometryQuery, region As CConcreteRegion, openings As Long, i As Long
    Set edges = RectangleEdges(0#, 0#, 100#, 100#)
    Set curve = New CFakeAcadContour: curve.InitializeCircle 50#, 50#, 20#: edges.Add curve
    Set island = RectangleEdges(45#, 45#, 10#, 10#)
    For Each curve In island: edges.Add curve: Next curve
    Set entity = Contour("AcDbRegion", edges)
    Set reader = New CAutoCADContourReader: Set query = New CSectionGeometryQuery
    Set region = reader.ReadRegion(entity)
    CheckArea "nested.area", region.Area, 10100# - 400# * GEOM_PI
    Check "nested.threeLoops", region.LoopCount = 3
    For i = 1 To region.LoopCount
        If query.LoopIsOpening(region, i) Then openings = openings + 1
    Next i
    Check "nested.oneOpening", openings = 1
    Check "nested.holeExcluded", Not query.ContainsPoint(region, 65#, 50#)
    Check "nested.islandIncluded", query.ContainsPoint(region, 50#, 50#)
    Check "nested.outerIncluded", query.ContainsPoint(region, 5#, 50#)
    Check "nested.sourceKept", Not entity.Deleted
End Sub

' ДЛЯ ТЕСТОВ: неподдержанные/неплоские/дублирующие границы не превращаются
' в приблизительный успешный контур; собственная Copy очищается и при ошибке.
Private Sub CheckInvalidContours()
    Dim reader As CAutoCADContourReader, entity As CFakeAcadContour, edges As Collection, curve As CFakeAcadContour
    Dim region As CConcreteRegion, mode As Long, number As Long, description As String
    Set reader = New CAutoCADContourReader
    For mode = 1 To 8
        Set edges = RectangleEdges(0#, 0#, 100#, 80#)
        Set entity = Contour("AcDbRegion", edges)
        Select Case mode
            Case 1: edges.Remove 1
            Case 2: edges.Add edges(1)
            Case 3: entity.SetPlane 1#, 0#, 0#, 0#
            Case 4: Set entity = Contour("AcDbSpline", edges)
            Case 5: entity.InitializeLoop "AcDbPolyline", "OUTER", "OPEN", edges, False
            Case 6: entity.InitializeLoop "AcDbPolyline", "OUTER", "Z", edges: entity.SetPlane 0#, 0#, 1#, 1#
            Case 7
                Set curve = New CFakeAcadContour: curve.InitializeLoop "AcDbEllipse", "OUTER", "ELLIPSE", Nothing: edges.Add curve
            Case 8
                Set edges = New Collection
                Set curve = New CFakeAcadContour: curve.InitializeLine 0#, 0#, 100#, 80#: edges.Add curve
                Set curve = New CFakeAcadContour: curve.InitializeLine 100#, 80#, 0#, 80#: edges.Add curve
                Set curve = New CFakeAcadContour: curve.InitializeLine 0#, 80#, 100#, 0#: edges.Add curve
                Set curve = New CFakeAcadContour: curve.InitializeLine 100#, 0#, 0#, 0#: edges.Add curve
                Set entity = Contour("AcDbRegion", edges)
        End Select
        On Error Resume Next
        Set region = reader.ReadRegion(entity)
        number = Err.Number: description = Err.Description: Err.Clear
        On Error GoTo 0
        Check "invalid." & CStr(mode) & ".error", number <> 0
        Check "invalid." & CStr(mode) & ".context", InStr(description, "OUTER") > 0 And Len(description) > 25
        Check "invalid." & CStr(mode) & ".sourceKept", Not entity.Deleted
        If Not entity.LastCopy Is Nothing Then Check "invalid." & CStr(mode) & ".copyDeleted", entity.LastCopy.Deleted
    Next mode
End Sub

' ДЛЯ ТЕСТОВ: реальный importer сохраняет Region с opening, не меняя A/I
' дискретных элементов. Повторный geometry-query читает его semantic-снимок.
Private Sub CheckImportedSnapshot()
    Dim space As Collection, layers As Collection, regionEntity As CFakeAcadRegion, contourEntity As CFakeAcadContour
    Dim edges As Collection, roundOpening As CFakeAcadContour, importer As CAutoCADSectionModelImporter
    Dim section As CSectionModel, query As CSectionGeometryQuery, domain As CConcreteRegion, layer As CFakeAcadContour
    Set space = New Collection: Set layers = New Collection
    Set regionEntity = New CFakeAcadRegion
    regionEntity.Initialize 10000#, 50#, 50#, 1000000#, 1200000#, 0#, "CONCRETE", "C1"
    space.Add regionEntity
    Set regionEntity = New CFakeAcadRegion
    regionEntity.Initialize GEOM_PI * 25#, 10#, 10#, 100#, 100#, 0#, "REBAR", "R1"
    space.Add regionEntity
    Set edges = RectangleEdges(0#, 0#, 100#, 100#)
    Set roundOpening = New CFakeAcadContour: roundOpening.InitializeCircle 50#, 50#, 20#: edges.Add roundOpening
    Set contourEntity = Contour("AcDbRegion", edges): space.Add contourEntity
    Set layer = New CFakeAcadContour: layers.Add layer, "OUTER"
    Set layer = New CFakeAcadContour: layers.Add layer, "OPENING"
    Set importer = New CAutoCADSectionModelImporter
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    Set query = New CSectionGeometryQuery: query.Initialize section: Set domain = query.ConcreteDomain
    CheckArea "import.domain", domain.Area, 10000# - 400# * GEOM_PI
    CheckArea "import.mechanicalAreaKept", section.ConcreteArea(1), 10000#
    Check "import.holeKept", domain.LoopCount = 2 And Not query.ContainsPoint(domain, 50#, 50#)
    Check "import.sourceKept", Not contourEntity.Deleted
    Check "import.emptyOpeningExplained", InStr(section.Contours.Comment(1), "объектов нет") > 0
End Sub

' ДЛЯ ТЕСТОВ: проверяет настоящий Autodesk AutoCAD в собственном новом
' документе. Существующие чертежи не изменяются; source-объекты сохраняют
' handles, площадь и число сущностей после Copy/Explode production reader-а.
Public Function RunRealAutoCADContourTests(Optional ByVal ownedApplication As Object = Nothing) As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    Dim acad As Object, doc As Object, previous As Object, polyline As Object, source As Object
    Dim circleEntity As Object, openingRegion As Object, reader As CAutoCADContourReader, query As CSectionGeometryQuery
    Dim region As CConcreteRegion, curves(0 To 0) As Object, regions As Variant
    Dim countBefore As Long, handleBefore As String, areaBefore As Double, mode As Long, i As Long, openings As Long
    Dim path As String, center(0 To 2) As Double, semi(0 To 3) As Double
    On Error GoTo Failed
    If ownedApplication Is Nothing Then
        Set acad = GetObject(, "AutoCAD.Application.24.2")
    Else
        Set acad = ownedApplication
    End If
    Check "native.actualAutoCAD", InStr(1, CStr(acad.FullName), "\AutoCAD 2023\acad.exe", vbTextCompare) > 0
    If InStr(1, CStr(acad.FullName), "\AutoCAD 2023\acad.exe", vbTextCompare) = 0 Then Err.Raise vbObjectError + 5502, , "Тест требует Autodesk AutoCAD 2023, а не SOFiPLUS."
    If acad.Documents.Count > 0 Then Set previous = acad.ActiveDocument
    Set doc = acad.Documents.Add
    WaitForCAD acad
    Set reader = New CAutoCADContourReader: Set query = New CSectionGeometryQuery
    Set polyline = NativeRectangle(doc, 100#, 80#, 0.37, 137#, -73#)
    Set curves(0) = polyline: regions = doc.ModelSpace.AddRegion(curves): Set source = regions(LBound(regions))
    For mode = 0 To 1
        If mode = 0 Then Set circleEntity = polyline Else Set circleEntity = source
        countBefore = doc.ModelSpace.Count: handleBefore = CStr(circleEntity.Handle): areaBefore = CDbl(circleEntity.Area)
        Set region = reader.ReadRegion(circleEntity)
        CheckArea "native.rectangle." & CStr(mode), region.Area, 8000#
        Check "native.rectangle.source." & CStr(mode), CStr(circleEntity.Handle) = handleBefore And doc.ModelSpace.Count = countBefore
        CheckArea "native.rectangle.unchanged." & CStr(mode), CDbl(circleEntity.Area), areaBefore
    Next mode
    ' Один bulge задает точную полуокружность; обе формы источника должны
    ' сохранить круговую дугу, а не заменять ее цепочкой прямых.
    semi(0) = 0#: semi(1) = 0#: semi(2) = 100#: semi(3) = 0#
    Set polyline = doc.ModelSpace.AddLightWeightPolyline(semi): polyline.Closed = True: polyline.SetBulge 0, 1#
    Set curves(0) = polyline: regions = doc.ModelSpace.AddRegion(curves): Set source = regions(LBound(regions))
    For mode = 0 To 1
        If mode = 0 Then Set circleEntity = polyline Else Set circleEntity = source
        countBefore = doc.ModelSpace.Count
        Set region = reader.ReadRegion(circleEntity)
        CheckArea "native.semicircle." & CStr(mode), region.Area, 1250# * GEOM_PI
        Check "native.semicircle.analytic." & CStr(mode), region.SegmentCount = 2
        Check "native.semicircle.cleanup." & CStr(mode), doc.ModelSpace.Count = countBefore
    Next mode
    Set polyline = NativeRectangle(doc, 100#, 100#, 0#, 250#, 100#)
    Set curves(0) = polyline: regions = doc.ModelSpace.AddRegion(curves): Set source = regions(LBound(regions))
    For mode = 0 To 1
        center(0) = 225# + 45# * mode: center(1) = 85# + 25# * mode
        Set circleEntity = doc.ModelSpace.AddCircle(center, 10# - 5# * mode)
        Set curves(0) = circleEntity: regions = doc.ModelSpace.AddRegion(curves): Set openingRegion = regions(LBound(regions))
        source.Boolean 2, openingRegion
    Next mode
    countBefore = doc.ModelSpace.Count: handleBefore = CStr(source.Handle): areaBefore = CDbl(source.Area)
    Set region = reader.ReadRegion(source)
    CheckArea "native.twoOpenings.area", region.Area, 10000# - 125# * GEOM_PI
    Check "native.twoOpenings.loops", region.LoopCount = 3
    For i = 1 To region.LoopCount
        If query.LoopIsOpening(region, i) Then openings = openings + 1
    Next i
    Check "native.twoOpenings.roles", openings = 2
    Check "native.twoOpenings.excluded", Not query.ContainsPoint(region, 225#, 85#) And Not query.ContainsPoint(region, 270#, 110#)
    Check "native.twoOpenings.material", query.ContainsPoint(region, 250#, 100#)
    Check "native.twoOpenings.source", CStr(source.Handle) = handleBefore And doc.ModelSpace.Count = countBefore
    CheckArea "native.twoOpenings.sourceArea", CDbl(source.Area), areaBefore
    CheckNativeSeparateOpenings doc, reader, query
    CheckNativeSavedContours doc, reader
    CheckNativeSharedContourOwnership doc
    CheckNativeRegionBoundaryDedup doc
    CheckNativeMixedContourRoundTrip doc
    path = ThisWorkbook.Path & "\SP35_ContourImport_" & Format$(Now, "yyyymmdd_hhnnss") & ".dwg"
    If Len(Dir$(path)) > 0 Then Err.Raise vbObjectError + 5502, , "Тестовый DWG уже существует; перезапись запрещена."
    doc.SaveAs path
    mReport = mReport & "NATIVE_DWG: " & path & vbCrLf
    GoTo Cleanup
Failed:
    Check "native.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Cleanup:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    If Not previous Is Nothing Then previous.Activate
    On Error GoTo 0
    mReport = mReport & "TOTAL_REAL_AUTOCAD_CONTOURS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
    RunRealAutoCADContourTests = mReport
End Function

' ДЛЯ ТЕСТОВ: проходит production import -> Results -> export -> import
' для обоих смешанных источников и для однородных вариантов. Два отверстия
' остаются на своем слое; экспорт добавляет новые объекты и сохраняет
' исходники. OUTPUT в метрах не меняет размеры DWG в миллиметрах.
Private Sub CheckNativeMixedContourRoundTrip(ByVal doc As Object)
    Dim importer As CAutoCADSectionModelImporter, writer As CNDMResultsWriter, units As CUnitSystem
    Dim query As CSectionGeometryQuery, section As CSectionModel, region As CConcreteRegion
    Dim outer As Object, mesh As Object, bar As Object, hole As Object, curve As Object, entity As Object
    Dim curves(0 To 0) As Object, regions As Variant, center(0 To 2) As Double
    Dim mode As Long, i As Long, countBefore As Long, exported As Long, expectedExported As Long, outerCount As Long, openingCount As Long
    Dim prefix As String, outerLayer As String, openingLayer As String, meshLayer As String, rebarLayer As String, outerHandle As String
    Dim unitRange As Object, savedUnits As Variant, settings As CSystemSettingsReader, number As Long, description As String
    On Error GoTo Failed
    Set unitRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange: savedUnits = unitRange.Formula
    For i = 2 To unitRange.Rows.Count
        If CStr(unitRange.Cells(i, 1).Value2) = "Length" Then unitRange.Cells(i, 4).Value2 = "m"
        If CStr(unitRange.Cells(i, 1).Value2) = "Area" Then unitRange.Cells(i, 4).Value2 = "m2"
    Next i
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set importer = New CAutoCADSectionModelImporter: Set writer = New CNDMResultsWriter
    Set query = New CSectionGeometryQuery: Set units = New CUnitSystem: units.LoadFromSettings settings
    For mode = 0 To 3
        prefix = "native.mixedRoundTrip." & CStr(mode)
        outerLayer = "NDM_MIX_OUTER_" & CStr(mode): openingLayer = "NDM_MIX_OPENING_" & CStr(mode)
        meshLayer = "NDM_MIX_MESH_" & CStr(mode): rebarLayer = "NDM_MIX_REBAR_" & CStr(mode)
        doc.Layers.Add outerLayer: doc.Layers.Add openingLayer: doc.Layers.Add meshLayer: doc.Layers.Add rebarLayer
        Set outer = NativeRectangle(doc, 100#, 100#, 0#, 1100# + 150# * mode, 100#): outer.Layer = outerLayer
        Set curves(0) = outer: regions = doc.ModelSpace.AddRegion(curves): Set mesh = regions(LBound(regions)): mesh.Layer = meshLayer
        If mode = 1 Or mode = 3 Then MarkNDMContourOutput outer
        outerHandle = CStr(outer.Handle)
        For i = 0 To 1
            Set hole = NativeRectangle(doc, 10# + 10# * i, 10#, 0#, 1080# + 150# * mode + 35# * i, 85# + 30# * i)
            hole.Layer = openingLayer
            If mode = 0 Or mode = 3 Then MarkNDMContourOutput hole
        Next i
        center(0) = 1060# + 150# * mode: center(1) = 60#
        Set curve = doc.ModelSpace.AddCircle(center, 5#)
        Set curves(0) = curve: regions = doc.ModelSpace.AddRegion(curves): Set bar = regions(LBound(regions))
        bar.Layer = rebarLayer: curve.Delete
        countBefore = doc.ModelSpace.Count
        Set section = importer.ImportFromModelSpace(doc.ModelSpace, meshLayer, rebarLayer, "Rebar", 0#, Nothing, outerLayer, openingLayer, doc.Layers)
        query.Initialize section: Set region = query.ConcreteDomain
        CheckArea prefix & ".importArea", region.Area, 9700#
        Check prefix & ".importThreeLoops", region.LoopCount = 3 And doc.ModelSpace.Count = countBefore
        writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section), units
        expectedExported = 3
        exported = SP35ExportSavedMaterialContoursForTests(doc, outerLayer, openingLayer)
        Check prefix & ".exportAddsAll", exported = expectedExported And doc.ModelSpace.Count = countBefore + 3
        Check prefix & ".originalOuterKept", CStr(doc.HandleToObject(outerHandle).Handle) = outerHandle
        outerCount = 0: openingCount = 0
        For Each entity In doc.ModelSpace
            If CStr(entity.Layer) = outerLayer Then outerCount = outerCount + 1
            If CStr(entity.Layer) = openingLayer Then openingCount = openingCount + 1
        Next entity
        Check prefix & ".exportBothLayers", outerCount = 2 And openingCount = 4
        ' Тест сам подготавливает чистый импорт, как теперь обязан пользователь.
        For i = doc.ModelSpace.Count - 1 To countBefore Step -1: doc.ModelSpace.Item(i).Delete: Next i
        Set section = importer.ImportFromModelSpace(doc.ModelSpace, meshLayer, rebarLayer, "Rebar", 0#, Nothing, outerLayer, openingLayer, doc.Layers)
        query.Initialize section: Set region = query.ConcreteDomain
        CheckArea prefix & ".reimportArea", region.Area, 9700#
        Check prefix & ".reimportThreeLoops", region.LoopCount = 3
        exported = SP35ExportSavedMaterialContoursForTests(doc, outerLayer, openingLayer)
        Check prefix & ".repeatExportAddsAll", exported = expectedExported And doc.ModelSpace.Count = countBefore + 3
    Next mode
    GoTo Restore
Failed:
    number = Err.Number: description = Err.Description
Restore:
    On Error Resume Next
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "CheckNativeMixedContourRoundTrip", description
End Sub

' ДЛЯ ТЕСТОВ: отверстие исходного Region уже является достоверной границей.
' Экспорт добавляет его полилинию независимо от существующего исходника.
Private Sub CheckNativeRegionBoundaryDedup(ByVal doc As Object)
    Dim polyline As Object, source As Object, circleEntity As Object, hole As Object
    Dim curves(0 To 0) As Object, regions As Variant, center(0 To 2) As Double
    Dim reader As CAutoCADContourReader, region As CConcreteRegion, section As CSectionModel
    Dim writer As CNDMResultsWriter, countBefore As Long, exported As Long, handle As String
    doc.Layers.Add "SP35_DEDUP_OUTER": doc.Layers.Add "SP35_DEDUP_OPENING"
    Set polyline = NativeRectangle(doc, 100#, 100#, 0.23, 800#, 100#)
    Set curves(0) = polyline: regions = doc.ModelSpace.AddRegion(curves): Set source = regions(LBound(regions))
    source.Layer = "SP35_DEDUP_OUTER": polyline.Delete
    center(0) = 800#: center(1) = 100#
    Set circleEntity = doc.ModelSpace.AddCircle(center, 10#)
    Set curves(0) = circleEntity: regions = doc.ModelSpace.AddRegion(curves): Set hole = regions(LBound(regions))
    source.Boolean 2, hole: circleEntity.Delete
    Set reader = New CAutoCADContourReader: Set region = reader.ReadRegion(source)
    Set section = New CSectionModel: section.SourceType = "Generated"
    section.AddConcreteElement 800#, 100#, CDbl(source.Area), 1, , , "Rectangle", 100#, 100#, 0.23
    section.Contours.AddMaterialRegion region, "SAME_CONTRACT", "Достоверный контур независимо от источника."
    Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
    countBefore = doc.ModelSpace.Count: handle = CStr(source.Handle)
    exported = SP35ExportSavedMaterialContoursForTests(doc, "SP35_DEDUP_OUTER", "SP35_DEDUP_OPENING")
    Check "native.append.regionAndOpeningAdded", exported = 2 And doc.ModelSpace.Count = countBefore + 2
    Check "native.dedup.regionPreserved", doc.HandleToObject(handle).Handle = handle
    CheckArea "native.dedup.regionAreaPreserved", CDbl(source.Area), 10000# - 100# * GEOM_PI
End Sub

' ДЛЯ ТЕСТОВ: настоящая XData защищает пользовательскую полилинию общего
' слоя при отдельной ручной очистке. Экспорт не удаляет прежние объекты.
Private Sub CheckNativeSharedContourOwnership(ByVal doc As Object)
    Dim system As Object, saved As Variant, settings As CSystemSettingsReader, i As Long
    Dim original As Object, copy As Object, number As Long, description As String, handle As String
    Dim writer As CNDMResultsWriter, section As CSectionModel, reader As CAutoCADContourReader, region As CConcreteRegion
    Dim exported As Long, countBefore As Long, ownedCount As Long, entity As Object
    On Error GoTo Failed
    Set system = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange: saved = system.Formula
    For i = 2 To system.Rows.Count
        If CStr(system.Cells(i, 1).Value2) = "AutoCAD.Common.SectionContourLayer" Then system.Cells(i, 2).Value2 = "SP35_SAVED_CONTOURS"
    Next i
    Set original = NativeRectangle(doc, 100#, 100#, 0#, 50#, 50#)
    original.Layer = "SP35_SAVED_CONTOURS": handle = CStr(original.Handle)
    Set copy = original.Copy: MarkNDMContourOutput copy
    Check "native.ownership.originalUnmarked", Not IsNDMContourOutput(original)
    Check "native.ownership.copyMarked", IsNDMContourOutput(copy)
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    exported = Audit03CleanupAutoCADModelSpaceForTests(doc.ModelSpace, settings)
    Check "native.ownership.cleanupOwned", exported >= 1
    Check "native.ownership.sourcePreserved", doc.HandleToObject(handle).Handle = handle
    Set section = New CSectionModel
    section.AddConcreteElement 50#, 50#, 10000#, 1, , , "Rectangle", 100#, 100#, 0#
    section.AddRebarElement 10#, 10#, 10#, GEOM_PI * 25#, "Rebar"
    Set reader = New CAutoCADContourReader: Set region = reader.ReadRegion(original)
    section.Contours.AddMaterialRegion region, "OWNERSHIP", "Исходный контур."
    Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
    countBefore = doc.ModelSpace.Count
    exported = SP35ExportSavedMaterialContoursForTests(doc)
    Check "native.ownership.existingSourceDoesNotBlockExport", exported = 1 And doc.ModelSpace.Count = countBefore + 1
    exported = SP35ExportSavedMaterialContoursForTests(doc)
    Check "native.ownership.repeatAddsContour", exported = 1 And doc.ModelSpace.Count = countBefore + 2
    For Each entity In doc.ModelSpace
        If CStr(entity.Layer) = "SP35_SAVED_CONTOURS" And IsNDMContourOutput(entity) Then ownedCount = ownedCount + 1
    Next entity
    Check "native.ownership.previousOutputKept", ownedCount = 2
    Check "native.ownership.originalStillExists", doc.HandleToObject(handle).Handle = handle
    GoTo Restore
Failed:
    number = Err.Number: description = Err.Description
Restore:
    On Error Resume Next
    If Not system Is Nothing Then system.Formula = saved
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "CheckNativeSharedContourOwnership", description
End Sub

' ДЛЯ ТЕСТОВ: настоящий export строит три отдельные замкнутые полилинии
' из сохраненных outer/opening, в том числе при output в метрах. Сумма
' площадей CAD совпадает с query; текущие Config-единицы не читаются.
Private Sub CheckNativeSavedContours(ByVal doc As Object, ByVal reader As CAutoCADContourReader)
    Dim section As CSectionModel, writer As CNDMResultsWriter, restored As CSectionModel
    Dim units As CUnitSystem, query As CSectionGeometryQuery, region As CConcreteRegion, entity As Object
    Dim edges As Collection, curve As CFakeAcadContour, contourEntity As CFakeAcadContour
    Dim countBefore As Long, exported As Long, i As Long, maxArea As Double, totalArea As Double, value As Double
    Set section = New CSectionModel: section.SourceType = "AutoCADImport"
    section.AddConcreteElement 50#, 50#, 10000#, 1, , , "Rectangle", 100#, 100#, 0#
    section.AddRebarElement 10#, 10#, 10#, GEOM_PI * 25#, "Rebar"
    Set edges = RectangleEdges(0#, 0#, 100#, 100#)
    Set curve = New CFakeAcadContour: curve.InitializeCircle 35#, 35#, 10#: edges.Add curve
    Set curve = New CFakeAcadContour: curve.InitializeCircle 70#, 65#, 5#: edges.Add curve
    Set contourEntity = Contour("AcDbRegion", edges): Set region = reader.ReadRegion(contourEntity)
    section.Contours.AddMaterialRegion region, "NATIVE_SNAPSHOT", "Точный контур с двумя проемами."
    Set units = New CUnitSystem: units.InitializeDefaults
    Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section), units
    Set restored = ReadSectionGeometryFromResults(ThisWorkbook, "AutoCADImport")
    Set query = New CSectionGeometryQuery: query.Initialize restored
    CheckArea "native.snapshot.area", query.ConcreteDomain.Area, region.Area
    countBefore = doc.ModelSpace.Count
    exported = SP35ExportSavedMaterialContoursForTests(doc)
    Check "native.snapshot.exportThreeLoops", exported = 3 And doc.ModelSpace.Count = countBefore + 3
    For i = countBefore To doc.ModelSpace.Count - 1
        Set entity = doc.ModelSpace.Item(i)
        Check "native.snapshot.closed." & CStr(i - countBefore), CBool(entity.Closed)
        value = CDbl(entity.Area): totalArea = totalArea + value
        If value > maxArea Then maxArea = value
        Set region = reader.ReadRegion(entity)
        CheckArea "native.snapshot.exactArc." & CStr(i - countBefore), region.Area, value
    Next i
    CheckArea "native.snapshot.outerMM", maxArea, 10000#
    CheckArea "native.snapshot.openingsMM", totalArea - maxArea, 125# * GEOM_PI
    CheckArea "native.snapshot.netArea", 2# * maxArea - totalArea, query.ConcreteDomain.Area
End Sub

' ДЛЯ ТЕСТОВ: два opening на одном реальном слое импортируются и сохраняются
' вместе. Бетонный Region сетки продолжает использовать прежнюю эквивалентную
' оболочку; достоверный контур не меняет его фактическую площадь и инерции.
Private Sub CheckNativeSeparateOpenings(ByVal doc As Object, ByVal reader As CAutoCADContourReader, ByVal query As CSectionGeometryQuery)
    Dim outer As Object, mesh As Object, bar As Object, hole As Object, holeCopy As Object, polyline As Object
    Dim curves(0 To 0) As Object, regions As Variant, center(0 To 2) As Double
    Dim importer As CAutoCADSectionModelImporter, section As CSectionModel, region As CConcreteRegion
    Dim i As Long, width As Double, height As Double, angle As Double, ix As Double, iy As Double, ixy As Double
    Dim raw As Variant, expectedIx As Double, expectedIy As Double, countBefore As Long
    doc.Layers.Add "SP35_OUTER": doc.Layers.Add "SP35_OPENING"
    doc.Layers.Add "SP35_MESH": doc.Layers.Add "SP35_REBAR"
    Set outer = NativeRectangle(doc, 100#, 100#, 0#, 500#, 100#): outer.Layer = "SP35_OUTER"
    Set curves(0) = outer: regions = doc.ModelSpace.AddRegion(curves): Set mesh = regions(LBound(regions))
    mesh.Layer = "SP35_MESH"
    For i = 0 To 1
        center(0) = 475# + 45# * i: center(1) = 85# + 25# * i
        Set polyline = doc.ModelSpace.AddCircle(center, 10# - 5# * i)
        Set curves(0) = polyline: regions = doc.ModelSpace.AddRegion(curves): Set hole = regions(LBound(regions))
        hole.Layer = "SP35_OPENING"
        Set holeCopy = hole.Copy: mesh.Boolean 2, holeCopy
    Next i
    center(0) = 460#: center(1) = 60#
    Set polyline = doc.ModelSpace.AddCircle(center, 5#)
    Set curves(0) = polyline: regions = doc.ModelSpace.AddRegion(curves): Set bar = regions(LBound(regions))
    bar.Layer = "SP35_REBAR"
    raw = mesh.MomentOfInertia
    Dim centroid As Variant
    centroid = mesh.Centroid
    expectedIx = CDbl(raw(LBound(raw))) - CDbl(mesh.Area) * CDbl(centroid(LBound(centroid) + 1)) ^ 2
    expectedIy = CDbl(raw(LBound(raw) + 1)) - CDbl(mesh.Area) * CDbl(centroid(LBound(centroid))) ^ 2
    countBefore = doc.ModelSpace.Count
    Set importer = New CAutoCADSectionModelImporter
    Set section = importer.ImportFromModelSpace(doc.ModelSpace, "SP35_MESH", "SP35_REBAR", "Rebar", 0#, Nothing, "SP35_OUTER", "SP35_OPENING", doc.Layers)
    query.Initialize section: Set region = query.ConcreteDomain
    CheckArea "native.separateOpenings.area", region.Area, 10000# - 125# * GEOM_PI
    Check "native.separateOpenings.threeLoops", region.LoopCount = 3
    Check "native.separateOpenings.excluded", Not query.ContainsPoint(region, 475#, 85#) And Not query.ContainsPoint(region, 520#, 110#)
    Check "native.separateOpenings.sourceKept", doc.ModelSpace.Count = countBefore
    CheckArea "native.mesh.actualAreaKept", section.ConcreteArea(1), CDbl(mesh.Area)
    Check "native.mesh.rectangleKept", section.ConcreteBoundaryRectangle(1, width, height, angle)
    CheckArea "native.mesh.equivalentArea", width * height, CDbl(mesh.Area)
    section.ConcreteLocalInertiaComponents 1, ix, iy, ixy
    Check "native.mesh.actualInertiaKept", Abs(ix - expectedIx) < 0.00001 And Abs(iy - expectedIy) < 0.00001
End Sub

' ДЛЯ ТЕСТОВ: строит исходные вершины независимо от геометрического движка.
Private Function NativeRectangle(ByVal doc As Object, ByVal width As Double, ByVal height As Double, _
        ByVal angle As Double, ByVal cx As Double, ByVal cy As Double) As Object
    Dim coordinates(0 To 7) As Double, px As Variant, py As Variant, i As Long, result As Object
    px = Array(-width / 2#, width / 2#, width / 2#, -width / 2#)
    py = Array(-height / 2#, -height / 2#, height / 2#, height / 2#)
    For i = 0 To 3
        coordinates(2 * i) = cx + CDbl(px(i)) * Cos(angle) - CDbl(py(i)) * Sin(angle)
        coordinates(2 * i + 1) = cy + CDbl(px(i)) * Sin(angle) + CDbl(py(i)) * Cos(angle)
    Next i
    Set result = doc.ModelSpace.AddLightWeightPolyline(coordinates): result.Closed = True
    Set NativeRectangle = result
End Function

' ДЛЯ ТЕСТОВ: ожидает готовность собственного нового CAD-документа с timeout.
Private Sub WaitForCAD(ByVal acad As Object)
    Dim started As Single, elapsed As Double
    started = Timer
    Do
        If acad.GetAcadState().IsQuiescent Then Exit Sub
        DoEvents: Application.Wait DateAdd("s", 1, Now)
        elapsed = Timer - started: If elapsed < 0# Then elapsed = elapsed + 86400#
        If elapsed > 30# Then Err.Raise vbObjectError + 5502, , "AutoCAD не завершил инициализацию за 30 секунд."
    Loop
End Sub
