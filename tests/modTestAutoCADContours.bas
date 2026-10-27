Attribute VB_Name = "modTestAutoCADContours"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: точные контуры AutoCAD Region и Polyline
' ==========================================================================
' Проверяет чтение сегментов, отверстий, направления обхода, ошибки
' и удаление временных объектов общими классами без решения НДС.
' Работа с подставными объектами и с настоящим AutoCAD запускается раздельно.

Private mPassed As Long
Private mFailed As Long
Private mReport As String
Private mNativeProgress As Boolean ' ДЛЯ ТЕСТОВ: журнал собственного CAD-прогона до возврата из COM.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохранённая метка RC_SECTION_NDM не должна влиять на выбор объектов при импорте.
Private Sub TestSetLegacyContourTag(ByVal entity As Object)
    Dim app As Object, codes(0 To 1) As Integer, values(0 To 1) As Variant
    On Error Resume Next
    Set app = entity.Document.RegisteredApplications.Item("RC_SECTION_NDM")
    On Error GoTo 0
    If app Is Nothing Then Set app = entity.Document.RegisteredApplications.Add("RC_SECTION_NDM")
    codes(0) = 1001: codes(1) = 1000
    values(0) = "RC_SECTION_NDM": values(1) = "MATERIAL_CONTOUR_OUTPUT"
    entity.SetXData codes, values
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Экспорт не записывает метки владения: все объекты заданных слоёв остаются равноправными.
Private Function TestHasLegacyContourTag(ByVal entity As Object) As Boolean
    Dim codes As Variant, values As Variant
    entity.GetXData "RC_SECTION_NDM", codes, values
    TestHasLegacyContourTag = IsArray(values)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' запускает parity/negative/import проверки без AutoCAD.
Public Function RunAutoCADContourTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    TestCheckEquivalentSources
    TestCheckNestedRegion
    TestCheckInvalidContours
    TestCheckImportedSnapshot
    TestCheckMultipleOpenings
    TestCheckFiveOpenings
    TestCheckOptionalLayersAndSeparateConcrete
    TestCheckEmptyContourLayers
    TestCheckMixedContourSources
    TestCheckMixedContourTopology
    TestCheckResultsRoundTrip
    TestCheckAdditiveContourExport
    TestCheckContourFormatSettings
    TestCheckImportedContourDimensions
    TestCheckOpeningWithoutOuter
    TestCheckKnownOuterWithoutOpenings
    TestCheckPhysicalPolylineClosure
    TestCheckEmptyStandaloneContourSegments
    TestCheckWorldCoordinateGuard
    TestCheckExactComponentAreaFilter
    TestCheckGeometryOnlyExportSource
    GoTo Finish
Failed:
    TestCheck "contour.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finish:
    mReport = mReport & "TOTAL_AUTOCAD_CONTOURS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
    RunAutoCADContourTests = mReport
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет формат контуров после сохранения и открытия отдельной
' копии книги без нового импорта, расчёта и изменения пользовательских Results.
Public Function RunPostAudit03ContourSnapshotTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    Dim fixture As Object, sheet As Object, writer As CNDMResultsWriter, model As CSectionModel, restored As CSectionModel
    Dim contours As CSectionContours, readBack As CSectionContours, query As CSectionGeometryQuery
    Dim props As CSectionPropertiesCalculator, anchor As Object, unitText As Variant, marker As Object
    Dim solves As Long, path As String, value As Variant, number As Long, description As String, i As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, unitSheet As Object
    Dim geometryAnchor As Object, contourAnchor As Object, geometryData As Variant, geometryRows As Long, geometryColumns As Long
    On Error GoTo Failed
    solves = SectionEquilibriumSolveCount()
    Set fixture = Application.Workbooks.Add(-4167): Set sheet = fixture.Worksheets(1): sheet.Name = "Results"
    TestCreatePostAudit03SnapshotAnchors fixture, sheet
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
    Set geometryAnchor = fixture.Names.Item("rngNDMSectionGeometry").RefersToRange
    Set contourAnchor = fixture.Names.Item("rngNDMSectionContours").RefersToRange
    geometryColumns = writer.OutputColumnCount("rngNDMSectionGeometry"): geometryRows = model.ConcreteCount + model.RebarCount + 1
    geometryData = geometryAnchor.Resize(geometryRows, geometryColumns).Value2
    Set marker = sheet.Cells(geometryAnchor.Row + 2, contourAnchor.Column + writer.OutputColumnCount("rngNDMSectionContours") + 1)
    marker.Value2 = "USER-GAP"
    TestCheck "postAudit03.contours.markerOutside17Columns", writer.OutputColumnCount("rngNDMSectionContours") = 17 And _
        marker.Column > contourAnchor.Column + 16 And marker.Column < fixture.Names.Item("rngNDMSectionProperties").RefersToRange.Column
    TestCheck "postAudit03.contours.geometry16", geometryColumns = 16 And AnchoredColumnCount(geometryAnchor) = 16 And _
        Len(CStr(geometryAnchor.Offset(0, geometryColumns).Value2)) = 0
    TestCheck "postAudit03.contours.geometry16.localIxyHeader", ResultColumn(geometryData, "LocalIxy") = 14
    Set restored = ReadSectionGeometryFromResults(fixture)
    TestCheckArea "postAudit03.contours.geometry16.localIxyRoundTrip", restored.ConcreteLocalIxy(1), 12#
    geometryAnchor.Offset(0, geometryColumns - 1).Resize(geometryRows, 1).ClearContents
    TestCheck "postAudit03.contours.legacy15.width", AnchoredColumnCount(geometryAnchor) = 15
    Set restored = ReadSectionGeometryFromResults(fixture, "Legacy15", True)
    TestCheckArea "postAudit03.contours.legacy15.area", restored.ConcreteArea(1), 10000#
    TestCheckArea "postAudit03.contours.legacy15.localIx", restored.ConcreteLocalIx(1), 123456#
    TestCheckArea "postAudit03.contours.legacy15.localIy", restored.ConcreteLocalIy(1), 654321#
    TestCheckArea "postAudit03.contours.legacy15.localIxy", restored.ConcreteLocalIxy(1), 12#
    geometryAnchor.Resize(geometryRows, geometryColumns).Value2 = geometryData
    TestCheck "postAudit03.contours.geometryOnlyElements", sheet.Cells(23, 14).Value2 = vbNullString
    For Each unitText In Array("mm", "cm", "m")
        unitSheet.Cells(2, 2).Value2 = CStr(unitText)
        Set settings = New CSystemSettingsReader: settings.LoadFromRange unitSheet.Range("A1:C2")
        Set units = New CUnitSystem: units.LoadFromSettings settings
        writer.WriteSectionContours fixture, contours, sheet.Cells(21, 12).Value2, units
        Set readBack = ReadSavedSectionContours(fixture)
        TestCheckPostAudit03ContoursEqual "postAudit03.contours." & CStr(unitText), contours, readBack
        Set restored = ReadSectionGeometryFromResults(fixture)
        TestCheckArea "postAudit03.contours.mechanicalA." & CStr(unitText), restored.ConcreteArea(1), 10000#
        TestCheckArea "postAudit03.contours.mechanicalIx." & CStr(unitText), restored.ConcreteLocalIx(1), 123456#
        TestCheckArea "postAudit03.contours.mechanicalIxy." & CStr(unitText), restored.ConcreteLocalIxy(1), 12#
        Set query = New CSectionGeometryQuery: query.Initialize restored
        TestCheckArea "postAudit03.contours.domain." & CStr(unitText), query.ConcreteDomain.Area, 12000# * GEOM_PI + 10000#
        TestCheck "postAudit03.contours.marker." & CStr(unitText), marker.Value2 = "USER-GAP"
    Next unitText
    path = ThisWorkbook.Path & "\PostAudit03_Contour_RoundTrip.xlsm"
    Application.DisplayAlerts = False: fixture.SaveAs path, 52: fixture.Close False
    Set fixture = Application.Workbooks.Open(path): Set sheet = fixture.Worksheets("Results")
    Set geometryAnchor = fixture.Names.Item("rngNDMSectionGeometry").RefersToRange
    Set readBack = ReadSavedSectionContours(fixture)
    TestCheckPostAudit03ContoursEqual "postAudit03.contours.reopen", contours, readBack
    Set anchor = fixture.Names.Item("rngNDMSectionContours").RefersToRange
    anchor.Offset(-1, 0).Resize(contours.Count + 2, 17).Cut sheet.Cells(19, 200)
    fixture.Names.Item("rngNDMSectionContours").RefersTo = "=Results!$GR$20"
    Set anchor = fixture.Names.Item("rngNDMSectionContours").RefersToRange
    Set readBack = ReadSavedSectionContours(fixture)
    TestCheckPostAudit03ContoursEqual "postAudit03.contours.moved", contours, readBack
    TestCheck "postAudit03.contours.canonicalHeader", anchor.Value2 = "RunID"
    anchor.Value2 = "RunID v1"
    On Error Resume Next
    Set readBack = ReadSavedSectionContours(fixture): number = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "postAudit03.contours.obsoleteHeaderRejected", number = vbObjectError + 4376
    anchor.Value2 = "RunID"
    value = anchor.Offset(1, 6).Value2: anchor.Offset(1, 6).Value2 = "bad"
    On Error Resume Next
    Set readBack = ReadSavedSectionContours(fixture): number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo Failed
    anchor.Offset(1, 6).Value2 = value
    TestCheck "postAudit03.contours.dynamicErrorAddress", number <> 0 And InStr(description, anchor.Offset(1, 6).Address(False, False)) > 0
    value = anchor.Value2: anchor.ClearContents
    On Error Resume Next
    Set readBack = ReadSavedSectionContours(fixture): number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo Failed
    anchor.Value2 = value
    TestCheck "postAudit03.contours.blankHeaderRejected", number = vbObjectError + 4376
    fixture.Names.Item("rngNDMSectionContours").RefersTo = "=#REF!"
    On Error Resume Next
    Set readBack = ReadSavedSectionContours(fixture): number = Err.Number: Err.Clear
    On Error GoTo Failed
    fixture.Names.Item("rngNDMSectionContours").RefersTo = "=Results!$GR$20"
    TestCheck "postAudit03.contours.brokenNameRejected", number = vbObjectError + 4376 And _
        fixture.Names.Item("rngNDMSectionProperties").RefersToRange.Column = 49
    Set readBack = New CSectionContours
    readBack.AddContourCircle "ONLY_HOLE", 5#, 5#, 2#, "Только отверстие", "ONLY_HOLE", "Opening"
    writer.WriteSectionContours fixture, readBack, sheet.Cells(21, 12).Value2
    Set readBack = ReadSavedSectionContours(fixture)
    TestCheck "postAudit03.contours.onlyOpening", readBack.Count = 1 And readBack.LoopRole(1) = "Opening"
    TestCheck "postAudit03.contours.largeSmallCleared", Len(CStr(anchor.Offset(2, 0).Value2)) = 0 And Len(CStr(anchor.Offset(10, 15).Value2)) = 0
    readBack.Clear: writer.WriteSectionContours fixture, readBack, "EMPTY"
    Set readBack = ReadSavedSectionContours(fixture)
    TestCheck "postAudit03.contours.emptyAvailable", readBack.Count = 0 And anchor.Value2 = "RunID"
    fixture.Names.Item("rngNDMSectionContours").Delete
    On Error Resume Next
    Set readBack = ReadSavedSectionContours(fixture): number = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "postAudit03.contours.missingTableRejected", number = vbObjectError + 4376
    TestCheck "postAudit03.contours.missingTableDoesNotInsert", fixture.Names.Count = 6 And _
        fixture.Names.Item("rngNDMSectionProperties").RefersToRange.Column = 49
    fixture.Names.Add "rngNDMSectionContours", "=Results!$GR$20"
    Set readBack = New CSectionContours: writer.WriteSectionContours fixture, readBack, Empty
    geometryAnchor.Offset(0, geometryColumns).Value2 = "UnexpectedGeometryField"
    TestCheck "postAudit03.contours.expandedGeometryIs17", AnchoredColumnCount(geometryAnchor) = 17
    On Error Resume Next
    Set restored = ReadSectionGeometryFromResults(fixture): number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo Failed
    TestCheck "postAudit03.contours.expandedGeometryRejected", number = vbObjectError + 4355
    TestCheck "postAudit03.contours.expandedGeometryAddress", InStr(description, geometryAnchor.Resize(1, 17).Address(False, False)) > 0
    geometryAnchor.Offset(0, geometryColumns).ClearContents
    Set restored = ReadSectionGeometryFromResults(fixture)
    TestCheckArea "postAudit03.contours.geometry16.restoredIxy", restored.ConcreteLocalIxy(1), 12#
    fixture.Close False: Set fixture = Nothing
    TestCheck "postAudit03.contours.noSolve", SectionEquilibriumSolveCount() = solves
    GoTo Finished
Failed:
    TestCheck "postAudit03.contours.runtime." & CStr(Err.Number) & "." & Err.Description, False
Finished:
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
    RunPostAudit03ContourSnapshotTests = mReport & "TOTAL_POSTAUDIT03_CONTOURS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Несвязный CAD Region фильтруется после разложения на конечные компоненты.
' Ровно пороговая площадь сохраняется; отверстие и дуги оставшейся части не меняются.
Private Sub TestCheckExactComponentAreaFilter()
    Dim fixture As Object: Set fixture = Application.Workbooks.Add(-4167)
    Dim number As Long, description As String
    On Error GoTo Failed
    Dim table(1 To 2, 1 To 3) As Variant
    table(1, 1) = "Key": table(1, 2) = "Value": table(1, 3) = "Unit"
    table(2, 1) = "Mesh.CellGeometry": table(2, 2) = "Contour": table(2, 3) = "-"
    fixture.Worksheets(1).Range("A1:C2").Value2 = table
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader
    settings.LoadFromRange fixture.Worksheets(1).Range("A1:C2")
    Dim query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    Dim ring As CConcreteRegion: Set ring = query.SubtractContainedRegion(query.CircleRegion(0#, 0#, 10#), query.CircleRegion(0#, 0#, 3#))
    Dim parts As Collection: Set parts = New Collection
    parts.Add ring: parts.Add query.CircleRegion(40#, 0#, 2#)
    parts.Add query.RectangleRegion(80#, 0#, 4#, 5#)
    Dim domain As CConcreteRegion: Set domain = query.MergeDisjointRegions(parts)
    Dim properties As TGeometryProperties: query.RegionProperties domain, properties
    Dim concrete As CTestFakeAcadRegion: Set concrete = New CTestFakeAcadRegion
    concrete.TestInitialize properties.Area, properties.CentroidX, properties.CentroidY, properties.Ixc, properties.Iyc, properties.Ixyc, "CONCRETE", "MULTIPART"
    concrete.TestSetGeometry domain
    Dim originalSourceArea As Double: originalSourceArea = concrete.Area
    Dim bar As CTestFakeAcadRegion, space As Collection: Set space = New Collection
    space.Add concrete
    Set bar = New CTestFakeAcadRegion: bar.TestInitialize 20#, 0#, -6#, 1#, 1#, 0#, "REBAR", "AT_LIMIT": space.Add bar
    Set bar = New CTestFakeAcadRegion: bar.TestInitialize 19#, 1#, -6#, 1#, 1#, 0#, "REBAR", "BELOW_LIMIT": space.Add bar
    Dim importer As CAutoCADSectionModelImporter: Set importer = New CAutoCADSectionModelImporter
    Dim model As CSectionModel, i As Long, foundRing As Boolean, foundRectangle As Boolean
    Set model = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 20#, meshSettings:=settings)
    TestCheck "exactComponentFilter.concreteCount", model.ConcreteCount = 2
    TestCheck "exactComponentFilter.rebarCount", model.RebarCount = 1 And model.RebarSourceHandle(1) = "AT_LIMIT"
    For i = 1 To model.ConcreteCount
        TestCheck "exactComponentFilter.atomicBoundary." & CStr(i), model.ConcreteHasBoundary(i)
        TestCheck "exactComponentFilter.ids." & CStr(i), model.ConcreteID(i) = "C" & CStr(i)
        TestCheck "exactComponentFilter.area." & CStr(i), model.ConcreteArea(i) >= 20#
        If model.ConcreteGeometry(i).LoopCount = 2 Then
            foundRing = query.RegionsCoincide(model.ConcreteGeometry(i), ring)
            TestCheck "exactComponentFilter.ringHole", Not query.ContainsPoint(model.ConcreteGeometry(i), 0#, 0#)
            query.RegionProperties ring, properties
            TestCheckArea "exactComponentFilter.ringA", model.ConcreteArea(i), properties.Area
            TestCheckArea "exactComponentFilter.ringIx", model.ConcreteLocalIx(i), properties.Ixc
            TestCheckArea "exactComponentFilter.ringIy", model.ConcreteLocalIy(i), properties.Iyc
        ElseIf Abs(model.ConcreteArea(i) - 20#) < 0.00000001 Then
            foundRectangle = query.RegionsCoincide(model.ConcreteGeometry(i), query.RectangleRegion(80#, 0#, 4#, 5#))
        End If
    Next i
    TestCheck "exactComponentFilter.ringPreserved", foundRing
    TestCheck "exactComponentFilter.equalThresholdPreserved", foundRectangle
    TestCheck "exactComponentFilter.noOrientationExplode", importer.TestDebugEdgeProbeCount = 0
    Set model = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, meshSettings:=settings)
    TestCheck "exactComponentFilter.zeroDisablesFilter", model.ConcreteCount = 3 And model.RebarCount = 2
    TestCheck "exactComponentFilter.sourceUnchanged", Not concrete.TestDeleted And Not bar.TestDeleted And _
        Not space.Item(2).TestDeleted And concrete.Area = originalSourceArea ' Строго сравнивается тот же CAD-источник до/после; domain.Area вычислена другим путём и не является исходным COM Area.
    fixture.Close False
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    On Error Resume Next: fixture.Close False: On Error GoTo 0
    Err.Raise number, "TestCheckExactComponentAreaFilter", description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Preview обоих представлений разрешает геометрический экспорт с обязательным статусом.
' Записи LC даже без element-results запрещают такой fallback; сохранённые таблицы не меняются.
Private Sub TestCheckGeometryOnlyExportSource()
    Dim query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    Dim model As CSectionModel, writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
    Dim mode As Long, number As Long, description As String, status As String
    Dim geometry As CConcreteRegion: Set geometry = query.RectangleRegion(0#, 0#, 100#, 80#)
    Dim properties As TGeometryProperties: query.RegionProperties geometry, properties
    Dim solves As Long: solves = SectionEquilibriumSolveCount()
    Dim names As Variant: names = Array("rngNDMSectionGeometry", "rngNDMElementGeometry", "rngNDMSectionContours", "rngNDMSectionProperties", "rngNDMElementResults")
    Dim savedTables As Object, key As Variant, actual As Variant
    For mode = 0 To 1
        Set model = New CSectionModel: model.SourceType = "Generated"
        If mode = 0 Then
            model.AddConcreteElement 0#, 0#, 8000#, 1, , , "Rectangle", 100#, 80#
        Else
            model.AddConcreteElement 0#, 0#, 8000#, 1, , , "Contour", , , , , properties.Ixc, properties.Iyc, properties.Ixyc, True, geometry
        End If
        model.AddRebarElement 0#, 20#, 10#, 0#, "Rebar"
        model.Contours.AddMaterialRegion geometry, "OUTER", "Наружная граница тестового сечения."
        writer.WriteGeometryPreview ThisWorkbook, model, PrepareSectionSnapshot(model)
        Set savedTables = CreateObject("Scripting.Dictionary")
        For Each key In names: savedTables.Add CStr(key), ReadPhysicalResultTable(ThisWorkbook, CStr(key)): Next key
        status = Audit02ReadExportSnapshotForTests(ThisWorkbook)
        TestCheck "geometryOnly.mode." & CStr(mode), Left$(status, 13) = "GeometryOnly|"
        TestCheck "geometryOnly.warning." & CStr(mode), InStr(status, "В AutoCAD передано только сечение, без рассчитанного состояния.") > 0
        TestCheck "geometryOnly.counts." & CStr(mode), Right$(status, 4) = "|1|1"
        For Each key In names
            actual = ReadPhysicalResultTable(ThisWorkbook, CStr(key))
            TestCheck "geometryOnly.sourceUntouched." & CStr(mode) & "." & CStr(key), TestGeometryExportTablesEqual(savedTables.Item(CStr(key)), actual)
        Next key
        TestCheck "geometryOnly.contextReleased." & CStr(mode), TestCalculationSnapshotReadReleased()
        TestCheck "geometryOnly.noSolve." & CStr(mode), SectionEquilibriumSolveCount() = solves
        Dim anchor As Object: Set anchor = ThisWorkbook.Names.Item("rngNDMSectionProperties").RefersToRange
        actual = savedTables.Item("rngNDMSectionProperties")
        anchor.Offset(UBound(actual, 1), 0).Value2 = actual(2, 1)
        anchor.Offset(UBound(actual, 1), 1).Value2 = "BROKEN_LC"
        anchor.Offset(UBound(actual, 1), 2).Value2 = "StrengthState.Status"
        anchor.Offset(UBound(actual, 1), 3).Value2 = "InvalidSavedState"
        number = 0: description = vbNullString: status = vbNullString
        On Error Resume Next
        status = Audit02ReadExportSnapshotForTests(ThisWorkbook)
        number = Err.Number: description = Err.Description: Err.Clear
        On Error GoTo 0
        TestCheck "geometryOnly.calculatedFailureNotHidden." & CStr(mode), number <> 0 And InStr(status, "GeometryOnly|") = 0
        TestCheck "geometryOnly.failureContextReleased." & CStr(mode), TestCalculationSnapshotReadReleased()
        writer.WriteGeometryPreview ThisWorkbook, model, PrepareSectionSnapshot(model)
    Next mode
    Dim data As Variant: data = ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionGeometry")
    ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Offset(1, 0).Resize(UBound(data, 1) - 1, UBound(data, 2)).ClearContents
    On Error Resume Next
    status = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo 0
    TestCheck "geometryOnly.notFormed.action", number <> 0 And InStr(description, "Сформировать сечение") > 0
    TestCheck "geometryOnly.notFormed.contextReleased", TestCalculationSnapshotReadReleased()
    writer.WriteGeometryPreview ThisWorkbook, model, PrepareSectionSnapshot(model)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает физические Variant-матрицы с сохранением типов и числовых значений.
' Проверка предназначена для малых независимых fixture, не для рабочего вывода Excel.
Private Function TestGeometryExportTablesEqual(ByVal expected As Variant, ByVal actual As Variant) As Boolean
    On Error GoTo Different
    If UBound(expected, 1) <> UBound(actual, 1) Or UBound(expected, 2) <> UBound(actual, 2) Then Exit Function
    Dim row As Long, column As Long
    For row = 1 To UBound(expected, 1)
        For column = 1 To UBound(expected, 2)
            If VarType(expected(row, column)) <> VarType(actual(row, column)) Then Exit Function
            If CStr(expected(row, column)) <> CStr(actual(row, column)) Then Exit Function
        Next column
    Next row
    TestGeometryExportTablesEqual = True
Different:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создает только тестовые якоря нижнего ряда; ширины не переназначаются.
Private Sub TestCreatePostAudit03SnapshotAnchors(ByVal workbook As Object, ByVal sheet As Object)
    Dim names As Variant, columns As Variant, i As Long
    names = Array("rngNDMElementResults", "rngNDMSectionGeometry", "rngNDMSectionProperties", "rngNDMMaterialDiagrams", "rngNDMSectionAnnotations", "rngNDMElementGeometry")
    columns = Array(1, 12, 49, 57, 70, 96)
    For i = 0 To UBound(names)
        workbook.Names.Add CStr(names(i)), "='" & sheet.Name & "'!" & sheet.Cells(20, CLng(columns(i))).Address
    Next i
    workbook.Names.Add "rngNDMSectionContours", "='" & sheet.Name & "'!$AD$20"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Поля проверяются по собственным ID и аналитической геометрии, не по площади одной фигуры.
Private Sub TestCheckPostAudit03ContoursEqual(ByVal prefix As String, ByVal expected As CSectionContours, ByVal actual As CSectionContours)
    TestCheck prefix & ".count", expected.Count = actual.Count
    If expected.Count <> actual.Count Then Exit Sub
    Dim i As Long
    For i = 1 To expected.Count
        TestCheck prefix & ".id." & CStr(i), expected.ContourID(i) = actual.ContourID(i) And expected.LoopID(i) = actual.LoopID(i)
        TestCheck prefix & ".context." & CStr(i), expected.LoopRole(i) = actual.LoopRole(i) And expected.SourceID(i) = actual.SourceID(i) And expected.Comment(i) = actual.Comment(i)
        TestCheck prefix & ".type." & CStr(i), expected.SegmentType(i) = actual.SegmentType(i)
        TestCheck prefix & ".geometry." & CStr(i), Abs(expected.StartX(i) - actual.StartX(i)) < 0.00000001 And Abs(expected.StartY(i) - actual.StartY(i)) < 0.00000001 And _
            Abs(expected.EndX(i) - actual.EndX(i)) < 0.00000001 And Abs(expected.EndY(i) - actual.EndY(i)) < 0.00000001 And _
            Abs(expected.Radius(i) - actual.Radius(i)) < 0.00000001 And Abs(expected.SweepAngle(i) - actual.SweepAngle(i)) < 0.000000000001
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отдельный opening допустим при уже пустой или приближенной
' сеточной границе. Точные наружные контуры здесь намеренно отсутствуют.
Private Sub TestCheckOpeningWithoutOuter()
    Dim space As Collection, layers As Collection, importer As CAutoCADSectionModelImporter
    Dim cell As CTestFakeAcadRegion, layer As CTestFakeAcadContour, opening As CTestFakeAcadContour
    Dim section As CSectionModel, query As CSectionGeometryQuery, region As CConcreteRegion
    Dim mode As Long, i As Long, expected As Double, writer As CNDMResultsWriter, restored As CSectionModel
    Set importer = New CAutoCADSectionModelImporter: Set query = New CSectionGeometryQuery
    For mode = 0 To 3
        Set space = New Collection: Set layers = New Collection
        Set cell = New CTestFakeAcadRegion
        cell.TestInitialize 2000#, 50#, 10#, 66666.6666666667, 1666666.66666667, 0#, "CONCRETE", "BOTTOM", 0#, True: space.Add cell
        Set cell = New CTestFakeAcadRegion
        cell.TestInitialize 2000#, 50#, 90#, 66666.6666666667, 1666666.66666667, 0#, "CONCRETE", "TOP", 0#, True: space.Add cell
        Set cell = New CTestFakeAcadRegion
        cell.TestInitialize 1200#, 10#, 50#, 360000#, 40000#, 0#, "CONCRETE", "LEFT", 0#, True: space.Add cell
        Set cell = New CTestFakeAcadRegion
        cell.TestInitialize 1200#, 90#, 50#, 360000#, 40000#, 0#, "CONCRETE", "RIGHT", 0#, True: space.Add cell
        Set cell = New CTestFakeAcadRegion
        cell.TestInitialize GEOM_PI * 4#, 10#, 10#, 10#, 10#, 0#, "REBAR", "BAR": space.Add cell
        Set layer = New CTestFakeAcadContour: layers.Add layer, "OPENING"
        If mode = 0 Or mode = 3 Then
            Set opening = TestContour("AcDbPolyline", TestRectangleEdges(20#, 20#, 60#, 60#), "OPENING"): expected = 6400#
        ElseIf mode = 1 Then
            Set opening = TestContour("AcDbRegion", TestRectangleEdges(15#, 15#, 70#, 70#), "OPENING"): expected = 5100#
        Else
            Set opening = TestContour("AcDbPolyline", TestRectangleEdges(25#, 25#, 50#, 50#), "OPENING"): expected = 7500#
            Set cell = New CTestFakeAcadRegion
            cell.TestInitialize GEOM_PI * 4#, 22#, 50#, 10#, 10#, 0#, "REBAR", "BAR_IN_MESH_VOID": space.Add cell
        End If
        space.Add opening
        If mode = 3 Then
            Set opening = TestContour("AcDbPolyline", TestRectangleEdges(85#, 40#, 10#, 10#), "OPENING")
            space.Add opening: expected = expected - 100#
        End If
        Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
        query.Initialize section: Set region = query.ConcreteDomain
        TestCheckArea "openingOnly.meshVoid." & CStr(mode), region.Area, expected
        TestCheck "openingOnly.noInventedOuter." & CStr(mode), InStr(region.Source, "AuthoritativeContour") = 0
        TestCheck "openingOnly.authoritativeHole." & CStr(mode), region.Source = "ElementBoundaryLoops+KnownOpening"
        For i = 1 To section.Contours.Count
            TestCheck "openingOnly.explicitRole." & CStr(mode) & "." & CStr(i), section.Contours.LoopRole(i) = "Opening"
        Next i
        TestCheck "openingOnly.actualHole." & CStr(mode), Not query.ContainsPoint(region, 50#, 50#)
        If mode = 2 Then
            TestCheck "openingOnly.noMeshOpening", query.ContainsPoint(region, 22#, 50#)
            TestCheck "openingOnly.rebarValidatedByFinalDomain", section.RebarCount = 2
            TestCheckArea "openingOnly.clippedExactHole", query.ClipHalfPlane(region, 0#, 1#, 50#).Area, 3750#
            Dim intervals As Variant
            intervals = query.LineIntervals(region, 0#, 50#, 1#, 0#)
            TestCheck "openingOnly.exactIntersections", UBound(intervals, 1) = 2 And _
                Abs(intervals(1, 2) - 25#) < 0.00001 And Abs(intervals(2, 1) - 75#) < 0.00001
        End If
        TestCheck "openingOnly.sourceKept." & CStr(mode), Not opening.TestDeleted And section.ConcreteCount = 4
        Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
        Set restored = ReadSectionGeometryFromResults(ThisWorkbook, "AutoCADImport")
        query.Initialize restored
        TestCheckArea "openingOnly.savedRoundTrip." & CStr(mode), query.ConcreteDomain.Area, expected
        TestCheckPostAudit03ContoursEqual "openingOnly.exactSaved." & CStr(mode), section.Contours, restored.Contours
    Next mode
    Set opening = TestContour("AcDbPolyline", TestRectangleEdges(87#, 42#, 5#, 5#), "OPENING"): space.Add opening
    Dim number As Long
    On Error Resume Next
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    number = Err.Number: Err.Clear
    On Error GoTo 0
    TestCheck "openingOnly.overlappingExactOpeningsRejected", number <> 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' заданный наружный контур без отверстий описывает сплошной бетон,
' даже если в сетке есть пустота. Проверяет отсутствующий и пустой слой отверстий,
' а затем сохранение того же смысла геометрии в Results.
Private Sub TestCheckKnownOuterWithoutOpenings()
    Dim space As Collection, layers As Collection, cell As CTestFakeAcadRegion
    Dim layer As CTestFakeAcadContour, outer As CTestFakeAcadContour, mode As Long, openingLayer As String
    Dim section As CSectionModel, restored As CSectionModel, query As CSectionGeometryQuery
    Dim domain As CConcreteRegion, importer As CAutoCADSectionModelImporter, writer As CNDMResultsWriter
    Set importer = New CAutoCADSectionModelImporter: Set query = New CSectionGeometryQuery
    Set writer = New CNDMResultsWriter
    For mode = 0 To 2
        Set space = New Collection: Set layers = New Collection
        Set cell = New CTestFakeAcadRegion
        cell.TestInitialize 2000#, 50#, 10#, 66666.6666666667, 1666666.66666667, 0#, "CONCRETE", "BOTTOM", 0#, True: space.Add cell
        Set cell = New CTestFakeAcadRegion
        cell.TestInitialize 2000#, 50#, 90#, 66666.6666666667, 1666666.66666667, 0#, "CONCRETE", "TOP", 0#, True: space.Add cell
        Set cell = New CTestFakeAcadRegion
        cell.TestInitialize 1200#, 10#, 50#, 360000#, 40000#, 0#, "CONCRETE", "LEFT", 0#, True: space.Add cell
        Set cell = New CTestFakeAcadRegion
        cell.TestInitialize 1200#, 90#, 50#, 360000#, 40000#, 0#, "CONCRETE", "RIGHT", 0#, True: space.Add cell
        Set cell = New CTestFakeAcadRegion
        cell.TestInitialize GEOM_PI * 4#, 10#, 10#, 10#, 10#, 0#, "REBAR", "BAR": space.Add cell
        Set layer = New CTestFakeAcadContour: layers.Add layer, "OUTER"
        openingLayer = "OPENING"
        If mode = 0 Then
            Set layer = New CTestFakeAcadContour: layers.Add layer, "OPENING"
        ElseIf mode = 2 Then
            openingLayer = vbNullString
        End If
        If mode = 1 Then
            Set outer = TestContour("AcDbRegion", TestRectangleEdges(0#, 0#, 100#, 100#), "OUTER")
        Else
            Set outer = TestContour("AcDbPolyline", TestRectangleEdges(0#, 0#, 100#, 100#), "OUTER")
        End If
        If mode = 2 Then outer.TestSetLegacyContourTag True
        space.Add outer
        Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", openingLayer, layers)
        query.Initialize section: Set domain = query.ConcreteDomain
        TestCheckArea "knownOuterOnly." & CStr(mode) & ".area", domain.Area, 10000#
        TestCheck "knownOuterOnly." & CStr(mode) & ".solid", domain.LoopCount = 1 And query.ContainsPoint(domain, 50#, 50#)
        TestCheck "knownOuterOnly." & CStr(mode) & ".source", domain.Source = "AuthoritativeContour"
        TestCheckArea "knownOuterOnly." & CStr(mode) & ".clip", query.ClipHalfPlane(domain, 0#, 1#, 50#).Area, 5000#
        TestCheck "knownOuterOnly." & CStr(mode) & ".mechanicalMeshKept", section.ConcreteCount = 4 And _
            section.ConcreteArea(1) + section.ConcreteArea(2) + section.ConcreteArea(3) + section.ConcreteArea(4) = 6400#
        TestCheck "knownOuterOnly." & CStr(mode) & ".comment", InStr(importer.ImportComment, "отверстия по сетке не добавляются") > 0
        TestCheck "knownOuterOnly." & CStr(mode) & ".shortMessage", InStr(importer.ResultMessage(section), "Наружных контуров: 1; отверстий: 0.") > 0 And _
            InStr(importer.ResultMessage(section), "Предупреждения:") = 0 And InStr(importer.ResultMessage(section), "Координаты") = 0
        TestCheck "knownOuterOnly." & CStr(mode) & ".noTypedWarning", Not importer.HasImportWarnings
        TestCheck "knownOuterOnly." & CStr(mode) & ".sourceKept", Not outer.TestDeleted
        writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
        Set restored = ReadSectionGeometryFromResults(ThisWorkbook, "AutoCADImport")
        query.Initialize restored: Set domain = query.ConcreteDomain
        TestCheckArea "knownOuterOnly." & CStr(mode) & ".snapshotArea", domain.Area, 10000#
        TestCheck "knownOuterOnly." & CStr(mode) & ".snapshotSolid", domain.LoopCount = 1 And query.ContainsPoint(domain, 50#, 50#)
    Next mode
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' настоящий exporter только добавляет выбранные объекты. Даже повторный
' запуск не читает ModelSpace.Count/Item, не удаляет существующие объекты, сохраняет
' разные цвета outer/opening/Ar и реальные дуги в миллиметрах.
Private Sub TestCheckAdditiveContourExport()
    Dim section As CSectionModel, writer As CNDMResultsWriter, space As CTestFakeAcadContourExport
    Dim region As CConcreteRegion, query As CSectionGeometryQuery, entity As CTestFakeAcadContourExport
    Dim points As Variant, exported As Long, i As Long, alive As Long, number As Long
    Set section = New CSectionModel
    section.AddConcreteElement 50#, 50#, 10000#, 1, , , "Rectangle", 100#, 100#
    Set query = New CSectionGeometryQuery
    Set region = query.PolygonRegion(TestContourRectanglePoints(0#, 0#, 100#, 100#))
    section.Contours.AddMaterialRegion region, "OUTER", "Наружный контур теста."
    section.Annotations.AddAnnotation "CRACK_REGION_CIRCLE", "CRACK_REGION_TEST", _
        50#, 50#, 12#, 0#, 0#, 0#, vbNullString, 0#, "mm", "Расчетная область теста."
    Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
    Set space = New CTestFakeAcadContourExport
    exported = SP35ExportSavedContoursToModelSpaceForTests(space)
    TestCheck "export.append.first", exported = 1 And space.TestCreatedEntities.Count = 2
    exported = SP35ExportSavedContoursToModelSpaceForTests(space)
    TestCheck "export.append.repeat", exported = 1 And space.TestCreatedEntities.Count = 4
    For i = 1 To 4
        Set entity = space.TestCreatedEntities(i)
        If entity.TestDeleted Then
            TestCheck "export.temporary.unmarked." & CStr(i), Not TestHasLegacyContourTag(entity)
        Else
            alive = alive + 1
            TestCheck "export.append.region." & CStr(i), entity.ObjectName = "AcDbRegion" And Not TestHasLegacyContourTag(entity)
            TestCheck "export.color.outer." & CStr(i), entity.Layer = "TEST_OUTER" And entity.Color = 30
            TestCheckArea "export.region.area." & CStr(i), entity.Area, 10000#
            points = entity.Coordinates
            TestCheck "export.mm.coordinates." & CStr(i), Abs(CDbl(points(2)) - CDbl(points(0))) = 100#
        End If
    Next i
    TestCheck "export.region.onlyFinalObjects", alive = 2
    exported = SP35ExportSavedContoursToModelSpaceForTests(space, True)
    TestCheck "export.append.crack", exported = 1 And space.TestCreatedEntities.Count = 5
    Set entity = space.TestCreatedEntities(5)
    TestCheck "export.color.crack", entity.Color = 31 And entity.Closed And Not entity.TestDeleted
    TestCheck "export.append.noExistingObjectRead", space.TestForbiddenReads = 0
    ' Boolean — метод настоящего AutoCAD. Подставной объект его не реализует;
    ' вычитание отверстий из Region проверяется отдельно в настоящем чертеже.
    section.Contours.AddContourCircle "CONTOUR_HOLE", 50#, 50#, 10#, "Круглое отверстие теста.", "HOLE", "Opening"
    writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
    Set space = New CTestFakeAcadContourExport
    For i = 1 To 2
        exported = SP35ExportSavedContoursToModelSpaceForTests(space, False, "Polyline")
        TestCheck "export.polyline.append." & CStr(i), exported = 2 And space.TestCreatedEntities.Count = 2 * i
    Next i
    For i = 1 To 4
        Set entity = space.TestCreatedEntities(i)
        TestCheck "export.polyline.type." & CStr(i), entity.ObjectName = "AcDbPolyline" And entity.Closed And _
            Not entity.TestDeleted And Not TestHasLegacyContourTag(entity)
        If i Mod 2 = 1 Then
            TestCheck "export.polyline.outer." & CStr(i), entity.Layer = "TEST_OUTER" And entity.Color = 30
        Else
            TestCheck "export.polyline.opening." & CStr(i), entity.Layer = "TEST_OPENING" And entity.Color = 4 And Abs(entity.TestBulgeAt(0) - Tan(GEOM_PI / 8#)) < 0.00000001
        End If
    Next i
    exported = SP35ExportSavedContoursToModelSpaceForTests(space, True, "Region")
    TestCheck "export.polyline.crackUnchanged", exported = 1 And space.TestCreatedEntities(5).ObjectName = "AcDbPolyline" And space.TestCreatedEntities(5).Color = 31
    TestCheck "export.polyline.noExistingRead", space.TestForbiddenReads = 0
    Set space = New CTestFakeAcadContourExport: space.TestFailRegionAt = 2 ' Ошибка AddRegion не оставляет частичный вывод и не затрагивает существующий.
    points = Array(0#, 0#, 1#, 0#, 1#, 1#, 0#, 1#)
    Set entity = space.AddLightWeightPolyline(points)
    On Error Resume Next
    exported = SP35ExportSavedContoursToModelSpaceForTests(space): number = Err.Number: Err.Clear
    On Error GoTo 0
    TestCheck "export.region.failureRaised", number <> 0
    TestCheck "export.region.previousObjectPreserved", Not entity.TestDeleted
    For i = 2 To space.TestCreatedEntities.Count
        TestCheck "export.region.failedObjectsDeleted." & CStr(i), space.TestCreatedEntities(i).TestDeleted
    Next i
    TestCheck "export.region.failureNoExistingRead", space.TestForbiddenReads = 0
    section.Contours.Clear ' Только известное отверстие остается отдельным голубым Region.
    section.Contours.AddContourCircle "ONLY_HOLE", 50#, 50#, 10#, , "HOLE", "Opening"
    writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
    Set space = New CTestFakeAcadContourExport
    exported = SP35ExportSavedContoursToModelSpaceForTests(space)
    Set entity = space.TestCreatedEntities(2)
    TestCheck "export.region.openingOnly", exported = 1 And entity.ObjectName = "AcDbRegion" And _
        entity.Layer = "TEST_OPENING" And entity.Color = 4 And Not entity.TestDeleted
    TestCheckArea "export.region.openingOnlyArea", entity.Area, GEOM_PI * 100#
    Set space = New CTestFakeAcadContourExport
    exported = SP35ExportSavedContoursToModelSpaceForTests(space, False, "Polyline")
    Set entity = space.TestCreatedEntities(1)
    TestCheck "export.polyline.openingOnly", exported = 1 And entity.ObjectName = "AcDbPolyline" And _
        entity.Layer = "TEST_OPENING" And entity.Color = 4 And Not entity.TestDeleted
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' активный формат читается штатным consumer-ом, а ошибки
' указывают настоящую ячейку даже после переноса именованной таблицы.
Private Sub TestCheckContourFormatSettings()
    Dim table As Object, shifted As Object, saved As Variant, savedShifted As Variant, originalName As String
    Dim enabledRow As Long, formatRow As Long, i As Long, position As Long, value As Variant
    Dim settings As CSystemSettingsReader, result As Object, number As Long, description As String
    On Error GoTo Failed
    Set table = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    saved = table.Formula: originalName = ThisWorkbook.Names.Item("rngSystemSettings").RefersTo
    Set shifted = table.Worksheet.Range("CR1100").Resize(table.Rows.Count, table.Columns.Count)
    savedShifted = shifted.Formula
    For i = 2 To table.Rows.Count
        If CStr(table.Cells(i, 1).Value2) = "AutoCAD.Export.ContourEnabled" Then enabledRow = i
        If CStr(table.Cells(i, 1).Value2) = "AutoCAD.Export.ContourFormat" Then formatRow = i
    Next i
    If enabledRow = 0 Or formatRow = 0 Then Err.Raise vbObjectError + 5505, , "В книге нет настройки формата контуров."
    For position = 0 To 1
        If position = 1 Then
            shifted.Formula = saved
            ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = "='" & shifted.Worksheet.Name & "'!" & shifted.Address
            Set table = shifted
        End If
        table.Cells(enabledRow, 2).Value2 = "Yes"
        For Each value In Array("Region", "Polyline", vbNullString, "invalid")
            table.Cells(formatRow, 2).Value2 = value
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            On Error Resume Next
            Set result = Audit03AutoCADPresentationForTests(settings, "Concrete", "Tension", "C1", 1#, 2)
            number = Err.Number: description = Err.Description: Err.Clear
            On Error GoTo Failed
            If value = "Region" Or value = "Polyline" Then
                TestCheck "export.format.setting." & CStr(position) & "." & CStr(value), number = 0
                If number = 0 Then TestCheck "export.format.consumer." & CStr(position) & "." & CStr(value), result.Item("ContourFormat") = value
            Else
                TestCheck "export.format.invalid." & CStr(position) & "." & CStr(value), number <> 0 And _
                    InStr(description, table.Cells(formatRow, 2).Address(False, False)) > 0 And InStr(description, "ContourFormat") > 0
            End If
        Next value
        table.Cells(enabledRow, 2).Value2 = "No"
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set result = Audit03AutoCADPresentationForTests(settings, "Concrete", "Tension", "C1", 1#, 2)
        TestCheck "export.format.inactive." & CStr(position), Not result.Item("Contour")
    Next position
    GoTo Restore
Failed:
    TestCheck "export.format.runtime; " & Err.Description, False
Restore:
    If Len(originalName) > 0 Then ThisWorkbook.Names.Item("rngSystemSettings").RefersTo = originalName
    If Not IsEmpty(saved) Then ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula = saved
    If Not shifted Is Nothing Then shifted.Formula = savedShifted
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' четыре независимые координаты для явного наружного контура.
Private Function TestContourRectanglePoints(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Variant
    Dim points(1 To 4, 1 To 2) As Double
    points(1, 1) = x1: points(1, 2) = y1: points(2, 1) = x2: points(2, 2) = y1
    points(3, 1) = x2: points(3, 2) = y2: points(4, 1) = x1: points(4, 2) = y2
    TestContourRectanglePoints = points
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' габариты окружности и повёрнутого наружного контура берутся
' по точной границе, а не по меньшей сетке. Без наружного контура размер
' остаётся приблизительным. Выбор единиц ввода и вывода не меняет координаты чертежа.
Private Sub TestCheckImportedContourDimensions()
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
            Set region = query.PolygonRegion(TestContourRectanglePoints(-100#, -40#, 100#, 40#))
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
            If annotations.AnnotationID(i) = "DIM_AUTO_BOUNDS_B" Then TestCheckArea "dimension.width." & CStr(kind), annotations.Value(i), expectedB
            If annotations.AnnotationID(i) = "DIM_AUTO_BOUNDS_H" Then TestCheckArea "dimension.height." & CStr(kind), annotations.Value(i), expectedH
            If annotations.AnnotationType(i) = "DIMENSION" Then
                TestCheck "dimension.approxOnlyWithoutOuter." & CStr(kind) & "." & CStr(i), _
                    (annotations.Text(i) = ChrW$(&H2248)) = (kind = 0)
            End If
        Next i
    Next kind
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' два габарита после настоящего конвейера импорта, записи Results и отрисовки Excel.
Public Function RunImportedDimensionTests(Optional ByVal imageDirectory As String = "") As String
    Dim section As CSectionModel, space As Collection, layers As Collection, cell As CTestFakeAcadRegion
    Dim importer As CAutoCADSectionModelImporter, outer As CTestFakeAcadContour, writer As CNDMResultsWriter
    Dim settings As CSystemSettingsReader, reader As CSectionPlotDataReader, plotter As CSectionPlotter
    Dim annotations As CSectionAnnotations, chart As Object, shape As Object, system As Object, plotSettings As Object
    Dim savedSystem As Variant, savedPlot As Variant, kind As Long, row As Long, i As Long, mode As Long
    Dim b As Long, h As Long, count As Long, horizontal As Long, vertical As Long, textCount As Long
    Dim enabledRow As Long, lineBottom As Double, contourBottom As Double, layerName As Variant
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = ""
    Set system = ThisWorkbook.Names("rngSystemSettings").RefersToRange: savedSystem = system.Formula
    Set plotSettings = ThisWorkbook.Names("rngPlotAnnotationSettings").RefersToRange: savedPlot = plotSettings.Formula
    For row = 2 To system.Rows.Count
        If CStr(system.Cells(row, 1).Value2) = "Plot.RebarLabels.Enabled" Then system.Cells(row, 2).Value2 = "No"
        If CStr(system.Cells(row, 1).Value2) = "Plot.Enabled" Then system.Cells(row, 2).Value2 = "Yes"
    Next row
    For row = 2 To plotSettings.Rows.Count
        If CStr(plotSettings.Cells(row, 1).Value2) = "Enabled" Then enabledRow = row: Exit For
    Next row
    If enabledRow = 0 Then Err.Raise 5, , "Нет строки Enabled в таблице аннотаций."
    plotSettings.Cells(enabledRow, 2).Value2 = "No"
    For kind = 0 To 1
        Set space = New Collection: Set layers = New Collection
        For Each layerName In Array("CONCRETE", "REBAR", "OUTER", "OPENING")
            Set outer = New CTestFakeAcadContour: layers.Add outer, CStr(layerName)
        Next layerName
        Set cell = New CTestFakeAcadRegion: cell.TestInitialize 19200#, 0#, 0#, 1#, 1#, 0#, "CONCRETE", "Mesh": space.Add cell
        Set cell = New CTestFakeAcadRegion: cell.TestInitialize GEOM_PI * 100#, 0#, 0#, 1#, 1#, 0#, "REBAR", "Bar": space.Add cell
        If kind = 1 Then
            Set outer = TestContour("AcDbPolyline", TestRectangleEdges(-100#, -150#, 200#, 300#), "OUTER"): space.Add outer
        End If
        Set importer = New CAutoCADSectionModelImporter
        Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
        Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
        Set annotations = section.Annotations: count = annotations.Count
        PrepareSectionSnapshot section
        TestCheck "import.dimension.idempotent." & CStr(kind), annotations.Count = count
        b = 0: h = 0
        For i = 1 To annotations.Count
            If annotations.AnnotationID(i) = "DIM_AUTO_BOUNDS_B" Then b = i
            If annotations.AnnotationID(i) = "DIM_AUTO_BOUNDS_H" Then h = i
        Next i
        TestCheck "import.dimension.both." & CStr(kind), b > 0 And h > 0
        If b = 0 Or h = 0 Then Err.Raise 5, , "Отсутствует габаритная аннотация."
        TestCheck "import.dimension.width.bottom." & CStr(kind), annotations.StartY(b) = annotations.EndY(b) And annotations.EndX(b) > annotations.StartX(b) And annotations.NormalX(b) = 0# And annotations.NormalY(b) = -1#
        TestCheck "import.dimension.height.left." & CStr(kind), annotations.StartX(h) = annotations.EndX(h) And annotations.EndY(h) > annotations.StartY(h) And annotations.NormalX(h) = -1# And annotations.NormalY(h) = 0#
        TestCheck "import.dimension.exact.sign." & CStr(kind), (annotations.Text(b) = ChrW$(&H2248)) = (kind = 0) And (annotations.Text(h) = ChrW$(&H2248)) = (kind = 0)
        For mode = 0 To 1
            plotSettings.Cells(enabledRow, 3).Value2 = IIf(mode = 0, "Yes", "No")
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set reader = New CSectionPlotDataReader: reader.LoadGeometryPreviewFromWorkbook ThisWorkbook, settings
            Set plotter = New CSectionPlotter: plotter.Draw ThisWorkbook, reader, settings
            Set chart = ThisWorkbook.Worksheets("Расчет").ChartObjects("chtNDMSectionPlot").Chart
            horizontal = 0: vertical = 0: lineBottom = 0#: contourBottom = 0#
            For Each shape In chart.Shapes
                If InStr(shape.Name, "AnnotationLine") > 0 Then
                    If shape.Line.BeginArrowheadStyle <> 1 And shape.Line.EndArrowheadStyle <> 1 Then
                        If shape.Width > 1# And shape.Height < 0.01 Then horizontal = horizontal + 1: lineBottom = shape.Top
                        TestCheck "import.dimension.line.in.chart", shape.Left >= 0# And shape.Top >= 0# And shape.Left + shape.Width <= chart.ChartArea.Width And shape.Top + shape.Height <= chart.ChartArea.Height
                        If shape.Height > 1# And shape.Width < 0.01 Then vertical = vertical + 1
                    End If
                End If
                If InStr(shape.Name, "Contour") > 0 Then
                    If shape.Top + shape.Height > contourBottom Then contourBottom = shape.Top + shape.Height
                End If
            Next shape
            TestCheck "import.dimension.render.horizontal." & CStr(kind) & "." & CStr(mode), horizontal = IIf(mode = 0, 1, 0)
            TestCheck "import.dimension.render.vertical." & CStr(kind) & "." & CStr(mode), vertical = IIf(mode = 0, 1, 0)
            If kind = 1 And mode = 0 Then TestCheck "import.dimension.render.below", contourBottom > 0# And lineBottom > contourBottom
            If Len(imageDirectory) > 0 Then TestCheck "import.dimension.png", chart.Export(imageDirectory & "\ImportDimensions_" & CStr(kind) & "_" & CStr(mode) & ".png", "PNG")
        Next mode
    Next kind
    GoTo Restore
Failed:
    TestCheck "import.dimension.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Restore:
    If Not system Is Nothing And Not IsEmpty(savedSystem) Then system.Formula = savedSystem
    If Not plotSettings Is Nothing And Not IsEmpty(savedPlot) Then plotSettings.Formula = savedPlot
    RunImportedDimensionTests = mReport & "TOTAL IMPORTED DIMENSIONS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' реальные линии обоих Offset для сечений меньше, равных и больше 1000 мм.
Public Function RunAnnotationSectionScaleTests() As String
    Dim system As Object, table As Object, savedSystem As Variant, saved As Variant, value As Variant, d As Double
    Dim row As Long, section As CSectionModel, query As CSectionGeometryQuery, region As CConcreteRegion
    Dim writer As CNDMResultsWriter, reader As CSectionPlotDataReader, settings As CSystemSettingsReader, plotter As CSectionPlotter
    Dim chart As Object, shape As Object, contourTop As Double, contourBottom As Double, dimensionY As Double, rebarY As Double, pixelsPerMm As Double, expected As Double
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = ""
    Set system = ThisWorkbook.Names("rngSystemSettings").RefersToRange: savedSystem = system.Formula
    Set table = ThisWorkbook.Names("rngPlotAnnotationSettings").RefersToRange: saved = table.Formula
    For row = 2 To table.Rows.Count
        Select Case CStr(table.Cells(row, 1).Value2)
            Case "Enabled": table.Cells(row, 2).Value2 = "Yes": table.Cells(row, 3).Value2 = "Yes"
            Case "Offset": table.Cells(row, 2).Value2 = 25#: table.Cells(row, 3).Value2 = 60#
            Case "Placement": table.Cells(row, 2).Value2 = "Outside": table.Cells(row, 3).Value2 = "Outside"
            Case "LineEnabled": table.Cells(row, 2).Value2 = "Yes"
        End Select
    Next row
    For Each value In Array(500#, 1000#, 2000#, 4000#)
        d = CDbl(value): Set section = New CSectionModel: section.SourceType = "AutoCADImport"
        section.AddConcreteElement 0#, 0#, d * d / 2#, 1, , , "Rectangle", d / 2#, d
        section.AddRebarElement 0#, 0#, 20#, GEOM_PI * 100#, "Rebar"
        Set query = New CSectionGeometryQuery: Set region = query.PolygonRegion(TestContourRectanglePoints(-d / 4#, -d / 2#, d / 4#, d / 2#))
        section.Contours.AddMaterialRegion region, "OUTER", "Габарит для проверки масштаба оформления."
        Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set reader = New CSectionPlotDataReader: reader.LoadGeometryPreviewFromWorkbook ThisWorkbook, settings
        Set plotter = New CSectionPlotter: plotter.Draw ThisWorkbook, reader, settings
        Set chart = ThisWorkbook.Worksheets("Расчет").ChartObjects("chtNDMSectionPlot").Chart
        contourTop = 1E+30: contourBottom = -1E+30: dimensionY = 0#: rebarY = 0#
        For Each shape In chart.Shapes
            If InStr(shape.Name, "Contour") > 0 Then
                If shape.Top < contourTop Then contourTop = shape.Top
                If shape.Top + shape.Height > contourBottom Then contourBottom = shape.Top + shape.Height
            End If
        Next shape
        For Each shape In chart.Shapes
            If InStr(shape.Name, "AnnotationLine") > 0 And shape.Width > 1# And shape.Height < 0.01 Then
                If shape.Line.BeginArrowheadStyle <> 1 And shape.Line.EndArrowheadStyle <> 1 Then dimensionY = shape.Top
                If shape.Line.BeginArrowheadStyle = 1 And shape.Top < contourTop - 0.1 Then rebarY = shape.Top
            End If
        Next shape
        pixelsPerMm = (contourBottom - contourTop) / d: expected = 1#: If d > 1000# Then expected = d / 1000#
        TestCheck "offset.pixelsPerMm.lines." & CStr(d), pixelsPerMm > 0# And dimensionY > contourBottom And rebarY < contourTop
        If pixelsPerMm > 0# Then
            TestCheck "offset.pixelsPerMm.dimension." & CStr(d), Abs((dimensionY - contourBottom) / pixelsPerMm - 60# * expected) < 0.01
            TestCheck "offset.pixelsPerMm.rebar." & CStr(d), Abs((contourTop - rebarY) / pixelsPerMm - 25# * expected) < 0.01 ' Общая привязка к бетону исключает исходное смещение опорной линии подписи.
        End If
    Next value
    GoTo Restore
Failed:
    TestCheck "offset.pixelsPerMm.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Restore:
    If Not system Is Nothing And Not IsEmpty(savedSystem) Then system.Formula = savedSystem
    If Not table Is Nothing And Not IsEmpty(saved) Then table.Formula = saved
    RunAnnotationSectionScaleTests = mReport & "TOTAL ANNOTATION SECTION SCALE: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Все объекты слоёв читаются независимо от меток RC_SECTION_NDM.
' Чистый набор сохраняется в Results; наложенные копии не пропускаются.
Private Sub TestCheckMixedContourSources()
    Dim space As Collection, layers As Collection, cell As CTestFakeAcadRegion, layer As CTestFakeAcadContour
    Dim outer As CTestFakeAcadContour, opening As CTestFakeAcadContour, duplicate As CTestFakeAcadContour
    Dim importer As CAutoCADSectionModelImporter, section As CSectionModel, restored As CSectionModel
    Dim query As CSectionGeometryQuery, region As CConcreteRegion, writer As CNDMResultsWriter
    Dim mode As Long, i As Long, countBefore As Long, prefix As String, outerCount As Long, openingCount As Long
    Dim number As Long, description As String
    Set importer = New CAutoCADSectionModelImporter: Set query = New CSectionGeometryQuery
    Set writer = New CNDMResultsWriter
    For mode = 0 To 3
        Set space = New Collection: Set layers = New Collection
        Set cell = New CTestFakeAcadRegion
        cell.TestInitialize 10000#, 50#, 50#, 8333333.33333333, 8333333.33333333, 0#, "CONCRETE", "CELL", 0#, True
        space.Add cell
        Set cell = New CTestFakeAcadRegion
        cell.TestInitialize GEOM_PI * 25#, 10#, 10#, 100#, 100#, 0#, "REBAR", "BAR": space.Add cell
        Set outer = TestContour("AcDbRegion", TestRectangleEdges(0#, 0#, 100#, 100#))
        outer.TestSetLegacyContourTag (mode = 1 Or mode = 3): space.Add outer
        For i = 0 To 1
            Set opening = TestContour("AcDbPolyline", TestRectangleEdges(25# + 35# * i, 30# + 30# * i, 10# + 10# * i, 10#), "OPENING")
            opening.TestSetLegacyContourTag (mode = 0 Or mode = 3): space.Add opening
        Next i
        Set layer = New CTestFakeAcadContour: layers.Add layer, "OUTER"
        Set layer = New CTestFakeAcadContour: layers.Add layer, "OPENING"
        countBefore = space.Count: prefix = "mixedSources." & CStr(mode)
        Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
        query.Initialize section: Set region = query.ConcreteDomain
        TestCheckArea prefix & ".area", region.Area, 9700#
        TestCheck prefix & ".authoritativeOuter", InStr(region.Source, "AuthoritativeContour") > 0
        TestCheck prefix & ".threeLoops", region.LoopCount = 3
        TestCheck prefix & ".bothOpeningsExcluded", Not query.ContainsPoint(region, 30#, 35#) And Not query.ContainsPoint(region, 70#, 65#)
        TestCheck prefix & ".sourcesPreserved", space.Count = countBefore And Not outer.TestDeleted And Not opening.TestDeleted
        TestCheck prefix & ".noOwnershipRead", outer.TestXDataReads = 0 And opening.TestXDataReads = 0
        TestCheckArea prefix & ".mechanicalArea", section.ConcreteArea(1), 10000#
        writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
        Set restored = ReadSectionGeometryFromResults(ThisWorkbook, "AutoCADImport")
        query.Initialize restored: Set region = query.ConcreteDomain
        TestCheckArea prefix & ".resultsArea", region.Area, 9700#
        TestCheck prefix & ".resultsLoops", region.LoopCount = 3
        outerCount = 0: openingCount = 0
        For i = 1 To restored.Contours.Count
            If restored.Contours.LoopRole(i) = "Opening" Then openingCount = openingCount + 1 Else outerCount = outerCount + 1
        Next i
        TestCheck prefix & ".resultsBothRoles", outerCount = 4 And openingCount = 8
    Next mode
    ' Дубли opening и outer одинаково ошибочны с меткой и без нее.
    For mode = 0 To 3
        If mode < 2 Then
            Set duplicate = TestContour("AcDbPolyline", TestRectangleEdges(25#, 30#, 10#, 10#), "OPENING")
        Else
            Set duplicate = TestContour("AcDbPolyline", TestRectangleEdges(0#, 0#, 100#, 100#))
        End If
        duplicate.TestSetLegacyContourTag (mode Mod 2 = 1): space.Add duplicate
        On Error Resume Next
        Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
        number = Err.Number: description = Err.Description: Err.Clear
        On Error GoTo 0
        TestCheck "mixedSources.duplicateRejected." & CStr(mode), number <> 0
        TestCheck "mixedSources.duplicatePreserved." & CStr(mode), Not duplicate.TestDeleted And duplicate.TestXDataReads = 0
        space.Remove space.Count
    Next mode
    Set duplicate = TestContour("AcDbPolyline", TestRectangleEdges(120#, 0#, 20#, 20#)) ' Метка не скрывает отдельную непересекающуюся бетонную часть.
    duplicate.TestSetLegacyContourTag True: space.Add duplicate
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    query.Initialize section
    TestCheckArea "mixedSources.allDisjointContoursImported", query.ConcreteDomain.Area, 10100#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' встроенные отверстия Region и бетонные части читаются
' универсально. Дубли и отверстия вне материала остаются ошибками.
Private Sub TestCheckMixedContourTopology()
    Dim space As Collection, layers As Collection, cell As CTestFakeAcadRegion, layer As CTestFakeAcadContour
    Dim outer As CTestFakeAcadContour, part As CTestFakeAcadContour, hole As CTestFakeAcadContour, edges As Collection, edge As Variant
    Dim importer As CAutoCADSectionModelImporter, section As CSectionModel, query As CSectionGeometryQuery
    Dim region As CConcreteRegion, number As Long, description As String
    Set space = New Collection: Set layers = New Collection
    Set cell = New CTestFakeAcadRegion
    cell.TestInitialize 10000#, 50#, 50#, 8333333.33333333, 8333333.33333333, 0#, "CONCRETE", "CELL", 0#, True: space.Add cell
    Set cell = New CTestFakeAcadRegion
    cell.TestInitialize GEOM_PI * 25#, 10#, 10#, 100#, 100#, 0#, "REBAR", "BAR": space.Add cell
    Set layer = New CTestFakeAcadContour: layers.Add layer, "OUTER"
    Set layer = New CTestFakeAcadContour: layers.Add layer, "OPENING"
    Set edges = TestRectangleEdges(0#, 0#, 100#, 100#)
    For Each edge In TestRectangleEdges(25#, 30#, 10#, 10#): edges.Add edge: Next edge
    Set outer = TestContour("AcDbRegion", edges): space.Add outer
    Set hole = TestContour("AcDbPolyline", TestRectangleEdges(25#, 30#, 10#, 10#), "OPENING")
    hole.TestSetLegacyContourTag True: space.Add hole
    Set hole = TestContour("AcDbPolyline", TestRectangleEdges(60#, 60#, 20#, 10#), "OPENING")
    hole.TestSetLegacyContourTag True: space.Add hole
    Set importer = New CAutoCADSectionModelImporter: Set query = New CSectionGeometryQuery
    On Error Resume Next
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    number = Err.Number: Err.Clear
    On Error GoTo 0
    TestCheck "mixedTopology.embeddedOpening.duplicateRejected", number <> 0
    space.Remove 4
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    query.Initialize section: Set region = query.ConcreteDomain
    TestCheckArea "mixedTopology.embeddedOpening.area", region.Area, 9700#
    TestCheck "mixedTopology.embeddedOpening.threeLoops", region.LoopCount = 3
    TestCheck "mixedTopology.embeddedOpening.bothExcluded", Not query.ContainsPoint(region, 30#, 35#) And Not query.ContainsPoint(region, 70#, 65#)
    space.Remove 4: space.Remove 3
    Set outer = TestContour("AcDbPolyline", TestRectangleEdges(0#, 0#, 100#, 100#)): outer.TestSetLegacyContourTag True: space.Add outer
    Set part = TestContour("AcDbPolyline", TestRectangleEdges(40#, 40#, 20#, 20#)): part.TestSetLegacyContourTag True: space.Add part
    Set hole = TestContour("AcDbRegion", TestRectangleEdges(20#, 20#, 60#, 60#), "OPENING"): space.Add hole
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    query.Initialize section: Set region = query.ConcreteDomain
    TestCheckArea "mixedTopology.separateConcrete.area", region.Area, 6800#
    TestCheck "mixedTopology.separateConcrete.threeLoops", region.LoopCount = 3
    TestCheck "mixedTopology.separateConcrete.roles", query.ContainsPoint(region, 50#, 50#) And Not query.ContainsPoint(region, 30#, 30#)
    space.Add TestContour("AcDbPolyline", TestRectangleEdges(120#, 20#, 10#, 10#), "OPENING")
    On Error Resume Next
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo 0
    TestCheck "mixedTopology.outsideOpening.rejected", number <> 0 And InStr(description, "слое") > 0
    TestCheck "mixedTopology.outsideOpening.oldModelKept", section.Contours.Count = 12 And Abs(query.ConcreteDomain.Area - 6800#) < 0.00001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отсутствующие общие слои не отменяют сетку; отдельный Region
' внутри уже вырезанного проема возвращает бетон, а не заполняет весь проем.
' Дубли не скрываются; чистоту импортируемого набора обеспечивает пользователь.
Private Sub TestCheckOptionalLayersAndSeparateConcrete()
    Dim space As Collection, layers As Collection, cell As CTestFakeAcadRegion, dummy As CTestFakeAcadContour
    Dim importer As CAutoCADSectionModelImporter, section As CSectionModel, query As CSectionGeometryQuery
    Dim region As CConcreteRegion, outer As CTestFakeAcadContour, holeEdges As Collection, edge As Variant
    Dim part As CTestFakeAcadContour, duplicate As CTestFakeAcadContour, countBefore As Long, number As Long
    Set space = New Collection: Set layers = New Collection
    Set cell = New CTestFakeAcadRegion
    cell.TestInitialize 10000#, 50#, 50#, 8333333.33333333, 8333333.33333333, 0#, "CONCRETE", "CELL", 0#, True
    space.Add cell
    Set cell = New CTestFakeAcadRegion
    cell.TestInitialize GEOM_PI * 25#, 50#, 50#, 100#, 100#, 0#, "REBAR", "BAR": space.Add cell
    Set importer = New CAutoCADSectionModelImporter
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    TestCheck "optional.absentLayers.meshKept", section.ConcreteCount = 1 And section.RebarCount = 1 And section.Contours.Count = 0
    TestCheck "optional.absentLayers.comment", InStr(importer.ImportComment, "OUTER") > 0 And InStr(importer.ImportComment, "OPENING") > 0
    Set holeEdges = TestRectangleEdges(0#, 0#, 100#, 100#)
    For Each edge In TestRectangleEdges(20#, 20#, 60#, 60#): holeEdges.Add edge: Next edge
    Set outer = TestContour("AcDbRegion", holeEdges): space.Add outer
    Set part = TestContour("AcDbRegion", TestRectangleEdges(40#, 40#, 20#, 20#)): space.Add part
    Set dummy = New CTestFakeAcadContour: layers.Add dummy, "OUTER"
    Set duplicate = TestContour("AcDbPolyline", holeEdges): duplicate.TestSetLegacyContourTag True: space.Add duplicate
    countBefore = space.Count
    On Error Resume Next
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    number = Err.Number: Err.Clear
    On Error GoTo 0
    TestCheck "separateConcrete.duplicateRejected", number <> 0 And Not duplicate.TestDeleted
    space.Remove 5: countBefore = space.Count
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    Set query = New CSectionGeometryQuery: query.Initialize section: Set region = query.ConcreteDomain
    TestCheckArea "separateConcrete.area", region.Area, 6800#
    TestCheck "separateConcrete.threeContours", region.LoopCount = 3
    TestCheck "separateConcrete.material", query.ContainsPoint(region, 50#, 50#) And query.ContainsPoint(region, 10#, 10#)
    TestCheck "separateConcrete.void", Not query.ContainsPoint(region, 30#, 30#)
    TestCheck "separateConcrete.sourcesKept", space.Count = countBefore And Not outer.TestDeleted And Not part.TestDeleted And Not duplicate.TestDeleted
    TestCheck "separateConcrete.openingAbsent", InStr(importer.ImportComment, "OPENING") > 0
    ' Новый чертеж содержит только экспорт: наружная граница, отверстие
    ' и отдельная бетонная часть должны восстановить те же три контура.
    space.Remove 4: space.Remove 3
    Set outer = TestContour("AcDbPolyline", TestRectangleEdges(0#, 0#, 100#, 100#)): outer.TestSetLegacyContourTag True: space.Add outer
    Set part = TestContour("AcDbPolyline", TestRectangleEdges(40#, 40#, 20#, 20#)): part.TestSetLegacyContourTag True: space.Add part
    Set duplicate = TestContour("AcDbPolyline", TestRectangleEdges(20#, 20#, 60#, 60#), "OPENING")
    duplicate.TestSetLegacyContourTag True: space.Add duplicate
    Set dummy = New CTestFakeAcadContour: layers.Add dummy, "OPENING"
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    query.Initialize section: Set region = query.ConcreteDomain
    TestCheckArea "exportOnly.separateConcrete.area", region.Area, 6800#
    TestCheck "exportOnly.separateConcrete.threeContours", region.LoopCount = 3
    TestCheck "exportOnly.separateConcrete.material", query.ContainsPoint(region, 50#, 50#) And Not query.ContainsPoint(region, 30#, 30#)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' существующие пустые слои не делают контуры обязательными.
' Проверяет пересечение по бетонной сетке, отдельные отверстия без outer
' и сохранение строгой ошибки для реально заданной незамкнутой полилинии.
Private Sub TestCheckEmptyContourLayers()
    Dim space As Collection, layers As Collection, cell As CTestFakeAcadRegion, bar As CTestFakeAcadRegion
    Dim layer As CTestFakeAcadContour, opening As CTestFakeAcadContour, invalid As CTestFakeAcadContour
    Dim importer As CAutoCADSectionModelImporter, section As CSectionModel, query As CSectionGeometryQuery
    Dim domain As CConcreteRegion, clipped As CConcreteRegion, mode As Long, expectedArea As Double
    Dim contourLayer As String, openingLayer As String, prefix As String, number As Long, description As String
    Set importer = New CAutoCADSectionModelImporter
    For mode = 0 To 3
        Set space = New Collection: Set layers = New Collection
        Set cell = New CTestFakeAcadRegion
        cell.TestInitialize 10000#, 50#, 50#, 8333333.33333333, 8333333.33333333, 0#, "CONCRETE", "CELL", 0#, True
        space.Add cell
        Set bar = New CTestFakeAcadRegion
        bar.TestInitialize GEOM_PI * 25#, 10#, 10#, 100#, 100#, 0#, "REBAR", "BAR": space.Add bar
        contourLayer = "OUTER": openingLayer = "OPENING": expectedArea = 10000#
        Set layer = New CTestFakeAcadContour: layers.Add layer, "OUTER"
        If mode > 0 Then
            Set layer = New CTestFakeAcadContour: layers.Add layer, "OPENING"
        End If
        If mode = 2 Then
            Set opening = TestContour("AcDbPolyline", TestRectangleEdges(60#, 60#, 20#, 20#), "OPENING")
            space.Add opening: expectedArea = 9600#
        ElseIf mode = 3 Then
            contourLayer = vbNullString: openingLayer = vbNullString
        End If
        Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, contourLayer, openingLayer, layers)
        Set query = New CSectionGeometryQuery: query.Initialize section: Set domain = query.ConcreteDomain
        prefix = "emptyLayers." & CStr(mode)
        TestCheck prefix & ".meshKept", section.ConcreteCount = 1 And section.RebarCount = 1
        TestCheck prefix & ".noInventedContour", section.Contours.Count = IIf(mode = 2, 4, 0)
        TestCheck prefix & ".meshFallback", InStr(domain.Source, "AuthoritativeContour") = 0
        TestCheck prefix & ".comment", InStr(importer.ImportComment, "бетонной сетке") > 0
        TestCheck prefix & ".sourceKept", Not cell.TestDeleted And Not bar.TestDeleted
        TestCheckArea prefix & ".area", domain.Area, expectedArea
        Set clipped = query.ClipHalfPlane(domain, 1#, 0#, 50#)
        TestCheckArea prefix & ".meshIntersection", clipped.Area, expectedArea - 5000#
        If mode = 2 Then TestCheck prefix & ".openingKept", Not query.ContainsPoint(domain, 70#, 70#) And Not opening.TestDeleted
    Next mode
    Set invalid = New CTestFakeAcadContour
    Dim openEdges As Collection
    Set openEdges = TestRectangleEdges(0#, 0#, 100#, 100#): openEdges.Remove 1
    invalid.TestInitializeLoop "AcDbPolyline", "OUTER", "OPEN", openEdges, False
    space.Add invalid
    On Error Resume Next
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo 0
    TestCheck "emptyLayers.invalidContourStillRejected", number <> 0 And InStr(description, "не замкнут") > 0
    TestCheck "emptyLayers.invalidSourceKept", Not invalid.TestDeleted
    TestCheck "emptyLayers.failedImportKeepsPreviousModel", section.ConcreteCount = 1 And section.RebarCount = 1 And section.Contours.Count = 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет несколько отверстий и дуги обычной записью предварительного
' просмотра, меняет текущие единицы и восстанавливает модель. Проверяет также
' модель только с отверстиями и адрес ошибки повреждённых сохранённых данных.
Private Sub TestCheckResultsRoundTrip()
    Dim unitRange As Object, savedUnits As Variant, settings As CSystemSettingsReader, units As CUnitSystem
    Dim section As CSectionModel, restored As CSectionModel, query As CSectionGeometryQuery, region As CConcreteRegion
    Dim reader As CAutoCADContourReader, writer As CNDMResultsWriter, contourEntity As CTestFakeAcadContour
    Dim edges As Collection, curve As CTestFakeAcadContour, mode As Long, row As Long, lengthUnit As String
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
            Set edges = TestRectangleEdges(0#, 0#, 100#, 100#)
            Set curve = New CTestFakeAcadContour: curve.TestInitializeCircle 35#, 35#, 10#: edges.Add curve
            Set curve = New CTestFakeAcadContour: curve.TestInitializeCircle 70#, 65#, 5#: edges.Add curve
            Set contourEntity = TestContour("AcDbRegion", edges)
            Set region = reader.ReadRegion(contourEntity)
            section.Contours.AddMaterialRegion region, "ROUNDTRIP", "Точный наружный контур с двумя проемами."
        Else
            For row = 0 To 1
                Set curve = New CTestFakeAcadContour: curve.TestInitializeCircle 35# + 35# * row, 35# + 30# * row, 10# - 5# * row
                Set region = reader.ReadRegion(curve)
                section.Contours.AddMaterialRegion region, "HOLE_" & CStr(row), "Отдельный проем.", True
            Next row
        End If
        writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section), units
        ' Новые настройки единиц ввода и вывода не меняют единицы уже сохранённых Results.
        For row = 2 To unitRange.Rows.Count
            If CStr(unitRange.Cells(row, 1).Value2) = "Length" Then
                unitRange.Cells(row, 2).Value2 = "cm"
                If lengthUnit = "m" Then unitRange.Cells(row, 4).Value2 = "mm" Else unitRange.Cells(row, 4).Value2 = "m"
            End If
        Next row
        Set restored = ReadSectionGeometryFromResults(ThisWorkbook, "AutoCADImport")
        Set query = New CSectionGeometryQuery: query.Initialize restored: Set region = query.ConcreteDomain
        TestCheckArea "snapshot." & CStr(mode) & ".area", region.Area, 10000# - 125# * GEOM_PI
        TestCheck "snapshot." & CStr(mode) & ".threeLoops", region.LoopCount = 3
        TestCheck "snapshot." & CStr(mode) & ".holes", Not query.ContainsPoint(region, 35#, 35#) And Not query.ContainsPoint(region, 70#, 65#)
        TestCheck "snapshot." & CStr(mode) & ".material", query.ContainsPoint(region, 50#, 50#)
        TestCheckArea "snapshot." & CStr(mode) & ".mechanicalArea", restored.ConcreteArea(1), 10000#
        restored.ConcreteLocalInertiaComponents 1, ix, iy, ixy
        TestCheck "snapshot." & CStr(mode) & ".mechanicalInertia", Abs(ix - 123456#) < 0.00001 And Abs(iy - 654321#) < 0.00001 And Abs(ixy - 12#) < 0.00001
        If mode < 2 Then
            TestCheck "snapshot." & CStr(mode) & ".authoritative", InStr(region.Source, "AuthoritativeContour") > 0
        Else
            TestCheck "snapshot." & CStr(mode) & ".meshFallback", InStr(region.Source, "AuthoritativeContour") = 0
        End If
        Set anchor = ThisWorkbook.Names.Item("rngNDMSectionContours").RefersToRange
        Dim contourRow As Long, annotationAnchor As Object
        contourRow = 1
        TestCheck "snapshot." & CStr(mode) & ".contoursOwnBlock", anchor.Value2 = "RunID" And anchor.Offset(contourRow, 5).Value2 Like "CONTOUR_*"
        TestCheck "snapshot." & CStr(mode) & ".explicitRole", Len(CStr(anchor.Offset(contourRow, 4).Value2)) > 0
        Set annotationAnchor = ThisWorkbook.Names.Item("rngNDMSectionAnnotations").RefersToRange
        TestCheck "snapshot." & CStr(mode) & ".noSourceAnnotations", Application.CountIf(annotationAnchor.Offset(1, 0).Resize(1000, 1), "CONTOUR_*") = 0
        savedValue = anchor.Offset(contourRow, 6).Value2: anchor.Offset(contourRow, 6).Value2 = "invalid"
        On Error Resume Next
        Set restored = ReadSectionGeometryFromResults(ThisWorkbook, "AutoCADImport")
        number = Err.Number: description = Err.Description: Err.Clear
        On Error GoTo Failed
        anchor.Offset(contourRow, 6).Value2 = savedValue
        TestCheck "snapshot." & CStr(mode) & ".invalidCoordinate", number <> 0 And InStr(description, anchor.Offset(contourRow, 6).Address(False, False)) > 0
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

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет утверждение и независимый численный эталон.
Private Sub TestCheck(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        mPassed = mPassed + 1: mReport = mReport & "OK: " & name & vbCrLf
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
    End If
    If mNativeProgress Then TestNativeProgress name
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' оставляет этап в отдельном журнале собственной тестовой книги.
Private Sub TestNativeProgress(ByVal phase As String)
    Dim file As Integer
    file = FreeFile
    Open ThisWorkbook.Path & "\NativeCAD.progress.txt" For Append As #file
    Print #file, Format$(Now, "hh:nn:ss") & " " & phase
    Close #file
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' несколько opening читаются отдельными Region/Polyline как с
' достоверным outer, так и при наружной границе, восстановленной по сетке. Последующий
' ошибочный импорт не меняет уже полученную пригодную модель.
Private Sub TestCheckMultipleOpenings()
    Dim space As Collection, layers As Collection, cell As CTestFakeAcadRegion, layer As CTestFakeAcadContour
    Dim firstOpening As CTestFakeAcadContour, secondOpening As CTestFakeAcadContour, outer As CTestFakeAcadContour
    Dim importer As CAutoCADSectionModelImporter, section As CSectionModel, query As CSectionGeometryQuery
    Dim region As CConcreteRegion, mode As Long, outerLayer As String, number As Long, description As String, annotationCount As Long
    For mode = 0 To 1
        Set space = New Collection: Set layers = New Collection
        Set cell = New CTestFakeAcadRegion
        cell.TestInitialize 10000#, 50#, 50#, 8333333.33333333, 8333333.33333333, 0#, "CONCRETE", "CELL", 0#, True
        space.Add cell
        Set cell = New CTestFakeAcadRegion
        cell.TestInitialize GEOM_PI * 25#, 10#, 10#, 100#, 100#, 0#, "REBAR", "BAR": space.Add cell
        outerLayer = vbNullString
        If mode = 1 Then
            Set outer = TestContour("AcDbRegion", TestRectangleEdges(0#, 0#, 100#, 100#))
            space.Add outer: outerLayer = "OUTER"
        End If
        Set firstOpening = TestContour("AcDbRegion", TestRectangleEdges(25#, 30#, 10#, 10#), "OPENING")
        Set secondOpening = TestContour("AcDbPolyline", TestRectangleEdges(60#, 60#, 20#, 10#), "OPENING")
        space.Add firstOpening: space.Add secondOpening
        Set layer = New CTestFakeAcadContour: layers.Add layer, "OUTER"
        Set layer = New CTestFakeAcadContour: layers.Add layer, "OPENING"
        Set importer = New CAutoCADSectionModelImporter
        Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, outerLayer, "OPENING", layers)
        Set query = New CSectionGeometryQuery: query.Initialize section: Set region = query.ConcreteDomain
        TestCheckArea "multiple." & CStr(mode) & ".area", region.Area, 9700#
        TestCheck "multiple." & CStr(mode) & ".threeLoops", region.LoopCount = 3
        TestCheck "multiple." & CStr(mode) & ".bothExcluded", Not query.ContainsPoint(region, 30#, 35#) And Not query.ContainsPoint(region, 70#, 65#)
        TestCheck "multiple." & CStr(mode) & ".material", query.ContainsPoint(region, 50#, 50#)
        TestCheck "multiple." & CStr(mode) & ".sourceKept", Not firstOpening.TestDeleted And Not secondOpening.TestDeleted
        If mode = 0 Then TestCheck "multiple.meshOuterNotAuthoritative", InStr(region.Source, "AuthoritativeContour") = 0
        annotationCount = section.Contours.Count
        space.Add TestContour("AcDbRegion", TestRectangleEdges(60#, 60#, 20#, 10#), "OPENING")
        On Error Resume Next
        Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, outerLayer, "OPENING", layers)
        number = Err.Number: description = Err.Description: Err.Clear
        On Error GoTo 0
        TestCheck "multiple." & CStr(mode) & ".duplicateRejected", number <> 0 And Len(description) > 25
        TestCheck "multiple." & CStr(mode) & ".oldModelKept", section.Contours.Count = annotationCount
        TestCheck "multiple." & CStr(mode) & ".oldAreaKept", Abs(query.ConcreteDomain.Area - 9700#) < 0.00001
    Next mode
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' число отверстий не ограничено двумя. Пять отверстий разных
' поддерживаемых видов остаются отдельными замкнутыми контурами и вычитаются
' один раз. После записи и чтения Results сохраняются их координаты и единицы.
Private Sub TestCheckFiveOpenings()
    Dim space As Collection, layers As Collection, cell As CTestFakeAcadRegion, layer As CTestFakeAcadContour
    Dim importer As CAutoCADSectionModelImporter, section As CSectionModel, restored As CSectionModel
    Dim query As CSectionGeometryQuery, region As CConcreteRegion, writer As CNDMResultsWriter
    Dim i As Long, kind As String, opening As CTestFakeAcadContour
    Set space = New Collection: Set layers = New Collection
    Set cell = New CTestFakeAcadRegion
    cell.TestInitialize 10000#, 50#, 50#, 8333333.33333333, 8333333.33333333, 0#, "CONCRETE", "CELL", 0#, True
    space.Add cell
    Set cell = New CTestFakeAcadRegion
    cell.TestInitialize GEOM_PI * 25#, 5#, 5#, 100#, 100#, 0#, "REBAR", "BAR": space.Add cell
    space.Add TestContour("AcDbRegion", TestRectangleEdges(0#, 0#, 100#, 100#))
    Set layer = New CTestFakeAcadContour: layers.Add layer, "OUTER"
    Set layer = New CTestFakeAcadContour: layers.Add layer, "OPENING"
    For i = 1 To 5
        kind = "AcDbRegion"
        If i Mod 2 = 0 Then kind = "AcDbPolyline"
        Set opening = TestContour(kind, TestRectangleEdges(10# + 15# * (i - 1), 40#, 5#, 10#), "OPENING")
        space.Add opening
    Next i
    Set importer = New CAutoCADSectionModelImporter
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    Set query = New CSectionGeometryQuery: query.Initialize section: Set region = query.ConcreteDomain
    TestCheckArea "fiveOpenings.area", region.Area, 9750#
    TestCheck "fiveOpenings.sixLoops", region.LoopCount = 6
    For i = 1 To 5
        TestCheck "fiveOpenings.excluded." & CStr(i), Not query.ContainsPoint(region, 12.5 + 15# * (i - 1), 45#)
    Next i
    Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
    Set restored = ReadSectionGeometryFromResults(ThisWorkbook, "AutoCADImport")
    query.Initialize restored: Set region = query.ConcreteDomain
    TestCheckArea "fiveOpenings.snapshotArea", region.Area, 9750#
    TestCheck "fiveOpenings.snapshotLoops", region.LoopCount = 6
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' независимо проверяет площадь в мм2, не вызывая проверяемую формулу программы.
Private Sub TestCheckArea(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    TestCheck name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) < 0.00001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' создает неупорядоченный прямоугольник с обращенными гранями.
Private Function TestRectangleEdges(ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal height As Double, _
        Optional ByVal angle As Double = 0#) As Collection
    Dim points(1 To 4, 1 To 2) As Double, i As Long, j As Long, lx As Double, ly As Double
    Dim edges As Collection, line As CTestFakeAcadContour, order As Variant
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
        Set line = New CTestFakeAcadContour
        If i Mod 2 = 0 Then
            line.TestInitializeLine points(CLng(lx), 1), points(CLng(lx), 2), points(j, 1), points(j, 2)
        Else
            line.TestInitializeLine points(j, 1), points(j, 2), points(CLng(lx), 1), points(CLng(lx), 2)
        End If
        edges.Add line
    Next i
    Set TestRectangleEdges = edges
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' представляет заданные сегменты как Region или Polyline
' с теми же доступными свойствами и методами, что у объектов AutoCAD.
Private Function TestContour(ByVal kind As String, ByVal edges As Collection, Optional ByVal layerName As String = "OUTER") As CTestFakeAcadContour
    Dim result As CTestFakeAcadContour
    Set result = New CTestFakeAcadContour
    result.TestInitializeLoop kind, layerName, "C_TEST", edges
    Set TestContour = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Region и обе 2D-полилинии дают одну геометрию после поворота.
Private Sub TestCheckEquivalentSources()
    Dim reader As CAutoCADContourReader, query As CSectionGeometryQuery, entity As CTestFakeAcadContour
    Dim region As CConcreteRegion, kind As Variant, curve As CTestFakeAcadContour
    Set reader = New CAutoCADContourReader: Set query = New CSectionGeometryQuery
    For Each kind In Array("AcDbRegion", "AcDbPolyline", "AcDb2dPolyline")
        Set entity = TestContour(CStr(kind), TestRectangleEdges(137#, -73#, 100#, 80#, 0.37))
        Set region = reader.ReadRegion(entity)
        TestCheckArea CStr(kind) & ".area", region.Area, 8000#
        TestCheck CStr(kind) & ".actualEdges", region.SegmentCount = 4 And region.LoopCount = 1
        TestCheck CStr(kind) & ".outer", Not query.LoopIsOpening(region, 1)
        TestCheck CStr(kind) & ".sourceKept", Not entity.TestDeleted
        TestCheck CStr(kind) & ".copyDeleted", entity.TestLastCopy.TestDeleted
        For Each curve In entity.TestLastCopy.TestExplodedPieces: TestCheck CStr(kind) & ".pieceDeleted", curve.TestDeleted: Next curve
    Next kind
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет внутреннюю окружность и вложенный остров материала.
' Начальные направления Explode не определяют семантику outer/opening.
Private Sub TestCheckNestedRegion()
    Dim edges As Collection, island As Collection, curve As CTestFakeAcadContour, entity As CTestFakeAcadContour
    Dim reader As CAutoCADContourReader, query As CSectionGeometryQuery, region As CConcreteRegion, openings As Long, i As Long
    Set edges = TestRectangleEdges(0#, 0#, 100#, 100#)
    Set curve = New CTestFakeAcadContour: curve.TestInitializeCircle 50#, 50#, 20#: edges.Add curve
    Set island = TestRectangleEdges(45#, 45#, 10#, 10#)
    For Each curve In island: edges.Add curve: Next curve
    Set entity = TestContour("AcDbRegion", edges)
    Set reader = New CAutoCADContourReader: Set query = New CSectionGeometryQuery
    Set region = reader.ReadRegion(entity)
    TestCheckArea "nested.area", region.Area, 10100# - 400# * GEOM_PI
    TestCheck "nested.threeLoops", region.LoopCount = 3
    For i = 1 To region.LoopCount
        If query.LoopIsOpening(region, i) Then openings = openings + 1
    Next i
    TestCheck "nested.oneOpening", openings = 1
    TestCheck "nested.holeExcluded", Not query.ContainsPoint(region, 65#, 50#)
    TestCheck "nested.islandIncluded", query.ContainsPoint(region, 50#, 50#)
    TestCheck "nested.outerIncluded", query.ContainsPoint(region, 5#, 50#)
    TestCheck "nested.sourceKept", Not entity.TestDeleted
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' неподдержанные/неплоские/дублирующие границы не превращаются
' в приблизительный успешный контур; собственная Copy очищается и при ошибке.
Private Sub TestCheckInvalidContours()
    Dim reader As CAutoCADContourReader, entity As CTestFakeAcadContour, edges As Collection, curve As CTestFakeAcadContour
    Dim region As CConcreteRegion, mode As Long, number As Long, description As String
    Set reader = New CAutoCADContourReader
    For mode = 1 To 8
        Set edges = TestRectangleEdges(0#, 0#, 100#, 80#)
        Set entity = TestContour("AcDbRegion", edges)
        Select Case mode
            Case 1: edges.Remove 1
            Case 2: edges.Add edges(1)
            Case 3: entity.TestSetPlane 1#, 0#, 0#, 0#
            Case 4: Set entity = TestContour("AcDbSpline", edges)
            Case 5
                edges.Remove 1
                entity.TestInitializeLoop "AcDbPolyline", "OUTER", "OPEN", edges, False
            Case 6: entity.TestInitializeLoop "AcDbPolyline", "OUTER", "Z", edges: entity.TestSetPlane 0#, 0#, 1#, 1#
            Case 7
                Set curve = New CTestFakeAcadContour: curve.TestInitializeLoop "AcDbEllipse", "OUTER", "ELLIPSE", Nothing: edges.Add curve
            Case 8
                Set edges = New Collection
                Set curve = New CTestFakeAcadContour: curve.TestInitializeLine 0#, 0#, 100#, 80#: edges.Add curve
                Set curve = New CTestFakeAcadContour: curve.TestInitializeLine 100#, 80#, 0#, 80#: edges.Add curve
                Set curve = New CTestFakeAcadContour: curve.TestInitializeLine 0#, 80#, 100#, 0#: edges.Add curve
                Set curve = New CTestFakeAcadContour: curve.TestInitializeLine 100#, 0#, 0#, 0#: edges.Add curve
                Set entity = TestContour("AcDbRegion", edges)
        End Select
        On Error Resume Next
        Set region = reader.ReadRegion(entity)
        number = Err.Number: description = Err.Description: Err.Clear
        On Error GoTo 0
        TestCheck "invalid." & CStr(mode) & ".error", number <> 0
        TestCheck "invalid." & CStr(mode) & ".context", InStr(description, "OUTER") > 0 And Len(description) > 25
        TestCheck "invalid." & CStr(mode) & ".sourceKept", Not entity.TestDeleted
        If Not entity.TestLastCopy Is Nothing Then TestCheck "invalid." & CStr(mode) & ".copyDeleted", entity.TestLastCopy.TestDeleted
    Next mode
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' обычный импорт сохраняет Region с отверстием без изменения площадей
' и моментов инерции элементов сетки. Повторные геометрические запросы
' используют сохранённые роли наружной границы и отверстия.
Private Sub TestCheckImportedSnapshot()
    Dim space As Collection, layers As Collection, regionEntity As CTestFakeAcadRegion, contourEntity As CTestFakeAcadContour
    Dim edges As Collection, roundOpening As CTestFakeAcadContour, importer As CAutoCADSectionModelImporter
    Dim section As CSectionModel, query As CSectionGeometryQuery, domain As CConcreteRegion, layer As CTestFakeAcadContour
    Set space = New Collection: Set layers = New Collection
    Set regionEntity = New CTestFakeAcadRegion
    regionEntity.TestInitialize 10000#, 50#, 50#, 1000000#, 1200000#, 0#, "CONCRETE", "C1"
    space.Add regionEntity
    Set regionEntity = New CTestFakeAcadRegion
    regionEntity.TestInitialize GEOM_PI * 25#, 10#, 10#, 100#, 100#, 0#, "REBAR", "R1"
    space.Add regionEntity
    Set edges = TestRectangleEdges(0#, 0#, 100#, 100#)
    Set roundOpening = New CTestFakeAcadContour: roundOpening.TestInitializeCircle 50#, 50#, 20#: edges.Add roundOpening
    Set contourEntity = TestContour("AcDbRegion", edges): space.Add contourEntity
    Set layer = New CTestFakeAcadContour: layers.Add layer, "OUTER"
    Set layer = New CTestFakeAcadContour: layers.Add layer, "OPENING"
    Set importer = New CAutoCADSectionModelImporter
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    Set query = New CSectionGeometryQuery: query.Initialize section: Set domain = query.ConcreteDomain
    TestCheckArea "import.domain", domain.Area, 10000# - 400# * GEOM_PI
    TestCheckArea "import.mechanicalAreaKept", section.ConcreteArea(1), 10000#
    TestCheck "import.holeKept", domain.LoopCount = 2 And Not query.ContainsPoint(domain, 50#, 50#)
    TestCheck "import.sourceKept", Not contourEntity.TestDeleted
    TestCheck "import.emptyOpeningExplained", InStr(section.Contours.Comment(1), "объектов нет") > 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Направленная проверка Region/Polyline без полной
' матрицы CAD. Документ и приложение принадлежат внешнему runner-у.
Public Function RunRealAutoCADContourFormatTests(ByVal ownedApplication As Object, _
        ByVal templatePath As String) As String
    mPassed = 0: mFailed = 0: mReport = vbNullString: mNativeProgress = True
    Dim doc As Object, reader As CAutoCADContourReader, path As String
    On Error GoTo Failed
    TestNativeProgress "format test start"
    TestCheck "native.actualAutoCAD", InStr(1, CStr(ownedApplication.FullName), "\AutoCAD 2023\acad.exe", vbTextCompare) > 0
    If InStr(1, CStr(ownedApplication.FullName), "\AutoCAD 2023\acad.exe", vbTextCompare) = 0 Then _
        Err.Raise vbObjectError + 5502, , "Тест требует Autodesk AutoCAD 2023."
    Set doc = ownedApplication.Documents.Add(templatePath)
    TestWaitForCAD ownedApplication
    Set reader = New CAutoCADContourReader
    TestCheckNativeContourFormatTopology doc
    TestCheckNativeConcreteComponentExport doc
    TestCheckNativeSavedContours doc, reader
    TestCheckNativeMixedContourRoundTrip doc
    TestCheckNativeSharedContours doc
    path = ThisWorkbook.Path & "\ContourFormat_" & Format$(Now, "yyyymmdd_hhnnss") & ".dwg"
    If Len(Dir$(path)) > 0 Then Err.Raise vbObjectError + 5502, , "Тестовый DWG уже существует."
    doc.SaveAs path
    mReport = mReport & "NATIVE_DWG: " & path & vbCrLf
    GoTo Cleanup
Failed:
    TestCheck "native.format.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Cleanup:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    On Error GoTo 0
    mReport = mReport & "TOTAL_REAL_AUTOCAD_CONTOURS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
    mNativeProgress = False
    RunRealAutoCADContourFormatTests = mReport
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Адресная native-проверка генерации по контуру без остальных CAD-пакетов.
Public Function RunRealAutoCADContourMeshTests(ByVal ownedApplication As Object, ByVal templatePath As String) As String
    RunRealAutoCADContourMeshTests = modTestContourMesh.RunNativeContourMeshTests(ownedApplication, templatePath)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' настоящий импорт одного сечения у начала координат и в 1000 м.
' Проверяет два отверстия, неизменные исходные координаты, НДС и трещины.
Public Function RunRealAutoCADFarOriginTests(ByVal ownedApplication As Object, _
        ByVal templatePath As String, Optional ByVal capacityOnly As Boolean = False) As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    mNativeProgress = True
    Dim doc As Object, fixture As Object, sheet As Object, entity As Object, region As Object
    Dim settings As CSystemSettingsReader, units As CUnitSystem, importer As CAutoCADSectionModelImporter
    Dim model As CSectionModel, baseline As CSectionModel, provider As CMaterialModelProvider
    Dim c As CMaterialDiagram, s As CMaterialDiagram, spec As CMaterialModelSpec
    Dim frame As CSectionNumericalFrame, solver As CSectionSolver, runner As CStateSolutionRunner
    Dim formation As CCrackFormationCalculator, crack As CCrackFormationResult, load As CSectionLoadState
    Dim keys As Variant, values As Variant, shifts As Variant, position As Long, i As Long, testCase As Long
    Dim curves(0 To 0) As Object, regions As Variant, center(0 To 2) As Double
    Dim sx As Double, sy As Double, targets As Variant, target As Variant, n As Double, mx As Double, my As Double
    Dim before As Long, reference(1 To 6, 1 To 5) As Double, statuses(1 To 6) As String
    Dim prefix As String, x As Variant, y As Variant, writer As CNDMResultsWriter, restored As CSectionModel
    Dim concrete As CConcreteMaterialParameters, steel As CSteelMaterialParameters
    Dim capacity As CCapacitySolver, capacityMethod As Variant, methodIndex As Long
    Dim capacityStates(1 To 6, 1 To 2) As String
    On Error GoTo Failed
    TestNativeProgress "far: start"
    If ownedApplication Is Nothing Then Err.Raise vbObjectError + 5502, , "Требуется собственный тестовый AutoCAD."
    TestCheck "far.actualAutoCAD", InStr(1, CStr(ownedApplication.FullName), "\AutoCAD 2023\acad.exe", vbTextCompare) > 0
    If Len(templatePath) > 0 Then Set doc = ownedApplication.Documents.Add(templatePath) Else Set doc = ownedApplication.Documents.Add
    TestWaitForCAD ownedApplication
    TestCheck "far.WCS", CLng(doc.GetVariable("WORLDUCS")) = 1
    Set fixture = Application.Workbooks.Add(-4167): Set sheet = fixture.Worksheets(1)
    keys = Array("AutoCAD.Common.ConcreteLayer", "AutoCAD.Common.RebarLayer", "AutoCAD.Common.SectionContourLayer", _
        "AutoCAD.Common.OpeningContourLayer", "Mesh.MinElementArea", "Mesh.StepX", "Mesh.StepY", "Mesh.BoundarySubdivisions", "Mesh.StepMode", "Mesh.CellGeometry")
    values = Array("FAR_EMPTY", "FAR_REBAR", "FAR_OUTER", "FAR_OPENING", 0#, 100#, 100#, 2, "User", "Adaptive")
    sheet.Cells(1, 1).Value2 = "Key": sheet.Cells(1, 2).Value2 = "Value": sheet.Cells(1, 3).Value2 = "Unit"
    For i = 0 To UBound(keys)
        sheet.Cells(i + 2, 1).Value2 = keys(i): sheet.Cells(i + 2, 2).Value2 = values(i): sheet.Cells(i + 2, 3).Value2 = "-"
        If i < 4 Then doc.Layers.Add CStr(values(i))
    Next i
    Set settings = New CSystemSettingsReader: settings.LoadFromRange sheet.Range("A1").Resize(UBound(keys) + 2, 3) ' Включает весь фактический набор ключей вместе с заголовком, включая Mesh.StepMode.
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set importer = New CAutoCADSectionModelImporter
    Set concrete = New CConcreteMaterialParameters: concrete.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    Set steel = New CSteelMaterialParameters: steel.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#
    Set provider = New CMaterialModelProvider: provider.TestInitializeFromParameters concrete, steel
    Set spec = New CMaterialModelSpec: spec.Initialize "SLS(II)", "ThreeLine", "UseDiagram", "TwoLine"
    Set c = provider.ConcreteMaterialForEquilibriumFromSpec(spec): Set s = provider.SteelMaterialForEquilibriumFromSpec(spec)
    If capacityOnly Then
        Set c = provider.ConcreteMaterial(cpStrength): Set s = provider.SteelMaterial(cpStrength)
    End If
    shifts = Array(Array(0#, 0#), Array(1000000#, -1000000#))
    targets = Array(Array(-500# * 9806.65, 0#, 0#), Array(-1000# * 9806.65, 0#, 0#), _
        Array(-200000#, 20000000#, 5000000#), Array(-200000#, -5000000#, 20000000#), _
        Array(-100000#, -15000000#, -15000000#), Array(50000#, 10000000#, -2000000#))
    For position = 0 To 1
        ' Только собственные объекты этого тестового документа.
        For i = doc.ModelSpace.Count - 1 To 0 Step -1: doc.ModelSpace.Item(i).Delete: Next i
        sx = CDbl(shifts(position)(0)): sy = CDbl(shifts(position)(1))
        Set entity = TestNativeRectangle(doc, 600#, 1000#, 0#, sx, sy): entity.Layer = "FAR_OUTER"
        For Each x In Array(-125#, 125#)
            center(0) = sx + CDbl(x): center(1) = sy
            Set entity = doc.ModelSpace.AddCircle(center, 50#): entity.Layer = "FAR_OPENING"
        Next x
        For Each x In Array(-250#, 250#)
            For Each y In Array(-450#, 0#, 450#)
                center(0) = sx + CDbl(x): center(1) = sy + CDbl(y)
                Set entity = doc.ModelSpace.AddCircle(center, 16#)
                Set curves(0) = entity: regions = doc.ModelSpace.AddRegion(curves)
                Set region = regions(LBound(regions)): region.Layer = "FAR_REBAR": entity.Delete
            Next y
        Next x
        before = doc.ModelSpace.Count
        TestNativeProgress "far: import position=" & CStr(position)
        Set model = importer.ImportFromDocument(doc, settings, units)
        TestNativeProgress "far: imported fibers=" & CStr(model.ConcreteCount)
        prefix = "far." & CStr(position)
        TestCheck prefix & ".sourcePreserved", doc.ModelSpace.Count = before
        TestCheck prefix & ".imported", model.SourceType = "AutoCADImport" And model.RebarCount = 6 And model.ConcreteCount > 0
        TestCheck prefix & ".contours", model.Contours.Count = 12
        If position = 0 Then
            Set baseline = model
        Else
            TestCheck prefix & ".meshCount", model.ConcreteCount = baseline.ConcreteCount
            For i = 1 To model.ConcreteCount
                TestCheckCloseFar prefix & ".fiberX." & CStr(i), model.ConcreteX(i) - sx, baseline.ConcreteX(i), 0.000001
                TestCheckCloseFar prefix & ".fiberY." & CStr(i), model.ConcreteY(i) - sy, baseline.ConcreteY(i), 0.000001
                TestCheckCloseFar prefix & ".fiberArea." & CStr(i), model.ConcreteArea(i), baseline.ConcreteArea(i), 0.000001
            Next i
        End If
        Set frame = New CSectionNumericalFrame: frame.Initialize model
        testCase = 0
        For Each target In targets
            testCase = testCase + 1: n = CDbl(target(0)): mx = CDbl(target(1)): my = CDbl(target(2))
            prefix = "far." & CStr(position) & ".load" & CStr(testCase)
            If capacityOnly Then
                methodIndex = 0
                For Each capacityMethod In Array("Auto", "LoadMultiplier")
                    methodIndex = methodIndex + 1
                    prefix = "far." & CStr(position) & ".load" & CStr(testCase) & "." & CStr(capacityMethod)
                    TestNativeProgress prefix & ": capacity start"
                    Set capacity = New CCapacitySolver
                    capacity.ConcreteCompressionLimit = c.PhysicalCompressionStrain
                    capacity.ConcreteTensionLimitEnabled = False
                    capacity.SteelCompressionLimit = Abs(s.PhysicalCompressionStrain)
                    capacity.SteelTensionLimit = s.PhysicalTensionStrain
                    capacity.StrainTolerance = 0.00000001: capacity.LambdaTolerance = 0.0001
                    capacity.SolverToleranceN = 0.1: capacity.SolverToleranceMx = 10#: capacity.SolverToleranceMy = 10#
                    capacity.SolverMaxIterations = 100: capacity.MaxRetries = 0: capacity.SolverBaseLoadSteps = 1
                    If CStr(capacityMethod) = "Auto" Then
                        capacity.SolveByAutoLoadPath model, c, s, 0#, n, 0#, mx + n * frame.OriginY, 0#, my + n * frame.OriginX
                    Else
                        capacity.SolveByLoadPathMultiplier model, c, s, 0#, n, 0#, mx + n * frame.OriginY, 0#, my + n * frame.OriginX
                    End If
                    TestCheck prefix & ".converged", capacity.Converged
                    If Not capacity.Converged Then Err.Raise vbObjectError + 5502, , prefix & ": " & capacity.StopReason
                    TestCheckCloseFar prefix & ".N", capacity.LastSolver.Nint, capacity.LambdaUltimate * n, 0.2
                    TestCheckCloseFar prefix & ".Mx", frame.CenterMx(capacity.LastSolver.Nint, capacity.LastSolver.Mxint), capacity.LambdaUltimate * mx, 20#
                    TestCheckCloseFar prefix & ".My", frame.CenterMy(capacity.LastSolver.Nint, capacity.LastSolver.Myint), capacity.LambdaUltimate * my, 20#
                    TestCheckCloseFar prefix & ".physicalLimit", capacity.TestCriticalStrainUtilization, 1#, 0.00001
                    If testCase <= 2 Then TestCheck prefix & ".axialStrategy", capacity.SolutionMethod = "LoadMultiplier"
                    If position = 0 Then
                        reference(testCase, methodIndex) = capacity.LambdaUltimate
                        capacityStates(testCase, methodIndex) = capacity.LimitState
                    Else
                        TestCheckCloseFar prefix & ".sameCapacity", capacity.LambdaUltimate, reference(testCase, methodIndex), 0.00001
                        TestCheck prefix & ".sameCriterion", capacity.LimitState = capacityStates(testCase, methodIndex)
                    End If
                    mReport = mReport & "FAR_CAPACITY: " & prefix & "; lambda=" & CStr(capacity.LambdaUltimate) & _
                        "; criterion=" & capacity.LimitState & "; comment=" & capacity.StopReason & vbCrLf
                Next capacityMethod
                GoTo NextFarLoad
            End If
            TestNativeProgress prefix & ": equilibrium start"
            Set runner = New CStateSolutionRunner: runner.LoadSteps = 8: runner.MaxIterations = 100
            runner.ToleranceN = 0.1: runner.ToleranceMx = 10#: runner.ToleranceMy = 10#
            runner.Solve model, c, s, n, mx + n * frame.OriginY, my + n * frame.OriginX
            Set solver = runner.ResultSolver
            TestCheck prefix & ".equilibrium", solver.Converged
            If Not solver.Converged Then Err.Raise vbObjectError + 5502, , prefix & ": " & solver.StopReason
            TestCheckCloseFar prefix & ".N", solver.Nint, n, 0.2
            TestCheckCloseFar prefix & ".Mx", frame.CenterMx(solver.Nint, solver.Mxint), mx, 20#
            TestCheckCloseFar prefix & ".My", frame.CenterMy(solver.Nint, solver.Myint), my, 20#
            Set load = New CSectionLoadState: load.Initialize n, mx, my, frame.OriginX, frame.OriginY
            Set formation = New CCrackFormationCalculator
            formation.SolverToleranceN = 0.1: formation.SolverToleranceMx = 10#: formation.SolverToleranceMy = 10#
            formation.SolverMaxIterations = 100: formation.SolverLoadSteps = 8
            Dim crackedSpec As CMaterialModelSpec
            Set crackedSpec = New CMaterialModelSpec: crackedSpec.Initialize "SLS(II)", "TwoLine", "Ignore", "TwoLine"
            TestNativeProgress prefix & ": formation start"
            Set crack = formation.CheckFormation(model, provider, crackedSpec, spec, n, load.InternalMx, load.InternalMy, _
                load, frame.OriginX, frame.OriginY)
            TestCheck prefix & ".noNumericalFailure", crack.ResultMeta.InternalStatus <> rsNumericalFailure
            TestNativeProgress prefix & ": formation finished"
            If position = 0 Then
                reference(testCase, 1) = solver.StrainAt(frame.OriginX, frame.OriginY)
                reference(testCase, 2) = solver.KappaX: reference(testCase, 3) = solver.KappaY
                reference(testCase, 4) = crack.LambdaCrc: reference(testCase, 5) = solver.Epsilon0
                statuses(testCase) = CStr(crack.ResultMeta.InternalStatus) & ":" & CStr(crack.CrackFormed)
            Else
                TestCheckCloseFar prefix & ".centerStrain", solver.StrainAt(frame.OriginX, frame.OriginY), reference(testCase, 1), 0.000000001
                TestCheckCloseFar prefix & ".kx", solver.KappaX, reference(testCase, 2), 0.00000000001
                TestCheckCloseFar prefix & ".ky", solver.KappaY, reference(testCase, 3), 0.00000000001
                TestCheckCloseFar prefix & ".lambdaCrc", crack.LambdaCrc, reference(testCase, 4), 0.00001
                TestCheckCloseFar prefix & ".originalEpsilon0", solver.Epsilon0, _
                    reference(testCase, 5) - solver.KappaX * sy - solver.KappaY * sx, 0.00000001
                TestCheck prefix & ".sameCrackStatus", CStr(crack.ResultMeta.InternalStatus) & ":" & CStr(crack.CrackFormed) = statuses(testCase)
            End If
            mReport = mReport & "FAR_LOAD: " & prefix & "; N=" & CStr(n) & "; Mx=" & CStr(mx) & "; My=" & CStr(my) & _
                "; formed=" & CStr(crack.CrackFormed) & "; lambda=" & CStr(crack.LambdaCrc) & "; comment=" & crack.ResultMeta.ResultComment & vbCrLf
NextFarLoad:
        Next target
        Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, model, PrepareSectionSnapshot(model), units
        Set restored = ReadSectionGeometryFromResults(ThisWorkbook)
        TestCheckPostAudit03ContoursEqual prefix & ".savedContours", model.Contours, restored.Contours
        TestCheckCloseFar prefix & ".savedOriginalX", restored.ConcreteX(1), model.ConcreteX(1), 0.000001
        TestCheckCloseFar prefix & ".savedOriginalY", restored.ConcreteY(1), model.ConcreteY(1), 0.000001
    Next position
    doc.SaveAs ThisWorkbook.Path & "\FarOrigin_1000m.dwg"
    GoTo Cleanup
Failed:
    TestCheck "far.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Cleanup:
    mNativeProgress = False
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
    RunRealAutoCADFarOriginTests = mReport & "TOTAL_REAL_AUTOCAD_FAR_ORIGIN: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' несущая того же настоящего импорта без повторения трещинообразования.
Public Function RunRealAutoCADFarCapacityTests(ByVal ownedApplication As Object, ByVal templatePath As String) As String
    RunRealAutoCADFarCapacityTests = RunRealAutoCADFarOriginTests(ownedApplication, templatePath, True)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает координату с явно переданным допуском, сохраняя величину расхождения в имени проверки.
Private Sub TestCheckCloseFar(ByVal name As String, ByVal actual As Double, ByVal expected As Double, ByVal tolerance As Double)
    TestCheck name & "; difference=" & CStr(Abs(actual - expected)), Abs(actual - expected) <= tolerance
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет настоящий AutoCAD в собственном новом документе.
' Существующие чертежи не меняются. После копирования и разбиения обычным
' классом чтения исходные объекты сохраняют номера, площади и количество.
Public Function RunRealAutoCADContourTests(Optional ByVal ownedApplication As Object = Nothing, _
        Optional ByVal templatePath As String = "") As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    Dim acad As Object, doc As Object, previous As Object, polyline As Object, source As Object
    Dim circleEntity As Object, openingRegion As Object, reader As CAutoCADContourReader, query As CSectionGeometryQuery
    Dim region As CConcreteRegion, curves(0 To 0) As Object, regions As Variant
    Dim countBefore As Long, handleBefore As String, areaBefore As Double, mode As Long, i As Long, openings As Long
    Dim path As String, center(0 To 2) As Double, semi(0 To 3) As Double
    On Error GoTo Failed
    mNativeProgress = True
    TestNativeProgress "start"
    If ownedApplication Is Nothing Then
        Set acad = GetObject(, "AutoCAD.Application.24.2")
    Else
        Set acad = ownedApplication
    End If
    TestCheck "native.actualAutoCAD", InStr(1, CStr(acad.FullName), "\AutoCAD 2023\acad.exe", vbTextCompare) > 0
    If InStr(1, CStr(acad.FullName), "\AutoCAD 2023\acad.exe", vbTextCompare) = 0 Then Err.Raise vbObjectError + 5502, , "Тест требует Autodesk AutoCAD 2023, а не SOFiPLUS."
    If acad.Documents.Count > 0 Then Set previous = acad.ActiveDocument
    TestNativeProgress "creating document"
    If Len(templatePath) > 0 Then Set doc = acad.Documents.Add(templatePath) Else Set doc = acad.Documents.Add
    TestNativeProgress "document created"
    TestWaitForCAD acad
    TestNativeProgress "document ready"
    Set reader = New CAutoCADContourReader: Set query = New CSectionGeometryQuery
    Set polyline = TestNativeRectangle(doc, 100#, 80#, 0.37, 137#, -73#)
    Set curves(0) = polyline: regions = doc.ModelSpace.AddRegion(curves): Set source = regions(LBound(regions))
    For mode = 0 To 1
        If mode = 0 Then Set circleEntity = polyline Else Set circleEntity = source
        countBefore = doc.ModelSpace.Count: handleBefore = CStr(circleEntity.Handle): areaBefore = CDbl(circleEntity.Area)
        Set region = reader.ReadRegion(circleEntity)
        TestCheckArea "native.rectangle." & CStr(mode), region.Area, 8000#
        TestCheck "native.rectangle.source." & CStr(mode), CStr(circleEntity.Handle) = handleBefore And doc.ModelSpace.Count = countBefore
        TestCheckArea "native.rectangle.unchanged." & CStr(mode), CDbl(circleEntity.Area), areaBefore
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
        TestCheckArea "native.semicircle." & CStr(mode), region.Area, 1250# * GEOM_PI
        TestCheck "native.semicircle.analytic." & CStr(mode), region.SegmentCount = 2
        TestCheck "native.semicircle.cleanup." & CStr(mode), doc.ModelSpace.Count = countBefore
    Next mode
    Set polyline = TestNativeRectangle(doc, 100#, 100#, 0#, 250#, 100#)
    Set curves(0) = polyline: regions = doc.ModelSpace.AddRegion(curves): Set source = regions(LBound(regions))
    For mode = 0 To 1
        center(0) = 225# + 45# * mode: center(1) = 85# + 25# * mode
        Set circleEntity = doc.ModelSpace.AddCircle(center, 10# - 5# * mode)
        Set curves(0) = circleEntity: regions = doc.ModelSpace.AddRegion(curves): Set openingRegion = regions(LBound(regions))
        source.Boolean 2, openingRegion
    Next mode
    countBefore = doc.ModelSpace.Count: handleBefore = CStr(source.Handle): areaBefore = CDbl(source.Area)
    Set region = reader.ReadRegion(source)
    TestCheckArea "native.twoOpenings.area", region.Area, 10000# - 125# * GEOM_PI
    TestCheck "native.twoOpenings.loops", region.LoopCount = 3
    For i = 1 To region.LoopCount
        If query.LoopIsOpening(region, i) Then openings = openings + 1
    Next i
    TestCheck "native.twoOpenings.roles", openings = 2
    TestCheck "native.twoOpenings.excluded", Not query.ContainsPoint(region, 225#, 85#) And Not query.ContainsPoint(region, 270#, 110#)
    TestCheck "native.twoOpenings.material", query.ContainsPoint(region, 250#, 100#)
    TestCheck "native.twoOpenings.source", CStr(source.Handle) = handleBefore And doc.ModelSpace.Count = countBefore
    TestCheckArea "native.twoOpenings.sourceArea", CDbl(source.Area), areaBefore
    TestCheckNativeSeparateOpenings doc, reader, query
    TestCheckNativeConcreteComponentExport doc
    TestCheckNativeSavedContours doc, reader
    TestCheckNativeSharedContours doc
    TestCheckNativeRegionBoundaryDedup doc
    TestCheckNativeMixedContourRoundTrip doc
    TestCheckNativeContourFormatTopology doc
    TestCheckNativePhysicalPolylineClosure doc
    TestCheckNativeEmptyOpeningSegment doc
    TestCheckNativeWorldCoordinateGuard doc
    path = ThisWorkbook.Path & "\SP35_ContourImport_" & Format$(Now, "yyyymmdd_hhnnss") & ".dwg"
    If Len(Dir$(path)) > 0 Then Err.Raise vbObjectError + 5502, , "Тестовый DWG уже существует; перезапись запрещена."
    doc.SaveAs path
    mReport = mReport & "NATIVE_DWG: " & path & vbCrLf
    GoTo Cleanup
Failed:
    TestCheck "native.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Cleanup:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    If Not previous Is Nothing Then previous.Activate
    On Error GoTo 0
    mReport = mReport & "TOTAL_REAL_AUTOCAD_CONTOURS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
    mNativeProgress = False
    RunRealAutoCADContourTests = mReport
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отказ по ПСК и невозможности прочитать WORLDUCS происходит
' до обращения к ModelSpace/слоям. Строки, Null и Boolean не заменяют код 1.
Private Sub TestCheckWorldCoordinateGuard()
    Dim document As CTestFakeAcadDocument, settings As CSystemSettingsReader
    Dim importer As CAutoCADSectionModelImporter, model As CSectionModel, value As Variant
    Dim number As Long, description As String, mode As Long, expectedError As Long
    Set document = New CTestFakeAcadDocument: Set settings = New CSystemSettingsReader
    Set importer = New CAutoCADSectionModelImporter
    For Each value In Array(0, Empty, Null, "1", 2, True, 0.5, CVErr(2015))
        mode = mode + 1: document.TestInitialize value
        On Error Resume Next
        Set model = importer.ImportFromDocument(document, settings)
        number = Err.Number: description = Err.Description: Err.Clear
        On Error GoTo 0
        expectedError = vbObjectError + 4408
        If mode = 1 Then expectedError = vbObjectError + 4407
        TestCheck "worldUCS.reject." & CStr(mode), number = expectedError And model Is Nothing
        TestCheck "worldUCS.noGeometryRead." & CStr(mode), document.TestGeometryReads = 0 And document.TestVariableReads = 1
        TestCheck "worldUCS.instruction." & CStr(mode), InStr(description, "_UCS") > 0 And InStr(description, "_World") > 0
    Next value
    document.TestInitialize 1, True
    On Error Resume Next
    Set model = importer.ImportFromDocument(document, settings): number = Err.Number: Err.Clear
    On Error GoTo 0
    TestCheck "worldUCS.unavailableRejected", number = vbObjectError + 4408 And document.TestGeometryReads = 0
    document.TestInitialize 1
    On Error Resume Next
    Set model = importer.ImportFromDocument(document, settings): number = Err.Number: Err.Clear
    On Error GoTo 0
    TestCheck "worldUCS.worldAcceptedBeforeGeometry", number = vbObjectError + 9106 And document.TestGeometryReads = 1 And document.TestVariableReads = 1
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' в собственном CAD-документе проверяет отказ смещенной,
' повернутой и смещенно-повернутой ПСК, затем импорт двух точных отверстий
' без outer в МСК. Координаты, исходники, настройки и текущая ПСК сохраняются.
Private Sub TestCheckNativeWorldCoordinateGuard(ByVal doc As Object)
    Dim fixture As Object, sheet As Object, settings As CSystemSettingsReader
    Dim importer As CAutoCADSectionModelImporter, section As CSectionModel, restored As CSectionModel
    Dim world As Object, localUCS As Object, polyline As Object, mesh As Object, bar As Object, hole As Object
    Dim origin(0 To 2) As Double, xAxis(0 To 2) As Double, yAxis(0 To 2) As Double, center(0 To 2) As Double
    Dim curves(0 To 0) As Object, regions As Variant, keys As Variant, values As Variant
    Dim mode As Long, i As Long, countBefore As Long, number As Long, description As String
    Dim failure As Long, failureDescription As String, query As CSectionGeometryQuery, region As CConcreteRegion
    Dim writer As CNDMResultsWriter, low As Double, high As Double, intervals As Variant
    On Error GoTo Failed
    TestNativeProgress "world coordinate guard"
    xAxis(0) = 1#: yAxis(1) = 1#
    TestNativeProgress "creating world UCS"
    Set world = doc.UserCoordinateSystems.Add(origin, xAxis, yAxis, "NDM_TEST_WORLD")
    TestNativeProgress "activating world UCS"
    doc.ActiveUCS = world
    TestCheck "native.worldUCS.initialWorld", CLng(doc.GetVariable("WORLDUCS")) = 1
    Set fixture = Application.Workbooks.Add(-4167): Set sheet = fixture.Worksheets(1)
    keys = Array("AutoCAD.Common.ConcreteLayer", "AutoCAD.Common.RebarLayer", "AutoCAD.Common.SectionContourLayer", _
        "AutoCAD.Common.OpeningContourLayer", "Mesh.MinElementArea", "Mesh.CellGeometry")
    values = Array("NDM_WCS_MESH", "NDM_WCS_REBAR", "NDM_WCS_OUTER", "NDM_WCS_OPENING", 0#, "Adaptive")
    sheet.Cells(1, 1).Value2 = "Key": sheet.Cells(1, 2).Value2 = "Value": sheet.Cells(1, 3).Value2 = "Unit"
    For i = 0 To UBound(keys)
        sheet.Cells(i + 2, 1).Value2 = keys(i): sheet.Cells(i + 2, 2).Value2 = values(i): sheet.Cells(i + 2, 3).Value2 = "-"
        If i < 4 Then doc.Layers.Add CStr(values(i))
    Next i
    Set settings = New CSystemSettingsReader: settings.LoadFromRange sheet.Range("A1:C7")
    Set importer = New CAutoCADSectionModelImporter
    Set polyline = TestNativeRectangle(doc, 100#, 80#, 0#, 5000#, 5000#)
    Set curves(0) = polyline: regions = doc.ModelSpace.AddRegion(curves): Set mesh = regions(LBound(regions))
    mesh.Layer = "NDM_WCS_MESH": polyline.Delete
    center(0) = 4970#: center(1) = 4980#
    Set polyline = doc.ModelSpace.AddCircle(center, 2#)
    Set curves(0) = polyline: regions = doc.ModelSpace.AddRegion(curves): Set bar = regions(LBound(regions))
    bar.Layer = "NDM_WCS_REBAR": polyline.Delete
    For i = 0 To 1
        center(0) = 5000# + 25# * i: center(1) = 5000#
        Set hole = doc.ModelSpace.AddCircle(center, 10# - 5# * i): hole.Layer = "NDM_WCS_OPENING"
    Next i
    countBefore = doc.ModelSpace.Count
    For mode = 0 To 2
        origin(0) = 0#: origin(1) = 0#: xAxis(0) = 1#: xAxis(1) = 0#: yAxis(0) = 0#: yAxis(1) = 1#
        If mode <> 1 Then origin(0) = 100#: origin(1) = -200#
        xAxis(0) = origin(0) + 1#: xAxis(1) = origin(1): yAxis(0) = origin(0): yAxis(1) = origin(1) + 1#
        If mode > 0 Then
            xAxis(0) = origin(0) + Cos(0.37): xAxis(1) = origin(1) + Sin(0.37)
            yAxis(0) = origin(0) - Sin(0.37): yAxis(1) = origin(1) + Cos(0.37)
        End If
        Set localUCS = doc.UserCoordinateSystems.Add(origin, xAxis, yAxis, "NDM_TEST_LOCAL_" & CStr(mode))
        doc.ActiveUCS = localUCS
        TestCheck "native.worldUCS.localFlag." & CStr(mode), CLng(doc.GetVariable("WORLDUCS")) = 0
        On Error Resume Next
        Set section = importer.ImportFromDocument(doc, settings)
        number = Err.Number: description = Err.Description: Err.Clear
        On Error GoTo Failed
        TestCheck "native.worldUCS.localRejected." & CStr(mode), number = vbObjectError + 4407 And section Is Nothing
        TestCheck "native.worldUCS.instruction." & CStr(mode), InStr(description, "_UCS") > 0 And InStr(description, "_World") > 0
        TestCheck "native.worldUCS.noMutation." & CStr(mode), doc.ModelSpace.Count = countBefore And CStr(doc.GetVariable("UCSNAME")) = localUCS.Name
    Next mode
    doc.ActiveUCS = world
    Set section = importer.ImportFromDocument(doc, settings)
    TestCheck "native.worldUCS.worldImported", section.ConcreteCount = 1 And section.RebarCount = 1
    TestCheckArea "native.worldUCS.coordinatesX", section.ConcreteX(1), 5000#
    TestCheckArea "native.worldUCS.coordinatesY", section.ConcreteY(1), 5000#
    TestCheck "native.worldUCS.sourcesKept", doc.ModelSpace.Count = countBefore And CLng(doc.GetVariable("WORLDUCS")) = 1
    Set query = New CSectionGeometryQuery: query.Initialize section: Set region = query.ConcreteDomain
    TestCheck "native.openingOnly.source", region.Source = "ElementBoundaryLoops+KnownOpening"
    TestCheck "native.openingOnly.twoExactHoles", region.LoopCount = 3 And region.SegmentCount = 12
    TestCheckArea "native.openingOnly.analyticArea", region.Area, 8000# - 125# * GEOM_PI
    TestCheck "native.openingOnly.holesExcluded", Not query.ContainsPoint(region, 5000#, 5000#) And Not query.ContainsPoint(region, 5025#, 5000#)
    query.ProjectionBounds region, 1#, 0#, low, high
    TestCheck "native.openingOnly.meshOuterBounds", Abs(low - 4950#) < 0.000001 And Abs(high - 5050#) < 0.000001
    intervals = query.LineIntervals(region, 5000#, 5000#, 1#, 0#)
    TestCheck "native.openingOnly.exactHoleIntersections", UBound(intervals, 1) = 3 And _
        Abs(intervals(1, 2) + 10#) < 0.000001 And Abs(intervals(2, 1) - 10#) < 0.000001 And _
        Abs(intervals(2, 2) - 20#) < 0.000001 And Abs(intervals(3, 1) - 30#) < 0.000001
    For i = 1 To section.Contours.Count
        TestCheck "native.openingOnly.explicitRole." & CStr(i), section.Contours.LoopRole(i) = "Opening"
    Next i
    Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
    Set restored = ReadSectionGeometryFromResults(ThisWorkbook, "AutoCADImport")
    TestCheckPostAudit03ContoursEqual "native.openingOnly.exactSaved", section.Contours, restored.Contours
    query.Initialize restored
    TestCheckArea "native.openingOnly.restoredArea", query.ConcreteDomain.Area, 8000# - 125# * GEOM_PI
    TestCheck "native.openingOnly.restoredSource", query.ConcreteDomain.Source = "ElementBoundaryLoops+KnownOpening"
    GoTo Cleanup
Failed:
    failure = Err.Number: failureDescription = Err.Description
Cleanup:
    On Error Resume Next
    If Not world Is Nothing Then doc.ActiveUCS = world
    If Not fixture Is Nothing Then fixture.Close False
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, "CheckNativeWorldCoordinateGuard", failureDescription
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' фактическое замыкание важнее флага Closed. Повторные прямые
' вершины не меняют область, но вырожденные Region и дуги остаются ошибками.
Private Sub TestCheckPhysicalPolylineClosure()
    Dim reader As CAutoCADContourReader, entity As CTestFakeAcadContour, curve As CTestFakeAcadContour
    Dim region As CConcreteRegion, edges As Collection, kind As Variant, mode As Long, prefix As String
    Dim number As Long, layerName As Variant
    Set reader = New CAutoCADContourReader
    For Each layerName In Array("OUTER", "OPENING")
    For Each kind In Array("AcDbPolyline", "AcDb2dPolyline")
        For mode = 0 To 2
            Set edges = TestRectangleEdges(0#, 0#, 100#, 80#)
            If mode > 0 Then
                Set curve = New CTestFakeAcadContour: curve.TestInitializeLine 0#, 0#, 0#, 0#: edges.Add curve
                Set curve = New CTestFakeAcadContour: curve.TestInitializeLine 100#, 0#, 100#, 0#: edges.Add curve
            End If
            Set entity = New CTestFakeAcadContour
            entity.TestInitializeLoop CStr(kind), CStr(layerName), "PHYSICAL", edges, (mode = 2)
            Set region = reader.ReadRegion(entity)
            prefix = "physicalPolyline." & CStr(layerName) & "." & CStr(kind) & "." & CStr(mode)
            TestCheckArea prefix & ".area", region.Area, 8000#
            TestCheck prefix & ".edges", region.SegmentCount = 4
            TestCheck prefix & ".sourceUnchanged", Not entity.TestDeleted And entity.Closed = (mode = 2)
            TestCheck prefix & ".copyDeleted", entity.TestLastCopy.TestDeleted
            For Each curve In entity.TestLastCopy.TestExplodedPieces: TestCheck prefix & ".pieceDeleted", curve.TestDeleted: Next curve
        Next mode
    Next kind
    Next layerName
    For mode = 0 To 2
        Set edges = TestRectangleEdges(0#, 0#, 100#, 80#)
        Set curve = New CTestFakeAcadContour
        If mode < 2 Then curve.TestInitializeLine 0#, 0#, 0#, 0# Else curve.TestInitializeArc 0#, 0#, 0#, 0#, GEOM_PI
        edges.Add curve
        Set entity = TestContour("AcDbRegion", edges)
        If mode = 1 Then
            Set edges = New Collection
            entity.TestInitializeLoop "AcDbPolyline", "OPENING", "EMPTY", edges, True
        End If
        If mode = 2 Then entity.TestInitializeLoop "AcDbPolyline", "OPENING", "BAD_ARC", edges, True
        On Error Resume Next
        Set region = reader.ReadRegion(entity): number = Err.Number: Err.Clear
        On Error GoTo 0
        TestCheck "physicalPolyline.invalidNotHidden." & CStr(mode), number <> 0
        TestCheck "physicalPolyline.invalidCleanup." & CStr(mode), Not entity.TestDeleted And entity.TestLastCopy.TestDeleted
    Next mode
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отдельный нулевой отрезок на каждом из контурных слоев не блокирует
' правильную полилинию. Ненулевой отрезок остается адресной ошибкой ввода.
Private Sub TestCheckEmptyStandaloneContourSegments()
    Dim space As Collection, layers As Collection, importer As CAutoCADSectionModelImporter, section As CSectionModel
    Dim zero As CTestFakeAcadContour, opening As CTestFakeAcadContour, layer As CTestFakeAcadContour, cell As CTestFakeAcadRegion
    Dim query As CSectionGeometryQuery, number As Long, description As String, layerName As Variant, prefix As String
    Set space = New Collection: Set layers = New Collection
    Set cell = New CTestFakeAcadRegion: cell.TestInitialize 10000#, 50#, 50#, 1000000#, 1200000#, 0#, "CONCRETE", "C1": space.Add cell
    Set cell = New CTestFakeAcadRegion: cell.TestInitialize 10#, 10#, 10#, 100#, 100#, 0#, "REBAR", "R1": space.Add cell
    space.Add TestContour("AcDbPolyline", TestRectangleEdges(0#, 0#, 100#, 100#))
    Set zero = New CTestFakeAcadContour: zero.TestInitializeLoop "AcDbLine", "OPENING", "ZERO_LINE", Nothing
    zero.TestInitializeLine 40#, 40#, 40#, 40#: space.Add zero
    Set opening = TestContour("AcDbPolyline", TestRectangleEdges(40#, 40#, 20#, 20#), "OPENING"): space.Add opening
    Set layer = New CTestFakeAcadContour: layers.Add layer, "OUTER": layers.Add layer, "OPENING"
    Set importer = New CAutoCADSectionModelImporter
    For Each layerName In Array("OUTER", "OPENING")
    zero.TestInitializeLoop "AcDbLine", CStr(layerName), "ZERO_LINE", Nothing
    zero.TestInitializeLine 40#, 40#, 40#, 40#
    prefix = "emptyContourSegment." & CStr(layerName)
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    Set query = New CSectionGeometryQuery: query.Initialize section
    TestCheckArea prefix & ".area", query.ConcreteDomain.Area, 9600#
    TestCheck prefix & ".sourceKept", Not zero.TestDeleted And Not opening.TestDeleted
    TestCheck prefix & ".warning", InStr(importer.ImportComment, "нулевой длины") > 0 And InStr(importer.ImportComment, "ZERO_LINE") > 0
    TestCheck prefix & ".typedWarning", importer.HasImportWarnings
    TestCheck prefix & ".shortMessage", InStr(importer.ResultMessage(section), "Наружных контуров: 1; отверстий: 1.") > 0 And _
        InStr(importer.ResultMessage(section), vbCrLf & vbCrLf & "Предупреждения:") > 0 And _
        InStr(importer.ResultMessage(section), "реальные линии/дуги") = 0
    zero.TestInitializeLine 40#, 40#, 41#, 40#
    On Error Resume Next
    Set section = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, Nothing, "OUTER", "OPENING", layers)
    number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo 0
    TestCheck prefix & ".nonzeroRejected", number <> 0 And InStr(description, "ZERO_LINE") > 0 And InStr(description, "PEDIT") > 0
    TestCheck prefix & ".nonzeroKept", Not zero.TestDeleted
    Next layerName
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' настоящий AutoCAD проверяет Closed=False с повтором первой
' вершины, Closed=True с таким же повтором, внутренний повтор и bulge-дугу.
' Повторное чтение и ошибочный импорт не оставляют фрагментов Explode.
Private Sub TestCheckNativePhysicalPolylineClosure(ByVal doc As Object)
    Dim reader As CAutoCADContourReader, entity As Object, region As CConcreteRegion
    Dim coordinates As Variant, values() As Double, stride As Long, kind As Long, mode As Long, i As Long, j As Long
    Dim countBefore As Long, number As Long, expectedArea As Double, prefix As String, description As String
    Set reader = New CAutoCADContourReader
    For kind = 0 To 1
        For mode = 0 To 4
            Select Case mode
                Case 0, 1: coordinates = Array(0#, 0#, 100#, 0#, 100#, 80#, 0#, 80#, 0#, 0#): expectedArea = 8000#
                Case 2: coordinates = Array(0#, 0#, 100#, 0#, 100#, 0#, 100#, 80#, 0#, 80#, 0#, 0#): expectedArea = 8000#
                Case 3, 4: coordinates = Array(0#, 0#, 20#, 0#, 0#, 0#): expectedArea = 50# * GEOM_PI
            End Select
            stride = 2 + kind
            ReDim values(0 To (UBound(coordinates) + 1) / 2 * stride - 1)
            For i = 0 To (UBound(coordinates) + 1) / 2 - 1
                values(i * stride) = coordinates(i * 2): values(i * stride + 1) = coordinates(i * 2 + 1)
            Next i
            If kind = 0 Then Set entity = doc.ModelSpace.AddLightWeightPolyline(values) Else Set entity = doc.ModelSpace.AddPolyline(values)
            entity.Closed = (mode = 1 Or mode = 2 Or mode = 4)
            If mode >= 3 Then entity.SetBulge 0, 1#
            countBefore = doc.ModelSpace.Count: prefix = "native.physicalPolyline." & CStr(kind) & "." & CStr(mode)
            For j = 1 To 3
                Set region = reader.ReadRegion(entity)
                TestCheckArea prefix & ".area." & CStr(j), region.Area, expectedArea
                TestCheck prefix & ".noLeakedEntities." & CStr(j), doc.ModelSpace.Count = countBefore
                TestCheck prefix & ".closedPreserved." & CStr(j), CBool(entity.Closed) = (mode = 1 Or mode = 2 Or mode = 4)
                If mode >= 3 Then TestCheckArea prefix & ".bulgePreserved." & CStr(j), entity.GetBulge(0), 1#
            Next j
        Next mode
    Next kind
    ReDim values(0 To 7): values(2) = 100#: values(4) = 100#: values(5) = 80#: values(7) = 80#
    Set entity = doc.ModelSpace.AddLightWeightPolyline(values): entity.Closed = False: countBefore = doc.ModelSpace.Count
    On Error Resume Next
    Set region = reader.ReadRegion(entity): number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo 0
    TestCheck "native.physicalPolyline.realGapRejected", number <> 0
    TestCheck "native.physicalPolyline.realGapCleanup", doc.ModelSpace.Count = countBefore And Not CBool(entity.Closed)
    ReDim values(0 To 3)
    Set entity = doc.ModelSpace.AddLightWeightPolyline(values): entity.Closed = True: countBefore = doc.ModelSpace.Count
    On Error Resume Next
    Set region = reader.ReadRegion(entity): number = Err.Number: Err.Clear
    On Error GoTo 0
    TestCheck "native.physicalPolyline.zeroAreaRejected", number <> 0
    TestCheck "native.physicalPolyline.zeroAreaCleanup", doc.ModelSpace.Count = countBefore
    Dim lockedLayer As Object
    Set lockedLayer = doc.Layers.Add("NP_LOCKED")
    Set entity = TestNativeRectangle(doc, 100#, 80#, 0#, 600#, 600#): entity.Layer = "NP_LOCKED"
    lockedLayer.Lock = True: countBefore = doc.ModelSpace.Count
    Set region = Nothing
    On Error Resume Next
    Set region = reader.ReadRegion(entity): number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo 0
    lockedLayer.Lock = False
    TestCheck "native.physicalPolyline.lockedNoLeaks", doc.ModelSpace.Count = countBefore
    TestCheck "native.physicalPolyline.lockedRejected", number <> 0 And InStr(description, "Разблокируйте") > 0
    Set region = reader.ReadRegion(entity)
    TestCheckArea "native.physicalPolyline.unlockedArea", region.Area, 8000#
    TestCheck "native.physicalPolyline.unlockedNoLeaks", doc.ModelSpace.Count = countBefore
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' воспроизводит видимые пользователю полилинию и отдельный
' пустой AcDbLine на каждом контурном слое. Полный importer сохраняет роли и область.
Private Sub TestCheckNativeEmptyOpeningSegment(ByVal doc As Object)
    Dim importer As CAutoCADSectionModelImporter, section As CSectionModel, query As CSectionGeometryQuery
    Dim entity As Object, outer As Object, zero As Object, curve As Object, regions As Variant, curves(0 To 0) As Object
    Dim point(0 To 2) As Double, endpoint(0 To 2) As Double, values(0 To 9) As Double
    Dim layer As Variant, targetLayer As Variant, countBefore As Long, number As Long, description As String, prefix As String
    For Each layer In Array("NP_CONCRETE", "NP_REBAR", "NP_OUTER", "NP_OPENING"): doc.Layers.Add CStr(layer): Next layer
    Set curve = TestNativeRectangle(doc, 100#, 100#, 0#, 50#, 50#)
    Set curves(0) = curve: regions = doc.ModelSpace.AddRegion(curves): Set entity = regions(LBound(regions)): entity.Layer = "NP_CONCRETE"
    point(0) = 10#: point(1) = 10#: Set curve = doc.ModelSpace.AddCircle(point, 2#)
    Set curves(0) = curve: regions = doc.ModelSpace.AddRegion(curves): Set entity = regions(LBound(regions)): entity.Layer = "NP_REBAR"
    values(2) = 100#: values(4) = 100#: values(5) = 100#: values(7) = 100#
    Set outer = doc.ModelSpace.AddLightWeightPolyline(values): outer.Layer = "NP_OUTER": outer.Closed = False
    values(0) = 40#: values(1) = 40#: values(2) = 60#: values(3) = 40#: values(4) = 60#: values(5) = 60#
    values(6) = 40#: values(7) = 60#: values(8) = 40#: values(9) = 40#
    Set entity = doc.ModelSpace.AddLightWeightPolyline(values): entity.Layer = "NP_OPENING": entity.Closed = False
    point(0) = 40#: point(1) = 40#: endpoint(0) = 40#: endpoint(1) = 40#
    Set importer = New CAutoCADSectionModelImporter
    For Each targetLayer In Array("NP_OUTER", "NP_OPENING")
    Set zero = doc.ModelSpace.AddLine(point, endpoint): zero.Layer = CStr(targetLayer): countBefore = doc.ModelSpace.Count
    prefix = "native.emptyContourSegment." & CStr(targetLayer)
    Set section = importer.ImportFromModelSpace(doc.ModelSpace, "NP_CONCRETE", "NP_REBAR", "Rebar", 0#, Nothing, "NP_OUTER", "NP_OPENING", doc.Layers)
    Set query = New CSectionGeometryQuery: query.Initialize section
    TestCheckArea prefix & ".area", query.ConcreteDomain.Area, 9600#
    TestCheck prefix & ".roles", query.ConcreteDomain.LoopCount = 2
    TestCheck prefix & ".warning", InStr(importer.ImportComment, CStr(zero.Handle)) > 0 And InStr(importer.ImportComment, "нулевой длины") > 0
    TestCheck prefix & ".noMutation", doc.ModelSpace.Count = countBefore And Not CBool(entity.Closed) And Not CBool(outer.Closed)
    zero.Delete: endpoint(0) = 41#: Set zero = doc.ModelSpace.AddLine(point, endpoint): zero.Layer = CStr(targetLayer)
    countBefore = doc.ModelSpace.Count
    On Error Resume Next
    Set section = importer.ImportFromModelSpace(doc.ModelSpace, "NP_CONCRETE", "NP_REBAR", "Rebar", 0#, Nothing, "NP_OUTER", "NP_OPENING", doc.Layers)
    number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo 0
    TestCheck prefix & ".nonzeroRejected", number <> 0 And InStr(description, CStr(zero.Handle)) > 0 And InStr(description, "PEDIT") > 0
    TestCheck prefix & ".failedNoMutation", doc.ModelSpace.Count = countBefore
    zero.Delete: endpoint(0) = 40#
    Next targetLayer
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' оба формата экспорта сохраняют несколько отверстий, отдельный
' бетонный участок внутри отверстия и его собственное отверстие. Повторный
' импорт выполняется обычным способом без исходных контуров и решения НДС.
Private Sub TestCheckNativeContourFormatTopology(ByVal doc As Object)
    Dim section As CSectionModel, writer As CNDMResultsWriter, query As CSectionGeometryQuery
    Dim importer As CAutoCADSectionModelImporter, restored As CSectionModel, region As CConcreteRegion
    Dim curve As Object, mesh As Object, bar As Object, entity As Object, imported As Collection
    Dim curves(0 To 0) As Object, regions As Variant, center(0 To 2) As Double
    Dim format As Variant, mode As Long, i As Long, before As Long, exported As Long, expected As Long
    Dim prefix As String, outerLayer As String, openingLayer As String, meshLayer As String, rebarLayer As String
    Set writer = New CNDMResultsWriter: Set query = New CSectionGeometryQuery
    Set importer = New CAutoCADSectionModelImporter
    For mode = 0 To 1
        For Each format In Array("Region", "Polyline")
            prefix = "native.format." & CStr(format) & "." & CStr(mode)
            outerLayer = prefix & "_OUTER": openingLayer = prefix & "_OPENING"
            meshLayer = prefix & "_MESH": rebarLayer = prefix & "_REBAR"
            doc.Layers.Add meshLayer: doc.Layers.Add rebarLayer
            Set section = New CSectionModel
            section.AddConcreteElement 4050#, 50#, 10000#, 1, , , "Rectangle", 100#, 100#
            section.AddRebarElement 4010#, 10#, 4#, GEOM_PI * 4#, "Rebar", 1
            If mode = 0 Then
                Set region = query.PolygonRegion(TestContourRectanglePoints(4000#, 0#, 4100#, 100#))
                section.Contours.AddMaterialRegion region, "MAIN", "Основная область теста."
                Set region = query.PolygonRegion(TestContourRectanglePoints(4044#, 54#, 4056#, 66#))
                section.Contours.AddMaterialRegion region, "ISLAND", "Отдельная бетонная часть внутри отверстия."
                Set region = query.PolygonRegion(TestContourRectanglePoints(4120#, 0#, 4150#, 50#))
                section.Contours.AddMaterialRegion region, "SECOND", "Вторая отдельная область теста."
                For i = 1 To 5
                    section.Contours.AddContourCircle "HOLE" & CStr(i), 4005# + 15# * i, 20#, 3#, , "H" & CStr(i), "Opening"
                Next i
                section.Contours.AddContourCircle "BIG_HOLE", 4050#, 60#, 15#, , "BIG", "Opening"
                section.Contours.AddContourCircle "ISLAND_HOLE", 4050#, 60#, 2#, , "SMALL", "Opening"
                expected = 10: If format = "Region" Then expected = 3
            Else
                section.Contours.AddContourCircle "HOLE1", 4025#, 50#, 10#, , "H1", "Opening"
                section.Contours.AddContourCircle "HOLE2", 4075#, 50#, 5#, , "H2", "Opening"
                expected = 2
            End If
            Set curve = TestNativeRectangle(doc, 100#, 100#, 0#, 4050#, 50#)
            Set curves(0) = curve: regions = doc.ModelSpace.AddRegion(curves)
            Set mesh = regions(LBound(regions)): mesh.Layer = meshLayer: curve.Delete
            center(0) = 4010#: center(1) = 10#: center(2) = 0#
            Set curve = doc.ModelSpace.AddCircle(center, 2#)
            Set curves(0) = curve: regions = doc.ModelSpace.AddRegion(curves)
            Set bar = regions(LBound(regions)): bar.Layer = rebarLayer: curve.Delete
            writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
            before = doc.ModelSpace.Count
            exported = SP35ExportSavedMaterialContoursForTests(doc, outerLayer, openingLayer, CStr(format))
            TestCheck prefix & ".count", exported = expected And doc.ModelSpace.Count = before + expected
            Set imported = New Collection: imported.Add mesh: imported.Add bar
            For i = before To doc.ModelSpace.Count - 1
                Set entity = doc.ModelSpace.Item(i): imported.Add entity
                If format = "Region" Then
                    TestCheck prefix & ".type." & CStr(i - before), CStr(entity.ObjectName) = "AcDbRegion"
                Else
                    TestCheck prefix & ".type." & CStr(i - before), CStr(entity.ObjectName) = "AcDbPolyline" And entity.Closed
                End If
                TestCheck prefix & ".noOwnershipTag." & CStr(i - before), Not TestHasLegacyContourTag(entity)
                If mode = 1 Or CStr(entity.Layer) = openingLayer Then
                    TestCheck prefix & ".openingColor." & CStr(i - before), CStr(entity.Layer) = openingLayer And entity.Color = 4
                Else
                    TestCheck prefix & ".outerColor." & CStr(i - before), CStr(entity.Layer) = outerLayer And entity.Color = 30
                End If
            Next i
            Set restored = importer.ImportFromModelSpace(imported, meshLayer, rebarLayer, "Rebar", 0#, Nothing, outerLayer, openingLayer, doc.Layers)
            query.Initialize restored: Set region = query.ConcreteDomain
            If mode = 0 Then
                TestCheckArea prefix & ".area", region.Area, 11644# - 274# * GEOM_PI
                TestCheck prefix & ".loops", region.LoopCount = 10
                TestCheck prefix & ".islandMaterial", query.ContainsPoint(region, 4055#, 60#)
                TestCheck prefix & ".islandHole", Not query.ContainsPoint(region, 4050#, 60#)
                TestCheck prefix & ".largeHole", Not query.ContainsPoint(region, 4060#, 60#)
                TestCheck prefix & ".secondPart", query.ContainsPoint(region, 4130#, 20#)
            Else
                TestCheckArea prefix & ".area", region.Area, 10000# - 125# * GEOM_PI
                TestCheck prefix & ".loops", region.LoopCount = 3
                For i = 1 To restored.Contours.Count
                    TestCheck prefix & ".noInventedOuter." & CStr(i), restored.Contours.LoopRole(i) = "Opening"
                Next i
                TestCheck prefix & ".twoKnownHoles", Not query.ContainsPoint(region, 4025#, 50#) And Not query.ContainsPoint(region, 4075#, 50#)
            End If
        Next format
    Next mode
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' выполняет обычные импорт, запись Results, экспорт и повторный импорт
' для смешанных и однородных источников контуров. Два отверстия вычитаются
' из наружной области Region. Исходники сохраняются; выходные метры
' не изменяют миллиметровые размеры объектов чертежа.
Private Sub TestCheckNativeMixedContourRoundTrip(ByVal doc As Object)
    Dim importer As CAutoCADSectionModelImporter, writer As CNDMResultsWriter, units As CUnitSystem, reader As CAutoCADContourReader
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
    Set importer = New CAutoCADSectionModelImporter: Set writer = New CNDMResultsWriter: Set reader = New CAutoCADContourReader
    Set query = New CSectionGeometryQuery: Set units = New CUnitSystem: units.LoadFromSettings settings
    For mode = 0 To 3
        prefix = "native.mixedRoundTrip." & CStr(mode)
        outerLayer = "NDM_MIX_OUTER_" & CStr(mode): openingLayer = "NDM_MIX_OPENING_" & CStr(mode)
        meshLayer = "NDM_MIX_MESH_" & CStr(mode): rebarLayer = "NDM_MIX_REBAR_" & CStr(mode)
        doc.Layers.Add outerLayer: doc.Layers.Add openingLayer: doc.Layers.Add meshLayer: doc.Layers.Add rebarLayer
        Set outer = TestNativeRectangle(doc, 100#, 100#, 0#, 1100# + 150# * mode, 100#): outer.Layer = outerLayer
        Set curves(0) = outer: regions = doc.ModelSpace.AddRegion(curves): Set mesh = regions(LBound(regions)): mesh.Layer = meshLayer
        If mode = 1 Or mode = 3 Then TestSetLegacyContourTag outer
        outerHandle = CStr(outer.Handle)
        For i = 0 To 1
            Set hole = TestNativeRectangle(doc, 10# + 10# * i, 10#, 0#, 1080# + 150# * mode + 35# * i, 85# + 30# * i)
            hole.Layer = openingLayer
            If mode = 0 Or mode = 3 Then TestSetLegacyContourTag hole
        Next i
        center(0) = 1060# + 150# * mode: center(1) = 60#
        Set curve = doc.ModelSpace.AddCircle(center, 5#)
        Set curves(0) = curve: regions = doc.ModelSpace.AddRegion(curves): Set bar = regions(LBound(regions))
        bar.Layer = rebarLayer: curve.Delete
        countBefore = doc.ModelSpace.Count
        Set section = importer.ImportFromModelSpace(doc.ModelSpace, meshLayer, rebarLayer, "Rebar", 0#, Nothing, outerLayer, openingLayer, doc.Layers)
        query.Initialize section: Set region = query.ConcreteDomain
        TestCheckArea prefix & ".importArea", region.Area, 9700#
        TestCheck prefix & ".importThreeLoops", region.LoopCount = 3 And doc.ModelSpace.Count = countBefore
        writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section), units
        expectedExported = 1
        exported = SP35ExportSavedMaterialContoursForTests(doc, outerLayer, openingLayer)
        TestCheck prefix & ".exportAddsAll", exported = expectedExported And doc.ModelSpace.Count = countBefore + 1
        Set entity = doc.ModelSpace.Item(countBefore)
        TestCheck prefix & ".exportRegion", CStr(entity.ObjectName) = "AcDbRegion" And CStr(entity.Layer) = outerLayer
        Set region = reader.ReadRegion(entity)
        TestCheckArea prefix & ".exportRegionArea", region.Area, 9700#
        TestCheck prefix & ".exportRegionHoles", region.LoopCount = 3
        Dim importedObjects As Collection: Set importedObjects = New Collection
        importedObjects.Add mesh: importedObjects.Add bar: importedObjects.Add entity
        Set section = importer.ImportFromModelSpace(importedObjects, meshLayer, rebarLayer, "Rebar", 0#, Nothing, outerLayer, openingLayer, doc.Layers)
        query.Initialize section: Set region = query.ConcreteDomain
        TestCheckArea prefix & ".exportRegionReimportArea", region.Area, 9700#
        TestCheck prefix & ".exportRegionReimportHoles", region.LoopCount = 3
        TestCheck prefix & ".originalOuterKept", CStr(doc.HandleToObject(outerHandle).Handle) = outerHandle
        outerCount = 0: openingCount = 0
        For Each entity In doc.ModelSpace
            If CStr(entity.Layer) = outerLayer Then outerCount = outerCount + 1
            If CStr(entity.Layer) = openingLayer Then openingCount = openingCount + 1
        Next entity
        TestCheck prefix & ".exportNoDuplicateHoleCurves", outerCount = 2 And openingCount = 2
        On Error Resume Next
        Set section = importer.ImportFromModelSpace(doc.ModelSpace, meshLayer, rebarLayer, "Rebar", 0#, Nothing, outerLayer, openingLayer, doc.Layers)
        number = Err.Number: Err.Clear
        On Error GoTo Failed
        TestCheck prefix & ".overlaidCopyNotIgnored", number <> 0
        number = 0
        ' Тест сам подготавливает чистый импорт, как теперь обязан пользователь.
        For i = doc.ModelSpace.Count - 1 To countBefore Step -1: doc.ModelSpace.Item(i).Delete: Next i
        Set section = importer.ImportFromModelSpace(doc.ModelSpace, meshLayer, rebarLayer, "Rebar", 0#, Nothing, outerLayer, openingLayer, doc.Layers)
        query.Initialize section: Set region = query.ConcreteDomain
        TestCheckArea prefix & ".reimportArea", region.Area, 9700#
        TestCheck prefix & ".reimportThreeLoops", region.LoopCount = 3
        exported = SP35ExportSavedMaterialContoursForTests(doc, outerLayer, openingLayer)
        TestCheck prefix & ".repeatExportAddsAll", exported = expectedExported And doc.ModelSpace.Count = countBefore + 1
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

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отверстие исходного Region уже является достоверной границей.
' Экспорт добавляет новый Region независимо от существующего исходника.
Private Sub TestCheckNativeRegionBoundaryDedup(ByVal doc As Object)
    Dim polyline As Object, source As Object, circleEntity As Object, hole As Object
    Dim curves(0 To 0) As Object, regions As Variant, center(0 To 2) As Double
    Dim reader As CAutoCADContourReader, region As CConcreteRegion, section As CSectionModel
    Dim writer As CNDMResultsWriter, countBefore As Long, exported As Long, handle As String
    doc.Layers.Add "SP35_DEDUP_OUTER": doc.Layers.Add "SP35_DEDUP_OPENING"
    Set polyline = TestNativeRectangle(doc, 100#, 100#, 0.23, 800#, 100#)
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
    TestCheck "native.append.regionWithOpeningAdded", exported = 1 And doc.ModelSpace.Count = countBefore + 1
    Set hole = doc.ModelSpace.Item(countBefore)
    TestCheck "native.append.regionType", CStr(hole.ObjectName) = "AcDbRegion"
    Set region = reader.ReadRegion(hole)
    TestCheck "native.append.regionOpeningKept", region.LoopCount = 2
    TestCheckArea "native.append.regionArea", region.Area, 10000# - 100# * GEOM_PI
    TestCheck "native.dedup.regionPreserved", doc.HandleToObject(handle).Handle = handle
    TestCheckArea "native.dedup.regionAreaPreserved", CDbl(source.Area), 10000# - 100# * GEOM_PI
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ручная очистка сохраняет все контуры и отверстия;
' Экспорт добавляет объекты без меток и без удаления существующих копий.
Private Sub TestCheckNativeSharedContours(ByVal doc As Object)
    Dim system As Object, saved As Variant, settings As CSystemSettingsReader, i As Long
    Dim original As Object, copy As Object, opening As Object, openingCopy As Object
    Dim number As Long, description As String, handle As String, copyHandle As String
    Dim openingHandle As String, openingCopyHandle As String
    Dim writer As CNDMResultsWriter, section As CSectionModel, reader As CAutoCADContourReader, region As CConcreteRegion
    Dim exported As Long, countBefore As Long, contourCount As Long, contourBefore As Long, entity As Object
    Dim name As Variant, layer As Object
    On Error GoTo Failed
    ' Тестовые исходные данные не зависят от ранее выполненных проверок.
    For Each name In Array("SP35_SAVED_CONTOURS", "SP35_OPENING")
        Set layer = Nothing
        On Error Resume Next
        Set layer = doc.Layers.Item(CStr(name))
        On Error GoTo Failed
        If layer Is Nothing Then Set layer = doc.Layers.Add(CStr(name))
    Next name
    Set system = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange: saved = system.Formula
    For i = 2 To system.Rows.Count
        If CStr(system.Cells(i, 1).Value2) = "AutoCAD.Common.SectionContourLayer" Then system.Cells(i, 2).Value2 = "SP35_SAVED_CONTOURS"
        If CStr(system.Cells(i, 1).Value2) = "AutoCAD.Common.OpeningContourLayer" Then system.Cells(i, 2).Value2 = "SP35_OPENING"
    Next i
    Set original = TestNativeRectangle(doc, 100#, 100#, 0#, 50#, 50#)
    original.Layer = "SP35_SAVED_CONTOURS": handle = CStr(original.Handle)
    Set copy = original.Copy: TestSetLegacyContourTag copy: copyHandle = CStr(copy.Handle)
    Set opening = TestNativeRectangle(doc, 20#, 20#, 0#, 20#, 20#)
    opening.Layer = "SP35_OPENING": openingHandle = CStr(opening.Handle)
    Set openingCopy = opening.Copy: TestSetLegacyContourTag openingCopy: openingCopyHandle = CStr(openingCopy.Handle)
    TestCheck "native.sharedContours.originalUnmarked", Not TestHasLegacyContourTag(original)
    TestCheck "native.sharedContours.legacyCopyMarked", TestHasLegacyContourTag(copy)
    For Each entity In doc.ModelSpace
        If CStr(entity.Layer) = "SP35_SAVED_CONTOURS" Then contourBefore = contourBefore + 1
    Next entity
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    exported = Audit03CleanupAutoCADModelSpaceForTests(doc.ModelSpace, settings)
    TestCheck "native.sharedContours.sourcePreserved", doc.HandleToObject(handle).Handle = handle
    TestCheck "native.sharedContours.outputContourPreserved", doc.HandleToObject(copyHandle).Handle = copyHandle
    TestCheck "native.sharedContours.openingPreserved", doc.HandleToObject(openingHandle).Handle = openingHandle
    TestCheck "native.sharedContours.outputOpeningPreserved", doc.HandleToObject(openingCopyHandle).Handle = openingCopyHandle
    Set section = New CSectionModel
    section.AddConcreteElement 50#, 50#, 10000#, 1, , , "Rectangle", 100#, 100#, 0#
    section.AddRebarElement 10#, 10#, 10#, GEOM_PI * 25#, "Rebar"
    Set reader = New CAutoCADContourReader: Set region = reader.ReadRegion(original)
    section.Contours.AddMaterialRegion region, "SHARED_CONTOURS", "Исходный контур."
    Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section)
    countBefore = doc.ModelSpace.Count
    exported = SP35ExportSavedMaterialContoursForTests(doc)
    TestCheck "native.sharedContours.existingSourceDoesNotBlockExport", exported = 1 And doc.ModelSpace.Count = countBefore + 1
    exported = SP35ExportSavedMaterialContoursForTests(doc)
    TestCheck "native.sharedContours.repeatAddsContour", exported = 1 And doc.ModelSpace.Count = countBefore + 2
    For Each entity In doc.ModelSpace
        If CStr(entity.Layer) = "SP35_SAVED_CONTOURS" Then contourCount = contourCount + 1
    Next entity
    TestCheck "native.sharedContours.previousOutputKept", contourCount = contourBefore + 2
    TestCheck "native.sharedContours.originalStillExists", doc.HandleToObject(handle).Handle = handle
    For i = countBefore To doc.ModelSpace.Count - 1
        TestCheck "native.sharedContours.newOutputUnmarked." & CStr(i), Not TestHasLegacyContourTag(doc.ModelSpace.Item(i))
    Next i
    GoTo Restore
Failed:
    number = Err.Number: description = Err.Description
Restore:
    On Error Resume Next
    If Not system Is Nothing Then system.Formula = saved
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "CheckNativeSharedContours", description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' экспорт создаёт один Region с двумя круглыми отверстиями.
' Повторное чтение AutoCAD даёт те же площадь и роли контуров, что общий геометрический расчёт.
Private Sub TestCheckNativeSavedContours(ByVal doc As Object, ByVal reader As CAutoCADContourReader)
    Dim section As CSectionModel, writer As CNDMResultsWriter, restored As CSectionModel
    Dim units As CUnitSystem, query As CSectionGeometryQuery, region As CConcreteRegion, entity As Object
    Dim edges As Collection, curve As CTestFakeAcadContour, contourEntity As CTestFakeAcadContour
    Dim countBefore As Long, exported As Long, i As Long, maxArea As Double, totalArea As Double, value As Double
    Set section = New CSectionModel: section.SourceType = "AutoCADImport"
    section.AddConcreteElement 50#, 50#, 10000#, 1, , , "Rectangle", 100#, 100#, 0#
    section.AddRebarElement 10#, 10#, 10#, GEOM_PI * 25#, "Rebar"
    Set edges = TestRectangleEdges(0#, 0#, 100#, 100#)
    Set curve = New CTestFakeAcadContour: curve.TestInitializeCircle 35#, 35#, 10#: edges.Add curve
    Set curve = New CTestFakeAcadContour: curve.TestInitializeCircle 70#, 65#, 5#: edges.Add curve
    Set contourEntity = TestContour("AcDbRegion", edges): Set region = reader.ReadRegion(contourEntity)
    section.Contours.AddMaterialRegion region, "NATIVE_SNAPSHOT", "Точный контур с двумя проемами."
    Set units = New CUnitSystem: units.InitializeDefaults
    Set writer = New CNDMResultsWriter: writer.WriteGeometryPreview ThisWorkbook, section, PrepareSectionSnapshot(section), units
    Set restored = ReadSectionGeometryFromResults(ThisWorkbook, "AutoCADImport")
    Set query = New CSectionGeometryQuery: query.Initialize restored
    TestCheckArea "native.snapshot.area", query.ConcreteDomain.Area, region.Area
    countBefore = doc.ModelSpace.Count
    exported = SP35ExportSavedMaterialContoursForTests(doc)
    TestCheck "native.snapshot.exportOneRegion", exported = 1 And doc.ModelSpace.Count = countBefore + 1
    For i = countBefore To doc.ModelSpace.Count - 1
        Set entity = doc.ModelSpace.Item(i)
        TestCheck "native.snapshot.regionType." & CStr(i - countBefore), CStr(entity.ObjectName) = "AcDbRegion"
        value = CDbl(entity.Area): totalArea = totalArea + value
        If value > maxArea Then maxArea = value
        Set region = reader.ReadRegion(entity)
        TestCheckArea "native.snapshot.exactArc." & CStr(i - countBefore), region.Area, value
        TestCheck "native.snapshot.twoOpenings." & CStr(i - countBefore), region.LoopCount = 3
    Next i
    TestCheckArea "native.snapshot.netAreaMM", totalArea, 10000# - 125# * GEOM_PI
    TestCheckArea "native.snapshot.netArea", totalArea, query.ConcreteDomain.Area
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' два opening на одном реальном слое импортируются и сохраняются
' вместе. В Adaptive бетонный Region сетки использует эквивалентную
' оболочку; достоверный контур не меняет его фактическую площадь и инерции.
Private Sub TestCheckNativeSeparateOpenings(ByVal doc As Object, ByVal reader As CAutoCADContourReader, ByVal query As CSectionGeometryQuery)
    Dim outer As Object, mesh As Object, bar As Object, hole As Object, holeCopy As Object, polyline As Object
    Dim curves(0 To 0) As Object, regions As Variant, center(0 To 2) As Double
    Dim importer As CAutoCADSectionModelImporter, section As CSectionModel, region As CConcreteRegion
    Dim i As Long, width As Double, height As Double, angle As Double, ix As Double, iy As Double, ixy As Double
    Dim raw As Variant, expectedIx As Double, expectedIy As Double, countBefore As Long
    doc.Layers.Add "SP35_OUTER": doc.Layers.Add "SP35_OPENING"
    doc.Layers.Add "SP35_MESH": doc.Layers.Add "SP35_REBAR"
    Set outer = TestNativeRectangle(doc, 100#, 100#, 0#, 500#, 100#): outer.Layer = "SP35_OUTER"
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
    TestCheckArea "native.separateOpenings.area", region.Area, 10000# - 125# * GEOM_PI
    TestCheck "native.separateOpenings.threeLoops", region.LoopCount = 3
    TestCheck "native.separateOpenings.excluded", Not query.ContainsPoint(region, 475#, 85#) And Not query.ContainsPoint(region, 520#, 110#)
    TestCheck "native.separateOpenings.sourceKept", doc.ModelSpace.Count = countBefore
    TestCheckArea "native.mesh.actualAreaKept", section.ConcreteArea(1), CDbl(mesh.Area)
    TestCheck "native.mesh.rectangleKept", section.ConcreteBoundaryRectangle(1, width, height, angle)
    TestCheckArea "native.mesh.equivalentArea", width * height, CDbl(mesh.Area)
    section.ConcreteLocalInertiaComponents 1, ix, iy, ixy
    TestCheck "native.mesh.actualInertiaKept", Abs(ix - expectedIx) < 0.00001 And Abs(iy - expectedIy) < 0.00001
    ' Отверстия на другом, не выбранном слое не попадают в достоверный
    ' контур. Пустой, отсутствующий и незаданный слой означают сплошную область.
    doc.Layers.Add "SP35_EMPTY_OPENING"
    Dim openingLayer As Variant
    For Each openingLayer In Array(vbNullString, "SP35_EMPTY_OPENING", "SP35_MISSING_OPENING")
        Set section = importer.ImportFromModelSpace(doc.ModelSpace, "SP35_MESH", "SP35_REBAR", "Rebar", 0#, Nothing, "SP35_OUTER", CStr(openingLayer), doc.Layers)
        query.Initialize section: Set region = query.ConcreteDomain
        TestCheckArea "native.knownOuterOnly." & CStr(openingLayer) & ".area", region.Area, 10000#
        TestCheck "native.knownOuterOnly." & CStr(openingLayer) & ".solid", region.LoopCount = 1 And _
            query.ContainsPoint(region, 475#, 85#) And query.ContainsPoint(region, 520#, 110#)
        TestCheckArea "native.knownOuterOnly." & CStr(openingLayer) & ".mechanicalMeshKept", section.ConcreteArea(1), CDbl(mesh.Area)
        TestCheck "native.knownOuterOnly." & CStr(openingLayer) & ".sourceKept", doc.ModelSpace.Count = countBefore
    Next openingLayer
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' строит исходные вершины независимо от геометрического движка.
Private Function TestNativeRectangle(ByVal doc As Object, ByVal width As Double, ByVal height As Double, _
        ByVal angle As Double, ByVal cx As Double, ByVal cy As Double) As Object
    Dim coordinates(0 To 7) As Double, px As Variant, py As Variant, i As Long, result As Object
    px = Array(-width / 2#, width / 2#, width / 2#, -width / 2#)
    py = Array(-height / 2#, -height / 2#, height / 2#, height / 2#)
    For i = 0 To 3
        coordinates(2 * i) = cx + CDbl(px(i)) * Cos(angle) - CDbl(py(i)) * Sin(angle)
        coordinates(2 * i + 1) = cy + CDbl(px(i)) * Sin(angle) + CDbl(py(i)) * Cos(angle)
    Next i
    Set result = doc.ModelSpace.AddLightWeightPolyline(coordinates): result.Closed = True
    Set TestNativeRectangle = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ожидает готовность собственного нового CAD-документа с timeout.
Private Sub TestWaitForCAD(ByVal acad As Object)
    Dim started As Single, elapsed As Double
    started = Timer
    Do
        If acad.GetAcadState().IsQuiescent Then Exit Sub
        DoEvents: Application.Wait DateAdd("s", 1, Now)
        elapsed = Timer - started: If elapsed < 0# Then elapsed = elapsed + 86400#
        If elapsed > 30# Then Err.Raise vbObjectError + 5502, , "AutoCAD не завершил инициализацию за 30 секунд."
    Loop
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ограниченный настоящий CAD-пакет; документ обязан принадлежать внешнему runner-у.
' Не создаёт/не закрывает приложения и удаляет только объекты собственного вывода.
Public Function RunNativeConcreteComponentExportTests(ByVal document As Object) As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    TestCheckNativeConcreteComponentExport document
    GoTo Finish
Failed:
    TestCheck "native.components.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finish:
    RunNativeConcreteComponentExportTests = mReport & "TOTAL_NATIVE_CONCRETE_COMPONENTS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет принадлежность отверстий внешнему кольцу и острову внутри его пустоты.
' Третья несвязная область сохраняет дуги; Area/c/I сравниваются после настоящего CAD export/read.
Private Sub TestCheckNativeConcreteComponentExport(ByVal document As Object)
    Dim query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    Dim parts As Collection: Set parts = New Collection
    parts.Add query.SubtractContainedRegion(query.CircleRegion(0#, 0#, 100#), query.CircleRegion(0#, 0#, 60#))
    parts.Add query.SubtractContainedRegion(query.CircleRegion(0#, 0#, 30#), query.CircleRegion(0#, 0#, 10#))
    parts.Add query.CircleRegion(300#, 0#, 20#)
    Dim geometry As CConcreteRegion: Set geometry = query.MergeDisjointRegions(parts)
    Dim expectedParts As Collection: Set expectedParts = query.ConnectedRegions(geometry)
    Dim before As Long: before = document.ModelSpace.Count
    Dim exported As Long, i As Long, j As Long, matched As Boolean, actualArea As Double
    Dim owned As Collection: Set owned = New Collection
    Dim entity As Object, actual As CConcreteRegion, expected As CConcreteRegion
    Dim reader As CAutoCADContourReader: Set reader = New CAutoCADContourReader
    Dim a As TGeometryProperties, b As TGeometryProperties
    Dim number As Long, description As String
    On Error GoTo Failed
    exported = TestExportConcreteElementGeometry(document.ModelSpace, geometry)
    For i = before To document.ModelSpace.Count - 1: owned.Add document.ModelSpace.Item(i): Next i
    TestCheck "native.components.count", exported = 3 And owned.Count = 3
    For i = 1 To owned.Count
        Set entity = owned(i): Set actual = reader.ReadRegion(entity)
        actualArea = actualArea + actual.Area: matched = False
        For j = 1 To expectedParts.Count
            Set expected = expectedParts(j)
            If query.RegionsCoincide(actual, expected) Then
                matched = True: query.RegionProperties actual, a: query.RegionProperties expected, b
                TestCheckArea "native.components.A." & CStr(i), a.Area, b.Area
                TestCheckArea "native.components.X." & CStr(i), a.CentroidX, b.CentroidX
                TestCheckArea "native.components.Y." & CStr(i), a.CentroidY, b.CentroidY
                TestCheckArea "native.components.Ix." & CStr(i), a.Ixc, b.Ixc
                TestCheckArea "native.components.Iy." & CStr(i), a.Iyc, b.Iyc
                TestCheck "native.components.Ixy." & CStr(i), Abs(a.Ixyc - b.Ixyc) <= _
                    64# * (2# ^ -52) * (Abs(b.Ixc) + Abs(b.Iyc)) + 0.000000000000001
                TestCheck "native.components.loops." & CStr(i), actual.LoopCount = expected.LoopCount
                TestCheck "native.components.arcs." & CStr(i), actual.SegmentCount = expected.SegmentCount
                Exit For
            End If
        Next j
        TestCheck "native.components.sameBoundary." & CStr(i), matched
    Next i
    TestCheckArea "native.components.totalArea", actualArea, 7600# * GEOM_PI
    GoTo Cleanup
Failed:
    number = Err.Number: description = Err.Description
    Resume Cleanup
Cleanup:
    On Error Resume Next
    For Each entity In owned: entity.Delete: Next entity
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "TestCheckNativeConcreteComponentExport", description
    TestCheck "native.components.ownedCleanup", document.ModelSpace.Count = before
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Настоящий geometry-only экспорт для Adaptive/Contour и выключенных/включённых контуров.
' Документ принадлежит внешнему runner; пользовательские приложения не создаются и не закрываются.
' Неиспользуемые строки намеренно имеют неверный формат; ошибки самих ячеек проверяет общий загрузчик Config.
Public Function RunNativeGeometryOnlyExportTests(ByVal document As Object) As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    Dim system As Object: Set system = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Dim savedSettings As Variant: savedSettings = system.Formula
    Dim owned As Collection: Set owned = New Collection
    Dim entity As Object, before As Long, mode As Long, contourMode As Long, added As Long, index As Long
    Dim query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    Dim region As CConcreteRegion: Set region = query.RectangleRegion(0#, 0#, 100#, 80#)
    Dim properties As TGeometryProperties: query.RegionProperties region, properties
    Dim model As CSectionModel, writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
    Dim names As Variant: names = Array("rngNDMSectionGeometry", "rngNDMElementGeometry", "rngNDMSectionContours", "rngNDMSectionProperties", "rngNDMElementResults")
    Dim tables As Object, key As Variant, values As Variant, warning As String, prefix As String
    Dim solves As Long: solves = SectionEquilibriumSolveCount()
    Dim number As Long, description As String, concreteCount As Long, rebarCount As Long, contourCount As Long
    On Error GoTo Failed
    TestSetGeometryExportSetting system, "AutoCAD.Common.ConcreteLayer", "GEO_ONLY_CONCRETE"
    TestSetGeometryExportSetting system, "AutoCAD.Common.RebarLayer", "GEO_ONLY_REBAR"
    TestSetGeometryExportSetting system, "General.NonCriticalMessagesEnabled", "No"
    TestSetGeometryExportSetting system, "AutoCAD.Color.ConcreteTension", "UNUSED_INVALID"
    TestSetGeometryExportSetting system, "AutoCAD.Export.NeutralLineEnabled", "UNUSED_INVALID"
    TestSetGeometryExportSetting system, "AutoCAD.Export.PrincipalAxesMode", "UNUSED_INVALID"
    TestSetGeometryExportSetting system, "AutoCAD.Export.ExportCrackInteractionContour", "UNUSED_INVALID"
    For mode = 0 To 1
        Set model = New CSectionModel: model.SourceType = "Generated"
        If mode = 0 Then
            model.AddConcreteElement 0#, 0#, 8000#, 1, , , "Rectangle", 100#, 80#
        Else
            model.AddConcreteElement 0#, 0#, 8000#, 1, , , "Contour", , , , , properties.Ixc, properties.Iyc, properties.Ixyc, True, region
        End If
        model.AddRebarElement 0#, 20#, 10#, 0#, "Rebar"
        model.Contours.AddMaterialRegion region, "OUTER", "Наружная граница тестового сечения."
        writer.WriteGeometryPreview ThisWorkbook, model, PrepareSectionSnapshot(model)
        For contourMode = 0 To 1
            prefix = "native.geometryOnly." & CStr(mode) & "." & CStr(contourMode)
            TestSetGeometryExportSetting system, "AutoCAD.Export.ContourEnabled", IIf(contourMode = 0, "No", "Yes")
            TestSetGeometryExportSetting system, "AutoCAD.Export.ContourFormat", IIf(contourMode = 0, "UNUSED_INVALID", "Polyline")
            TestSetGeometryExportSetting system, "AutoCAD.Common.SectionContourLayer", IIf(contourMode = 0, "*UNUSED_INVALID*", "GEO_ONLY_OUTER")
            TestSetGeometryExportSetting system, "AutoCAD.Common.OpeningContourLayer", "*UNUSED_NOT_NEEDED*"
            Set tables = CreateObject("Scripting.Dictionary")
            For Each key In names: tables.Add CStr(key), ReadPhysicalResultTable(ThisWorkbook, CStr(key)): Next key
            before = document.ModelSpace.Count: warning = vbNullString
            added = Audit03ExportAutoCADDocumentForTests(document, warning)
            Set owned = New Collection
            For index = before To document.ModelSpace.Count - 1: owned.Add document.ModelSpace.Item(index): Next index
            TestCheck prefix & ".warningAlways", warning = "В AutoCAD передано только сечение, без рассчитанного состояния."
            TestCheck prefix & ".objectCount", added = 2 + contourMode And owned.Count = added
            concreteCount = 0: rebarCount = 0: contourCount = 0
            For Each entity In owned
                TestCheck prefix & ".geometryOnlyType", CStr(entity.ObjectName) = "AcDbRegion" Or CStr(entity.ObjectName) = "AcDbPolyline"
                Select Case CStr(entity.Layer)
                    Case "GEO_ONLY_CONCRETE": concreteCount = concreteCount + 1: TestCheckArea prefix & ".concreteArea", CDbl(entity.Area), 8000#
                    Case "GEO_ONLY_REBAR": rebarCount = rebarCount + 1: TestCheckArea prefix & ".rebarArea", CDbl(entity.Area), GEOM_PI * 25#
                    Case "GEO_ONLY_OUTER": contourCount = contourCount + 1
                    Case Else: TestCheck prefix & ".unexpectedLayer." & CStr(entity.Layer), False
                End Select
            Next entity
            TestCheck prefix & ".layers", concreteCount = 1 And rebarCount = 1 And contourCount = contourMode
            For Each key In names
                values = ReadPhysicalResultTable(ThisWorkbook, CStr(key))
                TestCheck prefix & ".sourceUntouched." & CStr(key), TestGeometryExportTablesEqual(tables.Item(CStr(key)), values)
            Next key
            TestCheck prefix & ".noSolve", SectionEquilibriumSolveCount() = solves
            For Each entity In owned: entity.Delete: Next entity
            Set owned = New Collection
            TestCheck prefix & ".ownedCleanup", document.ModelSpace.Count = before
        Next contourMode
    Next mode
    GoTo Cleanup
Failed:
    number = Err.Number: description = Err.Description
    Resume Cleanup
Cleanup:
    On Error Resume Next
    For Each entity In owned: entity.Delete: Next entity
    system.Formula = savedSettings
    On Error GoTo 0
    If number <> 0 Then TestCheck "native.geometryOnly.runtime; " & CStr(number) & "; " & description, False
    RunNativeGeometryOnlyExportTests = mReport & "TOTAL_NATIVE_GEOMETRY_ONLY: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет ровно один ключ в частной книге; отсутствие настройки не заменяет молчаливым default.
Private Sub TestSetGeometryExportSetting(ByVal system As Object, ByVal key As String, ByVal value As Variant)
    Dim row As Long
    For row = 2 To system.Rows.Count
        If CStr(system.Cells(row, 1).Value2) = key Then system.Cells(row, 2).Value2 = value: Exit Sub
    Next row
    Err.Raise 5, "modTestAutoCADContours", "В частной книге отсутствует настройка " & key
End Sub
