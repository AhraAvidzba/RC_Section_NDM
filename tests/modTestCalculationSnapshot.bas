Attribute VB_Name = "modTestCalculationSnapshot"
Option Explicit

' ==================== ДЛЯ ТЕСТОВ ====================
' Полный цикл снимка: точные числа, оба хранилища, смена сечения и ошибки файла.
' Выполняется только runner-ом в его частной копии книги; внешние файлы принадлежат тесту.
Private Type TTestSnapshotDouble
    Value As Double ' Число, сравниваемое без округления.
End Type
Private Type TTestSnapshotBits
    Low As Long ' Первые четыре байта Double.
    High As Long ' Последние четыре байта Double.
End Type
Private mTestReport As String, mTestPassed As Long, mTestFailed As Long ' Итог и компактные адреса расхождений.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет реальные вызовы расчёта и кнопки, полную замену файла и защиту сохранённых Results.
Public Function RunCalculationSnapshotTests() As String
    mTestReport = "": mTestPassed = 0: mTestFailed = 0
    On Error GoTo Failed
    TestSnapshotBinaryBuffer
    TestSnapshotUTF16Strings
    TestSnapshotAdaptiveExtent
    TestSnapshotStatusSpans
    TestSnapshotOutputGuards
    TestSnapshotOutputPolicies
    TestSnapshotSeparatedProperties
    TestSnapshotOriginalLoadMemory
    Dim initial As CSystemSettingsReader: Set initial = New CSystemSettingsReader: initial.LoadFromWorkbook ThisWorkbook
    Dim extension As String: extension = initial.GetRequiredString("General.DiagramExtension")
    TestSnapshotFileGeometryAuthority extension
    modTestPerformance.TestConfigurePerformanceFixture "MIXED", 6, "Newton", "Circle", "Normal", extension
    TestSnapshotSetting "General.ExecutionReportEnabled", "No"
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    Dim token As String: token = "TestSnapshot_" & Format$(Now, "yyyymmdd_hhnnss")
    TestSnapshotSetting "Results.SnapshotStorage", "File"
    TestSnapshotSetting "Results.SnapshotName", token & "_" & ChrW$(&H3B1)
    Dim path As String: path = store.SnapshotPath(ThisWorkbook)
    modTestPlotConfig.TestSetOutputCombination ThisWorkbook, "PERF_01"
    Dim loads As Object: Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells(2, 1).Value2 = "Сочетание_" & ChrW$(&H3B1) & "_01": modTestPlotConfig.TestSetOutputCombination ThisWorkbook, "Сочетание_" & ChrW$(&H3B1) & "_01"
    Dim message As String: message = FormSectionForWorkbook(ThisWorkbook)
    message = RunSectionCalculationForWorkbook(ThisWorkbook)
    TestSnapshotCheck "readScope.afterCalculation", TestCalculationSnapshotReadReleased()
    Dim original As CCalculationSnapshot: Set original = store.LoadFile(ThisWorkbook)
    TestSnapshotSummaryLoads original
    TestSnapshotLegacySummary original
    TestSnapshotOneUseSource
    TestSnapshotNativeLayout original
    TestSnapshotValidationReuse original
    TestSnapshotElementIndex
    TestSnapshotGeometryPreview original, store
    TestSnapshotBoundaryFileCodec original, store
    Dim originalPlot As String: originalPlot = TestSnapshotPlot()
    modTestPlotConfig.TestSetOutputCombination ThisWorkbook, "Сочетание_" & ChrW$(&H3B1) & "_01"
    Dim originalExport As String: originalExport = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    TestSnapshotCheck "native.calculate.fileReference", store.HasFileReference(ThisWorkbook)
    TestSnapshotCheck "native.calculate.noSheetArchive", Not TestSnapshotNameExists("rngNDMLoadCaseResults")
    TestSnapshotCheck "native.calculate.states", UBound(original.Table("rngNDMElementResults"), 1) > 1
    TestSnapshotCheck "native.calculate.elapsed", original.ElapsedSeconds > 0#
    TestSnapshotCheck "native.common.onlyALL", TestSnapshotCommonOnly(original.Table("rngNDMSectionProperties"))
    TestSnapshotCheck "native.lc.separate", UBound(original.Table("LoadCase.Properties"), 1) > 1
    TestSnapshotUpper original
    TestSnapshotCorruptUpper
    message = vbNullString
    UpdateSectionResultsForWorkbook ThisWorkbook
    TestSnapshotUpper original
    TestSnapshotCheck "file.restored.plot.exact", TestSnapshotPlot() = originalPlot
    TestSnapshotSetting "Results.SnapshotName", token & "_" & ChrW$(&H3B1)
    TestSnapshotCheck "name.extensionAutomatic", path = ThisWorkbook.Path & Application.PathSeparator & token & "_" & ChrW$(&H3B1) & ".ndm"
    TestSnapshotSetting "Results.SnapshotName", token & "_" & ChrW$(&H3B1) & ".ndm"
    TestSnapshotCheck "name.explicitSuffixCompatible", store.SnapshotPath(ThisWorkbook) = path
    TestSnapshotSetting "Results.SnapshotName", token & "_" & ChrW$(&H3B1) & ".ndms"
    TestSnapshotCheck "name.previousSuffixCompatible", store.SnapshotPath(ThisWorkbook) = path
    TestSnapshotSetting "Results.SnapshotName", token & "_" & ChrW$(&H3B1)
    TestSnapshotLegacyFile original, store
    store.SaveFile ThisWorkbook, original
    Dim fileSnapshot As CCalculationSnapshot: Set fileSnapshot = store.LoadFile(ThisWorkbook)
    TestSnapshotEqual "file.allTables", original, fileSnapshot
    TestSnapshotSetting "Results.SnapshotStorage", "File"
    UpdateSectionResultsForWorkbook ThisWorkbook
    TestSnapshotCheck "file.linked", store.HasFileReference(ThisWorkbook)
    TestSnapshotCheck "file.noSheetArchive", Not TestSnapshotNameExists("rngNDMLoadCaseResults")
    TestSnapshotUpper original
    TestSnapshotCheck "file.plot.exact", TestSnapshotPlot() = originalPlot
    TestSnapshotCheck "file.export.exact", Audit02ReadExportSnapshotForTests(ThisWorkbook) = originalExport
    TestSnapshotCheck "readScope.afterExport", TestCalculationSnapshotReadReleased()
    TestSnapshotEqual "file.linkedExact", original, store.LoadLinkedFile(ThisWorkbook)
    Dim geometry As Variant: geometry = ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionGeometry")
    TestSnapshotCheck "file.mesh.alwaysShown", TestSnapshotSameMatrix(geometry, original.Table("rngNDMSectionGeometry"))
    TestSnapshotCheck "file.diagrams.alwaysShown", TestSnapshotSameMatrix(ReadPhysicalResultTable(ThisWorkbook, "rngNDMMaterialDiagrams"), original.Table("rngNDMMaterialDiagrams"))
    Dim profiles As Object: Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Dim header As Object: Set header = profiles.Find("PR1", LookAt:=1)
    If header Is Nothing Then Err.Raise 5, , "В тестовой таблице отсутствует PR1."
    header.Value2 = "PR91"
    TestSnapshotCheck "file.savedProfile.fallback", TestSnapshotPlot() = originalPlot
    TestSnapshotCheck "file.savedProfile.export", Audit02ReadExportSnapshotForTests(ThisWorkbook) = originalExport
    header.Value2 = "PR1"
    ' Новый выбранный файл не должен менять показанное НДС до нажатия кнопки.
    TestSnapshotSetting "Results.SnapshotName", token & "_other.ndm"
    TestSnapshotCheck "selection.missingFileRejectedImmediately", TestSnapshotFileReadRejected(1)
    TestSnapshotCheck "missing.rejected", TestSnapshotRefreshRejected()
    TestSnapshotCheck "readScope.afterFailure", TestCalculationSnapshotReadReleased()
    TestSnapshotCheck "missing.keepsGeometry", TestSnapshotSameMatrix(geometry, ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionGeometry"))
    TestSnapshotSetting "Results.SnapshotName", token & "_" & ChrW$(&H3B1) & ".ndm"
    Dim file As CBinarySnapshotFile: Set file = New CBinarySnapshotFile
    Dim value As Byte: file.OpenUpdate path: file.Position = 95: value = file.ReadByte
    file.Position = 95: file.WriteByte value Xor 1: file.CloseFile
    TestSnapshotCheck "corrupt.rejected", TestSnapshotRefreshRejected()
    TestSnapshotCheck "corrupt.keepsGeometry", TestSnapshotSameMatrix(geometry, ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionGeometry"))
    store.SaveFile ThisWorkbook, original
    UpdateSectionResultsForWorkbook ThisWorkbook
    ' Другая геометрия и единицы, тот же файл: строки заменяемого снимка должны исчезнуть полностью.
    TestSnapshotSetting "Geometry.Source", "Generated": TestSnapshotSetting "Geometry.Type", "Rectangles"
    TestSnapshotSetting "Units.Length.Output", "cm": TestSnapshotSetting "Units.Area.Output", "cm2"
    TestSnapshotSetting "Units.Stress.Output", "kgf/cm2"
    message = FormSectionForWorkbook(ThisWorkbook)
    message = RunSectionCalculationForWorkbook(ThisWorkbook)
    Dim second As CCalculationSnapshot: Set second = store.LoadFile(ThisWorkbook)
    TestSnapshotCheck "overwrite.newGeometry", Not TestSnapshotSameMatrix(original.Table("rngNDMSectionGeometry"), second.Table("rngNDMSectionGeometry"))
    TestSnapshotCheck "overwrite.units", second.OutputLengthUnit = "cm"
    TestSnapshotCheck "overwrite.oneFile", Not CreateObject("Scripting.FileSystemObject").FileExists(path & ".part") And CreateObject("Scripting.FileSystemObject").FileExists(path)
    TestSnapshotUpper second
    Dim secondPlot As String: secondPlot = TestSnapshotPlot()
    TestSnapshotSetting "Results.SnapshotName", token & "_other.ndm": store.SaveFile ThisWorkbook, original
    Dim inputs As Variant: inputs = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange.Formula
    TestSnapshotCorruptUpper
    UpdateSectionResultsForWorkbook ThisWorkbook
    TestSnapshotUpper original
    TestSnapshotCheck "switch.otherSection.plot", TestSnapshotPlot() = originalPlot
    TestSnapshotCheck "switch.inputPreserved", TestSnapshotSameMatrix(inputs, ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange.Formula)
    TestSnapshotSetting "Results.SnapshotName", token & "_" & ChrW$(&H3B1) & ".ndm": UpdateSectionResultsForWorkbook ThisWorkbook
    TestSnapshotCheck "switch.back.plot", TestSnapshotPlot() = secondPlot
    ' Никакого обращения к решателю при обновлении: даже ошибочный solver Config допустим.
    TestSnapshotSetting "Solver.MaxIterations", "not_a_number"
    UpdateSectionResultsForWorkbook ThisWorkbook
    TestSnapshotUpper second
    TestSnapshotSetting "Solver.MaxIterations", "100"
    TestSnapshotSetting "Results.SnapshotStorage", "Excel"
    Dim detached As CCalculationSnapshot: Set detached = CurrentCalculationSnapshot(ThisWorkbook)
    TestSnapshotCheck "excel.ignoresPreviousFileReference", detached Is Nothing
    UpdateSectionResultsForWorkbook ThisWorkbook, second ' Только явная запись подготовленного снимка создаёт полный Excel-набор; чтение Excel само файл не переносит.
    TestSnapshotCheck "excel.materializedBeforeDetach", Not store.HasFileReference(ThisWorkbook)
    TestSnapshotCheck "excel.materializedBoundaryExact", TestSnapshotSameMatrix(ReadPhysicalResultTable(ThisWorkbook, "rngNDMElementGeometry"), second.Table("rngNDMElementGeometry"))
    SaveAndDisplayCalculationSnapshot ThisWorkbook, second
    TestSnapshotExcelCurrent second
    Dim invalidName As Variant
    For Each invalidName In Array("../escape", "CON.ndm", "a:b", "", "NUL", "LPT1.ndm")
        TestSnapshotSetting "Results.SnapshotName", CStr(invalidName)
        TestSnapshotCheck "name.rejected." & CStr(invalidName), TestSnapshotPathRejected(store)
    Next invalidName
    TestSnapshotSetting "Results.SnapshotName", token & "_" & ChrW$(&H3B1) & ".ndm"
    Dim preview As String: preview = FormSectionForWorkbook(ThisWorkbook)
    TestSnapshotCheck "preview.clearsSnapshot", Not store.HasFileReference(ThisWorkbook)
    TestSnapshotCheck "readScope.afterPreview", TestCalculationSnapshotReadReleased()
    TestSnapshotCheck "layout.lowerOrder", TestSnapshotLowerOrder()
    TestSnapshotRepeatedImportedAnnotations extension
    TestSnapshotColumnCodec second, store
    TestSnapshotRepeatedElementIDs second, store
    TestSnapshotMinimalFormatting
    GoTo Finished
Failed:
    TestSnapshotCheck "runtime " & CStr(Err.Number) & ": " & Err.Description, False
Finished:
    On Error Resume Next
    If Not store Is Nothing Then store.ClearReference ThisWorkbook
    If Len(path) > 0 Then If CreateObject("Scripting.FileSystemObject").FileExists(path) Then CreateObject("Scripting.FileSystemObject").DeleteFile path
    Dim otherPath As String: otherPath = ThisWorkbook.Path & Application.PathSeparator & token & "_other.ndm"
    If Len(token) > 0 Then If CreateObject("Scripting.FileSystemObject").FileExists(otherPath) Then CreateObject("Scripting.FileSystemObject").DeleteFile otherPath
    On Error GoTo 0
    RunCalculationSnapshotTests = mTestReport & "TOTAL CALCULATION SNAPSHOT: passed=" & CStr(mTestPassed) & "; failed=" & CStr(mTestFailed)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет обязательные контуры элементов без нового расчёта: линии, дуги,
' отверстия, раздельные loops, три единицы, метаданные и повреждения матриц.
Public Function RunConcreteGeometrySnapshotTests() As String
    mTestReport = "": mTestPassed = 0: mTestFailed = 0
    On Error GoTo Failed
    Dim geometry As Variant, boundaries As Variant, expected As Collection, unit As Variant, map As Object
    Dim i As Long, query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    For Each unit In Array("mm", "cm", "m")
        TestSnapshotBoundaryFixture CStr(unit), geometry, boundaries, expected
        Set map = ReadConcreteBoundaryMap(boundaries, geometry)
        TestSnapshotCheck "boundary." & CStr(unit) & ".count", map.Count = 4
        For i = 1 To expected.Count
            TestSnapshotCheck "boundary." & CStr(unit) & ".geometry.C" & CStr(i), query.RegionsCoincide(expected(i), map.Item("C" & CStr(i)))
            TestSnapshotCheck "boundary." & CStr(unit) & ".loops.C" & CStr(i), expected(i).LoopCount = map.Item("C" & CStr(i)).LoopCount
        Next i
        TestSnapshotCheck "boundary." & CStr(unit) & ".quarterArea", Abs(map.Item("C1").Area - GEOM_PI * 2500#) < 0.000001
        TestSnapshotCheck "boundary." & CStr(unit) & ".holeArea", Abs(map.Item("C2").Area - (10000# - GEOM_PI * 225#)) < 0.000001
        TestSnapshotCheck "boundary." & CStr(unit) & ".islandsArea", Abs(map.Item("C3").Area - GEOM_PI * 244#) < 0.000001
        TestSnapshotCheck "boundary." & CStr(unit) & ".rectangleArea", Abs(map.Item("C4").Area - 2400#) < 0.000001
        TestSnapshotBoundaryDamage geometry, boundaries
    Next unit
    GoTo Finished
Failed:
    TestSnapshotCheck "boundary.runtime " & CStr(Err.Number) & ": " & Err.Description, False
Finished:
    RunConcreteGeometrySnapshotTests = mTestReport & "TOTAL_CONCRETE_GEOMETRY_SNAPSHOT: passed=" & CStr(mTestPassed) & "; failed=" & CStr(mTestFailed)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт независимый набор записей публичного формата, не вызывая writer.
' Аналитические области включают четверть круга, отверстие, два острова
' внутри одного элемента и повёрнутый прямоугольник.
Private Sub TestSnapshotBoundaryFixture(ByVal unit As String, ByRef geometry As Variant, ByRef boundaries As Variant, ByRef expected As Collection)
    Dim query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    Set expected = New Collection
    expected.Add query.ClipWindow(query.CircleRegion(0#, 0#, 100#), 1#, 0#, 0#, 100#, 0#, 100#)
    expected.Add query.SubtractContainedRegion(query.RectangleRegion(300#, 0#, 100#, 100#), query.CircleRegion(300#, 0#, 15#))
    Dim islands As Collection: Set islands = New Collection
    islands.Add query.CircleRegion(500#, 0#, 10#): islands.Add query.CircleRegion(550#, 0#, 12#)
    expected.Add query.MergeDisjointRegions(islands)
    expected.Add query.RectangleRegion(700#, 0#, 80#, 30#, 0.2)
    Dim factor As Double
    Select Case unit
        Case "mm": factor = 1#
        Case "cm": factor = 0.1
        Case "m": factor = 0.001
        Case Else: Err.Raise 5, , "Неизвестная единица тестовой матрицы."
    End Select
    Dim data() As Variant, edges() As Variant, headers As Variant, i As Long, j As Long, total As Long, row As Long
    ReDim data(1 To 5, 1 To 16)
    headers = Array("RunID", "ElementID", "MaterialType", "X, " & unit, "Y, " & unit, "Area, " & unit & "2", _
        "GeometryInterpretationStatus", "Width, " & unit, "Height, " & unit, "Diameter, " & unit, "Rotation, rad", _
        "LocalIx, " & unit & "4", "LocalIy, " & unit & "4", "LocalIxy, " & unit & "4", "Comment", "ExtraData")
    For i = 1 To 16: data(1, i) = headers(i - 1): Next i
    Dim groups As Collection: Set groups = New Collection
    Dim contour As CSectionContours, region As CConcreteRegion, props As TGeometryProperties
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    For i = 1 To expected.Count
        Set region = expected(i): query.RegionProperties region, props
        query.ProjectionBounds region, 1#, 0#, minX, maxX: query.ProjectionBounds region, 0#, 1#, minY, maxY
        row = i + 1
        data(row, 1) = 45000#: data(row, 2) = "C" & CStr(i): data(row, 3) = "Concrete"
        data(row, 4) = props.CentroidX * factor: data(row, 5) = props.CentroidY * factor
        data(row, 6) = props.Area * factor ^ 2: data(row, 7) = "Contour"
        data(row, 8) = 0#: data(row, 9) = 0#: data(row, 11) = 0#
        If i = 4 Then data(row, 8) = 80# * factor: data(row, 9) = 30# * factor: data(row, 11) = 0.2 ' Только настоящий прямоугольник имеет эти параметры вместо обязательного контура.
        data(row, 12) = props.Ixc * factor ^ 4: data(row, 13) = props.Iyc * factor ^ 4: data(row, 14) = props.Ixyc * factor ^ 4
        data(row, 15) = "Пример " & ChrW$(&H3B1): data(row, 16) = "{}"
        Set contour = New CSectionContours: contour.AddMaterialRegion region, "C" & CStr(i), ""
        groups.Add contour: total = total + contour.Count
    Next i
    headers = Array("RunID", "LoopID", "SegmentID", "Sequence", "LoopRole", "SegmentType", "StartX", "StartY", _
        "EndX", "EndY", "CenterX", "CenterY", "Radius", "SweepAngle, rad", "SourceID", "Comment", "LengthUnit (SectionXY)", "ElementID")
    ReDim edges(1 To total + 1, 1 To 18)
    For i = 1 To 18: edges(1, i) = headers(i - 1): Next i
    row = 1
    For i = 1 To groups.Count
        Set contour = groups(i)
        For j = 1 To contour.Count
            row = row + 1
            edges(row, 1) = 45000#: edges(row, 2) = contour.LoopID(j): edges(row, 3) = contour.ContourID(j)
            edges(row, 4) = row - 1: edges(row, 5) = contour.LoopRole(j): edges(row, 6) = contour.SegmentType(j)
            edges(row, 7) = contour.StartX(j) * factor: edges(row, 8) = contour.StartY(j) * factor
            If contour.SegmentType(j) = "CONTOUR_LINE" Then
                edges(row, 9) = contour.EndX(j) * factor: edges(row, 10) = contour.EndY(j) * factor
            ElseIf contour.SegmentType(j) = "CONTOUR_CIRCLE" Then
                edges(row, 13) = contour.Radius(j) * factor
            Else
                edges(row, 11) = contour.CenterX(j) * factor: edges(row, 12) = contour.CenterY(j) * factor: edges(row, 14) = contour.SweepAngle(j)
            End If
            edges(row, 15) = contour.SourceID(j): edges(row, 16) = contour.Comment(j): edges(row, 17) = unit: edges(row, 18) = "C" & CStr(i)
        Next j
    Next i
    geometry = data: boundaries = edges
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Изменяет по одному обязательному полю и требует адресного отказа.
' Повреждённые дополнительные подписи и JSON намеренно не влияют на форму.
Private Sub TestSnapshotBoundaryDamage(ByRef geometry As Variant, ByRef boundaries As Variant)
    Dim changed As Variant, changedGeometry As Variant, missingData As Variant, test As Long, map As Object
    TestSnapshotCheck "boundary.requiredTable", TestSnapshotBoundaryRejected(missingData, geometry)
    For test = 1 To 24
        changed = boundaries: changedGeometry = geometry
        Select Case test
            Case 1: changed(2, 18) = "R1"
            Case 2: changed(3, 3) = changed(2, 3)
            Case 3: changed(2, 4) = 2#
            Case 4: changed(2, 1) = 45001#
            Case 5: changed(2, 5) = "Opening"
            Case 6: changed(2, 17) = "unknown"
            Case 7: changed(2, 6) = "Spline"
            Case 8: changed(2, 7) = "1 234.5"
            Case 9: changed(2, 7) = CVErr(2015)
            Case 10: changedGeometry(2, 6) = changedGeometry(2, 6) * 1.01
            Case 11: changedGeometry(2, 4) = changedGeometry(2, 4) + 100#
            Case 12: changedGeometry(2, 12) = changedGeometry(2, 12) * 1.01
            Case 13: changedGeometry(2, 7) = "Region"
            Case 14: changed(2, 2) = "DifferentLoop"
            Case 15: changed(2, 18) = vbNullString
            Case 16: changed(1, 18) = "OptionalElementID"
            Case 17: changedGeometry(2, 8) = -1#
            Case 18: changed(2, 7) = changed(2, 7) + 1#
            Case 19
                Dim arc As Long
                For arc = 2 To UBound(changed, 1)
                    If CStr(changed(arc, 6)) = "CONTOUR_ARC" Then changed(arc, 14) = False: Exit For
                Next arc
            Case 20: changedGeometry(2, 14) = changedGeometry(2, 14) + changedGeometry(2, 12)
            Case 21: changed(2, 17) = vbNullString
            Case 22: changedGeometry(1, 4) = "X"
            Case 23: changedGeometry(1, 13) = "LocalIy, unknown4"
            Case 24: changedGeometry(1, 11) = "Rotation, deg"
        End Select
        TestSnapshotCheck "boundary.damage." & CStr(test), TestSnapshotBoundaryRejected(changed, changedGeometry)
    Next test
    changed = boundaries: changedGeometry = geometry
    For test = 2 To UBound(changedGeometry, 1): changedGeometry(test, 1) = CDate(45000#): Next test
    For test = 2 To UBound(changed, 1): changed(test, 1) = CDate(45000#): Next test
    Set map = ReadConcreteBoundaryMap(changed, changedGeometry)
    TestSnapshotCheck "boundary.nativeDateRunID", map.Count = 4
    changed = boundaries: changedGeometry = geometry
    For test = 2 To UBound(changedGeometry, 1)
        changedGeometry(test, 15) = CVErr(2015): changedGeometry(test, 16) = "{broken optional JSON"
    Next test
    For test = 2 To UBound(changed, 1): changed(test, 15) = CVErr(2015): changed(test, 16) = Null: Next test
    Set map = ReadConcreteBoundaryMap(changed, changedGeometry)
    TestSnapshotCheck "boundary.optionalIgnored", map.Count = 4
    For test = 2 To UBound(changedGeometry, 1): changedGeometry(test, 7) = "Region": Next test
    Set map = ReadConcreteBoundaryMap(missingData, changedGeometry)
    TestSnapshotCheck "boundary.legacyRegionWithoutTable", map.Count = 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет отдельный геометрический File без НДС и сохраняет его для внешнего close/reopen gate.
Public Function RunGeometryPreviewStorageTests() As String
    mTestReport = "": mTestPassed = 0: mTestFailed = 0
    On Error GoTo Failed
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    TestSnapshotSetting "Results.SnapshotStorage", "File"
    TestSnapshotSetting "Results.SnapshotName", "TestGeometryPreview_" & Format$(Now, "yyyymmdd_hhnnss")
    TestSnapshotGeometryPreview Nothing, store
    GoTo Finished
Failed:
    TestSnapshotCheck "preview.runtime " & CStr(Err.Number) & ": " & Err.Description, False
Finished:
    RunGeometryPreviewStorageTests = mTestReport & "TOTAL_GEOMETRY_PREVIEW_STORAGE: passed=" & CStr(mTestPassed) & "; failed=" & CStr(mTestFailed)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет линии, дуги, отверстие и два острова в File и в явно подготовленном Excel-снимке.
Private Sub TestSnapshotGeometryPreview(ByVal original As CCalculationSnapshot, ByVal store As CCalculationSnapshotStore)
    Dim query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    Dim geometry As Variant, boundaries As Variant, regions As Collection, values As Variant
    TestSnapshotBoundaryFixture "mm", geometry, boundaries, regions
    Dim model As CSectionModel: Set model = New CSectionModel: model.SourceType = "Generated"
    Dim region As CConcreteRegion, props As TGeometryProperties, i As Long
    For i = 1 To regions.Count
        Set region = regions(i): query.RegionProperties region, props
        model.AddConcreteElement props.CentroidX, props.CentroidY, props.Area, 1, shapeType:="Contour", _
            localIx:=props.Ixc, localIy:=props.Iyc, localIxy:=props.Ixyc, rotationKnown:=True, exactRegion:=region
    Next i
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Dim units As CUnitSystem: Set units = New CUnitSystem: units.LoadFromSettings settings
    Dim writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
    Dim before As Variant: before = ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionGeometry")
    Dim preview As CCalculationSnapshot: Set preview = writer.PrepareGeometryPreviewSnapshot(ThisWorkbook, model, PrepareSectionSnapshot(model), units)
    TestSnapshotCheck "preview.capture.noSheetWrite", TestSnapshotOptionalMatrixEqual(before, ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionGeometry"))
    TestSnapshotCheck "preview.explicitKind", preview.IsGeometryPreview And preview.HasTable("Snapshot.Kind")
    TestSnapshotCheck "preview.noFakeSummary", Not preview.HasTable("Summary") And Not preview.HasTable("Calculation.Profiles") And Not preview.HasTable("LoadCase.Properties")
    TestSnapshotCheck "preview.noStates", preview.TableRowCount("rngNDMElementResults") = 1 And preview.TableRowCount("rngNDMMaterialDiagrams") = 1
    Dim missing As CCalculationSnapshot: Set missing = New CCalculationSnapshot: missing.RunID = preview.RunID
    Dim undeclared As CCalculationSnapshot: Set undeclared = New CCalculationSnapshot: undeclared.RunID = preview.RunID
    Dim key As Variant
    For Each key In preview.Keys()
        If CStr(key) <> "rngNDMElementGeometry" Then missing.AddTable CStr(key), preview.Table(CStr(key))
        If CStr(key) <> "Snapshot.Kind" Then undeclared.AddTable CStr(key), preview.Table(CStr(key))
    Next key
    TestSnapshotCheck "preview.missingBoundaryRejected", TestSnapshotValidationRejected(missing)
    TestSnapshotCheck "preview.neverInferredFromMissingLC", TestSnapshotValidationRejected(undeclared)
    Dim unknown As CCalculationSnapshot: Set unknown = New CCalculationSnapshot: unknown.RunID = preview.RunID
    Dim oversized As CCalculationSnapshot: Set oversized = New CCalculationSnapshot: oversized.RunID = preview.RunID
    For Each key In preview.Keys()
        values = preview.Table(CStr(key))
        If CStr(key) = "Snapshot.Kind" Then values(1, 1) = "UnknownPreview"
        unknown.AddTable CStr(key), values
        oversized.AddTable CStr(key), preview.Table(CStr(key))
    Next key
    Dim fake(1 To 1, 1 To 1) As Variant: fake(1, 1) = "FakeSummary"
    oversized.AddTable "Summary", fake
    TestSnapshotCheck "preview.unknownKindRejected", TestSnapshotValidationRejected(unknown)
    TestSnapshotCheck "preview.cannotSmuggleFullTables", TestSnapshotValidationRejected(oversized)
    Dim withState As CCalculationSnapshot: Set withState = New CCalculationSnapshot: withState.RunID = preview.RunID
    For Each key In preview.Keys()
        values = preview.Table(CStr(key))
        If CStr(key) = "rngNDMSectionProperties" Then values(2, 3) = "State.StrengthState.Epsilon0"
        withState.AddTable CStr(key), values
    Next key
    TestSnapshotCheck "preview.cannotSmuggleStateProperties", TestSnapshotValidationRejected(withState)
    Dim failedPreparation As CCalculationSnapshot: Set failedPreparation = New CCalculationSnapshot: failedPreparation.RunID = preview.RunID
    For Each key In preview.Keys()
        values = preview.Table(CStr(key))
        If CStr(key) = "rngNDMSectionGeometry" Then values(2, 6) = CDbl(values(2, 6)) * 1.01
        failedPreparation.AddTable CStr(key), values
    Next key
    TestSnapshotDirectBoundaryRejection failedPreparation, preview
    TestSnapshotSetting "Results.SnapshotStorage", "Excel"
    SaveAndDisplayCalculationSnapshot ThisWorkbook, preview
    boundaries = preview.Table("rngNDMElementGeometry")
    TestSnapshotCheck "preview.excel.allBoundaries", TestSnapshotSameMatrix(boundaries, ReadPhysicalResultTable(ThisWorkbook, "rngNDMElementGeometry"))
    TestSnapshotSetting "Results.SnapshotStorage", "File"
    writer.RestoreSnapshotBase ThisWorkbook, preview, False, False
    Dim clearedBody As Object: Set clearedBody = ThisWorkbook.Names.Item("rngNDMElementGeometry").RefersToRange.Offset(1, 0).Resize(UBound(boundaries, 1) - 1, 18)
    TestSnapshotCheck "preview.file.noFormatClearsOldBorders", clearedBody.Cells(2, 2).Borders.Item(7).LineStyle = -4142 And clearedBody.Cells(2, 2).Borders.Item(8).LineStyle = -4142
    SaveAndDisplayCalculationSnapshot ThisWorkbook, preview
    Dim loaded As CCalculationSnapshot: Set loaded = store.LoadFile(ThisWorkbook)
    TestSnapshotEqual "preview.fileExact", preview, loaded
    Dim physical As Variant: physical = ReadPhysicalResultTable(ThisWorkbook, "rngNDMElementGeometry")
    TestSnapshotCheck "preview.file.header18", UBound(physical, 1) = 1 And UBound(physical, 2) = 18
    TestSnapshotCheck "preview.file.oldRowsCleared", Application.CountA(ThisWorkbook.Names.Item("rngNDMElementGeometry").RefersToRange.Offset(1, 0).Resize(UBound(boundaries, 1) - 1, 18)) = 0
    Dim file As CBinarySnapshotFile: Set file = New CBinarySnapshotFile: file.OpenRead store.SnapshotPath(ThisWorkbook)
    TestSnapshotCheck "preview.file.version05", file.ReadASCII(8) = "NDMSNP05" And file.ReadLong() = 5
    file.CloseFile
    Dim checksum As CFileChecksum: Set checksum = New CFileChecksum
    Dim path As String: path = store.SnapshotPath(ThisWorkbook)
    Dim fileHash As String: fileHash = checksum.FileSHA256(path)
    before = ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionGeometry")
    Dim fs As Object: Set fs = CreateObject("Scripting.FileSystemObject")
    Dim stream As Object, number As Long
    If fs.FileExists(path & ".part") Then Err.Raise 5, , "Чужой .part занимает путь отрицательного теста preview."
    Set stream = fs.CreateTextFile(path & ".part", False): stream.Write "TestPreviewSaveFailure": stream.Close
    On Error Resume Next
    SaveAndDisplayCalculationSnapshot ThisWorkbook, preview
    number = Err.Number: Err.Clear
    On Error GoTo 0
    fs.DeleteFile path & ".part"
    TestSnapshotCheck "preview.saveFailure.explicit", number = vbObjectError + 4950
    TestSnapshotCheck "preview.saveFailure.keepsFile", checksum.FileSHA256(path) = fileHash
    TestSnapshotCheck "preview.saveFailure.keepsResults", TestSnapshotOptionalMatrixEqual(before, ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionGeometry"))
    TestSnapshotCheck "preview.saveFailure.keepsLink", store.HasFileReference(ThisWorkbook)
    TestSnapshotMissingGeometryFile store
    Dim restored As CSectionModel: Set restored = ReadFormedSectionFromResults(ThisWorkbook, False)
    TestSnapshotCheck "preview.file.recalculationSource", restored.ConcreteCount = regions.Count
    For i = 1 To regions.Count
        TestSnapshotCheck "preview.file.exactRegion.C" & CStr(i), query.RegionsCoincide(restored.ConcreteGeometry(i), regions(i))
    Next i
    Dim beforePlot As String: beforePlot = TestSnapshotPlot()
    Dim beforeExport As String: beforeExport = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    TestSnapshotSetting "Results.SnapshotStorage", "Excel"
    Set loaded = CurrentCalculationSnapshot(ThisWorkbook)
    TestSnapshotCheck "preview.switch.excelIgnoresLink", loaded Is Nothing
    UpdateSectionResultsForWorkbook ThisWorkbook, preview ' Полный листовой fixture готовится явно: Excel сам не загружает ранее привязанный File.
    TestSnapshotCheck "preview.switch.excelDetached", Not store.HasFileReference(ThisWorkbook)
    TestSnapshotCheck "preview.switch.excelExact", TestSnapshotSameMatrix(boundaries, ReadPhysicalResultTable(ThisWorkbook, "rngNDMElementGeometry"))
    TestSnapshotSetting "Results.SnapshotStorage", "File"
    UpdateSectionResultsForWorkbook ThisWorkbook
    TestSnapshotCheck "preview.switch.fileHeaderOnly", UBound(ReadPhysicalResultTable(ThisWorkbook, "rngNDMElementGeometry"), 1) = 1
    TestSnapshotCheck "preview.switch.plotExact", TestSnapshotPlot() = beforePlot
    TestSnapshotCheck "preview.switch.exportExact", Audit02ReadExportSnapshotForTests(ThisWorkbook) = beforeExport
    TestSnapshotCheck "preview.readScopeReleased", TestCalculationSnapshotReadReleased()
    Dim compact As CSectionModel: Set compact = New CSectionModel: compact.SourceType = "Generated"
    compact.AddConcreteElement 0#, 0#, 200#, 1, shapeType:="Rectangle", width:=20#, height:=10#, rotationKnown:=True
    Dim adaptivePreview As CCalculationSnapshot
    Set adaptivePreview = writer.PrepareGeometryPreviewSnapshot(ThisWorkbook, compact, PrepareSectionSnapshot(compact), units)
    TestSnapshotCheck "preview.adaptive.noBoundaryRows", adaptivePreview.TableRowCount("rngNDMElementGeometry") = 1
    store.SaveFile ThisWorkbook, adaptivePreview
    file.OpenRead path
    TestSnapshotCheck "preview.adaptive.requires05", file.ReadASCII(8) = "NDMSNP05" And file.ReadLong() = 5
    file.CloseFile
    SaveAndDisplayCalculationSnapshot ThisWorkbook, preview ' Внешний reopen gate получает тот же точный контурный набор.
    If Not original Is Nothing Then SaveAndDisplayCalculationSnapshot ThisWorkbook, original
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Пустая физическая таблица допустима до первого расчёта; сравнение не объявляет её матрицей.
Private Function TestSnapshotOptionalMatrixEqual(ByRef expected As Variant, ByRef actual As Variant) As Boolean
    If IsEmpty(expected) Or IsEmpty(actual) Then
        TestSnapshotOptionalMatrixEqual = IsEmpty(expected) And IsEmpty(actual)
    Else
        TestSnapshotOptionalMatrixEqual = TestSnapshotSameMatrix(expected, actual)
    End If
End Function

 ' -------------------- ДЛЯ ТЕСТОВ --------------------
' Потерянный внешний источник не заменяется пустой таблицей или эквивалентным прямоугольником.
Private Sub TestSnapshotMissingGeometryFile(ByVal store As CCalculationSnapshotStore)
    Dim path As String: path = store.SnapshotPath(ThisWorkbook)
    Dim hidden As String: hidden = path & ".TestMissing"
    Dim fs As Object: Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(hidden) Then Err.Raise 5, , "Чужой файл занимает путь отрицательного теста отсутствия геометрии."
    Dim before As Variant: before = ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionGeometry")
    Dim number As Long, failed As Long, reason As String
    fs.MoveFile path, hidden
    On Error GoTo Cleanup
    Dim model As CSectionModel
    On Error Resume Next
    Set model = ReadFormedSectionFromResults(ThisWorkbook, False)
    number = Err.Number: Err.Clear
    On Error GoTo Cleanup
    TestSnapshotCheck "preview.missingFile.noRecalcFallback", number <> 0 And model Is Nothing
    On Error Resume Next
    UpdateSectionResultsForWorkbook ThisWorkbook
    number = Err.Number: Err.Clear
    On Error GoTo Cleanup
    TestSnapshotCheck "preview.missingFile.noRestoreFallback", number <> 0
    TestSnapshotCheck "preview.missingFile.keepsResults", TestSnapshotOptionalMatrixEqual(before, ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionGeometry"))
    TestSnapshotCheck "preview.missingFile.keepsLink", store.HasFileReference(ThisWorkbook)
    GoTo Cleanup
Cleanup:
    failed = Err.Number: reason = Err.Description
    fs.MoveFile hidden, path
    If failed <> 0 Then Err.Raise failed, "TestSnapshotMissingGeometryFile", reason
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Отличает договорной отказ повреждённой геометрии от посторонней runtime-ошибки.
Private Function TestSnapshotBoundaryRejected(ByRef boundaries As Variant, ByRef geometry As Variant) As Boolean
    On Error GoTo Rejected
    Dim map As Object: Set map = ReadConcreteBoundaryMap(boundaries, geometry)
    Exit Function
Rejected:
    TestSnapshotBoundaryRejected = Err.Number = vbObjectError + 4383 And InStr(Err.Description, "rngNDM") > 0
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет матрицу фактических границ кодеком .ndm и совместимым расширением .ndms.
' Остальные матрицы берутся из настоящего расчёта: тест меняет только записи
' геометрии для проверки хранения, новых напряжений по этому набору не считает.
Private Sub TestSnapshotBoundaryFileCodec(ByVal original As CCalculationSnapshot, ByVal store As CCalculationSnapshotStore)
    Dim geometry As Variant, boundaries As Variant, expected As Collection, sample As CCalculationSnapshot, key As Variant
    TestSnapshotBoundaryFixture "m", geometry, boundaries, expected
    Dim row As Long
    For row = 2 To UBound(geometry, 1): geometry(row, 1) = original.RunID: Next row
    For row = 2 To UBound(boundaries, 1): boundaries(row, 1) = original.RunID: Next row
    Set sample = New CCalculationSnapshot: sample.RunID = original.RunID: sample.ElapsedSeconds = original.ElapsedSeconds
    For Each key In original.Keys()
        Select Case CStr(key)
            Case "rngNDMSectionGeometry": sample.AddTable CStr(key), geometry
            Case "rngNDMElementGeometry"
            Case Else: sample.AddTable CStr(key), original.Table(CStr(key))
        End Select
    Next key
    sample.AddTable "rngNDMElementGeometry", boundaries
    Dim missing As CCalculationSnapshot: Set missing = New CCalculationSnapshot
    missing.RunID = sample.RunID: missing.ElapsedSeconds = sample.ElapsedSeconds
    For Each key In sample.Keys()
        If CStr(key) <> "rngNDMElementGeometry" Then missing.AddTable CStr(key), sample.Table(CStr(key))
    Next key
    TestSnapshotDirectBoundaryRejection missing, original
    Dim savedPath As String: savedPath = store.SnapshotPath(ThisWorkbook)
    Dim legacyPath As String: legacyPath = Left$(savedPath, Len(savedPath) - 4) & ".ndms"
    Dim fileSystem As Object: Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Dim ownsLegacy As Boolean
    If fileSystem.FileExists(legacyPath) Then Err.Raise 5, , "Чужой файл занимает путь проверки .ndms."
    On Error GoTo Failed
    Dim legacy As CCalculationSnapshot: Set legacy = New CCalculationSnapshot
    legacy.RunID = original.RunID: legacy.ElapsedSeconds = original.ElapsedSeconds
    Dim legacyGeometry As Variant: legacyGeometry = original.Table("rngNDMSectionGeometry")
    For row = 2 To UBound(legacyGeometry, 1)
        If CStr(legacyGeometry(row, 3)) = "Concrete" Then legacyGeometry(row, 7) = "Region"
    Next row
    For Each key In original.Keys()
        Select Case CStr(key)
            Case "rngNDMSectionGeometry": legacy.AddTable CStr(key), legacyGeometry
            Case "rngNDMElementGeometry"
            Case Else: legacy.AddTable CStr(key), original.Table(CStr(key))
        End Select
    Next key
    store.SaveFile ThisWorkbook, legacy
    Dim file As CBinarySnapshotFile: Set file = New CBinarySnapshotFile
    file.OpenRead savedPath
    TestSnapshotCheck "boundary.file.legacy04Header", file.ReadASCII(8) = "NDMSNP04" And file.ReadLong() = 4
    file.CloseFile
    Dim loaded As CCalculationSnapshot: Set loaded = store.LoadFile(ThisWorkbook)
    TestSnapshotCheck "boundary.file.legacy04Read", Not loaded.HasTable("rngNDMElementGeometry") And TestSnapshotSameMatrix(legacyGeometry, loaded.Table("rngNDMSectionGeometry"))
    store.SaveFile ThisWorkbook, sample
    file.OpenRead savedPath
    Dim signature As String: signature = file.ReadASCII(8)
    Dim version As Long: version = file.ReadLong()
    TestSnapshotCheck "boundary.file.new05Header", signature = "NDMSNP05" And version = 5
    TestSnapshotCheck "boundary.file.old04ReaderRejectsNew", signature <> "NDMSNP04" And version <> 4
    file.Position = file.Size - 8
    TestSnapshotCheck "boundary.file.new05Footer", file.ReadASCII(8) = signature
    file.CloseFile
    store.UpdateSavedElapsed ThisWorkbook, sample
    Set loaded = store.LoadFile(ThisWorkbook)
    TestSnapshotCheck "boundary.file.ndmMatrixExact", TestSnapshotSameMatrix(boundaries, loaded.Table("rngNDMElementGeometry"))
    TestSnapshotCheck "boundary.file.elapsed05", loaded.ElapsedSeconds = sample.ElapsedSeconds
    Dim map As Object, loadedGeometry As Variant, loadedBoundaries As Variant
    loadedGeometry = loaded.Table("rngNDMSectionGeometry"): loadedBoundaries = loaded.Table("rngNDMElementGeometry")
    Set map = ReadConcreteBoundaryMap(loadedBoundaries, loadedGeometry)
    TestSnapshotCheck "boundary.file.ndmGeometry", map.Count = 4
    Dim writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
    writer.RestoreSnapshotBase ThisWorkbook, loaded, False, False
    store.LinkFile ThisWorkbook
    TestSnapshotCheck "boundary.file.sheetHeaderOnly", UBound(ReadPhysicalResultTable(ThisWorkbook, "rngNDMElementGeometry"), 1) = 1
    Dim restored As CSectionModel: Set restored = ReadFormedSectionFromResults(ThisWorkbook, False)
    Dim i As Long, query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    TestSnapshotCheck "boundary.file.atomicRestore.count", restored.ConcreteCount = expected.Count
    For i = 1 To expected.Count
        TestSnapshotCheck "boundary.file.atomicRestore.region.C" & CStr(i), query.RegionsCoincide(expected(i), restored.ConcreteGeometry(i))
        TestSnapshotCheck "boundary.file.atomicRestore.marker.C" & CStr(i), restored.ConcreteHasBoundary(i) And restored.ConcreteShapeType(i) = "Contour"
        query.ValidateBoundaryProperties restored.ConcreteGeometry(i), restored.ConcreteX(i), restored.ConcreteY(i), _
            restored.ConcreteArea(i), restored.ConcreteLocalIx(i), restored.ConcreteLocalIy(i), restored.ConcreteLocalIxy(i)
    Next i
    fileSystem.MoveFile savedPath, legacyPath: ownsLegacy = True ' Как в TestSnapshotLegacyFile: BSTR-путь сохраняет символ α независимо от системной ANSI-страницы.
    Set loaded = store.LoadFile(ThisWorkbook)
    TestSnapshotCheck "boundary.file.ndmsMatrixExact", TestSnapshotSameMatrix(boundaries, loaded.Table("rngNDMElementGeometry"))
    fileSystem.MoveFile legacyPath, savedPath: ownsLegacy = False
    file.OpenUpdate savedPath: file.Position = file.Size - 8: file.WriteASCII "NDMSNP04": file.CloseFile
    TestSnapshotCheck "boundary.file.mismatchedFooterRejected", TestSnapshotFileRejected(store)
    writer.RestoreSnapshotBase ThisWorkbook, original, False, False
    store.SaveFile ThisWorkbook, original
    store.LinkFile ThisWorkbook
    Exit Sub
Failed:
    Dim number As Long, reason As String: number = Err.Number: reason = Err.Description
    On Error Resume Next
    If ownsLegacy And fileSystem.FileExists(legacyPath) Then
        If Not fileSystem.FileExists(savedPath) Then fileSystem.MoveFile legacyPath, savedPath
    End If
    If Not writer Is Nothing Then writer.RestoreSnapshotBase ThisWorkbook, original, False, False
    store.SaveFile ThisWorkbook, original
    store.LinkFile ThisWorkbook
    On Error GoTo 0
    Err.Raise number, "TestSnapshotBoundaryFileCodec", reason
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Напрямую вызывает публичное восстановление с настоящим Contour-снимком,
' у которого удалена обязательная матрица. Все семь физических таблиц Results
' должны остаться побитно неизменными; при регрессии тест восстанавливает оригинал.
Private Sub TestSnapshotDirectBoundaryRejection(ByVal invalid As CCalculationSnapshot, ByVal original As CCalculationSnapshot)
    Dim names As Variant: names = Array("rngNDMElementResults", "rngNDMSectionGeometry", "rngNDMSectionContours", _
        "rngNDMSectionProperties", "rngNDMMaterialDiagrams", "rngNDMSectionAnnotations", "rngNDMElementGeometry")
    Dim before As Collection: Set before = New Collection
    Dim i As Long, values As Variant
    For i = 0 To UBound(names)
        values = ReadPhysicalResultTable(ThisWorkbook, CStr(names(i))): before.Add values
    Next i
    Dim writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
    Dim number As Long, reason As String
    On Error Resume Next
    writer.RestoreSnapshotBase ThisWorkbook, invalid
    number = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo 0
    TestSnapshotCheck "boundary.restore.missingRejected", number = vbObjectError + 4383 And InStr(reason, "rngNDMElementGeometry") > 0
    Dim unchanged As Boolean, allUnchanged As Boolean, expected As Variant: allUnchanged = True
    For i = 0 To UBound(names)
        values = ReadPhysicalResultTable(ThisWorkbook, CStr(names(i)))
        expected = before.Item(i + 1)
        If IsEmpty(expected) Or IsEmpty(values) Then
            unchanged = IsEmpty(expected) And IsEmpty(values) ' В File физическая таблица состояний отсутствует; отсутствие с обеих сторон также должно сохраниться точно.
            mTestReport = mTestReport & "STATE boundary.restore." & CStr(names(i)) & "; expectedEmpty=" & CStr(IsEmpty(expected)) & "; actualEmpty=" & CStr(IsEmpty(values)) & vbCrLf
        Else
            unchanged = TestSnapshotSameMatrix(expected, values)
        End If
        TestSnapshotCheck "boundary.restore.keeps." & CStr(names(i)), unchanged
        allUnchanged = allUnchanged And unchanged
    Next i
    If number = 0 Or Not allUnchanged Then UpdateSectionResultsForWorkbook ThisWorkbook, original
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет пределы листов .xls и .xlsm, большие оценки и точную классификацию ошибки объединения.
Private Sub TestSnapshotOutputGuards()
    Dim limit As Variant
    For Each limit In Array(65536, ThisWorkbook.Worksheets("Results").Rows.Count)
        TestSnapshotCheck "rows.exact." & CStr(limit), ResultRowsFitSheet(156, CDbl(limit) - 155#, CLng(limit))
        TestSnapshotCheck "rows.offsetOverflow." & CStr(limit), Not ResultRowsFitSheet(156, CDbl(limit) - 154#, CLng(limit))
        TestSnapshotCheck "rows.totalOverflow." & CStr(limit), Not ResultRowsFitSheet(1, CDbl(limit) + 1#, CLng(limit))
        TestSnapshotCheck "rows.largeDouble." & CStr(limit), Not ResultRowsFitSheet(156, 5000000000#, CLng(limit))
    Next limit
    Dim writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
    TestSnapshotCheck "rows.estimateNoLongOverflow", writer.EstimatedElementResultRows(1000000, 5000) = 5000000001#
    Dim russian As String, english As String
    russian = ResultOutputErrorDescription("Мы не можем это сделать в объединенной ячейке.")
    english = ResultOutputErrorDescription("We can't do that to a merged cell.")
    TestSnapshotCheck "merged.russian", InStr(russian, "Не удалось вывести результаты на лист Results") > 0 And InStr(russian, "Снимите объединение") > 0
    TestSnapshotCheck "merged.english", english = russian
    TestSnapshotCheck "merged.other1004Unchanged", ResultOutputErrorDescription("Application-defined or object-defined error") = "Application-defined or object-defined error"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет политику крупных таблиц для обоих способов геометрии и хранилищ.
' Матрицы оценок не выделяют десятки тысяч строк и не показывают диалоги.
Private Sub TestSnapshotOutputPolicies()
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Dim originalMesh As String: originalMesh = settings.GetRequiredString("Mesh.CellGeometry")
    Dim originalMode As String: originalMode = settings.GetRequiredString("Results.SnapshotStorage")
    Dim mesh As Variant, mode As Variant, rows As Variant, label As String, expected As String
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    Dim number As Long, reason As String, limit As Long: limit = ThisWorkbook.Worksheets("Results").Rows.Count
    On Error GoTo Failed
    For Each mesh In Array("Adaptive", "Contour")
        TestSnapshotSetting "Mesh.CellGeometry", CStr(mesh)
        For Each mode In Array("Excel", "File")
            TestSnapshotSetting "Results.SnapshotStorage", CStr(mode)
            label = "rows.policy." & CStr(mesh) & "." & CStr(mode)
            rows = Array(Array("rngNDMSectionGeometry", 1001#, 156), Array("rngNDMElementGeometry", 1#, 156))
            TestSnapshotCheck label & ".small", CalculationSnapshotOutputPolicy(store.StorageMode(ThisWorkbook), 60000#, rows, limit) = ""
            expected = "": If CStr(mode) = "Excel" Then expected = "File"
            TestSnapshotCheck label & ".stateFile", CalculationSnapshotOutputPolicy(store.StorageMode(ThisWorkbook), 60001#, rows, limit) = expected
            rows = Array(Array("rngNDMSectionGeometry", 60001#, 156), Array("rngNDMElementGeometry", 1#, 156))
            TestSnapshotCheck label & ".mandatory", CalculationSnapshotOutputPolicy(store.StorageMode(ThisWorkbook), 100#, rows, limit) = "Mandatory"
            expected = "Mandatory": If CStr(mode) = "Excel" Then expected = "FileAndMandatory"
            TestSnapshotCheck label & ".both", CalculationSnapshotOutputPolicy(store.StorageMode(ThisWorkbook), 60001#, rows, limit) = expected
            rows = Array(Array("rngNDMSectionGeometry", CDbl(limit) - 155#, 156), Array("rngNDMElementGeometry", 1#, 156))
            TestSnapshotCheck label & ".exactLimit", CalculationSnapshotOutputPolicy(store.StorageMode(ThisWorkbook), 100#, rows, limit) = "Mandatory"
            rows = Array(Array("rngNDMSectionGeometry", 1001#, 156), Array("rngNDMElementGeometry", 1#, 156))
            expected = "": If CStr(mode) = "Excel" Then expected = "File"
            TestSnapshotCheck label & ".boundaryLargeFile", CalculationSnapshotOutputPolicy(store.StorageMode(ThisWorkbook), 100#, rows, limit, 60001#) = expected
            TestSnapshotCheck label & ".boundaryOverflowFile", CalculationSnapshotOutputPolicy(store.StorageMode(ThisWorkbook), 100#, rows, limit, CDbl(limit) + 1#) = expected
            rows = Array(Array("rngNDMSectionGeometry", CDbl(limit) - 154#, 156))
            TestSnapshotCheck label & ".geometryOverflowNoFile", CalculationSnapshotOutputPolicy(store.StorageMode(ThisWorkbook), 60001#, rows, limit) = "RowLimit"
            OfferExternalCalculationSnapshot ThisWorkbook, 60001#, False, rows
            TestSnapshotCheck label & ".noninteractivePreserved", store.StorageMode(ThisWorkbook) = CStr(mode)
        Next mode
    Next mesh
    GoTo Finished
Failed:
    number = Err.Number: reason = Err.Description
Finished:
    TestSnapshotSetting "Mesh.CellGeometry", originalMesh
    TestSnapshotSetting "Results.SnapshotStorage", originalMode
    If number <> 0 Then Err.Raise number, "TestSnapshotOutputPolicies", reason
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Запускает только граничную политику предупреждений, без решения НДС и файлов.
Public Function RunSnapshotOutputPolicyTests() As String
    mTestReport = "": mTestPassed = 0: mTestFailed = 0
    On Error GoTo Failed
    TestSnapshotOutputGuards
    TestSnapshotOutputPolicies
    GoTo Finished
Failed:
    TestSnapshotCheck "rows.policy.runtime " & CStr(Err.Number) & ": " & Err.Description, False
Finished:
    RunSnapshotOutputPolicyTests = mTestReport & "TOTAL_SNAPSHOT_OUTPUT_POLICY: passed=" & CStr(mTestPassed) & "; failed=" & CStr(mTestFailed)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет отказ до solver у последней строки, частичное объединение и оформление перенесённой шапки.
Private Sub TestSnapshotNativeLayout(ByVal original As CCalculationSnapshot)
    TestSnapshotSetting "Results.SnapshotStorage", "Excel"
    UpdateSectionResultsForWorkbook ThisWorkbook, original ' Испытывается вместимость полноценного листового снимка, а не минимального справочного File-вывода.
    Dim name As Object: Set name = ThisWorkbook.Names.Item("rngNDMElementResults")
    Dim previous As String: previous = name.RefersTo
    Dim anchor As Object: Set anchor = name.RefersToRange
    Dim limit As Long: limit = anchor.Worksheet.Rows.Count
    name.RefersTo = "=" & anchor.Worksheet.Cells(limit - 2, anchor.Column).Address(External:=True)
    Dim solves As Long: solves = SectionEquilibriumSolveCount()
    Dim reason As String: reason = TestSnapshotCalculationError()
    TestSnapshotCheck "rows.native.preflight", InStr(reason, "Results.SnapshotStorage = File") > 0 And InStr(reason, CStr(limit)) > 0 And InStr(reason, "1)") > 0
    TestSnapshotCheck "rows.native.noSolver", SectionEquilibriumSolveCount() = solves
    reason = TestSnapshotPreparedRefreshError(original)
    TestSnapshotCheck "rows.native.restoreAdvice", InStr(reason, "1) На Config установите Results.SnapshotStorage = File") > 0
    Dim mandatoryName As Object: Set mandatoryName = ThisWorkbook.Names.Item("rngNDMMaterialDiagrams")
    Dim mandatoryReference As String: mandatoryReference = mandatoryName.RefersTo
    Dim mandatoryAnchor As Object: Set mandatoryAnchor = mandatoryName.RefersToRange
    mandatoryName.RefersTo = "=" & mandatoryAnchor.Worksheet.Cells(limit, mandatoryAnchor.Column).Address(External:=True)
    reason = TestSnapshotPreparedRefreshError(original)
    TestSnapshotCheck "rows.native.mandatoryRestoreNoFileAdvice", InStr(reason, "rngNDMElementResults") > 0 And InStr(reason, "Results.SnapshotStorage = File") = 0
    reason = TestSnapshotCalculationError()
    TestSnapshotCheck "rows.native.mandatoryPreflightNoFileAdvice", InStr(reason, "rngNDMMaterialDiagrams") > 0 And InStr(reason, "Results.SnapshotStorage = File") = 0
    TestSnapshotCheck "rows.native.mandatoryNoSolver", SectionEquilibriumSolveCount() = solves
    mandatoryName.RefersTo = mandatoryReference
    name.RefersTo = previous
    UpdateSectionResultsForWorkbook ThisWorkbook, original
    TestSnapshotUpper original
    Set name = ThisWorkbook.Names.Item("rngNDMSectionProperties")
    previous = name.RefersTo: Set anchor = name.RefersToRange
    Dim common As Variant: common = original.Table("rngNDMSectionProperties")
    Dim cases As Variant: cases = original.Table("LoadCase.Properties")
    Dim staged() As Variant, row As Long, column As Long, nextRow As Long
    ReDim staged(1 To UBound(common, 1) + 2, 1 To 6)
    For row = 1 To UBound(common, 1)
        For column = 1 To 6: staged(row, column) = common(row, column): Next column
    Next row
    nextRow = UBound(common, 1) + 1
    For row = 2 To UBound(cases, 1)
        If CStr(cases(row, 3)) = "LoadReferenceX" Or CStr(cases(row, 3)) = "LoadReferenceY" Then
            For column = 1 To 6: staged(nextRow, column) = cases(row, column): Next column
            nextRow = nextRow + 1
        End If
    Next row
    name.RefersTo = "=" & anchor.Worksheet.Cells(limit - UBound(staged, 1) + 1, anchor.Column).Address(External:=True)
    name.RefersToRange.Resize(UBound(staged, 1), 6).Value2 = staged ' Сохраняем метаданные сформированного сечения: проверяется переполнение вывода, а не пустой источник геометрии.
    solves = SectionEquilibriumSolveCount()
    reason = TestSnapshotCalculationError()
    TestSnapshotCheck "rows.readableLC.preflight", InStr(reason, "rngNDMSectionProperties") > 0 And InStr(reason, CStr(limit)) > 0
    TestSnapshotCheck "rows.readableLC.noSolver", SectionEquilibriumSolveCount() = solves
    reason = TestSnapshotPreparedRefreshError(original)
    TestSnapshotCheck "rows.readableLC.prepared", InStr(reason, "rngNDMSectionProperties") > 0 And InStr(reason, CStr(limit)) > 0
    name.RefersToRange.Resize(UBound(staged, 1), 6).Clear
    name.RefersTo = previous
    UpdateSectionResultsForWorkbook ThisWorkbook, original
    TestSnapshotSetting "General.ReformatResultsEnabled", "No"
    Set name = ThisWorkbook.Names.Item("rngBatchSummary")
    previous = name.RefersTo
    Set anchor = ThisWorkbook.Worksheets("Results").Cells(name.RefersToRange.Row, 500)
    anchor.Offset(0, -1).Resize(1, 2).Merge
    name.RefersTo = "=" & anchor.Address(External:=True)
    reason = TestSnapshotPreparedRefreshError(original)
    TestSnapshotCheck "merged.native.refresh", InStr(reason, "область вывода попадает в объединённые ячейки") > 0
    reason = TestSnapshotCalculationError()
    TestSnapshotCheck "merged.native.calculate", InStr(reason, "область вывода попадает в объединённые ячейки") > 0
    anchor.Offset(0, -1).Resize(1, 2).UnMerge
    name.RefersTo = previous
    UpdateSectionResultsForWorkbook ThisWorkbook, original
    TestSnapshotSetting "General.ReformatResultsEnabled", "No"
    Set name = ThisWorkbook.Names.Item("rngBatchSummary")
    previous = name.RefersTo
    Set anchor = ThisWorkbook.Worksheets("Results").Cells(name.RefersToRange.Row, 500)
    name.RefersTo = "=" & anchor.Address(External:=True)
    UpdateSectionResultsForWorkbook ThisWorkbook, original
    TestSnapshotUpper original
    TestSnapshotCheck "minimal.movedHeaderCreated", anchor.Cells(6, 1).Font.Bold
    anchor.Cells(6, 1).Font.Color = RGB(12, 34, 56)
    UpdateSectionResultsForWorkbook ThisWorkbook, original
    TestSnapshotCheck "minimal.recognizedHeaderPreserved", CLng(anchor.Cells(6, 1).Font.Color) = RGB(12, 34, 56)
    name.RefersTo = previous
    TestSnapshotSetting "General.ReformatResultsEnabled", "Yes"
    TestSnapshotSetting "Results.SnapshotStorage", "File"
    UpdateSectionResultsForWorkbook ThisWorkbook, original
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает реальную причину отказа расчёта в частной книге без изменения ожиданий.
Private Function TestSnapshotCalculationError() As String
    On Error GoTo Rejected
    Dim message As String: message = RunSectionCalculationForWorkbook(ThisWorkbook)
    Exit Function
Rejected:
    TestSnapshotCalculationError = Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Передаёт готовый снимок, чтобы проверить размещение без повторного чтения уже перенесённого якоря.
Private Function TestSnapshotPreparedRefreshError(ByVal original As CCalculationSnapshot) As String
    On Error GoTo Rejected
    UpdateSectionResultsForWorkbook ThisWorkbook, original
    Exit Function
Rejected:
    TestSnapshotPreparedRefreshError = Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Суффикс .ndms выбирается только при отсутствии .ndm; повреждение .ndm не маскируется файлом .ndms.
Private Sub TestSnapshotLegacyFile(ByVal original As CCalculationSnapshot, ByVal store As CCalculationSnapshotStore)
    Dim path As String: path = store.SnapshotPath(ThisWorkbook)
    Dim previous As String: previous = Left$(path, Len(path) - 4) & ".ndms"
    Dim fileSystem As Object: Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Dim checksum As CFileChecksum: Set checksum = New CFileChecksum
    store.SaveFile ThisWorkbook, original
    fileSystem.MoveFile path, previous ' FSO сохраняет Unicode имени; оператор VBA Name зависит от системной кодовой страницы.
    Dim hash As String: hash = checksum.FileSHA256(previous)
    TestSnapshotEqual "legacy.load", original, store.LoadFile(ThisWorkbook)
    store.LinkFile ThisWorkbook
    TestSnapshotEqual "legacy.link", original, store.LoadLinkedFile(ThisWorkbook)
    store.SaveFile ThisWorkbook, original
    TestSnapshotCheck "legacy.saveNewSuffix", fileSystem.FileExists(path) And checksum.FileSHA256(previous) = hash
    Dim file As CBinarySnapshotFile: Set file = New CBinarySnapshotFile
    Dim value As Byte: file.OpenUpdate path: file.Position = 95: value = file.ReadByte
    file.Position = 95: file.WriteByte value Xor 1: file.CloseFile
    TestSnapshotCheck "legacy.corruptPreferredRejected", TestSnapshotFileRejected(store)
    TestSnapshotCheck "legacy.oldFileUntouched", checksum.FileSHA256(previous) = hash
    store.SaveFile ThisWorkbook, original
    file.OpenUpdate previous: file.Position = 95: value = file.ReadByte
    file.Position = 95: file.WriteByte value Xor 1: file.CloseFile
    TestSnapshotEqual "legacy.preferredWins", original, store.LoadFile(ThisWorkbook)
    fileSystem.DeleteFile previous
    store.ClearReference ThisWorkbook
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Неудачное чтение выбранного файла возвращает ошибку без подстановки другого снимка.
Private Function TestSnapshotFileRejected(ByVal store As CCalculationSnapshotStore) As Boolean
    On Error GoTo Rejected
    Dim snapshot As CCalculationSnapshot: Set snapshot = store.LoadFile(ThisWorkbook)
    Exit Function
Rejected:
    TestSnapshotFileRejected = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимо заданные поля SP35 и точки нагрузки принадлежат LC-блоку, а не геометрии.
Private Sub TestSnapshotSeparatedProperties()
    Dim snapshot As CCalculationSnapshot: Set snapshot = New CCalculationSnapshot
    Dim data(1 To 7, 1 To 6) As Variant, keys As Variant, row As Long
    data(1, 1) = "RunID": data(1, 2) = "LoadCase": data(1, 3) = "Parameter"
    data(1, 4) = "Value": data(1, 5) = "Unit": data(1, 6) = "Comment"
    keys = Array("Gross.Area", "SP35.Slenderness1", "SP35.Slenderness2", "LoadReferenceX", "LoadReferenceY", "StateLoadN")
    For row = 2 To 7
        data(row, 1) = 1#: data(row, 2) = "ALL": data(row, 3) = keys(row - 2): data(row, 4) = CDbl(row)
    Next row
    data(7, 2) = "LC_01"
    snapshot.AddSeparatedTable "rngNDMSectionProperties", data
    Dim common As Variant, cases As Variant
    common = snapshot.Table("rngNDMSectionProperties"): cases = snapshot.Table("LoadCase.Properties")
    TestSnapshotCheck "separated.commonOnlyArea", UBound(common, 1) = 2 And common(2, 3) = "Gross.Area"
    TestSnapshotCheck "separated.slendernessInCases", UBound(cases, 1) = 6 And cases(2, 3) = "SP35.Slenderness1" And cases(3, 3) = "SP35.Slenderness2"
    TestSnapshotCheck "separated.logicalTableExact", TestSnapshotSameMatrix(data, snapshot.ResultTable("rngNDMSectionProperties"))
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Оба хранилища обязаны соблюдать выключенное переоформление общих таблиц Results.
Private Sub TestSnapshotMinimalFormatting()
    Dim mode As Variant, target As Object, message As String
    TestSnapshotSetting "General.ReformatResultsEnabled", "No"
    For Each mode In Array("Excel", "File")
        TestSnapshotSetting "Results.SnapshotStorage", CStr(mode)
        If CStr(mode) = "File" Then
            TestSnapshotSetting "Geometry.Source", "Generated" ' Тест оформления не зависит от оставшегося после импортного сценария AutoCAD-источника.
            message = FormSectionForWorkbook(ThisWorkbook) ' File рассчитывает геометрию своего файла, а не оставшиеся строки Excel.
        End If
        Set target = ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Cells(2, 4)
        target.Font.Color = RGB(12, 34, 56): target.NumberFormat = "0.0000000000"
        message = RunSectionCalculationForWorkbook(ThisWorkbook)
        TestSnapshotCheck "minimal." & CStr(mode) & ".font", CLng(target.Font.Color) = RGB(12, 34, 56)
        TestSnapshotCheck "minimal." & CStr(mode) & ".numberFormat", CStr(target.NumberFormat) = "0.0000000000"
    Next mode
    TestSnapshotSetting "General.ReformatResultsEnabled", "Yes"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимый байтовый эталон проверяет границу буфера, прямой большой блок и seek/read.
Private Sub TestSnapshotBinaryBuffer()
    Dim path As String: path = ThisWorkbook.Path & Application.PathSeparator & "TestSnapshotBuffer.bin"
    Dim file As CBinarySnapshotFile: Set file = New CBinarySnapshotFile
    On Error GoTo Failed
    Dim expected(0 To 131088) As Byte, actual() As Byte, row As Long, same As Boolean
    For row = 0 To UBound(expected): expected(row) = row Mod 251: Next row
    file.OpenWrite path: file.WriteASCII "NDMBUF01": file.WriteBytes expected
    For row = 1 To 20000: file.WriteLong row: Next row
    file.WriteDouble 1E-250
    TestSnapshotCheck "buffer.logicalSize", file.Size = 8 + 131089 + 80000 + 8 And file.Position = file.Size
    file.Position = 8 + 65535: file.WriteASCII "BUFFER": file.CloseFile
    For row = 0 To 5: expected(65535 + row) = Asc(Mid$("BUFFER", row + 1, 1)): Next row
    file.OpenRead path
    TestSnapshotCheck "buffer.magic", file.ReadASCII(8) = "NDMBUF01"
    actual = file.ReadBytes(131089): same = True
    For row = 0 To UBound(expected)
        If actual(row) <> expected(row) Then same = False: Exit For
    Next row
    TestSnapshotCheck "buffer.largeBlockAndSeek", same
    same = True
    For row = 1 To 20000
        If file.ReadLong <> row Then same = False
    Next row
    TestSnapshotCheck "buffer.smallWrites", same
    TestSnapshotCheck "buffer.doubleExact", TestSnapshotDouble(file.ReadDouble) = TestSnapshotDouble(1E-250)
    TestSnapshotCheck "buffer.eofRejected", TestSnapshotBinaryEOFRejected(file)
    file.CloseFile: file.OpenUpdate path: file.Position = 0: file.WriteByte 42: file.Position = 0
    TestSnapshotCheck "buffer.flushBeforeRead", file.ReadByte = 42
    file.WriteByte 99 ' После короткого чтения Windows-указатель находится дальше Position из-за предварительного чтения.
    TestSnapshotCheck "buffer.writeAtLogicalReadPosition", file.ReadByte = Asc("M") And file.Position = 3
    file.Position = 0
    TestSnapshotCheck "buffer.readAfterMixedWrite", file.ReadByte = 42 And file.ReadByte = 99
    Dim chunk As Variant
    For Each chunk In Array(127&, 32767&, 65535, 65536, 65537, 131089)
        file.Position = 8: actual = file.ReadBytes(CLng(chunk)): same = True
        For row = 0 To UBound(actual)
            If actual(row) <> expected(row) Then same = False: Exit For
        Next row
        TestSnapshotCheck "buffer.readChunk." & CStr(chunk), same And file.Position = 8 + CLng(chunk)
    Next chunk
    file.Position = 8 + 65534: actual = file.ReadBytes(19): same = True
    For row = 0 To UBound(actual)
        If actual(row) <> expected(65534 + row) Then same = False: Exit For
    Next row
    TestSnapshotCheck "buffer.backwardSeekAndBoundary", same
    file.CloseFile
    CreateObject("Scripting.FileSystemObject").DeleteFile path
    Exit Sub
Failed:
    Dim number As Long, reason As String: number = Err.Number: reason = Err.Description
    On Error Resume Next
    file.CloseFile
    If CreateObject("Scripting.FileSystemObject").FileExists(path) Then CreateObject("Scripting.FileSystemObject").DeleteFile path
    On Error GoTo 0
    Err.Raise number, "TestSnapshotBinaryBuffer", reason
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Чтение за физическим концом файла обязано завершиться ошибкой.
Private Function TestSnapshotBinaryEOFRejected(ByVal file As CBinarySnapshotFile) As Boolean
    On Error GoTo Rejected
    Dim value As Byte: value = file.ReadByte
    Exit Function
Rejected:
    TestSnapshotBinaryEOFRejected = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ускоренное чтение границ сохраняет первую пустую ячейку, пробелы и точный maxCount.
Private Sub TestSnapshotAdaptiveExtent()
    Dim sheet As Object, app As Object, alerts As Boolean
    Set app = ThisWorkbook.Application: alerts = app.DisplayAlerts
    On Error GoTo Failed
    Set sheet = ThisWorkbook.Worksheets.Add
    Dim values() As Variant, row As Long
    ReDim values(1 To 100001, 1 To 1)
    For row = 1 To UBound(values, 1): values(row, 1) = row: Next row
    sheet.Cells(1, 1).Resize(UBound(values, 1), 1).Value2 = values
    sheet.Cells(100003, 1).Value2 = CVErr(2015) ' Посторонняя ошибка за первой пустой ячейкой не входит в таблицу.
    TestSnapshotCheck "extent.large", SnapshotAnchoredExtent(sheet.Cells(1, 1), False) = 100001
    TestSnapshotCheck "extent.maxCount", SnapshotAnchoredExtent(sheet.Cells(1, 1), False, 65537) = 65537
    sheet.Cells(4097, 1).Value2 = " "
    TestSnapshotCheck "extent.spaceAtGrowingBoundary", SnapshotAnchoredExtent(sheet.Cells(1, 1), False) = 4096
    sheet.Cells(1024, 1).ClearContents
    TestSnapshotCheck "extent.emptyAtInitialBoundary", SnapshotAnchoredExtent(sheet.Cells(1, 1), False) = 1023
    sheet.Cells(1, 1).ClearContents
    TestSnapshotCheck "extent.emptyAnchor", SnapshotAnchoredExtent(sheet.Cells(1, 1), False) = 0
    ReDim values(1 To 1, 1 To 4098)
    For row = 1 To UBound(values, 2): values(1, row) = "Header" & CStr(row): Next row
    sheet.Cells(2, 2).Resize(1, UBound(values, 2)).Value2 = values
    sheet.Cells(2, 2 + 3072).ClearContents
    TestSnapshotCheck "extent.columns", SnapshotAnchoredExtent(sheet.Cells(2, 2), True) = 3072
    app.DisplayAlerts = False: sheet.Delete: app.DisplayAlerts = alerts
    Exit Sub
Failed:
    Dim number As Long, description As String: number = Err.Number: description = Err.Description
    On Error Resume Next
    app.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    app.DisplayAlerts = alerts
    On Error GoTo 0
    Err.Raise number, "TestSnapshotAdaptiveExtent", description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Объединение цветных интервалов сохраняет пропуски и работает без чтения заполненных ячеек.
Private Sub TestSnapshotStatusSpans()
    Dim sheet As Object, app As Object, alerts As Boolean
    Set app = ThisWorkbook.Application: alerts = app.DisplayAlerts
    On Error GoTo Failed
    Set sheet = ThisWorkbook.Worksheets.Add
    Dim target As Object: Set target = sheet.Cells(4, 5).Resize(13, 3)
    Dim values(1 To 13, 1 To 3) As Variant, body(1 To 11, 1 To 3) As Variant
    Dim rows(1 To 8) As Long, ids As Variant, row As Long, column As Long, testCase As Long, i As Long, selected As Boolean, same As Boolean
    ids = Array(3&, 4&, 10&, 6&, 7&, 8&, 13&, 12&)
    For i = 1 To 8: rows(i) = CLng(ids(i - 1)): Next i
    For row = 1 To 13
        values(row, 1) = "OK": values(row, 2) = "FAIL": values(row, 3) = "N/A"
    Next row
    values(4, 3) = "Cracked": values(10, 3) = "NotCracked": values(6, 3) = ""
    For row = 1 To 11
        For column = 1 To 3: body(row, column) = values(row + 2, column): Next column
    Next row
    For testCase = 0 To 2
        target.Interior.Color = RGB(1, 2, 3)
        If testCase = 0 Then
            target.Value2 = values: ApplyTableStatusFills target, rows, Array(1, 2, 3), True, 3
        ElseIf testCase = 1 Then
            target.Value2 = CVErr(2015): ApplyTableStatusFills target, rows, Array(1, 2, 3), True, 3, values
        Else
            target.Value2 = CVErr(2015): ApplyTableStatusFills target, rows, Array(1, 2, 3), True, 3, body, 2
        End If
        same = True
        For row = 1 To 13
            selected = False
            For i = 1 To 8: If rows(i) = row Then selected = True
            Next i
            For column = 1 To 3
                Dim color As Long: color = RGB(1, 2, 3)
                If selected Then
                    If Not (column = 3 And (row = 4 Or row = 6 Or row = 10)) Then color = StatusFillColor(CStr(values(row, column)))
                End If
                If CLng(target.Cells(row, column).Interior.Color) <> color Then same = False
            Next column
        Next row
        TestSnapshotCheck "statusSpans.gapsAndPreparedValues." & CStr(testCase), same
    Next testCase
    app.DisplayAlerts = False: sheet.Delete: app.DisplayAlerts = alerts
    Exit Sub
Failed:
    Dim number As Long, description As String: number = Err.Number: description = Err.Description
    On Error Resume Next
    app.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    app.DisplayAlerts = alerts
    On Error GoTo 0
    Err.Raise number, "TestSnapshotStatusSpans", description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет прямой строковый IO независимо от кодека: побайтовый интерфейс служит независимым эталоном.
Private Sub TestSnapshotUTF16Strings()
    Dim path As String: path = ThisWorkbook.Path & Application.PathSeparator & "TestUTF16_" & Format$(Now, "yyyymmdd_hhnnss") & ".bin"
    If Len(Dir$(path)) > 0 Then Err.Raise 5, , "Тестовый файл UTF16 уже существует."
    Dim file As CBinarySnapshotFile: Set file = New CBinarySnapshotFile
    On Error GoTo Failed
    Dim samples As Variant, sample As Variant, bytes() As Byte, actual As String, count As Long, index As Long
    samples = Array(vbNullString, "ASCII", ChrW$(&H416) & ChrW$(&H3B1), "left" & Chr$(0) & "right", _
        ChrW$(&HD83D) & ChrW$(&HDE00), ChrW$(&HD800), String$(32767, "x"), String$(65536, "z"))
    file.OpenWrite path
    For Each sample In samples: file.WriteSizedUTF16 CStr(sample): Next sample
    file.CloseFile: file.OpenRead path
    For Each sample In samples
        count = file.ReadLong: actual = vbNullString
        If count > 0 Then bytes = file.ReadBytes(count): actual = bytes
        TestSnapshotCheck "utf16.write.matchesBytes." & CStr(index), count = LenB(CStr(sample)) And StrComp(actual, CStr(sample), vbBinaryCompare) = 0
        index = index + 1
    Next sample
    file.CloseFile: file.OpenWrite path
    For Each sample In samples
        file.WriteLong LenB(CStr(sample))
        If LenB(CStr(sample)) > 0 Then bytes = CStr(sample): file.WriteBytes bytes
    Next sample
    file.CloseFile: file.OpenRead path: index = 0
    For Each sample In samples
        actual = file.ReadSizedUTF16(131072)
        TestSnapshotCheck "utf16.read.matchesBytes." & CStr(index), LenB(actual) = LenB(CStr(sample)) And StrComp(actual, CStr(sample), vbBinaryCompare) = 0
        index = index + 1
    Next sample
    file.CloseFile
    Dim testCase As Long
    For testCase = 0 To 3
        file.OpenWrite path
        Select Case testCase
            Case 0: file.WriteLong -1
            Case 1: file.WriteLong 3: file.WriteASCII "abc"
            Case 2: file.WriteSizedUTF16 String$(65537, "z")
            Case 3: file.WriteLong 20: file.WriteByte 7
        End Select
        file.CloseFile: file.OpenRead path
        TestSnapshotCheck "utf16.invalidLength." & CStr(testCase), TestSnapshotUTF16Rejected(file)
        file.CloseFile
    Next testCase
    Kill path
    Exit Sub
Failed:
    Dim number As Long, reason As String: number = Err.Number: reason = Err.Description
    On Error Resume Next
    file.CloseFile: Kill path
    On Error GoTo 0
    Err.Raise number, "TestSnapshotUTF16Strings", reason
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Перехватывает только ожидаемую ошибку чтения некорректной длины строки.
Private Function TestSnapshotUTF16Rejected(ByVal file As CBinarySnapshotFile) As Boolean
    On Error GoTo Rejected
    Dim text As String: text = file.ReadSizedUTF16(131072)
    Exit Function
Rejected:
    TestSnapshotUTF16Rejected = (Err.Number = 5)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторное использование writer-а читает новые входные строки, а не массив предыдущей операции.
Private Sub TestSnapshotOneUseSource()
    Dim target As Object, saved As Variant
    Set target = ThisWorkbook.Names("rngNDMSectionGeometry").RefersToRange.Cells(2, 15)
    saved = target.Formula
    On Error GoTo Failed
    Dim sourceStore As CCalculationSnapshotStore: Set sourceStore = New CCalculationSnapshotStore
    Dim previousMode As String: previousMode = sourceStore.StorageMode(ThisWorkbook)
    Dim wasLinked As Boolean: wasLinked = sourceStore.HasFileReference(ThisWorkbook)
    Dim sourceSnapshot As CCalculationSnapshot
    If wasLinked Then Set sourceSnapshot = sourceStore.LoadFile(ThisWorkbook)
    TestSnapshotSetting "Results.SnapshotStorage", "Excel" ' Этот тест проверяет обновление изменяемого Excel-источника; файл отдельно проверяется как неизменяемый источник.
    If wasLinked Then UpdateSectionResultsForWorkbook ThisWorkbook, sourceSnapshot ' Данные для отдельного Excel-теста записываются явно, без неявного файлового источника.
    Dim section As CSectionModel: Set section = ReadFormedSectionFromResults(ThisWorkbook)
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Dim units As CUnitSystem: Set units = New CUnitSystem: units.LoadFromSettings settings
    Dim provider As CMaterialModelProvider: Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Dim props As CSectionPropertiesCalculator: Set props = PrepareSectionSnapshot(section, provider)
    Dim writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
    Dim batch As CBatchSectionCalculator: Set batch = New CBatchSectionCalculator
    Dim first As CCalculationSnapshot: Set first = New CCalculationSnapshot
    target.Value2 = "TestSourceBefore"
    writer.PreparePreservedSnapshotSource ThisWorkbook
    Set writer.CaptureSnapshot = first
    writer.WriteResults ThisWorkbook, section, props, batch, units, True
    Dim values As Variant: values = first.Table("rngNDMSectionGeometry")
    TestSnapshotCheck "capture.oneUse.initial", CStr(values(2, 15)) = "TestSourceBefore"
    target.Value2 = "TestSourceAfter"
    Dim second As CCalculationSnapshot: Set second = New CCalculationSnapshot
    Set writer.CaptureSnapshot = second
    writer.WriteResults ThisWorkbook, section, props, batch, units, True
    values = second.Table("rngNDMSectionGeometry")
    TestSnapshotCheck "capture.oneUse.newSource", CStr(values(2, 15)) = "TestSourceAfter"
    target.Formula = saved
    TestSnapshotSetting "Results.SnapshotStorage", previousMode
    If wasLinked Then UpdateSectionResultsForWorkbook ThisWorkbook
    Exit Sub
Failed:
    Dim number As Long, reason As String: number = Err.Number: reason = Err.Description
    On Error Resume Next
    target.Formula = saved
    If Len(previousMode) > 0 Then TestSnapshotSetting "Results.SnapshotStorage", previousMode
    If wasLinked Then sourceStore.LinkFile ThisWorkbook
    On Error GoTo 0
    Err.Raise number, "TestSnapshotOneUseSource", reason
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Запускает отдельно те же File/Excel проверки источника, которые входят в Snapshot-приёмку.
Public Function RunFileGeometryAuthorityTests() As String
    mTestReport = "": mTestPassed = 0: mTestFailed = 0
    On Error GoTo Failed
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    TestSnapshotFileGeometryAuthority settings.GetRequiredString("General.DiagramExtension")
    RunFileGeometryAuthorityTests = mTestReport & "TOTAL_FILE_GEOMETRY_AUTHORITY: " & mTestPassed & "/" & mTestFailed
    Exit Function
Failed:
    RunFileGeometryAuthorityTests = mTestReport & "RUNTIME: " & Err.Number & " " & Err.Description & vbCrLf & _
        "TOTAL_FILE_GEOMETRY_AUTHORITY: " & mTestPassed & "/" & (mTestFailed + 1)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Удаляет все справочные геометрические таблицы и сравнивает File с неизменённым снимком.
' Проверяет preview, повторный расчёт всех проверок и независимый выбор File/Excel для обеих сеток.
Private Sub TestSnapshotFileGeometryAuthority(ByVal extension As String)
    TestSnapshotTableOrder
    TestSnapshotActionMessages
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    Dim geometryMode As Variant, token As String, message As String, path As String
    Dim original As CCalculationSnapshot, after As CCalculationSnapshot, preview As CCalculationSnapshot
    Dim beforeModel As CSectionModel, afterModel As CSectionModel, key As Variant, data As Variant, matrix As Variant
    Dim geometrySource As String, exportBefore As String, plotBefore As String
    For Each geometryMode In Array("Adaptive", "Contour")
        modTestPerformance.TestConfigurePerformanceFixture "MIXED", 2, "Newton", "Circle", "Normal", extension
        TestSnapshotSetting "Mesh.CellGeometry", CStr(geometryMode)
        TestSnapshotSetting "Results.SnapshotStorage", "File"
        token = "TestFileAuthority_" & CStr(geometryMode) & "_" & extension & "_" & Format$(Now, "yyyymmdd_hhnnss")
        TestSnapshotSetting "Results.SnapshotName", token
        TestSnapshotSetting "General.ExecutionReportEnabled", "No"
        message = FormSectionForWorkbook(ThisWorkbook)
        path = store.SnapshotPath(ThisWorkbook)
        TestSnapshotCheck CStr(geometryMode) & ".message.formed.File", InStr(message, "Сечение сформировано и сохранено в файле «" & token & ".ndm».") = 1
        Set preview = store.LoadFile(ThisWorkbook)
        store.ClearReference ThisWorkbook ' Имя и File достаточны и для GeometryPreview; предварительное обновление не требуется.
        Set beforeModel = ReadFormedSectionFromResults(ThisWorkbook, True)
        geometrySource = beforeModel.SourceType
        TestSnapshotEraseGeometryDisplay
        Set afterModel = ReadFormedSectionFromResults(ThisWorkbook, True)
        TestSnapshotCheck CStr(geometryMode) & ".preview.fileModel", TestSnapshotSameModel(beforeModel, afterModel)
        TestSnapshotCheck CStr(geometryMode) & ".preview.source", ResultsGeometrySource(ThisWorkbook) = geometrySource
        TestSnapshotCheck CStr(geometryMode) & ".preview.fileUnchanged", TestSnapshotSameMatrix(preview.Table("rngNDMSectionGeometry"), store.LoadFile(ThisWorkbook).Table("rngNDMSectionGeometry"))
        message = RunSectionCalculationForWorkbook(ThisWorkbook)
        Set original = store.LoadLinkedFile(ThisWorkbook)
        TestSnapshotCheck CStr(geometryMode) & ".message.calculated.source.File", InStr(message, "Геометрия для расчёта загружена из файла «" & token & ".ndm».") > 0
        TestSnapshotCheck CStr(geometryMode) & ".message.calculated.saved.File", Right$(message, Len("Данные сохранены в файле «" & token & ".ndm».")) = "Данные сохранены в файле «" & token & ".ndm»."
        TestSnapshotCheck CStr(geometryMode) & ".calculate.fromFilePreview", Not original.IsGeometryPreview And original.TableRowCount("rngNDMElementResults") > 1
        modTestPlotConfig.TestSetOutputCombination ThisWorkbook, "PERF_01"
        modTestPlotConfig.TestSetOutputCombination ThisWorkbook, "PERF_01"
        exportBefore = Audit02ReadExportSnapshotForTests(ThisWorkbook)
        plotBefore = TestSnapshotPlot()
        TestSnapshotEraseGeometryDisplay
        Set afterModel = ReadFormedSectionFromResults(ThisWorkbook, True)
        TestSnapshotCheck CStr(geometryMode) & ".calculated.fileModel", TestSnapshotSameModel(beforeModel, afterModel)
        TestSnapshotCheck CStr(geometryMode) & ".calculated.export", Audit02ReadExportSnapshotForTests(ThisWorkbook) = exportBefore
        TestSnapshotCheck CStr(geometryMode) & ".calculated.plot", TestSnapshotPlot() = plotBefore
        modTestPlotConfig.TestSetOutputCombination ThisWorkbook, "Worst"
        modTestPlotConfig.TestSetOutputCombination ThisWorkbook, "Worst"
        exportBefore = Audit02ReadExportSnapshotForTests(ThisWorkbook): plotBefore = TestSnapshotPlot()
        ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Cells(1, 1).Value2 = "FORGED_RESULT_ID"
        TestSnapshotCheck CStr(geometryMode) & ".worst.exportFromFile", Audit02ReadExportSnapshotForTests(ThisWorkbook) = exportBefore
        TestSnapshotCheck CStr(geometryMode) & ".worst.plotFromFile", TestSnapshotPlot() = plotBefore
        TestSnapshotFileViewControls CStr(geometryMode)
        UpdateSectionResultsForWorkbook ThisWorkbook
        TestSnapshotUpper original
        For Each key In Array("rngNDMSectionGeometry", "rngNDMSectionContours", "rngNDMSectionProperties", "rngNDMMaterialDiagrams")
            TestSnapshotCheck CStr(geometryMode) & ".display.refreshed." & CStr(key), TestSnapshotSameMatrix(original.Table(CStr(key)), ReadPhysicalResultTable(ThisWorkbook, CStr(key)))
        Next key
        TestSnapshotEraseGeometryDisplay
        message = RunSectionCalculationForWorkbook(ThisWorkbook)
        Set after = store.LoadLinkedFile(ThisWorkbook)
        For Each key In Array("rngNDMSectionGeometry", "rngNDMSectionContours", "rngNDMElementGeometry", "rngNDMElementResults", "rngNDMSectionProperties", "LoadCase.Properties", "rngNDMMaterialDiagrams", "rngNDMSectionAnnotations", "LoadCase.Annotations")
            TestSnapshotCheck CStr(geometryMode) & ".repeat." & CStr(key), TestSnapshotSameIgnoringRun(original.Table(CStr(key)), after.Table(CStr(key)))
        Next key
        TestSnapshotEraseGeometryDisplay
        TestSnapshotSetting "Results.SnapshotStorage", "Excel"
        TestSnapshotCheck CStr(geometryMode) & ".excel.noFileFallback", TestSnapshotFileReadRejected(0, vbObjectError + 4380)
        UpdateSectionResultsForWorkbook ThisWorkbook, after ' Подготовка полного листового снимка для отдельной проверки Excel, без скрытой загрузки файла этим режимом.
        message = RunSectionCalculationForWorkbook(ThisWorkbook)
        TestSnapshotCheck CStr(geometryMode) & ".message.calculated.source.Excel", InStr(message, "Геометрия для расчёта загружена с листа Results текущей книги.") > 0
        TestSnapshotCheck CStr(geometryMode) & ".message.calculated.saved.Excel", Right$(message, Len("Данные сохранены на листе Results текущей книги.")) = "Данные сохранены на листе Results текущей книги."
        TestSnapshotCheck CStr(geometryMode) & ".toExcel.referenceCleared", Not store.HasFileReference(ThisWorkbook)
        Set afterModel = ReadFormedSectionFromResults(ThisWorkbook, True)
        TestSnapshotCheck CStr(geometryMode) & ".toExcel.completeModel", TestSnapshotSameModel(beforeModel, afterModel)
        TestSnapshotCheck CStr(geometryMode) & ".toExcel.boundaries", TestSnapshotSameIgnoringRun(after.Table("rngNDMElementGeometry"), ReadPhysicalResultTable(ThisWorkbook, "rngNDMElementGeometry"))
        TestSnapshotCheck CStr(geometryMode) & ".scopeReleased", TestCalculationSnapshotReadReleased()
        TestSnapshotSetting "Results.SnapshotStorage", "File"
        UpdateSectionResultsForWorkbook ThisWorkbook
        For Each key In Array("rngNDMElementResults", "rngNDMElementGeometry", "rngNDMSectionAnnotations")
            data = ReadPhysicalResultTable(ThisWorkbook, CStr(key))
            TestSnapshotCheck CStr(geometryMode) & ".display.headerOnly." & CStr(key), UBound(data, 1) = 1
        Next key
        TestSnapshotSetting "Results.SnapshotStorage", "Excel"
        UpdateSectionResultsForWorkbook ThisWorkbook, after
        TestSnapshotSetting "Results.SnapshotStorage", "File"
        Set afterModel = ReadFormedSectionFromResults(ThisWorkbook, True)
        TestSnapshotCheck CStr(geometryMode) & ".file.noPreloadRequired", TestSnapshotSameModel(beforeModel, afterModel)
        message = RunSectionCalculationForWorkbook(ThisWorkbook)
        TestSnapshotCheck CStr(geometryMode) & ".file.directCalculation", store.LoadFile(ThisWorkbook).TableRowCount("rngNDMElementResults") > 1
        TestSnapshotSetting "Results.SnapshotName", token & "_missing"
        TestSnapshotCheck CStr(geometryMode) & ".saveAs.ignoresOtherName", TestSnapshotUnselectedLinkIgnored(store)
        store.ClearReference ThisWorkbook
        TestSnapshotCheck CStr(geometryMode) & ".noLinkedFile.modelRejected", TestSnapshotFileReadRejected(0)
        TestSnapshotCheck CStr(geometryMode) & ".noLinkedFile.plotRejected", TestSnapshotFileReadRejected(1)
        TestSnapshotCheck CStr(geometryMode) & ".noLinkedFile.exportRejected", TestSnapshotFileReadRejected(2)
        TestSnapshotMissingFileMessage store
        TestSnapshotSetting "Results.SnapshotName", token
        TestSnapshotSetting "Results.SnapshotStorage", "Excel"
        If CreateObject("Scripting.FileSystemObject").FileExists(path) Then Kill path
    Next geometryMode
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что SaveAs не сверяет и не копирует ранее показанный файл другого имени.
Private Function TestSnapshotUnselectedLinkIgnored(ByVal store As CCalculationSnapshotStore) As Boolean
    Dim hashProperty As Object: Set hashProperty = ThisWorkbook.CustomDocumentProperties.Item("NDM.CalculationSnapshot.Hash")
    Dim before As String: before = CStr(hashProperty.Value)
    On Error GoTo Finished
    hashProperty.Value = "TestUnselectedHash"
    store.EnsureBesideWorkbook ThisWorkbook
    TestSnapshotUnselectedLinkIgnored = True
Finished:
    hashProperty.Value = before
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет имя, действия пользователя и фактический адрес настройки при отсутствии выбранного файла.
Private Sub TestSnapshotMissingFileMessage(ByVal store As CCalculationSnapshotStore)
    Dim snapshot As CCalculationSnapshot, reason As String, number As Long
    On Error Resume Next
    Set snapshot = store.LoadFile(ThisWorkbook)
    reason = Err.Description: number = Err.Number
    On Error GoTo 0
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    TestSnapshotCheck "missingFile.error", number = vbObjectError + 4956
    TestSnapshotCheck "missingFile.name", InStr(reason, "«" & store.SelectedFileName(ThisWorkbook) & "» не найден.") > 0
    TestSnapshotCheck "missingFile.actions", InStr(reason, "Проверьте имя Results.SnapshotName") > 0 And InStr(reason, "«Сформировать сечение»") > 0
    TestSnapshotCheck "missingFile.location", InStr(reason, settings.InputLocation("Results.SnapshotName")) > 0
    TestSnapshotCheck "missingFile.path", InStr(reason, store.SnapshotPath(ThisWorkbook)) > 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет согласованный порядок семи нижних блоков и зазор между их реальными ширинами.
Private Sub TestSnapshotTableOrder()
    Dim key As Variant, anchor As Object, previousEnd As Long
    Dim writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
    For Each key In Array("rngNDMSectionGeometry", "rngNDMSectionProperties", "rngNDMMaterialDiagrams", "rngNDMSectionContours", "rngNDMElementResults", "rngNDMElementGeometry", "rngNDMSectionAnnotations")
        Set anchor = ThisWorkbook.Names.Item(CStr(key)).RefersToRange
        TestSnapshotCheck "layout.order." & CStr(key), anchor.Column > previousEnd
        previousEnd = anchor.Column + writer.OutputColumnCount(CStr(key)) + 1
    Next key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Удаляет всё содержимое Results, включая сводки и шапки, не трогая Config, имена и оформление.
Private Sub TestSnapshotEraseGeometryDisplay()
    ThisWorkbook.Worksheets.Item("Results").UsedRange.ClearContents
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Изменяет только просмотр: другие сохранённые LC/НДС, величину и точность; расчётные поля Config намеренно непригодны.
Private Sub TestSnapshotFileViewControls(ByVal prefix As String)
    Dim profiles As Object: Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Dim saved As Variant: saved = profiles.Formula
    Dim row As Long, column As Long, key As String
    For row = 1 To profiles.Rows.Count
        key = CStr(profiles.Cells(row, 2).Value2)
        For column = 3 To profiles.Columns.Count
            Select Case key
                Case "Visualization.State": profiles.Cells(row, column).Value2 = "CapacityState"
                Case "Visualization.Quantity": profiles.Cells(row, column).Value2 = "Strain"
                Case "Visualization.StrainPrecision": profiles.Cells(row, column).Value2 = 8
                Case "Calculation.Crack.Width", "MaterialModel.Strength.ValueSet": profiles.Cells(row, column).Value2 = "INVALID_FOR_NEW_CALCULATION"
            End Select
        Next column
    Next row
    modTestPlotConfig.TestSetOutputCombination ThisWorkbook, "PERF_02"
    modTestPlotConfig.TestSetOutputCombination ThisWorkbook, "PERF_02"
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Dim reader As CSectionPlotDataReader: Set reader = New CSectionPlotDataReader: reader.LoadFromWorkbook ThisWorkbook, settings
    TestSnapshotCheck prefix & ".view.currentLC", reader.LoadCase = "PERF_02"
    TestSnapshotCheck prefix & ".view.currentState", reader.StateType = "CapacityState"
    TestSnapshotCheck prefix & ".view.currentQuantity", reader.VisualizationQuantity = "Strain"
    TestSnapshotCheck prefix & ".view.currentPrecision", reader.ResultPrecision = 8
    TestSnapshotCheck prefix & ".view.currentExport", Left$(Audit02ReadExportSnapshotForTests(ThisWorkbook), Len("PERF_02|PR1|CapacityState|Strain|")) = "PERF_02|PR1|CapacityState|Strain|"
    profiles.Formula = saved
    modTestPlotConfig.TestSetOutputCombination ThisWorkbook, "Worst"
    modTestPlotConfig.TestSetOutputCombination ThisWorkbook, "Worst"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' File без привязки не имеет права подменять свой источник сохранившимися таблицами Excel.
Private Function TestSnapshotFileReadRejected(ByVal consumer As Long, Optional ByVal expectedError As Long = 0) As Boolean
    On Error GoTo Rejected
    Dim model As CSectionModel, text As String
    Select Case consumer
        Case 0: Set model = ReadFormedSectionFromResults(ThisWorkbook)
        Case 1: text = TestSnapshotPlot()
        Case 2: text = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    End Select
    Exit Function
Rejected:
    If expectedError = 0 Then expectedError = vbObjectError + 4956
    TestSnapshotFileReadRejected = Err.Number = expectedError
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет короткий источник для каждого действия и расположение абзаца без открытия файла и AutoCAD.
Private Sub TestSnapshotActionMessages()
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Dim mode As String: mode = settings.GetRequiredString("Results.SnapshotStorage")
    Dim name As String: name = settings.GetRequiredString("Results.SnapshotName")
    TestSnapshotSetting "Results.SnapshotStorage", "Excel"
    TestSnapshotCheck "message.calc.Excel", SnapshotActionDataOrigin(ThisWorkbook, "Calculation") = "Геометрия для расчёта загружена с листа Results текущей книги."
    TestSnapshotSetting "Results.SnapshotStorage", "File"
    TestSnapshotSetting "Results.SnapshotName", "Контрольный_снимок"
    Dim source As String: source = SnapshotActionDataOrigin(ThisWorkbook, "Calculation")
    TestSnapshotCheck "message.calc.File", source = "Геометрия для расчёта загружена из файла «Контрольный_снимок.ndm»."
    TestSnapshotCheck "message.export.full", InStr(SnapshotActionDataOrigin(ThisWorkbook, "Export"), "Геометрия и НДС для экспорта") = 1
    TestSnapshotCheck "message.export.geometryOnly", InStr(SnapshotActionDataOrigin(ThisWorkbook, "Export", True), "НДС") = 0
    TestSnapshotCheck "message.update", InStr(SnapshotActionDataOrigin(ThisWorkbook, "Update"), "Данные для схемы и таблиц") = 1
    TestSnapshotCheck "message.afterOutcome", WorkbookMessageAfterOutcome("Готово." & vbCrLf & vbCrLf & "Строк: 4.", source) = _
        "Готово." & vbCrLf & vbCrLf & source & vbCrLf & vbCrLf & "Строк: 4."
    TestSnapshotSetting "Results.SnapshotStorage", mode
    TestSnapshotSetting "Results.SnapshotName", name
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает сохранённые числа/строки буквально, исключая только новый RunID первого столбца.
Private Function TestSnapshotSameIgnoringRun(ByVal expected As Variant, ByVal actual As Variant) As Boolean
    If Not IsArray(expected) Or Not IsArray(actual) Then Exit Function
    If UBound(expected, 1) <> UBound(actual, 1) Or UBound(expected, 2) <> UBound(actual, 2) Then Exit Function
    Dim row As Long
    For row = 2 To UBound(expected, 1): expected(row, 1) = 0#: actual(row, 1) = 0#: Next row
    TestSnapshotSameIgnoringRun = TestSnapshotSameMatrix(expected, actual)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет все координаты, площади, собственные инерции и фактические границы без допуска и округления.
Private Function TestSnapshotSameModel(ByVal expected As CSectionModel, ByVal actual As CSectionModel) As Boolean
    If expected.ConcreteCount <> actual.ConcreteCount Or expected.RebarCount <> actual.RebarCount Or expected.Contours.Count <> actual.Contours.Count Or expected.AnnotationCount <> actual.AnnotationCount Then Exit Function
    Dim i As Long
    For i = 1 To expected.ConcreteCount
        If expected.ConcreteID(i) <> actual.ConcreteID(i) Or expected.ConcreteX(i) <> actual.ConcreteX(i) Or expected.ConcreteY(i) <> actual.ConcreteY(i) Or expected.ConcreteArea(i) <> actual.ConcreteArea(i) Then Exit Function
        If expected.ConcreteLocalIx(i) <> actual.ConcreteLocalIx(i) Or expected.ConcreteLocalIy(i) <> actual.ConcreteLocalIy(i) Or expected.ConcreteLocalIxy(i) <> actual.ConcreteLocalIxy(i) Then Exit Function
        If Not TestSnapshotOptionalMatrixEqual(TestSnapshotRegionValues(expected.ConcreteGeometry(i)), TestSnapshotRegionValues(actual.ConcreteGeometry(i))) Then Exit Function
    Next i
    For i = 1 To expected.RebarCount
        If expected.RebarID(i) <> actual.RebarID(i) Or expected.RebarX(i) <> actual.RebarX(i) Or expected.RebarY(i) <> actual.RebarY(i) Or expected.RebarArea(i) <> actual.RebarArea(i) Or expected.RebarDiameter(i) <> actual.RebarDiameter(i) Then Exit Function
    Next i
    For i = 1 To expected.Contours.Count
        If expected.Contours.SegmentType(i) <> actual.Contours.SegmentType(i) Or expected.Contours.LoopID(i) <> actual.Contours.LoopID(i) Or expected.Contours.LoopRole(i) <> actual.Contours.LoopRole(i) Then Exit Function
        If expected.Contours.StartX(i) <> actual.Contours.StartX(i) Or expected.Contours.StartY(i) <> actual.Contours.StartY(i) Or expected.Contours.EndX(i) <> actual.Contours.EndX(i) Or expected.Contours.EndY(i) <> actual.Contours.EndY(i) Then Exit Function
        If expected.Contours.CenterX(i) <> actual.Contours.CenterX(i) Or expected.Contours.CenterY(i) <> actual.Contours.CenterY(i) Or expected.Contours.Radius(i) <> actual.Contours.Radius(i) Or expected.Contours.SweepAngle(i) <> actual.Contours.SweepAngle(i) Then Exit Function
    Next i
    TestSnapshotSameModel = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Преобразует каждую фактическую прямую/дугу и роль петли в матрицу точного сравнения.
Private Function TestSnapshotRegionValues(ByVal region As CConcreteRegion) As Variant
    If region Is Nothing Then Exit Function
    Dim data() As Variant, i As Long, loopID As Long
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, cx As Double, cy As Double, sweep As Double
    ReDim data(1 To region.SegmentCount + 1, 1 To 10)
    data(1, 1) = region.Area: data(1, 2) = region.LoopCount: data(1, 3) = region.SegmentCount
    For i = 1 To region.SegmentCount
        region.GetSegment i, x1, y1, x2, y2, cx, cy, sweep, loopID
        data(i + 1, 1) = x1: data(i + 1, 2) = y1: data(i + 1, 3) = x2: data(i + 1, 4) = y2
        data(i + 1, 5) = cx: data(i + 1, 6) = cy: data(i + 1, 7) = sweep: data(i + 1, 8) = loopID
        data(i + 1, 9) = region.LoopIsOpening(loopID)
    Next i
    TestSnapshotRegionValues = data
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторная проверка сохраняет защиту при изменении даты и добавлении новой части.
Private Sub TestSnapshotValidationReuse(ByVal snapshot As CCalculationSnapshot)
    snapshot.Validate: snapshot.Validate
    Dim run As Double: run = snapshot.RunID
    snapshot.RunID = run + 1#
    TestSnapshotCheck "validate.changedRunRejected", TestSnapshotValidationRejected(snapshot)
    snapshot.RunID = run: snapshot.Validate
    Dim copy As Variant: copy = snapshot.Table("rngNDMElementResults")
    copy(2, 1) = run + 2#
    snapshot.Validate
    Dim actual As Variant: actual = snapshot.Table("rngNDMElementResults")
    TestSnapshotCheck "validate.returnedArrayIndependent", CDbl(actual(2, 1)) = run
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает отказ действительного валидатора, не подменяя ожидаемые статусы.
Private Function TestSnapshotValidationRejected(ByVal snapshot As CCalculationSnapshot) As Boolean
    On Error GoTo Rejected
    snapshot.Validate
    Exit Function
Rejected:
    TestSnapshotValidationRejected = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Индекс обязан сохранять несмежные группы, регистр, порядок и произвольные позиции колонок.
Private Sub TestSnapshotElementIndex()
    Dim snapshot As CCalculationSnapshot: Set snapshot = New CCalculationSnapshot
    Dim data(1 To 7, 1 To 9) As Variant, row As Long
    data(1, 1) = "RunID": data(1, 5) = "LoadCase": data(1, 3) = "StateType"
    For row = 2 To 7
        data(row, 1) = 1#: data(row, 2) = "Element" & CStr(row)
        data(row, 5) = IIf(row = 4 Or row = 7, "B", "A")
        data(row, 3) = IIf(row = 3 Or row = 7, "CrackedState", "StrengthState")
    Next row
    snapshot.AddTable "rngNDMElementResults", data
    Dim actual As Variant: actual = snapshot.SelectElementResults("a", "strengthstate", False)
    TestSnapshotCheck "index.nonAdjacentOrder", UBound(actual, 1) = 4 And actual(2, 2) = "Element2" And actual(3, 2) = "Element5" And actual(4, 2) = "Element6"
    actual = snapshot.SelectElementResults("", "", True)
    TestSnapshotCheck "index.firstRows", UBound(actual, 1) = 5 And actual(2, 2) = "Element2" And actual(3, 2) = "Element3" And actual(4, 2) = "Element4" And actual(5, 2) = "Element7"
    actual = snapshot.SelectElementResults("missing", "", False)
    TestSnapshotCheck "index.missingHeaderOnly", UBound(actual, 1) = 1
    actual = snapshot.SelectElementResults("", "", False)
    TestSnapshotCheck "index.fullExact", TestSnapshotSameMatrix(data, actual)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет плотные колонки и серии независимо заданных чисел, Unicode и типов.
Private Sub TestSnapshotColumnCodec(ByVal original As CCalculationSnapshot, ByVal store As CCalculationSnapshotStore)
    Dim sample As CCalculationSnapshot: Set sample = New CCalculationSnapshot
    sample.RunID = original.RunID: sample.ElapsedSeconds = original.ElapsedSeconds
    Dim key As Variant
    For Each key In original.Keys(): sample.AddTable CStr(key), original.Table(CStr(key)): Next key
    Dim data(1 To 1024, 1 To 8) As Variant, row As Long, negative As TTestSnapshotDouble, bits As TTestSnapshotBits
    bits.High = &H80000000: LSet negative = bits
    For row = 1 To 1024
        data(row, 1) = "Сечение_" & ChrW$(&H3B1)
        data(row, 2) = CDbl(row) / 7#
        data(row, 3) = (row < 513)
        data(row, 4) = 0#: If row > 512 Then data(row, 4) = negative.Value
        data(row, 5) = "": If row > 512 Then data(row, 5) = Empty
        data(row, 6) = CDbl(row Mod 23) * 1E-250
        data(row, 7) = IIf(row Mod 2 = 0, "one", "two")
        If row = 1 Then data(row, 8) = "header" Else data(row, 8) = original.RunID
    Next row
    sample.AddTable "Test.ColumnCodec", data
    store.SaveFile ThisWorkbook, sample
    Dim loaded As CCalculationSnapshot: Set loaded = store.LoadFile(ThisWorkbook)
    Dim actual As Variant: actual = loaded.Table("Test.ColumnCodec")
    TestSnapshotCheck "codec.mixed.bitExact", TestSnapshotSameMatrix(data, actual)
    TestSnapshotCheck "codec.emptyKinds", VarType(actual(1, 5)) = vbString And VarType(actual(1024, 5)) = vbEmpty
    TestSnapshotCheck "codec.booleanKind", VarType(actual(1, 3)) = vbBoolean And VarType(actual(1024, 3)) = vbBoolean
    store.SaveFile ThisWorkbook, original
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторяющиеся ElementID сохраняются точно; любое нарушение шаблона включает общий кодек.
Private Sub TestSnapshotRepeatedElementIDs(ByVal original As CCalculationSnapshot, ByVal store As CCalculationSnapshotStore)
    Dim count As Long: count = original.TableRowCount("rngNDMSectionGeometry") - 1
    Dim groups As Long: groups = 2 + 4096 \ count
    Dim expected As Variant: expected = original.Table("rngNDMElementResults")
    Dim data() As Variant, row As Long, column As Long, sourceRow As Long
    ReDim data(1 To groups * count + 1, 1 To 9)
    For column = 1 To 9: data(1, column) = expected(1, column): Next column
    For row = 2 To UBound(data, 1)
        sourceRow = 2 + (row - 2) Mod count
        For column = 1 To 9: data(row, column) = expected(sourceRow, column): Next column
    Next row
    Dim testCase As Long, sample As CCalculationSnapshot, loaded As CCalculationSnapshot, key As Variant
    Dim changed As Long, savedID As String
    For testCase = 0 To 3
        changed = 0
        Select Case testCase
            Case 1: changed = UBound(data, 1): savedID = CStr(data(changed, 5)): data(changed, 5) = "TestMismatch_" & savedID
            Case 2: changed = 3: savedID = CStr(data(changed, 5)): data(changed, 5) = data(2, 5)
            Case 3: changed = 2: savedID = CStr(data(changed, 5)): data(changed, 5) = "ElementID"
        End Select
        Set sample = New CCalculationSnapshot: sample.RunID = original.RunID: sample.ElapsedSeconds = original.ElapsedSeconds
        For Each key In original.Keys()
            If CStr(key) <> "rngNDMElementResults" Then sample.AddTable CStr(key), original.Table(CStr(key))
        Next key
        sample.AddTable "rngNDMElementResults", data
        store.SaveFile ThisWorkbook, sample: Set loaded = store.LoadFile(ThisWorkbook)
        TestSnapshotCheck "codec.repeatedIDs." & CStr(testCase) & ".exact", TestSnapshotSameMatrix(data, loaded.Table("rngNDMElementResults"))
        TestSnapshotCheck "codec.repeatedIDs." & CStr(testCase) & ".sourceUnchanged", TestSnapshotSameMatrix(data, sample.Table("rngNDMElementResults"))
        If changed > 0 Then data(changed, 5) = savedID
    Next testCase
    store.SaveFile ThisWorkbook, original
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторный расчёт сохранённого импорта не должен умножать автоматические подписи.
Private Sub TestSnapshotRepeatedImportedAnnotations(ByVal extension As String)
    modTestPerformance.TestConfigurePerformanceFixture "IMPORTED_CRACK_SP35", 3, "Newton", "Saved", "Normal", extension
    TestSnapshotSetting "Results.SnapshotStorage", "Excel"
    TestSnapshotSetting "General.ExecutionReportEnabled", "No"
    modTestPlotConfig.TestSetOutputCombination ThisWorkbook, "PERF_01"
    Dim message As String: message = RunSectionCalculationForWorkbook(ThisWorkbook)
    Dim first As Variant: first = ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionAnnotations")
    Dim expected As Variant: expected = ReadPhysicalResultTable(ThisWorkbook, "rngNDMElementResults")
    message = RunSectionCalculationForWorkbook(ThisWorkbook)
    Dim actual As Variant: actual = ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionAnnotations")
    TestSnapshotCheck "repeat.imported.annotationCount", UBound(first, 1) = UBound(actual, 1)
    Dim i As Long
    For i = 2 To UBound(first, 1): first(i, 1) = actual(i, 1): Next i
    TestSnapshotCheck "repeat.imported.annotationsExact", TestSnapshotSameMatrix(first, actual)
    actual = ReadElementResultTable(ThisWorkbook)
    For i = 2 To UBound(expected, 1): expected(i, 1) = actual(i, 1): Next i
    TestSnapshotCheck "repeat.imported.statesExact", TestSnapshotSameMatrix(expected, actual)
    TestSnapshotCheck "repeat.imported.contextReleased", TestCalculationSnapshotReadReleased()
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет текущие читаемые Excel-таблицы, отсутствие архива и неизменность верхнего вывода.
' Внешний runner вызывает метод только в своей частной рассчитанной копии.
Public Function TestExcelCurrentResultsRefresh() As String
    mTestReport = "": mTestPassed = 0: mTestFailed = 0
    On Error GoTo Failed
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    TestSnapshotCheck "excel.storage", store.StorageMode(ThisWorkbook) = "Excel"
    TestSnapshotCheck "excel.noArchiveName", Not TestSnapshotNameExists("rngNDMLoadCaseResults")
    TestSnapshotCheck "excel.noFileReference", Not store.HasFileReference(ThisWorkbook)
    Dim snapshot As CCalculationSnapshot: Set snapshot = CurrentCalculationSnapshot(ThisWorkbook)
    TestSnapshotCheck "excel.noPersistentFullSnapshot", snapshot Is Nothing
    Dim properties As Variant: properties = ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionProperties")
    TestSnapshotCheck "excel.readableLoadCaseProperties", Not TestSnapshotCommonOnly(properties)
    Dim names As Variant: names = Array("rngBatchSummary", "rngStrengthSummaryAnchor", "rngCrackSummaryAnchor", "rngStabilitySummaryAnchor")
    Dim widths As Variant: widths = Array(33, 49, 72, 84)
    Dim before(0 To 3) As Variant, i As Long
    Dim plot As String: plot = TestSnapshotPlot()
    Dim exported As String: exported = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    TestSnapshotCorruptUpper ' Удалённые значения не должны восстанавливаться из скрытого архива.
    For i = 0 To 3: before(i) = ThisWorkbook.Names.Item(CStr(names(i))).RefersToRange.Resize(32, CLng(widths(i))).Value2: Next i
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Dim iterations As String: iterations = settings.GetRequiredString("Solver.MaxIterations")
    Dim selectedName As String: selectedName = settings.GetRequiredString("Results.SnapshotName")
    TestSnapshotSetting "Results.SnapshotName", "../this_file_is_not_used_in_Excel"
    TestSnapshotSetting "Solver.MaxIterations", "not_a_number"
    Dim solves As Long: solves = SectionEquilibriumSolveCount()
    UpdateSectionResultsForWorkbook ThisWorkbook
    TestSnapshotCheck "excel.refresh.noSolver", SectionEquilibriumSolveCount() = solves
    TestSnapshotCheck "excel.refresh.plotExact", TestSnapshotPlot() = plot
    TestSnapshotCheck "excel.refresh.exportExact", Audit02ReadExportSnapshotForTests(ThisWorkbook) = exported
    For i = 0 To 3
        TestSnapshotCheck "excel.refresh.keepsUpper." & CStr(names(i)), TestSnapshotSameMatrix(before(i), ThisWorkbook.Names.Item(CStr(names(i))).RefersToRange.Resize(32, CLng(widths(i))).Value2)
    Next i
    TestSnapshotCheck "excel.refresh.keepsProperties", TestSnapshotSameMatrix(properties, ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionProperties"))
    TestSnapshotCheck "excel.refresh.contextReleased", TestCalculationSnapshotReadReleased()
    TestSnapshotSetting "Solver.MaxIterations", iterations
    TestSnapshotSetting "Results.SnapshotName", selectedName
    GoTo Finished
Failed:
    TestSnapshotCheck "excel.runtime " & CStr(Err.Number) & ": " & Err.Description, False
Finished:
    TestExcelCurrentResultsRefresh = mTestReport & "TOTAL EXCEL CURRENT RESULTS: passed=" & CStr(mTestPassed) & "; failed=" & CStr(mTestFailed)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет все шесть технических таблиц Excel с теми же готовыми массивами расчёта.
Private Sub TestSnapshotExcelCurrent(ByVal original As CCalculationSnapshot)
    Dim key As Variant
    TestSnapshotUpper original
    For Each key In original.BaseKeys()
        TestSnapshotCheck "excel.physical." & CStr(key), TestSnapshotSameMatrix(original.ResultTable(CStr(key)), ReadPhysicalResultTable(ThisWorkbook, CStr(key)))
    Next key
    Dim previousReport As String: previousReport = mTestReport
    Dim previousPassed As Long: previousPassed = mTestPassed
    Dim previousFailed As Long: previousFailed = mTestFailed
    Dim result As String: result = TestExcelCurrentResultsRefresh()
    mTestReport = previousReport & mTestReport: mTestPassed = previousPassed + mTestPassed: mTestFailed = previousFailed + mTestFailed
    If InStr(result, "failed=0") = 0 Then Err.Raise 5, , result
    SaveAndDisplayCalculationSnapshot ThisWorkbook, original
    TestSnapshotUpper original
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет отсутствие удалённого имени без создания диапазона или изменения книги.
Private Function TestSnapshotNameExists(ByVal name As String) As Boolean
    Dim target As Object
    On Error Resume Next
    Set target = ThisWorkbook.Names.Item(name)
    On Error GoTo 0
    TestSnapshotNameExists = Not target Is Nothing
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Готовит реальный расчёт для внешней проверки Save/Close/Open и Save As.
' Runner вызывает его только в своей записываемой копии; пользовательские книги не участвуют.
Public Function TestPrepareSnapshotReopen(ByVal mode As String) As String
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    modTestPerformance.TestConfigurePerformanceFixture "MIXED", 6, "Newton", "Circle", "Normal", settings.GetRequiredString("General.DiagramExtension")
    TestSnapshotSetting "Results.SnapshotStorage", mode
    TestSnapshotSetting "Results.SnapshotName", "TestReopen_" & ChrW$(&H3B1) & ".ndm"
    TestSnapshotSetting "General.ExecutionReportEnabled", "No"
    modTestPlotConfig.TestSetOutputCombination ThisWorkbook, "PERF_01"
    modTestPlotConfig.TestSetOutputCombination ThisWorkbook, "PERF_01"
    Dim message As String: message = FormSectionForWorkbook(ThisWorkbook)
    message = RunSectionCalculationForWorkbook(ThisWorkbook)
    TestPrepareSnapshotReopen = TestSnapshotReopenFingerprint()
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' File восстанавливает повреждённый вывод; Excel обновляет схему без изменения сводок.
Public Function TestRestoreSnapshotAfterReopen() As String
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    If store.StorageMode(ThisWorkbook) = "File" Then TestSnapshotCorruptUpper
    TestSnapshotSetting "Solver.MaxIterations", "not_a_number"
    Dim solves As Long: solves = SectionEquilibriumSolveCount()
    UpdateSectionResultsForWorkbook ThisWorkbook
    If SectionEquilibriumSolveCount() <> solves Then Err.Raise 5, , "Восстановление повторно вызвало solver."
    Dim snapshot As CCalculationSnapshot: Set snapshot = CurrentCalculationSnapshot(ThisWorkbook)
    mTestReport = "": mTestPassed = 0: mTestFailed = 0
    If Not snapshot Is Nothing Then
        If Not snapshot.IsGeometryPreview Then TestSnapshotUpper snapshot
    End If
    If mTestFailed <> 0 Then Err.Raise 5, , mTestReport
    If Not TestCalculationSnapshotReadReleased() Then Err.Raise 5, , "Контекст чтения не освобождён."
    TestRestoreSnapshotAfterReopen = TestSnapshotReopenFingerprint()
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает все части снимка, данные схемы и экспорта для побитного сравнения после открытия.
Public Function TestSnapshotReopenFingerprint() As String
    Dim snapshot As CCalculationSnapshot: Set snapshot = CurrentCalculationSnapshot(ThisWorkbook)
    Dim keys As Variant
    If snapshot Is Nothing Then
        keys = Array("rngNDMSectionGeometry", "rngNDMSectionProperties", "rngNDMMaterialDiagrams", "rngNDMElementResults", "rngNDMSectionContours", "rngNDMSectionAnnotations", "rngNDMElementGeometry")
    Else
        keys = snapshot.Keys()
    End If
    Dim key As Variant, values As Variant, row As Long, column As Long, text As String, value As Variant
    Dim lines() As String, line As String
    If Not snapshot Is Nothing Then text = TestSnapshotDouble(snapshot.RunID) & "|" & TestSnapshotDouble(snapshot.ElapsedSeconds)
    For Each key In keys
        If snapshot Is Nothing Then values = ReadPhysicalResultTable(ThisWorkbook, CStr(key)) Else values = snapshot.Table(CStr(key))
        ReDim lines(0 To UBound(values, 1))
        lines(0) = CStr(key) & "|" & CStr(UBound(values, 1)) & "|" & CStr(UBound(values, 2))
        For row = 1 To UBound(values, 1)
            line = vbNullString
            For column = 1 To UBound(values, 2)
                value = values(row, column)
                If VarType(value) = vbString Or IsEmpty(value) Then
                    line = line & "S" & CStr(Len(CStr(value))) & ":" & CStr(value) & "|"
                Else
                    line = line & "D" & TestSnapshotDouble(CDbl(value)) & "|"
                End If
            Next column
            lines(row) = line
        Next row
        text = text & vbLf & Join(lines, vbLf)
    Next key
    TestSnapshotReopenFingerprint = text & vbLf & TestSnapshotPlot() & vbLf & Audit02ReadExportSnapshotForTests(ThisWorkbook)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет все части одного логического снимка, включая каждое число IEEE 754.
Private Sub TestSnapshotEqual(ByVal label As String, ByVal expected As CCalculationSnapshot, ByVal actual As CCalculationSnapshot)
    Dim key As Variant
    TestSnapshotCheck label & ".run", TestSnapshotDouble(expected.RunID) = TestSnapshotDouble(actual.RunID)
    TestSnapshotCheck label & ".elapsed", TestSnapshotDouble(expected.ElapsedSeconds) = TestSnapshotDouble(actual.ElapsedSeconds)
    For Each key In expected.Keys()
        TestSnapshotCheck label & "." & CStr(key), actual.HasTable(CStr(key))
        If actual.HasTable(CStr(key)) Then TestSnapshotCheck label & ".values." & CStr(key), TestSnapshotSameMatrix(expected.Table(CStr(key)), actual.Table(CStr(key)))
    Next key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет физические верхние таблицы с сохранёнными данными, не решая НДС.
Private Sub TestSnapshotUpper(ByVal snapshot As CCalculationSnapshot)
    Dim key As Variant, name As String, values As Variant
    For Each key In Array("Summary", "Strength.Data", "Cracks.Data", "Stability.Data")
        Select Case CStr(key)
            Case "Summary": name = "rngBatchSummary"
            Case "Strength.Data": name = "rngStrengthSummaryAnchor"
            Case "Cracks.Data": name = "rngCrackSummaryAnchor"
            Case "Stability.Data": name = "rngStabilitySummaryAnchor"
        End Select
        values = snapshot.Table(CStr(key))
        TestSnapshotCheck "upper." & CStr(key), TestSnapshotSameMatrix(values, ThisWorkbook.Names.Item(name).RefersToRange.Resize(UBound(values, 1), UBound(values, 2)).Value2)
    Next key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Удаляет значения, чтобы обновление не могло пройти за счёт оставшихся физических значений.
Private Sub TestSnapshotCorruptUpper()
    Dim name As Variant
    For Each name In Array("rngStrengthSummaryAnchor", "rngCrackSummaryAnchor", "rngStabilitySummaryAnchor")
        ThisWorkbook.Names.Item(CStr(name)).RefersToRange.Resize(30, 3).ClearContents
    Next name
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Числа сравнивает побитно, Empty и пустую строку считает эквивалентным пустым Excel-полем.
Private Function TestSnapshotSameMatrix(ByVal expected As Variant, ByVal actual As Variant) As Boolean
    On Error GoTo Failed
    If UBound(expected, 1) <> UBound(actual, 1) Or UBound(expected, 2) <> UBound(actual, 2) Then Exit Function
    Dim i As Long, j As Long, a As Variant, b As Variant
    For i = 1 To UBound(expected, 1)
        For j = 1 To UBound(expected, 2)
            a = expected(i, j): b = actual(i, j)
            If VarType(a) = vbString Or VarType(b) = vbString Or IsEmpty(a) Or IsEmpty(b) Then
                If CStr(a) <> CStr(b) Then GoTo Different
            ElseIf TestSnapshotDouble(CDbl(a)) <> TestSnapshotDouble(CDbl(b)) Then
                GoTo Different
            End If
        Next j
    Next i
    TestSnapshotSameMatrix = True
    Exit Function
Different:
    mTestReport = mTestReport & "DIFF|row=" & CStr(i) & "|column=" & CStr(j) & "|expected=" & CStr(a) & "|actual=" & CStr(b) & vbCrLf
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Формирует точное представление Double без зависимости от локали и десятичных цифр.
Private Function TestSnapshotDouble(ByVal value As Double) As String
    Dim number As TTestSnapshotDouble, bits As TTestSnapshotBits
    number.Value = value: LSet bits = number
    TestSnapshotDouble = Right$("00000000" & Hex$(bits.High), 8) & Right$("00000000" & Hex$(bits.Low), 8)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает реальные данные потребителя схемы, а не только наличие ChartObject.
Private Function TestSnapshotPlot() As String
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Dim reader As CSectionPlotDataReader: Set reader = New CSectionPlotDataReader: reader.LoadFromWorkbook ThisWorkbook, settings
    Dim text As String, i As Long
    text = reader.LoadCase & "|" & reader.StateType & "|" & reader.LengthUnit & "|" & reader.ResultUnit
    For i = 1 To reader.Count
        text = text & vbLf & reader.ElementID(i) & "|" & TestSnapshotDouble(reader.X(i)) & "|" & TestSnapshotDouble(reader.Y(i)) & "|" & _
            TestSnapshotDouble(reader.Area(i)) & "|" & TestSnapshotDouble(reader.ResultValue(i)) & "|" & reader.PhysicalState(i)
    Next i
    TestSnapshotPlot = text
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что LC-зависимые поля не остались в общей геометрической таблице.
Private Function TestSnapshotCommonOnly(ByVal values As Variant) As Boolean
    Dim i As Long
    For i = 2 To UBound(values, 1)
        If CStr(values(i, 2)) <> "ALL" Then Exit Function
        If CStr(values(i, 3)) = "LoadReferenceX" Or CStr(values(i, 3)) = "LoadReferenceY" Or InStr(1, CStr(values(i, 3)), "SP35.Slenderness", vbBinaryCompare) = 1 Then Exit Function
    Next i
    TestSnapshotCommonOnly = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Неудачная загрузка обязана завершиться явной ошибкой до перезаписи Results.
Private Function TestSnapshotRefreshRejected() As Boolean
    On Error GoTo Rejected
    UpdateSectionResultsForWorkbook ThisWorkbook
    Exit Function
Rejected:
    TestSnapshotRefreshRejected = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет имя через тот же рабочий валидатор, что сохранение файла.
Private Function TestSnapshotPathRejected(ByVal store As CCalculationSnapshotStore) As Boolean
    On Error GoTo Rejected
    Dim path As String: path = store.SnapshotPath(ThisWorkbook)
    Exit Function
Rejected:
    TestSnapshotPathRejected = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Устанавливает настройку в частном fixture по ключу и реальному именованному диапазону.
Private Sub TestSnapshotSetting(ByVal key As String, ByVal value As String)
    Dim table As Object, field As String, column As Long: field = key: column = 2
    If Left$(key, 6) = "Units." Then
        Set table = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
        field = Split(key, ".")(1): column = 4
    Else
        Set table = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    End If
    Dim values As Variant: values = table.Value2
    Dim i As Long
    For i = 1 To UBound(values, 1)
        If CStr(values(i, 1)) = field Then table.Cells(i, column).Value2 = value: Exit Sub
    Next i
    Err.Raise 5, , "Не найдена тестовая настройка: " & key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет неизменённые исходные числа, включая малые моменты и ошибочную строку.
Private Sub TestSnapshotOriginalLoadMemory()
    Dim batch As CBatchSectionCalculator: Set batch = New CBatchSectionCalculator
    batch.AddCombination "RAW", -123.456#, 1E-20, -2E-20, "PR1", "Исходные числа"
    batch.AddInvalidCombination "BAD", "PR1", "Ошибка", "Некорректная нагрузка"
    TestSnapshotCheck "loads.original.available", batch.OriginalLoadsAvailable(1)
    TestSnapshotCheck "loads.original.force", TestSnapshotDouble(batch.OriginalN(1)) = TestSnapshotDouble(-123.456#)
    TestSnapshotCheck "loads.original.tinyMx", TestSnapshotDouble(batch.OriginalMx(1)) = TestSnapshotDouble(1E-20)
    TestSnapshotCheck "loads.original.tinyMy", TestSnapshotDouble(batch.OriginalMy(1)) = TestSnapshotDouble(-2E-20)
    TestSnapshotCheck "loads.invalid.notInvented", Not batch.OriginalLoadsAvailable(2)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает новые исходные числа с фактическим вводом, а рабочие — с таблицей устойчивости.
' Строка worst обязана принадлежать тому же LC, что указан в самом верху снимка.
Private Sub TestSnapshotSummaryLoads(ByVal snapshot As CCalculationSnapshot)
    Dim data As Variant: data = snapshot.Table("Summary")
    TestSnapshotCheck "loads.summary.width", UBound(data, 2) = 33
    TestSnapshotCheck "loads.summary.title", CStr(data(6, 28)) = "Нагрузки в точке приложения"
    TestSnapshotCheck "loads.summary.lineBreak", InStr(CStr(data(8, 31)), vbLf & "(если применимо)") > 0
    TestSnapshotCheck "loads.summary.originalMxComment", CStr(data(11, 29)) = "Исходный изгибающий момент Mx"
    TestSnapshotCheck "loads.summary.originalMyComment", CStr(data(11, 30)) = "Исходный изгибающий момент My"
    TestSnapshotCheck "loads.summary.forceNoEta", InStr(CStr(data(11, 31)), ChrW$(&H3B7)) = 0
    TestSnapshotCheck "loads.summary.momentsEta", InStr(CStr(data(11, 32)), ChrW$(&H3B7)) > 0
    TestSnapshotCheck "loads.summary.workingMxComment", CStr(data(11, 32)) = "Изгибающий момент Mx с учётом " & ChrW$(&H3B7) & " от продольного изгиба (если применимо)"
    TestSnapshotCheck "loads.summary.workingMyComment", CStr(data(11, 33)) = "Изгибающий момент My с учётом " & ChrW$(&H3B7) & " от продольного изгиба (если применимо)"
    Dim anchor As Object: Set anchor = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Cells(1, 1)
    TestSnapshotCheck "loads.summary.workingMxSeparate", Not CBool(anchor.Cells(11, 32).MergeCells)
    TestSnapshotCheck "loads.summary.workingMySeparate", Not CBool(anchor.Cells(11, 33).MergeCells)
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Dim units As CUnitSystem: Set units = New CUnitSystem: units.LoadFromSettings settings
    TestSnapshotCheck "loads.summary.forceUnit", CStr(data(10, 28)) = "N, " & units.OutputForceUnit
    TestSnapshotCheck "loads.summary.momentUnit", CStr(data(10, 29)) = "Mx, " & units.OutputMomentUnit
    Dim loads As Variant: loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange.Value2
    Dim stable As Variant: stable = snapshot.Table("Stability.Data")
    Dim row As Long, inputRow As Long, stableRow As Long, column As Long, expected As Double
    For row = 13 To UBound(data, 1)
        If Len(CStr(data(row, 1))) = 0 Then GoTo NextSummaryLoad
        For inputRow = 2 To UBound(loads, 1)
            If CStr(loads(inputRow, 1)) = CStr(data(row, 1)) Then Exit For
        Next inputRow
        If inputRow > UBound(loads, 1) Then Err.Raise 5, , "Ввод тестового сочетания не найден."
        expected = units.InternalForceToOutput(units.InputForceToInternal(CDbl(loads(inputRow, 2))))
        TestSnapshotCheck "loads.summary.originalN." & CStr(row), TestSnapshotDouble(data(row, 28)) = TestSnapshotDouble(expected)
        expected = units.InternalMomentMxToOutput(units.InputMomentMxToInternal(CDbl(loads(inputRow, 3))))
        TestSnapshotCheck "loads.summary.originalMx." & CStr(row), TestSnapshotDouble(data(row, 29)) = TestSnapshotDouble(expected)
        expected = units.InternalMomentMyToOutput(units.InputMomentMyToInternal(CDbl(loads(inputRow, 4))))
        TestSnapshotCheck "loads.summary.originalMy." & CStr(row), TestSnapshotDouble(data(row, 30)) = TestSnapshotDouble(expected)
        For stableRow = 1 To UBound(stable, 1)
            If CStr(stable(stableRow, 1)) = CStr(data(row, 1)) And CStr(stable(stableRow, 3)) = "OK" Then
                For column = 0 To 2
                    TestSnapshotCheck "loads.summary.afterStability." & CStr(row) & "." & CStr(column), TestSnapshotDouble(data(row, 31 + column)) = TestSnapshotDouble(stable(stableRow, 12 + column))
                Next column
            End If
        Next stableRow
        If CStr(data(row, 1)) = CStr(data(1, 1)) Then
            For column = 28 To 33
                TestSnapshotCheck "loads.summary.worst." & CStr(column), TestSnapshotDouble(data(12, column)) = TestSnapshotDouble(data(row, column))
            Next column
        End If
NextSummaryLoad:
    Next row
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает совместимую 26-колоночную сводку без подстановки текущих нагрузок Config.
Private Sub TestSnapshotLegacySummary(ByVal original As CCalculationSnapshot)
    Dim sample As CCalculationSnapshot: Set sample = New CCalculationSnapshot
    sample.RunID = original.RunID: sample.ElapsedSeconds = original.ElapsedSeconds
    Dim key As Variant, saved As Variant, data() As Variant, row As Long, column As Long
    For Each key In original.Keys()
        saved = original.Table(CStr(key))
        If CStr(key) = "Summary" Then
            ReDim data(1 To UBound(saved, 1), 1 To 26)
            For row = 1 To UBound(saved, 1)
                For column = 1 To 26: data(row, column) = saved(row, column): Next column
            Next row
            sample.AddTable CStr(key), data
        Else
            sample.AddTable CStr(key), saved
        End If
    Next key
    sample.Validate
    Dim writer As CBatchResultWriter: Set writer = New CBatchResultWriter
    writer.RestoreSummary ThisWorkbook, sample
    Dim target As Object: Set target = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Cells(1, 1)
    TestSnapshotCheck "loads.legacy.explicitMissing", CStr(target.Cells(11, 28).Value2) = "Не сохранено в этом снимке"
    TestSnapshotCheck "loads.legacy.noCurrentInput", Application.CountA(target.Offset(11, 27).Resize(31, 6)) = 0
    saved = original.Table("Summary")
    TestSnapshotCheck "loads.legacy.keepsWorst", CStr(target.Cells(1, 1).Value2) = CStr(saved(1, 1))
    writer.RestoreSummary ThisWorkbook, original
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет заданный горизонтальный порядок, не фиксируя адреса верхних таблиц.
Private Function TestSnapshotLowerOrder() As Boolean
    Dim name As Variant, last As Long, column As Long
    For Each name In Array("rngNDMSectionGeometry", "rngNDMSectionProperties", "rngNDMMaterialDiagrams", "rngNDMSectionContours", "rngNDMElementResults", "rngNDMElementGeometry", "rngNDMSectionAnnotations")
        column = ThisWorkbook.Names.Item(CStr(name)).RefersToRange.Column
        If column <= last Then Exit Function
        last = column
    Next name
    TestSnapshotLowerOrder = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет компактный итог; успешные числа отдельных ячеек не раздувают журнал.
Private Sub TestSnapshotCheck(ByVal label As String, ByVal value As Boolean)
    If value Then mTestPassed = mTestPassed + 1 Else mTestFailed = mTestFailed + 1
    mTestReport = mTestReport & IIf(value, "PASS ", "FAIL ") & label & vbCrLf
End Sub
