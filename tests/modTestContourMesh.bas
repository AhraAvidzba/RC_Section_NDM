Attribute VB_Name = "modTestContourMesh"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: общая сетка параметрических и импортированных контуров
' ==========================================================================
' Проверяет построение сетки и решение НДС рабочим ядром и диаграммами книги.
' Сравнивает каждый элемент. Config и Results изменяет только в отдельной тестовой копии.

#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#End If

Private mPassed As Long, mFailed As Long, mReport As String
Private mFixture As Object, mConfig As Object

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает сетки и НДС генерации и имитации импорта для шести форм; отдельно проверяет отверстия, ошибки ввода и сохранение модели в книге.
Public Function RunContourMeshTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    Set mFixture = Application.Workbooks.Add(-4167)
    Set mConfig = mFixture.Worksheets(1): mConfig.Name = "MeshConfig"
    Dim shape As Long, subdivision As Variant, geometry As CGeometryRegion
    For shape = 1 To 6
        Set geometry = TestControlGeometry(shape)
        For Each subdivision In Array(1, 2, 3)
            TestCompareGeneratedAndImported geometry, CStr(shape) & ".sub" & CStr(subdivision), CLng(subdivision), CLng(subdivision) = 2
        Next subdivision
    Next shape
    TestCheckArbitraryDomains
    TestCheckImportContracts
    TestCheckSnapshot
    GoTo Finish
Failed:
    TestCheck "mesh.runtime." & CStr(Err.Number) & "; " & Err.Source & "; " & Err.Description, False
Finish:
    On Error Resume Next
    If Not mFixture Is Nothing Then mFixture.Close False
    Set mFixture = Nothing: Set mConfig = Nothing
    On Error GoTo 0
    RunContourMeshTests = mReport & "TOTAL_CONTOUR_MESH: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' приложение AutoCAD и новый документ создаёт внешний сценарий проверки.
' Обычный экспорт строит контуры; обычный импорт читает созданные объекты чертежа.
Public Function RunNativeContourMeshTests(ByVal acad As Object, ByVal template As String) As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    Dim doc As Object, geometry As CGeometryRegion, settings As CSystemSettingsReader, units As CUnitSystem
    Dim mesh As CFiberMeshBuilder, bars As CRebarLayout, generated As CSectionModel, imported As CSectionModel, restored As CSectionModel
    Dim writer As CNDMResultsWriter, importer As CAutoCADSectionModelImporter, shape As Long, format As Variant
    Dim outerLayer As String, openingLayer As String, rebarLayer As String, meshLayer As String, key As String
    Dim i As Long, exported As Long, countBefore As Long, circleEntity As Object, regions As Variant, curves(0 To 0) As Object
    Dim center(0 To 2) As Double, region As Object, path As String
    On Error GoTo Failed
    TestCheck "native.actualAutoCAD", InStr(1, CStr(acad.FullName), "\AutoCAD 2023\acad.exe", vbTextCompare) > 0
    If InStr(1, CStr(acad.FullName), "\AutoCAD 2023\acad.exe", vbTextCompare) = 0 Then Err.Raise vbObjectError + 5295, , "Требуется Autodesk AutoCAD 2023."
    Set mFixture = Application.Workbooks.Add(-4167): Set mConfig = mFixture.Worksheets(1)
    Set doc = acad.Documents.Add(template)
    Set writer = New CNDMResultsWriter: Set importer = New CAutoCADSectionModelImporter
    For shape = 1 To 6
        For Each format In Array("Region", "Polyline")
            key = "CM_" & CStr(shape) & "_" & CStr(format)
            outerLayer = key & "_OUT": openingLayer = key & "_HOLE": rebarLayer = key & "_BAR": meshLayer = key & "_EMPTY"
            Set geometry = TestControlGeometry(shape)
            Set settings = TestMeshSettings(2, "m", 0.04, 0.03): Set units = TestSettingsUnits(settings)
            Set mesh = BuildConfiguredConcreteMesh(geometry, settings, units): Set bars = TestControlBars(mesh, geometry)
            Set generated = TestBuildGeneratedSectionModel(mesh, bars): geometry.BuildContours generated.Contours
            writer.WriteGeometryPreview ThisWorkbook, generated, PrepareSectionSnapshot(generated), units
            exported = SP35ExportSavedMaterialContoursForTests(doc, outerLayer, openingLayer, CStr(format))
            TestCheck key & ".exportedContours", exported > 0
            Audit03EnsureAutoCADLayerForTests doc, rebarLayer, 1
            Audit03EnsureAutoCADLayerForTests doc, meshLayer, 8
            For i = 1 To bars.Count
                center(0) = bars.X(i): center(1) = bars.Y(i)
                Set circleEntity = doc.ModelSpace.AddCircle(center, bars.Diameter(i) / 2#)
                Set curves(0) = circleEntity: regions = doc.ModelSpace.AddRegion(curves)
                Set region = regions(LBound(regions)): region.Layer = rebarLayer: circleEntity.Delete
            Next i
            countBefore = doc.ModelSpace.Count
            Set imported = importer.ImportFromModelSpace(doc.ModelSpace, meshLayer, rebarLayer, "Rebar", 0#, units, outerLayer, openingLayer, doc.Layers, settings)
            TestCheck key & ".sourcePreserved", doc.ModelSpace.Count = countBefore
            TestCheck key & ".source", imported.ConcreteMeshSource = "AutoCADContours"
            TestCheck key & ".notice", InStr(importer.ResultMessage(imported), CONTOUR_MESH_GENERATION_NOTICE) > 0
            If shape < 5 Then TestCompareModels key, generated, imported, False Else TestCompareCurvedModels key, generated, imported
            If shape = 3 Then TestCompareSolutions key, generated, imported, True
            If shape = 3 And CStr(format) = "Region" Then
                TestCheckBatchPair generated, imported
                TestCheckNativeFullExport doc, imported
            End If
            writer.WriteGeometryPreview ThisWorkbook, imported, PrepareSectionSnapshot(imported), units
            Set restored = ReadSectionGeometryFromResults(ThisWorkbook)
            TestCompareModels key & ".snapshot", imported, restored, False
            TestCheck key & ".restoredMeshSource", restored.ConcreteMeshSource = "AutoCADContours"
            exported = SP35ExportSavedMaterialContoursForTests(doc, key & "_OUT2", key & "_HOLE2", CStr(format))
            TestCheck key & ".secondExport", exported > 0
            Dim reimported As CSectionModel
            Set reimported = importer.ImportFromModelSpace(doc.ModelSpace, meshLayer, rebarLayer, "Rebar", 0#, units, key & "_OUT2", key & "_HOLE2", doc.Layers, settings)
            TestCompareModels key & ".CADRoundTrip", imported, reimported, False
            mReport = mReport & "NATIVE " & key & ": concrete=" & CStr(imported.ConcreteCount) & "; contours=" & CStr(imported.Contours.Count) & "; rebar=" & CStr(imported.RebarCount) & vbCrLf
        Next format
    Next shape
    TestCheckNativeTwoHoles doc, settings, units
    path = ThisWorkbook.Path & "\ContourMeshNative.dwg": doc.SaveAs path
    mReport = mReport & "NATIVE_DWG: " & path & vbCrLf
    GoTo Finish
Failed:
    TestCheck "native.mesh.runtime." & CStr(Err.Number) & "; " & Err.Source & "; " & Err.Description, False
Finish:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    If Not mFixture Is Nothing Then mFixture.Close False
    Set mFixture = Nothing: Set mConfig = Nothing
    On Error GoTo 0
    RunNativeContourMeshTests = mReport & "TOTAL_NATIVE_CONTOUR_MESH: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет настоящий экспорт волокон и напряжений из полного Results:
' миллиметры, количество/площадь Region и отсутствие нового решения НДС.
Private Sub TestCheckNativeFullExport(ByVal doc As Object, ByVal model As CSectionModel)
    Dim table As Object, profiles As Object, saved As Variant, savedProfiles As Variant
    Dim i As Long, settings As CSystemSettingsReader, concreteLayer As String, rebarLayer As String
    Dim entity As Object, count As Long, bars As Long, area As Double, expectedArea As Double, exported As Long, solves As Long
    Dim exportCell As Object, snapshot As String, failure As Long, reason As String, objectsBefore As Long
    On Error GoTo Failed
    Set table = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange: saved = table.Formula
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange: savedProfiles = profiles.Formula
    For i = 2 To table.Rows.Count
        If CStr(table.Cells(i, 1).Value2) = "Output.CombinationID" Then Set exportCell = table.Cells(i, 2): exportCell.Value2 = "CM_BEND"
    Next i
    modTestPlotConfig.TestSetOutputCombinationInTable table, "CM_BEND"
    For i = 2 To profiles.Rows.Count
        If CStr(profiles.Cells(i, 2).Value2) = "Visualization.State" Then profiles.Cells(i, 3).Value2 = "StrengthState"
    Next i
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    concreteLayer = settings.GetRequiredString("AutoCAD.Common.ConcreteLayer")
    rebarLayer = settings.GetRequiredString("AutoCAD.Common.RebarLayer")
    solves = SectionEquilibriumSolveCount()
    ' ДЛЯ ТЕСТОВ: CM_BEND остаётся нагруженным контрольным сочетанием. В режиме
    ' без продолжения диаграмм его НДС прочности отсутствует из-за несходимости.
    ' Экспорт должен отказать без нового расчёта и создания объектов.
    objectsBefore = doc.ModelSpace.Count
    If Not settings.GetBoolean("General.DiagramExtension", False) Then
        On Error Resume Next
        exported = Audit03ExportAutoCADDocumentForTests(doc)
        failure = Err.Number: reason = Err.Description: Err.Clear
        On Error GoTo Failed
        TestCheck "native.fullExport.missingState.rejected", failure <> 0 And InStr(1, reason, "StrengthState", vbBinaryCompare) > 0 And InStr(1, reason, "CM_BEND", vbBinaryCompare) > 0
        TestCheck "native.fullExport.missingState.noObjects", doc.ModelSpace.Count = objectsBefore
    Else
        snapshot = Audit02ReadExportSnapshotForTests(ThisWorkbook)
        TestCheck "native.fullExport.extendedState.available", InStr(1, snapshot, "CM_BEND|PR1|StrengthState|", vbBinaryCompare) = 1
    End If
    TestCheck "native.fullExport.stateSelection.noSolve", solves = SectionEquilibriumSolveCount()
    ' Успешный экспорт в обоих режимах использует пригодное сохранённое НДС.
    ' Нагрузки исходного CM_BEND, его статусы и численные сравнения не меняются.
    exportCell.Value2 = "CM_COMP"
    snapshot = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    TestCheck "native.fullExport.validState.available", InStr(1, snapshot, "CM_COMP|PR1|StrengthState|", vbBinaryCompare) = 1
    exported = Audit03ExportAutoCADDocumentForTests(doc)
    TestCheck "native.fullExport.objects", exported > model.ConcreteCount + model.RebarCount
    TestCheck "native.fullExport.noSolve", solves = SectionEquilibriumSolveCount()
    For Each entity In doc.ModelSpace
        If CStr(entity.ObjectName) = "AcDbRegion" Then
            If StrComp(CStr(entity.Layer), concreteLayer, vbTextCompare) = 0 Then
                count = count + 1: area = area + CDbl(entity.Area)
            End If
            If StrComp(CStr(entity.Layer), rebarLayer, vbTextCompare) = 0 Then bars = bars + 1
        End If
    Next entity
    For i = 1 To model.ConcreteCount: expectedArea = expectedArea + model.ConcreteArea(i): Next i
    TestCheck "native.fullExport.count", count = model.ConcreteCount And bars = model.RebarCount
    TestNear "native.fullExport.areaInMm2", expectedArea, area, 0.0000001
    Dim restored As CSectionModel
    Set restored = ReadSectionGeometryFromResults(ThisWorkbook)
    TestCheck "native.fullExport.meshSourcePreserved", restored.ConcreteMeshSource = "AutoCADContours"
    GoTo Finish
Failed:
    TestCheck "native.fullExport.runtime." & CStr(Err.Number) & "; " & Err.Description, False
Finish:
    If Not IsEmpty(saved) Then table.Formula = saved
    If Not IsEmpty(savedProfiles) Then profiles.Formula = savedProfiles
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Один настоящий Region содержит два отверстия; отдельный слой отверстий пуст.
Private Sub TestCheckNativeTwoHoles(ByVal doc As Object, ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim outer As Object, hole As Object, center(0 To 2) As Double, curves(0 To 0) As Object, regions As Variant, i As Long
    Dim source As Object, query As CSectionGeometryQuery, domain As CConcreteRegion, geometry As CGeometryRegion
    Dim importer As CAutoCADSectionModelImporter, model As CSectionModel, mesh As CFiberMeshBuilder, bars As CRebarLayout
    Dim generated As CSectionModel, x As Double, contourReader As CAutoCADContourReader
    Audit03EnsureAutoCADLayerForTests doc, "CM_TWO_OUT", 30
    Audit03EnsureAutoCADLayerForTests doc, "CM_TWO_BAR", 1
    Audit03EnsureAutoCADLayerForTests doc, "CM_TWO_EMPTY", 8
    Set source = doc.ModelSpace.AddCircle(center, 120#): Set curves(0) = source
    regions = doc.ModelSpace.AddRegion(curves): Set outer = regions(LBound(regions)): source.Delete: outer.Layer = "CM_TWO_OUT"
    For i = 1 To 2
        center(0) = IIf(i = 1, -45#, 45#)
        Set source = doc.ModelSpace.AddCircle(center, IIf(i = 1, 23#, 31#)): Set curves(0) = source
        regions = doc.ModelSpace.AddRegion(curves): Set hole = regions(LBound(regions)): source.Delete
        outer.Boolean 2, hole
    Next i
    Set contourReader = New CAutoCADContourReader: Set domain = contourReader.ReadRegion(outer)
    Set geometry = New CGeometryRegion: geometry.Initialize domain
    Set mesh = BuildConfiguredConcreteMesh(geometry, settings, units): Set bars = New CRebarLayout
    For i = 1 To 3
        x = (i - 2) * 70#: center(0) = x: center(1) = -60#
        bars.AddBar "R" & CStr(i), x, -60#, 16#, 0#, "Rebar", "", geometry
        Set source = doc.ModelSpace.AddCircle(center, 8#): Set curves(0) = source
        regions = doc.ModelSpace.AddRegion(curves): Set source = regions(LBound(regions))
        source.Layer = "CM_TWO_BAR": curves(0).Delete
    Next i
    Set importer = New CAutoCADSectionModelImporter
    Set model = importer.ImportFromModelSpace(doc.ModelSpace, "CM_TWO_EMPTY", "CM_TWO_BAR", "Rebar", 0#, units, "CM_TWO_OUT", "", doc.Layers, settings)
    Set generated = TestBuildGeneratedSectionModel(mesh, bars): TestCompareModels "native.twoHoles", generated, model, False
    Set query = New CSectionGeometryQuery: query.Initialize model
    TestNear "native.twoHoles.area", query.ConcreteDomain.Area, GEOM_PI * (120# * 120# - 23# * 23# - 31# * 31#), 0.0000001
    TestCheck "native.twoHoles.empty", Not query.ContainsPoint(query.ConcreteDomain, -45#, 0#) And Not query.ContainsPoint(query.ConcreteDomain, 45#, 0#)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт шесть контрольных областей: круг, прямоугольник, два объединения прямоугольников, скруглённый контур и контур с отверстием.
Private Function TestControlGeometry(ByVal shape As Long) As CGeometryRegion
    Dim polygonSpec As CPolygonSpec, polygonPoint As Variant
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, definition As CSectionDefinition
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "ContourMeshFixture"
    Select Case shape
        Case 1
            Set path = facade.CirclePath("Boundary", 0#, 0#, 150#)
        Case 2
            Set path = facade.RectanglePath("Boundary", 0#, 0#, 300#, 400#)
        Case 3
            Set path = facade.AdjacentRectangles("Boundary", Array(Array(0#, 0#, 300#, 120#), Array(0#, 120#, 120#, 240#)))
        Case 4
            Set path = facade.AdjacentRectangles("Boundary", Array(Array(-83#, 41#, 300#, 120#), Array(-23#, 161#, 200#, 240#)))
        Case 5
            Set polygonSpec = New CPolygonSpec
            polygonSpec.AddVertex "BL", -137#, -227#, 40#
            polygonSpec.AddVertex "BR", 163#, -227#, 40#
            polygonSpec.AddVertex "TR", 163#, 173#, 40#
            polygonSpec.AddVertex "TL", -137#, 173#, 40#
            polygonSpec.Seal
            Set path = facade.RoundedPolygon("Boundary", polygonSpec)
        Case 6
            Set polygonSpec = New CPolygonSpec
            polygonSpec.AddVertex "BL", -200#, -250#, 35#
            polygonSpec.AddVertex "BR", 200#, -250#, 35#
            polygonSpec.AddVertex "TR", 200#, 250#, 35#
            polygonSpec.AddVertex "TL", -200#, 250#, 35#
            polygonSpec.Seal
            Set path = facade.RoundedPolygon("Boundary", polygonSpec)
    End Select
    facade.Outer "OUTER", "Boundary"
    If shape = 6 Then
        Set polygonSpec = New CPolygonSpec
        polygonSpec.AddVertex "BL", -83#, -163#, 25#
        polygonSpec.AddVertex "BR", 117#, -163#, 25#
        polygonSpec.AddVertex "TR", 117#, 117#, 25#
        polygonSpec.AddVertex "TL", -83#, 117#, 25#
        polygonSpec.Seal
        Set path = facade.RoundedPolygon("VoidPath", polygonSpec)
        facade.Opening "VOID", "VoidPath", "OUTER"
    End If
    Set definition = facade.Finish
    Set TestControlGeometry = TestFixtureGeometry(definition.ConcreteDomain)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт минимальную настоящую таблицу настроек сетки и единиц в собственной временной книге.
Private Function TestMeshSettings(Optional ByVal subdivision As Long = 2, Optional ByVal unit As String = "mm", _
        Optional ByVal stepX As Double = 40#, Optional ByVal stepY As Double = 30#) As CSystemSettingsReader
    mConfig.Cells.ClearContents
    Dim data(1 To 10, 1 To 3) As Variant
    data(1, 1) = "Key": data(1, 2) = "Value": data(1, 3) = "Unit"
    data(2, 1) = "Mesh.StepX": data(2, 2) = stepX
    data(3, 1) = "Mesh.StepY": data(3, 2) = stepY
    data(4, 1) = "Mesh.BoundarySubdivisions": data(4, 2) = subdivision
    data(5, 1) = "Units.Length.Input": data(5, 2) = unit
    data(6, 1) = "Units.Length.Output": data(6, 2) = unit
    data(7, 1) = "Units.Area.Output": data(7, 2) = unit & "2"
    data(8, 1) = "Mesh.StepMode": data(8, 2) = "User"
    data(9, 1) = "Mesh.MinElementArea": data(9, 2) = 0#
    data(10, 1) = "Units.Area.Input": data(10, 2) = unit & "2"
    mConfig.Range("A1:C10").Value2 = data
    Set TestMeshSettings = New CSystemSettingsReader
    TestMeshSettings.LoadFromRange mConfig.Range("A1:C10")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает автоматический шаг с независимым ручным эталоном на шести формах,
' включая отверстие и смещения, оба порядка габаритов, единицы и дробление границы.
' Ошибочные неиспользуемые ручные поля не должны мешать Auto; неизвестный режим отклоняется.
Public Function RunAutomaticMeshTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    Set mFixture = Application.Workbooks.Add(-4167)
    Set mConfig = mFixture.Worksheets(1): mConfig.Name = "AutoMeshConfig"
    Dim shape As Long, subdivision As Variant, factor As Variant, geometry As CGeometryRegion
    For shape = 1 To 8
        If shape <= 6 Then
            Set geometry = TestControlGeometry(shape)
        ElseIf shape = 7 Then
            Set geometry = TestFixtureBox(500#, 200#, 1000000000#, -1000000000#)
        Else
            Set geometry = TestFixtureBox(20#, 60#)
        End If
        For Each subdivision In Array(1, 2)
            For Each factor In Array(0.05, 0.1, 0.2)
                TestAutomaticMeshCase geometry, CDbl(factor), CLng(subdivision), shape
            Next factor
        Next subdivision
    Next shape
    TestAutomaticMeshInputs
    TestMinimumElementAreaInput
    TestAutomaticSectionBuilder
    GoTo Finish
Failed:
    TestCheck "autoMesh.runtime." & CStr(Err.Number) & "; " & Err.Source & "; " & Err.Description, False
Finish:
    On Error Resume Next
    If Not mFixture Is Nothing Then mFixture.Close False
    Set mFixture = Nothing: Set mConfig = Nothing
    On Error GoTo 0
    RunAutomaticMeshTests = mReport & "TOTAL_AUTO_MESH: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed)
End Function


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет один физический порог в мм2, см2 и м2 при обоих способах ячеек,
' обе десятичные записи и адресный отказ до построения при ошибочном значении.
Private Sub TestMinimumElementAreaInput()
    Dim settings As CSystemSettingsReader, units As CUnitSystem, options As TSectionMeshOptions
    Dim cellMode As Variant, areaUnit As Variant, areaValue As Double, value As Variant
    Dim number As Long, description As String, invalidIndex As Long, invalidLabels As Variant
    For Each cellMode In Array("Adaptive", "Contour")
        For Each areaUnit In Array("mm2", "cm2", "m2")
            Set settings = TestMeshSettings()
            mConfig.Range("A11").Value2 = "Mesh.CellGeometry": mConfig.Range("B11").Value2 = CStr(cellMode)
            Select Case CStr(areaUnit)
                Case "mm2": areaValue = 25#
                Case "cm2": areaValue = 0.25
                Case "m2": areaValue = 0.000025
            End Select
            mConfig.Range("B9").Value2 = areaValue: mConfig.Range("B10").Value2 = CStr(areaUnit)
            settings.LoadFromRange mConfig.Range("A1:C11"): Set units = TestSettingsUnits(settings)
            options = ReadSectionMeshOptions(settings, units)
            TestCheck "minArea.units." & CStr(cellMode) & "." & CStr(areaUnit), Abs(options.MinArea - 25#) <= 0.000000000001
            TestCheck "minArea.mode." & CStr(cellMode) & "." & CStr(areaUnit), options.CellGeometry = CStr(cellMode)
        Next areaUnit
    Next cellMode
    For Each value In Array("0.25", "0,25", 0.25)
        Set settings = TestMeshSettings(): mConfig.Range("B9").Value2 = value: mConfig.Range("B10").Value2 = "cm2"
        settings.LoadFromRange mConfig.Range("A1:C10"): Set units = TestSettingsUnits(settings)
        TestCheck "minArea.decimal." & CStr(value), Abs(ReadMinimumElementArea(settings, units) - 25#) <= 0.000000000001
    Next value
    invalidLabels = Array("negative", "blank", "text", "excelError", "overflow")
    For Each value In Array(-1#, "", "abc", CVErr(xlErrValue), "1e308")
        Set settings = TestMeshSettings(): mConfig.Range("B9").Value2 = value: mConfig.Range("B10").Value2 = "m2"
        On Error Resume Next
        Err.Clear
        settings.LoadFromRange mConfig.Range("A1:C10")
        number = Err.Number: description = Err.Description ' Сохраняем первый отказ: повторный запрос незагруженного ключа уничтожил бы точный адрес исходной ошибки Excel.
        If number = 0 Then
            Set units = TestSettingsUnits(settings)
            number = Err.Number: description = Err.Description
        End If
        If number = 0 Then
            areaValue = ReadMinimumElementArea(settings, units)
            number = Err.Number: description = Err.Description
        End If
        Err.Clear
        On Error GoTo 0
        TestCheck "minArea.invalid." & CStr(invalidLabels(invalidIndex)) & ".address; Err=" & CStr(number) & _
            "; " & Replace(Replace(description, vbCr, " "), vbLf, " "), _
            number <> 0 And InStr(description, "Mesh.MinElementArea") > 0 And InStr(description, "B9") > 0
        invalidIndex = invalidIndex + 1
    Next value
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит одну Auto-сетку через рабочее чтение и её ручной эталон; сравнивает каждый
' элемент и импорт тех же контуров. Нечисловые ручные шаги намеренно оставлены в Auto.
Private Sub TestAutomaticMeshCase(ByVal geometry As CGeometryRegion, ByVal factor As Double, ByVal subdivision As Long, ByVal shape As Long)
    Dim settings As CSystemSettingsReader, units As CUnitSystem, actual As CFiberMeshBuilder, expected As CFiberMeshBuilder
    Dim step As Double, i As Long, label As String, contours As CSectionContours, bars As CRebarLayout
    Dim imported As CSectionModel, generated As CSectionModel, importer As CAutoCADSectionModelImporter, unit As String
    Select Case shape Mod 3
        Case 0: unit = "mm"
        Case 1: unit = "cm"
        Case 2: unit = "m"
    End Select
    Set settings = TestMeshSettings(subdivision, unit)
    mConfig.Range("B8").Value2 = "Auto " & Replace(CStr(factor), ",", ".")
    mConfig.Range("B2").Formula = "=1/0": mConfig.Range("B3").Value2 = "unused"
    settings.LoadFromRange mConfig.Range("A1:C10"): Set units = TestSettingsUnits(settings)
    Set actual = BuildConfiguredConcreteMesh(geometry, settings, units)
    step = geometry.Width: If geometry.Height < step Then step = geometry.Height
    step = step * factor
    Set expected = New CFiberMeshBuilder: expected.BuildMesh geometry, step, step, 1, subdivision
    label = "autoMesh." & CStr(shape) & "." & CStr(factor) & ".sub" & CStr(subdivision)
    TestCheck label & ".count", actual.FiberCount = expected.FiberCount
    For i = 1 To expected.FiberCount
        TestCheck label & ".fiber" & CStr(i), actual.FiberX(i) = expected.FiberX(i) And actual.FiberY(i) = expected.FiberY(i) And _
            actual.FiberWidth(i) = expected.FiberWidth(i) And actual.FiberHeight(i) = expected.FiberHeight(i) And actual.FiberArea(i) = expected.FiberArea(i)
    Next i
    Set contours = New CSectionContours: geometry.BuildContours contours
    Set bars = TestControlBars(expected, geometry): Set generated = TestBuildGeneratedSectionModel(expected, bars)
    Set importer = New CAutoCADSectionModelImporter
    Set imported = importer.ImportFromModelSpace(TestFakeSpace(contours, bars), "CONCRETE", "REBAR", "Rebar", 0#, units, _
        "OUTER", "OPENING", TestFakeLayers(), settings)
    TestCompareModels label & ".import", generated, imported
    TestCheck label & ".manualKept", mConfig.Range("B2").Formula = "=1/0" And mConfig.Range("B3").Value2 = "unused"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет обязательность режима, его фактический адрес, строгие ручные шаги
' и сохранение независимого дробления границы в автоматическом режиме.
Private Sub TestAutomaticMeshInputs()
    Dim settings As CSystemSettingsReader, units As CUnitSystem, mesh As CFiberMeshBuilder, geometry As CGeometryRegion
    Dim mode As Variant, number As Long, description As String, options As TSectionMeshOptions
    For Each mode In Array(vbNullString, "TODO", "Auto 0.3", CVErr(2015))
        Set settings = TestMeshSettings(): mConfig.Range("B8").Value2 = mode
        On Error Resume Next
        settings.LoadFromRange mConfig.Range("A1:C10")
        number = Err.Number: description = Err.Description: Err.Clear
        If number = 0 Then
            Set units = TestSettingsUnits(settings): options = ReadSectionMeshOptions(settings, units)
            number = Err.Number: description = Err.Description: Err.Clear
        End If
        On Error GoTo 0
        TestCheck "autoMesh.invalidMode; " & description, number <> 0 And InStr(description, "B8") > 0 And InStr(description, "Mesh.StepMode") > 0
    Next mode
    Set settings = TestMeshSettings(2, "cm", 4#, 3#): Set units = TestSettingsUnits(settings)
    options = ReadSectionMeshOptions(settings, units)
    TestCheck "autoMesh.userExact", options.StepX = 40# And options.StepY = 30# And options.AutoStepFactor = 0# And options.BoundarySubdivisions = 2
    mConfig.Range("B2").Value2 = 0#: settings.LoadFromRange mConfig.Range("A1:C10")
    On Error Resume Next
    options = ReadSectionMeshOptions(settings, units): number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo 0
    TestCheck "autoMesh.userInvalid", number <> 0 And InStr(description, "B2") > 0
    mConfig.Range("B8").Value2 = "Auto 0.1": mConfig.Range("B4").Value2 = 0#
    settings.LoadFromRange mConfig.Range("A1:C10")
    On Error Resume Next
    options = ReadSectionMeshOptions(settings, units): number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo 0
    TestCheck "autoMesh.autoInvalidSubdivision", number <> 0 And InStr(description, "B4") > 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет передачу AutoStepFactor через полноценный CSectionBuilder:
' ручные шаги при Auto не требуются, два уровня сохраняются; основной шаг 30 мм, минимальный 7.5 мм.
Private Sub TestAutomaticSectionBuilder()
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, definition As CSectionDefinition
    Dim builder As CSectionBuilder, options As TSectionMeshOptions, model As CSectionModel, mesh As CFiberMeshBuilder
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "TestAutoMesh"
    Set path = facade.RectanglePath("Boundary", 0#, 0#, 300#, 500#): facade.Outer "OUTER", path.ID
    Set path = facade.LinePath("Axis", 40#, 40#, 260#, 40#): facade.RebarByCount "Main", path.ID, 3, 16#
    Set definition = facade.Finish
    options.AutoStepFactor = 0.1: options.BoundarySubdivisions = 2: options.MaterialID = 1
    Set builder = New CSectionBuilder: Set model = builder.Build(definition, options): Set mesh = builder.Mesh
    TestCheck "autoMesh.sectionBuilder.bars", model.RebarCount = 3
    TestCheck "autoMesh.sectionBuilder.mainStep", mesh.FiberWidth(1) = 30# And mesh.FiberHeight(1) = 30#
    Dim i As Long, refined As Boolean
    For i = 1 To mesh.FiberCount
        If mesh.FiberWidth(i) = 7.5 And mesh.FiberHeight(i) = 7.5 Then refined = True
    Next i
    TestCheck "autoMesh.sectionBuilder.refined", refined
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт систему единиц из того же снимка настроек, который используется для контрольной сетки.
Private Function TestSettingsUnits(ByVal settings As CSystemSettingsReader) As CUnitSystem
    Set TestSettingsUnits = New CUnitSystem: TestSettingsUnits.LoadFromSettings settings
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт именованные слои бетона, арматуры и контуров для импорта без установленного AutoCAD.
Private Function TestFakeLayers() As Collection
    Set TestFakeLayers = New Collection
    Dim name As Variant, layer As CTestFakeAcadContour
    For Each name In Array("CONCRETE", "REBAR", "OUTER", "OPENING")
        Set layer = New CTestFakeAcadContour
        layer.TestInitializeLoop "Layer", CStr(name), CStr(name), Nothing
        TestFakeLayers.Add layer, CStr(name)
    Next name
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Представляет сохранённые линии, дуги, окружности и арматуру как объекты имитируемого ModelSpace.
Private Function TestFakeSpace(ByVal contours As CSectionContours, ByVal bars As CRebarLayout) As Collection
    Dim loops As Object, edges As Collection, i As Long, key As Variant, segment As CTestFakeAcadContour
    Dim entity As CTestFakeAcadContour, rebar As CTestFakeAcadRegion, angle As Double, layer As String
    Set TestFakeSpace = New Collection: Set loops = CreateObject("Scripting.Dictionary")
    If Not bars Is Nothing Then
        For i = 1 To bars.Count
            Set rebar = New CTestFakeAcadRegion
            rebar.TestInitialize bars.Area(i), bars.X(i), bars.Y(i), 1#, 1#, 0#, "REBAR", "BAR" & CStr(i)
            TestFakeSpace.Add rebar
        Next i
    End If
    For i = 1 To contours.Count
        layer = "OUTER": If contours.LoopRole(i) = "Opening" Then layer = "OPENING"
        key = layer & "|" & contours.LoopID(i)
        If Not loops.Exists(key) Then
            Set edges = New Collection: loops.Add key, edges
        End If
        Set edges = loops(key): Set segment = New CTestFakeAcadContour
        Select Case contours.SegmentType(i)
            Case "CONTOUR_LINE"
                segment.TestInitializeLine contours.StartX(i), contours.StartY(i), contours.EndX(i), contours.EndY(i)
            Case "CONTOUR_ARC"
                angle = TestAtan2(contours.StartY(i) - contours.CenterY(i), contours.StartX(i) - contours.CenterX(i))
                segment.TestInitializeArc contours.CenterX(i), contours.CenterY(i), contours.Radius(i), angle, contours.SweepAngle(i)
            Case "CONTOUR_CIRCLE"
                segment.TestInitializeCircle contours.StartX(i), contours.StartY(i), contours.Radius(i)
        End Select
        edges.Add segment
    Next i
    For Each key In loops.Keys
        Set entity = New CTestFakeAcadContour
        entity.TestInitializeLoop "AcDbRegion", Left$(CStr(key), InStr(CStr(key), "|") - 1), CStr(key), loops(key)
        TestFakeSpace.Add entity
    Next key
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Определяет угол радиус-вектора по обеим координатам, сохраняя нужную четверть для имитируемой дуги.
Private Function TestAtan2(ByVal y As Double, ByVal x As Double) As Double
    If x = 0# Then
        TestAtan2 = Sgn(y) * GEOM_PI / 2#
    Else
        TestAtan2 = Atn(y / x)
        If x < 0# Then TestAtan2 = TestAtan2 + GEOM_PI
    End If
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выбирает четыре положения из существующих центров сетки, чтобы обе сравниваемые модели получили одну арматуру.
Private Function TestControlBars(ByVal mesh As CFiberMeshBuilder, ByVal geometry As CGeometryRegion) As CRebarLayout
    Set TestControlBars = New CRebarLayout
    Dim index As Long, i As Long
    For i = 1 To 4
        index = 1 + CLng(Fix(CDbl(mesh.FiberCount - 1) * CDbl(i) / 5#)) ' Положения 1/5–4/5 списка не привязаны к конкретному числу волокон.
        TestControlBars.AddBar "SOURCE" & CStr(i), mesh.FiberX(index), mesh.FiberY(index), 16#, 0#, "Rebar", "", geometry
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает модель общего генератора с импортом тех же контуров; проверяет происхождение сетки, центры волокон и выбранные расчётные случаи.
Private Sub TestCompareGeneratedAndImported(ByVal geometry As CGeometryRegion, ByVal label As String, _
        ByVal subdivision As Long, ByVal solve As Boolean)
    Dim settings As CSystemSettingsReader, units As CUnitSystem, mesh As CFiberMeshBuilder, bars As CRebarLayout
    Dim generated As CSectionModel, imported As CSectionModel, importer As CAutoCADSectionModelImporter
    Dim query As CSectionGeometryQuery, adapter As CGeometryRegion, available As Boolean
    Set settings = TestMeshSettings(subdivision): Set units = TestSettingsUnits(settings)
    Set mesh = BuildConfiguredConcreteMesh(geometry, settings, units): Set bars = TestControlBars(mesh, geometry)
    Set generated = TestBuildGeneratedSectionModel(mesh, bars)
    geometry.BuildContours generated.Contours
    Set importer = New CAutoCADSectionModelImporter
    Set imported = importer.ImportFromModelSpace(TestFakeSpace(generated.Contours, bars), "CONCRETE", "REBAR", "Rebar", 0#, units, "OUTER", "OPENING", TestFakeLayers(), settings)
    TestCompareModels label, generated, imported
    TestCheck label & ".source", imported.SourceType = "AutoCADImport" And imported.ConcreteMeshSource = "AutoCADContours"
    TestCheck label & ".notice", InStr(importer.ResultMessage(imported), vbCrLf & vbCrLf & CONTOUR_MESH_GENERATION_NOTICE) > 0
    TestCheck label & ".typedWarning", importer.HasImportWarnings
    Set query = New CSectionGeometryQuery: query.Initialize imported
    Set adapter = New CGeometryRegion: adapter.Initialize query.ConcreteDomain
    TestNear label & ".exactArea", adapter.AnalyticalArea(available), geometry.AnalyticalArea(available), 0.0000001
    TestCheckCellCenters label, mesh, geometry
    Dim i As Long
    For i = 1 To imported.ConcreteCount
        TestCheck label & ".importedCenter." & CStr(i), adapter.ContainsPoint(imported.ConcreteX(i), imported.ConcreteY(i))
    Next i
    If solve Then TestCompareSolutions label, generated, imported, True
    If label = "3.sub2" Then TestCheckBatchPair generated, imported
    mReport = mReport & "SECTION " & label & ": concrete=" & CStr(imported.ConcreteCount) & "; rebar=" & CStr(imported.RebarCount) & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Полный расчёт двух одинаковых Г-сечений: НДС при заданной нагрузке, несущая
' способность и трещины СП 35. Исходная книга — тестовая копия; настройки
' восстанавливаются также при ошибке.
Private Sub TestCheckBatchPair(ByVal first As CSectionModel, ByVal second As CSectionModel)
    Dim profileTable As Object, systemTable As Object, savedProfiles As Variant, savedSystem As Variant
    Dim i As Long, key As String, settings As CSystemSettingsReader, units As CUnitSystem, provider As CMaterialModelProvider
    Dim profiles As CCalculationProfileCatalog, a As CBatchSectionCalculator, b As CBatchSectionCalculator
    Dim ra As CCombinationResult, rb As CCombinationResult, props As CSectionPropertiesCalculator, writer As CNDMResultsWriter
    On Error GoTo Failed
    Set profileTable = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Set systemTable = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    savedProfiles = profileTable.Formula: savedSystem = systemTable.Formula
    For i = 2 To profileTable.Rows.Count
        key = CStr(profileTable.Cells(i, 2).Value2)
        Select Case key
            Case "Calculation.Strength.DirectState", "Calculation.Strength.Capacity", "Calculation.Crack.Width"
                profileTable.Cells(i, 3).Value2 = "Yes"
            Case "Calculation.Stability.Enabled", "Calculation.Crack.Longitudinal"
                profileTable.Cells(i, 3).Value2 = "No"
        End Select
    Next i
    For i = 2 To systemTable.Rows.Count
        If CStr(systemTable.Cells(i, 1).Value2) = "SLS.Crack.Code" Then systemTable.Cells(i, 2).Value2 = "SP35"
    Next i
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = TestSettingsUnits(settings)
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
    Set props = New CSectionPropertiesCalculator: props.CalculateConcrete first
    Set a = New CBatchSectionCalculator: a.Initialize first, provider: Set a.ProfileCatalog = profiles: a.ApplySettings settings, units
    Set b = New CBatchSectionCalculator: b.Initialize second, provider: Set b.ProfileCatalog = profiles: b.ApplySettings settings, units
    a.ApplyLoadReference props.CentroidX, props.CentroidY, 0#, 0#
    b.ApplyLoadReference props.CentroidX, props.CentroidY, 0#, 0#
    a.AddCombination "CM_COMP", -500000#, 0#, 0#, "PR1", "Contour mesh control", "Auto"
    b.AddCombination "CM_COMP", -500000#, 0#, 0#, "PR1", "Contour mesh control", "Auto"
    a.AddCombination "CM_BEND", -50000#, 25000000#, -10000000#, "PR1", "Contour mesh control", "Auto"
    b.AddCombination "CM_BEND", -50000#, 25000000#, -10000000#, "PR1", "Contour mesh control", "Auto"
    a.Execute: b.Execute
    For i = 1 To 2
        Set ra = a.ResultAt(i): Set rb = b.ResultAt(i)
        TestCheck "batch.status." & CStr(i), ra.Status = rb.Status
        TestCheck "batch.widthStatus." & CStr(i), ra.NormalCrackStatus = rb.NormalCrackStatus
        TestCheck "batch.noInputErr." & CStr(i) & "; " & rb.OverallMeta.ResultComment, rb.Status <> "InputErr"
        TestCheck "batch.capacityCalculated." & CStr(i), rb.StrengthResult.Capacity.ResultMeta.Calculated
        TestNear "batch.capacity." & CStr(i), ra.StrengthResult.Capacity.LambdaCapacity, rb.StrengthResult.Capacity.LambdaCapacity, 0.0000000001
        TestNear "batch.formation." & CStr(i), ra.CrackResult.Formation.LambdaCrc, rb.CrackResult.Formation.LambdaCrc, 0.0000000001
        TestNear "batch.width." & CStr(i), ra.CrackResult.Width.CrackWidth, rb.CrackResult.Width.CrackWidth, 0.0000000001
        TestNear "batch.Ar." & CStr(i), ra.CrackResult.Width.Abt, rb.CrackResult.Width.Abt, 0.00000001
        mReport = mReport & "BATCH " & CStr(i) & ": status=" & rb.Status & "; normalCrack=" & rb.NormalCrackStatus & _
            "; capacity=" & CStr(rb.StrengthResult.Capacity.LambdaCapacity) & "; width=" & CStr(rb.CrackResult.Width.CrackWidth) & "; Ar=" & CStr(rb.CrackResult.Width.Abt) & _
            "; reason=" & rb.OverallMeta.ResultComment & vbCrLf
    Next i
    TestCheck "batch.widthActuallyCalculated", b.ResultAt(2).CrackResult.Width.ResultMeta.Calculated
    TestCheck "batch.SP35", b.ResultAt(2).CrackResult.Width.StandardCode = "SP35"
    TestNear "batch.actualRegionArea", b.ResultAt(2).CrackResult.Width.Abt, b.ResultAt(2).CrackResult.Width.InteractionRegion.Area, 0.00000001
    Set writer = New CNDMResultsWriter
    writer.WriteResults ThisWorkbook, second, PrepareSectionSnapshot(second, provider), b, units
    Dim restored As CSectionModel
    Set restored = ReadSectionGeometryFromResults(ThisWorkbook)
    TestCheck "batch.fullWriterSource", restored.ConcreteMeshSource = "AutoCADContours"
    GoTo Finish
Failed:
    TestCheck "batch.runtime." & CStr(Err.Number) & "; " & Err.Description, False
Finish:
    If Not IsEmpty(savedProfiles) Then profileTable.Formula = savedProfiles
    If Not IsEmpty(savedSystem) Then systemTable.Formula = savedSystem
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимый контроль площади и инерции после нативного CAD-преобразования;
' величины отклонений сохраняются в отчете.
Private Sub TestCompareCurvedModels(ByVal label As String, ByVal first As CSectionModel, ByVal second As CSectionModel)
    Dim a As CSectionPropertiesCalculator, b As CSectionPropertiesCalculator
    Set a = New CSectionPropertiesCalculator: a.CalculateConcrete first
    Set b = New CSectionPropertiesCalculator: b.CalculateConcrete second
    TestNear label & ".curvedArea", a.Area, b.Area, a.Area * 0.01
    TestNear label & ".curvedIx", a.Ixc, b.Ixc, a.Ixc * 0.01
    TestNear label & ".curvedIy", a.Iyc, b.Iyc, a.Iyc * 0.01
    TestCheck label & ".curvedBars", first.RebarCount = second.RebarCount
    mReport = mReport & "CURVED " & label & ": GeneratedCount=" & CStr(first.ConcreteCount) & "; CADCount=" & CStr(second.ConcreteCount) & _
        "; relativeArea=" & CStr(Abs(a.Area - b.Area) / a.Area) & "; relativeIx=" & CStr(Abs(a.Ixc - b.Ixc) / a.Ixc) & "; relativeIy=" & CStr(Abs(a.Iyc - b.Iyc) / a.Iyc) & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает каждый бетонный элемент и стержень в исходном порядке; проверку исходных CAD-идентификаторов можно отключить для сохранённой модели.
Private Sub TestCompareModels(ByVal label As String, ByVal first As CSectionModel, ByVal second As CSectionModel, Optional ByVal checkHandles As Boolean = True)
    TestCheck label & ".counts", first.ConcreteCount = second.ConcreteCount And first.RebarCount = second.RebarCount
    If first.ConcreteCount <> second.ConcreteCount Or first.RebarCount <> second.RebarCount Then Exit Sub
    Dim i As Long
    For i = 1 To first.ConcreteCount
        TestNear label & ".x." & CStr(i), first.ConcreteX(i), second.ConcreteX(i)
        TestNear label & ".y." & CStr(i), first.ConcreteY(i), second.ConcreteY(i)
        TestNear label & ".a." & CStr(i), first.ConcreteArea(i), second.ConcreteArea(i)
        TestNear label & ".width." & CStr(i), first.ConcreteWidth(i), second.ConcreteWidth(i)
        TestNear label & ".height." & CStr(i), first.ConcreteHeight(i), second.ConcreteHeight(i)
        TestNear label & ".Ix." & CStr(i), first.ConcreteLocalIx(i), second.ConcreteLocalIx(i)
        TestNear label & ".Iy." & CStr(i), first.ConcreteLocalIy(i), second.ConcreteLocalIy(i)
        TestCheck label & ".id." & CStr(i), first.ConcreteID(i) = second.ConcreteID(i)
    Next i
    For i = 1 To first.RebarCount
        TestNear label & ".barX." & CStr(i), first.RebarX(i), second.RebarX(i)
        TestNear label & ".barY." & CStr(i), first.RebarY(i), second.RebarY(i)
        TestNear label & ".barA." & CStr(i), first.RebarArea(i), second.RebarArea(i)
        If checkHandles Then TestCheck label & ".barHandle." & CStr(i), second.RebarSourceHandle(i) = "BAR" & CStr(i)
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает четыре нагрузки и три режима материалов у каждой переданной пары моделей; в основном наборе это 72 сравнения.
Private Sub TestCompareSolutions(ByVal label As String, ByVal first As CSectionModel, ByVal second As CSectionModel, ByVal exactPair As Boolean)
    Dim settings As CSystemSettingsReader, units As CUnitSystem, provider As CMaterialModelProvider
    Dim a As CSectionSolver, b As CSectionSolver, props As CSectionPropertiesCalculator
    Dim purpose As Variant, loadIndex As Long, n As Double, mx As Double, my As Double, height As Double, width As Double
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double, prefix As String
    Dim relativeTolerance As Double
    If Not exactPair Then relativeTolerance = 0.01 ' Нативный CAD может изменить дискретизацию дуговой границы; для такой пары проверяются интегральные величины с 1 %.
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Set props = New CSectionPropertiesCalculator: props.CalculateConcrete first
    props.CalculateProjection first, 1#, 0#, False, minX, maxX
    props.CalculateProjection first, 0#, 1#, False, minY, maxY
    width = maxX - minX: height = maxY - minY
    For Each purpose In Array(cpStrength, cpMcrc, cpCrackedNDS)
        For loadIndex = 0 To 3
            n = -2# * props.Area
            mx = n * props.CentroidY: my = n * props.CentroidX
            If loadIndex = 1 Or loadIndex = 3 Then mx = mx + props.Area * height * 0.2
            If loadIndex = 2 Or loadIndex = 3 Then my = my - props.Area * width * 0.15
            Set a = New CSectionSolver: Set b = New CSectionSolver
            a.ApplySettings settings, units: b.ApplySettings settings, units
            a.Solve first, provider.ConcreteMaterial(CLng(purpose)), provider.SteelMaterial(CLng(purpose)), n, mx, my
            b.Solve second, provider.ConcreteMaterial(CLng(purpose)), provider.SteelMaterial(CLng(purpose)), n, mx, my
            prefix = label & ".purpose" & CStr(purpose) & ".load" & CStr(loadIndex)
            TestCheck prefix & ".converged; " & a.StopReason & "; " & b.StopReason, a.Converged And b.Converged
            TestNear prefix & ".eps", a.Epsilon0, b.Epsilon0, 0.000000000001 + Abs(a.Epsilon0) * relativeTolerance
            TestNear prefix & ".kx", a.KappaX, b.KappaX, 0.000000000000001 + Abs(a.KappaX) * relativeTolerance
            TestNear prefix & ".ky", a.KappaY, b.KappaY, 0.000000000000001 + Abs(a.KappaY) * relativeTolerance
            TestNear prefix & ".N", b.Nint, n, 2#
            TestNear prefix & ".Mx", b.Mxint, mx, 2000#
            TestNear prefix & ".My", b.Myint, my, 2000#
            TestNear prefix & ".minConcrete", a.MinConcreteStress, b.MinConcreteStress, 0.00000001 + Abs(a.MinConcreteStress) * relativeTolerance
            TestNear prefix & ".maxConcrete", a.MaxConcreteStress, b.MaxConcreteStress, 0.00000001 + GeomMax(Abs(a.MaxConcreteStress), Abs(a.MinConcreteStress)) * relativeTolerance
            TestNear prefix & ".minSteel", a.MinSteelStress, b.MinSteelStress, 0.00000001 + Abs(a.MinSteelStress) * relativeTolerance
            TestNear prefix & ".maxSteel", a.MaxSteelStress, b.MaxSteelStress, 0.00000001 + Abs(a.MaxSteelStress) * relativeTolerance
            mReport = mReport & "SOLVE " & prefix & ": eps=" & CStr(b.Epsilon0) & "; kx=" & CStr(b.KappaX) & "; ky=" & CStr(b.KappaY) & "; iterations=" & CStr(b.Iterations) & vbCrLf
        Next loadIndex
    Next purpose
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет сетку вогнутого контура, двух отверстий и раздельных областей, а также очистку ранее подготовленной области после ошибочной инициализации.
Private Sub TestCheckArbitraryDomains()
    Dim query As CSectionGeometryQuery, region As CConcreteRegion, geometry As CGeometryRegion, mesh As CFiberMeshBuilder
    Set query = New CSectionGeometryQuery
    Set region = TestFlatPolygon(Array(0#, 0#, 100#, 0#, 100#, 100#, 60#, 100#, 60#, 40#, 40#, 40#, 40#, 100#, 0#, 100#))
    Set geometry = New CGeometryRegion: geometry.Initialize region
    Set mesh = New CFiberMeshBuilder: mesh.BuildMesh geometry, 100#, 100#, 1, 3
    TestNear "concave.subcells.area", TestAdaptiveMeshArea(mesh), 8437.5 ' Независимый растр 8x8: 64 - 10 центров выреза, каждый элемент площадью 156.25 мм2.
    TestCheckCellCenters "concave", mesh, geometry
    Set region = TestFlatPolygon(Array(0#, 0#, 300#, 0#, 300#, 200#, 0#, 200#))
    Set region = query.SubtractContainedRegion(region, query.CircleRegion(70#, 70#, 23#))
    Set region = query.SubtractContainedRegion(region, query.CircleRegion(230#, 130#, 31#))
    geometry.Initialize region: mesh.BuildMesh geometry, 40#, 30#, 1, 3
    TestCheckCellCenters "twoHoles", mesh, geometry
    TestCheck "twoHoles.exclude", Not geometry.ContainsPoint(70#, 70#) And Not geometry.ContainsPoint(230#, 130#)
    TestCheck "twoHoles.outerBoundary", geometry.ContainsPoint(0#, 0#)
    TestCheck "twoHoles.openingBoundary", Not geometry.ContainsPoint(93#, 70#)
    Dim parts As Collection
    Set parts = New Collection: parts.Add region
    parts.Add TestFlatPolygon(Array(400#, 0#, 500#, 0#, 500#, 100#, 400#, 100#))
    Set region = query.MergeDisjointRegions(parts): geometry.Initialize region
    mesh.BuildMesh geometry, 40#, 30#, 1, 3
    TestCheckCellCenters "disjoint", mesh, geometry
    TestCheck "disjoint.material", geometry.ContainsPoint(450#, 50#) And Not geometry.ContainsPoint(350#, 50#)
    Dim pointX() As Double, pointY() As Double, i As Long, message As String, number As Long
    geometry.GetExtremePoints pointX, pointY
    For i = 1 To UBound(pointX)
        TestCheck "extrema.realBoundary." & CStr(i), query.ContainsPoint(region, pointX(i), pointY(i), True)
    Next i
    On Error Resume Next
    geometry.Initialize Nothing: number = Err.Number: Err.Clear
    On Error GoTo 0
    TestCheck "adapter.invalidInitializeClearsOldShape", number <> 0 And Not geometry.IsValid(message)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Преобразует плоский список x,y в массив точек, используемый общим построителем бетонной области.
Private Function TestFlatPolygon(ByVal values As Variant) As CConcreteRegion
    Dim points() As Double, i As Long, query As CSectionGeometryQuery
    ReDim points(1 To (UBound(values) + 1) \ 2, 1 To 2)
    For i = 1 To UBound(points, 1)
        points(i, 1) = values((i - 1) * 2): points(i, 2) = values((i - 1) * 2 + 1)
    Next i
    Set query = New CSectionGeometryQuery: Set TestFlatPolygon = query.PolygonRegion(points)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет принадлежность всех центров бетону, положительные размеры ячеек и отсутствие повторных координат.
Private Sub TestCheckCellCenters(ByVal label As String, ByVal mesh As CFiberMeshBuilder, ByVal geometry As CGeometryRegion)
    Dim i As Long, seen As Object, key As String
    Set seen = CreateObject("Scripting.Dictionary")
    For i = 1 To mesh.FiberCount
        TestCheck label & ".center." & CStr(i), geometry.ContainsPoint(mesh.FiberX(i), mesh.FiberY(i))
        TestCheck label & ".positive." & CStr(i), mesh.FiberArea(i) > 0# And mesh.FiberWidth(i) > 0# And mesh.FiberHeight(i) > 0#
        key = CStr(mesh.FiberX(i)) & "|" & CStr(mesh.FiberY(i))
        TestCheck label & ".unique." & CStr(i), Not seen.Exists(key): seen(key) = True
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет независимость CAD-миллиметров от единиц Config и отказ при непригодной геометрии, неполных контурах и ошибках сетки.
Private Sub TestCheckImportContracts()
    Dim geometry As CGeometryRegion, contours As CSectionContours, bars As CRebarLayout, mesh As CFiberMeshBuilder
    Dim importer As CAutoCADSectionModelImporter, settings As CSystemSettingsReader, units As CUnitSystem
    Dim space As Collection, model As CSectionModel, baseline As CSectionModel, cell As CTestFakeAcadRegion, bad As CTestFakeAcadContour
    Dim number As Long, description As String, mode As Long, factor As Double, unit As Variant
    Set geometry = TestControlGeometry(3): Set contours = New CSectionContours: geometry.BuildContours contours
    Set settings = TestMeshSettings(): Set units = TestSettingsUnits(settings)
    Set mesh = BuildConfiguredConcreteMesh(geometry, settings, units): Set bars = TestControlBars(mesh, geometry)
    Set importer = New CAutoCADSectionModelImporter: Set space = TestFakeSpace(contours, bars)
    Set baseline = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, units, "OUTER", "OPENING", TestFakeLayers(), settings)
    For Each unit In Array("cm", "m")
        factor = 10#: If CStr(unit) = "m" Then factor = 1000#
        Set settings = TestMeshSettings(2, CStr(unit), 40# / factor, 30# / factor): Set units = TestSettingsUnits(settings)
        Set model = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, units, "OUTER", "OPENING", TestFakeLayers(), settings)
        TestCompareModels "units." & CStr(unit), baseline, model
    Next unit
    For mode = 0 To 6
        Set settings = TestMeshSettings(): Set units = TestSettingsUnits(settings)
        Set space = TestFakeSpace(contours, bars)
        If mode = 0 Then Set space = TestFakeSpace(New CSectionContours, bars)
        If mode = 1 Then
            Set bad = New CTestFakeAcadContour: bad.TestInitializeLoop "AcDbLine", "CONCRETE", "BAD", Nothing: space.Add bad
        End If
        If mode = 2 Or mode = 6 Then
            Set cell = New CTestFakeAcadRegion: cell.TestInitialize 100#, 40#, 40#, 1000#, 1000#, 0#, "CONCRETE", "CELL": space.Add cell
        End If
        If mode = 3 Or mode = 6 Then mConfig.Cells(2, 2).Value2 = 0
        If mode = 4 Then mConfig.Cells(4, 2).Value2 = 1.5
        If mode = 5 Then
            Set contours = New CSectionContours: contours.AddContourCircle "H", 40#, 40#, 10#, "", "H", "Opening"
            Set space = TestFakeSpace(contours, bars)
        End If
        settings.LoadFromRange mConfig.Range("A1:C10")
        On Error Resume Next
        Set model = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", IIf(mode = 2, 200#, 0#), units, "OUTER", "OPENING", TestFakeLayers(), settings)
        number = Err.Number: description = Err.Description: Err.Clear
        On Error GoTo 0
        If mode = 6 Then
            TestCheck "existing.noMeshSettingsNeeded", number = 0
            If number = 0 Then TestCheck "existing.notRegenerated", model.ConcreteCount = 1 And model.ConcreteMeshSource = "AutoCADRegions"
        Else
            TestCheck "invalid." & CStr(mode) & "; " & description, number <> 0
            If mode = 3 Then TestCheck "invalid.dynamicAddress", InStr(description, "B2") > 0 And InStr(description, "Mesh.StepX") > 0
        End If
        geometry.BuildContours contours
    Next mode
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет и заново открывает собственную книгу; проверяет восстановление элементов и признака происхождения бетонной сетки.
Private Sub TestCheckSnapshot()
    Dim geometry As CGeometryRegion, settings As CSystemSettingsReader, units As CUnitSystem, bars As CRebarLayout
    Dim mesh As CFiberMeshBuilder, contours As CSectionContours, importer As CAutoCADSectionModelImporter
    Dim model As CSectionModel, restored As CSectionModel, writer As CNDMResultsWriter, sheet As Object, names As Variant, columns As Variant, i As Long
    Set geometry = TestControlGeometry(3): Set settings = TestMeshSettings(): Set units = TestSettingsUnits(settings)
    Set mesh = BuildConfiguredConcreteMesh(geometry, settings, units): Set bars = TestControlBars(mesh, geometry)
    Set contours = New CSectionContours: geometry.BuildContours contours
    Set importer = New CAutoCADSectionModelImporter
    Set model = importer.ImportFromModelSpace(TestFakeSpace(contours, bars), "CONCRETE", "REBAR", "Rebar", 0#, units, "OUTER", "OPENING", TestFakeLayers(), settings)
    Set sheet = mFixture.Worksheets.Add: sheet.Name = "Results"
    sheet.Range("A1").Value2 = "Key": sheet.Range("B1").Value2 = "Value": sheet.Range("C1").Value2 = "Default"
    sheet.Range("A2").Value2 = "Results.SnapshotStorage": sheet.Range("B2").Value2 = "Excel"
    mFixture.Names.Add Name:="rngSystemSettings", RefersTo:="='Results'!" & sheet.Range("A1:C2").Address ' Явный Excel-источник переживает SaveAs/reopen вместе с контрольной геометрией.
    names = Array("rngNDMElementResults", "rngNDMSectionGeometry", "rngNDMSectionContours", "rngNDMSectionProperties", "rngNDMMaterialDiagrams", "rngNDMSectionAnnotations", "rngNDMElementGeometry")
    columns = Array(1, 12, 29, 48, 56, 69, 96)
    For i = 0 To UBound(names)
        mFixture.Names.Add CStr(names(i)), "=Results!" & sheet.Cells(20, CLng(columns(i))).Address
    Next i
    Set writer = New CNDMResultsWriter
    writer.WriteGeometryPreview mFixture, model, PrepareSectionSnapshot(model), units
    Dim path As String
    path = ThisWorkbook.Path & "\ContourMeshSnapshot.xlsm"
    mFixture.SaveAs path, 52: mFixture.Close False
    Set mFixture = Application.Workbooks.Open(path): Set mConfig = mFixture.Worksheets("MeshConfig")
    Set restored = ReadSectionGeometryFromResults(mFixture)
    TestCompareModels "snapshot.reopen", model, restored, False
    TestCheck "snapshot.source", restored.ConcreteMeshSource = "AutoCADContours"
    TestCheck "snapshot.importNotice", InStr(importer.ResultMessage(restored), vbCrLf & vbCrLf & CONTOUR_MESH_GENERATION_NOTICE) > 0
    restored.ConcreteMeshSource = "AutoCADRegions"
    TestCheck "snapshot.noFalseImportNotice", InStr(importer.ResultMessage(restored), CONTOUR_MESH_GENERATION_NOTICE) = 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает фактическую и ожидаемую величины с заданным в этом наборе допуском, сохраняя числа в протоколе.
Private Sub TestNear(ByVal label As String, ByVal a As Double, ByVal b As Double, Optional ByVal tolerance As Double = 0.00000001)
    TestCheck label & "; actual=" & CStr(b) & "; expected=" & CStr(a), Abs(a - b) <= tolerance
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Учитывает отдельное условие проверки и сохраняет его имя в итоговом протоколе.
Private Sub TestCheck(ByVal label As String, ByVal condition As Boolean)
    If condition Then
        mPassed = mPassed + 1
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & label & vbCrLf
    End If
End Sub


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет уровни 1/2/4 по независимому равномерному растру h/2^N:
' каждый конечный пиксель покрыт ровно один раз и только при центре в бетоне.
Public Function RunAdaptiveBoundaryMeshTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    Dim shape As Long, depth As Variant, geometry As CGeometryRegion, query As CSectionGeometryQuery
    Dim region As CConcreteRegion, parts As Collection, mesh As CFiberMeshBuilder
    For shape = 1 To 6
        If shape <= 3 Then
            Set geometry = TestControlGeometry(Choose(shape, 1, 3, 6))
        Else
            Set query = New CSectionGeometryQuery
            Set region = TestFlatPolygon(Array(0#, 0#, 100#, 0#, 100#, 100#, 0#, 100#))
            If shape = 4 Then
                Set region = query.SubtractContainedRegion(region, query.CircleRegion(13#, 17#, 3#))
                Set region = query.SubtractContainedRegion(region, query.CircleRegion(73#, 81#, 7#))
            ElseIf shape = 5 Then
                Set parts = New Collection: parts.Add region
                parts.Add query.CircleRegion(173#, 31#, 3#): Set region = query.MergeDisjointRegions(parts)
            Else
                Set region = TestFlatPolygon(Array(0#, 0#, 100#, 0#, 100#, 100#, 60#, 100#, 60#, 40#, 40#, 40#, 40#, 100#, 0#, 100#))
            End If
            Set geometry = TestFixtureGeometry(region)
        End If
        For Each depth In Array(1, 2, 4)
            Set mesh = New CFiberMeshBuilder: mesh.BuildMesh geometry, 100#, 80#, 7, CLng(depth)
            TestAdaptiveCoverage geometry, mesh, 100#, 80#, CLng(depth), "adaptive." & CStr(shape) & "." & CStr(depth)
        Next depth
    Next shape
    Dim number As Long, reason As String
    On Error Resume Next
    mesh.BuildMesh geometry, 100#, 80#, 1, 2147483647, "Config!TestDepth"
    number = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    TestCheck "adaptive.depthOverflow.beforePower", number <> 0 And InStr(reason, "TestDepth") > 0 And mesh.FiberCount = 0
    GoTo Finish
Failed:
    TestCheck "adaptive.runtime." & CStr(Err.Number) & "; " & Err.Source & "; " & Err.Description, False
Finish:
    RunAdaptiveBoundaryMeshTests = mReport & "TOTAL_ADAPTIVE_BOUNDARY_MESH: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сопоставляет все мелкие клетки с покрытием прямоугольниками адаптивной сетки;
' отдельно сверяет площадь и моменты интегрирования, включая собственную инерцию.
Private Sub TestAdaptiveCoverage(ByVal geometry As CGeometryRegion, ByVal mesh As CFiberMeshBuilder, _
        ByVal stepX As Double, ByVal stepY As Double, ByVal depth As Long, ByVal label As String)
    Dim fineScale As Long, nx As Long, ny As Long, i As Long, ix As Long, iy As Long, firstX As Long, firstY As Long
    Dim width As Long, height As Long, occupied As Object, key As String, overlaps As Boolean, valid As Boolean
    Dim hx As Double, hy As Double, x As Double, y As Double, area As Double, expectedArea As Double
    Dim expectedIx As Double, expectedIy As Double, actualIx As Double, actualIy As Double
    fineScale = CLng(2# ^ depth): hx = stepX / fineScale: hy = stepY / fineScale
    nx = (CLng(Fix(geometry.Width / stepX)) + 1) * fineScale
    ny = (CLng(Fix(geometry.Height / stepY)) + 1) * fineScale
    Set occupied = CreateObject("Scripting.Dictionary"): valid = True
    For i = 1 To mesh.FiberCount
        width = CLng(mesh.FiberWidth(i) / hx): height = CLng(mesh.FiberHeight(i) / hy)
        valid = valid And width >= 1 And height >= 1 And mesh.FiberMaterialID(i) = 7
        valid = valid And Abs(mesh.FiberWidth(i) - width * hx) < 0.0000001 And Abs(mesh.FiberHeight(i) - height * hy) < 0.0000001
        firstX = CLng((mesh.FiberX(i) - mesh.FiberWidth(i) / 2# - geometry.MinX) / hx)
        firstY = CLng((mesh.FiberY(i) - mesh.FiberHeight(i) / 2# - geometry.MinY) / hy)
        For iy = firstY To firstY + height - 1
            For ix = firstX To firstX + width - 1
                key = CStr(ix) & ":" & CStr(iy)
                If occupied.Exists(key) Then overlaps = True Else occupied.Add key, True
            Next ix
        Next iy
        area = mesh.FiberArea(i)
        actualIx = actualIx + area * (mesh.FiberY(i) ^ 2 + mesh.FiberHeight(i) ^ 2 / 12#)
        actualIy = actualIy + area * (mesh.FiberX(i) ^ 2 + mesh.FiberWidth(i) ^ 2 / 12#)
    Next i
    TestCheck label & ".positiveAlignedMaterial", valid
    TestCheck label & ".noOverlap", Not overlaps
    Dim identical As Boolean: identical = True
    For iy = 0 To ny - 1
        y = geometry.MinY + (iy + 0.5) * hy
        For ix = 0 To nx - 1
            x = geometry.MinX + (ix + 0.5) * hx: key = CStr(ix) & ":" & CStr(iy)
            If geometry.ContainsPoint(x, y) <> occupied.Exists(key) Then identical = False
            If geometry.ContainsPoint(x, y) Then
                area = hx * hy: expectedArea = expectedArea + area
                expectedIx = expectedIx + area * (y * y + hy * hy / 12#)
                expectedIy = expectedIy + area * (x * x + hx * hx / 12#)
            End If
        Next ix
    Next iy
    TestCheck label & ".uniformPixelsExact", identical
    TestNear label & ".area", TestAdaptiveMeshArea(mesh), expectedArea, 0.000001
    TestNear label & ".IxOrigin", actualIx, expectedIx, 0.000001 + Abs(expectedIx) * 0.000000000001
    TestNear label & ".IyOrigin", actualIy, expectedIy, 0.000001 + Abs(expectedIy) * 0.000000000001
    mReport = mReport & "ADAPTIVE " & label & "; concrete=" & CStr(mesh.FiberCount) & "; minStepX=" & CStr(hx) & "; area=" & CStr(expectedArea) & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Суммирует фактические площади элементов без округления или восстановления по контуру.
Private Function TestAdaptiveMeshArea(ByVal mesh As CFiberMeshBuilder) As Double
    Dim i As Long
    For i = 1 To mesh.FiberCount
        TestAdaptiveMeshArea = TestAdaptiveMeshArea + mesh.FiberArea(i)
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнительный стенд: настройка не входит в замер. Сетка измеряется пакетами
' по 20 построений, штатное формирование — со схемой, Results и восстановлением Excel.
Public Function TestMeasureAdaptiveMeshBenchmark(ByVal shape As String, ByVal subdivisions As String) As String
    Dim settings As CSystemSettingsReader, units As CUnitSystem, definition As CSectionDefinition, geometry As CGeometryRegion
    Dim mesh As CFiberMeshBuilder, options As TSectionMeshOptions, section As CSectionModel, props As CSectionPropertiesCalculator
    Dim builder As CSectionBuilder, r As Long, repeat As Long, started As Double, seconds As Double, message As String, report As String
    TestAdaptiveSetSetting "Geometry.Source", "Generated"
    TestAdaptiveSetSetting "Geometry.Type", IIf(shape = "OneRect", "Rectangles", shape)
    TestAdaptiveSetSetting "Mesh.StepMode", "Auto 0.1"
    TestAdaptiveSetSetting "Mesh.BoundarySubdivisions", CLng(subdivisions)
    TestAdaptiveSetSetting "General.NonCriticalMessagesEnabled", "No"
    Dim table As Object, row As Long
    Set table = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    For row = 1 To table.Rows.Count
        If CStr(table.Cells(row, 1).Value2) = "Length" Then table.Cells(row, 3).Value2 = "mm"
    Next row
    Set table = ThisWorkbook.Names.Item("rngRectanglesGeometry").RefersToRange
    For row = 1 To table.Rows.Count
        If CStr(table.Cells(row, 1).Value2) = "Shape" Then table.Cells(row, 3).Value2 = IIf(shape = "OneRect", "OneRect", "TwoRect")
    Next row
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
    Set geometry = TestFixtureGeometry(definition.ConcreteDomain): options = ReadSectionMeshOptions(settings, units)
    Set builder = New CSectionBuilder: Set section = builder.Build(definition, options)
    Set props = New CSectionPropertiesCalculator: props.CalculateConcrete section
    Dim analyticAvailable As Boolean
    report = "BENCH_INFO|" & shape & "|" & subdivisions & "|" & CStr(section.ConcreteCount) & "|" & CStr(section.RebarCount) & "|" & TestAdaptiveNumber(geometry.Width) & "|" & TestAdaptiveNumber(geometry.Height) & "|" & TestAdaptiveNumber(0.1 * IIf(geometry.Width < geometry.Height, geometry.Width, geometry.Height)) & "|" & TestAdaptiveNumber(props.Area) & "|" & TestAdaptiveNumber(geometry.AnalyticalArea(analyticAvailable)) & "|" & TestAdaptiveNumber(props.Ixc) & "|" & TestAdaptiveNumber(props.Iyc) & vbCrLf
    For r = 0 To 5
        started = TestAdaptiveClock()
        For repeat = 1 To 20
            Set mesh = New CFiberMeshBuilder
            mesh.BuildMesh geometry, options.StepX, options.StepY, options.MaterialID, options.BoundarySubdivisions, options.SourceReference, options.AutoStepFactor
        Next repeat
        seconds = (TestAdaptiveClock() - started) / 20#
        If r > 0 Then report = report & "BENCH_MESH|" & shape & "|" & subdivisions & "|" & CStr(r) & "|" & TestAdaptiveNumber(seconds) & vbCrLf
    Next r
    For r = 0 To 5
        started = TestAdaptiveClock(): message = FormSectionForWorkbook(ThisWorkbook)
        seconds = TestAdaptiveClock() - started
        If InStr(message, "Предупреждение:") > 0 Then Err.Raise vbObjectError + 5295, "TestMeasureAdaptiveMeshBenchmark", message
        If r > 0 Then report = report & "BENCH_FORM|" & shape & "|" & subdivisions & "|" & CStr(r) & "|" & TestAdaptiveNumber(seconds) & vbCrLf
    Next r
    Dim restored As CSectionModel
    Set restored = ReadSectionGeometryFromResults(ThisWorkbook)
    If restored.ConcreteCount <> section.ConcreteCount Or restored.RebarCount <> section.RebarCount Then Err.Raise vbObjectError + 5295, "TestMeasureAdaptiveMeshBenchmark", "Повторное чтение Results изменило количество элементов."
    TestMeasureAdaptiveMeshBenchmark = report & "TOTAL_ADAPTIVE_BENCHMARK: passed=1; failed=0"
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет только значение найденного ключа в тестовой книге; отсутствие строки является ошибкой стенда.
Private Sub TestAdaptiveSetSetting(ByVal key As String, ByVal value As Variant)
    Dim table As Object, row As Long
    Set table = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    For row = 2 To table.Rows.Count
        If CStr(table.Cells(row, 1).Value2) = key Then table.Cells(row, 2).Value2 = value: Exit Sub
    Next row
    Err.Raise vbObjectError + 5295, "TestAdaptiveSetSetting", "Отсутствует ключ " & key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Монотонный счётчик Windows для коротких участков, где разрешения VBA.Timer недостаточно.
Private Function TestAdaptiveClock() As Double
    Dim counter As Currency, frequency As Currency
    QueryPerformanceCounter counter: QueryPerformanceFrequency frequency
    TestAdaptiveClock = counter / frequency
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Пишет дробные секунды с точкой независимо от разделителя локали Excel.
Private Function TestAdaptiveNumber(ByVal value As Double) As String
    TestAdaptiveNumber = Replace$(Format$(value, "0.000000000"), ",", ".")
End Function
