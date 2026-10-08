Attribute VB_Name = "modTestContourMesh"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: общая сетка Generated/CAD-контуров и реальные решения НДС
' ==========================================================================
' Настройки и Results изменяются только в собственной fixture-книге.
' Сравнения поэлементные; решения используют ядро и диаграммы текущей книги.

Private mPassed As Long, mFailed As Long, mReport As String
Private mFixture As Object, mConfig As Object

Public Function RunContourMeshTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    Set mFixture = Application.Workbooks.Add(-4167)
    Set mConfig = mFixture.Worksheets(1): mConfig.Name = "MeshConfig"
    Dim shape As Long, subdivision As Variant, geometry As CGeometryRegion
    For shape = 1 To 6
        Set geometry = ControlGeometry(shape)
        For Each subdivision In Array(1, 2, 3)
            CompareGeneratedAndImported geometry, CStr(shape) & ".sub" & CStr(subdivision), CLng(subdivision), CLng(subdivision) = 2
        Next subdivision
    Next shape
    CheckArbitraryDomains
    CheckImportContracts
    CheckSnapshot
    GoTo Finish
Failed:
    Check "mesh.runtime." & CStr(Err.Number) & "; " & Err.Source & "; " & Err.Description, False
Finish:
    On Error Resume Next
    If Not mFixture Is Nothing Then mFixture.Close False
    Set mFixture = Nothing: Set mConfig = Nothing
    On Error GoTo 0
    RunContourMeshTests = mReport & "TOTAL_CONTOUR_MESH: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed)
End Function

' ДЛЯ ТЕСТОВ: приложение и новый документ принадлежат внешнему runner-у.
' Контуры выводит production export, импорт читает настоящие CAD-сущности.
Public Function RunNativeContourMeshTests(ByVal acad As Object, ByVal template As String) As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    Dim doc As Object, geometry As CGeometryRegion, settings As CSystemSettingsReader, units As CUnitSystem
    Dim mesh As CFiberMeshBuilder, bars As CRebarLayout, generated As CSectionModel, imported As CSectionModel, restored As CSectionModel
    Dim writer As CNDMResultsWriter, importer As CAutoCADSectionModelImporter, shape As Long, format As Variant
    Dim outerLayer As String, openingLayer As String, rebarLayer As String, meshLayer As String, key As String
    Dim i As Long, exported As Long, countBefore As Long, circleEntity As Object, regions As Variant, curves(0 To 0) As Object
    Dim center(0 To 2) As Double, region As Object, path As String
    On Error GoTo Failed
    Check "native.actualAutoCAD", InStr(1, CStr(acad.FullName), "\AutoCAD 2023\acad.exe", vbTextCompare) > 0
    If InStr(1, CStr(acad.FullName), "\AutoCAD 2023\acad.exe", vbTextCompare) = 0 Then Err.Raise vbObjectError + 5295, , "Требуется Autodesk AutoCAD 2023."
    Set mFixture = Application.Workbooks.Add(-4167): Set mConfig = mFixture.Worksheets(1)
    Set doc = acad.Documents.Add(template)
    Set writer = New CNDMResultsWriter: Set importer = New CAutoCADSectionModelImporter
    For shape = 1 To 6
        For Each format In Array("Region", "Polyline")
            key = "CM_" & CStr(shape) & "_" & CStr(format)
            outerLayer = key & "_OUT": openingLayer = key & "_HOLE": rebarLayer = key & "_BAR": meshLayer = key & "_EMPTY"
            Set geometry = ControlGeometry(shape)
            Set settings = MeshSettings(2, "m", 0.04, 0.03): Set units = SettingsUnits(settings)
            Set mesh = BuildConfiguredConcreteMesh(geometry, settings, units): Set bars = ControlBars(mesh, geometry)
            Set generated = BuildGeneratedSectionModel(mesh, bars): geometry.BuildContours generated.Contours
            writer.WriteGeometryPreview ThisWorkbook, generated, PrepareSectionSnapshot(generated), units
            exported = SP35ExportSavedMaterialContoursForTests(doc, outerLayer, openingLayer, CStr(format))
            Check key & ".exportedContours", exported > 0
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
            Check key & ".sourcePreserved", doc.ModelSpace.Count = countBefore
            Check key & ".source", imported.ConcreteMeshSource = "AutoCADContours"
            Check key & ".notice", InStr(importer.ResultMessage(imported), CONTOUR_MESH_GENERATION_NOTICE) > 0
            If shape < 5 Then CompareModels key, generated, imported, False Else CompareCurvedModels key, generated, imported
            If shape = 3 Then CompareSolutions key, generated, imported, True
            If shape = 3 And CStr(format) = "Region" Then
                CheckBatchPair generated, imported
                CheckNativeFullExport doc, imported
            End If
            writer.WriteGeometryPreview ThisWorkbook, imported, PrepareSectionSnapshot(imported), units
            Set restored = ReadSectionGeometryFromResults(ThisWorkbook)
            CompareModels key & ".snapshot", imported, restored, False
            Check key & ".restoredNotice", InStr(WithConcreteMeshGenerationNotice("Export", restored), vbCrLf & vbCrLf & CONTOUR_MESH_GENERATION_NOTICE) > 0
            exported = SP35ExportSavedMaterialContoursForTests(doc, key & "_OUT2", key & "_HOLE2", CStr(format))
            Check key & ".secondExport", exported > 0
            Dim reimported As CSectionModel
            Set reimported = importer.ImportFromModelSpace(doc.ModelSpace, meshLayer, rebarLayer, "Rebar", 0#, units, key & "_OUT2", key & "_HOLE2", doc.Layers, settings)
            CompareModels key & ".CADRoundTrip", imported, reimported, False
            mReport = mReport & "NATIVE " & key & ": concrete=" & CStr(imported.ConcreteCount) & "; contours=" & CStr(imported.Contours.Count) & "; rebar=" & CStr(imported.RebarCount) & vbCrLf
        Next format
    Next shape
    CheckNativeTwoHoles doc, settings, units
    path = ThisWorkbook.Path & "\ContourMeshNative.dwg": doc.SaveAs path
    mReport = mReport & "NATIVE_DWG: " & path & vbCrLf
    GoTo Finish
Failed:
    Check "native.mesh.runtime." & CStr(Err.Number) & "; " & Err.Source & "; " & Err.Description, False
Finish:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    If Not mFixture Is Nothing Then mFixture.Close False
    Set mFixture = Nothing: Set mConfig = Nothing
    On Error GoTo 0
    RunNativeContourMeshTests = mReport & "TOTAL_NATIVE_CONTOUR_MESH: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed)
End Function

' Проверяет настоящий экспорт волокон и напряжений из полного Results:
' миллиметры, количество/площадь Region и отсутствие нового решения НДС.
Private Sub CheckNativeFullExport(ByVal doc As Object, ByVal model As CSectionModel)
    Dim table As Object, profiles As Object, saved As Variant, savedProfiles As Variant
    Dim i As Long, settings As CSystemSettingsReader, concreteLayer As String, rebarLayer As String
    Dim entity As Object, count As Long, bars As Long, area As Double, expectedArea As Double, exported As Long, solves As Long
    On Error GoTo Failed
    Set table = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange: saved = table.Formula
    Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange: savedProfiles = profiles.Formula
    For i = 2 To table.Rows.Count
        If CStr(table.Cells(i, 1).Value2) = "AutoCAD.Export.CombinationID" Then table.Cells(i, 2).Value2 = "CM_BEND"
    Next i
    For i = 2 To profiles.Rows.Count
        If CStr(profiles.Cells(i, 2).Value2) = "Visualization.State" Then profiles.Cells(i, 3).Value2 = "StrengthState"
    Next i
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    concreteLayer = settings.GetRequiredString("AutoCAD.Common.ConcreteLayer")
    rebarLayer = settings.GetRequiredString("AutoCAD.Common.RebarLayer")
    solves = SectionEquilibriumSolveCount()
    exported = Audit03ExportAutoCADDocumentForTests(doc)
    Check "native.fullExport.objects", exported > model.ConcreteCount + model.RebarCount
    Check "native.fullExport.noSolve", solves = SectionEquilibriumSolveCount()
    For Each entity In doc.ModelSpace
        If CStr(entity.ObjectName) = "AcDbRegion" Then
            If StrComp(CStr(entity.Layer), concreteLayer, vbTextCompare) = 0 Then
                count = count + 1: area = area + CDbl(entity.Area)
            End If
            If StrComp(CStr(entity.Layer), rebarLayer, vbTextCompare) = 0 Then bars = bars + 1
        End If
    Next entity
    For i = 1 To model.ConcreteCount: expectedArea = expectedArea + model.ConcreteArea(i): Next i
    Check "native.fullExport.count", count = model.ConcreteCount And bars = model.RebarCount
    Near "native.fullExport.areaInMm2", expectedArea, area, 0.0000001
    Dim restored As CSectionModel
    Set restored = ReadSectionGeometryFromResults(ThisWorkbook)
    Check "native.fullExport.notice", InStr(WithConcreteMeshGenerationNotice("Экспорт завершен.", restored), vbCrLf & vbCrLf & CONTOUR_MESH_GENERATION_NOTICE) > 0
    GoTo Finish
Failed:
    Check "native.fullExport.runtime." & CStr(Err.Number) & "; " & Err.Description, False
Finish:
    If Not IsEmpty(saved) Then table.Formula = saved
    If Not IsEmpty(savedProfiles) Then profiles.Formula = savedProfiles
End Sub

' Единый native Region содержит сразу два отверстия; отдельный opening-слой пуст.
Private Sub CheckNativeTwoHoles(ByVal doc As Object, ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
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
    Set generated = BuildGeneratedSectionModel(mesh, bars): CompareModels "native.twoHoles", generated, model, False
    Set query = New CSectionGeometryQuery: query.Initialize model
    Near "native.twoHoles.area", query.ConcreteDomain.Area, GEOM_PI * (120# * 120# - 23# * 23# - 31# * 31#), 0.0000001
    Check "native.twoHoles.empty", Not query.ContainsPoint(query.ConcreteDomain, -45#, 0#) And Not query.ContainsPoint(query.ConcreteDomain, 45#, 0#)
End Sub

Private Function ControlGeometry(ByVal shape As Long) As CGeometryRegion
    Dim circleGeometry As CGeometryCircle, rect As CGeometryRoundedRectangle
    Dim hollow As CGeometryHollowRectangle, rectset As CGeometryRectSet
    Select Case shape
        Case 1
            Set circleGeometry = New CGeometryCircle: circleGeometry.InitializeByDiameter 300#
            Set ControlGeometry = BuildConcreteGeometry(circleGeometry)
        Case 2, 3, 4
            Set rectset = New CGeometryRectSet
            If shape = 2 Then rectset.Initialize 300#, 400#, 0#, 0#, 0#, 0#, 0#, "Rectangle"
            If shape = 3 Then rectset.Initialize 120#, 240#, 300#, 120#, 0#, 0#, 0#, "LSection"
            If shape = 4 Then rectset.Initialize 200#, 240#, 300#, 120#, -83#, 41#, 60#, "TwoRectangles"
            Set ControlGeometry = BuildConcreteGeometry(rectset)
        Case 5
            Set rect = New CGeometryRoundedRectangle: rect.Initialize 300#, 400#, 40#, 40#, 40#, 40#, 13#, -27#
            Set ControlGeometry = BuildConcreteGeometry(rect)
        Case 6
            Set hollow = New CGeometryHollowRectangle: hollow.Initialize 400#, 500#, 35#, 200#, 280#, 25#, 17#, -23#
            Set ControlGeometry = BuildConcreteGeometry(hollow)
    End Select
End Function

Private Function MeshSettings(Optional ByVal subdivision As Long = 2, Optional ByVal unit As String = "mm", _
        Optional ByVal stepX As Double = 40#, Optional ByVal stepY As Double = 30#) As CSystemSettingsReader
    mConfig.Cells.ClearContents
    Dim data(1 To 7, 1 To 3) As Variant
    data(1, 1) = "Key": data(1, 2) = "Value": data(1, 3) = "Unit"
    data(2, 1) = "Mesh.StepX": data(2, 2) = stepX
    data(3, 1) = "Mesh.StepY": data(3, 2) = stepY
    data(4, 1) = "Mesh.BoundarySubdivisions": data(4, 2) = subdivision
    data(5, 1) = "Units.Length.Input": data(5, 2) = unit
    data(6, 1) = "Units.Length.Output": data(6, 2) = unit
    data(7, 1) = "Units.Area.Output": data(7, 2) = unit & "2"
    mConfig.Range("A1:C7").Value2 = data
    Set MeshSettings = New CSystemSettingsReader
    MeshSettings.LoadFromRange mConfig.Range("A1:C7")
End Function

Private Function SettingsUnits(ByVal settings As CSystemSettingsReader) As CUnitSystem
    Set SettingsUnits = New CUnitSystem: SettingsUnits.LoadFromSettings settings
End Function

Private Function FakeLayers() As Collection
    Set FakeLayers = New Collection
    Dim name As Variant, layer As CFakeAcadContour
    For Each name In Array("CONCRETE", "REBAR", "OUTER", "OPENING")
        Set layer = New CFakeAcadContour
        layer.InitializeLoop "Layer", CStr(name), CStr(name), Nothing
        FakeLayers.Add layer, CStr(name)
    Next name
End Function

Private Function FakeSpace(ByVal contours As CSectionContours, ByVal bars As CRebarLayout) As Collection
    Dim loops As Object, edges As Collection, i As Long, key As Variant, segment As CFakeAcadContour
    Dim entity As CFakeAcadContour, rebar As CFakeAcadRegion, angle As Double, layer As String
    Set FakeSpace = New Collection: Set loops = CreateObject("Scripting.Dictionary")
    If Not bars Is Nothing Then
        For i = 1 To bars.Count
            Set rebar = New CFakeAcadRegion
            rebar.Initialize bars.Area(i), bars.X(i), bars.Y(i), 1#, 1#, 0#, "REBAR", "BAR" & CStr(i)
            FakeSpace.Add rebar
        Next i
    End If
    For i = 1 To contours.Count
        layer = "OUTER": If contours.LoopRole(i) = "Opening" Then layer = "OPENING"
        key = layer & "|" & contours.LoopID(i)
        If Not loops.Exists(key) Then
            Set edges = New Collection: loops.Add key, edges
        End If
        Set edges = loops(key): Set segment = New CFakeAcadContour
        Select Case contours.SegmentType(i)
            Case "CONTOUR_LINE"
                segment.InitializeLine contours.StartX(i), contours.StartY(i), contours.EndX(i), contours.EndY(i)
            Case "CONTOUR_ARC"
                angle = Atan2(contours.StartY(i) - contours.CenterY(i), contours.StartX(i) - contours.CenterX(i))
                segment.InitializeArc contours.CenterX(i), contours.CenterY(i), contours.Radius(i), angle, contours.SweepAngle(i)
            Case "CONTOUR_CIRCLE"
                segment.InitializeCircle contours.StartX(i), contours.StartY(i), contours.Radius(i)
        End Select
        edges.Add segment
    Next i
    For Each key In loops.Keys
        Set entity = New CFakeAcadContour
        entity.InitializeLoop "AcDbRegion", Left$(CStr(key), InStr(CStr(key), "|") - 1), CStr(key), loops(key)
        FakeSpace.Add entity
    Next key
End Function

Private Function Atan2(ByVal y As Double, ByVal x As Double) As Double
    If x = 0# Then
        Atan2 = Sgn(y) * GEOM_PI / 2#
    Else
        Atan2 = Atn(y / x)
        If x < 0# Then Atan2 = Atan2 + GEOM_PI
    End If
End Function

Private Function ControlBars(ByVal mesh As CFiberMeshBuilder, ByVal geometry As CGeometryRegion) As CRebarLayout
    Set ControlBars = New CRebarLayout
    Dim index As Long, i As Long
    For i = 1 To 4
        index = 1 + CLng(Fix(CDbl(mesh.FiberCount - 1) * CDbl(i) / 5#))
        ControlBars.AddBar "SOURCE" & CStr(i), mesh.FiberX(index), mesh.FiberY(index), 16#, 0#, "Rebar", "", geometry
    Next i
End Function

Private Sub CompareGeneratedAndImported(ByVal geometry As CGeometryRegion, ByVal label As String, _
        ByVal subdivision As Long, ByVal solve As Boolean)
    Dim settings As CSystemSettingsReader, units As CUnitSystem, mesh As CFiberMeshBuilder, bars As CRebarLayout
    Dim generated As CSectionModel, imported As CSectionModel, importer As CAutoCADSectionModelImporter
    Dim query As CSectionGeometryQuery, adapter As CGeometryRegion, available As Boolean
    Set settings = MeshSettings(subdivision): Set units = SettingsUnits(settings)
    Set mesh = BuildConfiguredConcreteMesh(geometry, settings, units): Set bars = ControlBars(mesh, geometry)
    Set generated = BuildGeneratedSectionModel(mesh, bars)
    geometry.BuildContours generated.Contours
    Set importer = New CAutoCADSectionModelImporter
    Set imported = importer.ImportFromModelSpace(FakeSpace(generated.Contours, bars), "CONCRETE", "REBAR", "Rebar", 0#, units, "OUTER", "OPENING", FakeLayers(), settings)
    CompareModels label, generated, imported
    Check label & ".source", imported.SourceType = "AutoCADImport" And imported.ConcreteMeshSource = "AutoCADContours"
    Check label & ".notice", InStr(importer.ResultMessage(imported), vbCrLf & vbCrLf & CONTOUR_MESH_GENERATION_NOTICE) > 0
    Set query = New CSectionGeometryQuery: query.Initialize imported
    Set adapter = New CGeometryRegion: adapter.Initialize query.ConcreteDomain
    Near label & ".exactArea", adapter.AnalyticalArea(available), geometry.AnalyticalArea(available), 0.0000001
    CheckCellCenters label, mesh, geometry
    Dim i As Long
    For i = 1 To imported.ConcreteCount
        Check label & ".importedCenter." & CStr(i), adapter.ContainsPoint(imported.ConcreteX(i), imported.ConcreteY(i))
    Next i
    If solve Then CompareSolutions label, generated, imported, True
    If label = "3.sub2" Then CheckBatchPair generated, imported
    mReport = mReport & "SECTION " & label & ": concrete=" & CStr(imported.ConcreteCount) & "; rebar=" & CStr(imported.RebarCount) & vbCrLf
End Sub

' Полный расчет двух одинаковых Г-сечений: direct, capacity и трещины СП 35.
' Настройки восстанавливаются даже при ошибке; исходная книга теста - копия.
Private Sub CheckBatchPair(ByVal first As CSectionModel, ByVal second As CSectionModel)
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
    Set units = SettingsUnits(settings)
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
        Check "batch.status." & CStr(i), ra.Status = rb.Status
        Check "batch.widthStatus." & CStr(i), ra.NormalCrackStatus = rb.NormalCrackStatus
        Check "batch.noInputErr." & CStr(i) & "; " & rb.OverallMeta.ResultComment, rb.Status <> "InputErr"
        Check "batch.capacityCalculated." & CStr(i), rb.StrengthResult.Capacity.ResultMeta.Calculated
        Near "batch.capacity." & CStr(i), ra.StrengthResult.Capacity.LambdaCapacity, rb.StrengthResult.Capacity.LambdaCapacity, 0.0000000001
        Near "batch.formation." & CStr(i), ra.CrackResult.Formation.LambdaCrc, rb.CrackResult.Formation.LambdaCrc, 0.0000000001
        Near "batch.width." & CStr(i), ra.CrackResult.Width.CrackWidth, rb.CrackResult.Width.CrackWidth, 0.0000000001
        Near "batch.Ar." & CStr(i), ra.CrackResult.Width.Abt, rb.CrackResult.Width.Abt, 0.00000001
        mReport = mReport & "BATCH " & CStr(i) & ": status=" & rb.Status & "; normalCrack=" & rb.NormalCrackStatus & _
            "; capacity=" & CStr(rb.StrengthResult.Capacity.LambdaCapacity) & "; width=" & CStr(rb.CrackResult.Width.CrackWidth) & "; Ar=" & CStr(rb.CrackResult.Width.Abt) & _
            "; reason=" & rb.OverallMeta.ResultComment & vbCrLf
    Next i
    Check "batch.widthActuallyCalculated", b.ResultAt(2).CrackResult.Width.ResultMeta.Calculated
    Check "batch.SP35", b.ResultAt(2).CrackResult.Width.StandardCode = "SP35"
    Near "batch.actualRegionArea", b.ResultAt(2).CrackResult.Width.Abt, b.ResultAt(2).CrackResult.Width.InteractionRegion.Area, 0.00000001
    Set writer = New CNDMResultsWriter
    writer.WriteResults ThisWorkbook, second, PrepareSectionSnapshot(second, provider), b, units
    Dim restored As CSectionModel
    Set restored = ReadSectionGeometryFromResults(ThisWorkbook)
    Check "batch.fullWriterSource", restored.ConcreteMeshSource = "AutoCADContours"
    GoTo Finish
Failed:
    Check "batch.runtime." & CStr(Err.Number) & "; " & Err.Description, False
Finish:
    If Not IsEmpty(savedProfiles) Then profileTable.Formula = savedProfiles
    If Not IsEmpty(savedSystem) Then systemTable.Formula = savedSystem
End Sub

' Независимый контроль площади и инерции после нативного CAD-преобразования;
' величины отклонений сохраняются в отчете.
Private Sub CompareCurvedModels(ByVal label As String, ByVal first As CSectionModel, ByVal second As CSectionModel)
    Dim a As CSectionPropertiesCalculator, b As CSectionPropertiesCalculator
    Set a = New CSectionPropertiesCalculator: a.CalculateConcrete first
    Set b = New CSectionPropertiesCalculator: b.CalculateConcrete second
    Near label & ".curvedArea", a.Area, b.Area, a.Area * 0.01
    Near label & ".curvedIx", a.Ixc, b.Ixc, a.Ixc * 0.01
    Near label & ".curvedIy", a.Iyc, b.Iyc, a.Iyc * 0.01
    Check label & ".curvedBars", first.RebarCount = second.RebarCount
    mReport = mReport & "CURVED " & label & ": GeneratedCount=" & CStr(first.ConcreteCount) & "; CADCount=" & CStr(second.ConcreteCount) & _
        "; relativeArea=" & CStr(Abs(a.Area - b.Area) / a.Area) & "; relativeIx=" & CStr(Abs(a.Ixc - b.Ixc) / a.Ixc) & "; relativeIy=" & CStr(Abs(a.Iyc - b.Iyc) / a.Iyc) & vbCrLf
End Sub

Private Sub CompareModels(ByVal label As String, ByVal first As CSectionModel, ByVal second As CSectionModel, Optional ByVal checkHandles As Boolean = True)
    Check label & ".counts", first.ConcreteCount = second.ConcreteCount And first.RebarCount = second.RebarCount
    If first.ConcreteCount <> second.ConcreteCount Or first.RebarCount <> second.RebarCount Then Exit Sub
    Dim i As Long
    For i = 1 To first.ConcreteCount
        Near label & ".x." & CStr(i), first.ConcreteX(i), second.ConcreteX(i)
        Near label & ".y." & CStr(i), first.ConcreteY(i), second.ConcreteY(i)
        Near label & ".a." & CStr(i), first.ConcreteArea(i), second.ConcreteArea(i)
        Near label & ".width." & CStr(i), first.ConcreteWidth(i), second.ConcreteWidth(i)
        Near label & ".height." & CStr(i), first.ConcreteHeight(i), second.ConcreteHeight(i)
        Near label & ".Ix." & CStr(i), first.ConcreteLocalIx(i), second.ConcreteLocalIx(i)
        Near label & ".Iy." & CStr(i), first.ConcreteLocalIy(i), second.ConcreteLocalIy(i)
        Check label & ".id." & CStr(i), first.ConcreteID(i) = second.ConcreteID(i)
    Next i
    For i = 1 To first.RebarCount
        Near label & ".barX." & CStr(i), first.RebarX(i), second.RebarX(i)
        Near label & ".barY." & CStr(i), first.RebarY(i), second.RebarY(i)
        Near label & ".barA." & CStr(i), first.RebarArea(i), second.RebarArea(i)
        If checkHandles Then Check label & ".barHandle." & CStr(i), second.RebarSourceHandle(i) = "BAR" & CStr(i)
    Next i
End Sub

' 72 пары: шесть форм, четыре сочетания, три состояния материала.
Private Sub CompareSolutions(ByVal label As String, ByVal first As CSectionModel, ByVal second As CSectionModel, ByVal exactPair As Boolean)
    Dim settings As CSystemSettingsReader, units As CUnitSystem, provider As CMaterialModelProvider
    Dim a As CSectionSolver, b As CSectionSolver, props As CSectionPropertiesCalculator
    Dim purpose As Variant, loadIndex As Long, n As Double, mx As Double, my As Double, height As Double, width As Double
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double, prefix As String
    Dim relativeTolerance As Double
    If Not exactPair Then relativeTolerance = 0.01
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
            Check prefix & ".converged; " & a.StopReason & "; " & b.StopReason, a.Converged And b.Converged
            Near prefix & ".eps", a.Epsilon0, b.Epsilon0, 0.000000000001 + Abs(a.Epsilon0) * relativeTolerance
            Near prefix & ".kx", a.KappaX, b.KappaX, 0.000000000000001 + Abs(a.KappaX) * relativeTolerance
            Near prefix & ".ky", a.KappaY, b.KappaY, 0.000000000000001 + Abs(a.KappaY) * relativeTolerance
            Near prefix & ".N", b.Nint, n, 2#
            Near prefix & ".Mx", b.Mxint, mx, 2000#
            Near prefix & ".My", b.Myint, my, 2000#
            Near prefix & ".minConcrete", a.MinConcreteStress, b.MinConcreteStress, 0.00000001 + Abs(a.MinConcreteStress) * relativeTolerance
            Near prefix & ".maxConcrete", a.MaxConcreteStress, b.MaxConcreteStress, 0.00000001 + GeomMax(Abs(a.MaxConcreteStress), Abs(a.MinConcreteStress)) * relativeTolerance
            Near prefix & ".minSteel", a.MinSteelStress, b.MinSteelStress, 0.00000001 + Abs(a.MinSteelStress) * relativeTolerance
            Near prefix & ".maxSteel", a.MaxSteelStress, b.MaxSteelStress, 0.00000001 + Abs(a.MaxSteelStress) * relativeTolerance
            mReport = mReport & "SOLVE " & prefix & ": eps=" & CStr(b.Epsilon0) & "; kx=" & CStr(b.KappaX) & "; ky=" & CStr(b.KappaY) & "; iterations=" & CStr(b.Iterations) & vbCrLf
        Next loadIndex
    Next purpose
End Sub

Private Sub CheckArbitraryDomains()
    Dim query As CSectionGeometryQuery, region As CConcreteRegion, geometry As CGeometryRegion, mesh As CFiberMeshBuilder
    Set query = New CSectionGeometryQuery
    Set region = FlatPolygon(Array(0#, 0#, 100#, 0#, 100#, 100#, 60#, 100#, 60#, 40#, 40#, 40#, 40#, 100#, 0#, 100#))
    Set geometry = New CGeometryRegion: geometry.Initialize region
    Set mesh = New CFiberMeshBuilder: mesh.BuildMesh geometry, 100#, 100#, 1, 10
    Check "concave.subcells", mesh.FiberCount = 88
    CheckCellCenters "concave", mesh, geometry
    Set region = FlatPolygon(Array(0#, 0#, 300#, 0#, 300#, 200#, 0#, 200#))
    Set region = query.SubtractContainedRegion(region, query.CircleRegion(70#, 70#, 23#))
    Set region = query.SubtractContainedRegion(region, query.CircleRegion(230#, 130#, 31#))
    geometry.Initialize region: mesh.BuildMesh geometry, 40#, 30#, 1, 3
    CheckCellCenters "twoHoles", mesh, geometry
    Check "twoHoles.exclude", Not geometry.ContainsPoint(70#, 70#) And Not geometry.ContainsPoint(230#, 130#)
    Check "twoHoles.outerBoundary", geometry.ContainsPoint(0#, 0#)
    Check "twoHoles.openingBoundary", Not geometry.ContainsPoint(93#, 70#)
    Dim parts As Collection
    Set parts = New Collection: parts.Add region
    parts.Add FlatPolygon(Array(400#, 0#, 500#, 0#, 500#, 100#, 400#, 100#))
    Set region = query.MergeDisjointRegions(parts): geometry.Initialize region
    mesh.BuildMesh geometry, 40#, 30#, 1, 3
    CheckCellCenters "disjoint", mesh, geometry
    Check "disjoint.material", geometry.ContainsPoint(450#, 50#) And Not geometry.ContainsPoint(350#, 50#)
    Dim pointX() As Double, pointY() As Double, i As Long, message As String, number As Long
    geometry.GetExtremePoints pointX, pointY
    For i = 1 To UBound(pointX)
        Check "extrema.realBoundary." & CStr(i), query.ContainsPoint(region, pointX(i), pointY(i), True)
    Next i
    On Error Resume Next
    geometry.Initialize Nothing: number = Err.Number: Err.Clear
    On Error GoTo 0
    Check "adapter.invalidInitializeClearsOldShape", number <> 0 And Not geometry.IsValid(message)
End Sub

Private Function FlatPolygon(ByVal values As Variant) As CConcreteRegion
    Dim points() As Double, i As Long, query As CSectionGeometryQuery
    ReDim points(1 To (UBound(values) + 1) \ 2, 1 To 2)
    For i = 1 To UBound(points, 1)
        points(i, 1) = values((i - 1) * 2): points(i, 2) = values((i - 1) * 2 + 1)
    Next i
    Set query = New CSectionGeometryQuery: Set FlatPolygon = query.PolygonRegion(points)
End Function

Private Sub CheckCellCenters(ByVal label As String, ByVal mesh As CFiberMeshBuilder, ByVal geometry As CGeometryRegion)
    Dim i As Long, seen As Object, key As String
    Set seen = CreateObject("Scripting.Dictionary")
    For i = 1 To mesh.FiberCount
        Check label & ".center." & CStr(i), geometry.ContainsPoint(mesh.FiberX(i), mesh.FiberY(i))
        Check label & ".positive." & CStr(i), mesh.FiberArea(i) > 0# And mesh.FiberWidth(i) > 0# And mesh.FiberHeight(i) > 0#
        key = CStr(mesh.FiberX(i)) & "|" & CStr(mesh.FiberY(i))
        Check label & ".unique." & CStr(i), Not seen.Exists(key): seen(key) = True
    Next i
End Sub

Private Sub CheckImportContracts()
    Dim geometry As CGeometryRegion, contours As CSectionContours, bars As CRebarLayout, mesh As CFiberMeshBuilder
    Dim importer As CAutoCADSectionModelImporter, settings As CSystemSettingsReader, units As CUnitSystem
    Dim space As Collection, model As CSectionModel, baseline As CSectionModel, cell As CFakeAcadRegion, bad As CFakeAcadContour
    Dim number As Long, description As String, mode As Long, factor As Double, unit As Variant
    Set geometry = ControlGeometry(3): Set contours = New CSectionContours: geometry.BuildContours contours
    Set settings = MeshSettings(): Set units = SettingsUnits(settings)
    Set mesh = BuildConfiguredConcreteMesh(geometry, settings, units): Set bars = ControlBars(mesh, geometry)
    Set importer = New CAutoCADSectionModelImporter: Set space = FakeSpace(contours, bars)
    Set baseline = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, units, "OUTER", "OPENING", FakeLayers(), settings)
    For Each unit In Array("cm", "m")
        factor = 10#: If CStr(unit) = "m" Then factor = 1000#
        Set settings = MeshSettings(2, CStr(unit), 40# / factor, 30# / factor): Set units = SettingsUnits(settings)
        Set model = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", 0#, units, "OUTER", "OPENING", FakeLayers(), settings)
        CompareModels "units." & CStr(unit), baseline, model
    Next unit
    For mode = 0 To 6
        Set settings = MeshSettings(): Set units = SettingsUnits(settings)
        Set space = FakeSpace(contours, bars)
        If mode = 0 Then Set space = FakeSpace(New CSectionContours, bars)
        If mode = 1 Then
            Set bad = New CFakeAcadContour: bad.InitializeLoop "AcDbLine", "CONCRETE", "BAD", Nothing: space.Add bad
        End If
        If mode = 2 Or mode = 6 Then
            Set cell = New CFakeAcadRegion: cell.Initialize 100#, 40#, 40#, 1000#, 1000#, 0#, "CONCRETE", "CELL": space.Add cell
        End If
        If mode = 3 Or mode = 6 Then mConfig.Cells(2, 2).Value2 = 0
        If mode = 4 Then mConfig.Cells(4, 2).Value2 = 1.5
        If mode = 5 Then
            Set contours = New CSectionContours: contours.AddContourCircle "H", 40#, 40#, 10#, "", "H", "Opening"
            Set space = FakeSpace(contours, bars)
        End If
        settings.LoadFromRange mConfig.Range("A1:C7")
        On Error Resume Next
        Set model = importer.ImportFromModelSpace(space, "CONCRETE", "REBAR", "Rebar", IIf(mode = 2, 200#, 0#), units, "OUTER", "OPENING", FakeLayers(), settings)
        number = Err.Number: description = Err.Description: Err.Clear
        On Error GoTo 0
        If mode = 6 Then
            Check "existing.noMeshSettingsNeeded", number = 0
            If number = 0 Then Check "existing.notRegenerated", model.ConcreteCount = 1 And model.ConcreteMeshSource = "AutoCADRegions"
        Else
            Check "invalid." & CStr(mode) & "; " & description, number <> 0
            If mode = 3 Then Check "invalid.dynamicAddress", InStr(description, "B2") > 0 And InStr(description, "Mesh.StepX") > 0
        End If
        geometry.BuildContours contours
    Next mode
End Sub

Private Sub CheckSnapshot()
    Dim geometry As CGeometryRegion, settings As CSystemSettingsReader, units As CUnitSystem, bars As CRebarLayout
    Dim mesh As CFiberMeshBuilder, contours As CSectionContours, importer As CAutoCADSectionModelImporter
    Dim model As CSectionModel, restored As CSectionModel, writer As CNDMResultsWriter, sheet As Object, names As Variant, columns As Variant, i As Long
    Set geometry = ControlGeometry(3): Set settings = MeshSettings(): Set units = SettingsUnits(settings)
    Set mesh = BuildConfiguredConcreteMesh(geometry, settings, units): Set bars = ControlBars(mesh, geometry)
    Set contours = New CSectionContours: geometry.BuildContours contours
    Set importer = New CAutoCADSectionModelImporter
    Set model = importer.ImportFromModelSpace(FakeSpace(contours, bars), "CONCRETE", "REBAR", "Rebar", 0#, units, "OUTER", "OPENING", FakeLayers(), settings)
    Set sheet = mFixture.Worksheets.Add: sheet.Name = "Results"
    names = Array("rngNDMElementResults", "rngNDMSectionGeometry", "rngNDMSectionContours", "rngNDMSectionProperties", "rngNDMMaterialDiagrams", "rngNDMSectionAnnotations")
    columns = Array(1, 12, 29, 48, 56, 69)
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
    CompareModels "snapshot.reopen", model, restored, False
    Check "snapshot.source", restored.ConcreteMeshSource = "AutoCADContours"
    Check "snapshot.exportNotice", WithConcreteMeshGenerationNotice("Экспорт завершен.", restored) = "Экспорт завершен." & vbCrLf & vbCrLf & CONTOUR_MESH_GENERATION_NOTICE
    restored.ConcreteMeshSource = "AutoCADRegions"
    Check "snapshot.noFalseNotice", WithConcreteMeshGenerationNotice("Экспорт завершен.", restored) = "Экспорт завершен."
End Sub

Private Sub Near(ByVal label As String, ByVal a As Double, ByVal b As Double, Optional ByVal tolerance As Double = 0.00000001)
    Check label & "; actual=" & CStr(b) & "; expected=" & CStr(a), Abs(a - b) <= tolerance
End Sub

Private Sub Check(ByVal label As String, ByVal condition As Boolean)
    If condition Then
        mPassed = mPassed + 1
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & label & vbCrLf
    End If
End Sub
