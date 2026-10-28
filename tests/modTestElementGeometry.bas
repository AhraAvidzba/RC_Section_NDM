Attribute VB_Name = "modTestElementGeometry"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: точные области бетонных элементов с прямыми и дугами
' ==========================================================================
' Аналитические эталоны проверяют площадь, центр, собственные моменты и
' разбиение на связные части без Excel, AutoCAD и расчёта напряжений.

Private Type TTestStats
    Passed As Long
    Failed As Long
    Report As String
End Type

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет общую геометрию независимо от режима сетки и источника контура.
Public Function RunElementGeometryTests() As String
    On Error GoTo Failed
    Dim stats As TTestStats
    TestExactRectangle stats
    TestExactCircle stats
    TestExactOpening stats
    TestExactComponents stats
    TestClippedCellClosure stats
    TestTangentArcEndpoints stats
    TestHollowTangentCells stats
    TestSmallArcSegment stats
    TestOwnerReadQueries stats
    stats.Report = stats.Report & "TOTAL: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & vbCrLf
    RunElementGeometryTests = stats.Report
    Exit Function
Failed:
    RunElementGeometryTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; source=" & Err.Source & "; description=" & Err.Description & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет запрос владельца с прежним запросом независимой Geometry для прямоугольника, дуг и отверстия; выдаваемая наружу область остаётся копией.
Private Sub TestOwnerReadQueries(ByRef stats As TTestStats)
    Dim query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    Dim regions As Collection: Set regions = New Collection
    regions.Add query.RectangleRegion(17#, -9#, 50#, 100#, 0.37)
    regions.Add query.CircleRegion(-11#, 7#, 35#)
    regions.Add query.SubtractContainedRegion(query.CircleRegion(0#, 0#, 100#), query.CircleRegion(20#, 0#, 15#))
    Dim item As Variant, region As CConcreteRegion, model As CSectionModel, properties As TGeometryProperties
    Dim axis As Variant, minimum As Double, maximum As Double, oldMinimum As Double, oldMaximum As Double
    Dim cx As Double, cy As Double, width As Double, height As Double, rotation As Double, number As Long
    For Each item In regions
        Set region = item: query.RegionProperties region, properties
        Set model = New CSectionModel
        model.AddConcreteElement properties.CentroidX, properties.CentroidY, properties.Area, shapeType:="Contour", _
            localIx:=properties.Ixc, localIy:=properties.Iyc, localIxy:=properties.Ixyc, exactRegion:=region
        number = number + 1
        For Each axis In Array(Array(1#, 0#), Array(0#, 1#), Array(0.6, 0.8), Array(-0.8, 0.6))
            model.ConcreteProjectionBounds 1, CDbl(axis(0)), CDbl(axis(1)), minimum, maximum
            query.ProjectionBounds model.ConcreteGeometry(1), CDbl(axis(0)), CDbl(axis(1)), oldMinimum, oldMaximum
            TestTrue stats, "ownerQuery.min." & CStr(number), minimum = oldMinimum
            TestTrue stats, "ownerQuery.max." & CStr(number), maximum = oldMaximum
        Next axis
        TestTrue stats, "ownerQuery.rectangle." & CStr(number), _
            model.ConcreteBoundaryRectangle(1, width, height, rotation) = query.TryRectangle(model.ConcreteGeometry(1), cx, cy, width, height, rotation)
    Next item
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет полный перенос точных ячеек через CFiberMeshBuilder и сборщик модели,
' а также согласованность компактного Adaptive-вызова при двух и четырёх уровнях.
Public Function RunContourCellMeshTests() As String
    On Error GoTo Failed
    Dim stats As TTestStats, query As CSectionGeometryQuery, region As CConcreteRegion
    Dim outer As CConcreteRegion, opening As CConcreteRegion, coordinates(1 To 8, 1 To 2) As Double
    Set query = New CSectionGeometryQuery
    Set region = query.CircleRegion(13#, -7#, 150#)
    TestContourMeshFixture stats, "Circle.50x100", region, 50#, 100#
    Set region = TestRoundedRegion(query, 400#, 220#, 35#, 17#, -13#)
    TestContourMeshFixture stats, "Rounded.50x100", region, 50#, 100#
    Set outer = TestRoundedRegion(query, 500#, 800#, 180#, 0#, 0#)
    Set opening = TestRoundedRegion(query, 200#, 500#, 30#, 0#, 0#)
    Set region = query.SubtractContainedRegion(outer, opening)
    TestContourMeshFixture stats, "HollowRounded.50x100", region, 50#, 100#
    Set region = query.RectangleRegion(0#, 0#, 120#, 140#)
    Set opening = query.CircleRegion(-35#, -20#, 5#)
    Set region = query.SubtractContainedRegion(region, opening)
    Set opening = query.CircleRegion(20#, 60#, 3#)
    Set region = query.SubtractContainedRegion(region, opening)
    TestContourMeshFixture stats, "MultipleContainedHoles.50x100", region, 50#, 100#
    coordinates(1, 1) = -3#: coordinates(1, 2) = -3#
    coordinates(2, 1) = 3#: coordinates(2, 2) = -3#
    coordinates(3, 1) = 3#: coordinates(3, 2) = 3#
    coordinates(4, 1) = 1#: coordinates(4, 2) = 3#
    coordinates(5, 1) = 1#: coordinates(5, 2) = -1#
    coordinates(6, 1) = -1#: coordinates(6, 2) = -1#
    coordinates(7, 1) = -1#: coordinates(7, 2) = 3#
    coordinates(8, 1) = -3#: coordinates(8, 2) = 3#
    Set region = query.PolygonRegion(coordinates)
    TestContourMeshFixture stats, "Concave.TwoPartsInCell", region, 10#, 3#, 3
    Set outer = query.CircleRegion(0#, 0#, 100#)
    Set region = query.ClipWindow(outer, 1#, 0#, 0#, 100#, 0#, 100#)
    TestContourMeshFixture stats, "Quarter.ExactBounds", region, 100#, 100#, 1
    stats.Report = stats.Report & "STAGE ThinArc.constructBoundary" & vbCrLf
    Set region = TestArcSegmentRegion(query, 1000#, 0.002)
    stats.Report = stats.Report & "STAGE ThinArc.integrateAndBuild" & vbCrLf
    TestContourMeshFixture stats, "ThinArc.NoEquivalentFallback", region, 1#, 1#
    TestLegacyAdaptiveMesh stats, 2, 256
    TestLegacyAdaptiveMesh stats, 4, 980
    TestContourSetterContract stats
    TestContourAtomicAdd stats
    TestMinimumAreaMesh stats
    TestFiberGeometryIsolation stats
    TestConcreteElementEntity stats
    TestLegacyEntityInterpretation stats
    TestTinyLinearProjection stats
    TestContourIgnoresAdaptiveDepth stats
    TestHollowTangentMesh stats
    stats.Report = stats.Report & "TOTAL: passed=" & CStr(stats.Passed) & "; failed=" & CStr(stats.Failed) & vbCrLf
    RunContourCellMeshTests = stats.Report
    Exit Function
Failed:
    RunContourCellMeshTests = stats.Report & "RUNTIME ERROR: " & CStr(Err.Number) & "; source=" & Err.Source & "; description=" & Err.Description & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Луч через касательный конец должен считать нижнюю вершину по направлению
' внутренней части дуги. Наружный контур и отверстие проверяют оба знака Sweep.
Private Sub TestTangentArcEndpoints(ByRef stats As TTestStats)
    Dim query As CSectionGeometryQuery, region As CConcreteRegion, outer As CConcreteRegion
    Dim caseIndex As Long, atStart As Boolean, upper As Boolean, prefix As String
    Dim edges(1 To 2) As TRegionEdge
    Set query = New CSectionGeometryQuery
    For caseIndex = 0 To 3
        atStart = (caseIndex Mod 2 = 0): upper = (caseIndex >= 2)
        prefix = "tangent.corner." & CStr(caseIndex)
        Set region = TestTangentCornerRegion(query, atStart, upper)
        TestTrue stats, prefix & ".left", query.ContainsPoint(region, -5#, 0#, False) = atStart
        TestTrue stats, prefix & ".right", query.ContainsPoint(region, 5#, 0#, False) = (Not atStart)
        TestTrue stats, prefix & ".outside", Not query.ContainsPoint(region, -15#, 0#, False)
        Set outer = query.RectangleRegion(0#, 0#, 100#, 100#)
        Set region = query.SubtractContainedRegion(outer, region)
        TestTrue stats, prefix & ".hole.left", query.ContainsPoint(region, -5#, 0#, False) = (Not atStart)
        TestTrue stats, prefix & ".hole.right", query.ContainsPoint(region, 5#, 0#, False) = atStart
    Next caseIndex
    Set region = query.CircleRegion(0#, 0#, 10#)
    TestTrue stats, "tangent.fullCircle.lowerOutside", Not query.ContainsPoint(region, -15#, -10#, False)
    TestTrue stats, "tangent.fullCircle.upperOutside", Not query.ContainsPoint(region, -15#, 10#, False)
    TestTrue stats, "tangent.fullCircle.middleOutside", Not query.ContainsPoint(region, -15#, 0#, False)
    TestTrue stats, "tangent.fullCircle.centerInside", query.ContainsPoint(region, 0#, 0#, False)
    edges(1).X1 = -10#: edges(1).Y1 = 0#: edges(1).X2 = 10#: edges(1).Y2 = 0#: edges(1).Sweep = GEOM_PI
    edges(2).X1 = 10#: edges(2).Y1 = 0#: edges(2).X2 = -10#: edges(2).Y2 = 0#
    Set region = query.BoundaryRegion(edges, 2)
    TestTrue stats, "tangent.interior.lowerOutside", Not query.ContainsPoint(region, -15#, -10#, False)
    TestTrue stats, "tangent.interior.lowerInside", query.ContainsPoint(region, 0#, -5#, False)
    edges(1).Sweep = -GEOM_PI
    Set region = query.BoundaryRegion(edges, 2)
    TestTrue stats, "tangent.interior.upperOutside", Not query.ContainsPoint(region, -15#, 10#, False)
    TestTrue stats, "tangent.interior.upperInside", query.ContainsPoint(region, 0#, 5#, False)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Прямые замыкают одну четверть так, чтобы луч y=0 проходил через стык
' вертикальной грани и горизонтально касательной дуги, но не через тестовую точку.
Private Function TestTangentCornerRegion(ByVal query As CSectionGeometryQuery, ByVal atStart As Boolean, _
        ByVal upper As Boolean) As CConcreteRegion
    Dim edges(1 To 5) As TRegionEdge, i As Long
    If atStart Then
        edges(1).X1 = 0#: edges(1).Y1 = 0#: edges(1).X2 = 10#: edges(1).Y2 = 10#
        edges(2).X1 = 10#: edges(2).Y1 = 10#: edges(2).X2 = -10#: edges(2).Y2 = 10#
        edges(3).X1 = -10#: edges(3).Y1 = 10#: edges(3).X2 = -10#: edges(3).Y2 = -10#
        edges(4).X1 = -10#: edges(4).Y1 = -10#: edges(4).X2 = 0#: edges(4).Y2 = -10#
        edges(5).X1 = 0#: edges(5).Y1 = -10#: edges(5).X2 = 0#: edges(5).Y2 = 0#
    Else
        edges(1).X1 = -10#: edges(1).Y1 = 10#: edges(1).X2 = 0#: edges(1).Y2 = 0#
        edges(2).X1 = 0#: edges(2).Y1 = 0#: edges(2).X2 = 0#: edges(2).Y2 = -10#
        edges(3).X1 = 0#: edges(3).Y1 = -10#: edges(3).X2 = 10#: edges(3).Y2 = -10#
        edges(4).X1 = 10#: edges(4).Y1 = -10#: edges(4).X2 = 10#: edges(4).Y2 = 10#
        edges(5).X1 = 10#: edges(5).Y1 = 10#: edges(5).X2 = -10#: edges(5).Y2 = 10#
    End If
    edges(1).CenterY = 10#: edges(1).Sweep = GEOM_PI / 2#
    If upper Then
        For i = 1 To 5
            edges(i).Y1 = -edges(i).Y1: edges(i).Y2 = -edges(i).Y2
            edges(i).CenterY = -edges(i).CenterY: edges(i).Sweep = -edges(i).Sweep
        Next i
    End If
    Set TestTangentCornerRegion = query.BoundaryRegion(edges, 5)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторяет ячейку Workflow02: срез y=-250 совпадает с нижним касанием
' отверстия R30 и обязан сохранить замыкающий отрезок [-175,-170].
Private Sub TestHollowTangentCells(ByRef stats As TTestStats)
    Dim query As CSectionGeometryQuery, region As CConcreteRegion, cell As CConcreteRegion
    Dim properties As TGeometryProperties, expected As Double, index As Long
    Set query = New CSectionGeometryQuery: Set region = TestHollowTangentRegion(query)
    For index = 0 To 1
        Set cell = query.ClipWindow(region, 1#, 0#, -200# + index * 25#, -175# + index * 25#, -250#, -200#)
        query.RegionProperties cell, properties
        TestTrue stats, "tangent.hollow.cell.positive." & CStr(index), properties.Area > 0#
        TestClose stats, "tangent.hollow.cell.oneLoop." & CStr(index), CDbl(cell.LoopCount), 1#, 0#
    Next index
    expected = 150# - (5# * Sqr(875#) + 900# * Atn(5# / Sqr(875#))) / 2# ' Интеграл 30-sqrt(900-x²) от -5 до 0, без геометрического ядра как эталона.
    TestClose stats, "tangent.hollow.cell.arcIntegral", properties.Area, expected, 0.00000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит точное сечение из подтверждённого рабочего сбоя без Excel и Config.
Private Function TestHollowTangentRegion(ByVal query As CSectionGeometryQuery) As CConcreteRegion
    Dim outer As CConcreteRegion, opening As CConcreteRegion
    Set outer = TestRoundedRegion(query, 1000#, 1200#, 60#, 0#, 0#)
    Set opening = TestRoundedRegion(query, 400#, 500#, 30#, 0#, 0#)
    Set TestHollowTangentRegion = query.SubtractContainedRegion(outer, opening)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Полная сетка 25x50 должна замкнуть все ячейки и сохранить A/c/I исходного
' HollowCentered дополнительно к более крупной сетке 50x100.
Private Sub TestHollowTangentMesh(ByRef stats As TTestStats)
    Dim query As CSectionGeometryQuery, region As CConcreteRegion, geometry As CGeometryRegion
    Dim mesh As CFiberMeshBuilder, model As CSectionModel, builder As CSectionModelBuilder
    Dim calculator As CSectionPropertiesCalculator, reference As TGeometryProperties
    Set query = New CSectionGeometryQuery: Set region = TestHollowTangentRegion(query)
    query.RegionProperties region, reference
    Set geometry = New CGeometryRegion: geometry.Initialize region
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geometry, 25#, 50#, 1, 2, "Test Workflow02 HollowCentered25x50", 0#, "Contour"
    TestClose stats, "tangent.hollow.meshCount", CDbl(mesh.FiberCount), 808#, 0#
    Set builder = New CSectionModelBuilder: Set model = builder.BuildFromGenerated(mesh, Nothing, "TestHollowTangent")
    Set calculator = New CSectionPropertiesCalculator: calculator.CalculateConcrete model
    TestClose stats, "tangent.hollow.meshArea", calculator.Area, reference.Area, reference.Area * 0.000000001
    TestClose stats, "tangent.hollow.meshCx", calculator.CentroidX, reference.CentroidX, 0.000001
    TestClose stats, "tangent.hollow.meshCy", calculator.CentroidY, reference.CentroidY, 0.000001
    TestClose stats, "tangent.hollow.meshIx", calculator.Ixc, reference.Ixc, reference.Ixc * 0.000000001
    TestClose stats, "tangent.hollow.meshIy", calculator.Iyc, reference.Iyc, reference.Iyc * 0.000000001
    TestClose stats, "tangent.hollow.meshIxy", calculator.Ixyc, reference.Ixyc, GeomMax(reference.Ixc, reference.Iyc) * 0.000000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет неизменяемую сущность: метаданные и A/c/I сохраняются при замене,
' входная область копируется, повторный Initialize и поворот точной формы запрещены.
Private Sub TestConcreteElementEntity(ByRef stats As TTestStats)
    Dim original As CConcreteElement, rotated As CConcreteElement, exact As CConcreteElement
    Dim query As CSectionGeometryQuery, region As CConcreteRegion, saved As CConcreteRegion, errorNumber As Long
    Dim replacement As CConcreteRegion, changedEdges() As TRegionEdge, edgeCount As Long, openingFlags(1 To 1) As Boolean
    Set query = New CSectionGeometryQuery: Set original = New CConcreteElement
    original.Initialize "C7", 1#, 2#, 800#, 3, "SourceName", "ABC123", "Rectangle", _
        40#, 20#, 0.25, False, "Комментарий", 11#, 22#, -3#, Nothing, True
    Set region = original.Geometry
    TestClose stats, "entity.legacyLazyArea", region.Area, 800#, 0.000001
    TestTrue stats, "entity.legacyNotExact", Not original.HasExactGeometry
    Set rotated = original.WithRotation(0.5)
    TestClose stats, "entity.originalRotationUnchanged", original.Rotation, 0.25, 0#
    TestClose stats, "entity.newRotation", rotated.Rotation, 0.5, 0#
    TestTrue stats, "entity.rotationUnknownPreserved", Not rotated.RotationKnown
    TestTrue stats, "entity.knownRectanglePreserved", original.HasKnownRectangle And rotated.HasKnownRectangle
    Set region = query.RectangleRegion(1#, 2#, 40#, 20#, 0.5)
    region.SetOwner "BeforeCopy", "Test", "Anchor"
    Set exact = rotated.WithGeometry(region, 40#, 20#, 0.5)
    region.SetOwner "AfterCopy", "Changed", "Changed"
    Set saved = exact.Geometry
    TestTrue stats, "entity.independentInputRegion", saved.OwnerID = "BeforeCopy" And saved.Standard = "Test" And saved.AnchorID = "Anchor"
    TestTrue stats, "entity.exactMarker", exact.HasExactGeometry And exact.ShapeType = "Contour"
    TestTrue stats, "entity.exactDoesNotInventLegacyFact", Not exact.HasKnownRectangle
    TestTrue stats, "entity.sourceIdentity", exact.ElementID = "C7" And exact.SourceName = "SourceName" And exact.SourceHandle = "ABC123"
    TestClose stats, "entity.cxPreserved", exact.X, 1#, 0#
    TestClose stats, "entity.cyPreserved", exact.Y, 2#, 0#
    TestClose stats, "entity.areaPreserved", exact.Area, 800#, 0#
    TestClose stats, "entity.materialPreserved", CDbl(exact.MaterialID), 3#, 0#
    TestClose stats, "entity.IxPreserved", exact.LocalIx, 11#, 0#
    TestClose stats, "entity.IyPreserved", exact.LocalIy, 22#, 0#
    TestClose stats, "entity.IxyPreserved", exact.LocalIxy, -3#, 0#
    TestTrue stats, "entity.commentPreserved", exact.Comment = "Комментарий"
    TestTrue stats, "entity.oldEntityStillLegacy", Not rotated.HasExactGeometry And rotated.ShapeType = "Rectangle"
    Set replacement = query.RectangleRegion(100#, 200#, 10#, 20#)
    replacement.CopyEdges changedEdges, edgeCount
    saved.SetGeometry changedEdges, edgeCount, 1, replacement.Area, "ChangedCopy", 99, openingFlags
    saved.SetOwner "MutatedOutputCopy", "Changed", "Changed"
    Set saved = exact.Geometry
    TestClose stats, "entity.outputCopyPreservesArea", saved.Area, 800#, 0.000001
    TestClose stats, "entity.outputCopyPreservesEntityArea", exact.Area, 800#, 0#
    TestTrue stats, "entity.outputCopyPreservesMetadata", saved.OwnerID = "BeforeCopy" And saved.Source <> "ChangedCopy"
    Set saved = original.Geometry: saved.SetOwner "MutatedLegacyCopy", "Changed", "Changed"
    Set saved = original.Geometry
    TestTrue stats, "entity.legacyOutputCopyIndependent", saved.OwnerID <> "MutatedLegacyCopy"
    On Error Resume Next
    Set rotated = exact.WithRotation(1#)
    errorNumber = Err.Number: Err.Clear
    On Error GoTo 0
    TestTrue stats, "entity.exactRotationRejected", errorNumber <> 0
    On Error Resume Next
    original.Initialize "Changed", 0#, 0#, 1#, 1, "", "", "Rectangle", 1#, 1#, 0#, True, "", 1#, 1#, 0#
    errorNumber = Err.Number: Err.Clear
    On Error GoTo 0
    TestTrue stats, "entity.reinitializationRejected", errorNumber <> 0 And original.ElementID = "C7" And original.Area = 800#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Прямой API Contour не зависит от параметра глубины Adaptive: ноль и четыре
' уровня дают те же центры, площади и собственные инерции каждой точной ячейки.
Private Sub TestContourIgnoresAdaptiveDepth(ByRef stats As TTestStats)
    Dim query As CSectionGeometryQuery, geometry As CGeometryRegion, region As CConcreteRegion
    Dim zeroDepth As CFiberMeshBuilder, fourDepth As CFiberMeshBuilder, i As Long, equal As Boolean
    Dim firstIx As Double, firstIy As Double, firstIxy As Double, secondIx As Double, secondIy As Double, secondIxy As Double
    Set query = New CSectionGeometryQuery: Set region = query.CircleRegion(13#, -7#, 150#)
    Set geometry = New CGeometryRegion: geometry.Initialize region
    Set zeroDepth = New CFiberMeshBuilder: Set fourDepth = New CFiberMeshBuilder
    zeroDepth.BuildMesh geometry, 50#, 100#, 7, 0, "Ignored Adaptive depth", 0#, "Contour"
    fourDepth.BuildMesh geometry, 50#, 100#, 7, 4, "Ignored Adaptive depth", 0#, "Contour"
    equal = zeroDepth.FiberCount = fourDepth.FiberCount
    If equal Then
        For i = 1 To zeroDepth.FiberCount
            zeroDepth.FiberLocalInertia i, firstIx, firstIy, firstIxy
            fourDepth.FiberLocalInertia i, secondIx, secondIy, secondIxy
            If zeroDepth.FiberX(i) <> fourDepth.FiberX(i) Or zeroDepth.FiberY(i) <> fourDepth.FiberY(i) Or _
                    zeroDepth.FiberArea(i) <> fourDepth.FiberArea(i) Or firstIx <> secondIx Or firstIy <> secondIy Or firstIxy <> secondIxy Then equal = False
            Set region = zeroDepth.FiberGeometry(i): If region Is Nothing Then equal = False
        Next i
    End If
    TestTrue stats, "contour.zeroAdaptiveDepth.exactEquality", equal
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Подготовленная эквивалентная ширина не превращает компактную запись Rectangle
' без raw-размеров в достоверный прямоугольник. Усреднённые направления кругов
' остаются неизвестными источнику и не участвуют в следующем усреднении.
Private Sub TestLegacyEntityInterpretation(ByRef stats As TTestStats)
    Dim model As CSectionModel, element As CConcreteElement, previousRevision As Long
    Set model = New CSectionModel
    model.AddConcreteElement 0#, 0#, 100#, 1, vbNullString, vbNullString, "Rectangle", _
        0#, 0#, 0#, vbNullString, 1000#, 1000#, 0#
    Set element = model.ConcreteElement(1)
    TestTrue stats, "legacy.implicitRectangle.notKnown", Not element.HasKnownRectangle
    TestTrue stats, "legacy.implicitRectangle.equivalentSquare", model.ConcreteBoundaryShapeType(1) = "Equivalent square"
    TestClose stats, "legacy.implicitRectangle.preparedWidth", element.Width, 10#, 0.000000001
    Set model = New CSectionModel
    model.AddConcreteElement 0#, 0#, 200#, 1, vbNullString, vbNullString, "Rectangle", 10#, 20#, 0.4
    model.AddConcreteElement 100#, 0#, 100#, 1, vbNullString, vbNullString, "Region", _
        0#, 0#, 0#, vbNullString, 1000#, 1000#, 0#
    model.ApplyAverageRotationToEquivalentAreaFallbacks
    Set element = model.ConcreteElement(2)
    TestClose stats, "legacy.average.first", element.Rotation, 0.4, 0.000000001
    TestTrue stats, "legacy.average.staysUnknown", Not element.RotationKnown
    model.AddConcreteElement 200#, 0#, 200#, 1, vbNullString, vbNullString, "Rectangle", 10#, 20#, -0.2
    TestClose stats, "legacy.average.excludesAssignedCircle", model.AverageKnownConcreteElementRotation, 0.1, 0.000000001
    model.ApplyAverageRotationToEquivalentAreaFallbacks
    Set element = model.ConcreteElement(2)
    TestClose stats, "legacy.average.recomputed", element.Rotation, 0.1, 0.000000001
    TestTrue stats, "legacy.average.stillUnknown", Not element.RotationKnown
    previousRevision = model.Revision
    model.ApplyAverageRotationToEquivalentAreaFallbacks
    TestClose stats, "legacy.average.idempotentRevision", CDbl(model.Revision), CDbl(previousRevision), 0#
    Set model = New CSectionModel
    model.AddConcreteElement 0#, 0#, 200#, 1, vbNullString, vbNullString, " Rectangle ", 20#, 10#
    Set element = model.ConcreteElement(1)
    TestTrue stats, "legacy.trimmedRectangle.known", element.HasKnownRectangle
    TestTrue stats, "legacy.trimmedRectangle.status", model.ConcreteBoundaryShapeType(1) = "Rectangle"
    TestClose stats, "legacy.trimmedRectangle.Ix", model.ConcreteLocalIx(1), 200# * 10# ^ 2 / 12#, 0.000000001
    TestClose stats, "legacy.trimmedRectangle.width", model.ConcreteWidth(1), 20#, 0.000000001
    TestClose stats, "legacy.trimmedRectangle.height", model.ConcreteHeight(1), 10#, 0.000000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Линейные коэффициенты ядра сечения имеют размерность 1/мм3 и могут быть
' намного меньше геометрического допуска длины. Их масштаб нельзя отбрасывать.
Private Sub TestTinyLinearProjection(ByRef stats As TTestStats)
    Dim query As CSectionGeometryQuery, circleRegion As CConcreteRegion, region As CConcreteRegion
    Dim properties As TGeometryProperties, model As CSectionModel, minimum As Double, maximum As Double
    Set query = New CSectionGeometryQuery: Set circleRegion = query.CircleRegion(0#, 0#, 100#)
    Set region = query.ClipWindow(circleRegion, 1#, 0#, 0#, 100#, 0#, 100#)
    query.RegionProperties region, properties
    Set model = New CSectionModel
    model.AddConcreteElement properties.CentroidX, properties.CentroidY, properties.Area, 1, "", "", "Contour", _
        0#, 0#, 0#, "", properties.Ixc, properties.Iyc, properties.Ixyc, True, region
    model.ConcreteProjectionBounds 1, 0.000000000001, 0#, minimum, maximum
    TestClose stats, "projection.tinyAxis.min", minimum, 0#, 0.00000000000000000001
    TestClose stats, "projection.tinyAxis.max", maximum, 0.0000000001, 0.00000000000000000001
    model.ConcreteProjectionBounds 1, -0.000000000001, 0#, minimum, maximum
    TestClose stats, "projection.tinyNegativeAxis.min", minimum, -0.0000000001, 0.00000000000000000001
    TestClose stats, "projection.tinyNegativeAxis.max", maximum, 0#, 0.00000000000000000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт аналитический скруглённый прямоугольник через общий построитель путей.
Private Function TestRoundedRegion(ByVal query As CSectionGeometryQuery, ByVal width As Double, _
        ByVal height As Double, ByVal radius As Double, ByVal cx As Double, ByVal cy As Double) As CConcreteRegion
    Dim polygon As CPolygonSpec, path As CGeometryPath
    Set polygon = New CPolygonSpec
    polygon.AddVertex "BL", cx - width / 2#, cy - height / 2#, radius
    polygon.AddVertex "BR", cx + width / 2#, cy - height / 2#, radius
    polygon.AddVertex "TR", cx + width / 2#, cy + height / 2#, radius
    polygon.AddVertex "TL", cx - width / 2#, cy + height / 2#, radius
    polygon.Seal
    Set path = query.RoundedPolygon("TestRounded", polygon)
    Set TestRoundedRegion = query.RegionFromPath(path)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает A/c/I модели с областью и каждый элемент с его настоящей границей.
' Точность не зависит от глубины Adaptive: в Contour уровни 2 и 4 дают одну сетку.
Private Sub TestContourMeshFixture(ByRef stats As TTestStats, ByVal label As String, _
        ByVal region As CConcreteRegion, ByVal stepX As Double, ByVal stepY As Double, _
        Optional ByVal expectedCount As Long = 0)
    Dim query As CSectionGeometryQuery, geometry As CGeometryRegion, mesh As CFiberMeshBuilder
    Dim builder As CSectionModelBuilder, model As CSectionModel, calculator As CSectionPropertiesCalculator
    Dim reference As TGeometryProperties, properties As TGeometryProperties, cell As CConcreteRegion
    Dim depth As Long, firstCount As Long, i As Long, localIx As Double, localIy As Double, localIxy As Double
    Dim minExpected As Double, maxExpected As Double, minActual As Double, maxActual As Double
    Dim tolerance As Double, prefix As String, reconstructed As CConcreteRegion
    Set query = New CSectionGeometryQuery: query.RegionProperties region, reference
    Set geometry = New CGeometryRegion: geometry.Initialize region
    For depth = 2 To 4 Step 2
        prefix = label & ".depth" & CStr(depth)
        Set mesh = New CFiberMeshBuilder
        mesh.BuildMesh geometry, stepX, stepY, 7, depth, "Test exact cells", 0#, "Contour"
        TestTrue stats, prefix & ".nonEmpty", mesh.FiberCount > 0
        If depth = 2 Then firstCount = mesh.FiberCount
        TestClose stats, prefix & ".depthDoesNotChangeCount", CDbl(mesh.FiberCount), CDbl(firstCount), 0#
        If expectedCount > 0 Then TestClose stats, prefix & ".expectedCount", CDbl(mesh.FiberCount), CDbl(expectedCount), 0#
        Set builder = New CSectionModelBuilder: Set model = builder.BuildFromGenerated(mesh, Nothing, "TestContour")
        TestClose stats, prefix & ".modelCount", CDbl(model.ConcreteCount), CDbl(mesh.FiberCount), 0#
        For i = 1 To mesh.FiberCount
            Set cell = mesh.FiberGeometry(i)
            TestTrue stats, prefix & ".fiberBoundary." & CStr(i), Not cell Is Nothing
            If cell Is Nothing Then Err.Raise vbObjectError + 5990, "modTestElementGeometry", "Потеряна точная граница тестового волокна."
            query.RegionProperties cell, properties
            mesh.FiberLocalInertia i, localIx, localIy, localIxy
            TestClose stats, prefix & ".fiberArea." & CStr(i), mesh.FiberArea(i), properties.Area, GeomMax(properties.Area * 0.000000001, 0.000000000001)
            TestClose stats, prefix & ".fiberIx." & CStr(i), localIx, properties.Ixc, GeomMax(Abs(properties.Ixc) * 0.000000001, 0.00000000000000000001)
            TestClose stats, prefix & ".fiberIy." & CStr(i), localIy, properties.Iyc, GeomMax(Abs(properties.Iyc) * 0.000000001, 0.00000000000000000001)
            TestClose stats, prefix & ".fiberIxy." & CStr(i), localIxy, properties.Ixyc, GeomMax(Abs(properties.Ixyc) * 0.000000001, 0.00000000000000000001)
            TestTrue stats, prefix & ".modelExact." & CStr(i), model.ConcreteHasBoundary(i) And model.ConcreteShapeType(i) = "Contour"
            TestClose stats, prefix & ".material." & CStr(i), CDbl(model.ConcreteMaterialID(i)), 7#, 0#
            model.ConcreteLocalInertiaComponents i, localIx, localIy, localIxy
            TestClose stats, prefix & ".modelIx." & CStr(i), localIx, properties.Ixc, GeomMax(Abs(properties.Ixc) * 0.000000001, 0.00000000000000000001)
            TestClose stats, prefix & ".modelIy." & CStr(i), localIy, properties.Iyc, GeomMax(Abs(properties.Iyc) * 0.000000001, 0.00000000000000000001)
            TestClose stats, prefix & ".modelIxy." & CStr(i), localIxy, properties.Ixyc, GeomMax(Abs(properties.Ixyc) * 0.000000001, 0.00000000000000000001)
        Next i
        Set calculator = New CSectionPropertiesCalculator: calculator.CalculateConcrete model
        TestClose stats, prefix & ".areaBalance", calculator.Area, reference.Area, GeomMax(reference.Area * 0.000000001, 0.000000000001)
        TestClose stats, prefix & ".cxBalance", calculator.CentroidX, reference.CentroidX, 0.0000001
        TestClose stats, prefix & ".cyBalance", calculator.CentroidY, reference.CentroidY, 0.0000001
        TestClose stats, prefix & ".IxBalance", calculator.Ixc, reference.Ixc, GeomMax(Abs(reference.Ixc) * 0.000000001, 0.00000000000000000001)
        TestClose stats, prefix & ".IyBalance", calculator.Iyc, reference.Iyc, GeomMax(Abs(reference.Iyc) * 0.000000001, 0.00000000000000000001)
        tolerance = GeomMax(GeomMax(Abs(reference.Ixc), Abs(reference.Iyc)) * 0.000000001, 0.00000000000000000001)
        TestClose stats, prefix & ".IxyBalance", calculator.Ixyc, reference.Ixyc, tolerance
        query.ProjectionBounds region, 1#, 0#, minExpected, maxExpected
        calculator.CalculateProjection model, 1#, 0#, False, minActual, maxActual
        TestClose stats, prefix & ".minX", minActual, minExpected, 0.0000001
        TestClose stats, prefix & ".maxX", maxActual, maxExpected, 0.0000001
        query.ProjectionBounds region, 0#, 1#, minExpected, maxExpected
        calculator.CalculateProjection model, 0#, 1#, False, minActual, maxActual
        TestClose stats, prefix & ".minY", minActual, minExpected, 0.0000001
        TestClose stats, prefix & ".maxY", maxActual, maxExpected, 0.0000001
        query.Initialize model: Set reconstructed = query.ConcreteDomain
        TestClose stats, prefix & ".meshBoundaryArea", reconstructed.Area, reference.Area, GeomMax(reference.Area * 0.000000001, 0.000000000001)
    Next depth
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает вызов с режимом по умолчанию и явным Adaptive побитно по
' данным элементов и с ранее принятыми количествами круга D300 при шаге 30.
Private Sub TestLegacyAdaptiveMesh(ByRef stats As TTestStats, ByVal depth As Long, ByVal expectedCount As Long)
    Dim query As CSectionGeometryQuery, geometry As CGeometryRegion, region As CConcreteRegion
    Dim original As CFiberMeshBuilder, explicitMode As CFiberMeshBuilder, cell As CConcreteRegion
    Dim equal As Boolean, noBoundary As Boolean, i As Long, firstMismatch As Long
    Set query = New CSectionGeometryQuery: Set region = query.CircleRegion(0#, 0#, 150#)
    Set geometry = New CGeometryRegion: geometry.Initialize region
    Set original = New CFiberMeshBuilder: Set explicitMode = New CFiberMeshBuilder
    original.BuildMesh geometry, 30#, 30#, 3, depth
    explicitMode.BuildMesh geometry, 30#, 30#, 3, depth, vbNullString, 0#, "Adaptive"
    TestClose stats, "adaptive.count." & CStr(depth), CDbl(original.FiberCount), CDbl(expectedCount), 0#
    equal = (original.FiberCount = explicitMode.FiberCount): noBoundary = True
    If equal Then
        For i = 1 To original.FiberCount
            If original.FiberX(i) <> explicitMode.FiberX(i) Or original.FiberY(i) <> explicitMode.FiberY(i) Or _
                    original.FiberWidth(i) <> explicitMode.FiberWidth(i) Or original.FiberHeight(i) <> explicitMode.FiberHeight(i) Or _
                    original.FiberArea(i) <> explicitMode.FiberArea(i) Or original.FiberMaterialID(i) <> explicitMode.FiberMaterialID(i) Then
                equal = False: If firstMismatch = 0 Then firstMismatch = i
            End If
            Set cell = original.FiberGeometry(i): If Not cell Is Nothing Then noBoundary = False
        Next i
    End If
    TestTrue stats, "adaptive.exactLegacyEquality." & CStr(depth) & ".firstMismatch=" & CStr(firstMismatch), equal
    TestTrue stats, "adaptive.noForcedBoundary." & CStr(depth), noBoundary
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Несогласованная граница не должна частично менять готовый элемент или ревизию.
' У записи Contour потеря обязательной границы запрещает эквивалентную оболочку.
Private Sub TestContourSetterContract(ByRef stats As TTestStats)
    Dim query As CSectionGeometryQuery, model As CSectionModel, region As CConcreteRegion
    Dim properties As TGeometryProperties, previousRevision As Long, errorNumber As Long
    Set query = New CSectionGeometryQuery: Set region = query.RectangleRegion(0#, 0#, 40#, 20#)
    query.RegionProperties region, properties
    Set model = New CSectionModel
    model.AddConcreteElement 0#, 0#, properties.Area, 1, vbNullString, vbNullString, "Rectangle", 40#, 20#, 0#, vbNullString, properties.Ixc, properties.Iyc, properties.Ixyc
    previousRevision = model.Revision
    Set region = query.RectangleRegion(10#, 0#, 40#, 20#)
    On Error Resume Next
    model.SetConcreteGeometry 1, region
    errorNumber = Err.Number: Err.Clear
    On Error GoTo 0
    TestTrue stats, "setter.rejectsWrongCenter", errorNumber <> 0
    TestTrue stats, "setter.rejectionKeepsBoundary", Not model.ConcreteHasBoundary(1)
    TestClose stats, "setter.rejectionKeepsRevision", CDbl(model.Revision), CDbl(previousRevision), 0#
    Set model = New CSectionModel
    On Error Resume Next
    model.AddConcreteElement 0#, 0#, 800#, 1, vbNullString, vbNullString, "Contour", 40#, 20#, 0#, vbNullString, properties.Ixc, properties.Iyc, properties.Ixyc
    errorNumber = Err.Number: Err.Clear
    On Error GoTo 0
    TestTrue stats, "setter.missingContourRejectedBeforePublication", errorNumber <> 0 And model.ConcreteCount = 0 And model.Revision = 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Отказы до публикации не расходуют имя Cn, не меняют существующую сущность,
' числовую проекцию, ревизию и подготовленные центры/границы общей геометрии.
Private Sub TestContourAtomicAdd(ByRef stats As TTestStats)
    Dim query As CSectionGeometryQuery, cachedQuery As CSectionGeometryQuery, model As CSectionModel
    Dim region As CConcreteRegion, candidate As CConcreteRegion, domain As CConcreteRegion, properties As TGeometryProperties
    Dim original As CConcreteElement, current As CConcreteElement, previousRevision As Long, buildCount As Long
    Dim caseIndex As Long, errorNumber As Long, x As Double, y As Double, area As Double
    Dim localIx As Double, localIy As Double, localIxy As Double, cx As Double, cy As Double, totalArea As Double
    Dim minimum As Double, maximum As Double, prefix As String
    Set query = New CSectionGeometryQuery: Set model = New CSectionModel
    model.AddConcreteElement 0#, 0#, 800#, 1, "Existing", "Handle1", "Rectangle", 40#, 20#
    Set original = model.ConcreteElement(1)
    model.ConcreteMeshCentroid cx, cy, totalArea
    model.ConcreteProjectionBounds 1, 1#, 0#, minimum, maximum
    TestTrue stats, "atomic.initialPreparedBounds", minimum = -20# And maximum = 20#
    Set cachedQuery = New CSectionGeometryQuery: cachedQuery.Initialize model
    Set domain = cachedQuery.ConcreteDomain: buildCount = cachedQuery.TestDomainBuildCount
    previousRevision = model.Revision
    Set region = query.CircleRegion(70#, 30#, 7#): query.RegionProperties region, properties
    For caseIndex = 1 To 10
        Set candidate = region
        x = properties.CentroidX: y = properties.CentroidY: area = properties.Area
        localIx = properties.Ixc: localIy = properties.Iyc: localIxy = properties.Ixyc
        Select Case caseIndex
            Case 1: Set candidate = Nothing
            Case 2: Set candidate = New CConcreteRegion
            Case 3: x = x + 1#
            Case 4: y = y + 1#
            Case 5: area = area * 1.1
            Case 6: localIx = localIx * 1.1
            Case 7: localIy = localIy * 1.1
            Case 8: localIxy = Sqr(localIx) * Sqr(localIy) * 2#
            Case 9: localIx = 0#
            Case 10: area = 0#
        End Select
        errorNumber = 0
        On Error Resume Next
        model.AddConcreteElement x, y, area, 2, "Rejected", "Rejected", "Contour", 0#, 0#, 0#, "", localIx, localIy, localIxy, True, candidate
        errorNumber = Err.Number: Err.Clear
        On Error GoTo 0
        prefix = "atomic.case" & CStr(caseIndex)
        TestTrue stats, prefix & ".reject", errorNumber <> 0
        TestTrue stats, prefix & ".countRevision", model.ConcreteCount = 1 And model.Revision = previousRevision
        Set current = model.ConcreteElement(1)
        TestTrue stats, prefix & ".sameEntity", current Is original
        TestTrue stats, prefix & ".sameIDAndProjection", model.ConcreteID(1) = "C1" And model.ConcreteX(1) = 0# And model.ConcreteY(1) = 0# And model.ConcreteArea(1) = 800#
        model.ConcreteMeshCentroid cx, cy, totalArea
        TestTrue stats, prefix & ".sameCenter", cx = 0# And cy = 0# And totalArea = 800#
        Set domain = cachedQuery.ConcreteDomain
        TestTrue stats, prefix & ".sameDomainCache", cachedQuery.TestDomainBuildCount = buildCount And domain.Area = 800#
        model.ConcreteProjectionBounds 1, 1#, 0#, minimum, maximum
        TestTrue stats, prefix & ".samePreparedBounds", minimum = -20# And maximum = 20#
    Next caseIndex
    model.AddConcreteElement properties.CentroidX, properties.CentroidY, properties.Area, 2, "Accepted", "Handle2", "Contour", _
        0#, 0#, 0#, "", properties.Ixc, properties.Iyc, properties.Ixyc, True, region
    TestTrue stats, "atomic.acceptUsesNextID", model.ConcreteCount = 2 And model.ConcreteID(2) = "C2"
    TestTrue stats, "atomic.acceptSingleRevision", model.Revision = previousRevision + 1
    TestTrue stats, "atomic.acceptHasBoundary", model.ConcreteHasBoundary(2) And model.ConcreteShapeType(2) = "Contour"
    region.SetOwner "MutatedInput", "", ""
    Set candidate = model.ConcreteGeometry(2)
    TestTrue stats, "atomic.acceptOwnsBoundaryCopy", candidate.OwnerID <> "MutatedInput"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Порог относится к площади каждой конечной части: равенство сохраняется,
' отверстия уменьшают площадь, а несвязные фрагменты проверяются раздельно.
Private Sub TestMinimumAreaMesh(ByRef stats As TTestStats)
    Dim query As CSectionGeometryQuery, geometry As CGeometryRegion, region As CConcreteRegion, opening As CConcreteRegion
    Dim mesh As CFiberMeshBuilder, baseline As CFiberMeshBuilder, filtered As CFiberMeshBuilder
    Dim modeIndex As Long, modeText As String, errorNumber As Long, i As Long, expectedCount As Long, acceptedIndex As Long
    Dim coordinates(1 To 8, 1 To 2) As Double
    Set query = New CSectionGeometryQuery: Set geometry = New CGeometryRegion
    For modeIndex = 1 To 2
        modeText = "Adaptive": If modeIndex = 2 Then modeText = "Contour"
        Set region = query.RectangleRegion(0#, 0#, 20#, 10#): geometry.Initialize region
        Set mesh = New CFiberMeshBuilder
        mesh.BuildMesh geometry, 10#, 10#, 1, 1, "Test exact threshold", 0#, modeText, 100#
        TestClose stats, "minArea." & modeText & ".equalityKeep", CDbl(mesh.FiberCount), 2#, 0#
        Set mesh = New CFiberMeshBuilder
        On Error Resume Next
        mesh.BuildMesh geometry, 10#, 10#, 1, 1, "Test threshold exclusion", 0#, modeText, 100.0001
        errorNumber = Err.Number: Err.Clear
        On Error GoTo 0
        TestTrue stats, "minArea." & modeText & ".smallerRejected", errorNumber <> 0 And mesh.FiberCount = 0
        Set region = query.CircleRegion(0#, 0#, 15#): geometry.Initialize region
        Set baseline = New CFiberMeshBuilder: Set filtered = New CFiberMeshBuilder
        baseline.BuildMesh geometry, 10#, 10#, 1, 1, "Test unfiltered final cells", 0#, modeText, 0#
        filtered.BuildMesh geometry, 10#, 10#, 1, 1, "Test filtered final cells", 0#, modeText, 100#
        expectedCount = 0: acceptedIndex = 0
        For i = 1 To baseline.FiberCount
            If baseline.FiberArea(i) >= 100# Then
                expectedCount = expectedCount + 1: acceptedIndex = acceptedIndex + 1
                If acceptedIndex <= filtered.FiberCount Then
                    TestTrue stats, "minArea." & modeText & ".acceptedUnchanged." & CStr(i), _
                        baseline.FiberX(i) = filtered.FiberX(acceptedIndex) And baseline.FiberY(i) = filtered.FiberY(acceptedIndex) And baseline.FiberArea(i) = filtered.FiberArea(acceptedIndex)
                End If
            End If
        Next i
        TestTrue stats, "minArea." & modeText & ".eachFinalCell", filtered.FiberCount = expectedCount And filtered.FiberCount < baseline.FiberCount
    Next modeIndex
    coordinates(1, 1) = -3#: coordinates(1, 2) = -3#: coordinates(2, 1) = 3#: coordinates(2, 2) = -3#
    coordinates(3, 1) = 3#: coordinates(3, 2) = 3#: coordinates(4, 1) = 2#: coordinates(4, 2) = 3#
    coordinates(5, 1) = 2#: coordinates(5, 2) = -1#: coordinates(6, 1) = -1#: coordinates(6, 2) = -1#
    coordinates(7, 1) = -1#: coordinates(7, 2) = 3#: coordinates(8, 1) = -3#: coordinates(8, 2) = 3#
    Set region = query.PolygonRegion(coordinates): geometry.Initialize region
    Set mesh = New CFiberMeshBuilder: mesh.BuildMesh geometry, 10#, 3#, 1, 2, "Test disconnected threshold", 0#, "Contour", 4#
    TestClose stats, "minArea.disconnected.independentParts", CDbl(mesh.FiberCount), 2#, 0#
    TestClose stats, "minArea.disconnected.keptArea", mesh.FiberArea(1) + mesh.FiberArea(2), 21#, 0.000000001
    Set region = query.RectangleRegion(0#, 0#, 20#, 10#): Set opening = query.CircleRegion(0#, 0#, 2#)
    Set region = query.SubtractContainedRegion(region, opening): geometry.Initialize region
    Set mesh = New CFiberMeshBuilder: mesh.BuildMesh geometry, 20#, 10#, 1, 2, "Test hole net area", 0#, "Contour", 0#
    TestClose stats, "minArea.hole.netArea", mesh.FiberArea(1), 200# - 4# * GEOM_PI, 0.000000001
    Set region = mesh.FiberGeometry(1)
    TestTrue stats, "minArea.hole.keptBoundary", region.LoopCount = 2
    Set mesh = New CFiberMeshBuilder
    On Error Resume Next
    mesh.BuildMesh geometry, 20#, 10#, 1, 2, "Test hole threshold", 0#, "Contour", 195#
    errorNumber = Err.Number: Err.Clear
    On Error GoTo 0
    TestTrue stats, "minArea.hole.rejectsNetAreaBelowThreshold", errorNumber <> 0 And mesh.FiberCount = 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Изменение выданной области не меняет сетку, её A/c/I и следующие копии;
' компактный Adaptive-прямоугольник не создаёт область только ради чтения.
Private Sub TestFiberGeometryIsolation(ByRef stats As TTestStats)
    Dim query As CSectionGeometryQuery, geometry As CGeometryRegion, mesh As CFiberMeshBuilder
    Dim region As CConcreteRegion, returned As CConcreteRegion, replacement As CConcreteRegion
    Dim edges() As TRegionEdge, edgeCount As Long, openingFlags(1 To 1) As Boolean
    Dim localIx As Double, localIy As Double, localIxy As Double
    Set query = New CSectionGeometryQuery: Set geometry = New CGeometryRegion
    Set region = query.RectangleRegion(0#, 0#, 40#, 20#): geometry.Initialize region
    Set mesh = New CFiberMeshBuilder: mesh.BuildMesh geometry, 40#, 20#, 1, 1, "Test getter copy", 0#, "Contour"
    Set returned = mesh.FiberGeometry(1): Set replacement = query.RectangleRegion(100#, 200#, 10#, 10#)
    replacement.CopyEdges edges, edgeCount
    returned.SetGeometry edges, edgeCount, 1, replacement.Area, "MutatedFiberCopy", 99, openingFlags
    returned.SetOwner "MutatedFiberCopy", "", ""
    Set returned = mesh.FiberGeometry(1)
    TestTrue stats, "mesh.copy.preservesBoundary", returned.Area = 800# And returned.OwnerID <> "MutatedFiberCopy"
    TestClose stats, "mesh.copy.preservesArea", mesh.FiberArea(1), 800#, 0#
    mesh.FiberLocalInertia 1, localIx, localIy, localIxy
    query.ValidateBoundaryProperties returned, mesh.FiberX(1), mesh.FiberY(1), mesh.FiberArea(1), localIx, localIy, localIxy
    TestTrue stats, "mesh.copy.preservesProperties", True
    Set mesh = New CFiberMeshBuilder: mesh.BuildMesh geometry, 40#, 20#, 1, 1, "Test compact getter", 0#, "Adaptive"
    Set returned = mesh.FiberGeometry(1)
    TestTrue stats, "mesh.compactGetterNothing", returned Is Nothing
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт точный малый сегмент окружности с замыкающей хордой для независимых
' проверок свойств и переноса тонких элементов без эквивалентного квадрата.
Private Function TestArcSegmentRegion(ByVal query As CSectionGeometryQuery, ByVal radius As Double, _
        ByVal theta As Double) As CConcreteRegion
    Dim edges(1 To 2) As TRegionEdge
    edges(1).X1 = radius * Cos(theta / 2#): edges(1).Y1 = -radius * Sin(theta / 2#)
    edges(1).X2 = edges(1).X1: edges(1).Y2 = -edges(1).Y1: edges(1).Sweep = theta
    edges(2).X1 = edges(1).X2: edges(2).Y1 = edges(1).Y2
    edges(2).X2 = edges(1).X1: edges(2).Y2 = edges(1).Y1
    Set TestArcSegmentRegion = query.BoundaryRegion(edges, 2)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает повёрнутые прямоугольники с независимыми формулами и переносом WCS.
Private Sub TestExactRectangle(ByRef stats As TTestStats)
    Dim query As CSectionGeometryQuery, region As CConcreteRegion, properties As TGeometryProperties
    Dim mode As Long, cx As Double, cy As Double, rotation As Double, cosine As Double, sine As Double
    Dim actualX As Double, actualY As Double, width As Double, height As Double, angle As Double
    Set query = New CSectionGeometryQuery
    For mode = 0 To 2
        cx = 10#: cy = -5#: rotation = 0#
        If mode > 0 Then rotation = GEOM_PI / 6#
        If mode = 2 Then cx = cx + 100000000#: cy = cy - 200000000#
        cosine = Cos(rotation): sine = Sin(rotation)
        Set region = query.RectangleRegion(cx, cy, 40#, 20#, rotation)
        query.RegionProperties region, properties
        TestClose stats, "rectangle.area." & CStr(mode), properties.Area, 800#, 0.00001
        TestClose stats, "rectangle.cx." & CStr(mode), properties.CentroidX, cx, 0.000001
        TestClose stats, "rectangle.cy." & CStr(mode), properties.CentroidY, cy, 0.000001
        TestClose stats, "rectangle.Ix." & CStr(mode), properties.Ixc, 800# / 12# * (400# * cosine ^ 2 + 1600# * sine ^ 2), 0.01
        TestClose stats, "rectangle.Iy." & CStr(mode), properties.Iyc, 800# / 12# * (1600# * cosine ^ 2 + 400# * sine ^ 2), 0.01
        TestClose stats, "rectangle.Ixy." & CStr(mode), properties.Ixyc, 800# / 12# * 1200# * sine * cosine, 0.01
        TestTrue stats, "rectangle.recognized." & CStr(mode), query.TryRectangle(region, actualX, actualY, width, height, angle)
        TestClose stats, "rectangle.width." & CStr(mode), width, 40#, 0.000001
        TestClose stats, "rectangle.height." & CStr(mode), height, 20#, 0.000001
        TestClose stats, "rectangle.angle." & CStr(mode), angle, rotation, 0.00000001
    Next mode
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Полная окружность и четверть проверяются по аналитическим интегралам, включая Ixy.
Private Sub TestExactCircle(ByRef stats As TTestStats)
    Dim query As CSectionGeometryQuery, region As CConcreteRegion, clipped As CConcreteRegion
    Dim properties As TGeometryProperties, radius As Double, area As Double, offset As Double
    Dim cx As Double, cy As Double, width As Double, height As Double, angle As Double
    Set query = New CSectionGeometryQuery
    radius = 100#: Set region = query.CircleRegion(1234#, -4567#, radius)
    query.RegionProperties region, properties
    TestClose stats, "circle.area", properties.Area, GEOM_PI * radius ^ 2, 0.000001
    TestClose stats, "circle.cx", properties.CentroidX, 1234#, 0.000001
    TestClose stats, "circle.cy", properties.CentroidY, -4567#, 0.000001
    TestClose stats, "circle.Ix", properties.Ixc, GEOM_PI * radius ^ 4 / 4#, 0.0001
    TestClose stats, "circle.Iy", properties.Iyc, GEOM_PI * radius ^ 4 / 4#, 0.0001
    TestClose stats, "circle.Ixy", properties.Ixyc, 0#, 0.0001
    TestTrue stats, "circle.notRectangle", Not query.TryRectangle(region, cx, cy, width, height, angle)
    Set region = query.CircleRegion(0#, 0#, radius)
    Set clipped = query.ClipWindow(region, 1#, 0#, 0#, radius, 0#, radius)
    query.RegionProperties clipped, properties
    area = GEOM_PI * radius ^ 2 / 4#: offset = 4# * radius / (3# * GEOM_PI)
    TestClose stats, "quarter.area", properties.Area, area, 0.000001
    TestClose stats, "quarter.cx", properties.CentroidX, offset, 0.000001
    TestClose stats, "quarter.cy", properties.CentroidY, offset, 0.000001
    TestClose stats, "quarter.Ix", properties.Ixc, GEOM_PI * radius ^ 4 / 16# - area * offset ^ 2, 0.0001
    TestClose stats, "quarter.Iy", properties.Iyc, properties.Ixc, 0.0001
    TestClose stats, "quarter.Ixy", properties.Ixyc, radius ^ 4 / 8# - area * offset ^ 2, 0.0001
    Set clipped = query.ClipWindow(region, 1#, 0#, radius, radius + 5#, -radius, radius)
    TestClose stats, "tangent.empty", clipped.Area, 0#, 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Смещённое отверстие проверяет знаки моментов и то, что центр области может
' находиться внутри пустоты без потери компонента при разборе границ.
Private Sub TestExactOpening(ByRef stats As TTestStats)
    Dim query As CSectionGeometryQuery, region As CConcreteRegion, outer As CConcreteRegion, opening As CConcreteRegion
    Dim properties As TGeometryProperties, area As Double, cx As Double, cy As Double, components As Collection
    Set query = New CSectionGeometryQuery
    Set outer = query.CircleRegion(0#, 0#, 20#): Set opening = query.CircleRegion(4#, -3#, 5#)
    Set region = query.SubtractContainedRegion(outer, opening)
    query.RegionProperties region, properties
    area = GEOM_PI * 375#: cx = -100# / 375#: cy = 75# / 375#
    TestClose stats, "opening.area", properties.Area, area, 0.000001
    TestClose stats, "opening.cx", properties.CentroidX, cx, 0.000001
    TestClose stats, "opening.cy", properties.CentroidY, cy, 0.000001
    TestClose stats, "opening.Ix", properties.Ixc, GEOM_PI * (20# ^ 4 - 5# ^ 4) / 4# - GEOM_PI * 25# * 9# - area * cy ^ 2, 0.000001
    TestClose stats, "opening.Iy", properties.Iyc, GEOM_PI * (20# ^ 4 - 5# ^ 4) / 4# - GEOM_PI * 25# * 16# - area * cx ^ 2, 0.000001
    TestClose stats, "opening.Ixy", properties.Ixyc, GEOM_PI * 300# - area * cx * cy, 0.000001
    Set components = query.ConnectedRegions(region)
    TestClose stats, "opening.componentCount", CDbl(components.Count), 1#, 0#
    Set region = components(1)
    TestClose stats, "opening.loopCount", CDbl(region.LoopCount), 2#, 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Разделение сохраняет отверстия своих наружных контуров и различает два
' фрагмента одной исходной ячейки после отсечения нижней перемычки U-контура.
Private Sub TestExactComponents(ByRef stats As TTestStats)
    Dim query As CSectionGeometryQuery, first As CConcreteRegion, second As CConcreteRegion
    Dim region As CConcreteRegion, parts As Collection, inputs As Collection, opening As CConcreteRegion
    Dim coordinates(1 To 8, 1 To 2) As Double, i As Long, area As Double
    Set query = New CSectionGeometryQuery
    Set first = query.RectangleRegion(0#, 0#, 40#, 30#)
    Set opening = query.CircleRegion(0#, 0#, 3#)
    Set first = query.SubtractContainedRegion(first, opening)
    Set second = query.RectangleRegion(100#, 0#, 10#, 20#)
    Set inputs = New Collection: inputs.Add first: inputs.Add second
    Set region = query.MergeDisjointRegions(inputs): Set parts = query.ConnectedRegions(region)
    TestClose stats, "disconnected.count", CDbl(parts.Count), 2#, 0#
    For i = 1 To parts.Count
        Set region = parts(i): area = area + region.Area
        If region.LoopCount = 2 Then TestClose stats, "disconnected.holeKept", region.Area, 1200# - GEOM_PI * 9#, 0.000001
    Next i
    TestClose stats, "disconnected.area", area, 1400# - GEOM_PI * 9#, 0.000001
    Set first = query.CircleRegion(0#, 0#, 100#)
    Set opening = query.CircleRegion(0#, 0#, 50#)
    Set first = query.SubtractContainedRegion(first, opening)
    Set second = query.CircleRegion(0#, 0#, 20#)
    Set opening = query.CircleRegion(0#, 0#, 5#)
    Set second = query.SubtractContainedRegion(second, opening)
    Set inputs = New Collection: inputs.Add first: inputs.Add second
    Set region = query.MergeDisjointRegions(inputs): Set parts = query.ConnectedRegions(region)
    TestClose stats, "nested.count", CDbl(parts.Count), 2#, 0#
    area = 0#
    For i = 1 To parts.Count
        Set region = parts(i): area = area + region.Area
        TestClose stats, "nested.holeAssigned." & CStr(i), CDbl(region.LoopCount), 2#, 0#
    Next i
    TestClose stats, "nested.area", area, GEOM_PI * (10000# - 2500# + 400# - 25#), 0.000001
    coordinates(1, 1) = -3#: coordinates(1, 2) = -3#
    coordinates(2, 1) = 3#: coordinates(2, 2) = -3#
    coordinates(3, 1) = 3#: coordinates(3, 2) = 3#
    coordinates(4, 1) = 1#: coordinates(4, 2) = 3#
    coordinates(5, 1) = 1#: coordinates(5, 2) = -1#
    coordinates(6, 1) = -1#: coordinates(6, 2) = -1#
    coordinates(7, 1) = -1#: coordinates(7, 2) = 3#
    coordinates(8, 1) = -3#: coordinates(8, 2) = 3#
    Set region = query.PolygonRegion(coordinates)
    Set region = query.ClipWindow(region, 1#, 0#, -4#, 4#, 0#, 4#)
    Set parts = query.ConnectedRegions(region)
    TestClose stats, "concave.clipped.count", CDbl(parts.Count), 2#, 0#
    area = 0#
    For i = 1 To parts.Count
        Set region = parts(i): area = area + region.Area
        TestClose stats, "concave.component.area." & CStr(i), region.Area, 6#, 0.000001
    Next i
    TestClose stats, "concave.clipped.area", area, 12#, 0.000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сумма несмежно хранимых ячеек точно восстанавливает A, статические моменты
' и инерции исходного круга, включая прямоугольные базовые ячейки и дуги.
Private Sub TestClippedCellClosure(ByRef stats As TTestStats)
    Dim query As CSectionGeometryQuery, geometry As CGeometryRegion, region As CConcreteRegion
    Dim cells As Collection, cell As CConcreteRegion, properties As TGeometryProperties
    Dim i As Long, j As Long, k As Long, totalArea As Double, totalX As Double, totalY As Double
    Dim totalIx As Double, totalIy As Double, totalIxy As Double
    Set query = New CSectionGeometryQuery: Set geometry = New CGeometryRegion
    Set region = query.CircleRegion(3#, -2#, 20#): geometry.Initialize region
    For i = 0 To 4
        For j = 0 To 3
            Set cells = geometry.ClipCell(-22# + i * 10#, -24# + j * 13#, -12# + i * 10#, -11# + j * 13#)
            For k = 1 To cells.Count
                Set cell = cells(k): query.RegionProperties cell, properties
                totalArea = totalArea + properties.Area
                totalX = totalX + properties.StaticMomentY: totalY = totalY + properties.StaticMomentX
                totalIx = totalIx + properties.Ix: totalIy = totalIy + properties.Iy: totalIxy = totalIxy + properties.Ixy
                TestTrue stats, "cell.area.positive." & CStr(i) & "." & CStr(j) & "." & CStr(k), properties.Area > 0#
            Next k
        Next j
    Next i
    TestClose stats, "closure.area", totalArea, GEOM_PI * 400#, 0.000001
    TestClose stats, "closure.Sy", totalX, GEOM_PI * 400# * 3#, 0.00001
    TestClose stats, "closure.Sx", totalY, -GEOM_PI * 400# * 2#, 0.00001
    TestClose stats, "closure.Ix", totalIx, GEOM_PI * 20# ^ 4 / 4# + GEOM_PI * 400# * 4#, 0.0001
    TestClose stats, "closure.Iy", totalIy, GEOM_PI * 20# ^ 4 / 4# + GEOM_PI * 400# * 9#, 0.0001
    TestClose stats, "closure.Ixy", totalIxy, -GEOM_PI * 400# * 6#, 0.0001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Малый круговой сегмент проверяет устойчивость при theta-sin(theta) и
' центральном радиальном моменте, который на семь порядков меньше поперечного.
Private Sub TestSmallArcSegment(ByRef stats As TTestStats)
    Dim query As CSectionGeometryQuery, region As CConcreteRegion, properties As TGeometryProperties
    Dim edges(1 To 2) As TRegionEdge, radius As Double, theta As Double, square As Double
    Dim expectedArea As Double, expectedIx As Double, expectedIy As Double
    Set query = New CSectionGeometryQuery
    radius = 1000#: theta = 0.002: square = theta * theta
    edges(1).X1 = radius * Cos(theta / 2#): edges(1).Y1 = -radius * Sin(theta / 2#)
    edges(1).X2 = edges(1).X1: edges(1).Y2 = -edges(1).Y1: edges(1).Sweep = theta
    edges(2).X1 = edges(1).X2: edges(2).Y1 = edges(1).Y2
    edges(2).X2 = edges(1).X1: edges(2).Y2 = edges(1).Y1
    stats.Report = stats.Report & "STAGE smallArc.constructBoundary" & vbCrLf
    Set region = query.BoundaryRegion(edges, 2)
    stats.Report = stats.Report & "STAGE smallArc.integrateProperties" & vbCrLf
    query.RegionProperties region, properties
    expectedArea = radius ^ 2 * theta ^ 3 / 12# * (1# - square / 20# + square ^ 2 / 840#)
    expectedIx = radius ^ 4 * theta ^ 5 * (1# / 240# - square / 2016# + square ^ 2 / 34560#)
    expectedIy = radius ^ 4 * theta ^ 7 * (1# / 11200# - square / 126000# + square ^ 2 * 173# / 517440000#)
    TestClose stats, "smallArc.area", properties.Area, expectedArea, expectedArea * 0.00000001
    TestClose stats, "smallArc.Ix", properties.Ixc, expectedIx, expectedIx * 0.00000001
    TestClose stats, "smallArc.Iy", properties.Iyc, expectedIy, expectedIy * 0.000001
    TestClose stats, "smallArc.Ixy", properties.Ixyc, 0#, expectedIx * 0.00000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает фактическое отклонение без округления проверяемых значений.
Private Sub TestClose(ByRef stats As TTestStats, ByVal checkLabel As String, ByVal actual As Double, _
        ByVal expected As Double, ByVal tolerance As Double)
    TestTrue stats, checkLabel & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) <= tolerance
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет каждую проверку для первичного отчёта и итогового счётчика.
Private Sub TestTrue(ByRef stats As TTestStats, ByVal checkLabel As String, ByVal passed As Boolean)
    If passed Then
        stats.Passed = stats.Passed + 1: stats.Report = stats.Report & "PASS " & checkLabel & vbCrLf
    Else
        stats.Failed = stats.Failed + 1: stats.Report = stats.Report & "FAIL " & checkLabel & vbCrLf
    End If
End Sub
