Attribute VB_Name = "modTestGeometryQuery"
Option Explicit

' ==========================================================================
' Адресные тесты общего geometry-query/topology движка
' ==========================================================================
' Проверяются аналитические площади/дуги, отверстия, notch, замкнутость,
' T-соединения сетки, интервалы прямой/луча/отрезка и локальная стенка.
' НДС не решается, нормативные коэффициенты не назначаются, AutoCAD не нужен.
' Реальный AutoCAD import/export проверяется отдельным integration-сценарием.

Private mPassed As Long
Private mFailed As Long
Private mReport As String

' ДЛЯ ТЕСТОВ: сравнивает геометрический скаляр с независимым эталоном.
Private Sub CheckNear(ByVal name As String, ByVal actual As Double, ByVal expected As Double, Optional ByVal tolerance As Double = 0.00001)
    Check name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) <= tolerance
End Sub

' ДЛЯ ТЕСТОВ: сохраняет проверку в короткий воспроизводимый протокол.
Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        mPassed = mPassed + 1: mReport = mReport & "OK: " & name & vbCrLf
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
    End If
End Sub

' ДЛЯ ТЕСТОВ: готовит четыре точки контура без геометрического калькулятора.
Private Function Rectangle(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Variant
    Dim points(1 To 4, 1 To 2) As Double
    points(1, 1) = x1: points(1, 2) = y1
    points(2, 1) = x2: points(2, 2) = y1
    points(3, 1) = x2: points(3, 2) = y2
    points(4, 1) = x1: points(4, 2) = y2
    Rectangle = points
End Function

' ДЛЯ ТЕСТОВ: проверяет каждый замкнутый контур готового снимка, включая границы отверстий.
Private Sub CheckClosed(ByVal name As String, ByVal region As CConcreteRegion)
    Dim i As Long, loopID As Long, previousLoop As Long
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, cx As Double, cy As Double, sweep As Double
    Dim firstX As Double, firstY As Double, previousX As Double, previousY As Double
    For i = 1 To region.SegmentCount
        region.GetSegment i, x1, y1, x2, y2, cx, cy, sweep, loopID
        If loopID <> previousLoop Then
            If previousLoop > 0 Then
                CheckNear name & ".loopEndX", previousX, firstX
                CheckNear name & ".loopEndY", previousY, firstY
            End If
            firstX = x1: firstY = y1
        Else
            CheckNear name & ".joinedX", x1, previousX
            CheckNear name & ".joinedY", y1, previousY
        End If
        previousLoop = loopID: previousX = x2: previousY = y2
    Next i
    If previousLoop > 0 Then
        CheckNear name & ".lastEndX", previousX, firstX
        CheckNear name & ".lastEndY", previousY, firstY
    End If
End Sub

' ДЛЯ ТЕСТОВ: проверяет независимые эталоны прямоугольника, полосы и точной дуги.
Private Sub TestAnalyticClipping(ByVal query As CSectionGeometryQuery)
    Dim region As CConcreteRegion, clipped As CConcreteRegion, copy As CConcreteRegion
    Set region = query.PolygonRegion(Rectangle(0#, 0#, 100#, 80#))
    CheckNear "rectangle.area", region.Area, 8000#
    Set copy = query.PolygonRegion(Rectangle(0#, 0#, 100#, 80#))
    Check "regionEquality.same", query.RegionsCoincide(region, copy)
    Set copy = query.PolygonRegion(Rectangle(100#, 0#, 200#, 80#))
    Check "regionEquality.sameAreaDifferentPosition", Not query.RegionsCoincide(region, copy)
    Set copy = query.PolygonRegion(Rectangle(0#, 0#, 80#, 100#))
    Check "regionEquality.sameAreaDifferentShape", Not query.RegionsCoincide(region, copy)
    Set clipped = query.ClipWindow(region, 0#, 1#, 20#, 70#, -80#, -10#)
    CheckNear "rectangle.window", clipped.Area, 3500#
    CheckClosed "rectangle.window", clipped
    Set clipped = query.ClipHalfPlane(region, 1#, 1#, 100#)
    CheckNear "rectangle.diagonal", clipped.Area, 3200#
    Set clipped = query.ClipHalfPlane(region, 1#, 0#, 100#)
    CheckNear "rectangle.tangent.empty", clipped.Area, 0#
    Check "rectangle.tangent.noLoop", clipped.LoopCount = 0
    Set clipped = query.ClipHalfPlane(region, 1#, 0#, 150#)
    CheckNear "rectangle.outside.empty", clipped.Area, 0#
    Set clipped = query.ClipHalfPlane(region, 1#, 0#, -1#)
    CheckNear "rectangle.inside.unchanged", clipped.Area, 8000#
    Set region = query.CircleRegion(31#, -17#, 100#)
    CheckNear "circle.area", region.Area, GEOM_PI * 10000#
    Set clipped = query.ClipHalfPlane(region, 1#, 0#, 31#)
    CheckNear "circle.half.analytic", clipped.Area, GEOM_PI * 5000#
    CheckClosed "circle.half", clipped
    Set clipped = query.ClipHalfPlane(region, 1#, 0#, 81#)
    CheckNear "circle.segment.analytic", clipped.Area, 10000# * GEOM_PI / 3# - 50# * Sqr(7500#)
    CheckClosed "circle.segment", clipped
    Set copy = clipped.Clone
    CheckNear "copy.sameArea", copy.Area, clipped.Area
    Check "copy.sameSegments", copy.SegmentCount = clipped.SegmentCount
    clipped.SetOwner "LC1", "TEST", "R1"
    Set copy = clipped.Clone
    Check "copy.owner", copy.OwnerID = "LC1" And copy.AnchorID = "R1"
End Sub

' ДЛЯ ТЕСТОВ: отверстие сохраняется inner loop либо становится вырезом при
' пересечении границы clipping. Площадь одного и того же отверстия не удваивается.
Private Sub TestOpenings(ByVal query As CSectionGeometryQuery)
    Dim region As CConcreteRegion, hole As CConcreteRegion, withHole As CConcreteRegion, clipped As CConcreteRegion
    Set region = query.PolygonRegion(Rectangle(0#, 0#, 100#, 100#))
    Set hole = query.PolygonRegion(Rectangle(20#, 30#, 60#, 70#))
    Set withHole = query.SubtractRegion(region, hole)
    CheckNear "hole.area", withHole.Area, 8400#
    Check "hole.twoLoops", withHole.LoopCount = 2
    Check "hole.insideExcluded", Not query.ContainsPoint(withHole, 40#, 50#)
    Check "hole.edgeIncluded", query.ContainsPoint(withHole, 20#, 50#)
    Check "hole.materialIncluded", query.ContainsPoint(withHole, 10#, 50#)
    CheckClosed "hole", withHole
    Set clipped = query.ClipHalfPlane(withHole, 0#, 1#, 50#)
    CheckNear "notch.area", clipped.Area, 4200#
    Check "notch.oneLoop", clipped.LoopCount = 1
    CheckClosed "notch", clipped
    Set withHole = query.SubtractRegion(withHole, hole)
    CheckNear "hole.noDoubleSubtract", withHole.Area, 8400#
    Set hole = query.CircleRegion(40#, 50#, 10#)
    Set withHole = query.SubtractRegion(region, hole)
    CheckNear "circleHole.area", withHole.Area, 10000# - 100# * GEOM_PI
    Check "circleHole.centerExcluded", Not query.ContainsPoint(withHole, 40#, 50#)
    Check "circleHole.leftMaterial", query.ContainsPoint(withHole, 5#, 50#)
    Check "circleHole.rightMaterial", query.ContainsPoint(withHole, 95#, 50#)
    Check "circleHole.leftOutside", Not query.ContainsPoint(withHole, -5#, 50#)
    Set clipped = query.ClipHalfPlane(withHole, 1#, 0#, 40#)
    CheckNear "circleNotch.area", clipped.Area, 6000# - 50# * GEOM_PI
    Check "circleNotch.oneLoop", clipped.LoopCount = 1
    CheckClosed "circleNotch", clipped
End Sub

' ДЛЯ ТЕСТОВ: одна линия дает одинаковые интервалы при ограничении до луча
' или отрезка; изменение направления не превращает пустоту в материал.
Private Sub TestIntervals(ByVal query As CSectionGeometryQuery)
    Dim region As CConcreteRegion, hole As CConcreteRegion, intervals As Variant
    Set region = query.PolygonRegion(Rectangle(0#, 0#, 100#, 100#))
    Set hole = query.PolygonRegion(Rectangle(20#, 30#, 60#, 70#))
    Set region = query.SubtractRegion(region, hole)
    intervals = query.LineIntervals(region, 10#, 50#, 1#, 0#)
    Check "line.twoIntervals", UBound(intervals, 1) = 2
    CheckNear "line.leftStart", intervals(1, 1), -10#
    CheckNear "line.leftEnd", intervals(1, 2), 10#
    CheckNear "line.rightStart", intervals(2, 1), 50#
    CheckNear "line.rightEnd", intervals(2, 2), 90#
    intervals = query.LineIntervals(region, 10#, 50#, 1#, 0#, 0#)
    CheckNear "ray.startZero", intervals(1, 1), 0#
    CheckNear "ray.firstOpening", intervals(1, 2), 10#
    intervals = query.LineIntervals(region, 10#, 50#, 1#, 0#, 0#, 5#)
    Check "segment.oneInterval", UBound(intervals, 1) = 1
    CheckNear "segment.cut", intervals(1, 2), 5#
    intervals = query.LineIntervals(region, 10#, 50#, -1#, 0#, 0#)
    CheckNear "reverse.firstExit", intervals(1, 2), 10#
    intervals = query.LineIntervals(region, 10#, 150#, 1#, 0#)
    Check "line.outside.empty", IsEmpty(intervals)
End Sub

' ДЛЯ ТЕСТОВ: неравные соседние КЭ образуют T-соединение без внутреннего ребра.
' Затем к той же session добавляется элемент, что проверяет revision invalidation.
Private Sub TestMesh(ByVal query As CSectionGeometryQuery)
    Dim model As CSectionModel, region As CConcreteRegion
    Set model = New CSectionModel
    model.AddConcreteElement 5#, 10#, 200#, 1, , , "Rectangle", 10#, 20#
    model.AddConcreteElement 15#, 5#, 100#, 1, , , "Rectangle", 10#, 10#
    model.AddConcreteElement 15#, 15#, 100#, 1, , , "Rectangle", 10#, 10#
    query.Initialize model
    Set region = query.ConcreteDomain
    CheckNear "mesh.tJoint.area", region.Area, 400#
    Check "mesh.tJoint.oneLoop", region.LoopCount = 1
    CheckClosed "mesh.tJoint", region
    CheckNear "mesh.cover", query.ConcreteCoverFromPointAlongDirection(5#, 5#, 1#, 0#), 15#
    model.AddConcreteElement 25#, 5#, 100#, 1, , , "Rectangle", 10#, 10#
    Set region = query.ConcreteDomain
    CheckNear "mesh.revision.area", region.Area, 500#
    CheckNear "mesh.revision.cover", query.ConcreteCoverFromPointAlongDirection(5#, 5#, 1#, 0#), 25#

    Set model = New CSectionModel
    Dim i As Long, j As Long
    For i = 0 To 2
        For j = 0 To 2
            If i <> 1 Or j <> 1 Then model.AddConcreteElement i * 10# + 5#, j * 10# + 5#, 100#, 1, , , "Rectangle", 10#, 10#
        Next j
    Next i
    query.Initialize model
    Set region = query.ConcreteDomain
    CheckNear "mesh.opening.area", region.Area, 800#
    Check "mesh.opening.twoLoops", region.LoopCount = 2
    CheckNear "mesh.opening.firstWall", query.ConcreteCoverFromPointAlongDirection(5#, 15#, 1#, 0#), 5#
    CheckNear "mesh.outside.noSeed", query.ConcreteCoverFromPointAlongDirection(15#, 15#, 1#, 0#), -1#
    CheckClosed "mesh.opening", region
End Sub

' ДЛЯ ТЕСТОВ: сравнивает точное отсечение с прежним выбором целых КЭ по центру.
' Независимый эталон площади прямоугольника и сумма пересеченных ячеек объясняют
' допустимое изменение СП 63 без произвольного допуска к нормативной формуле.
Private Sub TestCenterCellAreaDifference(ByVal query As CSectionGeometryQuery)
    Dim model As CSectionModel, region As CConcreteRegion, clipped As CConcreteRegion
    Dim offset As Variant, x As Double, y As Double, i As Long, j As Long
    Dim oldArea As Double, boundaryArea As Double, exactArea As Double, prefix As String
    Set model = New CSectionModel
    For i = 0 To 19
        For j = 0 To 19
            model.AddConcreteElement -190# + i * 20#, -190# + j * 20#, 400#, 1, , , "Rectangle", 20#, 20#
        Next j
    Next i
    query.Initialize model
    Set region = query.ConcreteDomain
    CheckNear "sp63.mesh.fullArea", region.Area, 160000#
    For Each offset In Array(-197#, -3#, 0#, 3#, 197#)
        oldArea = 0#: boundaryArea = 0#
        For i = 0 To 19
            For j = 0 To 19
                y = -190# + j * 20#
                If y > CDbl(offset) Then oldArea = oldArea + 400#
                If y - 10# < CDbl(offset) And y + 10# > CDbl(offset) Then boundaryArea = boundaryArea + 400#
            Next j
        Next i
        exactArea = 400# * (200# - CDbl(offset))
        Set clipped = query.ClipHalfPlane(region, 0#, 1#, CDbl(offset))
        prefix = "sp63.mesh.cut." & CStr(offset)
        CheckNear prefix & ".analytic", clipped.Area, exactArea
        Check prefix & ".boundaryBound", Abs(clipped.Area - oldArea) <= boundaryArea + 0.00001
        CheckClosed prefix, clipped
    Next offset
End Sub

' ДЛЯ ТЕСТОВ: запускает адресные geometry-query проверки без решения НДС.
Public Function RunGeometryQueryTests() As String
    On Error GoTo Failed
    mPassed = 0: mFailed = 0: mReport = vbNullString
    Dim query As CSectionGeometryQuery
    Set query = New CSectionGeometryQuery
    TestAnalyticClipping query
    TestHollowArcJunctions query
    TestOpenings query
    TestIntervals query
    TestMesh query
    TestCenterCellAreaDifference query
    TestKnownAndUnknownContours query
    TestRotatedAndTouchingGeometry query
    TestFirstMaterialWall query
    TestRotatedMaterialWall query
    TestConnectedRegionSelection query
    TestPostAudit03RegionClone query
    GoTo Finished
Failed:
    Check "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    mReport = mReport & "TOTAL: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
    RunGeometryQueryTests = mReport
End Function

' ДЛЯ ТЕСТОВ: копия нескольких частей с дугами/проемами хранит готовые роли,
' владельца и пробы. Изменение внешних массивов либо следующей копии не меняет оригинал.
Private Sub TestPostAudit03RegionClone(ByVal query As CSectionGeometryQuery)
    Dim regions As Collection, region As CConcreteRegion, copy As CConcreteRegion, nextCopy As CConcreteRegion
    Set region = query.CircleRegion(0#, 0#, 100#)
    Set region = query.SubtractRegion(region, query.CircleRegion(25#, 0#, 10#))
    Set regions = New Collection: regions.Add region
    regions.Add query.PolygonRegion(Rectangle(150#, 0#, 200#, 50#))
    Set region = query.MergeDisjointRegions(regions)
    region.SetOwner "COPY_LC", "SP35", "G7"
    Dim probes(1 To 2, 1 To 5) As Variant
    probes(1, 1) = True: probes(1, 2) = 1#: probes(1, 3) = 2#: probes(1, 4) = 3#: probes(1, 5) = 4#
    probes(2, 1) = False
    region.SetBoundaryProbes probes
    Set copy = region.Clone: Set nextCopy = copy.Clone
    Check "postAudit03.clone.multiLoop", region.LoopCount = 3
    CheckNear "postAudit03.clone.area", copy.Area, region.Area, 0#
    Check "postAudit03.clone.context", copy.Source = region.Source And copy.SectionRevision = region.SectionRevision
    Check "postAudit03.clone.owner", copy.OwnerID = "COPY_LC" And copy.Standard = "SP35" And copy.AnchorID = "G7"
    Dim edges() As TRegionEdge, copiedEdges() As TRegionEdge, count As Long, copiedCount As Long, i As Long
    region.CopyEdges edges, count: nextCopy.CopyEdges copiedEdges, copiedCount
    Check "postAudit03.clone.segmentCount", count = copiedCount
    For i = 1 To count
        Check "postAudit03.clone.edge." & CStr(i), edges(i).X1 = copiedEdges(i).X1 And edges(i).Y1 = copiedEdges(i).Y1 And _
            edges(i).X2 = copiedEdges(i).X2 And edges(i).Y2 = copiedEdges(i).Y2 And _
            edges(i).CenterX = copiedEdges(i).CenterX And edges(i).CenterY = copiedEdges(i).CenterY And _
            edges(i).Sweep = copiedEdges(i).Sweep And edges(i).LoopID = copiedEdges(i).LoopID
    Next i
    For i = 1 To region.LoopCount
        Check "postAudit03.clone.role." & CStr(i), copy.LoopIsOpening(i) = region.LoopIsOpening(i) And nextCopy.LoopIsOpening(i) = region.LoopIsOpening(i)
    Next i
    copiedEdges(1).X1 = -10000#: copy.SetOwner "OTHER", "OTHER", "OTHER"
    Dim readProbes As Variant: readProbes = copy.BoundaryProbes
    readProbes(1, 2) = 999#: copy.SetBoundaryProbes readProbes
    readProbes = region.BoundaryProbes
    Check "postAudit03.clone.probesIndependent", readProbes(1, 2) = 1# And region.BoundaryProbeCount = 2
    Check "postAudit03.clone.ownerIndependent", region.OwnerID = "COPY_LC" And nextCopy.OwnerID = "COPY_LC"
    region.CopyEdges copiedEdges, copiedCount
    Check "postAudit03.clone.edgesIndependent", copiedEdges(1).X1 = edges(1).X1
    Dim start As Double: start = Timer
    For i = 1 To 1000: Set copy = region.Clone: Next i
    mReport = mReport & "POSTAUDIT03_CLONE_1000_SECONDS=" & CStr(Timer - start) & vbCrLf
End Sub

' ДЛЯ ТЕСТОВ: связная часть сохраняет свои отверстия, но не присоединяет
' отдельный бетон внутри проема или соседнюю область. После clipping
' сквозной проем отделяет стенку, а боковой вырез не создает длинную тень.
Private Sub TestConnectedRegionSelection(ByVal query As CSectionGeometryQuery)
    Dim region As CConcreteRegion, island As CConcreteRegion, separate As CConcreteRegion
    Dim selected As CConcreteRegion, regions As Collection
    Set region = query.PolygonRegion(Rectangle(0#, 0#, 100#, 100#))
    Set region = query.SubtractRegion(region, query.PolygonRegion(Rectangle(20#, 20#, 80#, 80#)))
    Set island = query.PolygonRegion(Rectangle(30#, 30#, 70#, 70#))
    Set island = query.SubtractRegion(island, query.PolygonRegion(Rectangle(40#, 40#, 60#, 60#)))
    Set separate = query.PolygonRegion(Rectangle(120#, 0#, 140#, 20#))
    Set regions = New Collection: regions.Add region: regions.Add island: regions.Add separate
    Set region = query.MergeDisjointRegions(regions)
    region.SetOwner "fixture", "SP35", "G1"
    Set selected = query.ConnectedRegionAtPoint(region, 10#, 50#)
    CheckNear "connected.outer.area", selected.Area, 6400#
    Check "connected.outer.onlyOwnOpening", selected.LoopCount = 2
    Check "connected.outer.noIsland", Not query.ContainsPoint(selected, 35#, 35#)
    Check "connected.outer.noSeparate", Not query.ContainsPoint(selected, 130#, 10#)
    Check "connected.metadata", selected.OwnerID = "fixture" And selected.Standard = "SP35" And selected.AnchorID = "G1"
    CheckClosed "connected.outer", selected
    Set selected = query.ConnectedRegionAtPoint(region, 35#, 35#)
    CheckNear "connected.island.area", selected.Area, 1200#
    Check "connected.island.onlyOwnOpening", selected.LoopCount = 2
    Check "connected.island.noSurrounding", Not query.ContainsPoint(selected, 10#, 50#)
    Check "connected.island.hole", Not query.ContainsPoint(selected, 50#, 50#)
    CheckClosed "connected.island", selected
    Set selected = query.ConnectedRegionAtPoint(region, 50#, 50#)
    Check "connected.emptyAtOpening", selected.Area = 0# And selected.SegmentCount = 0
    Set selected = query.ConnectedRegionAtPoint(selected, 0#, 0#)
    Check "connected.emptyInput", selected.Area = 0# And selected.SegmentCount = 0
    Set selected = query.ConnectedRegionAtPoint(region, 120#, 10#)
    CheckNear "connected.onBoundary", selected.Area, 400#

    Set region = query.PolygonRegion(Rectangle(0#, 0#, 100#, 100#))
    Set region = query.SubtractRegion(region, query.PolygonRegion(Rectangle(20#, 30#, 80#, 70#)))
    Set selected = query.ClipWindow(region, 0#, 1#, 0#, 100#, -60#, -40#)
    Set selected = query.ConnectedRegionAtPoint(selected, 50#, 85#)
    CheckNear "connected.strip.separatedWall", selected.Area, 600#
    Check "connected.strip.noOppositeWall", Not query.ContainsPoint(selected, 50#, 15#)
    CheckClosed "connected.strip", selected
    Set selected = query.ClipWindow(region, 0#, 1#, 0#, 100#, -30#, -10#)
    Set selected = query.ConnectedRegionAtPoint(selected, 15#, 85#)
    CheckNear "connected.notch.localArea", selected.Area, 1600#
    Check "connected.notch.concreteBelow", query.ContainsPoint(selected, 25#, 15#)
    Check "connected.notch.holeStillExcluded", Not query.ContainsPoint(selected, 25#, 50#)
    CheckClosed "connected.notch", selected
End Sub

' ДЛЯ ТЕСТОВ: координата R22 из сохраненной HollowRectangle лежит в бетоне
' на уровне стыка прямой и скругления. Проверяем также отверстие, обе стороны
' стыка, перенос и поворот без подгонки допуска и без запуска НДС.
Private Sub TestHollowArcJunctions(ByVal query As CSectionGeometryQuery)
    Dim geometry As CGeometryHollowRectangle, section As CSectionModel
    Dim annotations As CHollowRectAnnotationBuilder, rebars As CRebarLayout
    Dim domain As CConcreteRegion, rotated As CConcreteRegion, raw() As TRegionEdge, count As Long
    Dim angle As Variant, offset As Variant, side As Variant, i As Long, x As Double, y As Double
    Set geometry = New CGeometryHollowRectangle
    geometry.Initialize 500#, 800#, 180#, 200#, 500#, 30#, 0#, 0#
    Set section = New CSectionModel
    section.AddConcreteElement 0#, -300#, 10000#, 1, , , "Rectangle", 100#, 100#
    Set annotations = New CHollowRectAnnotationBuilder: Set rebars = New CRebarLayout
    annotations.Build section, geometry, rebars
    query.Initialize section: Set domain = query.ConcreteDomain
    For Each side In Array(-1#, 1#)
        For Each offset In Array(-0.00001, -0.00000001, 0#, 0.00000001, 0.00001)
            Check "hollow.original.junction.concrete." & CStr(side) & "." & CStr(offset), query.ContainsPoint(domain, side * 210#, 220# + offset)
            Check "hollow.original.junction.opening." & CStr(side) & "." & CStr(offset), Not query.ContainsPoint(domain, 0#, 220# + offset)
        Next offset
    Next side
    For Each angle In Array(0#, 0.47, GEOM_PI / 2#)
        domain.CopyEdges raw, count
        For i = 1 To count
            x = raw(i).X1: y = raw(i).Y1
            raw(i).X1 = 100000# + x * Cos(angle) - y * Sin(angle)
            raw(i).Y1 = -200000# + x * Sin(angle) + y * Cos(angle)
            x = raw(i).X2: y = raw(i).Y2
            raw(i).X2 = 100000# + x * Cos(angle) - y * Sin(angle)
            raw(i).Y2 = -200000# + x * Sin(angle) + y * Cos(angle)
            x = raw(i).CenterX: y = raw(i).CenterY
            raw(i).CenterX = 100000# + x * Cos(angle) - y * Sin(angle)
            raw(i).CenterY = -200000# + x * Sin(angle) + y * Cos(angle)
        Next i
        Set rotated = query.BoundaryRegion(raw, count, "HollowJunctionFixture")
        For Each side In Array(-1#, 1#)
            For Each offset In Array(-0.00001, 0#, 0.00001)
                x = side * 210#: y = 220# + offset
                Check "hollow.junction.concrete." & CStr(angle) & "." & CStr(side) & "." & CStr(offset), _
                    query.ContainsPoint(rotated, 100000# + x * Cos(angle) - y * Sin(angle), -200000# + x * Sin(angle) + y * Cos(angle))
                x = side * 270#
                Check "hollow.junction.outside." & CStr(angle) & "." & CStr(side) & "." & CStr(offset), _
                    Not query.ContainsPoint(rotated, 100000# + x * Cos(angle) - y * Sin(angle), -200000# + x * Sin(angle) + y * Cos(angle))
                x = 0#
                Check "hollow.junction.opening." & CStr(angle) & "." & CStr(side) & "." & CStr(offset), _
                    Not query.ContainsPoint(rotated, 100000# + x * Cos(angle) - y * Sin(angle), -200000# + x * Sin(angle) + y * Cos(angle))
            Next offset
        Next side
    Next angle
End Sub

' ДЛЯ ТЕСТОВ: сохраняет достоверный прямоугольный контур независимо от сетки.
' Роль opening задается машинным ID, а не знаком пользовательских координат.
Private Sub AddContourRectangle(ByVal model As CSectionModel, ByVal id As String, _
        ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
    Dim points As Variant, i As Long, j As Long
    points = Rectangle(x1, y1, x2, y2)
    For i = 1 To 4
        j = i + 1: If j > 4 Then j = 1
        model.Contours.AddContourLine id & "_" & CStr(i), _
            points(i, 1), points(i, 2), points(j, 1), points(j, 2), vbNullString, id, _
            IIf(InStr(id, "OPENING") > 0, "Opening", "Outer")
    Next i
End Sub

' ДЛЯ ТЕСТОВ: поворачивает исходный контур с общим переносом. Это независимый
' эталон координат; query не участвует в подготовке преобразования.
Private Function RotatedRectangle(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, _
        ByVal y2 As Double, ByVal angle As Double, ByVal offsetX As Double, ByVal offsetY As Double) As Variant
    Dim source As Variant, points(1 To 4, 1 To 2) As Double, i As Long
    source = Rectangle(x1, y1, x2, y2)
    For i = 1 To 4
        points(i, 1) = offsetX + source(i, 1) * Cos(angle) - source(i, 2) * Sin(angle)
        points(i, 2) = offsetY + source(i, 1) * Sin(angle) + source(i, 2) * Cos(angle)
    Next i
    RotatedRectangle = points
End Function

' ДЛЯ ТЕСТОВ: площадь первой стенки инвариантна к наклону нормали и переносу.
' Проверяются прямолинейная и круговая пустоты одним и тем же публичным API.
Private Sub TestRotatedMaterialWall(ByVal query As CSectionGeometryQuery)
    Dim angle As Variant, nx As Double, ny As Double, seedPlane As Double
    Dim region As CConcreteRegion, opening As CConcreteRegion, wall As CConcreteRegion, prefix As String
    For Each angle In Array(0.13, 0.47, 1.22)
        nx = -Sin(CDbl(angle)): ny = Cos(CDbl(angle))
        seedPlane = 85# + nx * 500# + ny * -800#
        prefix = "wall.rotated." & CStr(angle)
        Set region = query.PolygonRegion(RotatedRectangle(0#, 0#, 100#, 100#, CDbl(angle), 500#, -800#))
        Set opening = query.PolygonRegion(RotatedRectangle(20#, 30#, 60#, 70#, CDbl(angle), 500#, -800#))
        Set region = query.SubtractRegion(region, opening)
        Set wall = query.MaterialRegionThroughPlane(region, nx, ny, seedPlane)
        CheckNear prefix & ".rectangle.area", wall.Area, 7200#
        CheckClosed prefix & ".rectangle", wall
        Set region = query.PolygonRegion(RotatedRectangle(0#, 0#, 100#, 100#, CDbl(angle), 500#, -800#))
        Set opening = query.CircleRegion(500# + 50# * Cos(CDbl(angle)) - 50# * Sin(CDbl(angle)), _
            -800# + 50# * Sin(CDbl(angle)) + 50# * Cos(CDbl(angle)), 10#)
        Set region = query.SubtractRegion(region, opening)
        Set wall = query.MaterialRegionThroughPlane(region, nx, ny, seedPlane)
        CheckNear prefix & ".circle.area", wall.Area, 9000# - 50# * GEOM_PI
        CheckClosed prefix & ".circle", wall
    Next angle
End Sub

' ДЛЯ ТЕСТОВ: локальная стенка заканчивается перед отверстием, даже если
' радиус будущего нормативного окна пересекает дальнюю стенку. Площадь
' криволинейной тени сравнивается с аналитическим полукругом.
Private Sub TestFirstMaterialWall(ByVal query As CSectionGeometryQuery)
    Dim region As CConcreteRegion, opening As CConcreteRegion, wall As CConcreteRegion
    Set region = query.PolygonRegion(Rectangle(0#, 0#, 100#, 100#))
    Set opening = query.PolygonRegion(Rectangle(20#, 30#, 60#, 70#))
    Set region = query.SubtractRegion(region, opening)
    Set wall = query.MaterialRegionThroughPlane(region, 0#, 1#, 85#)
    CheckNear "wall.rectangular.area", wall.Area, 7200#
    Check "wall.rectangular.nearIncluded", query.ContainsPoint(wall, 40#, 85#)
    Check "wall.rectangular.farExcluded", Not query.ContainsPoint(wall, 40#, 15#)
    Check "wall.rectangular.sideAccessible", query.ContainsPoint(wall, 10#, 15#)
    CheckClosed "wall.rectangular", wall
    Set wall = query.ClipWindow(wall, 0#, 1#, 0#, 100#, -50#, -30#)
    CheckNear "wall.rectangular.localStrip", wall.Area, 600#
    CheckClosed "wall.rectangular.localStrip", wall
    Set wall = query.MaterialRegionThroughPlane(region, 0#, 1#, 150#)
    CheckNear "wall.missingSeed.empty", wall.Area, 0#

    Set region = query.PolygonRegion(Rectangle(0#, 0#, 100#, 100#))
    Set opening = query.CircleRegion(50#, 50#, 10#)
    Set region = query.SubtractRegion(region, opening)
    Set wall = query.MaterialRegionThroughPlane(region, 0#, 1#, 85#)
    CheckNear "wall.circular.area", wall.Area, 9000# - 50# * GEOM_PI
    Check "wall.circular.farExcluded", Not query.ContainsPoint(wall, 50#, 15#)
    Check "wall.circular.nearIncluded", query.ContainsPoint(wall, 50#, 85#)
    CheckClosed "wall.circular", wall
End Sub

' ДЛЯ ТЕСТОВ: проверяет смешанные источники и инвалидирование только при
' изменении материала. Подпись результата не должна перестраивать domain.
Private Sub TestKnownAndUnknownContours(ByVal query As CSectionGeometryQuery)
    Dim model As CSectionModel, region As CConcreteRegion, i As Long, j As Long, revision As Long
    Set model = New CSectionModel
    For i = 0 To 2
        For j = 0 To 2
            If i <> 1 Or j <> 1 Then model.AddConcreteElement i * 10# + 5#, j * 10# + 5#, 100#, 1, , , "Rectangle", 10#, 10#
        Next j
    Next i
    AddContourRectangle model, "CONTOUR_OUTER", 0#, 0#, 30#, 30#
    query.Initialize model
    Set region = query.ConcreteDomain
    CheckNear "knownOuter.unknownOpening.area", region.Area, 800#
    Check "knownOuter.unknownOpening.loops", region.LoopCount = 2
    Check "knownOuter.unknownOpening.empty", Not query.ContainsPoint(region, 15#, 15#)
    CheckClosed "knownOuter.unknownOpening", region

    Set model = New CSectionModel
    model.AddConcreteElement 50#, 50#, 10000#, 1, , , "Rectangle", 100#, 100#
    query.Initialize model
    Set region = query.ConcreteDomain
    revision = region.SectionRevision
    model.Annotations.AddAnnotation "CRACK_REGION_LABEL", "RESULT1", 0#, 0#, 1#, 1#, 0#, 0#, "TEST", 0#, "-", vbNullString
    Set region = query.ConcreteDomain
    Check "resultAnnotation.noInvalidation", region.SectionRevision = revision
    AddContourRectangle model, "CONTOUR_OPENING", 20#, 30#, 60#, 70#
    Set region = query.ConcreteDomain
    Check "openingAnnotation.invalidates", region.SectionRevision > revision
    CheckNear "unknownOuter.knownOpening.area", region.Area, 8400#
    Check "unknownOuter.knownOpening.loops", region.LoopCount = 2
    Check "unknownOuter.knownOpening.empty", Not query.ContainsPoint(region, 40#, 50#)
    CheckClosed "unknownOuter.knownOpening", region
    model.Contours.Clear
    Set region = query.ConcreteDomain
    CheckNear "removedContour.rebuilt", region.Area, 10000#
End Sub

' ДЛЯ ТЕСТОВ: наклон и перенос не меняют площадь; касающиеся в вершине
' компоненты не получают фиктивную перемычку. Экстремумы дуг считаются точно.
Private Sub TestRotatedAndTouchingGeometry(ByVal query As CSectionGeometryQuery)
    Dim model As CSectionModel, region As CConcreteRegion, minimum As Double, maximum As Double
    Dim points(1 To 4, 1 To 2) As Double, original As Variant, i As Long, angle As Double
    original = Rectangle(-60#, -40#, 60#, 40#): angle = 0.47
    For i = 1 To 4
        points(i, 1) = 100000# + original(i, 1) * Cos(angle) - original(i, 2) * Sin(angle)
        points(i, 2) = -200000# + original(i, 1) * Sin(angle) + original(i, 2) * Cos(angle)
    Next i
    Set region = query.PolygonRegion(points)
    CheckNear "rotated.translated.area", region.Area, 9600#, 0.0001
    Set region = query.ClipHalfPlane(region, -Sin(angle), Cos(angle), _
        -Sin(angle) * 100000# + Cos(angle) * -200000#)
    CheckNear "rotated.translated.half", region.Area, 4800#, 0.0001
    CheckClosed "rotated.translated.half", region
    Set region = query.CircleRegion(31#, -17#, 100#)
    query.ProjectionBounds region, Cos(angle), Sin(angle), minimum, maximum
    CheckNear "arc.projection.minimum", minimum, 31# * Cos(angle) - 17# * Sin(angle) - 100#
    CheckNear "arc.projection.maximum", maximum, 31# * Cos(angle) - 17# * Sin(angle) + 100#

    Set model = New CSectionModel
    model.AddConcreteElement 5#, 5#, 100#, 1, , , "Rectangle", 10#, 10#
    model.AddConcreteElement 15#, 15#, 100#, 1, , , "Rectangle", 10#, 10#
    query.Initialize model
    Set region = query.ConcreteDomain
    CheckNear "touching.vertex.area", region.Area, 200#
    Check "touching.vertex.twoLoops", region.LoopCount = 2
    CheckClosed "touching.vertex", region
End Sub
