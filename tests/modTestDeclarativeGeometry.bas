Attribute VB_Name = "modTestDeclarativeGeometry"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: точные линии и дуги общего геометрического слоя без Excel и сетки.
' ==========================================================================
Private mPassed As Long, mFailed As Long, mReport As String

' -------------------- ДЛЯ ТЕСТОВ --------------------
' единый короткий набор общего геометрического слоя.
Public Function RunCoreGeometryTests() As String
    Dim result As String
    result = RunPathTests()
    result = result & RunPathAnchorTests()
    result = result & RunRoundedPolygonTests()
    result = result & RunRoundedOffsetTests()
    result = result & RunInputTableTests()
    RunCoreGeometryTests = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' выделение именованных участков контура, переход через его начало
' и сохранение соответствия сегментов исходным линиям.
Public Function RunPathAnchorTests() As String
    Dim polygonSpec As CPolygonSpec, polygonPoint As Variant
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    Dim path As CGeometryPath, part As CGeometryPath, changed As CGeometryPath, edge As TRegionEdge
    Dim query As CSectionGeometryQuery, facade As CSectionDefinitionBuilder, definition As CSectionDefinition
    Dim x As Double, y As Double, tx As Double, ty As Double, scenario As Long
    Set path = TestHelperCompoundPath
    TestNear "anchor.named.station", path.AnchorStation("Corner", 0.5), 100# + 12.5 * GEOM_PI
    Set part = path.BetweenAnchors("Chain", "Bottom", 0.25, "Corner", 0.5)
    TestCheck "anchor.chain.id", part.ID = "Chain" And Not part.IsClosed
    TestCheck "anchor.chain.count", part.SegmentCount = 2
    TestNear "anchor.chain.length", part.Length, 75# + 12.5 * GEOM_PI
    TestCheck "anchor.chain.source.path", part.SourcePathID(1) = "Outer.Mixed" And part.SourcePathID(2) = "Outer.Mixed"
    TestCheck "anchor.chain.source.segment", part.SourceSegmentID(1) = "Bottom" And part.SourceSegmentID(2) = "Corner"
    TestNear "anchor.chain.source.first", part.SourceStartParameter(1), 0.25
    TestNear "anchor.chain.source.last", part.SourceEndParameter(2), 0.5
    Set part = path.Slice(20#, 80#): Set changed = part.Reversed
    TestNear "anchor.reverse.source.first", changed.SourceStartParameter(1), 0.8
    TestNear "anchor.reverse.source.last", changed.SourceEndParameter(1), 0.2
    Set part = changed.Slice(10#, 50#)
    TestNear "anchor.nested.source.first", part.SourceStartParameter(1), 0.7
    TestNear "anchor.nested.source.last", part.SourceEndParameter(1), 0.3
    Set changed = part.Transformed(1000#, -2000#, GEOM_PI / 2#, True): Set part = changed.WithID("Moved")
    TestCheck "anchor.transform.id", part.ID = "Moved" And part.SegmentID(1) = "Bottom"
    TestCheck "anchor.transform.source", part.SourcePathID(1) = "Outer.Mixed" And part.SourceSegmentID(1) = "Bottom"
    TestNear "anchor.transform.source.first", part.SourceStartParameter(1), 0.7
    TestNear "anchor.transform.source.last", part.SourceEndParameter(1), 0.3
    part.AtDistance 0#, x, y, tx, ty
    TestNear "anchor.transform.x", x, 1000#: TestNear "anchor.transform.y", y, -1930#
    TestNear "anchor.source.unchanged", path.Length, 150# + 25# * GEOM_PI
    Set path = TestCirclePath
    Set part = path.BetweenAnchors("Wrap", "Circumference", 0.75, "Circumference", 0.25)
    TestCheck "anchor.wrap.count", part.SegmentCount = 2 And Not part.IsClosed
    TestCheck "anchor.wrap.unique.ids", part.SegmentID(1) <> part.SegmentID(2)
    TestCheck "anchor.wrap.source", part.SourceSegmentID(1) = "Circumference" And part.SourceSegmentID(2) = "Circumference"
    TestNear "anchor.wrap.length", part.Length, 10# * GEOM_PI
    TestNear "anchor.wrap.first.start", part.SourceStartParameter(1), 0.75
    TestNear "anchor.wrap.first.end", part.SourceEndParameter(1), 1#
    TestNear "anchor.wrap.second.start", part.SourceStartParameter(2), 0#
    TestNear "anchor.wrap.second.end", part.SourceEndParameter(2), 0.25
    part.AtDistance part.Length / 2#, x, y, tx, ty
    TestNear "anchor.wrap.no.chord.x", x, 10#: TestNear "anchor.wrap.no.chord.y", y, 0#
    Set part = path.BetweenAnchors("Reverse", "Circumference", 0.75, "Circumference", 0.25, False)
    TestCheck "anchor.reverse.count", part.SegmentCount = 1
    TestNear "anchor.reverse.length", part.Length, 10# * GEOM_PI
    TestNear "anchor.reverse.first", part.SourceStartParameter(1), 0.75
    TestNear "anchor.reverse.last", part.SourceEndParameter(1), 0.25
    part.AtDistance part.Length / 2#, x, y, tx, ty
    TestNear "anchor.reverse.other.side.x", x, -10#: TestNear "anchor.reverse.other.side.y", y, 0#
    part.CopySegment 1, edge: TestNear "anchor.reverse.sweep", edge.Sweep, -GEOM_PI
    Set part = path.BetweenAnchors("Loop", "Circumference", 0.25, "Circumference", 0.25, True, True)
    TestCheck "anchor.full.loop.closed", part.IsClosed And part.SegmentCount = 2
    TestNear "anchor.full.loop.length", part.Length, path.Length
    Set part = path.BetweenAnchors("Seam", "Circumference", 1#, "Circumference", 0#, True, True)
    TestCheck "anchor.seam.loop.closed", part.IsClosed And part.SegmentCount = 1
    TestNear "anchor.seam.loop.length", part.Length, path.Length
    Set part = path.BetweenAnchors("SeamReverse", "Circumference", 0#, "Circumference", 1#, False, True)
    TestCheck "anchor.seam.reverse.closed", part.IsClosed
    TestNear "anchor.seam.reverse.source.first", part.SourceStartParameter(1), 1#
    Set query = New CSectionGeometryQuery
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, 0#
    polygonSpec.AddVertex "BR", 100#, 0#, 0#
    polygonSpec.AddVertex "TR", 100#, 80#, 0#
    polygonSpec.AddVertex "TL", 0#, 80#, 0#
    polygonSpec.Seal
    Set path = query.RoundedPolygon("Sharp", polygonSpec)
    Set part = path.BetweenAnchors("NamedSide", "Edge.BR", 0#, "Edge.TR", 0#)
    TestCheck "anchor.zero.radius.named.edge", part.SegmentCount = 1 And part.SegmentID(1) = "Edge.BR"
    TestNear "anchor.zero.radius.length", part.Length, 80#
    TestCheck "anchor.zero.radius.no.phantom", Not path.HasSegmentID("Corner.BL") ' Нулевая дуга не является сегментом, но её границы сохраняют стык.
    TestNear "anchor.zero.radius.junction.start", path.AnchorStation("Corner.BL", 0#), path.AnchorStation("Edge.BL", 0#)
    TestNear "anchor.zero.radius.junction.end", path.AnchorStation("Corner.BL", 1#), path.AnchorStation("Edge.BL", 0#)
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "AnchorFixture"
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, 0#
    polygonSpec.AddVertex "BR", 100#, 0#, 0#
    polygonSpec.AddVertex "TR", 100#, 80#, 0#
    polygonSpec.AddVertex "TL", 0#, 80#, 0#
    polygonSpec.Seal
    Set path = facade.RoundedPolygon("Outer", polygonSpec)
    facade.Outer "Concrete", path.ID
    Set part = facade.Chain("Side", "Outer", "Edge.BR", 0#, "Edge.TR", 0#)
    Set changed = facade.Transform("MovedSide", "Side", 1000000#, -1000000#)
    Set definition = facade.Finish
    TestCheck "anchor.facade.registered", definition.PathCount = 3 And definition.IsSealed
    TestCheck "anchor.facade.source", changed.SourcePathID(1) = "Outer" And changed.SourceSegmentID(1) = "Edge.BR"
    changed.AtDistance 0#, x, y, tx, ty
    TestNear "anchor.facade.moved.x", x, 1000100#: TestNear "anchor.facade.moved.y", y, -1000000#
    For scenario = 0 To 9: TestCheck "anchor.invalid." & CStr(scenario), TestInvalidAnchorRejected(scenario): Next scenario
    GoTo Finished
Failed:
    TestCheck "anchor.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunPathAnchorTests = mReport & "TOTAL PATH ANCHORS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет адресный отказ при неизвестной или недопустимой привязке, ошибке направления и попытке получить нулевую дугу как сегмент.
Private Function TestInvalidAnchorRejected(ByVal scenario As Long) As Boolean
    Dim polygonSpec As CPolygonSpec, polygonPoint As Variant
    On Error GoTo Rejected
    Dim path As CGeometryPath, part As CGeometryPath, query As CSectionGeometryQuery, station As Double
    Set path = TestCirclePath
    Select Case scenario
        Case 0: station = path.AnchorStation("Missing")
        Case 1: station = path.AnchorStation("Circumference", -0.1)
        Case 2: station = path.AnchorStation("Circumference", 1.1)
        Case 3: Set part = path.BetweenAnchors("Same", "Circumference", 0.25, "Circumference", 0.25)
        Case 4: Set part = path.BetweenAnchors("Seam", "Circumference", 1#, "Circumference", 0#)
        Case 5: Set part = path.BetweenAnchors("Seam", "Circumference", 0#, "Circumference", 1#)
        Case 6
            Set path = TestHelperCompoundPath
            Set part = path.BetweenAnchors("WrongDirection", "Right", 0#, "Bottom", 0#)
        Case 7
            Set path = TestHelperCompoundPath
            Set part = path.BetweenAnchors("OpenLoop", "Bottom", 0.5, "Bottom", 0.5, True, True)
        Case 8
            Set query = New CSectionGeometryQuery
            Set polygonSpec = New CPolygonSpec
            polygonSpec.AddVertex "BL", 0#, 0#, 0#
            polygonSpec.AddVertex "BR", 100#, 0#, 0#
            polygonSpec.AddVertex "TR", 100#, 80#, 0#
            polygonSpec.AddVertex "TL", 0#, 80#, 0#
            polygonSpec.Seal
            Set path = query.RoundedPolygon("Sharp", polygonSpec)
            station = path.SegmentIndex("Corner.BL")
        Case 9: Set part = path.WithID("")
    End Select
    Exit Function
Rejected:
    TestInvalidAnchorRejected = (Err.Source = "CGeometryPath" And Err.Number = vbObjectError + 5520)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет прямые, дуги, составные пути, касательные в стыках, разбиение дуги и пересечения лучей, затем недопустимые входы.
Public Function RunPathTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    TestLine
    TestArcs
    TestCompoundPath
    TestJointTangents
    TestSegmentation
    TestRays
    Dim scenario As Long
    For scenario = 0 To 13
        TestCheck "path.invalid." & CStr(scenario), TestInvalidPathRejected(scenario)
    Next scenario
    GoTo Finished
Failed:
    TestCheck "path.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunPathTests = mReport & "TOTAL ANALYTIC PATHS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' общий алгоритм скругления, включая вогнутые вершины и нулевой радиус.
Public Function RunRoundedPolygonTests() As String
    Dim polygonSpec As CPolygonSpec, polygonPoint As Variant
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    Dim query As CSectionGeometryQuery, path As CGeometryPath, region As CConcreteRegion
    Dim edge As TRegionEdge, i As Long, radius As Double, scenario As Long
    Set query = New CSectionGeometryQuery
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, 10#
    polygonSpec.AddVertex "BR", 100#, 0#, 10#
    polygonSpec.AddVertex "TR", 100#, 80#, 10#
    polygonSpec.AddVertex "TL", 0#, 80#, 10#
    polygonSpec.Seal
    Set path = query.RoundedPolygon("Outer", polygonSpec)
    TestCheck "rounded.count", path.SegmentCount = 8
    TestCheck "rounded.closed", path.IsClosed
    TestNear "rounded.length", path.Length, 280# + 20# * GEOM_PI
    Set region = query.RegionFromPath(path)
    TestNear "rounded.area", region.Area, 8000# - 400# + 100# * GEOM_PI
    For i = 1 To path.SegmentCount
        path.CopySegment i, edge
        If edge.Sweep <> 0# Then
            radius = Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2)
            TestNear "rounded.radius." & CStr(i), radius, 10#
            TestNear "rounded.sweep." & CStr(i), edge.Sweep, GEOM_PI / 2#
        End If
    Next i
    TestCheck "rounded.named.edge", path.SegmentID(path.SegmentIndex("Edge.BR")) = "Edge.BR"
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, 0#
    polygonSpec.AddVertex "BR", 100#, 0#, 10#
    polygonSpec.AddVertex "TR", 100#, 80#, 20#
    polygonSpec.AddVertex "TL", 0#, 80#, 5#
    polygonSpec.Seal
    Set path = query.RoundedPolygon("Mixed", polygonSpec)
    TestCheck "rounded.zero.radius.count", path.SegmentCount = 7
    TestCheck "rounded.zero.radius.edge", path.SegmentIndex("Edge.BL") = 1
    Set region = query.RegionFromPath(path)
    TestNear "rounded.individual.area", region.Area, 8000# - 525# * (1# - GEOM_PI / 4#)
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, 0#
    polygonSpec.AddVertex "BR", 100#, 0#, 0#
    polygonSpec.AddVertex "TR", 100#, 80#, 0#
    polygonSpec.AddVertex "TL", 0#, 80#, 0#
    polygonSpec.Seal
    Set path = query.RoundedPolygon("Sharp", polygonSpec)
    TestNear "rounded.sharp.length", path.Length, 360#
    Set region = query.RegionFromPath(path)
    TestNear "rounded.sharp.area", region.Area, 8000#
    Set polygonSpec = New CPolygonSpec ' Вогнутый угол сохраняет противоположное направление своей дуги.
    polygonSpec.AddVertex "BL", 0#, 0#, 10#
    polygonSpec.AddVertex "BR", 200#, 0#, 10#
    polygonSpec.AddVertex "Arm", 200#, 100#, 10#
    polygonSpec.AddVertex "Inside", 100#, 100#, 10#
    polygonSpec.AddVertex "TR", 100#, 200#, 10#
    polygonSpec.AddVertex "TL", 0#, 200#, 10#
    polygonSpec.Seal
    Set path = query.RoundedPolygon("Concave", polygonSpec)
    path.CopySegment path.SegmentIndex("Corner.Inside"), edge
    TestNear "rounded.concave.sweep", edge.Sweep, -GEOM_PI / 2#
    TestNear "rounded.concave.cx", edge.CenterX, 110#: TestNear "rounded.concave.cy", edge.CenterY, 110#
    TestNear "rounded.concave.start.x", edge.X1, 110#: TestNear "rounded.concave.start.y", edge.Y1, 100#
    TestNear "rounded.concave.end.x", edge.X2, 100#: TestNear "rounded.concave.end.y", edge.Y2, 110#
    Set region = query.RegionFromPath(path)
    TestNear "rounded.concave.area", region.Area, 30000# - 400# * (1# - GEOM_PI / 4#)
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, 10#
    polygonSpec.AddVertex "TL", 0#, 80#, 10#
    polygonSpec.AddVertex "TR", 100#, 80#, 10#
    polygonSpec.AddVertex "BR", 100#, 0#, 10#
    polygonSpec.Seal
    Set path = query.RoundedPolygon("Clockwise", polygonSpec)
    path.CopySegment path.SegmentIndex("Corner.BL"), edge
    TestNear "rounded.clockwise.sweep", edge.Sweep, -GEOM_PI / 2#
    Set region = query.RegionFromPath(path)
    TestNear "rounded.clockwise.area", region.Area, 8000# - 400# + 100# * GEOM_PI
    Set polygonSpec = New CPolygonSpec ' Касание соседних дуг допустимо; исчезнувшая прямая не выбирается по индексу.
    polygonSpec.AddVertex "BL", 0#, 0#, 50#
    polygonSpec.AddVertex "BR", 100#, 0#, 50#
    polygonSpec.AddVertex "TR", 100#, 100#, 50#
    polygonSpec.AddVertex "TL", 0#, 100#, 50#
    polygonSpec.Seal
    Set path = query.RoundedPolygon("Touching", polygonSpec)
    TestCheck "rounded.touching.count", path.SegmentCount = 4
    Set region = query.RegionFromPath(path)
    TestNear "rounded.touching.area", region.Area, 2500# * GEOM_PI
    For scenario = 0 To 5
        TestCheck "rounded.invalid." & CStr(scenario), TestInvalidPolygonRejected(scenario)
    Next scenario
    GoTo Finished
Failed:
    TestCheck "rounded.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRoundedPolygonTests = mReport & "TOTAL ROUNDED POLYGONS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' точные offset-дуги, вогнутый угол и независимость от обхода.
Public Function RunRoundedOffsetTests() As String
    Dim polygonSpec As CPolygonSpec, polygonPoint As Variant
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    Dim query As CSectionGeometryQuery, path As CGeometryPath, region As CConcreteRegion
    Dim original As CGeometryPath, originalRegion As CConcreteRegion, edge As TRegionEdge
    Dim vertices As Variant, radii As Variant, names As Variant, scenario As Long, x As Double, y As Double, tx As Double, ty As Double
    Set query = New CSectionGeometryQuery
    radii = Array(10#, 10#, 10#, 10#)
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, CDbl(radii(LBound(radii) + 0))
    polygonSpec.AddVertex "BR", 100#, 0#, CDbl(radii(LBound(radii) + 1))
    polygonSpec.AddVertex "TR", 100#, 80#, CDbl(radii(LBound(radii) + 2))
    polygonSpec.AddVertex "TL", 0#, 80#, CDbl(radii(LBound(radii) + 3))
    polygonSpec.Seal
    Set original = query.RoundedPolygon("Base", polygonSpec)
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, CDbl(radii(LBound(radii) + 0))
    polygonSpec.AddVertex "BR", 100#, 0#, CDbl(radii(LBound(radii) + 1))
    polygonSpec.AddVertex "TR", 100#, 80#, CDbl(radii(LBound(radii) + 2))
    polygonSpec.AddVertex "TL", 0#, 80#, CDbl(radii(LBound(radii) + 3))
    polygonSpec.Seal
    Set path = query.OffsetRoundedPolygon("Inward", polygonSpec, 5#)
    TestCheck "offset.closed", path.IsClosed And path.SegmentCount = 8
    TestNear "offset.rectangle.length", path.Length, 280# + 10# * GEOM_PI
    Set region = query.RegionFromPath(path)
    TestNear "offset.rectangle.area", region.Area, 6300# - 100# + 25# * GEOM_PI
    path.CopySegment path.SegmentIndex("Corner.BL"), edge
    TestNear "offset.rectangle.radius", Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2), 5#
    TestNear "offset.rectangle.center.x", edge.CenterX, 10#: TestNear "offset.rectangle.center.y", edge.CenterY, 10#
    TestNear "offset.original.unchanged", original.Length, 280# + 20# * GEOM_PI
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, CDbl(radii(LBound(radii) + 0))
    polygonSpec.AddVertex "BR", 100#, 0#, CDbl(radii(LBound(radii) + 1))
    polygonSpec.AddVertex "TR", 100#, 80#, CDbl(radii(LBound(radii) + 2))
    polygonSpec.AddVertex "TL", 0#, 80#, CDbl(radii(LBound(radii) + 3))
    polygonSpec.Seal
    Set path = query.OffsetRoundedPolygon("Outward", polygonSpec, -5#)
    Set region = query.RegionFromPath(path)
    TestNear "offset.outward.area", region.Area, 9900# - 900# + 225# * GEOM_PI
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, CDbl(radii(LBound(radii) + 0))
    polygonSpec.AddVertex "BR", 100#, 0#, CDbl(radii(LBound(radii) + 1))
    polygonSpec.AddVertex "TR", 100#, 80#, CDbl(radii(LBound(radii) + 2))
    polygonSpec.AddVertex "TL", 0#, 80#, CDbl(radii(LBound(radii) + 3))
    polygonSpec.Seal
    Set path = query.OffsetRoundedPolygon("VanishedArc", polygonSpec, 15#)
    TestCheck "offset.vanished.arc.count", path.SegmentCount = 4
    TestCheck "offset.vanished.arc.named.edge", path.SegmentID(path.SegmentIndex("Edge.BL")) = "Edge.BL"
    Set region = query.RegionFromPath(path)
    TestNear "offset.vanished.arc.area", region.Area, 3500#
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, 0#
    polygonSpec.AddVertex "BR", 100#, 0#, 0#
    polygonSpec.AddVertex "TR", 100#, 80#, 0#
    polygonSpec.AddVertex "TL", 0#, 80#, 0#
    polygonSpec.Seal
    Set path = query.OffsetRoundedPolygon("SharpOutward", polygonSpec, -5#)
    TestCheck "offset.sharp.preserved", path.SegmentCount = 4
    Set region = query.RegionFromPath(path): TestNear "offset.sharp.area", region.Area, 9900#
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, CDbl(radii(LBound(radii) + 0))
    polygonSpec.AddVertex "TL", 0#, 80#, CDbl(radii(LBound(radii) + 1))
    polygonSpec.AddVertex "TR", 100#, 80#, CDbl(radii(LBound(radii) + 2))
    polygonSpec.AddVertex "BR", 100#, 0#, CDbl(radii(LBound(radii) + 3))
    polygonSpec.Seal
    Set path = query.OffsetRoundedPolygon("Clockwise", polygonSpec, 5#)
    Set region = query.RegionFromPath(path): TestNear "offset.clockwise.area", region.Area, 6300# - 100# + 25# * GEOM_PI
    path.CopySegment path.SegmentIndex("Corner.BL"), edge
    TestNear "offset.clockwise.sweep", edge.Sweep, -GEOM_PI / 2#
    vertices = Array(Array(0#, 0#), Array(100#, 0#), Array(100#, 40#), Array(40#, 40#), Array(40#, 100#), Array(0#, 100#))
    radii = Array(5#, 5#, 5#, 5#, 5#, 5#): names = Array("BL", "BR", "Arm", "Inside", "TR", "TL")
    Set polygonSpec = New CPolygonSpec
    polygonPoint = vertices(LBound(vertices) + 0)
    polygonSpec.AddVertex CStr(names(LBound(names) + 0)), CDbl(polygonPoint(LBound(polygonPoint))), CDbl(polygonPoint(LBound(polygonPoint) + 1)), CDbl(radii(LBound(radii) + 0))
    polygonPoint = vertices(LBound(vertices) + 1)
    polygonSpec.AddVertex CStr(names(LBound(names) + 1)), CDbl(polygonPoint(LBound(polygonPoint))), CDbl(polygonPoint(LBound(polygonPoint) + 1)), CDbl(radii(LBound(radii) + 1))
    polygonPoint = vertices(LBound(vertices) + 2)
    polygonSpec.AddVertex CStr(names(LBound(names) + 2)), CDbl(polygonPoint(LBound(polygonPoint))), CDbl(polygonPoint(LBound(polygonPoint) + 1)), CDbl(radii(LBound(radii) + 2))
    polygonPoint = vertices(LBound(vertices) + 3)
    polygonSpec.AddVertex CStr(names(LBound(names) + 3)), CDbl(polygonPoint(LBound(polygonPoint))), CDbl(polygonPoint(LBound(polygonPoint) + 1)), CDbl(radii(LBound(radii) + 3))
    polygonPoint = vertices(LBound(vertices) + 4)
    polygonSpec.AddVertex CStr(names(LBound(names) + 4)), CDbl(polygonPoint(LBound(polygonPoint))), CDbl(polygonPoint(LBound(polygonPoint) + 1)), CDbl(radii(LBound(radii) + 4))
    polygonPoint = vertices(LBound(vertices) + 5)
    polygonSpec.AddVertex CStr(names(LBound(names) + 5)), CDbl(polygonPoint(LBound(polygonPoint))), CDbl(polygonPoint(LBound(polygonPoint) + 1)), CDbl(radii(LBound(radii) + 5))
    polygonSpec.Seal
    Set original = query.RoundedPolygon("ConcaveBase", polygonSpec)
    Set originalRegion = query.RegionFromPath(original)
    Set polygonSpec = New CPolygonSpec
    polygonPoint = vertices(LBound(vertices) + 0)
    polygonSpec.AddVertex CStr(names(LBound(names) + 0)), CDbl(polygonPoint(LBound(polygonPoint))), CDbl(polygonPoint(LBound(polygonPoint) + 1)), CDbl(radii(LBound(radii) + 0))
    polygonPoint = vertices(LBound(vertices) + 1)
    polygonSpec.AddVertex CStr(names(LBound(names) + 1)), CDbl(polygonPoint(LBound(polygonPoint))), CDbl(polygonPoint(LBound(polygonPoint) + 1)), CDbl(radii(LBound(radii) + 1))
    polygonPoint = vertices(LBound(vertices) + 2)
    polygonSpec.AddVertex CStr(names(LBound(names) + 2)), CDbl(polygonPoint(LBound(polygonPoint))), CDbl(polygonPoint(LBound(polygonPoint) + 1)), CDbl(radii(LBound(radii) + 2))
    polygonPoint = vertices(LBound(vertices) + 3)
    polygonSpec.AddVertex CStr(names(LBound(names) + 3)), CDbl(polygonPoint(LBound(polygonPoint))), CDbl(polygonPoint(LBound(polygonPoint) + 1)), CDbl(radii(LBound(radii) + 3))
    polygonPoint = vertices(LBound(vertices) + 4)
    polygonSpec.AddVertex CStr(names(LBound(names) + 4)), CDbl(polygonPoint(LBound(polygonPoint))), CDbl(polygonPoint(LBound(polygonPoint) + 1)), CDbl(radii(LBound(radii) + 4))
    polygonPoint = vertices(LBound(vertices) + 5)
    polygonSpec.AddVertex CStr(names(LBound(names) + 5)), CDbl(polygonPoint(LBound(polygonPoint))), CDbl(polygonPoint(LBound(polygonPoint) + 1)), CDbl(radii(LBound(radii) + 5))
    polygonSpec.Seal
    Set path = query.OffsetRoundedPolygon("ConcaveOffset", polygonSpec, 2#)
    path.CopySegment path.SegmentIndex("Corner.Inside"), edge
    TestNear "offset.concave.radius", Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2), 7#
    TestNear "offset.concave.center.x", edge.CenterX, 45#: TestNear "offset.concave.center.y", edge.CenterY, 45#
    TestNear "offset.concave.sweep", edge.Sweep, -GEOM_PI / 2#
    Set region = query.RegionFromPath(path): TestNear "offset.concave.area", region.Area, 5616# + 4# - GEOM_PI
    path.AtDistance path.Length / 7#, x, y, tx, ty
    TestNear "offset.boundary.distance", query.BoundaryDistance(originalRegion, x, y), 2#
    vertices = Array(Array(1000000#, -1000000#), Array(1000100#, -1000000#), _
        Array(1000100#, -999920#), Array(1000000#, -999920#))
    Set polygonSpec = New CPolygonSpec
    polygonPoint = vertices(LBound(vertices) + 0)
    polygonSpec.AddVertex "BL", CDbl(polygonPoint(LBound(polygonPoint))), CDbl(polygonPoint(LBound(polygonPoint) + 1)), 10#
    polygonPoint = vertices(LBound(vertices) + 1)
    polygonSpec.AddVertex "BR", CDbl(polygonPoint(LBound(polygonPoint))), CDbl(polygonPoint(LBound(polygonPoint) + 1)), 10#
    polygonPoint = vertices(LBound(vertices) + 2)
    polygonSpec.AddVertex "TR", CDbl(polygonPoint(LBound(polygonPoint))), CDbl(polygonPoint(LBound(polygonPoint) + 1)), 10#
    polygonPoint = vertices(LBound(vertices) + 3)
    polygonSpec.AddVertex "TL", CDbl(polygonPoint(LBound(polygonPoint))), CDbl(polygonPoint(LBound(polygonPoint) + 1)), 10#
    polygonSpec.Seal
    Set path = query.OffsetRoundedPolygon("Far", polygonSpec, 5#)
    Set region = query.RegionFromPath(path): TestNear "offset.far.area", region.Area, 6300# - 100# + 25# * GEOM_PI
    path.CopySegment path.SegmentIndex("Corner.BL"), edge
    TestNear "offset.far.center.x", edge.CenterX, 1000010#: TestNear "offset.far.center.y", edge.CenterY, -999990#
    For scenario = 0 To 2: TestCheck "offset.invalid." & CStr(scenario), TestInvalidOffsetRejected(scenario): Next scenario
    GoTo Finished
Failed:
    TestCheck "offset.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRoundedOffsetTests = mReport & "TOTAL ROUNDED OFFSET: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет отказ общего смещения при недопустимой геометрии или исчезновении области; успешная пустая подстановка не допускается.
Private Function TestInvalidOffsetRejected(ByVal scenario As Long) As Boolean
    Dim polygonSpec As CPolygonSpec, polygonPoint As Variant
    On Error GoTo Rejected
    Dim query As CSectionGeometryQuery, path As CGeometryPath
    Set query = New CSectionGeometryQuery
    Select Case scenario
        Case 0
            Set polygonSpec = New CPolygonSpec
            polygonSpec.AddVertex "BL", 0#, 0#, 10#
            polygonSpec.AddVertex "BR", 100#, 0#, 10#
            polygonSpec.AddVertex "TR", 100#, 80#, 10#
            polygonSpec.AddVertex "TL", 0#, 80#, 10#
            polygonSpec.Seal
            Set path = query.OffsetRoundedPolygon("Collapsed", polygonSpec, 50#)
        Case 1
            Set polygonSpec = New CPolygonSpec
            polygonSpec.AddVertex "BL", 0#, 0#, 10#
            polygonSpec.AddVertex "BR", 100#, 0#, 10#
            polygonSpec.AddVertex "TR", 100#, 80#, 10#
            polygonSpec.AddVertex "TL", 0#, 80#, 10#
            polygonSpec.Seal
            Set path = query.OffsetRoundedPolygon("Reversed", polygonSpec, 100#)
        Case 2
            Set polygonSpec = New CPolygonSpec
            polygonSpec.AddVertex "BL", 0#, 0#, 60#
            polygonSpec.AddVertex "BR", 100#, 0#, 60#
            polygonSpec.AddVertex "TR", 100#, 80#, 60#
            polygonSpec.AddVertex "TL", 0#, 80#, 60#
            polygonSpec.Seal
            Set path = query.OffsetRoundedPolygon("InvalidBase", polygonSpec, 50#)
    End Select
    Exit Function
Rejected:
    TestInvalidOffsetRejected = (Err.Source = "CSectionGeometryQuery" And Err.Number = vbObjectError + 5401)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает вершины контрольного прямоугольника в постоянном порядке обхода.
Private Function TestRectangleVertices() As Variant
    TestRectangleVertices = Array(Array(0#, 0#), Array(100#, 0#), Array(100#, 80#), Array(0#, 80#))
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает устойчивые имена углов контрольного прямоугольника для проверки источников сегментов.
Private Function TestRectangleNames() As Variant
    TestRectangleNames = Array("BL", "BR", "TR", "TL")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' чтение in-memory ввода и независимость от исходных массивов.
Public Function RunInputTableTests() As String
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    Dim cells(1 To 4, 1 To 3) As Variant, addresses(1 To 4, 1 To 3) As String, row As Long, column As Long
    Dim table As CSectionInputTable, units As CUnitSystem, scenario As Long
    cells(1, 1) = "Key": cells(1, 2) = "Value": cells(1, 3) = "Mode"
    cells(2, 1) = "Count": cells(2, 2) = 4#: cells(2, 3) = "stacked"
    cells(3, 1) = "Disabled": cells(3, 2) = 0#
    cells(4, 1) = "Diameter": cells(4, 2) = 16#
    For row = 1 To 4
        For column = 1 To 3
            addresses(row, column) = "Config!" & Chr$(64 + column) & CStr(row + 10)
        Next column
    Next row
    Set table = New CSectionInputTable: table.Initialize "rngFixture", cells, addresses
    Set units = New CUnitSystem: units.InitializeDefaults
    TestCheck "input.rows", table.RowCount = 4
    TestCheck "input.columns", table.ColumnCount = 3
    TestCheck "input.caption", table.ColumnByCaption(1, "value") = 2
    TestCheck "input.key", table.RowByKey("diameter") = 4
    TestCheck "input.key.scoped", table.RowByKey("Count", 1, 2, 3) = 2
    TestCheck "input.key.alias", table.RowByAnyKey(Array("Unknown", "Diameter"), 1, 3, 4) = 4
    TestNear "input.number", table.Number(2, 2), 4#
    TestCheck "input.count", table.Count(2, 2) = 4
    TestCheck "input.choice", table.Choice(2, 3, Array("Stacked", "SideBySide")) = "Stacked"
    TestCheck "input.choice.default", table.Choice(4, 3, Array("Stacked", "SideBySide"), "SideBySide") = "SideBySide"
    TestCheck "input.source", table.SourceReference(2, 2) = "Config!B12"
    TestCheck "input.blank", table.IsBlank(4, 3)
    TestNear "input.number.default", table.Number(4, 3, 7#), 7#
    TestNear "input.optional.zero", table.OptionalDiameter(3, 2, units), 0#
    TestNear "input.optional.blank", table.OptionalDiameter(3, 3, units), 0#
    TestNear "input.optional.active", table.OptionalDiameter(4, 2, units), 16#
    TestNear "input.length", table.Length(4, 2, units), 16#
    cells(2, 2) = 999#: addresses(2, 2) = "Changed"
    TestNear "input.value.immutable", table.Number(2, 2), 4#
    TestCheck "input.source.immutable", table.SourceReference(2, 2) = "Config!B12"
    cells(2, 2) = 2147483647#: addresses(2, 2) = "Config!B12"
    Set table = New CSectionInputTable: table.Initialize "rngLongBoundary", cells, addresses
    TestCheck "input.count.maxLong", table.Count(2, 2) = 2147483647
    For scenario = 0 To 24
        TestCheck "input.invalid." & CStr(scenario), TestInvalidInputRejected(scenario)
    Next scenario
    GoTo Finished
Failed:
    TestCheck "input.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunInputTableTests = mReport & "TOTAL INPUT TABLES: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет типы, диапазоны, повторные ключи и границы индексов общей таблицы ввода с сохранением владельца ошибки.
Private Function TestInvalidInputRejected(ByVal scenario As Long) As Boolean
    On Error GoTo Rejected
    Dim cells(1 To 2, 1 To 2) As Variant, table As CSectionInputTable, value As Variant, units As CUnitSystem
    cells(1, 1) = "Key": cells(1, 2) = "Value": cells(2, 1) = "Count": cells(2, 2) = 4#
    Select Case scenario
        Case 0: cells(2, 2) = "four"
        Case 1: cells(2, 2) = -1#
        Case 2: cells(2, 2) = 1.5
        Case 3: cells(2, 2) = 2147483648#
        Case 4: cells(2, 2) = True
        Case 5: cells(2, 2) = CVErr(2042)
        Case 6: cells(1, 1) = "Value"
        Case 7: cells(1, 1) = "Count"
        Case 8: cells(2, 2) = Empty
        Case 9: cells(2, 2) = "Unknown"
        Case 13: cells(2, 2) = -1#
        Case 14: cells(2, 2) = Array(1#, 2#)
        Case 15: Set cells(2, 2) = New CUnitSystem
        Case 22: cells(2, 2) = "-"
        Case 23: cells(2, 2) = "TODO"
        Case 24: cells(2, 2) = "diameter"
    End Select
    Set table = New CSectionInputTable: table.Initialize "rngInvalid", cells
    Set units = New CUnitSystem: units.InitializeDefaults
    Select Case scenario
        Case 0, 4, 5, 8: value = table.Number(2, 2)
        Case 1, 2, 3: value = table.Count(2, 2)
        Case 6: value = table.ColumnByCaption(1, "Value")
        Case 7: value = table.RowByKey("Count")
        Case 9: value = table.Choice(2, 2, Array("Stacked", "SideBySide"))
        Case 10: table.Initialize "Second", cells
        Case 11: value = table.Number(3, 2)
        Case 12: value = table.RowByKey("Missing")
        Case 13, 22, 23, 24: value = table.OptionalDiameter(2, 2, units)
        Case 16: value = table.ColumnByCaption(1, "Missing")
        Case 17: value = table.RowByKey("Count", 1, 2, 1)
        Case 18: value = table.RowByKey("Count", 1, 3, 4)
        Case 19: value = table.RowByAnyKey(Array("Key", "Count"))
        Case 20: value = table.RowByAnyKey("Count")
        Case 21: value = table.RowByAnyKey(Array(""))
    End Select
    Exit Function
Rejected:
    TestInvalidInputRejected = (Err.Source = "CSectionInputTable" And Err.Number = vbObjectError + 5525)
    If scenario = 0 Then TestInvalidInputRejected = TestInvalidInputRejected And InStr(1, Err.Description, "rngInvalid[2,2]", vbBinaryCompare) > 0
    If scenario >= 22 Then TestInvalidInputRejected = TestInvalidInputRejected And InStr(1, Err.Description, "rngInvalid[2,2]", vbBinaryCompare) > 0
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Разделяет ошибки структуры описателя и геометрические ошибки скругления, проверяя код и источник отказа.
Private Function TestInvalidPolygonRejected(ByVal scenario As Long) As Boolean
    Dim polygonSpec As CPolygonSpec, polygonPoint As Variant
    On Error GoTo Rejected
    Dim query As CSectionGeometryQuery, path As CGeometryPath
    Set query = New CSectionGeometryQuery
    Select Case scenario
        Case 0
            Set polygonSpec = New CPolygonSpec
            polygonSpec.AddVertex "BL", 0#, 0#, -1#
            polygonSpec.AddVertex "BR", 100#, 0#, 0#
            polygonSpec.AddVertex "TR", 100#, 80#, 0#
            polygonSpec.AddVertex "TL", 0#, 80#, 0#
            polygonSpec.Seal
            Set path = query.RoundedPolygon("Negative", polygonSpec)
        Case 1
            Set polygonSpec = New CPolygonSpec
            polygonSpec.AddVertex "BL", 0#, 0#, 60#
            polygonSpec.AddVertex "BR", 100#, 0#, 60#
            polygonSpec.AddVertex "TR", 100#, 80#, 60#
            polygonSpec.AddVertex "TL", 0#, 80#, 60#
            polygonSpec.Seal
            Set path = query.RoundedPolygon("Overlap", polygonSpec)
        Case 2
            Set polygonSpec = New CPolygonSpec
            polygonSpec.AddVertex "A", 0#, 0#, 0#
            polygonSpec.AddVertex "A", 100#, 0#, 0#
            polygonSpec.AddVertex "B", 100#, 80#, 0#
            polygonSpec.AddVertex "C", 0#, 80#, 0#
            polygonSpec.Seal
            Set path = query.RoundedPolygon("Duplicate", polygonSpec)
        Case 3
            Set polygonSpec = New CPolygonSpec
            polygonSpec.AddVertex "A", 0#, 0#, 0#
            polygonSpec.AddVertex "B", 50#, 0#, 10#
            polygonSpec.AddVertex "C", 100#, 0#, 0#
            polygonSpec.AddVertex "D", 100#, 80#, 0#
            polygonSpec.AddVertex "E", 0#, 80#, 0#
            polygonSpec.Seal
            Set path = query.RoundedPolygon("Collinear", polygonSpec)
        Case 4
            Set polygonSpec = New CPolygonSpec
            polygonSpec.AddVertex "BL", 0#, 0#, 0#
            polygonSpec.AddVertex "BR", 100#, 0#, 0#
            polygonSpec.Seal
            Set path = query.RoundedPolygon("Missing", polygonSpec)
        Case 5
            Set polygonSpec = New CPolygonSpec
            polygonSpec.AddVertex "BL", 0#, 0#, 0#
            polygonSpec.AddVertex "BR", 100#, 80#, 0#
            polygonSpec.AddVertex "TR", 0#, 80#, 0#
            polygonSpec.AddVertex "TL", 100#, 0#, 0#
            polygonSpec.Seal
            Set path = query.RoundedPolygon("Crossed", polygonSpec)
    End Select
    Exit Function
Rejected:
    If scenario = 0 Or scenario = 2 Or scenario = 4 Then
        TestInvalidPolygonRejected = (Err.Source = "CPolygonSpec" And Err.Number = vbObjectError + 5470) ' Структура теперь проверяется описателем до передачи геометрическому ядру.
    Else
        TestInvalidPolygonRejected = (Err.Source = "CSectionGeometryQuery" And Err.Number = vbObjectError + 5401)
    End If
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет отрезок 3–4–5: длину, единичную касательную, проекции, обращение и поворот без изменения исходника.
Private Sub TestLine()
    Dim path As CGeometryPath, changed As CGeometryPath
    Dim x As Double, y As Double, tx As Double, ty As Double, minimum As Double, maximum As Double
    Set path = New CGeometryPath: path.Initialize "Outer.Inclined"
    path.AddLine "Inclined", 0#, 0#, 3#, 4#: path.Seal
    TestNear "line.length", path.Length, 5#
    path.AtDistance 2.5, x, y, tx, ty
    TestNear "line.x", x, 1.5: TestNear "line.y", y, 2#
    TestNear "line.tx", tx, 0.6: TestNear "line.ty", ty, 0.8
    path.ProjectionBounds 2#, 0#, minimum, maximum
    TestNear "line.minimum", minimum, 0#: TestNear "line.maximum", maximum, 3#
    Set changed = path.Reversed
    changed.AtDistance 0#, x, y, tx, ty
    TestNear "line.reverse.x", x, 3#: TestNear "line.reverse.y", y, 4#
    TestNear "line.reverse.tx", tx, -0.6: TestNear "line.reverse.ty", ty, -0.8
    Set changed = path.Transformed(10#, 20#, GEOM_PI / 2#)
    changed.AtDistance changed.Length, x, y, tx, ty
    TestNear "line.rotate.x", x, 6#: TestNear "line.rotate.y", y, 23#
    TestNear "line.rotate.tx", tx, -0.8: TestNear "line.rotate.ty", ty, 0.6
    TestCheck "line.transform.id", changed.SegmentID(1) = "Inclined"
    TestNear "line.transform.length", changed.Length, path.Length
    TestCheck "line.sealed", changed.IsSealed And path.IsSealed
    TestCheck "line.open", Not changed.IsClosed
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет аналитические точки, касательные и экстремумы дуг обоих направлений, отражение и замыкание полной окружности.
Private Sub TestArcs()
    Dim path As CGeometryPath, changed As CGeometryPath, edge As TRegionEdge
    Dim x As Double, y As Double, tx As Double, ty As Double, minimum As Double, maximum As Double
    Set path = New CGeometryPath: path.Initialize "Arc.Positive"
    path.AddArc "Curve", 0#, 0#, 10#, GEOM_PI / 4#, GEOM_PI: path.Seal
    TestNear "arc.length", path.Length, 10# * GEOM_PI
    path.AtDistance path.Length / 4#, x, y, tx, ty
    TestNear "arc.internal.x", x, 0#: TestNear "arc.internal.y", y, 10#
    TestNear "arc.internal.tx", tx, -1#: TestNear "arc.internal.ty", ty, 0#
    path.ProjectionBounds 0#, 1#, minimum, maximum
    TestNear "arc.bounds.y.min", minimum, -10# / Sqr(2#)
    TestNear "arc.bounds.y.max", maximum, 10#
    path.ProjectionBounds 1#, 0#, minimum, maximum
    TestNear "arc.bounds.x.min", minimum, -10#
    TestNear "arc.bounds.x.max", maximum, 10# / Sqr(2#)
    Set changed = path.Reversed
    changed.AtDistance changed.Length, x, y, tx, ty
    TestNear "arc.reverse.end.x", x, 10# / Sqr(2#)
    TestNear "arc.reverse.end.y", y, 10# / Sqr(2#)
    TestNear "arc.reverse.end.tx", tx, 1# / Sqr(2#)
    TestNear "arc.reverse.end.ty", ty, -1# / Sqr(2#)
    Set changed = path.Transformed(1000000#, -1000000#, 0#, True)
    changed.CopySegment 1, edge
    TestNear "arc.reflection.sweep", edge.Sweep, -GEOM_PI
    changed.AtDistance changed.Length / 4#, x, y, tx, ty
    TestNear "arc.reflection.x", x, 1000000#
    TestNear "arc.reflection.y", y, -1000010#
    TestNear "arc.reflection.length", changed.Length, path.Length
    TestCheck "arc.reflection.id", changed.SegmentID(1) = path.SegmentID(1)
    Set path = New CGeometryPath: path.Initialize "Arc.Negative"
    path.AddArc "Curve", 0#, 0#, 10#, GEOM_PI / 4#, -GEOM_PI: path.Seal
    path.AtDistance path.Length / 4#, x, y, tx, ty
    TestNear "arc.negative.x", x, 10#: TestNear "arc.negative.y", y, 0#
    TestNear "arc.negative.tx", tx, 0#: TestNear "arc.negative.ty", ty, -1#
    path.ProjectionBounds 1#, 0#, minimum, maximum
    TestNear "arc.negative.max", maximum, 10#
    TestNear "arc.negative.min", minimum, -10# / Sqr(2#)
    Set path = TestCirclePath
    TestCheck "circle.closed", path.IsClosed
    TestNear "circle.length", path.Length, 20# * GEOM_PI
    path.AtDistance 0#, x, y, tx, ty
    TestNear "circle.start.x", x, 10#: TestNear "circle.start.y", y, 0#
    path.AtDistance path.Length, x, y, tx, ty
    TestNear "circle.end.x", x, 10#: TestNear "circle.end.y", y, 0#
    path.ProjectionBounds 0#, 1#, minimum, maximum
    TestNear "circle.min", minimum, -10#: TestNear "circle.max", maximum, 10#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет путь из двух прямых и четверти окружности, его срез и преобразование с сохранением исходной геометрии.
Private Sub TestCompoundPath()
    Dim path As CGeometryPath, part As CGeometryPath
    Dim x As Double, y As Double, tx As Double, ty As Double
    Set path = TestHelperCompoundPath
    TestNear "compound.length", path.Length, 150# + 25# * GEOM_PI
    TestCheck "compound.count", path.SegmentCount = 3
    path.AtDistance 100# + 12.5 * GEOM_PI, x, y, tx, ty
    TestNear "compound.arc.x", x, 100# + 50# / Sqr(2#)
    TestNear "compound.arc.y", y, 50# - 50# / Sqr(2#)
    TestNear "compound.arc.tx", tx, 1# / Sqr(2#)
    TestNear "compound.arc.ty", ty, 1# / Sqr(2#)
    path.AtDistance 100#, x, y, tx, ty
    TestNear "compound.join.x", x, 100#: TestNear "compound.join.y", y, 0#
    TestNear "compound.join.tx", tx, 1#: TestNear "compound.join.ty", ty, 0#
    Set part = path.Slice(50#, path.Length - 25#)
    TestNear "compound.slice.length", part.Length, path.Length - 75#
    TestCheck "compound.slice.segments", part.SegmentCount = 3
    TestCheck "compound.slice.id", part.SegmentID(2) = "Corner"
    part.AtDistance 0#, x, y, tx, ty
    TestNear "compound.slice.start.x", x, 50#: TestNear "compound.slice.start.y", y, 0#
    part.AtDistance part.Length, x, y, tx, ty
    TestNear "compound.slice.end.x", x, 150#: TestNear "compound.slice.end.y", y, 75#
    TestNear "compound.source.unchanged", path.Length, 150# + 25# * GEOM_PI
    Set part = path.Transformed(1000#, -2000#, GEOM_PI / 2#, True)
    part.AtDistance part.Length, x, y, tx, ty
    TestNear "compound.transform.x", x, 1100#: TestNear "compound.transform.y", y, -1850#
    TestNear "compound.transform.tx", tx, 1#: TestNear "compound.transform.ty", ty, 0#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет выбор входящей касательной в точном углу и при малом шуме координаты стыка.
Private Sub TestJointTangents()
    Dim path As CGeometryPath, x As Double, y As Double, tx As Double, ty As Double
    Set path = New CGeometryPath: path.Initialize "Joint"
    path.AddLine "Incoming", 0#, 0#, 100#, 0#
    path.AddLine "Outgoing", 100#, 0#, 100#, 100#
    path.Seal
    path.AtDistance 100#, x, y, tx, ty
    TestNear "joint.default.tx", tx, 0#: TestNear "joint.default.ty", ty, 1#
    path.AtDistance 100#, x, y, tx, ty, True
    TestNear "joint.incoming.x", x, 100#: TestNear "joint.incoming.y", y, 0#
    TestNear "joint.incoming.tx", tx, 1#: TestNear "joint.incoming.ty", ty, 0#
    path.AtDistance 100# + GEOM_TOLERANCE / 2#, x, y, tx, ty, True
    TestNear "joint.roundoff.x", x, 100#: TestNear "joint.roundoff.y", y, 0#
    path.AtDistance 100# + 2# * GEOM_TOLERANCE, x, y, tx, ty, True
    TestNear "joint.after.tx", tx, 0#: TestNear "joint.after.ty", ty, 1#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает одну дугу с её разбиением на 25 частей: длина, точки и касательные должны совпадать.
Private Sub TestSegmentation()
    Dim whole As CGeometryPath, split As CGeometryPath, i As Long, startAngle As Double, sweep As Double
    Dim x1 As Double, y1 As Double, tx1 As Double, ty1 As Double, x2 As Double, y2 As Double, tx2 As Double, ty2 As Double
    startAngle = 7# * GEOM_PI / 4#: sweep = 5# * GEOM_PI / 4#
    Set whole = New CGeometryPath: whole.Initialize "Whole"
    whole.AddArc "Curve", 25#, -40#, 100#, startAngle, sweep: whole.Seal
    Set split = New CGeometryPath: split.Initialize "Split"
    For i = 0 To 24
        split.AddArc "Curve." & CStr(i), 25#, -40#, 100#, startAngle + sweep * CDbl(i) / 25#, sweep / 25#
    Next i
    split.Seal
    TestNear "split.length", split.Length, whole.Length
    For i = 0 To 40
        whole.AtDistance whole.Length * CDbl(i) / 40#, x1, y1, tx1, ty1
        split.AtDistance split.Length * CDbl(i) / 40#, x2, y2, tx2, ty2
        TestNear "split.x." & CStr(i), x2, x1
        TestNear "split.y." & CStr(i), y2, y1
        TestNear "split.tx." & CStr(i), tx2, tx1
        TestNear "split.ty." & CStr(i), ty2, ty1
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет ближайшее попадание луча на конечную прямую или дугу, касание и отсутствие попадания вне участка.
Private Sub TestRays()
    Dim path As CGeometryPath, station As Double, distance As Double
    Set path = TestCirclePath
    TestCheck "ray.circle.hit", path.RayHit(-20#, 0#, 2#, 0#, station, distance)
    TestNear "ray.circle.nearest", distance, 10#
    TestNear "ray.circle.station", station, 10# * GEOM_PI
    TestCheck "ray.circle.away", Not path.RayHit(-20#, 0#, -1#, 0#, station, distance)
    TestCheck "ray.circle.tangent", path.RayHit(-20#, 10#, 1#, 0#, station, distance)
    TestNear "ray.circle.tangent.distance", distance, 20#
    Set path = TestHelperCompoundPath
    TestCheck "ray.compound.hit", path.RayHit(200#, 75#, -1#, 0#, station, distance)
    TestNear "ray.compound.distance", distance, 50#
    TestNear "ray.compound.station", station, 125# + 25# * GEOM_PI
    Set path = New CGeometryPath: path.Initialize "Limited"
    path.AddLine "Line", 0#, 0#, 3#, 4#: path.Seal
    TestCheck "ray.line.beyond.target", Not path.RayHit(10#, 10#, 0#, -1#, station, distance)
    TestCheck "ray.line.collinear", path.RayHit(1.5, 2#, 0.6, 0.8, station, distance)
    TestNear "ray.line.collinear.distance", distance, 0#
    TestNear "ray.line.collinear.station", station, 2.5
    Set path = New CGeometryPath: path.Initialize "LimitedArc"
    path.AddArc "Curve", 0#, 0#, 10#, GEOM_PI / 4#, GEOM_PI: path.Seal
    TestCheck "ray.arc.beyond.target", Not path.RayHit(0#, 0#, 1#, 0#, station, distance)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт полную окружность радиусом 10 мм с началом на положительной полуоси X для проверки общего API пути.
Private Function TestCirclePath() As CGeometryPath
    Dim path As CGeometryPath
    Set path = New CGeometryPath: path.Initialize "Circle"
    path.AddArc "Circumference", 0#, 0#, 10#, 0#, 2# * GEOM_PI: path.Seal True
    Set TestCirclePath = path
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт связный путь из горизонтальной прямой, четверти окружности и вертикальной прямой.
Private Function TestHelperCompoundPath() As CGeometryPath
    Dim path As CGeometryPath
    Set path = New CGeometryPath: path.Initialize "Outer.Mixed"
    path.AddLine "Bottom", 0#, 0#, 100#, 0#
    path.AddArc "Corner", 100#, 50#, 50#, -GEOM_PI / 2#, GEOM_PI / 2#
    path.AddLine "Right", 150#, 50#, 150#, 100#: path.Seal
    Set TestHelperCompoundPath = path
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет отказ пути при разрыве, повторном имени, неверной дуге, незамкнутости, недопустимой длине или изменении после фиксации.
Private Function TestInvalidPathRejected(ByVal scenario As Long) As Boolean
    On Error GoTo Rejected
    Dim path As CGeometryPath, part As CGeometryPath, edge As TRegionEdge, value As Double, x As Double, y As Double, tx As Double, ty As Double
    Set path = New CGeometryPath: path.Initialize "Invalid"
    Select Case scenario
        Case 0
            path.AddLine "Duplicate", 0#, 0#, 1#, 0#: path.AddLine "Duplicate", 1#, 0#, 2#, 0#
        Case 1
            path.AddLine "First", 0#, 0#, 1#, 0#: path.AddLine "Disconnected", 2#, 0#, 3#, 0#
        Case 2: path.AddArc "NegativeRadius", 0#, 0#, -1#, 0#, 1#
        Case 3: path.AddArc "ZeroSweep", 0#, 0#, 1#, 0#, 0#
        Case 4: path.AddArc "DoubleTurn", 0#, 0#, 1#, 0#, 4# * GEOM_PI
        Case 5: path.AddLine "Open", 0#, 0#, 1#, 0#: path.Seal True
        Case 6: value = path.Length
        Case 7: path.AddLine "First", 0#, 0#, 1#, 0#: path.Seal: path.AddLine "Late", 1#, 0#, 2#, 0#
        Case 8: path.AddLine "First", 0#, 0#, 1#, 0#: path.Seal: path.AtDistance -1#, x, y, tx, ty
        Case 9: path.Seal
        Case 10: path.AddLine "First", 0#, 0#, 1#, 0#: path.Seal: Set part = path.Slice(0.5, 0.5)
        Case 11: path.AddLine "First", 0#, 0#, 1#, 0#: path.Seal: path.ProjectionBounds 0#, 0#, x, y
        Case 12: path.AddLine "First", 0#, 0#, 1#, 0#: path.Seal: value = path.SegmentIndex("Missing")
        Case 13: path.AddLine "First", 0#, 0#, 1#, 0#: path.CopySegment 1, edge
    End Select
    Exit Function
Rejected:
    TestInvalidPathRejected = (Err.Source = "CGeometryPath" And Err.Number = vbObjectError + 5520)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает фактическую и ожидаемую величины с заданным в этом наборе допуском, сохраняя числа в протоколе.
Private Sub TestNear(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    TestCheck name & "; actual=" & Format$(actual, "0.000000000000000") & "; expected=" & Format$(expected, "0.000000000000000"), _
        Abs(actual - expected) <= 0.0000001
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Учитывает отдельное условие проверки и сохраняет его имя в итоговом протоколе.
Private Sub TestCheck(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        mPassed = mPassed + 1: mReport = mReport & "PASS: " & name & vbCrLf
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
    End If
End Sub
