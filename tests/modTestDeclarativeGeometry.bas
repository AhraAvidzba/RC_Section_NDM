Attribute VB_Name = "modTestDeclarativeGeometry"
Option Explicit

' ДЛЯ ТЕСТОВ: точные линии и дуги общего геометрического слоя без Excel и сетки.
Private mPassed As Long, mFailed As Long, mReport As String

' ДЛЯ ТЕСТОВ: единый короткий набор общего геометрического слоя.
Public Function RunCoreGeometryTests() As String
    Dim result As String
    result = RunPathTests()
    result = result & RunPathAnchorTests()
    result = result & RunRoundedPolygonTests()
    result = result & RunRoundedOffsetTests()
    result = result & RunInputTableTests()
    RunCoreGeometryTests = result
End Function

' ДЛЯ ТЕСТОВ: выделение именованных участков контура, переход через его начало
' и сохранение соответствия сегментов исходным линиям.
Public Function RunPathAnchorTests() As String
    Dim polygonSpec As CPolygonSpec, polygonPoint As Variant
    mPassed = 0: mFailed = 0: mReport = vbNullString
    On Error GoTo Failed
    Dim path As CGeometryPath, part As CGeometryPath, changed As CGeometryPath, edge As TRegionEdge
    Dim query As CSectionGeometryQuery, facade As CSectionDefinitionBuilder, definition As CSectionDefinition
    Dim x As Double, y As Double, tx As Double, ty As Double, scenario As Long
    Set path = CompoundPath
    Near "anchor.named.station", path.AnchorStation("Corner", 0.5), 100# + 12.5 * GEOM_PI
    Set part = path.BetweenAnchors("Chain", "Bottom", 0.25, "Corner", 0.5)
    Check "anchor.chain.id", part.ID = "Chain" And Not part.IsClosed
    Check "anchor.chain.count", part.SegmentCount = 2
    Near "anchor.chain.length", part.Length, 75# + 12.5 * GEOM_PI
    Check "anchor.chain.source.path", part.SourcePathID(1) = "Outer.Mixed" And part.SourcePathID(2) = "Outer.Mixed"
    Check "anchor.chain.source.segment", part.SourceSegmentID(1) = "Bottom" And part.SourceSegmentID(2) = "Corner"
    Near "anchor.chain.source.first", part.SourceStartParameter(1), 0.25
    Near "anchor.chain.source.last", part.SourceEndParameter(2), 0.5
    Set part = path.Slice(20#, 80#): Set changed = part.Reversed
    Near "anchor.reverse.source.first", changed.SourceStartParameter(1), 0.8
    Near "anchor.reverse.source.last", changed.SourceEndParameter(1), 0.2
    Set part = changed.Slice(10#, 50#)
    Near "anchor.nested.source.first", part.SourceStartParameter(1), 0.7
    Near "anchor.nested.source.last", part.SourceEndParameter(1), 0.3
    Set changed = part.Transformed(1000#, -2000#, GEOM_PI / 2#, True): Set part = changed.WithID("Moved")
    Check "anchor.transform.id", part.ID = "Moved" And part.SegmentID(1) = "Bottom"
    Check "anchor.transform.source", part.SourcePathID(1) = "Outer.Mixed" And part.SourceSegmentID(1) = "Bottom"
    Near "anchor.transform.source.first", part.SourceStartParameter(1), 0.7
    Near "anchor.transform.source.last", part.SourceEndParameter(1), 0.3
    part.AtDistance 0#, x, y, tx, ty
    Near "anchor.transform.x", x, 1000#: Near "anchor.transform.y", y, -1930#
    Near "anchor.source.unchanged", path.Length, 150# + 25# * GEOM_PI
    Set path = CirclePath
    Set part = path.BetweenAnchors("Wrap", "Circumference", 0.75, "Circumference", 0.25)
    Check "anchor.wrap.count", part.SegmentCount = 2 And Not part.IsClosed
    Check "anchor.wrap.unique.ids", part.SegmentID(1) <> part.SegmentID(2)
    Check "anchor.wrap.source", part.SourceSegmentID(1) = "Circumference" And part.SourceSegmentID(2) = "Circumference"
    Near "anchor.wrap.length", part.Length, 10# * GEOM_PI
    Near "anchor.wrap.first.start", part.SourceStartParameter(1), 0.75
    Near "anchor.wrap.first.end", part.SourceEndParameter(1), 1#
    Near "anchor.wrap.second.start", part.SourceStartParameter(2), 0#
    Near "anchor.wrap.second.end", part.SourceEndParameter(2), 0.25
    part.AtDistance part.Length / 2#, x, y, tx, ty
    Near "anchor.wrap.no.chord.x", x, 10#: Near "anchor.wrap.no.chord.y", y, 0#
    Set part = path.BetweenAnchors("Reverse", "Circumference", 0.75, "Circumference", 0.25, False)
    Check "anchor.reverse.count", part.SegmentCount = 1
    Near "anchor.reverse.length", part.Length, 10# * GEOM_PI
    Near "anchor.reverse.first", part.SourceStartParameter(1), 0.75
    Near "anchor.reverse.last", part.SourceEndParameter(1), 0.25
    part.AtDistance part.Length / 2#, x, y, tx, ty
    Near "anchor.reverse.other.side.x", x, -10#: Near "anchor.reverse.other.side.y", y, 0#
    part.CopySegment 1, edge: Near "anchor.reverse.sweep", edge.Sweep, -GEOM_PI
    Set part = path.BetweenAnchors("Loop", "Circumference", 0.25, "Circumference", 0.25, True, True)
    Check "anchor.full.loop.closed", part.IsClosed And part.SegmentCount = 2
    Near "anchor.full.loop.length", part.Length, path.Length
    Set part = path.BetweenAnchors("Seam", "Circumference", 1#, "Circumference", 0#, True, True)
    Check "anchor.seam.loop.closed", part.IsClosed And part.SegmentCount = 1
    Near "anchor.seam.loop.length", part.Length, path.Length
    Set part = path.BetweenAnchors("SeamReverse", "Circumference", 0#, "Circumference", 1#, False, True)
    Check "anchor.seam.reverse.closed", part.IsClosed
    Near "anchor.seam.reverse.source.first", part.SourceStartParameter(1), 1#
    Set query = New CSectionGeometryQuery
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, 0#
    polygonSpec.AddVertex "BR", 100#, 0#, 0#
    polygonSpec.AddVertex "TR", 100#, 80#, 0#
    polygonSpec.AddVertex "TL", 0#, 80#, 0#
    polygonSpec.Seal
    Set path = query.RoundedPolygon("Sharp", polygonSpec)
    Set part = path.BetweenAnchors("NamedSide", "Edge.BR", 0#, "Edge.TR", 0#)
    Check "anchor.zero.radius.named.edge", part.SegmentCount = 1 And part.SegmentID(1) = "Edge.BR"
    Near "anchor.zero.radius.length", part.Length, 80#
    ' Нулевая дуга не является сегментом, но её границы сохраняют стык.
    Check "anchor.zero.radius.no.phantom", Not path.HasSegmentID("Corner.BL")
    Near "anchor.zero.radius.junction.start", path.AnchorStation("Corner.BL", 0#), path.AnchorStation("Edge.BL", 0#)
    Near "anchor.zero.radius.junction.end", path.AnchorStation("Corner.BL", 1#), path.AnchorStation("Edge.BL", 0#)
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
    Check "anchor.facade.registered", definition.PathCount = 3 And definition.IsSealed
    Check "anchor.facade.source", changed.SourcePathID(1) = "Outer" And changed.SourceSegmentID(1) = "Edge.BR"
    changed.AtDistance 0#, x, y, tx, ty
    Near "anchor.facade.moved.x", x, 1000100#: Near "anchor.facade.moved.y", y, -1000000#
    For scenario = 0 To 9: Check "anchor.invalid." & CStr(scenario), InvalidAnchorRejected(scenario): Next scenario
    GoTo Finished
Failed:
    Check "anchor.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunPathAnchorTests = mReport & "TOTAL PATH ANCHORS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Private Function InvalidAnchorRejected(ByVal scenario As Long) As Boolean
    Dim polygonSpec As CPolygonSpec, polygonPoint As Variant
    On Error GoTo Rejected
    Dim path As CGeometryPath, part As CGeometryPath, query As CSectionGeometryQuery, station As Double
    Set path = CirclePath
    Select Case scenario
        Case 0: station = path.AnchorStation("Missing")
        Case 1: station = path.AnchorStation("Circumference", -0.1)
        Case 2: station = path.AnchorStation("Circumference", 1.1)
        Case 3: Set part = path.BetweenAnchors("Same", "Circumference", 0.25, "Circumference", 0.25)
        Case 4: Set part = path.BetweenAnchors("Seam", "Circumference", 1#, "Circumference", 0#)
        Case 5: Set part = path.BetweenAnchors("Seam", "Circumference", 0#, "Circumference", 1#)
        Case 6
            Set path = CompoundPath
            Set part = path.BetweenAnchors("WrongDirection", "Right", 0#, "Bottom", 0#)
        Case 7
            Set path = CompoundPath
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
    InvalidAnchorRejected = (Err.Source = "CGeometryPath" And Err.Number = vbObjectError + 5520)
End Function

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
        Check "path.invalid." & CStr(scenario), InvalidPathRejected(scenario)
    Next scenario
    GoTo Finished
Failed:
    Check "path.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunPathTests = mReport & "TOTAL ANALYTIC PATHS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: общий алгоритм скругления, включая вогнутые вершины и нулевой радиус.
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
    Check "rounded.count", path.SegmentCount = 8
    Check "rounded.closed", path.IsClosed
    Near "rounded.length", path.Length, 280# + 20# * GEOM_PI
    Set region = query.RegionFromPath(path)
    Near "rounded.area", region.Area, 8000# - 400# + 100# * GEOM_PI
    For i = 1 To path.SegmentCount
        path.CopySegment i, edge
        If edge.Sweep <> 0# Then
            radius = Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2)
            Near "rounded.radius." & CStr(i), radius, 10#
            Near "rounded.sweep." & CStr(i), edge.Sweep, GEOM_PI / 2#
        End If
    Next i
    Check "rounded.named.edge", path.SegmentID(path.SegmentIndex("Edge.BR")) = "Edge.BR"
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, 0#
    polygonSpec.AddVertex "BR", 100#, 0#, 10#
    polygonSpec.AddVertex "TR", 100#, 80#, 20#
    polygonSpec.AddVertex "TL", 0#, 80#, 5#
    polygonSpec.Seal
    Set path = query.RoundedPolygon("Mixed", polygonSpec)
    Check "rounded.zero.radius.count", path.SegmentCount = 7
    Check "rounded.zero.radius.edge", path.SegmentIndex("Edge.BL") = 1
    Set region = query.RegionFromPath(path)
    Near "rounded.individual.area", region.Area, 8000# - 525# * (1# - GEOM_PI / 4#)
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, 0#
    polygonSpec.AddVertex "BR", 100#, 0#, 0#
    polygonSpec.AddVertex "TR", 100#, 80#, 0#
    polygonSpec.AddVertex "TL", 0#, 80#, 0#
    polygonSpec.Seal
    Set path = query.RoundedPolygon("Sharp", polygonSpec)
    Near "rounded.sharp.length", path.Length, 360#
    Set region = query.RegionFromPath(path)
    Near "rounded.sharp.area", region.Area, 8000#
    ' Вогнутый угол сохраняет противоположное направление своей дуги.
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, 10#
    polygonSpec.AddVertex "BR", 200#, 0#, 10#
    polygonSpec.AddVertex "Arm", 200#, 100#, 10#
    polygonSpec.AddVertex "Inside", 100#, 100#, 10#
    polygonSpec.AddVertex "TR", 100#, 200#, 10#
    polygonSpec.AddVertex "TL", 0#, 200#, 10#
    polygonSpec.Seal
    Set path = query.RoundedPolygon("Concave", polygonSpec)
    path.CopySegment path.SegmentIndex("Corner.Inside"), edge
    Near "rounded.concave.sweep", edge.Sweep, -GEOM_PI / 2#
    Near "rounded.concave.cx", edge.CenterX, 110#: Near "rounded.concave.cy", edge.CenterY, 110#
    Near "rounded.concave.start.x", edge.X1, 110#: Near "rounded.concave.start.y", edge.Y1, 100#
    Near "rounded.concave.end.x", edge.X2, 100#: Near "rounded.concave.end.y", edge.Y2, 110#
    Set region = query.RegionFromPath(path)
    Near "rounded.concave.area", region.Area, 30000# - 400# * (1# - GEOM_PI / 4#)
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, 10#
    polygonSpec.AddVertex "TL", 0#, 80#, 10#
    polygonSpec.AddVertex "TR", 100#, 80#, 10#
    polygonSpec.AddVertex "BR", 100#, 0#, 10#
    polygonSpec.Seal
    Set path = query.RoundedPolygon("Clockwise", polygonSpec)
    path.CopySegment path.SegmentIndex("Corner.BL"), edge
    Near "rounded.clockwise.sweep", edge.Sweep, -GEOM_PI / 2#
    Set region = query.RegionFromPath(path)
    Near "rounded.clockwise.area", region.Area, 8000# - 400# + 100# * GEOM_PI
    ' Касание соседних дуг допустимо; исчезнувшая прямая не выбирается по индексу.
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, 50#
    polygonSpec.AddVertex "BR", 100#, 0#, 50#
    polygonSpec.AddVertex "TR", 100#, 100#, 50#
    polygonSpec.AddVertex "TL", 0#, 100#, 50#
    polygonSpec.Seal
    Set path = query.RoundedPolygon("Touching", polygonSpec)
    Check "rounded.touching.count", path.SegmentCount = 4
    Set region = query.RegionFromPath(path)
    Near "rounded.touching.area", region.Area, 2500# * GEOM_PI
    For scenario = 0 To 5
        Check "rounded.invalid." & CStr(scenario), InvalidPolygonRejected(scenario)
    Next scenario
    GoTo Finished
Failed:
    Check "rounded.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRoundedPolygonTests = mReport & "TOTAL ROUNDED POLYGONS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: точные offset-дуги, вогнутый угол и независимость от обхода.
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
    Check "offset.closed", path.IsClosed And path.SegmentCount = 8
    Near "offset.rectangle.length", path.Length, 280# + 10# * GEOM_PI
    Set region = query.RegionFromPath(path)
    Near "offset.rectangle.area", region.Area, 6300# - 100# + 25# * GEOM_PI
    path.CopySegment path.SegmentIndex("Corner.BL"), edge
    Near "offset.rectangle.radius", Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2), 5#
    Near "offset.rectangle.center.x", edge.CenterX, 10#: Near "offset.rectangle.center.y", edge.CenterY, 10#
    Near "offset.original.unchanged", original.Length, 280# + 20# * GEOM_PI
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, CDbl(radii(LBound(radii) + 0))
    polygonSpec.AddVertex "BR", 100#, 0#, CDbl(radii(LBound(radii) + 1))
    polygonSpec.AddVertex "TR", 100#, 80#, CDbl(radii(LBound(radii) + 2))
    polygonSpec.AddVertex "TL", 0#, 80#, CDbl(radii(LBound(radii) + 3))
    polygonSpec.Seal
    Set path = query.OffsetRoundedPolygon("Outward", polygonSpec, -5#)
    Set region = query.RegionFromPath(path)
    Near "offset.outward.area", region.Area, 9900# - 900# + 225# * GEOM_PI
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, CDbl(radii(LBound(radii) + 0))
    polygonSpec.AddVertex "BR", 100#, 0#, CDbl(radii(LBound(radii) + 1))
    polygonSpec.AddVertex "TR", 100#, 80#, CDbl(radii(LBound(radii) + 2))
    polygonSpec.AddVertex "TL", 0#, 80#, CDbl(radii(LBound(radii) + 3))
    polygonSpec.Seal
    Set path = query.OffsetRoundedPolygon("VanishedArc", polygonSpec, 15#)
    Check "offset.vanished.arc.count", path.SegmentCount = 4
    Check "offset.vanished.arc.named.edge", path.SegmentID(path.SegmentIndex("Edge.BL")) = "Edge.BL"
    Set region = query.RegionFromPath(path)
    Near "offset.vanished.arc.area", region.Area, 3500#
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, 0#
    polygonSpec.AddVertex "BR", 100#, 0#, 0#
    polygonSpec.AddVertex "TR", 100#, 80#, 0#
    polygonSpec.AddVertex "TL", 0#, 80#, 0#
    polygonSpec.Seal
    Set path = query.OffsetRoundedPolygon("SharpOutward", polygonSpec, -5#)
    Check "offset.sharp.preserved", path.SegmentCount = 4
    Set region = query.RegionFromPath(path): Near "offset.sharp.area", region.Area, 9900#
    Set polygonSpec = New CPolygonSpec
    polygonSpec.AddVertex "BL", 0#, 0#, CDbl(radii(LBound(radii) + 0))
    polygonSpec.AddVertex "TL", 0#, 80#, CDbl(radii(LBound(radii) + 1))
    polygonSpec.AddVertex "TR", 100#, 80#, CDbl(radii(LBound(radii) + 2))
    polygonSpec.AddVertex "BR", 100#, 0#, CDbl(radii(LBound(radii) + 3))
    polygonSpec.Seal
    Set path = query.OffsetRoundedPolygon("Clockwise", polygonSpec, 5#)
    Set region = query.RegionFromPath(path): Near "offset.clockwise.area", region.Area, 6300# - 100# + 25# * GEOM_PI
    path.CopySegment path.SegmentIndex("Corner.BL"), edge
    Near "offset.clockwise.sweep", edge.Sweep, -GEOM_PI / 2#
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
    Near "offset.concave.radius", Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2), 7#
    Near "offset.concave.center.x", edge.CenterX, 45#: Near "offset.concave.center.y", edge.CenterY, 45#
    Near "offset.concave.sweep", edge.Sweep, -GEOM_PI / 2#
    Set region = query.RegionFromPath(path): Near "offset.concave.area", region.Area, 5616# + 4# - GEOM_PI
    path.AtDistance path.Length / 7#, x, y, tx, ty
    Near "offset.boundary.distance", query.BoundaryDistance(originalRegion, x, y), 2#
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
    Set region = query.RegionFromPath(path): Near "offset.far.area", region.Area, 6300# - 100# + 25# * GEOM_PI
    path.CopySegment path.SegmentIndex("Corner.BL"), edge
    Near "offset.far.center.x", edge.CenterX, 1000010#: Near "offset.far.center.y", edge.CenterY, -999990#
    For scenario = 0 To 2: Check "offset.invalid." & CStr(scenario), InvalidOffsetRejected(scenario): Next scenario
    GoTo Finished
Failed:
    Check "offset.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRoundedOffsetTests = mReport & "TOTAL ROUNDED OFFSET: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Private Function InvalidOffsetRejected(ByVal scenario As Long) As Boolean
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
    InvalidOffsetRejected = (Err.Source = "CSectionGeometryQuery" And Err.Number = vbObjectError + 5401)
End Function

Private Function RectangleVertices() As Variant
    RectangleVertices = Array(Array(0#, 0#), Array(100#, 0#), Array(100#, 80#), Array(0#, 80#))
End Function

Private Function RectangleNames() As Variant
    RectangleNames = Array("BL", "BR", "TR", "TL")
End Function

' ДЛЯ ТЕСТОВ: чтение in-memory ввода и независимость от исходных массивов.
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
    Check "input.rows", table.RowCount = 4
    Check "input.columns", table.ColumnCount = 3
    Check "input.caption", table.ColumnByCaption(1, "value") = 2
    Check "input.key", table.RowByKey("diameter") = 4
    Check "input.key.scoped", table.RowByKey("Count", 1, 2, 3) = 2
    Check "input.key.alias", table.RowByAnyKey(Array("Unknown", "Diameter"), 1, 3, 4) = 4
    Near "input.number", table.Number(2, 2), 4#
    Check "input.count", table.Count(2, 2) = 4
    Check "input.choice", table.Choice(2, 3, Array("Stacked", "SideBySide")) = "Stacked"
    Check "input.choice.default", table.Choice(4, 3, Array("Stacked", "SideBySide"), "SideBySide") = "SideBySide"
    Check "input.source", table.SourceReference(2, 2) = "Config!B12"
    Check "input.blank", table.IsBlank(4, 3)
    Near "input.number.default", table.Number(4, 3, 7#), 7#
    Near "input.optional.zero", table.OptionalDiameter(3, 2, units), 0#
    Near "input.optional.blank", table.OptionalDiameter(3, 3, units), 0#
    Near "input.optional.active", table.OptionalDiameter(4, 2, units), 16#
    Near "input.length", table.Length(4, 2, units), 16#
    cells(2, 2) = 999#: addresses(2, 2) = "Changed"
    Near "input.value.immutable", table.Number(2, 2), 4#
    Check "input.source.immutable", table.SourceReference(2, 2) = "Config!B12"
    cells(2, 2) = 2147483647#: addresses(2, 2) = "Config!B12"
    Set table = New CSectionInputTable: table.Initialize "rngLongBoundary", cells, addresses
    Check "input.count.maxLong", table.Count(2, 2) = 2147483647
    For scenario = 0 To 24
        Check "input.invalid." & CStr(scenario), InvalidInputRejected(scenario)
    Next scenario
    GoTo Finished
Failed:
    Check "input.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunInputTableTests = mReport & "TOTAL INPUT TABLES: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Private Function InvalidInputRejected(ByVal scenario As Long) As Boolean
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
    InvalidInputRejected = (Err.Source = "CSectionInputTable" And Err.Number = vbObjectError + 5525)
    If scenario = 0 Then InvalidInputRejected = InvalidInputRejected And InStr(1, Err.Description, "rngInvalid[2,2]", vbBinaryCompare) > 0
    If scenario >= 22 Then InvalidInputRejected = InvalidInputRejected And InStr(1, Err.Description, "rngInvalid[2,2]", vbBinaryCompare) > 0
End Function

Private Function InvalidPolygonRejected(ByVal scenario As Long) As Boolean
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
        ' Структура теперь проверяется описателем до передачи геометрическому ядру.
        InvalidPolygonRejected = (Err.Source = "CPolygonSpec" And Err.Number = vbObjectError + 5470)
    Else
        InvalidPolygonRejected = (Err.Source = "CSectionGeometryQuery" And Err.Number = vbObjectError + 5401)
    End If
End Function

Private Sub TestLine()
    Dim path As CGeometryPath, changed As CGeometryPath
    Dim x As Double, y As Double, tx As Double, ty As Double, minimum As Double, maximum As Double
    Set path = New CGeometryPath: path.Initialize "Outer.Inclined"
    path.AddLine "Inclined", 0#, 0#, 3#, 4#: path.Seal
    Near "line.length", path.Length, 5#
    path.AtDistance 2.5, x, y, tx, ty
    Near "line.x", x, 1.5: Near "line.y", y, 2#
    Near "line.tx", tx, 0.6: Near "line.ty", ty, 0.8
    path.ProjectionBounds 2#, 0#, minimum, maximum
    Near "line.minimum", minimum, 0#: Near "line.maximum", maximum, 3#
    Set changed = path.Reversed
    changed.AtDistance 0#, x, y, tx, ty
    Near "line.reverse.x", x, 3#: Near "line.reverse.y", y, 4#
    Near "line.reverse.tx", tx, -0.6: Near "line.reverse.ty", ty, -0.8
    Set changed = path.Transformed(10#, 20#, GEOM_PI / 2#)
    changed.AtDistance changed.Length, x, y, tx, ty
    Near "line.rotate.x", x, 6#: Near "line.rotate.y", y, 23#
    Near "line.rotate.tx", tx, -0.8: Near "line.rotate.ty", ty, 0.6
    Check "line.transform.id", changed.SegmentID(1) = "Inclined"
    Near "line.transform.length", changed.Length, path.Length
    Check "line.sealed", changed.IsSealed And path.IsSealed
    Check "line.open", Not changed.IsClosed
End Sub

Private Sub TestArcs()
    Dim path As CGeometryPath, changed As CGeometryPath, edge As TRegionEdge
    Dim x As Double, y As Double, tx As Double, ty As Double, minimum As Double, maximum As Double
    Set path = New CGeometryPath: path.Initialize "Arc.Positive"
    path.AddArc "Curve", 0#, 0#, 10#, GEOM_PI / 4#, GEOM_PI: path.Seal
    Near "arc.length", path.Length, 10# * GEOM_PI
    path.AtDistance path.Length / 4#, x, y, tx, ty
    Near "arc.internal.x", x, 0#: Near "arc.internal.y", y, 10#
    Near "arc.internal.tx", tx, -1#: Near "arc.internal.ty", ty, 0#
    path.ProjectionBounds 0#, 1#, minimum, maximum
    Near "arc.bounds.y.min", minimum, -10# / Sqr(2#)
    Near "arc.bounds.y.max", maximum, 10#
    path.ProjectionBounds 1#, 0#, minimum, maximum
    Near "arc.bounds.x.min", minimum, -10#
    Near "arc.bounds.x.max", maximum, 10# / Sqr(2#)
    Set changed = path.Reversed
    changed.AtDistance changed.Length, x, y, tx, ty
    Near "arc.reverse.end.x", x, 10# / Sqr(2#)
    Near "arc.reverse.end.y", y, 10# / Sqr(2#)
    Near "arc.reverse.end.tx", tx, 1# / Sqr(2#)
    Near "arc.reverse.end.ty", ty, -1# / Sqr(2#)
    Set changed = path.Transformed(1000000#, -1000000#, 0#, True)
    changed.CopySegment 1, edge
    Near "arc.reflection.sweep", edge.Sweep, -GEOM_PI
    changed.AtDistance changed.Length / 4#, x, y, tx, ty
    Near "arc.reflection.x", x, 1000000#
    Near "arc.reflection.y", y, -1000010#
    Near "arc.reflection.length", changed.Length, path.Length
    Check "arc.reflection.id", changed.SegmentID(1) = path.SegmentID(1)
    Set path = New CGeometryPath: path.Initialize "Arc.Negative"
    path.AddArc "Curve", 0#, 0#, 10#, GEOM_PI / 4#, -GEOM_PI: path.Seal
    path.AtDistance path.Length / 4#, x, y, tx, ty
    Near "arc.negative.x", x, 10#: Near "arc.negative.y", y, 0#
    Near "arc.negative.tx", tx, 0#: Near "arc.negative.ty", ty, -1#
    path.ProjectionBounds 1#, 0#, minimum, maximum
    Near "arc.negative.max", maximum, 10#
    Near "arc.negative.min", minimum, -10# / Sqr(2#)
    Set path = CirclePath
    Check "circle.closed", path.IsClosed
    Near "circle.length", path.Length, 20# * GEOM_PI
    path.AtDistance 0#, x, y, tx, ty
    Near "circle.start.x", x, 10#: Near "circle.start.y", y, 0#
    path.AtDistance path.Length, x, y, tx, ty
    Near "circle.end.x", x, 10#: Near "circle.end.y", y, 0#
    path.ProjectionBounds 0#, 1#, minimum, maximum
    Near "circle.min", minimum, -10#: Near "circle.max", maximum, 10#
End Sub

Private Sub TestCompoundPath()
    Dim path As CGeometryPath, part As CGeometryPath
    Dim x As Double, y As Double, tx As Double, ty As Double
    Set path = CompoundPath
    Near "compound.length", path.Length, 150# + 25# * GEOM_PI
    Check "compound.count", path.SegmentCount = 3
    path.AtDistance 100# + 12.5 * GEOM_PI, x, y, tx, ty
    Near "compound.arc.x", x, 100# + 50# / Sqr(2#)
    Near "compound.arc.y", y, 50# - 50# / Sqr(2#)
    Near "compound.arc.tx", tx, 1# / Sqr(2#)
    Near "compound.arc.ty", ty, 1# / Sqr(2#)
    path.AtDistance 100#, x, y, tx, ty
    Near "compound.join.x", x, 100#: Near "compound.join.y", y, 0#
    Near "compound.join.tx", tx, 1#: Near "compound.join.ty", ty, 0#
    Set part = path.Slice(50#, path.Length - 25#)
    Near "compound.slice.length", part.Length, path.Length - 75#
    Check "compound.slice.segments", part.SegmentCount = 3
    Check "compound.slice.id", part.SegmentID(2) = "Corner"
    part.AtDistance 0#, x, y, tx, ty
    Near "compound.slice.start.x", x, 50#: Near "compound.slice.start.y", y, 0#
    part.AtDistance part.Length, x, y, tx, ty
    Near "compound.slice.end.x", x, 150#: Near "compound.slice.end.y", y, 75#
    Near "compound.source.unchanged", path.Length, 150# + 25# * GEOM_PI
    Set part = path.Transformed(1000#, -2000#, GEOM_PI / 2#, True)
    part.AtDistance part.Length, x, y, tx, ty
    Near "compound.transform.x", x, 1100#: Near "compound.transform.y", y, -1850#
    Near "compound.transform.tx", tx, 1#: Near "compound.transform.ty", ty, 0#
End Sub

' G09: разбиение той же дуги на 25 частей не меняет точки или касательные.
Private Sub TestJointTangents()
    Dim path As CGeometryPath, x As Double, y As Double, tx As Double, ty As Double
    Set path = New CGeometryPath: path.Initialize "Joint"
    path.AddLine "Incoming", 0#, 0#, 100#, 0#
    path.AddLine "Outgoing", 100#, 0#, 100#, 100#
    path.Seal
    path.AtDistance 100#, x, y, tx, ty
    Near "joint.default.tx", tx, 0#: Near "joint.default.ty", ty, 1#
    path.AtDistance 100#, x, y, tx, ty, True
    Near "joint.incoming.x", x, 100#: Near "joint.incoming.y", y, 0#
    Near "joint.incoming.tx", tx, 1#: Near "joint.incoming.ty", ty, 0#
    path.AtDistance 100# + GEOM_TOLERANCE / 2#, x, y, tx, ty, True
    Near "joint.roundoff.x", x, 100#: Near "joint.roundoff.y", y, 0#
    path.AtDistance 100# + 2# * GEOM_TOLERANCE, x, y, tx, ty, True
    Near "joint.after.tx", tx, 0#: Near "joint.after.ty", ty, 1#
End Sub

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
    Near "split.length", split.Length, whole.Length
    For i = 0 To 40
        whole.AtDistance whole.Length * CDbl(i) / 40#, x1, y1, tx1, ty1
        split.AtDistance split.Length * CDbl(i) / 40#, x2, y2, tx2, ty2
        Near "split.x." & CStr(i), x2, x1
        Near "split.y." & CStr(i), y2, y1
        Near "split.tx." & CStr(i), tx2, tx1
        Near "split.ty." & CStr(i), ty2, ty1
    Next i
End Sub

Private Sub TestRays()
    Dim path As CGeometryPath, station As Double, distance As Double
    Set path = CirclePath
    Check "ray.circle.hit", path.RayHit(-20#, 0#, 2#, 0#, station, distance)
    Near "ray.circle.nearest", distance, 10#
    Near "ray.circle.station", station, 10# * GEOM_PI
    Check "ray.circle.away", Not path.RayHit(-20#, 0#, -1#, 0#, station, distance)
    Check "ray.circle.tangent", path.RayHit(-20#, 10#, 1#, 0#, station, distance)
    Near "ray.circle.tangent.distance", distance, 20#
    Set path = CompoundPath
    Check "ray.compound.hit", path.RayHit(200#, 75#, -1#, 0#, station, distance)
    Near "ray.compound.distance", distance, 50#
    Near "ray.compound.station", station, 125# + 25# * GEOM_PI
    Set path = New CGeometryPath: path.Initialize "Limited"
    path.AddLine "Line", 0#, 0#, 3#, 4#: path.Seal
    Check "ray.line.beyond.target", Not path.RayHit(10#, 10#, 0#, -1#, station, distance)
    Check "ray.line.collinear", path.RayHit(1.5, 2#, 0.6, 0.8, station, distance)
    Near "ray.line.collinear.distance", distance, 0#
    Near "ray.line.collinear.station", station, 2.5
    Set path = New CGeometryPath: path.Initialize "LimitedArc"
    path.AddArc "Curve", 0#, 0#, 10#, GEOM_PI / 4#, GEOM_PI: path.Seal
    Check "ray.arc.beyond.target", Not path.RayHit(0#, 0#, 1#, 0#, station, distance)
End Sub

Private Function CirclePath() As CGeometryPath
    Dim path As CGeometryPath
    Set path = New CGeometryPath: path.Initialize "Circle"
    path.AddArc "Circumference", 0#, 0#, 10#, 0#, 2# * GEOM_PI: path.Seal True
    Set CirclePath = path
End Function

Private Function CompoundPath() As CGeometryPath
    Dim path As CGeometryPath
    Set path = New CGeometryPath: path.Initialize "Outer.Mixed"
    path.AddLine "Bottom", 0#, 0#, 100#, 0#
    path.AddArc "Corner", 100#, 50#, 50#, -GEOM_PI / 2#, GEOM_PI / 2#
    path.AddLine "Right", 150#, 50#, 150#, 100#: path.Seal
    Set CompoundPath = path
End Function

Private Function InvalidPathRejected(ByVal scenario As Long) As Boolean
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
    InvalidPathRejected = (Err.Source = "CGeometryPath" And Err.Number = vbObjectError + 5520)
End Function

Private Sub Near(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    Check name & "; actual=" & Format$(actual, "0.000000000000000") & "; expected=" & Format$(expected, "0.000000000000000"), _
        Abs(actual - expected) <= 0.0000001
End Sub

Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        mPassed = mPassed + 1: mReport = mReport & "PASS: " & name & vbCrLf
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
    End If
End Sub
