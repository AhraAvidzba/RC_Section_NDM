Attribute VB_Name = "modTestPolygonSpec"
Option Explicit

' ДЛЯ ТЕСТОВ: структура общего описания вершин многоугольника.
Private mPassed As Long, mFailed As Long, mReport As String
Private Type TDoubleValue
    Value As Double
End Type
Private Type TDoubleWords
    Low As Long
    High As Long
End Type

' ДЛЯ ТЕСТОВ: порядок/источники, более шести вершин, ошибки и повторное чтение снимка.
Public Function RunPolygonSpecTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim polygon As CPolygonSpec, n As Variant, i As Long, errorCode As Long, reason As String
    Dim vertices As Variant, radii As Variant, names As Variant, number As TDoubleValue, words As TDoubleWords
    For Each n In Array(4, 5, 6, 9)
        Set polygon = New CPolygonSpec
        Check "new.open." & CStr(n), Not polygon.IsSealed And polygon.Count = 0
        For i = 1 To CLng(n)
            polygon.AddVertex "V" & CStr(i), i * 10#, -i * 20#, i - 1#, "Config!C" & CStr(i)
        Next i
        Check "count.before.seal." & CStr(n), polygon.Count = CLng(n)
        polygon.Seal: polygon.Seal
        Check "sealed." & CStr(n), polygon.IsSealed
        For i = 1 To CLng(n)
            Check "record." & CStr(n) & "." & CStr(i), polygon.VertexID(i) = "V" & CStr(i) And _
                polygon.X(i) = i * 10# And polygon.Y(i) = -i * 20# And polygon.FilletRadius(i) = i - 1# And _
                polygon.SourceReference(i) = "Config!C" & CStr(i)
        Next i
        polygon.CopyKernelData vertices, radii, names
        vertices(0) = Array(-1#, -2#): radii(0) = 999#: names(0) = "Changed"
        Check "kernel.copy.independent." & CStr(n), polygon.X(1) = 10# And polygon.FilletRadius(1) = 0# And polygon.VertexID(1) = "V1"
        On Error Resume Next
        polygon.AddVertex "After", 0#, 0#, 0#, "Config!After"
        errorCode = Err.Number: reason = Err.Description: Err.Clear
        On Error GoTo Failed
        Check "reject.after.seal." & CStr(n), errorCode <> 0 And polygon.Count = CLng(n) And InStr(reason, "Config!After") > 0
    Next n
    Set polygon = New CPolygonSpec
    polygon.AddVertex "A", 0#, 0#
    polygon.AddVertex "a", 1#, 0#
    Check "ids.binary.case", polygon.Count = 2
    On Error Resume Next
    polygon.AddVertex "A", 0#, 1#, 0#, "Config!Duplicate"
    errorCode = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    Check "reject.duplicate", errorCode <> 0 And polygon.Count = 2 And InStr(reason, "Config!Duplicate") > 0
    On Error Resume Next
    polygon.AddVertex " ", 0#, 1#, 0#, "Config!Empty"
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "reject.empty.id", errorCode <> 0 And polygon.Count = 2
    On Error Resume Next
    polygon.AddVertex "Negative", 0#, 1#, -1#, "Config!Radius"
    errorCode = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    Check "reject.negative.radius", errorCode <> 0 And polygon.Count = 2 And InStr(reason, "Config!Radius") > 0
    On Error Resume Next
    polygon.Seal
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "reject.incomplete", errorCode <> 0 And Not polygon.IsSealed
    words.High = &H7FF00000: LSet number = words
    On Error Resume Next
    polygon.AddVertex "Inf", number.Value, 1#
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "reject.infinity", errorCode <> 0 And polygon.Count = 2
    words.Low = 1: LSet number = words
    On Error Resume Next
    polygon.AddVertex "NaN", 1#, 1#, number.Value
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "reject.nan", errorCode <> 0 And polygon.Count = 2
    polygon.AddVertex "B", 0#, 1#: polygon.Seal
    Check "recover.after.failed.seal", polygon.Count = 3 And polygon.IsSealed
    GoTo Finished
Failed:
    Check "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunPolygonSpecTests = mReport & "TOTAL POLYGON SPEC: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function
' ДЛЯ ТЕСТОВ: явные имена, прежние вызовы, преобразования, происхождение и публикация контуров.
Public Function RunNamedSegmentTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    TestNamedPolygonInput
    TestNamedPolygonOperations
    TestNamedRectangles
    TestNamedPublication
    GoTo Finished
Failed:
    Check "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunNamedSegmentTests = mReport & "TOTAL NAMED SEGMENTS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: проверяем эффективные имена до геометрии, включая отсутствующие дуги и грани.
Private Sub TestNamedPolygonInput()
    Dim p As CPolygonSpec, code As Long, reason As String
    Dim vertices As Variant, radii As Variant, names As Variant, edges As Variant, corners As Variant
    Set p = New CPolygonSpec
    p.AddVertex "A", 0#, 0#, outgoingEdgeID:="Bottom", cornerID:="Round.A"
    On Error Resume Next
    p.AddVertex "B", 100#, 0#, 0#, "Config!Duplicate", "Round.A"
    code = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo 0
    Check "duplicate.collapsed.corner", code <> 0 And p.Count = 1 And InStr(reason, "Config!Duplicate") > 0
    On Error Resume Next
    p.AddVertex "B", 100#, 0#, 0#, "", "X", "X"
    code = Err.Number: Err.Clear
    On Error GoTo 0
    Check "duplicate.same.vertex", code <> 0 And p.Count = 1
    p.AddVertex "B", 100#, 0#, outgoingEdgeID:="", cornerID:="Round.B"
    On Error Resume Next
    p.AddVertex "C", 100#, 80#, 40#, "", "Edge.B"
    code = Err.Number: Err.Clear
    On Error GoTo 0
    Check "duplicate.explicit.default", code <> 0 And p.Count = 2
    On Error Resume Next
    p.AddVertex "C", 100#, 80#, 40#, "", "Edge.B", "Unused"
    code = Err.Number: Err.Clear
    On Error GoTo 0
    Check "duplicate.reserved.vanishing.edge", code <> 0 And p.Count = 2
    p.AddVertex "C", 100#, 80#, 0#, "", "bottom" ' Регистр имён различается, как и прежде.
    p.AddVertex "D", 0#, 80#, 0#, "", " ", " "
    p.Seal
    Check "partial.names", p.OutgoingEdgeID(1) = "Bottom" And p.CornerID(1) = "Round.A" And p.VertexID(1) = "A"
    Check "empty.names.defaults", p.OutgoingEdgeID(2) = "Edge.B" And p.CornerID(3) = "Corner.C"
    Check "whitespace.names.defaults", p.OutgoingEdgeID(4) = "Edge.D" And p.CornerID(4) = "Corner.D"
    p.CopyKernelData vertices, radii, names, edges, corners
    edges(0) = "Changed": corners(0) = "Changed": names(0) = "Changed"
    Check "kernel.names.independent", p.OutgoingEdgeID(1) = "Bottom" And p.CornerID(1) = "Round.A" And p.VertexID(1) = "A"
    On Error Resume Next
    p.AddVertex "E", 1#, 1#, outgoingEdgeID:="New"
    code = Err.Number: Err.Clear
    On Error GoTo 0
    Check "sealed.names.immutable", code <> 0 And p.Count = 4 And p.OutgoingEdgeID(1) = "Bottom"
End Sub

' ДЛЯ ТЕСТОВ: создаём одинаковую геометрию с автоматическими либо явно объявленными именами.
Private Function NamedBoxPolygon(ByVal radius As Double, Optional ByVal named As Boolean = True) As CPolygonSpec
    Dim p As CPolygonSpec, points As Variant, point As Variant, i As Long, id As String, edgeNames As Variant
    Set p = New CPolygonSpec
    points = Array(Array(0#, 0#), Array(100#, 0#), Array(100#, 80#), Array(0#, 80#))
    edgeNames = Array("Bottom", "Right", "Top", "Left")
    For i = 0 To 3
        point = points(i): id = Chr$(65 + i)
        If named Then
            p.AddVertex id, CDbl(point(0)), CDbl(point(1)), radius, "Config!" & id, CStr(edgeNames(i)), "Round." & id
        Else
            p.AddVertex id, CDbl(point(0)), CDbl(point(1)), radius, "Config!" & id
        End If
    Next i
    p.Seal: Set NamedBoxPolygon = p
End Function

' ДЛЯ ТЕСТОВ: имена не меняют точные линии и дуги, а производные пути сохраняют исходные параметры.
Private Sub TestNamedPolygonOperations()
    Dim query As CSectionGeometryQuery, named As CPolygonSpec, legacy As CPolygonSpec
    Dim path As CGeometryPath, old As CGeometryPath, part As CGeometryPath, moved As CGeometryPath, reverse As CGeometryPath
    Dim i As Long, radius As Variant, distance As Variant, first As Double, code As Long, a As TRegionEdge, b As TRegionEdge
    Set query = New CSectionGeometryQuery
    For Each radius In Array(0#, 10#, 40#)
        Set named = NamedBoxPolygon(CDbl(radius)): Set legacy = NamedBoxPolygon(CDbl(radius), False)
        For Each distance In Array(0#, 5#, -5#, 15#)
            Set path = query.OffsetRoundedPolygon("Box", named, CDbl(distance))
            Set old = query.OffsetRoundedPolygon("Box", legacy, CDbl(distance))
            CompareNamedGeometry "polygon." & CStr(radius) & ".offset." & CStr(distance), old, path
            Check "polygon.source.unchanged", named.X(1) = 0# And named.FilletRadius(1) = CDbl(radius) And named.OutgoingEdgeID(1) = "Bottom"
        Next distance
        Set path = query.RoundedPolygon("Box", named)
        Set old = query.RoundedPolygon("Box", legacy)
        CompareNamedGeometry "polygon.base." & CStr(radius), old, path
        If CDbl(radius) = 0# Then
            Check "zero.corner.anchor", path.HasCollapsedSegmentID("Round.A") And path.AnchorStation("Round.A", 0#) = path.AnchorStation("Round.A", 1#)
            Set moved = query.OffsetPath("ExistingOffset", path, 5#)
            Check "offset.keeps.collapsed.corner", moved.HasCollapsedSegmentID("Round.A") And moved.AnchorStation("Round.B") = moved.AnchorStation("Right")
        ElseIf CDbl(radius) = 40# Then
            Check "zero.edge.anchor", path.HasCollapsedSegmentID("Right") And Not path.HasSegmentID("Right")
            Set moved = query.OffsetPath("ExistingOffset", path, 5#)
            Check "offset.keeps.collapsed.edge", moved.HasCollapsedSegmentID("Right") And moved.AnchorStation("Right") = moved.AnchorStation("Round.C")
            Check "zero.edge.no.fake.geometry", moved.SegmentCount = path.SegmentCount
        Else
            Set moved = query.OffsetPath("ExistingOffset", path, 5#)
            Check "offset.name.origin", moved.SegmentID(1) = "Round.A" And moved.SourcePathID(1) = "Box" And moved.SourceSegmentID(1) = "Round.A"
        End If
        Set moved = path.Transformed(15#, -20#, GEOM_PI / 3#, True)
        Set reverse = moved.Reversed
        CompareNamedGeometry "transform.named." & CStr(radius), old.Transformed(15#, -20#, GEOM_PI / 3#, True), moved
        CompareNamedGeometry "reverse.named." & CStr(radius), moved.Reversed, reverse
        For i = 1 To path.SegmentCount
            Check "transform.name.origin", moved.SegmentID(i) = path.SegmentID(i) And moved.SourcePathID(i) = path.SourcePathID(i) And moved.SourceSegmentID(i) = path.SourceSegmentID(i)
            path.CopySegment i, a: moved.CopySegment i, b
            Check "reflection.arc.sweep", b.Sweep = -a.Sweep
            Check "reverse.name.origin", reverse.SegmentID(path.SegmentCount - i + 1) = path.SegmentID(i) And reverse.SourceStartParameter(path.SegmentCount - i + 1) = path.SourceEndParameter(i)
        Next i
        If path.HasCollapsedSegmentID("Right") Then Check "transform.collapsed.edge", moved.AnchorStation("Right") = path.AnchorStation("Right")
        If path.HasCollapsedSegmentID("Round.A") Then Check "reverse.collapsed.corner", reverse.AnchorStation("Round.A") = path.Length - path.AnchorStation("Round.A")
    Next radius
    Set named = NamedBoxPolygon(10#): Set path = query.RoundedPolygon("Box", named)
    first = path.AnchorStation("Bottom", 0.25)
    Set part = path.Slice(first, path.AnchorStation("Bottom", 0.75))
    Check "slice.name.origin", part.SegmentID(1) = "Bottom" And part.SourcePathID(1) = "Box" And part.SourceSegmentID(1) = "Bottom"
    Check "slice.parameters", Abs(part.SourceStartParameter(1) - 0.25) < 0.000000001 And Abs(part.SourceEndParameter(1) - 0.75) < 0.000000001
    Set part = path.BetweenAnchors("Lap", "Bottom", 0.25, "Bottom", 0.25, True, True)
    Check "full.loop.split", part.IsClosed And part.SegmentCount = path.SegmentCount + 1 And part.SegmentID(part.SegmentCount) = "Bottom.Part2"
    Check "full.loop.origin", part.SourceSegmentID(part.SegmentCount) = "Bottom" And part.SourcePathID(part.SegmentCount) = "Box"
    Check "full.loop.parameters", Abs(part.SourceStartParameter(1) - 0.25) < 0.000000001 And Abs(part.SourceEndParameter(part.SegmentCount) - 0.25) < 0.000000001
    Set reverse = part.Reversed: Set moved = reverse.Transformed(1#, 2#, GEOM_PI / 2#, True)
    Check "parts.transform.origin", moved.SegmentID(1) = "Bottom.Part2" And moved.SourceSegmentID(1) = "Bottom" And moved.SourceStartParameter(1) = part.SourceEndParameter(part.SegmentCount)
    Set part = path.BetweenAnchors("Backwards", "Top", 0.8, "Bottom", 0.2, False)
    Check "chain.backwards", part.SegmentID(1) = "Top" And part.SourceStartParameter(1) > part.SourceEndParameter(1)
    Set named = NamedBoxPolygon(0#): Set path = query.RoundedPolygon("Sharp", named)
    Set part = path.BetweenAnchors("SharpSide", "Top", 0.5, "Bottom", 0.5)
    Check "chain.collapsed.inside", part.HasCollapsedSegmentID("Round.D") And part.HasCollapsedSegmentID("Round.A")
    Check "chain.collapsed.outside", Not part.HasCollapsedSegmentID("Round.B") And Not part.HasCollapsedSegmentID("Round.C")
    On Error Resume Next
    i = path.SegmentIndex("Round.A")
    code = Err.Number: Err.Clear
    On Error GoTo 0
    Check "collapsed.not.real.segment", code <> 0
    On Error Resume Next
    Set part = path.BetweenAnchors("Invalid", "Top", 0#, "Top", 0#)
    code = Err.Number: Err.Clear
    On Error GoTo 0
    Check "full.loop.explicit.required", code <> 0
End Sub

' ДЛЯ ТЕСТОВ: сравниваем координаты и точные дуги независимо от намеренно разных имён.
Private Sub CompareNamedGeometry(ByVal title As String, ByVal old As CGeometryPath, ByVal actual As CGeometryPath)
    Dim i As Long, a As TRegionEdge, b As TRegionEdge
    Check title & ".count", old.SegmentCount = actual.SegmentCount And old.IsClosed = actual.IsClosed
    Check title & ".length", old.Length = actual.Length
    For i = 1 To old.SegmentCount
        old.CopySegment i, a: actual.CopySegment i, b
        Check title & ".edge." & CStr(i), a.X1 = b.X1 And a.Y1 = b.Y1 And a.X2 = b.X2 And a.Y2 = b.Y2 And a.CenterX = b.CenterX And a.CenterY = b.CenterY And a.Sweep = b.Sweep
    Next i
End Sub

' ДЛЯ ТЕСТОВ: стороны узнаём по фактическим координатам; прежний порядок и начало не меняются.
Private Sub TestNamedRectangles()
    Dim facade As CSectionDefinitionBuilder, old As CGeometryPath, named As CGeometryPath, partial As CGeometryPath
    Dim part As CGeometryPath, i As Long, edge As TRegionEdge, code As Long
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Names"
    Set old = facade.RectanglePath("Legacy", -20#, 30#, 100#, 80#)
    Set named = facade.RectanglePath("Named", -20#, 30#, 100#, 80#, "Bottom", "Right", "Top", "Left")
    Set partial = facade.RectanglePath("Partial", -20#, 30#, 100#, 80#, topEdgeID:="Top")
    CompareNamedGeometry "rectangle", old, named
    For i = 1 To named.SegmentCount
        named.CopySegment i, edge
        Select Case named.SegmentID(i)
            Case "Bottom": Check "rectangle.bottom.direction", edge.Y1 = 30# And edge.Y2 = 30# And edge.X2 > edge.X1
            Case "Right": Check "rectangle.right.direction", edge.X1 = 80# And edge.X2 = 80# And edge.Y2 > edge.Y1
            Case "Top": Check "rectangle.top.direction", edge.Y1 = 110# And edge.Y2 = 110# And edge.X2 < edge.X1
            Case "Left": Check "rectangle.left.direction", edge.X1 = -20# And edge.X2 = -20# And edge.Y2 < edge.Y1
            Case Else: Check "rectangle.unknown.side", False
        End Select
        Check "rectangle.origin.named", named.SourceSegmentID(i) = named.SegmentID(i) And named.SourcePathID(i) = "Named"
        If named.SegmentID(i) <> "Top" Then Check "rectangle.partial.defaults", partial.SegmentID(i) = "Partial." & CStr(i)
        Check "rectangle.legacy.defaults", old.SegmentID(i) = "Legacy." & CStr(i)
    Next i
    Set part = facade.Chain("Across", named.ID, "Top", 0.5, "Bottom", 0.5)
    Check "rectangle.chain.top.bottom", part.SourceSegmentID(1) = "Top" And part.SourceSegmentID(part.SegmentCount) = "Bottom"
    Set part = facade.Transform("Turned", named.ID, 0#, 0#, GEOM_PI / 2#, True)
    Check "rectangle.transformed.original.meaning", part.HasSegmentID("Top") And part.SourcePathID(part.SegmentIndex("Top")) = "Named"
    On Error Resume Next
    Set part = facade.RectanglePath("Duplicate", 0#, 0#, 10#, 10#, "Side", "Side")
    code = Err.Number: Err.Clear
    On Error GoTo 0
    Check "rectangle.duplicates", code <> 0
    On Error Resume Next
    Set part = facade.RectanglePath("DefaultCollision", 0#, 0#, 10#, 10#, "DefaultCollision.2")
    code = Err.Number: Err.Clear
    On Error GoTo 0
    Check "rectangle.explicit.default.duplicate", code <> 0
    Set part = facade.RectanglePath("Repeat", 0#, 0#, 100#, 80#, "Side", "Side.Part2", "Top", "Left")
    Set part = facade.Chain("RepeatLap", part.ID, "Side", 0.5, "Side", 0.5, True, True)
    Check "parts.skip.reserved.suffix", part.SegmentID(part.SegmentCount) = "Side.Part3" And part.SourceSegmentID(part.SegmentCount) = "Side"
    Dim polygon As CPolygonSpec
    Set polygon = New CPolygonSpec
    polygon.AddVertex "A", 0#, 0#, outgoingEdgeID:="Bottom", cornerID:="Bottom.Part2"
    polygon.AddVertex "B", 100#, 0#, outgoingEdgeID:="Right"
    polygon.AddVertex "C", 100#, 80#, outgoingEdgeID:="Top"
    polygon.AddVertex "D", 0#, 80#, outgoingEdgeID:="Left"
    polygon.Seal
    Set part = facade.RoundedPolygon("ReservedCorner", polygon)
    Set part = facade.Chain("ReservedLap", part.ID, "Bottom", 0.5, "Bottom", 0.5, True, True)
    Check "parts.skip.collapsed.name", part.SegmentID(part.SegmentCount) = "Bottom.Part3" And part.HasCollapsedSegmentID("Bottom.Part2")
    Check "parts.collapsed.origin", part.SourceSegmentID(part.SegmentCount) = "Bottom" And part.SourcePathID(part.SegmentCount) = "ReservedCorner"
End Sub

' ДЛЯ ТЕСТОВ: локальные имена могут повторяться; роли и опубликованные имена их не меняют.
Private Sub TestNamedPublication()
    Dim facade As CSectionDefinitionBuilder, a As CGeometryPath, b As CGeometryPath, hole As CGeometryPath
    Dim definition As CSectionDefinition, contours As CSectionContours, i As Long
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Publication"
    Set a = facade.RectanglePath("A", 0#, 0#, 100#, 100#, "Bottom", "Right", "Top", "Left")
    Set b = facade.RectanglePath("B", 200#, 0#, 100#, 100#, "Bottom", "Right", "Top", "Left")
    Set hole = facade.RectanglePath("Hole", 20#, 20#, 20#, 20#, "Bottom", "Right", "Top", "Left")
    facade.Outer "OuterA", a.ID
    facade.Outer "OuterB", b.ID, namingPolicy:=SequentialByKind, namingPrefix:="B"
    facade.Opening "OpeningA", hole.ID, "OuterA", namingPolicy:=SequentialByKind, namingPrefix:="H"
    Set hole = facade.RectanglePath("Hole2", 250#, 20#, 20#, 20#, "Bottom", "Right", "Top", "Left")
    facade.Opening "OpeningB", hole.ID, "OuterB", contourIDs:=Array("H2.1", "H2.2", "H2.3", "H2.4")
    Set definition = facade.Finish: Set contours = definition.MaterialContours()
    Check "publication.four.loops", contours.Count = 16
    For i = 1 To a.SegmentCount
        Check "publication.preserve", contours.ContourID(i) = a.SegmentID(i)
        Check "publication.sequential", contours.ContourID(i + 4) = "B_LINE_" & CStr(i)
        Check "publication.opening.sequential", contours.ContourID(i + 8) = "H_LINE_" & CStr(i)
        Check "publication.explicit", contours.ContourID(i + 12) = "H2." & CStr(i)
        Check "publication.roles", contours.LoopID(i) = "OuterA" And contours.LoopRole(i + 8) = "Opening"
        Check "publication.paths.unchanged", a.SegmentID(i) = b.SegmentID(i) And b.SegmentID(i) = hole.SegmentID(i)
    Next i
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "MixedPublication"
    Set a = facade.RoundedPolygon("Mixed", NamedBoxPolygon(40#))
    facade.Outer "Outer", a.ID, namingPolicy:=SequentialByKind, namingPrefix:="Mixed"
    Set definition = facade.Finish: Set contours = definition.MaterialContours()
    Dim edge As TRegionEdge, kind As String
    For i = 1 To a.SegmentCount
        a.CopySegment i, edge
        If edge.Sweep = 0# Then kind = "LINE" Else kind = "ARC"
        Check "publication.mixed.common.counter", contours.ContourID(i) = "Mixed_" & kind & "_" & CStr(i)
        Check "publication.mixed.local.unchanged", a.SourceSegmentID(i) = a.SegmentID(i)
    Next i
End Sub

' ДЛЯ ТЕСТОВ: добавляем результат одной независимой проверки в журнал.
Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        mPassed = mPassed + 1: mReport = mReport & "PASS: " & name & vbCrLf
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
    End If
End Sub
