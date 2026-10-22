Attribute VB_Name = "modTestRecipeOperations"
Option Explicit

' ДЛЯ ТЕСТОВ: смещение и ограничение линий, а также понятные причины ошибок размещения арматуры.
Private mPassed As Long, mFailed As Long, mReport As String

Public Function RunRecipeOperationTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, source As CGeometryPath, definition As CSectionDefinition
    Dim edge As TRegionEdge, x As Double, y As Double, tx As Double, ty As Double, i As Long, query As CSectionGeometryQuery
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Operations"
    Set path = facade.RectanglePath("Concrete", -200#, -200#, 400#, 400#): facade.Outer "Outer", path.ID
    Set path = facade.LinePath("FullFace", 0#, 0#, 100#, 0#)
    Set path = facade.Offset("Axis", "FullFace", 20#)
    path.AtDistance 0#, x, y, tx, ty: Near "offset.full.start.x", x, 0#: Near "offset.normal.y", y, 20#
    Near "offset.full.length", path.Length, 100#
    Set path = facade.ConstrainLine("Extended", "Axis", 1#, 0#, -15#, 1#, 0#, 125#)
    path.AtDistance 0#, x, y, tx, ty: Near "constraint.extension.start", x, -15#: Near "constraint.extension.y", y, 20#
    Near "constraint.extension.length", path.Length, 140#
    Near "constraint.source.start", path.SourceStartParameter(1), -0.15
    Near "constraint.source.end", path.SourceEndParameter(1), 1.25
    Check "constraint.source.identity", path.SourcePathID(1) = "FullFace" And path.SourceSegmentID(1) = "FullFace"
    Set path = facade.ConstrainLine("Trimmed", "Axis", 1#, 0#, 10#, 1#, 0#, 80#)
    Near "constraint.trim.length", path.Length, 70#
    Set source = New CGeometryPath: source.Initialize "Corner"
    source.AddLine "Bottom", 0#, 0#, 100#, 0#: source.AddLine "Right", 100#, 0#, 100#, 100#: source.Seal
    facade.AddPath source: Set path = facade.Offset("Miter", "Corner", 10#)
    path.CopySegment 1, edge: Near "miter.x", edge.X2, 90#: Near "miter.y", edge.Y2, 10#
    path.CopySegment 2, edge: Near "miter.continuous.x", edge.X1, 90#: Near "miter.continuous.y", edge.Y1, 10#
    Near "miter.length", path.Length, 180#
    Near "miter.source.first.end", path.SourceEndParameter(1), 0.9
    Near "miter.source.second.start", path.SourceStartParameter(2), 0.1
    Set source = New CGeometryPath: source.Initialize "Smooth"
    source.AddLine "Straight", 0#, 0#, 100#, 0#
    source.AddArc "Fillet", 100#, 50#, 50#, -GEOM_PI / 2#, GEOM_PI / 2#: source.Seal
    facade.AddPath source: Set path = facade.Offset("SmoothOffset", source.ID, 10#)
    Near "offset.analytic.length", path.Length, 100# + 20# * GEOM_PI
    path.CopySegment 2, edge: Near "offset.arc.radius", Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2), 40#
    Set path = facade.CirclePath("Circle", 0#, 0#, 100#): Set path = facade.Offset("CircleInner", path.ID, 20#)
    Near "offset.circle.length", path.Length, 160# * GEOM_PI: Check "offset.circle.closed", path.IsClosed
    Set source = facade.Transform("MovedFace", "FullFace", 500#, -200#, GEOM_PI / 2#, True)
    facade.Anchor "CornerAnchor", 100#, 0#: facade.TransformAnchor "MovedAnchor", "CornerAnchor", 500#, -200#, GEOM_PI / 2#, True
    Set definition = facade.Finish: definition.AnchorCoordinates "MovedAnchor", x, y
    Near "transform.anchor.x", x, 500#: Near "transform.anchor.y", y, -100#
    source.AtDistance source.Length, x, y, tx, ty
    Near "transform.path.anchor.x", x, 500#: Near "transform.path.anchor.y", y, -100#
    Set query = New CSectionGeometryQuery
    Set source = New CGeometryPath: source.Initialize "CCW"
    source.AddLine "Bottom", 0#, 0#, 100#, 0#: source.AddLine "Right", 100#, 0#, 100#, 100#
    source.AddLine "Top", 100#, 100#, 0#, 100#: source.AddLine "Left", 0#, 100#, 0#, 0#: source.Seal True
    Set path = query.OffsetPath("Inset", source, 10#): Near "offset.closed.area", query.RegionFromPath(path).Area, 6400#
    Set source = source.Reversed: Set path = query.OffsetPath("OutsetCW", source, 10#)
    Near "offset.direction.left", query.RegionFromPath(path).Area, 14400#
    For i = 0 To 4: Check "operations.invalid." & CStr(i), InvalidOperation(i): Next i
    GoTo Finished
Failed:
    Check "operations.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRecipeOperationTests = mReport & "TOTAL RECIPE OPERATIONS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Private Function InvalidOperation(ByVal scenario As Long) As Boolean
    On Error GoTo Rejected
    Dim query As CSectionGeometryQuery, path As CGeometryPath, result As CGeometryPath, diagnostic As CGeometryDiagnostic, expected As String
    Set query = New CSectionGeometryQuery: Set path = New CGeometryPath: path.Initialize "Source"
    Select Case scenario
        Case 0
            expected = "CONSTRAINT_PARALLEL": path.AddLine "Line", 0#, 0#, 100#, 0#: path.Seal
            Set result = query.ConstrainLine("Failed", path, 0#, 1#, 10#, 1#, 0#, 90#)
        Case 1
            expected = "CONSTRAINT_DISAPPEARED": path.AddLine "Line", 0#, 0#, 100#, 0#: path.Seal
            Set result = query.ConstrainLine("Failed", path, 1#, 0#, 90#, 1#, 0#, 10#)
        Case 2
            expected = "OFFSET_DISAPPEARED": path.AddArc "Arc", 0#, 0#, 10#, 0#, GEOM_PI: path.Seal
            Set result = query.OffsetPath("Failed", path, 20#)
        Case 3
            expected = "CONSTRAINT_TARGET": path.AddArc "Arc", 0#, 0#, 10#, 0#, GEOM_PI: path.Seal
            Set result = query.ConstrainLine("Failed", path, 1#, 0#, 0#, 1#, 0#, 10#)
        Case 4
            expected = "OFFSET_JOINT": path.AddLine "Line", 0#, 0#, 100#, 0#
            path.AddArc "NonTangent", 100#, 10#, 10#, -GEOM_PI / 2#, -GEOM_PI / 2#: path.Seal
            Set result = query.OffsetPath("Failed", path, 1#)
    End Select
    Exit Function
Rejected:
    If Err.Number <> vbObjectError + 5401 Then Exit Function
    Set diagnostic = query.LastDiagnostic
    InvalidOperation = (Not diagnostic Is Nothing)
    If InvalidOperation Then InvalidOperation = diagnostic.Code = expected And diagnostic.EntityID = "Failed" And Len(diagnostic.RelatedEntityID) > 0
    If scenario = 2 And InvalidOperation Then InvalidOperation = diagnostic.Detail1 = 10# And diagnostic.Detail2 = 20#
    Err.Clear
End Function

Public Function RunRecipeDiagnosticTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, definition As CSectionDefinition, options As CRebarRuleOptions
    Dim rules As CRebarRuleExecutor, layout As CRebarLayout, diagnostic As CGeometryDiagnostic, i As Long, mode As Long
    For mode = 0 To 2
        Set facade = DiagnosticFacade: Set options = New CRebarRuleOptions
        If mode = 2 Then options.OverlapPolicy = RejectOverlap
        facade.Source "Bars", "ProgramInputs.Spacing"
        facade.RebarAtStations "Bars", "Row", Array(10#, IIf(mode = 0, 25#, 20#)), 20#, options
        Set definition = facade.Finish: Set rules = New CRebarRuleExecutor
        If mode = 2 Then
            On Error Resume Next
            Set layout = FixtureAnnotatedLayout(definition, rules)
            i = Err.Number: Err.Clear
            On Error GoTo Failed
            Set diagnostic = rules.LastDiagnostic
            Check "overlap.reject", i = vbObjectError + 5530 And diagnostic.Code = "BAR_OVERLAP"
            Check "overlap.source", diagnostic.EntityID = "Bars" And diagnostic.SourceReference = "ProgramInputs.Spacing"
            Near "overlap.numeric.gap", diagnostic.Detail1, -10#
            Check "overlap.error.actual.created", rules.RequestedPositions = 2 And rules.CreatedPositions = 1 And rules.Conflicts = 1
        Else
            Set layout = FixtureAnnotatedLayout(definition, rules)
            Check "overlap.physical." & CStr(mode), layout.Count = 2 And rules.RequestedPositions = 2 And rules.CreatedPositions = 2
            Check "overlap.report." & CStr(mode), rules.BarOverlaps = 1
            Check "overlap.machine." & CStr(mode), rules.DiagnosticAt(1).Code = "BAR_OVERLAP"
            Check "trace.rule." & CStr(mode), InStr(rules.RuleTrace("Bars"), "Created=2") > 0 And InStr(rules.RuleTrace("Bars"), "StationStepMax=") > 0
            Check "trace.owner." & CStr(mode), rules.PositionPhysicalOwner(1) = 1 And rules.PositionPhysicalOwner(2) = 2
            Check "trace.ids." & CStr(mode), rules.PositionRuleID(1) = "Bars" And rules.PositionPathID(1) = "Row" And rules.PositionSourceID(1) = "Bars.1"
        End If
    Next mode
    Set facade = DiagnosticFacade
    facade.RebarAtStations "Contact", "Row", Array(10#, 30#), 20#
    Set definition = facade.Finish: Set rules = New CRebarRuleExecutor: Set layout = FixtureAnnotatedLayout(definition, rules)
    Check "contact.allowed", rules.BarOverlaps = 0 And layout.Count = 2
    Set facade = DiagnosticFacade
    facade.RebarByCount "Base", "Row", 3, 10#
    Set path = facade.LinePath("Target", 20#, 50#, 80#, 50#)
    Set options = New CRebarRuleOptions: options.DirectionY = 1#: options.SkipMissingProjection = True
    facade.Source "Projected", "ProgramInputs.Target"
    facade.RebarProjected "Projected", "Target", "Base", 10#, options
    Set definition = facade.Finish: Set rules = New CRebarRuleExecutor: Set layout = FixtureAnnotatedLayout(definition, rules)
    Check "projection.counts", rules.RequestedPositions = 6 And rules.CreatedPositions = 4 And rules.SkippedProjections = 2
    Check "projection.trace", InStr(rules.RuleTrace("Projected"), "Requested=3; Created=1; Merged=0; Skipped=2") > 0
    Check "projection.machine", rules.DiagnosticCount = 2 And rules.DiagnosticAt(1).Code = "PROJECTION_MISSING"
    Check "projection.related", rules.DiagnosticAt(1).RelatedEntityID = "Base.1" And rules.DiagnosticAt(2).RelatedEntityID = "Base.3"
    Check "projection.source", rules.DiagnosticAt(1).SourceReference = "ProgramInputs.Target"
    Set facade = DiagnosticFacade: Set path = facade.CirclePath("Ring", 0#, 0#, 40#)
    Set options = New CRebarRuleOptions: options.EndMode = ClosedUniform
    facade.RebarAtStations "RingBars", path.ID, Array(0#, 10#, 20#), 4#, options
    Set definition = facade.Finish: Set rules = New CRebarRuleExecutor: Set layout = FixtureAnnotatedLayout(definition, rules)
    Dim stepCount As Long, minStep As Double, maxStep As Double, minChord As Double, maxChord As Double
    rules.RuleStepBounds "RingBars", stepCount, minStep, maxStep, minChord, maxChord
    Check "trace.closed.count", stepCount = 3
    Near "trace.closed.minimum", minStep, 10#
    Near "trace.closed.wrap.maximum", maxStep, 80# * GEOM_PI - 20#
    Near "trace.closed.chord", maxChord, 80# * Sin(0.25)
    Set facade = DiagnosticFacade
    Dim dense() As Double
    ReDim dense(0 To 99)
    For i = 0 To 99: dense(i) = CDbl(i): Next i
    facade.RebarAtStations "Dense", "Row", dense, 20#
    Set definition = facade.Finish: Set rules = New CRebarRuleExecutor: Set layout = FixtureAnnotatedLayout(definition, rules)
    Check "diagnostics.bounded.full.counts", rules.CreatedPositions = 100 And rules.BarOverlaps = 1710
    Check "diagnostics.bounded.records", rules.DiagnosticCount = 128 And rules.DiagnosticRecordsOmitted = 1582
    Check "diagnostics.bounded.latest", rules.LastDiagnostic.Code = "BAR_OVERLAP" And rules.LastDiagnostic.EntityID = "Dense"
    Set facade = DiagnosticFacade: Set options = New CRebarRuleOptions: options.Coincidence = MergeIdentical
    facade.RebarAtStations "First", "Row", Array(50#), 10#
    facade.RebarAtStations "Second", "Row", Array(50#), 10#, options
    Set definition = facade.Finish: Set rules = New CRebarRuleExecutor: Set layout = FixtureAnnotatedLayout(definition, rules)
    Check "merge.counts", rules.RequestedPositions = 2 And rules.CreatedPositions = 1 And rules.MergedPositions = 1
    Check "merge.physical.owner", rules.PositionPhysicalOwner(1) = 1 And rules.PositionPhysicalOwner(2) = 1
    Check "merge.trace", InStr(rules.RuleTrace("Second"), "Created=0; Merged=1") > 0
    ' Неподдерживаемый способ распределения вызывает ошибку без произвольной подстановки.
    Set facade = DiagnosticFacade: Set options = New CRebarRuleOptions: options.DistributionScope = 2
    On Error Resume Next
    facade.RebarByCount "Unsupported", "Row", 3, 10#, options
    i = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "scope.invalid", i = vbObjectError + 5530
    On Error Resume Next
    Set layout = rules.Build(Nothing, Nothing)
    i = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "diagnostics.failed.reuse", i = vbObjectError + 5530 And rules.LastDiagnostic.Code = "RULE_INVALID"
    Check "diagnostics.failed.reuse.counts", rules.RequestedPositions = 0 And rules.CreatedPositions = 0 And rules.MergedPositions = 0 And rules.SkippedProjections = 0
    GoTo Finished
Failed:
    Check "diagnostics.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRecipeDiagnosticTests = mReport & "TOTAL RECIPE DIAGNOSTICS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

Private Function DiagnosticFacade() As CSectionDefinitionBuilder
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Diagnostics"
    Set path = facade.RectanglePath("Concrete", -100#, -100#, 300#, 300#): facade.Outer "Outer", path.ID
    Set path = facade.LinePath("Row", 0#, 0#, 100#, 0#)
    Set DiagnosticFacade = facade
End Function

' ДЛЯ ТЕСТОВ: поиск по координатам не изменяет линии; выделение участка через Chain сохраняет
' свои правила.
Public Function RunResolveAnchorTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, polygon As CPolygonSpec
    Dim rightID As String, t As Double, right As CGeometryPath, x As Double, y As Double, tx As Double, ty As Double
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "ResolveAnchor"
    Set path = facade.RectanglePath("Rect", 0#, 0#, 400#, 600#)
    path.ResolveAnchor 400#, 300#, rightID, t
    Near "rectangle.right.parameter", t, 0.5
    Set right = facade.Chain("Right", path.ID, rightID, 1#, rightID, 0#, False)
    Near "rectangle.right.length", right.Length, 600#
    right.AtDistance 0#, x, y, tx, ty: Near "rectangle.right.start.x", x, 400#: Near "rectangle.right.start.y", y, 600#
    right.AtDistance right.Length, x, y, tx, ty: Near "rectangle.right.end.y", y, 0#: Check "rectangle.right.direction", ty < 0#
    CheckResolvedPath facade, path
    Set path = facade.AdjacentRectangles("Adjacent", Array(Array(0#, 0#, 400#, 200#), Array(0#, 200#, 200#, 400#)))
    CheckResolvedPath facade, path
    Set polygon = New CPolygonSpec
    polygon.AddVertex "BL", 0#, 0#, 30#: polygon.AddVertex "BR", 400#, 0#, 40#
    polygon.AddVertex "TR", 400#, 600#, 50#: polygon.AddVertex "TL", 0#, 600#, 60#: polygon.Seal
    Set path = facade.RoundedPolygon("Rounded", polygon): CheckResolvedPath facade, path
    Set path = facade.OffsetRoundedPolygon("RoundedOffset", polygon, 10#): CheckResolvedPath facade, path
    Set path = facade.CirclePath("Circle", 25#, -15#, 100#): CheckResolvedPath facade, path
    Set path = facade.ArcPath("ArcCCW", 15#, 30#, 80#, 0.9 * GEOM_PI, 0.8 * GEOM_PI): CheckResolvedPath facade, path
    ExpectAnchorError path, 15# + 80#, 30#, "ANCHOR_NOT_ON_PATH"
    Set path = facade.ArcPath("ArcCW", -10#, 40#, 80#, -0.9 * GEOM_PI, -0.8 * GEOM_PI): CheckResolvedPath facade, path
    Set path = facade.LinePath("Line", 0#, 0#, 100#, 0#): CheckResolvedPath facade, path
    ' Преобразования проверяются и для прямых, и для скруглений; отражение меняет знак дуги.
    Set path = facade.Transform("MovedRect", "Rect", 1000#, -700#, GEOM_PI / 3#, True): CheckResolvedPath facade, path
    Set path = facade.Transform("MovedRounded", "Rounded", -900#, 700#, GEOM_PI / 3#, True): CheckResolvedPath facade, path
    Set path = facade.Transform("MovedCircle", "Circle", 900#, -1700#, -GEOM_PI / 4#, True): CheckResolvedPath facade, path
    Set path = facade.Transform("MovedArc", "ArcCW", -900#, -700#, GEOM_PI / 4#): CheckResolvedPath facade, path
    ' Реальный общий допуск границы допускает малую поперечную погрешность, но не ближайшую произвольную точку.
    Set path = facade.LinePath("Tolerance", 0#, 0#, 100#, 0#)
    path.ResolveAnchor 50#, GEOM_BOUNDARY_TOLERANCE / 2#, rightID, t: Near "tolerance.accepted", t, 0.5
    ExpectAnchorError path, 50#, GEOM_BOUNDARY_TOLERANCE * 2#, "ANCHOR_NOT_ON_PATH"
    Set path = New CGeometryPath: path.Initialize "Retraced"
    path.AddLine "Out", 0#, 0#, 100#, 0#: path.AddLine "Back", 100#, 0#, 0#, 0#: path.Seal
    ExpectAnchorError path, 50#, 0#, "ANCHOR_AMBIGUOUS"
    Set path = New CGeometryPath: path.Initialize "Unsealed": path.AddLine "Only", 0#, 0#, 100#, 0#
    ExpectAnchorError path, 50#, 0#, "PATH_INVALID"
    TestPositionOrderOwner
    GoTo Finished
Failed:
    Check "anchor.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunResolveAnchorTests = mReport & "TOTAL RESOLVE ANCHOR: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' Проверяет сегменты, оба направления обхода Chain, длины, концы и соответствие участкам
' исходных линий.
Private Sub CheckResolvedPath(ByVal facade As CSectionDefinitionBuilder, ByVal path As CGeometryPath)
    Dim i As Long, edge As TRegionEdge, segment As String, parameter As Double, x As Double, y As Double
    Dim chain As CGeometryPath, station As Double, tx As Double, ty As Double, a As Double, b As Double
    Dim expectedX As Double, expectedY As Double, label As String, reverse As Long
    station = 0#
    For i = 1 To path.SegmentCount
        label = path.ID & "." & CStr(i): path.CopySegment i, edge
        modGeometryPrimitives.EdgePoint edge, 0.25, x, y
        path.ResolveAnchor x, y, segment, parameter
        Check label & ".id", segment = path.SegmentID(i): Near label & ".parameter", parameter, 0.25
        For reverse = 0 To 1
            a = 0.25: b = 0.75: If reverse = 1 Then a = 0.75: b = 0.25
            Set chain = facade.Chain(label & ".Chain" & CStr(reverse), path.ID, segment, a, segment, b, reverse = 0)
            Check label & ".chain.id", chain.SegmentCount = 1 And chain.SegmentID(1) = segment
            Check label & ".source.id", chain.SourcePathID(1) = path.SourcePathID(i) And chain.SourceSegmentID(1) = path.SourceSegmentID(i)
            Near label & ".source.start", chain.SourceStartParameter(1), path.SourceStartParameter(i) + a * (path.SourceEndParameter(i) - path.SourceStartParameter(i))
            Near label & ".source.end", chain.SourceEndParameter(1), path.SourceStartParameter(i) + b * (path.SourceEndParameter(i) - path.SourceStartParameter(i))
            Near label & ".length", chain.Length, modGeometryPrimitives.EdgeLength(edge) / 2#
            modGeometryPrimitives.EdgePoint edge, a, expectedX, expectedY: chain.AtDistance 0#, x, y, tx, ty
            Near label & ".start.x", x, expectedX: Near label & ".start.y", y, expectedY
            modGeometryPrimitives.EdgePoint edge, b, expectedX, expectedY: chain.AtDistance chain.Length, x, y, tx, ty
            Near label & ".end.x", x, expectedX: Near label & ".end.y", y, expectedY
            modGeometryPrimitives.EdgeTangent edge, b, expectedX, expectedY
            If reverse = 1 Then expectedX = -expectedX: expectedY = -expectedY
            Check label & ".direction", tx * expectedX + ty * expectedY > 0#
        Next reverse
        If i > 1 Or path.IsClosed And path.SegmentCount > 1 Then
            ExpectAnchorError path, edge.X1, edge.Y1, "ANCHOR_AMBIGUOUS"
        Else
            path.ResolveAnchor edge.X1, edge.Y1, segment, parameter: Near label & ".start.parameter", parameter, 0#
        End If
        If i = path.SegmentCount And Not path.IsClosed Then
            path.ResolveAnchor edge.X2, edge.Y2, segment, parameter: Near label & ".end.parameter", parameter, 1#
        End If
        station = station + modGeometryPrimitives.EdgeLength(edge)
    Next i
    Near path.ID & ".unchanged.length", path.Length, station
    ExpectAnchorError path, 100000#, -100000#, "ANCHOR_NOT_ON_PATH"
    If path.SegmentCount > 1 Then
        Dim first As TRegionEdge, last As TRegionEdge, startID As String, endID As String, startT As Double, endT As Double
        path.CopySegment 1, first: path.CopySegment 2, last
        modGeometryPrimitives.EdgePoint first, 0.75, x, y: path.ResolveAnchor x, y, startID, startT
        modGeometryPrimitives.EdgePoint last, 0.25, x, y: path.ResolveAnchor x, y, endID, endT
        Set chain = facade.Chain(path.ID & ".JointChain", path.ID, startID, startT, endID, endT)
        Check path.ID & ".joint.chain.ids", chain.SegmentCount = 2 And chain.SegmentID(1) = startID And chain.SegmentID(2) = endID
        Near path.ID & ".joint.chain.length", chain.Length, (modGeometryPrimitives.EdgeLength(first) + modGeometryPrimitives.EdgeLength(last)) / 4#
        Check path.ID & ".joint.chain.origin", chain.SourceSegmentID(1) = path.SourceSegmentID(1) And chain.SourceSegmentID(2) = path.SourceSegmentID(2)
        modGeometryPrimitives.EdgePoint first, 0.75, expectedX, expectedY: chain.AtDistance 0#, x, y, tx, ty
        Near path.ID & ".joint.start.x", x, expectedX: Near path.ID & ".joint.start.y", y, expectedY
        modGeometryPrimitives.EdgePoint last, 0.25, expectedX, expectedY: chain.AtDistance chain.Length, x, y, tx, ty
        Near path.ID & ".joint.end.x", x, expectedX: Near path.ID & ".joint.end.y", y, expectedY
        If path.IsClosed Then
            path.CopySegment path.SegmentCount, last
            modGeometryPrimitives.EdgePoint last, 0.75, x, y: path.ResolveAnchor x, y, startID, startT
            modGeometryPrimitives.EdgePoint first, 0.25, x, y: path.ResolveAnchor x, y, endID, endT
            Set chain = facade.Chain(path.ID & ".WrapChain", path.ID, startID, startT, endID, endT)
            Check path.ID & ".wrap.ids", chain.SegmentCount = 2 And chain.SegmentID(1) = startID And chain.SegmentID(2) = endID
            Near path.ID & ".wrap.length", chain.Length, (modGeometryPrimitives.EdgeLength(first) + modGeometryPrimitives.EdgeLength(last)) / 4#
            Check path.ID & ".wrap.origin", chain.SourceSegmentID(1) = path.SourceSegmentID(path.SegmentCount) And chain.SourceSegmentID(2) = path.SourceSegmentID(1)
            modGeometryPrimitives.EdgePoint last, 0.75, expectedX, expectedY: chain.AtDistance 0#, x, y, tx, ty
            Near path.ID & ".wrap.start.x", x, expectedX: Near path.ID & ".wrap.start.y", y, expectedY
            modGeometryPrimitives.EdgePoint first, 0.25, expectedX, expectedY: chain.AtDistance chain.Length, x, y, tx, ty
            Near path.ID & ".wrap.end.x", x, expectedX: Near path.ID & ".wrap.end.y", y, expectedY
        End If
    End If
End Sub

' Ошибка оставляет переданные выходные значения неизменными и содержит координаты точки.
Private Sub ExpectAnchorError(ByVal path As CGeometryPath, ByVal x As Double, ByVal y As Double, ByVal code As String)
    Dim segment As String, parameter As Double, number As Long, reason As String, diagnostic As CGeometryDiagnostic
    segment = "unchanged": parameter = -1#
    On Error Resume Next
    path.ResolveAnchor x, y, segment, parameter: number = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo 0
    Set diagnostic = path.LastDiagnostic
    Check path.ID & ".error.number", number = vbObjectError + 5520 And InStr(reason, path.ID) > 0
    Check path.ID & ".error.outputs", segment = "unchanged" And parameter = -1#
    Check path.ID & ".error.diagnostic", Not diagnostic Is Nothing
    If Not diagnostic Is Nothing Then
        Check path.ID & ".error.code", diagnostic.Code = code And diagnostic.EntityID = path.ID
        If code <> "PATH_INVALID" Then
            Near path.ID & ".error.x", diagnostic.Detail1, x: Near path.ID & ".error.y", diagnostic.Detail2, y
        End If
        If code = "ANCHOR_AMBIGUOUS" Then Check path.ID & ".error.segments", InStr(diagnostic.RelatedEntityID, ", ") > 0
    End If
End Sub

' Контрпример: при совпадении стержней KeepPublishedOwner сохраняет диаметр того, который
' добавлен первым.
Private Sub TestPositionOrderOwner()
    Dim facade As CSectionDefinitionBuilder, options As CRebarRuleOptions, path As CGeometryPath
    Dim definition As CSectionDefinition, rules As CRebarRuleExecutor, layout As CRebarLayout, mode As Long, owner As Long
    For mode = AlongPath To FromPathMiddle
        Set facade = New CSectionDefinitionBuilder: facade.Initialize "PublicationOwner"
        Set path = facade.RectanglePath("Concrete", -50#, -50#, 150#, 100#): facade.Outer "Outer", path.ID
        Set path = facade.LinePath("Row", 0#, 0#, 30#, 0#)
        Set options = New CRebarRuleOptions
        options.PositionOrder = mode: options.Publication = PublishPerSourcePosition: options.Coincidence = KeepPublishedOwner
        facade.RebarByCount "Main", path.ID, 3, 10#, options
        Set options = New CRebarRuleOptions
        options.Publication = PublishPerSourcePosition: options.Coincidence = KeepPublishedOwner: options.TangentOffset = 15#
        facade.RebarDependent "Extra", "Main", 20#, options
        Set definition = facade.Finish: Set rules = New CRebarRuleExecutor: Set layout = FixtureAnnotatedLayout(definition, rules)
        owner = layout.FindBarAt(15#, 0#): Check "position.order.owner.exists", owner > 0
        If owner > 0 Then
            If mode = AlongPath Then Near "position.order.along.diameter", layout.Diameter(owner), 20# Else Near "position.order.middle.diameter", layout.Diameter(owner), 10#
        End If
    Next mode
End Sub

Private Sub Check(ByVal name As String, ByVal value As Boolean)
    If value Then mPassed = mPassed + 1 Else mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
End Sub

Private Sub Near(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    Check name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) <= 0.000000001 * (1# + Abs(expected))
End Sub

' ДЛЯ ТЕСТОВ: исчезновение угловой дуги сохраняет точный стык и имена привязок.
Public Function RunCollapsingOffsetTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim source As CGeometryPath, result As CGeometryPath, query As CSectionGeometryQuery
    Dim distance As Variant, x As Double, y As Double, tx As Double, ty As Double, polygon As CPolygonSpec
    Set query = New CSectionGeometryQuery
    Set source = New CGeometryPath: source.Initialize "RoundedCorner"
    source.AddLine "First", 0#, 0#, 100#, 0#
    source.AddArc "Corner", 100#, 40#, 40#, -GEOM_PI / 2#, GEOM_PI / 2#
    source.AddLine "Last", 140#, 40#, 140#, 140#: source.Seal
    For Each distance In Array(40#, 50#, 80#)
        Set result = query.OffsetPath("CornerOffset", source, CDbl(distance))
        Check "collapse.segmentCount." & CStr(distance), result.SegmentCount = 2
        Check "collapse.anchor." & CStr(distance), result.HasCollapsedSegmentID("Corner") And Not result.HasSegmentID("Corner")
        result.AtDistance result.AnchorStation("Corner"), x, y, tx, ty
        Near "collapse.joint.x." & CStr(distance), x, 140# - CDbl(distance)
        Near "collapse.joint.y." & CStr(distance), y, CDbl(distance)
        Check "collapse.joint.belongs." & CStr(distance), result.ContainsPoint(x, y)
        Near "collapse.length." & CStr(distance), result.Length, 280# - 2# * CDbl(distance)
    Next distance
    Set result = query.OffsetPath("BeforeCollapse", source, 39.99)
    Check "collapse.before.arcKept", result.HasSegmentID("Corner") And result.SegmentCount = 3
    Set result = query.OffsetPath("AfterCollapse", source, 50#)
    source.AtDistance 100# + 10# * GEOM_PI, x, y, tx, ty
    Check "collapse.inverted.point.rejected", Not result.ContainsPoint(x - 50# * ty, y + 50# * tx)
    Check "collapse.source.unchanged", source.SegmentCount = 3 And source.HasSegmentID("Corner")
    Set result = query.OffsetPath("ReverseOffset", source.Reversed, -50#)
    Near "collapse.reverse.length", result.Length, 180#
    Set source = source.Transformed(500#, -200#, GEOM_PI / 3#, True)
    Set result = query.OffsetPath("TransformedOffset", source, -50#)
    result.AtDistance result.AnchorStation("Corner"), x, y, tx, ty
    tx = 90#: ty = 50#: query.TransformCoordinates tx, ty, 500#, -200#, GEOM_PI / 3#, True
    Near "collapse.transform.x", x, tx: Near "collapse.transform.y", y, ty
    Set polygon = New CPolygonSpec
    polygon.AddVertex "A", 0#, 0#, 10#: polygon.AddVertex "B", 100#, 0#, 10#
    polygon.AddVertex "C", 100#, 100#, 10#: polygon.AddVertex "D", 0#, 100#, 10#: polygon.Seal
    Set source = query.RoundedPolygon("Closed", polygon)
    Set result = query.OffsetPath("ClosedOffset", source, 20#)
    Check "collapse.closed", result.IsClosed And result.SegmentCount = 4
    Near "collapse.closed.area", query.RegionFromPath(result).Area, 3600#
    GoTo Finished
Failed:
    Check "collapse.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunCollapsingOffsetTests = mReport & "TOTAL_COLLAPSING_OFFSETS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: сохранённый PointedSection при 40 мм сохраняет стержень в стыке,
' а при большем смещении пропускает только позицию, вышедшую за новую линию.
Public Function RunSavedPointedOffsetTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim table As Object, config As Object, saved As Variant, values As Variant, value As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim i As Long, count As Long, expected As Long, inputCell As Object
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = FixtureSaveInputs(config)
    PrepareSavedOffsetCase
    Set table = ThisWorkbook.Names("rngPointedSectionGeometry").RefersToRange
    Set inputCell = FixtureInputCell(table, "Extra", "Right", "a2")
    For Each value In Array(39.99, 40#, 50#, 80#)
        inputCell.Value2 = CDbl(value)
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
        count = 0
        For i = 1 To model.RebarCount
            If InStr(model.RebarRecipeTrace(i), "RuleId=H.Right.Row2") > 0 Then count = count + 1
        Next i
        expected = 5: If CDbl(value) > 40# Then expected = 4
        Check "saved.right.created." & CStr(value), count = expected
        Check "saved.right.warning." & CStr(value), SavedRuleSkipped(model.GeometryDiagnosticsText, "H.Right.Row2") = IIf(CDbl(value) > 40#, 1, 0)
        Check "saved.warning.summary." & CStr(value), (Len(model.GeometryWarningText) > 0) = (InStr(model.GeometryDiagnosticsText, "Предупреждение:") > 0)
        If CDbl(value) > 40# Then Check "saved.right.warning.address." & CStr(value), InStr(model.GeometryDiagnosticsText, inputCell.Worksheet.Name & "!" & inputCell.Address(False, False)) > 0
    Next value
    GoTo Finished
Failed:
    Check "saved.offset.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then FixtureRestoreInputs config, saved
    RunSavedPointedOffsetTests = mReport & "TOTAL_SAVED_POINTED_OFFSETS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: соседи прямой стороны ограничивают дополнительные ряды даже без дуг.
' Проверяет общий контур, обращённый обход, отражение, поворот и связь третьего ряда со вторым.
Public Function RunFullBoundaryOffsetTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim reverse As Long, reflect As Long, binding As Long, chain As Long
    Dim facade As CSectionDefinitionBuilder, boundary As CGeometryPath, whole As CGeometryPath, side As CGeometryPath
    Dim options As CRebarRuleOptions, query As CSectionGeometryQuery, executor As CRebarRuleExecutor, layout As CRebarLayout
    Dim definition As CSectionDefinition, actual As CGeometryPath, i As Long, row2 As Long, row3 As Long, normal As Double
    Dim x As Double, y As Double, prefix As String
    For reverse = 0 To 1
        For reflect = 0 To 1
            For binding = 0 To 1
                For chain = 0 To 1
                    Set facade = New CSectionDefinitionBuilder: facade.Initialize "WholeBoundaryTest"
                    Set boundary = facade.RectanglePath("Boundary", 0#, 0#, 1000#, 600#, "Bottom", "Right", "Top", "Left")
                    Set query = New CSectionGeometryQuery
                    Set whole = query.OffsetPath("Whole", boundary, 50#)
                    Set side = whole.BetweenAnchors("Side", "Bottom", 0#, "Bottom", 1#)
                    If reverse = 1 Then Set side = side.Reversed
                    Set side = side.WithID("Axis").Transformed(1700#, -900#, GEOM_PI / 6#, reflect = 1)
                    Set actual = boundary.Transformed(1700#, -900#, GEOM_PI / 6#, reflect = 1).WithID("Concrete")
                    facade.AddPath actual: facade.Outer "Outer", actual.ID
                    facade.AddPath side
                    Set options = New CRebarRuleOptions
                    options.AnnotationGroup = "Side"
                    facade.RebarByCount "Main", side.ID, 5, 10#, options
                    normal = (1# - 2# * reverse) * (1# - 2# * reflect)
                    options.NormalOffset = normal * 100#: options.EverySecond = (binding = 1)
                    options.SourceReference = "WholeBoundary[Extra,2]"
                    facade.RebarDependent "Row2", "Main", 10#, options
                    options.EverySecond = False: options.NormalOffset = normal * 200#
                    If chain = 1 Then options.NormalOffset = normal * 100#
                    options.SourceReference = "WholeBoundary[Extra,3]"
                    facade.RebarDependent "Row3", IIf(chain = 1, "Row2", "Main"), 10#, options
                    Set definition = facade.Finish: Set executor = New CRebarRuleExecutor
                    Set layout = FixtureAnnotatedLayout(definition, executor)
                    row2 = 0: row3 = 0
                    For i = 1 To layout.Count
                        x = layout.X(i) - 1700#: y = layout.Y(i) + 900#
                        ' Обратный поворот и отражение дают координаты до преобразования.
                        query.TransformCoordinates x, y, 0#, 0#, -GEOM_PI / 6#
                        If reflect = 1 Then y = -y
                        If Abs(y - 150#) < 0.000001 Then
                            row2 = row2 + 1: Check "whole.row2.bounds", x >= 150# And x <= 850#
                        ElseIf Abs(y - 250#) < 0.000001 Then
                            row3 = row3 + 1: Check "whole.row3.bounds", x >= 250# And x <= 750#
                        End If
                    Next i
                    prefix = "whole." & CStr(reverse) & CStr(reflect) & CStr(binding) & CStr(chain)
                    Check prefix & ".row2", row2 = IIf(binding = 1, 1, 3)
                    Check prefix & ".row3", row3 = IIf(binding = 1 And chain = 1, 1, 3)
                    Check prefix & ".warning", InStr(executor.DiagnosticsText, "Предупреждение:") > 0
                    Check prefix & ".address", InStr(executor.DiagnosticsText, "WholeBoundary[Extra,2]") > 0
                    Check prefix & ".source.unchanged", whole.IsClosed And whole.SegmentCount = 4 And side.SegmentCount = 1
                Next chain
            Next binding
        Next reflect
    Next reverse
    GoTo Finished
Failed:
    Check "whole.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunFullBoundaryOffsetTests = mReport & "TOTAL_FULL_BOUNDARY_OFFSETS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: обычный расчёт сохранённого примера проходит через запись Results и схему.
' Предупреждение о пропуске стержня проверяется в ответе формирования и в Results.
Public Function RunSavedOffsetCalculationTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim area As Object, config As Object, inputCell As Object, saved As Variant, value As Variant, message As String
    Dim summary As Object, i As Long, j As Long, text As String, diagnostics As String, formation As String
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = FixtureSaveInputs(config)
    PrepareSavedOffsetCase
    Set area = ThisWorkbook.Names("rngPointedSectionGeometry").RefersToRange
    Set inputCell = FixtureInputCell(area, "Extra", "Right", "a2")
    For Each value In Array(40#, 50#)
        inputCell.Value2 = value
        formation = FormSectionForWorkbook(ThisWorkbook)
        message = RunSectionCalculationForWorkbook(ThisWorkbook, True)
        Check "saved.calculation.completed." & CStr(value), InStr(message, "завершен.") > 0 And InStr(message, "ошиб") = 0
        diagnostics = SavedRecipeDiagnostics()
        Check "saved.calculation.warning." & CStr(value), (InStr(formation, "Предупреждение:") > 0) = (InStr(diagnostics, "Предупреждение:") > 0)
        Check "saved.calculation.rowWarning." & CStr(value), SavedRuleSkipped(diagnostics, "H.Right.Row2") = IIf(CDbl(value) > 40#, 1, 0)
        Set summary = ThisWorkbook.Names("rngBatchSummary").RefersToRange
        text = ""
        For i = 1 To summary.Rows.Count
            For j = 1 To summary.Columns.Count
                text = text & CStr(summary.Cells(i, j).Value2) & "|"
            Next j
        Next i
        Check "saved.calculation.status." & CStr(value), InStr(text, "NumFail") = 0 And InStr(text, "InputErr") = 0 And InStr(text, "CalcErr") = 0
        mReport = mReport & "OBSERVED_SAVED_CALCULATION: a2=" & CStr(value) & "; " & Replace(message, vbCrLf, " | ") & vbCrLf
    Next value
    GoTo Finished
Failed:
    Check "saved.calculation.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then FixtureRestoreInputs config, saved
    RunSavedOffsetCalculationTests = mReport & "TOTAL_SAVED_OFFSET_CALCULATION: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: воспроизводит пример R80/as40 без дополнительных рядов на соседних сторонах.
' Их независимые отступы не относятся к проверке исчезновения правого скругления.
Private Sub PrepareSavedOffsetCase()
    Dim table As Object, key As Variant
    ThisWorkbook.Names("rngUnitSettings").RefersToRange.Cells(2, 2).Value2 = "mm"
    modTestShapeArchitecture.ConfigureShapeFixture "RoundedMixed"
    Set table = ThisWorkbook.Names("rngPointedSectionGeometry").RefersToRange
    FixtureInputCell(table, "General", "B").Value2 = 1000#: FixtureInputCell(table, "General", "H").Value2 = 500#
    FixtureInputCell(table, "General", "Left.W").Value2 = 0#: FixtureInputCell(table, "General", "Left.R1").Value2 = 0#
    FixtureInputCell(table, "General", "Left.R2").Value2 = 0#: FixtureInputCell(table, "General", "Right.W").Value2 = 150#
    FixtureInputCell(table, "General", "Right.R1").Value2 = 0#: FixtureInputCell(table, "General", "Right.R2").Value2 = 80#
    For Each key In Array("Top", "Bottom", "Left", "Right")
        FixtureInputCell(table, "Main", CStr(key), "as").Value2 = 40#
        FixtureInputCell(table, "Main", CStr(key), "d").Value2 = 25#
        FixtureInputCell(table, "Main", CStr(key), "n").Value2 = IIf(CStr(key) = "Top" Or CStr(key) = "Bottom", 10, 5)
        FixtureInputCell(table, "Main", CStr(key), "s0").Value2 = 0#: FixtureInputCell(table, "Main", CStr(key), "s1").Value2 = 0#
        FixtureInputCell(table, "Extra", CStr(key), "d2").Value2 = IIf(CStr(key) = "Right", 16#, 0#)
        FixtureInputCell(table, "Extra", CStr(key), "d3").Value2 = 0#
    Next key
    FixtureInputCell(table, "Extra", "Right", "pos2").Value2 = "Stacked"
    FixtureInputCell(table, "Extra", "Right", "bind2").Value2 = "EachBar"
End Sub

' ДЛЯ ТЕСТОВ: читает полный счётчик пропусков конкретного правила.
' Предупреждение другого ряда не означает, что проверяемый ряд потерял стержень.
Private Function SavedRuleSkipped(ByVal diagnostics As String, ByVal ruleID As String) As Long
    Dim line As Variant, prefix As String, marker As String, position As Long
    prefix = "Rule=" & ruleID & ";": marker = "; Skipped="
    For Each line In Split(diagnostics, vbCrLf)
        If Left$(CStr(line), Len(prefix)) = prefix Then
            position = InStr(CStr(line), marker)
            If position > 0 Then
                SavedRuleSkipped = CLng(Split(Mid$(CStr(line), position + Len(marker)), ";")(0))
                Exit Function
            End If
        End If
    Next line
    Err.Raise 5, "modTestRecipeOperations", "В диагностике отсутствует счётчик пропусков ряда " & ruleID
End Function

' ДЛЯ ТЕСТОВ: берёт диагностику из записанного Results после штатного расчёта.
Private Function SavedRecipeDiagnostics() As String
    Dim anchor As Object, row As Long
    Set anchor = ThisWorkbook.Names("rngNDMSectionProperties").RefersToRange
    row = 2
    Do While Len(CStr(anchor.Cells(row, 1).Value2)) > 0
        If CStr(anchor.Cells(row, 3).Value2) = "Geometry.RecipeDiagnostics" Then
            SavedRecipeDiagnostics = CStr(anchor.Cells(row, 4).Value2): Exit Function
        End If
        row = row + 1
    Loop
    Err.Raise 5, "modTestRecipeOperations", "В Results отсутствует сохранённая диагностика геометрии."
End Function
