Attribute VB_Name = "modTestRecipeOperations"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: смещение и ограничение линий, а также понятные причины ошибок размещения арматуры.
' ==========================================================================
Private mPassed As Long, mFailed As Long, mReport As String

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет смещение, обрезку и продление путей, сопряжение угла, преобразование привязки и влияние направления обхода.
Public Function RunRecipeOperationTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, source As CGeometryPath, definition As CSectionDefinition
    Dim edge As TRegionEdge, x As Double, y As Double, tx As Double, ty As Double, i As Long, query As CSectionGeometryQuery
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Operations"
    Set path = facade.RectanglePath("Concrete", -200#, -200#, 400#, 400#): facade.Outer "Outer", path.ID
    Set path = facade.LinePath("FullFace", 0#, 0#, 100#, 0#)
    Set path = facade.Offset("Axis", "FullFace", 20#)
    path.AtDistance 0#, x, y, tx, ty: TestNear "offset.full.start.x", x, 0#: TestNear "offset.normal.y", y, 20#
    TestNear "offset.full.length", path.Length, 100#
    Set path = facade.ConstrainLine("Extended", "Axis", 1#, 0#, -15#, 1#, 0#, 125#)
    path.AtDistance 0#, x, y, tx, ty: TestNear "constraint.extension.start", x, -15#: TestNear "constraint.extension.y", y, 20#
    TestNear "constraint.extension.length", path.Length, 140#
    TestNear "constraint.source.start", path.SourceStartParameter(1), -0.15
    TestNear "constraint.source.end", path.SourceEndParameter(1), 1.25
    TestCheck "constraint.source.identity", path.SourcePathID(1) = "FullFace" And path.SourceSegmentID(1) = "FullFace"
    Set path = facade.ConstrainLine("Trimmed", "Axis", 1#, 0#, 10#, 1#, 0#, 80#)
    TestNear "constraint.trim.length", path.Length, 70#
    Set source = New CGeometryPath: source.Initialize "Corner"
    source.AddLine "Bottom", 0#, 0#, 100#, 0#: source.AddLine "Right", 100#, 0#, 100#, 100#: source.Seal
    facade.AddPath source: Set path = facade.Offset("Miter", "Corner", 10#)
    path.CopySegment 1, edge: TestNear "miter.x", edge.X2, 90#: TestNear "miter.y", edge.Y2, 10#
    path.CopySegment 2, edge: TestNear "miter.continuous.x", edge.X1, 90#: TestNear "miter.continuous.y", edge.Y1, 10#
    TestNear "miter.length", path.Length, 180#
    TestNear "miter.source.first.end", path.SourceEndParameter(1), 0.9
    TestNear "miter.source.second.start", path.SourceStartParameter(2), 0.1
    Set source = New CGeometryPath: source.Initialize "Smooth"
    source.AddLine "Straight", 0#, 0#, 100#, 0#
    source.AddArc "Fillet", 100#, 50#, 50#, -GEOM_PI / 2#, GEOM_PI / 2#: source.Seal
    facade.AddPath source: Set path = facade.Offset("SmoothOffset", source.ID, 10#)
    TestNear "offset.analytic.length", path.Length, 100# + 20# * GEOM_PI
    path.CopySegment 2, edge: TestNear "offset.arc.radius", Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2), 40#
    Set path = facade.CirclePath("Circle", 0#, 0#, 100#): Set path = facade.Offset("CircleInner", path.ID, 20#)
    TestNear "offset.circle.length", path.Length, 160# * GEOM_PI: TestCheck "offset.circle.closed", path.IsClosed
    Set source = facade.Transform("MovedFace", "FullFace", 500#, -200#, GEOM_PI / 2#, True)
    facade.Anchor "CornerAnchor", 100#, 0#: facade.TransformAnchor "MovedAnchor", "CornerAnchor", 500#, -200#, GEOM_PI / 2#, True
    Set definition = facade.Finish: definition.AnchorCoordinates "MovedAnchor", x, y
    TestNear "transform.anchor.x", x, 500#: TestNear "transform.anchor.y", y, -100#
    source.AtDistance source.Length, x, y, tx, ty
    TestNear "transform.path.anchor.x", x, 500#: TestNear "transform.path.anchor.y", y, -100#
    Set query = New CSectionGeometryQuery
    Set source = New CGeometryPath: source.Initialize "CCW"
    source.AddLine "Bottom", 0#, 0#, 100#, 0#: source.AddLine "Right", 100#, 0#, 100#, 100#
    source.AddLine "Top", 100#, 100#, 0#, 100#: source.AddLine "Left", 0#, 100#, 0#, 0#: source.Seal True
    Set path = query.OffsetPath("Inset", source, 10#): TestNear "offset.closed.area", query.RegionFromPath(path).Area, 6400#
    Set source = source.Reversed: Set path = query.OffsetPath("OutsetCW", source, 10#)
    TestNear "offset.direction.left", query.RegionFromPath(path).Area, 14400#
    For i = 0 To 4: TestCheck "operations.invalid." & CStr(i), TestInvalidOperation(i): Next i
    GoTo Finished
Failed:
    TestCheck "operations.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRecipeOperationTests = mReport & "TOTAL RECIPE OPERATIONS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет код, связанные объекты и численные подробности отказа смещения или ограничения пути.
Private Function TestInvalidOperation(ByVal scenario As Long) As Boolean
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
    TestInvalidOperation = (Not diagnostic Is Nothing)
    If TestInvalidOperation Then TestInvalidOperation = diagnostic.Code = expected And diagnostic.EntityID = "Failed" And Len(diagnostic.RelatedEntityID) > 0
    If scenario = 2 And TestInvalidOperation Then TestInvalidOperation = diagnostic.Detail1 = 10# And diagnostic.Detail2 = 20#
    Err.Clear
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет наложения, пропущенные проекции, объединённые позиции и ограничение журнала при сохранении полных счётчиков.
Public Function RunRecipeDiagnosticTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, definition As CSectionDefinition, options As CRebarRuleOptions
    Dim rules As CRebarRuleExecutor, layout As CRebarLayout, diagnostic As CGeometryDiagnostic, i As Long, mode As Long
    For mode = 0 To 2
        Set facade = TestDiagnosticFacade: Set options = New CRebarRuleOptions
        If mode = 2 Then options.OverlapPolicy = RejectOverlap
        facade.Source "Bars", "ProgramInputs.Spacing"
        facade.RebarAtStations "Bars", "Row", Array(10#, IIf(mode = 0, 25#, 20#)), 20#, options
        Set definition = facade.Finish: Set rules = New CRebarRuleExecutor
        If mode = 2 Then
            On Error Resume Next
            Set layout = TestFixtureAnnotatedLayout(definition, rules)
            i = Err.Number: Err.Clear
            On Error GoTo Failed
            Set diagnostic = rules.LastDiagnostic
            TestCheck "overlap.reject", i = vbObjectError + 5530 And diagnostic.Code = "BAR_OVERLAP"
            TestCheck "overlap.source", diagnostic.EntityID = "Bars" And diagnostic.SourceReference = "ProgramInputs.Spacing"
            TestNear "overlap.numeric.gap", diagnostic.Detail1, -10#
            TestCheck "overlap.error.actual.created", rules.RequestedPositions = 2 And rules.CreatedPositions = 1 And rules.Conflicts = 1
        Else
            Set layout = TestFixtureAnnotatedLayout(definition, rules)
            TestCheck "overlap.physical." & CStr(mode), layout.Count = 2 And rules.RequestedPositions = 2 And rules.CreatedPositions = 2
            TestCheck "overlap.report." & CStr(mode), rules.BarOverlaps = 1
            TestCheck "overlap.machine." & CStr(mode), rules.DiagnosticAt(1).Code = "BAR_OVERLAP"
            TestCheck "trace.rule." & CStr(mode), InStr(rules.RuleTrace("Bars"), "Created=2") > 0 And InStr(rules.RuleTrace("Bars"), "StationStepMax=") > 0
            TestCheck "trace.owner." & CStr(mode), rules.PositionPhysicalOwner(1) = 1 And rules.PositionPhysicalOwner(2) = 2
            TestCheck "trace.ids." & CStr(mode), rules.PositionRuleID(1) = "Bars" And rules.PositionPathID(1) = "Row" And rules.PositionSourceID(1) = "Bars.1"
        End If
    Next mode
    Set facade = TestDiagnosticFacade
    facade.RebarAtStations "Contact", "Row", Array(10#, 30#), 20#
    Set definition = facade.Finish: Set rules = New CRebarRuleExecutor: Set layout = TestFixtureAnnotatedLayout(definition, rules)
    TestCheck "contact.allowed", rules.BarOverlaps = 0 And layout.Count = 2
    Set facade = TestDiagnosticFacade
    facade.RebarByCount "Base", "Row", 3, 10#
    Set path = facade.LinePath("Target", 20#, 50#, 80#, 50#)
    Set options = New CRebarRuleOptions: options.DirectionY = 1#: options.SkipMissingProjection = True
    facade.Source "Projected", "ProgramInputs.Target"
    facade.RebarProjected "Projected", "Target", "Base", 10#, options
    Set definition = facade.Finish: Set rules = New CRebarRuleExecutor: Set layout = TestFixtureAnnotatedLayout(definition, rules)
    TestCheck "projection.counts", rules.RequestedPositions = 6 And rules.CreatedPositions = 4 And rules.SkippedProjections = 2
    TestCheck "projection.trace", InStr(rules.RuleTrace("Projected"), "Requested=3; Created=1; Merged=0; Skipped=2") > 0
    TestCheck "projection.machine", rules.DiagnosticCount = 2 And rules.DiagnosticAt(1).Code = "PROJECTION_MISSING"
    TestCheck "projection.related", rules.DiagnosticAt(1).RelatedEntityID = "Base.1" And rules.DiagnosticAt(2).RelatedEntityID = "Base.3"
    TestCheck "projection.source", rules.DiagnosticAt(1).SourceReference = "ProgramInputs.Target"
    Set facade = TestDiagnosticFacade: Set path = facade.CirclePath("Ring", 0#, 0#, 40#)
    Set options = New CRebarRuleOptions: options.EndMode = IncludeStart
    facade.RebarAtStations "RingBars", path.ID, Array(0#, 10#, 20#), 4#, options
    Set definition = facade.Finish: Set rules = New CRebarRuleExecutor: Set layout = TestFixtureAnnotatedLayout(definition, rules)
    Dim stepCount As Long, minStep As Double, maxStep As Double, minChord As Double, maxChord As Double
    rules.RuleStepBounds "RingBars", stepCount, minStep, maxStep, minChord, maxChord
    TestCheck "trace.closed.count", stepCount = 3
    TestNear "trace.closed.minimum", minStep, 10#
    TestNear "trace.closed.wrap.maximum", maxStep, 80# * GEOM_PI - 20#
    TestNear "trace.closed.chord", maxChord, 80# * Sin(0.25)
    Set facade = TestDiagnosticFacade
    Dim dense() As Double
    ReDim dense(0 To 99)
    For i = 0 To 99: dense(i) = CDbl(i): Next i
    facade.RebarAtStations "Dense", "Row", dense, 20#
    Set definition = facade.Finish: Set rules = New CRebarRuleExecutor: Set layout = TestFixtureAnnotatedLayout(definition, rules)
    TestCheck "diagnostics.bounded.full.counts", rules.CreatedPositions = 100 And rules.BarOverlaps = 1710
    TestCheck "diagnostics.bounded.records", rules.DiagnosticCount = 128 And rules.DiagnosticRecordsOmitted = 1582
    TestCheck "diagnostics.bounded.latest", rules.LastDiagnostic.Code = "BAR_OVERLAP" And rules.LastDiagnostic.EntityID = "Dense"
    Set facade = TestDiagnosticFacade: Set options = New CRebarRuleOptions: options.Coincidence = MergeIdentical
    facade.RebarAtStations "First", "Row", Array(50#), 10#
    facade.RebarAtStations "Second", "Row", Array(50#), 10#, options
    Set definition = facade.Finish: Set rules = New CRebarRuleExecutor: Set layout = TestFixtureAnnotatedLayout(definition, rules)
    TestCheck "merge.counts", rules.RequestedPositions = 2 And rules.CreatedPositions = 1 And rules.MergedPositions = 1
    TestCheck "merge.physical.owner", rules.PositionPhysicalOwner(1) = 1 And rules.PositionPhysicalOwner(2) = 1
    TestCheck "merge.trace", InStr(rules.RuleTrace("Second"), "Created=0; Merged=1") > 0
    Set facade = TestDiagnosticFacade: Set options = New CRebarRuleOptions: options.DistributionScope = 2 ' Неподдерживаемый способ распределения вызывает ошибку без произвольной подстановки.
    On Error Resume Next
    facade.RebarByCount "Unsupported", "Row", 3, 10#, options
    i = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "scope.invalid", i = vbObjectError + 5530
    On Error Resume Next
    Set layout = rules.Build(Nothing, Nothing)
    i = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck "diagnostics.failed.reuse", i = vbObjectError + 5530 And rules.LastDiagnostic.Code = "RULE_INVALID"
    TestCheck "diagnostics.failed.reuse.counts", rules.RequestedPositions = 0 And rules.CreatedPositions = 0 And rules.MergedPositions = 0 And rules.SkippedProjections = 0
    GoTo Finished
Failed:
    TestCheck "diagnostics.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRecipeDiagnosticTests = mReport & "TOTAL RECIPE DIAGNOSTICS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт общую бетонную область и прямой ряд для изолированной проверки размещения и диагностики.
Private Function TestDiagnosticFacade() As CSectionDefinitionBuilder
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Diagnostics"
    Set path = facade.RectanglePath("Concrete", -100#, -100#, 300#, 300#): facade.Outer "Outer", path.ID
    Set path = facade.LinePath("Row", 0#, 0#, 100#, 0#)
    Set TestDiagnosticFacade = facade
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' поиск по координатам не изменяет линии; выделение участка через Chain сохраняет
' свои правила.
Public Function RunResolveAnchorTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, polygon As CPolygonSpec
    Dim rightID As String, t As Double, right As CGeometryPath, x As Double, y As Double, tx As Double, ty As Double
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "ResolveAnchor"
    Set path = facade.RectanglePath("Rect", 0#, 0#, 400#, 600#)
    path.ResolveAnchor 400#, 300#, rightID, t
    TestNear "rectangle.right.parameter", t, 0.5
    Set right = facade.Chain("Right", path.ID, rightID, 1#, rightID, 0#, False)
    TestNear "rectangle.right.length", right.Length, 600#
    right.AtDistance 0#, x, y, tx, ty: TestNear "rectangle.right.start.x", x, 400#: TestNear "rectangle.right.start.y", y, 600#
    right.AtDistance right.Length, x, y, tx, ty: TestNear "rectangle.right.end.y", y, 0#: TestCheck "rectangle.right.direction", ty < 0#
    TestCheckResolvedPath facade, path
    Set path = facade.AdjacentRectangles("Adjacent", Array(Array(0#, 0#, 400#, 200#), Array(0#, 200#, 200#, 400#)))
    TestCheckResolvedPath facade, path
    Set polygon = New CPolygonSpec
    polygon.AddVertex "BL", 0#, 0#, 30#: polygon.AddVertex "BR", 400#, 0#, 40#
    polygon.AddVertex "TR", 400#, 600#, 50#: polygon.AddVertex "TL", 0#, 600#, 60#: polygon.Seal
    Set path = facade.RoundedPolygon("Rounded", polygon): TestCheckResolvedPath facade, path
    Set path = facade.OffsetRoundedPolygon("RoundedOffset", polygon, 10#): TestCheckResolvedPath facade, path
    Set path = facade.CirclePath("Circle", 25#, -15#, 100#): TestCheckResolvedPath facade, path
    Set path = facade.ArcPath("ArcCCW", 15#, 30#, 80#, 0.9 * GEOM_PI, 0.8 * GEOM_PI): TestCheckResolvedPath facade, path
    TestExpectAnchorError path, 15# + 80#, 30#, "ANCHOR_NOT_ON_PATH"
    Set path = facade.ArcPath("ArcCW", -10#, 40#, 80#, -0.9 * GEOM_PI, -0.8 * GEOM_PI): TestCheckResolvedPath facade, path
    Set path = facade.LinePath("Line", 0#, 0#, 100#, 0#): TestCheckResolvedPath facade, path
    Set path = facade.Transform("MovedRect", "Rect", 1000#, -700#, GEOM_PI / 3#, True): TestCheckResolvedPath facade, path ' Преобразования проверяются и для прямых, и для скруглений; отражение меняет знак дуги.
    Set path = facade.Transform("MovedRounded", "Rounded", -900#, 700#, GEOM_PI / 3#, True): TestCheckResolvedPath facade, path
    Set path = facade.Transform("MovedCircle", "Circle", 900#, -1700#, -GEOM_PI / 4#, True): TestCheckResolvedPath facade, path
    Set path = facade.Transform("MovedArc", "ArcCW", -900#, -700#, GEOM_PI / 4#): TestCheckResolvedPath facade, path
    Set path = facade.LinePath("Tolerance", 0#, 0#, 100#, 0#) ' Реальный общий допуск границы допускает малую поперечную погрешность, но не ближайшую произвольную точку.
    path.ResolveAnchor 50#, GEOM_BOUNDARY_TOLERANCE / 2#, rightID, t: TestNear "tolerance.accepted", t, 0.5
    TestExpectAnchorError path, 50#, GEOM_BOUNDARY_TOLERANCE * 2#, "ANCHOR_NOT_ON_PATH"
    Set path = New CGeometryPath: path.Initialize "Retraced"
    path.AddLine "Out", 0#, 0#, 100#, 0#: path.AddLine "Back", 100#, 0#, 0#, 0#: path.Seal
    TestExpectAnchorError path, 50#, 0#, "ANCHOR_AMBIGUOUS"
    Set path = New CGeometryPath: path.Initialize "Unsealed": path.AddLine "Only", 0#, 0#, 100#, 0#
    TestExpectAnchorError path, 50#, 0#, "PATH_INVALID"
    TestPositionOrderOwner
    GoTo Finished
Failed:
    TestCheck "anchor.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunResolveAnchorTests = mReport & "TOTAL RESOLVE ANCHOR: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет сегменты, оба направления обхода Chain, длины, концы и соответствие участкам
' исходных линий.
Private Sub TestCheckResolvedPath(ByVal facade As CSectionDefinitionBuilder, ByVal path As CGeometryPath)
    Dim i As Long, edge As TRegionEdge, segment As String, parameter As Double, x As Double, y As Double
    Dim chain As CGeometryPath, station As Double, tx As Double, ty As Double, a As Double, b As Double
    Dim expectedX As Double, expectedY As Double, label As String, reverse As Long
    station = 0#
    For i = 1 To path.SegmentCount
        label = path.ID & "." & CStr(i): path.CopySegment i, edge
        modGeometryPrimitives.EdgePoint edge, 0.25, x, y
        path.ResolveAnchor x, y, segment, parameter
        TestCheck label & ".id", segment = path.SegmentID(i): TestNear label & ".parameter", parameter, 0.25
        For reverse = 0 To 1
            a = 0.25: b = 0.75: If reverse = 1 Then a = 0.75: b = 0.25
            Set chain = facade.Chain(label & ".Chain" & CStr(reverse), path.ID, segment, a, segment, b, reverse = 0)
            TestCheck label & ".chain.id", chain.SegmentCount = 1 And chain.SegmentID(1) = segment
            TestCheck label & ".source.id", chain.SourcePathID(1) = path.SourcePathID(i) And chain.SourceSegmentID(1) = path.SourceSegmentID(i)
            TestNear label & ".source.start", chain.SourceStartParameter(1), path.SourceStartParameter(i) + a * (path.SourceEndParameter(i) - path.SourceStartParameter(i))
            TestNear label & ".source.end", chain.SourceEndParameter(1), path.SourceStartParameter(i) + b * (path.SourceEndParameter(i) - path.SourceStartParameter(i))
            TestNear label & ".length", chain.Length, modGeometryPrimitives.EdgeLength(edge) / 2#
            modGeometryPrimitives.EdgePoint edge, a, expectedX, expectedY: chain.AtDistance 0#, x, y, tx, ty
            TestNear label & ".start.x", x, expectedX: TestNear label & ".start.y", y, expectedY
            modGeometryPrimitives.EdgePoint edge, b, expectedX, expectedY: chain.AtDistance chain.Length, x, y, tx, ty
            TestNear label & ".end.x", x, expectedX: TestNear label & ".end.y", y, expectedY
            modGeometryPrimitives.EdgeTangent edge, b, expectedX, expectedY
            If reverse = 1 Then expectedX = -expectedX: expectedY = -expectedY
            TestCheck label & ".direction", tx * expectedX + ty * expectedY > 0#
        Next reverse
        If i > 1 Or path.IsClosed And path.SegmentCount > 1 Then
            TestExpectAnchorError path, edge.X1, edge.Y1, "ANCHOR_AMBIGUOUS"
        Else
            path.ResolveAnchor edge.X1, edge.Y1, segment, parameter: TestNear label & ".start.parameter", parameter, 0#
        End If
        If i = path.SegmentCount And Not path.IsClosed Then
            path.ResolveAnchor edge.X2, edge.Y2, segment, parameter: TestNear label & ".end.parameter", parameter, 1#
        End If
        station = station + modGeometryPrimitives.EdgeLength(edge)
    Next i
    TestNear path.ID & ".unchanged.length", path.Length, station
    TestExpectAnchorError path, 100000#, -100000#, "ANCHOR_NOT_ON_PATH"
    If path.SegmentCount > 1 Then
        Dim first As TRegionEdge, last As TRegionEdge, startID As String, endID As String, startT As Double, endT As Double
        path.CopySegment 1, first: path.CopySegment 2, last
        modGeometryPrimitives.EdgePoint first, 0.75, x, y: path.ResolveAnchor x, y, startID, startT
        modGeometryPrimitives.EdgePoint last, 0.25, x, y: path.ResolveAnchor x, y, endID, endT
        Set chain = facade.Chain(path.ID & ".JointChain", path.ID, startID, startT, endID, endT)
        TestCheck path.ID & ".joint.chain.ids", chain.SegmentCount = 2 And chain.SegmentID(1) = startID And chain.SegmentID(2) = endID
        TestNear path.ID & ".joint.chain.length", chain.Length, (modGeometryPrimitives.EdgeLength(first) + modGeometryPrimitives.EdgeLength(last)) / 4#
        TestCheck path.ID & ".joint.chain.origin", chain.SourceSegmentID(1) = path.SourceSegmentID(1) And chain.SourceSegmentID(2) = path.SourceSegmentID(2)
        modGeometryPrimitives.EdgePoint first, 0.75, expectedX, expectedY: chain.AtDistance 0#, x, y, tx, ty
        TestNear path.ID & ".joint.start.x", x, expectedX: TestNear path.ID & ".joint.start.y", y, expectedY
        modGeometryPrimitives.EdgePoint last, 0.25, expectedX, expectedY: chain.AtDistance chain.Length, x, y, tx, ty
        TestNear path.ID & ".joint.end.x", x, expectedX: TestNear path.ID & ".joint.end.y", y, expectedY
        If path.IsClosed Then
            path.CopySegment path.SegmentCount, last
            modGeometryPrimitives.EdgePoint last, 0.75, x, y: path.ResolveAnchor x, y, startID, startT
            modGeometryPrimitives.EdgePoint first, 0.25, x, y: path.ResolveAnchor x, y, endID, endT
            Set chain = facade.Chain(path.ID & ".WrapChain", path.ID, startID, startT, endID, endT)
            TestCheck path.ID & ".wrap.ids", chain.SegmentCount = 2 And chain.SegmentID(1) = startID And chain.SegmentID(2) = endID
            TestNear path.ID & ".wrap.length", chain.Length, (modGeometryPrimitives.EdgeLength(first) + modGeometryPrimitives.EdgeLength(last)) / 4#
            TestCheck path.ID & ".wrap.origin", chain.SourceSegmentID(1) = path.SourceSegmentID(path.SegmentCount) And chain.SourceSegmentID(2) = path.SourceSegmentID(1)
            modGeometryPrimitives.EdgePoint last, 0.75, expectedX, expectedY: chain.AtDistance 0#, x, y, tx, ty
            TestNear path.ID & ".wrap.start.x", x, expectedX: TestNear path.ID & ".wrap.start.y", y, expectedY
            modGeometryPrimitives.EdgePoint first, 0.25, expectedX, expectedY: chain.AtDistance chain.Length, x, y, tx, ty
            TestNear path.ID & ".wrap.end.x", x, expectedX: TestNear path.ID & ".wrap.end.y", y, expectedY
        End If
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ошибка оставляет переданные выходные значения неизменными и содержит координаты точки.
Private Sub TestExpectAnchorError(ByVal path As CGeometryPath, ByVal x As Double, ByVal y As Double, ByVal code As String)
    Dim segment As String, parameter As Double, number As Long, reason As String, diagnostic As CGeometryDiagnostic
    segment = "unchanged": parameter = -1#
    On Error Resume Next
    path.ResolveAnchor x, y, segment, parameter: number = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo 0
    Set diagnostic = path.LastDiagnostic
    TestCheck path.ID & ".error.number", number = vbObjectError + 5520 And InStr(reason, path.ID) > 0
    TestCheck path.ID & ".error.outputs", segment = "unchanged" And parameter = -1#
    TestCheck path.ID & ".error.diagnostic", Not diagnostic Is Nothing
    If Not diagnostic Is Nothing Then
        TestCheck path.ID & ".error.code", diagnostic.Code = code And diagnostic.EntityID = path.ID
        If code <> "PATH_INVALID" Then
            TestNear path.ID & ".error.x", diagnostic.Detail1, x: TestNear path.ID & ".error.y", diagnostic.Detail2, y
        End If
        If code = "ANCHOR_AMBIGUOUS" Then TestCheck path.ID & ".error.segments", InStr(diagnostic.RelatedEntityID, ", ") > 0
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
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
        Set definition = facade.Finish: Set rules = New CRebarRuleExecutor: Set layout = TestFixtureAnnotatedLayout(definition, rules)
        owner = layout.FindBarAt(15#, 0#): TestCheck "position.order.owner.exists", owner > 0
        If owner > 0 Then
            If mode = AlongPath Then TestNear "position.order.along.diameter", layout.Diameter(owner), 20# Else TestNear "position.order.middle.diameter", layout.Diameter(owner), 10#
        End If
    Next mode
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Учитывает отдельное условие проверки и сохраняет его имя в итоговом протоколе.
Private Sub TestCheck(ByVal name As String, ByVal value As Boolean)
    If value Then mPassed = mPassed + 1 Else mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает фактическую и ожидаемую величины с заданным в этом наборе допуском, сохраняя числа в протоколе.
Private Sub TestNear(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    TestCheck name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) <= 0.000000001 * (1# + Abs(expected))
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' исчезновение угловой дуги сохраняет точный стык и имена привязок.
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
        TestCheck "collapse.segmentCount." & CStr(distance), result.SegmentCount = 2
        TestCheck "collapse.anchor." & CStr(distance), result.HasCollapsedSegmentID("Corner") And Not result.HasSegmentID("Corner")
        result.AtDistance result.AnchorStation("Corner"), x, y, tx, ty
        TestNear "collapse.joint.x." & CStr(distance), x, 140# - CDbl(distance)
        TestNear "collapse.joint.y." & CStr(distance), y, CDbl(distance)
        TestCheck "collapse.joint.belongs." & CStr(distance), result.ContainsPoint(x, y)
        TestNear "collapse.length." & CStr(distance), result.Length, 280# - 2# * CDbl(distance)
    Next distance
    Set result = query.OffsetPath("BeforeCollapse", source, 39.99)
    TestCheck "collapse.before.arcKept", result.HasSegmentID("Corner") And result.SegmentCount = 3
    Set result = query.OffsetPath("AfterCollapse", source, 50#)
    source.AtDistance 100# + 10# * GEOM_PI, x, y, tx, ty
    TestCheck "collapse.inverted.point.rejected", Not result.ContainsPoint(x - 50# * ty, y + 50# * tx)
    TestCheck "collapse.source.unchanged", source.SegmentCount = 3 And source.HasSegmentID("Corner")
    Set result = query.OffsetPath("ReverseOffset", source.Reversed, -50#)
    TestNear "collapse.reverse.length", result.Length, 180#
    Set source = source.Transformed(500#, -200#, GEOM_PI / 3#, True)
    Set result = query.OffsetPath("TransformedOffset", source, -50#)
    result.AtDistance result.AnchorStation("Corner"), x, y, tx, ty
    tx = 90#: ty = 50#: query.TransformCoordinates tx, ty, 500#, -200#, GEOM_PI / 3#, True
    TestNear "collapse.transform.x", x, tx: TestNear "collapse.transform.y", y, ty
    Set polygon = New CPolygonSpec
    polygon.AddVertex "A", 0#, 0#, 10#: polygon.AddVertex "B", 100#, 0#, 10#
    polygon.AddVertex "C", 100#, 100#, 10#: polygon.AddVertex "D", 0#, 100#, 10#: polygon.Seal
    Set source = query.RoundedPolygon("Closed", polygon)
    Set result = query.OffsetPath("ClosedOffset", source, 20#)
    TestCheck "collapse.closed", result.IsClosed And result.SegmentCount = 4
    TestNear "collapse.closed.area", query.RegionFromPath(result).Area, 3600#
    GoTo Finished
Failed:
    TestCheck "collapse.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunCollapsingOffsetTests = mReport & "TOTAL_COLLAPSING_OFFSETS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохранённый PointedSection при 40 мм сохраняет стержень в стыке,
' а при большем смещении пропускает только позицию, вышедшую за новую линию.
Public Function RunSavedPointedOffsetTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim table As Object, config As Object, saved As Variant, values As Variant, value As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim i As Long, count As Long, expected As Long, inputCell As Object
    Dim attention As Boolean, information As Boolean, message As String
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    TestPrepareSavedOffsetCase
    Set table = ThisWorkbook.Names("rngPointedSectionGeometry").RefersToRange
    Set inputCell = TestFixtureInputCell(table, "Extra", "Right", "a2")
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
        TestCheck "saved.right.created." & CStr(value), count = expected
        TestCheck "saved.right.warning." & CStr(value), TestSavedRuleSkipped(model.GeometryDiagnosticsText, "H.Right.Row2") = IIf(CDbl(value) > 40#, 1, 0)
        TestCheck "saved.warning.summary." & CStr(value), (Len(model.GeometryWarningText) > 0) = (InStr(model.GeometryDiagnosticsText, "Предупреждение:") > 0)
        If CDbl(value) > 40# Then TestCheck "saved.right.warning.address." & CStr(value), InStr(model.GeometryDiagnosticsText, inputCell.Worksheet.Name & "!" & inputCell.Address(False, False)) > 0
        TestCheck "saved.warning.typed." & CStr(value), model.HasGeometryWarnings = (CDbl(value) > 40#)
        model.GeometryDiagnosticsText = "Localized text without warning keyword"
        TestCheck "saved.warning.localized." & CStr(value), model.HasGeometryWarnings = (CDbl(value) > 40#)
        model.GeometryDiagnosticsText = ""
        TestCheck "saved.warning.empty." & CStr(value), model.HasGeometryWarnings = (CDbl(value) > 40#)
        attention = Not (CDbl(value) > 40#)
        message = FormSectionForWorkbook(ThisWorkbook, -1#, information, attention)
        TestCheck "saved.warning.buttonFlag." & CStr(value), attention = (CDbl(value) > 40#)

    Next value
    Set model = New CSectionModel
    model.GeometryDiagnosticsText = "Предупреждение: исторический текст без кодов."
    TestCheck "saved.warning.historicalNotClassified", Not model.HasGeometryWarnings And model.GeometryWarningText = ""
    model.SetGeometryWarnings True, ""
    TestCheck "saved.warning.emptyDisplayStillEvent", model.HasGeometryWarnings
    model.Clear
    TestCheck "saved.warning.clearResetsEvent", Not model.HasGeometryWarnings
    GoTo Finished
Failed:
    TestCheck "saved.offset.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    RunSavedPointedOffsetTests = mReport & "TOTAL_SAVED_POINTED_OFFSETS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' соседи прямой стороны ограничивают дополнительные ряды даже без дуг.
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
                    Set layout = TestFixtureAnnotatedLayout(definition, executor)
                    row2 = 0: row3 = 0
                    For i = 1 To layout.Count
                        x = layout.X(i) - 1700#: y = layout.Y(i) + 900#
                        query.TransformCoordinates x, y, 0#, 0#, -GEOM_PI / 6# ' Обратный поворот и отражение дают координаты до преобразования.
                        If reflect = 1 Then y = -y
                        If Abs(y - 150#) < 0.000001 Then
                            row2 = row2 + 1: TestCheck "whole.row2.bounds", x >= 150# And x <= 850#
                        ElseIf Abs(y - 250#) < 0.000001 Then
                            row3 = row3 + 1: TestCheck "whole.row3.bounds", x >= 250# And x <= 750#
                        End If
                    Next i
                    prefix = "whole." & CStr(reverse) & CStr(reflect) & CStr(binding) & CStr(chain)
                    TestCheck prefix & ".row2", row2 = IIf(binding = 1, 1, 3)
                    TestCheck prefix & ".row3", row3 = IIf(binding = 1 And chain = 1, 1, 3)
                    TestCheck prefix & ".warning", InStr(executor.DiagnosticsText, "Предупреждение:") > 0
                    TestCheck prefix & ".typedWarning", executor.HasGeometryWarnings
                    TestCheck prefix & ".address", InStr(executor.DiagnosticsText, "WholeBoundary[Extra,2]") > 0
                    TestCheck prefix & ".source.unchanged", whole.IsClosed And whole.SegmentCount = 4 And side.SegmentCount = 1
                Next chain
            Next binding
        Next reflect
    Next reverse
    GoTo Finished
Failed:
    TestCheck "whole.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunFullBoundaryOffsetTests = mReport & "TOTAL_FULL_BOUNDARY_OFFSETS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' обычный расчёт сохранённого примера проходит через запись Results и схему.
' Предупреждение о пропуске стержня проверяется в ответе формирования и в Results.
Public Function RunSavedOffsetCalculationTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim area As Object, config As Object, inputCell As Object, saved As Variant, value As Variant, message As String
    Dim summary As Object, i As Long, j As Long, text As String, diagnostics As String, formation As String
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = TestFixtureSaveInputs(config)
    TestPrepareSavedOffsetCase
    TestPrepareSavedOffsetCalculationLoad
    Set area = ThisWorkbook.Names("rngPointedSectionGeometry").RefersToRange
    Set inputCell = TestFixtureInputCell(area, "Extra", "Right", "a2")
    For Each value In Array(40#, 50#)
        inputCell.Value2 = value
        formation = FormSectionForWorkbook(ThisWorkbook)
        message = RunSectionCalculationForWorkbook(ThisWorkbook, True)
        TestCheck "saved.calculation.completed." & CStr(value), InStr(message, "завершен.") > 0 And InStr(message, "ошиб") = 0
        diagnostics = TestSavedRecipeDiagnostics()
        TestCheck "saved.calculation.warning." & CStr(value), (InStr(formation, "Предупреждение:") > 0) = (InStr(diagnostics, "Предупреждение:") > 0)
        TestCheck "saved.calculation.rowWarning." & CStr(value), TestSavedRuleSkipped(diagnostics, "H.Right.Row2") = IIf(CDbl(value) > 40#, 1, 0)
        Set summary = ThisWorkbook.Names("rngBatchSummary").RefersToRange
        text = ""
        For i = 1 To summary.Rows.Count
            For j = 1 To summary.Columns.Count
                text = text & CStr(summary.Cells(i, j).Value2) & "|"
            Next j
        Next i
        TestCheck "saved.calculation.status." & CStr(value), InStr(text, "NumFail") = 0 And InStr(text, "InputErr") = 0 And InStr(text, "CalcErr") = 0
        mReport = mReport & "OBSERVED_SAVED_CALCULATION: a2=" & CStr(value) & "; " & Replace(message, vbCrLf, " | ") & vbCrLf
    Next value
    GoTo Finished
Failed:
    TestCheck "saved.calculation.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then TestFixtureRestoreInputs config, saved
    RunSavedOffsetCalculationTests = mReport & "TOTAL_SAVED_OFFSET_CALCULATION: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' воспроизводит пример R80/as40 без дополнительных рядов на соседних сторонах.
' Их независимые отступы не относятся к проверке исчезновения правого скругления.
Private Sub TestPrepareSavedOffsetCase()
    Dim table As Object, key As Variant
    ThisWorkbook.Names("rngUnitSettings").RefersToRange.Cells(2, 2).Value2 = "mm"
    modTestShapeArchitecture.TestConfigureShapeFixture "RoundedMixed"
    Set table = ThisWorkbook.Names("rngPointedSectionGeometry").RefersToRange
    TestFixtureInputCell(table, "General", "B").Value2 = 1000#: TestFixtureInputCell(table, "General", "H").Value2 = 500#
    TestFixtureInputCell(table, "General", "Left.W").Value2 = 0#: TestFixtureInputCell(table, "General", "Left.R1").Value2 = 0#
    TestFixtureInputCell(table, "General", "Left.R2").Value2 = 0#: TestFixtureInputCell(table, "General", "Right.W").Value2 = 150#
    TestFixtureInputCell(table, "General", "Right.R1").Value2 = 0#: TestFixtureInputCell(table, "General", "Right.R2").Value2 = 80#
    For Each key In Array("Top", "Bottom", "Left", "Right")
        TestFixtureInputCell(table, "Main", CStr(key), "as").Value2 = 40#
        TestFixtureInputCell(table, "Main", CStr(key), "d").Value2 = 25#
        TestFixtureInputCell(table, "Main", CStr(key), "n").Value2 = IIf(CStr(key) = "Top" Or CStr(key) = "Bottom", 10, 5)
        TestFixtureInputCell(table, "Main", CStr(key), "s0").Value2 = 0#: TestFixtureInputCell(table, "Main", CStr(key), "s1").Value2 = 0#
        TestFixtureInputCell(table, "Extra", CStr(key), "d2").Value2 = IIf(CStr(key) = "Right", 16#, 0#)
        TestFixtureInputCell(table, "Extra", CStr(key), "d3").Value2 = 0#
    Next key
    TestFixtureInputCell(table, "Extra", "Right", "pos2").Value2 = "Stacked"
    TestFixtureInputCell(table, "Extra", "Right", "bind2").Value2 = "EachBar"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' читает полный счётчик пропусков конкретного правила.
' Предупреждение другого ряда не означает, что проверяемый ряд потерял стержень.
Private Function TestSavedRuleSkipped(ByVal diagnostics As String, ByVal ruleID As String) As Long
    Dim line As Variant, prefix As String, marker As String, position As Long
    prefix = "Rule=" & ruleID & ";": marker = "; Skipped="
    For Each line In Split(diagnostics, vbCrLf)
        If Left$(CStr(line), Len(prefix)) = prefix Then
            position = InStr(CStr(line), marker)
            If position > 0 Then
                TestSavedRuleSkipped = CLng(Split(Mid$(CStr(line), position + Len(marker)), ";")(0))
                Exit Function
            End If
        End If
    Next line
    Err.Raise 5, "modTestRecipeOperations", "В диагностике отсутствует счётчик пропусков ряда " & ruleID
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' берёт диагностику из записанного Results после штатного расчёта.
Private Function TestSavedRecipeDiagnostics() As String
    Dim anchor As Object, row As Long
    Set anchor = ThisWorkbook.Names("rngNDMSectionProperties").RefersToRange
    row = 2
    Do While Len(CStr(anchor.Cells(row, 1).Value2)) > 0
        If CStr(anchor.Cells(row, 3).Value2) = "Geometry.RecipeDiagnostics" Then
            TestSavedRecipeDiagnostics = CStr(anchor.Cells(row, 4).Value2): Exit Function
        End If
        row = row + 1
    Loop
    Err.Raise 5, "modTestRecipeOperations", "В Results отсутствует сохранённая диагностика геометрии."
End Function


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задаёт собственное двухосное сочетание независимо от пользовательских строк чистой книги.
' Все усилия сначала заданы во внутренних единицах, затем переводятся единой CUnitSystem.
Private Sub TestPrepareSavedOffsetCalculationLoad()
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Dim units As CUnitSystem: Set units = New CUnitSystem: units.LoadFromSettings settings
    Dim loads As Object: Set loads = ThisWorkbook.Names("rngLoadCombinations").RefersToRange
    Dim values() As Variant: ReDim values(1 To loads.Rows.Count - 1, 1 To loads.Columns.Count)
    values(1, 1) = "LC_SAVED_OFFSETS"
    values(1, 2) = -111000# / units.InputForceToInternal(1#)
    values(1, 3) = 6000000# / units.InputMomentMxToInternal(1#)
    values(1, 4) = 3000000# / units.InputMomentMyToInternal(1#)
    values(1, 5) = "PR1": values(1, 6) = "Auto"
    values(1, 7) = "ДЛЯ ТЕСТОВ: сохранённые смещения, собственное сочетание"
    loads.Offset(1, 0).Resize(UBound(values, 1), UBound(values, 2)).Value2 = values
End Sub
