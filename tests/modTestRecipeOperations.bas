Attribute VB_Name = "modTestRecipeOperations"
Option Explicit

' ДЛЯ ТЕСТОВ: общие offsets/constraints и проверяемая диагностика раскладки.
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
    ' Явный неподдерживаемый scope не получает молчаливое распределение.
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

Private Sub Check(ByVal name As String, ByVal value As Boolean)
    If value Then mPassed = mPassed + 1 Else mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
End Sub

Private Sub Near(ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    Check name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected), Abs(actual - expected) <= 0.000000001 * (1# + Abs(expected))
End Sub
