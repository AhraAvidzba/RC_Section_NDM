Attribute VB_Name = "modTestAudit02Api"
Option Explicit

' ДЛЯ ТЕСТОВ: новые контракты описателя, options, технических имён и прежних измерений.
Private mPassed As Long, mFailed As Long, mReport As String

' ДЛЯ ТЕСТОВ: повторное использование spec, атомарная регистрация и политика имён без зависимости от типа.
Public Function RunAudit02PolygonApiTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim polygon As CPolygonSpec, crossed As CPolygonSpec, query As CSectionGeometryQuery
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, inside As CGeometryPath, outside As CGeometryPath
    Dim definition As CSectionDefinition, contours As CSectionContours, before As String, errorCode As Long, i As Long
    Dim region As CConcreteRegion, n As Variant, angle As Double, x As Double, y As Double
    Set query = New CSectionGeometryQuery
    For Each n In Array(4, 5, 6, 9)
        Set polygon = New CPolygonSpec
        For i = 1 To CLng(n)
            angle = 2# * GEOM_PI * CDbl(i - 1) / CDbl(n)
            polygon.AddVertex "V" & CStr(i), 100# * Cos(angle), 100# * Sin(angle), i - 1#, "Vertex." & CStr(i)
        Next i
        polygon.Seal
        Set path = query.RoundedPolygon("Many." & CStr(n), polygon)
        Check "P01.kernel." & CStr(n), path.IsClosed And path.IsSealed And path.SegmentCount = 2 * CLng(n) - 1
    Next n
    Set polygon = RectangleSpec(20#): before = SpecSnapshot(polygon)
    Set path = query.RoundedPolygon("Base", polygon)
    Set inside = query.OffsetRoundedPolygon("Inward", polygon, 5#)
    Set outside = query.OffsetRoundedPolygon("Outward", polygon, -5#)
    Check "P05.spec.unchanged", SpecSnapshot(polygon) = before
    Check "P05.paths.independent", Not path Is inside And Not inside Is outside
    Set region = query.RegionFromPath(path)
    Check "P02.radius.geometry", Abs(region.Area - (10000# - 4# * 400# + GEOM_PI * 400#)) < 0.000001
    Set crossed = New CPolygonSpec
    crossed.AddVertex "A", 0#, 0#: crossed.AddVertex "B", 100#, 100#
    crossed.AddVertex "C", 0#, 100#: crossed.AddVertex "D", 100#, 0#: crossed.Seal
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Atomic"
    On Error Resume Next
    Set path = facade.RoundedPolygon("Boundary", crossed)
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "P06.invalid.topology.rejected", errorCode <> 0
    Set path = facade.RoundedPolygon("Boundary", polygon)
    On Error Resume Next
    facade.AddPath path
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "P07.creator.registers.once", errorCode <> 0
    facade.Outer "Outer", path.ID, namingPolicy:=SequentialByKind, namingPrefix:="OUT_A"
    Set path = facade.RectanglePath("Other", 200#, 0#, 100#, 100#)
    facade.Outer "OtherOuter", path.ID, namingPolicy:=SequentialByKind, namingPrefix:="OUT_B"
    Set path = New CGeometryPath: path.Initialize "External"
    path.AddLine "External.Edge", 1#, 1#, 20#, 20#: path.Seal: facade.AddPath path
    Set definition = facade.Finish: Set contours = definition.MaterialContours
    Check "P06.no.partial.path", definition.PathCount = 3
    Check "P08.external.path", definition.Path("External").Length = Sqr(19# * 19# * 2#)
    Check "G10.multiple.prefixes", contours.ContourID(1) = "OUT_A_ARC_1" And contours.ContourID(9) = "OUT_B_LINE_1"
    Check "G10.loop.roles", contours.LoopID(1) = "Outer" And contours.LoopID(9) = "OtherOuter" And contours.LoopRole(9) = "Outer"
    On Error Resume Next
    definition.AnchorCoordinates "A", x, y
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "A03.no.implicit.anchor", errorCode <> 0
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "NamingConflict"
    Set path = facade.RectanglePath("Boundary", 0#, 0#, 100#, 100#)
    On Error Resume Next
    facade.Outer "Outer", path.ID, contourIDs:=Array("A", "B", "C", "D"), namingPolicy:=SequentialByKind
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "G10.explicit.policy.conflict", errorCode <> 0
    facade.Outer "Outer", path.ID, namingPolicy:=SequentialByKind
    Set path = facade.RectanglePath("Other", 200#, 0#, 100#, 100#)
    On Error Resume Next
    facade.Outer "Other", path.ID, namingPolicy:=SequentialByKind
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "G10.duplicate.prefix.rejected", errorCode <> 0
    GoTo Finished
Failed:
    Check "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunAudit02PolygonApiTests = mReport & "TOTAL AUDIT02 POLYGON API: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: контактный helper меняет только компоненты смещения и EverySecond.
Public Function RunAudit02OptionsTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim options As CRebarRuleOptions, errorCode As Long, axis As Variant, sign As Variant
    Set options = New CRebarRuleOptions
    options.StartSetback = 100#: options.EndSetback = 150#: options.EndMode = ExcludeBoth
    options.PositionOrder = FromPathMiddle: options.IncomingTangentAtJoint = True
    options.SourceNormalSign = -1#: options.SuccessfulProjectionOrdinal = True
    options.SkipMissingProjection = True: options.NewPublicationRoot = True
    options.SetPublication "Group", PublishPerSourcePosition, KeepPublishedOwner
    Check "R08.publication.only", options.StartSetback = 100# And options.EndMode = ExcludeBoth And options.SourceNormalSign = -1#
    For Each axis In Array(RowNormal, RowTangent)
        For Each sign In Array(-1#, 1#)
            options.ConfigureDependentContact 20#, 16#, CLng(axis), CDbl(sign), False
            Check "R03.contact.18." & CStr(axis) & "." & CStr(sign), _
                IIf(axis = RowNormal, options.NormalOffset, options.TangentOffset) = sign * 18# And _
                IIf(axis = RowNormal, options.TangentOffset, options.NormalOffset) = 0#
            options.ConfigureDependentContact 20#, 12#, CLng(axis), CDbl(sign), True, 16#
            Check "R04.contact.32." & CStr(axis) & "." & CStr(sign), IIf(axis = RowNormal, options.NormalOffset, options.TangentOffset) = sign * 32#
            options.ConfigureDependentContact 20#, 12#, CLng(axis), CDbl(sign), False
            Check "R04.other.direction.16." & CStr(axis) & "." & CStr(sign), IIf(axis = RowNormal, options.NormalOffset, options.TangentOffset) = sign * 16#
        Next sign
    Next axis
    Check "R08.independent.fields", options.StartSetback = 100# And options.EndSetback = 150# And options.EndMode = ExcludeBoth And _
        options.PositionOrder = FromPathMiddle And options.IncomingTangentAtJoint And options.SuccessfulProjectionOrdinal And _
        options.SkipMissingProjection And options.NewPublicationRoot And options.SourceNormalSign = -1# And _
        options.AnnotationGroup = "Group" And options.Publication = PublishPerSourcePosition And options.Coincidence = KeepPublishedOwner
    On Error Resume Next
    options.ConfigureDependentContact 20#, 12#, RowNormal, 0#, False
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "R08.reject.invalid.sign", errorCode <> 0 And options.TangentOffset = 16# And options.NormalOffset = 0#
    GoTo Finished
Failed:
    Check "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunAudit02OptionsTests = mReport & "TOTAL AUDIT02 OPTIONS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: положительные/неотрицательные длины сохраняют обязательность ввода и его источник.
Public Function RunAudit02InputTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim values(1 To 4, 1 To 1) As Variant, addresses(1 To 4, 1 To 1) As Variant
    Dim table As CSectionInputTable, units As CUnitSystem, number As Double, errorCode As Long, reason As String, i As Long
    values(1, 1) = 25#: values(2, 1) = 0#: values(3, 1) = -1#: values(4, 1) = "-"
    For i = 1 To 4: addresses(i, 1) = "Config!A" & CStr(i): Next i
    Set table = New CSectionInputTable: table.Initialize "Input", values, addresses
    Set units = New CUnitSystem: units.InitializeDefaults
    Check "input.positive", table.PositiveLength(1, 1, units) = 25#
    Check "input.nonnegative.zero", table.NonNegativeLength(2, 1, units) = 0#
    For i = 2 To 4
        On Error Resume Next
        number = table.PositiveLength(i, 1, units)
        errorCode = Err.Number: reason = Err.Description: Err.Clear
        On Error GoTo Failed
        Check "input.reject.address." & CStr(i), errorCode <> 0 And InStr(reason, "Config!A" & CStr(i)) > 0
    Next i
    GoTo Finished
Failed:
    Check "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunAudit02InputTests = mReport & "TOTAL AUDIT02 INPUT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: три линейных постановки, действительная грань и смешанная длина пути.
Public Function RunAudit02AnnotationTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim facade As CSectionDefinitionBuilder, definition As CSectionDefinition, path As CGeometryPath, polygon As CPolygonSpec
    Dim builder As CSectionAnnotationBuilder, annotations As CSectionAnnotations, rebars As CRebarLayout, edge As TRegionEdge
    Dim x As Double, y As Double, errorCode As Long
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Dimensions"
    Set path = facade.RectanglePath("Boundary", -100#, -100#, 1000#, 1000#): facade.Outer "Outer", path.ID
    facade.Anchor "A", 0#, 0#: facade.Anchor "B", 300#, 400#
    facade.Dimension "H", "A", "B", "X", "Horizontal"
    facade.Dimension "V", "A", "B", "Y", "Vertical"
    facade.Dimension "L", "A", "B", "L", "Aligned"
    Set polygon = RectangleSpec(20#): Set path = facade.RoundedPolygon("Filleted", polygon)
    facade.PathAnchor "Touch", path.ID, "Edge.A", 0#
    Set path = New CGeometryPath: path.Initialize "Mixed"
    path.AddLine "Line", 0#, 0#, 100#, 0#: path.AddArc "Arc", 100#, 50#, 50#, -GEOM_PI / 2#, GEOM_PI / 2#: path.Seal
    facade.AddPath path: facade.PathLength "PL", path.ID, "A", "PL"
    Set definition = facade.Finish
    Set builder = New CSectionAnnotationBuilder: Set annotations = New CSectionAnnotations: Set rebars = New CRebarLayout
    builder.Build definition, rebars, annotations
    Check "A01.horizontal.300", annotations.Value(1) = 300#
    Check "A01.vertical.400", annotations.Value(2) = 400#
    Check "A01.aligned.500", annotations.Value(3) = 500#
    Check "A02.aligned.endpoints", annotations.StartX(3) = 0# And annotations.EndX(3) = 300# And annotations.EndY(3) = 400#
    Check "A02.aligned.normal", annotations.NormalX(3) = -0.8 And annotations.NormalY(3) = 0.6
    definition.AnchorCoordinates "Touch", x, y
    Check "A03.virtual.not.touch", x <> polygon.X(1) Or y <> polygon.Y(1)
    Check "A04.mixed.length", Abs(annotations.Value(4) - (100# + 25# * GEOM_PI)) < 0.000001
    Check "A04.not.chord", Abs(annotations.Value(4) - Sqr(150# * 150# + 50# * 50#)) > 10#
    builder.Build definition, rebars, annotations
    Check "A05.repeat", annotations.Count = 4 And rebars.Count = 0
    GoTo Finished
Failed:
    Check "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunAudit02AnnotationTests = mReport & "TOTAL AUDIT02 ANNOTATIONS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' Исходный квадрат в прежнем CCW-порядке; имена не являются автоматическими anchors.
Private Function RectangleSpec(ByVal radius As Double) As CPolygonSpec
    Dim polygon As CPolygonSpec: Set polygon = New CPolygonSpec
    polygon.AddVertex "A", 0#, 0#, radius, "Input.A"
    polygon.AddVertex "B", 100#, 0#, radius, "Input.B"
    polygon.AddVertex "C", 100#, 100#, radius, "Input.C"
    polygon.AddVertex "D", 0#, 100#, radius, "Input.D"
    polygon.Seal: Set RectangleSpec = polygon
End Function
Private Function SpecSnapshot(ByVal polygon As CPolygonSpec) As String
    Dim i As Long
    For i = 1 To polygon.Count
        SpecSnapshot = SpecSnapshot & polygon.VertexID(i) & "|" & CStr(polygon.X(i)) & "|" & CStr(polygon.Y(i)) & _
            "|" & CStr(polygon.FilletRadius(i)) & "|" & polygon.SourceReference(i) & vbCrLf
    Next i
End Function
Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        mPassed = mPassed + 1: mReport = mReport & "PASS: " & name & vbCrLf
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
    End If
End Sub
