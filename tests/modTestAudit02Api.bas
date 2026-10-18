Attribute VB_Name = "modTestAudit02Api"
Option Explicit

' ДЛЯ ТЕСТОВ: описание сечения, параметры арматуры, технические имена и совместимость размеров.
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

' ДЛЯ ТЕСТОВ: независимые отступы, контактные группы, знаки и добавка 2,5 мм на каждое соседство.
Public Function RunIndependentRowOffsetBuilderTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim cases As Variant, item As Variant, firstLocation As Variant, lastLocation As Variant
    Dim contact As CRebarContactBuilder, defaults As CRebarRuleOptions
    Dim second As CRebarRuleOptions, third As CRebarRuleOptions, copy As CRebarRuleOptions
    Dim expectedNormal As Double, expectedTangent As Double, groupBase As Double, intermediate As Double, tag As String
    cases = Array(Array(0#, 0#), Array(100#, 200#), Array(200#, 100#), Array(100#, 100#), Array(0#, 100#), Array(100#, 0#))
    For Each item In cases
        For Each firstLocation In Array("Stacked", "SideBySide")
            For Each lastLocation In Array("Stacked", "SideBySide")
                Set defaults = New CRebarRuleOptions
                Set contact = New CRebarContactBuilder: contact.Initialize 32#, -1#, 1#, defaults
                Set second = contact.AppendRow(20#, CStr(firstLocation), "EverySecondBar", "Offsets[2]", CDbl(item(0)))
                Set third = contact.AppendRow(16#, CStr(lastLocation), "EachBar", "Offsets[3]", CDbl(item(1)))
                tag = CStr(item(0)) & "/" & CStr(item(1)) & "/" & CStr(firstLocation) & "/" & CStr(lastLocation)
                expectedNormal = CDbl(item(1)): expectedTangent = 0#: intermediate = 0#: groupBase = 0#
                If CDbl(item(1)) = 0# Then groupBase = 32#
                If CDbl(item(1)) = CDbl(item(0)) Then
                    If CDbl(item(1)) > 0# Then
                        groupBase = 20#
                    ElseIf CStr(firstLocation) = CStr(lastLocation) Then
                        intermediate = 22.5
                    End If
                End If
                If groupBase > 0# Then
                    If CStr(lastLocation) = "Stacked" Then
                        expectedNormal = expectedNormal + groupBase / 2# + intermediate + 8# + 2.5
                    Else
                        expectedTangent = groupBase / 2# + intermediate + 8# + 2.5
                    End If
                End If
                Check "offsets.normal." & tag, third.NormalOffset = -expectedNormal
                Check "offsets.tangent." & tag, third.TangentOffset = expectedTangent
                Check "offsets.binding." & tag, second.EverySecond And Not third.EverySecond
                Set copy = third.Clone
                Check "offsets.clone." & tag, copy.ValidateRowOffset = third.ValidateRowOffset And copy.NormalOffset = third.NormalOffset
            Next lastLocation
        Next firstLocation
    Next item
    Set contact = New CRebarContactBuilder: contact.Initialize 32#, 1#, -1#, defaults
    Set second = contact.AppendRow(20#, "Stacked", "EachBar", "Tolerance[2]", 100#)
    Set third = contact.AppendRow(16#, "Stacked", "EachBar", "Tolerance[3]", 100# + GEOM_TOLERANCE / 2#)
    Check "offsets.tolerance.same.group", Abs(third.NormalOffset - 120.5) < GEOM_TOLERANCE
    Set contact = New CRebarContactBuilder: contact.Initialize 32#, 1#, -1#, defaults
    Set third = contact.AppendRow(16#, "SideBySide", "EverySecondBar", "DisabledSecond[3]", 100#)
    Check "offsets.disabled.second", third.NormalOffset = 100# And third.TangentOffset = 0# And third.EverySecond
    Set contact = New CRebarContactBuilder: contact.Initialize 32#, 1#, -1#, defaults
    Set third = contact.AppendRow(16#, "SideBySide", "EachBar", "DisabledSecondZero[3]")
    Check "offsets.profile.disabled.second", third.NormalOffset = 0# And third.TangentOffset = -26.5
    GoTo Finished
Failed:
    Check "offsets.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunIndependentRowOffsetBuilderTests = mReport & "TOTAL INDEPENDENT ROW OFFSET BUILDER: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
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
    options.AnnotationGroup = "Group" ' В одной подписи группы суммируем стержни каждого диаметра, включая все её ряды.
    options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
    options.Coincidence = KeepPublishedOwner ' При совпадении центров сохраняем стержень, который был добавлен раньше.
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

' ДЛЯ ТЕСТОВ: полные примеры для комментарийного гайда и реальная наклонная грань.
Public Function RunAudit02GuideExamplesTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim definition As CSectionDefinition, builder As CSectionBuilder, model As CSectionModel
    Dim mesh As TSectionMeshOptions, path As CGeometryPath, edge As TRegionEdge, i As Long
    Dim annotations As CSectionAnnotations, annotationBuilder As CSectionAnnotationBuilder, rebars As CRebarLayout
    Dim facade As CSectionDefinitionBuilder, polygon As CPolygonSpec, errorCode As Long, reason As String
    mesh.StepX = 50#: mesh.StepY = 50#: mesh.BoundarySubdivisions = 2: mesh.MaterialID = 1
    Set definition = BuildAudit02DemoDefinition()
    Set builder = New CSectionBuilder: Set model = builder.Build(definition, mesh)
    Check "D01.demo.bars", model.RebarCount = 15
    Set path = definition.Path("Boundary"): path.CopySegment path.SegmentIndex("Edge.TR"), edge
    For i = 1 To model.Annotations.Count
        If model.Annotations.AnnotationID(i) = "DIM_SLOPE" Then
            Check "A02.slope.start", model.Annotations.StartX(i) = edge.X1 And model.Annotations.StartY(i) = edge.Y1
            Check "A02.slope.end", model.Annotations.EndX(i) = edge.X2 And model.Annotations.EndY(i) = edge.Y2
            Check "A02.slope.length", Abs(model.Annotations.Value(i) - Sqr((edge.X2 - edge.X1) ^ 2 + (edge.Y2 - edge.Y1) ^ 2)) < 0.000001
            Check "A02.slope.parallel", Abs((edge.X2 - edge.X1) * model.Annotations.NormalX(i) + _
                (edge.Y2 - edge.Y1) * model.Annotations.NormalY(i)) < 0.000001
        End If
    Next i
    Set definition = BuildAudit02OperationsDefinition()
    Set model = builder.Build(definition, mesh)
    Check "D01.operations.build", model.RebarCount > 20 And model.ConcreteCount > 0
    Set annotations = New CSectionAnnotations: Set annotationBuilder = New CSectionAnnotationBuilder
    annotationBuilder.Build definition, builder.Rebars, annotations, builder.Rules
    Check "A05.annotations.repeat", annotations.Count = model.Annotations.Count And builder.Rebars.Count = model.RebarCount
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Zero"
    Set path = facade.RectanglePath("Boundary", 0#, 0#, 100#, 100#): facade.Outer "Outer", path.ID
    facade.Anchor "A", 20#, 20#: facade.Anchor "B", 20#, 20#
    facade.Source "ZERO", "Config!TestZero": facade.Dimension "ZERO", "A", "B", "L", "Aligned"
    Set definition = facade.Finish: Set rebars = New CRebarLayout
    On Error Resume Next
    annotationBuilder.Build definition, rebars, annotations
    errorCode = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    Check "A06.zero.rejected.address", errorCode <> 0 And InStr(reason, "ZERO") > 0 And InStr(reason, "Config!TestZero") > 0
    Check "A06.no.replacement", annotations.Count = 0
    Set polygon = New CPolygonSpec
    polygon.AddVertex "TL", -300#, 500#, 300#: polygon.AddVertex "TR", 300#, 500#, 300#
    polygon.AddVertex "BR", 300#, -500#, 300#: polygon.AddVertex "BL", -300#, -500#, 300#: polygon.Seal
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "CollapsedDimension"
    Set path = facade.RoundedPolygon("OuterPath", polygon): facade.Outer "Outer", path.ID
    Set path = facade.OffsetRoundedPolygon("OffsetPath", polygon, 50#)
    facade.PathAnchor "A", path.ID, "Edge.TL", 0#: facade.PathAnchor "B", path.ID, "Edge.TL", 1#
    facade.Source "COLLAPSED", "Config!TestCollapsed": facade.Dimension "COLLAPSED", "A", "B", "L", "Aligned"
    Set definition = facade.Finish
    On Error Resume Next
    annotationBuilder.Build definition, rebars, annotations
    errorCode = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    Check "A06.collapsed.rejected.address", errorCode <> 0 And InStr(reason, "COLLAPSED") > 0 And InStr(reason, "Config!TestCollapsed") > 0
    Check "A06.collapsed.no.replacement", annotations.Count = 0
    GoTo Finished
Failed:
    Check "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunAudit02GuideExamplesTests = mReport & "TOTAL AUDIT02 GUIDE EXAMPLES: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: пример из Audit02, без подбора достаточного армирования.
Private Function BuildAudit02DemoDefinition() As CSectionDefinition
    Dim facade As CSectionDefinitionBuilder, polygon As CPolygonSpec
    Dim shapePath As CGeometryPath, offsetPath As CGeometryPath, rightAxis As CGeometryPath
    Dim options As CRebarRuleOptions, b As Double, h As Double, w As Double
    b = 600#: h = 1000#: w = 200#
    Set polygon = New CPolygonSpec
    polygon.AddVertex "TL", -b / 2#, h / 2#, 80#, outgoingEdgeID:="Edge.TL" ' Верхняя грань для привязки бокового ряда, мм.
    polygon.AddVertex "TR", b / 2#, h / 2#, 120#, outgoingEdgeID:="Edge.TR" ' Наклонная грань для размерной аннотации, мм.
    polygon.AddVertex "RightNose", b / 2# + w, 0#, 80#
    polygon.AddVertex "BR", b / 2#, -h / 2#, 120#, outgoingEdgeID:="Edge.BR" ' Нижняя грань для конца бокового ряда, мм.
    polygon.AddVertex "BL", -b / 2#, -h / 2#, 80#, outgoingEdgeID:="Edge.BL" ' Левая грань от нижнего угла к верхнему, мм.
    polygon.Seal
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Audit02Demo"
    Set shapePath = facade.RoundedPolygon("Boundary", polygon)
    facade.Outer "OUTER", shapePath.ID
    Set offsetPath = facade.OffsetRoundedPolygon("Offset.H.Right", polygon, 50#)
    Set rightAxis = facade.Chain("H.Right", offsetPath.ID, "Edge.TL", 1#, "Edge.BR", 0#)
    Set options = New CRebarRuleOptions
    options.AnnotationGroup = "H.Right" ' В одной подписи группы суммируем стержни каждого диаметра, включая все её ряды.
    options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
    options.Coincidence = KeepPublishedOwner ' При совпадении центров сохраняем стержень, который был добавлен раньше.
    options.EndMode = ExcludeBoth: options.StartSetback = 100#: options.EndSetback = 150#
    options.IncomingTangentAtJoint = True
    facade.RebarByCount "H.Right.Main", rightAxis.ID, 15, 20#, options
    facade.RebarRowExtentLabel "H.Right", "H.Right.Main", 1#, 0#, True
    facade.PathAnchor "Slope.A", shapePath.ID, "Edge.TR", 0#
    facade.PathAnchor "Slope.B", shapePath.ID, "Edge.TR", 1#
    facade.Dimension "DIM_SLOPE", "Slope.A", "Slope.B", "L", "Aligned"
    facade.HorizontalDistance "DIM_B", "B", -b / 2#, b / 2#, h / 2#, 1#
    facade.VerticalExtent "DIM_H", shapePath.ID, "H", -b / 2#, -1#
    Set BuildAudit02DemoDefinition = facade.Finish
End Function

' ДЛЯ ТЕСТОВ: независимые примитивы, signed-left offset, распределения и проекция.
Private Function BuildAudit02OperationsDefinition() As CSectionDefinition
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "OperationsExample"
    Set path = facade.AdjacentRectangles("Boundary", Array(Array(0#, 0#, 1000#, 1000#), Array(1000#, 0#, 200#, 600#)))
    facade.Outer "Outer", path.ID
    Set path = facade.CirclePath("Hole", 900#, 800#, 50#): facade.Opening "Opening", path.ID, "Outer"
    Set path = facade.LinePath("MainAxis", 100#, 100#, 700#, 100#)
    Set path = facade.ArcPath("Arc", 500#, 600#, 100#, 0#, GEOM_PI / 2#)
    Set path = facade.Offset("ProjectedAxis", "MainAxis", 100#)
    Set path = facade.Transform("Rotated", "MainAxis", 900#, 100#, GEOM_PI / 2#, True)
    Set path = facade.ConstrainLine("Limited", "MainAxis", 1#, 0#, 150#, 1#, 0#, 650#)
    Set path = facade.LinePath("MaxStepAxis", 100#, 300#, 700#, 300#)
    Set path = facade.LinePath("FixedStepAxis", 100#, 400#, 700#, 400#)
    Set path = facade.LinePath("ExplicitAxis", 100#, 500#, 700#, 500#)
    facade.PathGroup "TwoLines", Array("MaxStepAxis", "FixedStepAxis")
    Set options = New CRebarRuleOptions
    options.AnnotationGroup = "Bottom" ' В одной подписи группы суммируем стержни каждого диаметра, включая все её ряды.
    options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
    options.Coincidence = KeepPublishedOwner ' При совпадении центров сохраняем стержень, который был добавлен раньше.
    facade.RebarByCount "Main", "MainAxis", 5, 20#, options
    Set options = New CRebarRuleOptions
    options.AnnotationGroup = "Bottom" ' В одной подписи группы суммируем стержни каждого диаметра, включая все её ряды.
    options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
    options.Coincidence = KeepPublishedOwner ' При совпадении центров сохраняем стержень, который был добавлен раньше.
    options.ConfigureDependentContact 20#, 16#, RowNormal, 1#, False
    facade.RebarDependent "Contact", "Main", 16#, options
    Set options = New CRebarRuleOptions
    options.NormalOffset = 40#: options.TangentOffset = 10#
    facade.RebarDependent "Shifted", "Main", 12#, options
    Set options = New CRebarRuleOptions: options.DirectionY = 1#
    facade.RebarProjected "Projected", "ProjectedAxis", "Main", 16#, options
    facade.RebarByMaxStep "MaxStep", "MaxStepAxis", 150#, 16#
    facade.RebarByFixedStep "FixedStep", "FixedStepAxis", 150#, 16#
    facade.RebarAtStations "Explicit", "ExplicitAxis", Array(0#, 150#, 600#), 12#
    facade.Anchor "A", 100#, 100#: facade.Anchor "B", 700#, 500#
    facade.TransformAnchor "MovedA", "A", 0#, 700#
    facade.Dimension "H", "A", "B", "X", "Horizontal"
    facade.Dimension "V", "A", "B", "Y", "Vertical"
    facade.Dimension "L", "A", "B", "L", "Aligned"
    facade.PathLength "Length", "TwoLines", "MovedA", "SumL"
    facade.Radius "Radius", "Arc", "Arc", 0.5, "R"
    facade.Diameter "Diameter", "Hole", "Hole", 0.5, "D"
    facade.Text "Note", "MovedA", "Учебный пример"
    facade.RebarRowLabel "Bottom", "Main", 0#, -1#
    Set BuildAudit02OperationsDefinition = facade.Finish
End Function

' ДЛЯ ТЕСТОВ: реальная Excel-схема наклонного размера и проверка повторного Draw.
Public Function DumpAudit02DemoPreview(ByVal outputDirectory As String) As String
    Dim definition As CSectionDefinition, builder As CSectionBuilder, mesh As TSectionMeshOptions, model As CSectionModel
    Set definition = BuildAudit02DemoDefinition(): Set builder = New CSectionBuilder
    mesh.StepX = 50#: mesh.StepY = 50#: mesh.BoundarySubdivisions = 2: mesh.MaterialID = 1
    Set model = builder.Build(definition, mesh)
    Dim first As String, second As String
    first = modTestShapeArchitecture.PreparedPreviewSnapshot("Audit02Demo", model, outputDirectory)
    second = modTestShapeArchitecture.PreparedPreviewSnapshot("Audit02Demo", model, outputDirectory)
    If first <> second Then Err.Raise vbObjectError + 5490, "modTestAudit02Api", "Повторный Draw изменил данные/Shapes."
    DumpAudit02DemoPreview = first
End Function

' ДЛЯ ТЕСТОВ: настройка контактного ряда отвергает нулевые диаметры,
' а при отказе не изменяет переданные параметры.
Public Function RunAudit02CloseoutContactTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim options As CRebarRuleOptions, cases As Variant, item As Variant, code As Long, reason As String
    Set options = New CRebarRuleOptions
    options.SourceReference = "Program.Contact": options.StartSetback = 100#: options.EndSetback = 150#
    options.AnnotationGroup = "Group" ' В одной подписи группы суммируем стержни каждого диаметра, включая все её ряды.
    options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
    options.Coincidence = KeepPublishedOwner ' При совпадении центров сохраняем стержень, который был добавлен раньше.
    options.NormalOffset = 12#: options.TangentOffset = 34#: options.EverySecond = True
    cases = Array(Array(0#, 16#, RowNormal, 1#, 0#), Array(-20#, 16#, RowNormal, 1#, 0#), _
        Array(20#, 0#, RowNormal, 1#, 0#), Array(20#, -16#, RowNormal, 1#, 0#), _
        Array(20#, 16#, RowNormal, 1#, -1#), Array(20#, 16#, 0, 1#, 0#), Array(20#, 16#, RowNormal, 0#, 0#))
    For Each item In cases
        On Error Resume Next
        options.ConfigureDependentContact CDbl(item(0)), CDbl(item(1)), CLng(item(2)), CDbl(item(3)), False, CDbl(item(4))
        code = Err.Number: reason = Err.Description: Err.Clear
        On Error GoTo Failed
        Check "contact.reject.address." & CStr(item(0)) & "." & CStr(item(1)) & "." & CStr(item(2)) & "." & CStr(item(3)), _
            code <> 0 And InStr(reason, "Program.Contact") > 0
        Check "contact.reject.atomic", options.NormalOffset = 12# And options.TangentOffset = 34# And options.EverySecond
        Check "contact.unrelated.fields", options.StartSetback = 100# And options.EndSetback = 150# And _
            options.AnnotationGroup = "Group" And options.Publication = PublishPerSourcePosition And options.Coincidence = KeepPublishedOwner
    Next item
    options.ConfigureDependentContact 20#, 16#, RowNormal, -1#, False
    Check "contact.valid.18", options.NormalOffset = -18# And options.TangentOffset = 0# And Not options.EverySecond
    options.ConfigureDependentContact 20#, 12#, RowTangent, 1#, True, 16#
    Check "contact.valid.32", options.NormalOffset = 0# And options.TangentOffset = 32# And options.EverySecond
    GoTo Finished
Failed:
    Check "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunAudit02CloseoutContactTests = mReport & "TOTAL AUDIT02 CLOSEOUT CONTACT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: адресная матрица размеров, неактивных полей, знаков и порядка диагностик RectSet.
Public Function RunAudit02CloseoutRectSetInputTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim records As String
    records = CloseoutRectSetInputMatrix()
    RunAudit02CloseoutRectSetInputTests = mReport & "TOTAL AUDIT02 CLOSEOUT RECTSET INPUT: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: полный снимок значений/ошибок той же матрицы для побайтной сверки baseline/candidate.
Public Function DumpAudit02CloseoutInputContract() As String
    mPassed = 0: mFailed = 0: mReport = ""
    DumpAudit02CloseoutInputContract = CloseoutRectSetInputMatrix()
    If mFailed > 0 Then Err.Raise vbObjectError + 5491, "modTestAudit02Api", mReport
End Function

' ДЛЯ ТЕСТОВ: только три схемы RectSet, без полного набора UI/ON-OFF.
Public Function DumpAudit02CloseoutRectSetPreview(ByVal outputDirectory As String) As String
    Dim fixture As Variant
    For Each fixture In Array("Rectangle", "LSection", "TwoRectangles")
        DumpAudit02CloseoutRectSetPreview = DumpAudit02CloseoutRectSetPreview & _
            modTestShapeArchitecture.DumpShapePreviewSnapshot(outputDirectory, CStr(fixture))
    Next fixture
End Function

' Записываем номер ошибки, её источник и полный текст с фактическим адресом ячейки.
' Сравнение должно выявлять любое различие.
Private Function CloseoutRectSetInputMatrix() As String
    Dim keyed As CSectionInputReader, fieldKeys As Variant
    Dim config As Object, saved As Variant, range As Object, block As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, definition As CSectionDefinition
    Dim reader As CSectionInputSnapshotReader, context As CSectionRecipeContext, table As CSectionInputTable
    Dim fixture As Variant, valueKind As Variant, unit As Variant, aliasName As Variant
    Dim row As Long, column As Long, maxColumn As Long, offsetRow As Long, shapeRow As Long
    Dim code As Long, source As String, reason As String, expectedAddress As String, area As Double
    Dim result As String, factor As Double, pair As Variant, number As Double, otherCode As Long, otherReason As String
    fieldKeys = Array("H1", "B1", "H2", "B2")
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    For Each unit In Array("mm", "cm", "m")
        config.Formula = saved
        ThisWorkbook.Names("rngUnitSettings").RefersToRange.Cells(2, 2).Value2 = CStr(unit)
        For Each fixture In Array("Rectangle", "LSection", "TwoRectangles")
            modTestShapeArchitecture.ConfigureShapeFixture "Circle"
            modTestShapeArchitecture.ConfigureShapeFixture CStr(fixture)
            Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
            Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
            Set range = ThisWorkbook.Names("rngRectSetGeometry").RefersToRange: block = range.Formula
            Set reader = New CSectionInputSnapshotReader
            Set context = reader.Capture(ThisWorkbook, Array("rngRectSetGeometry", "rngRebarInput"), units): Set table = context.Table("rngRectSetGeometry"): Set keyed = context.Reader("rngRectSetGeometry")
            offsetRow = keyed.RowIndex("Detail", "XOffset")
            shapeRow = keyed.RowIndex("General", "Shape")
            Set definition = RectSetInputOutcome(settings, units, code, source, reason)
            Check "rect.valid." & CStr(unit) & "." & CStr(fixture), code = 0
            If code <> 0 Then GoTo Failed
            area = definition.ConcreteDomain.Area
            result = result & InputContractRow(CStr(unit), CStr(fixture), "base", code, source, reason, area)
            maxColumn = 4: If CStr(fixture) = "Rectangle" Then maxColumn = 2
            For column = 1 To maxColumn
                expectedAddress = keyed.SourceReference("General", CStr(fieldKeys(column - 1)))
                For Each valueKind In Array("positive", "zero", "negative", "empty", "text")
                    range.Formula = block
                    Select Case CStr(valueKind)
                        Case "zero": FixtureInputCell(range, "General", CStr(fieldKeys(column - 1))).Value2 = 0#
                        Case "negative": FixtureInputCell(range, "General", CStr(fieldKeys(column - 1))).Value2 = -1# / factor
                        Case "empty": FixtureInputCell(range, "General", CStr(fieldKeys(column - 1))).ClearContents
                        Case "text": FixtureInputCell(range, "General", CStr(fieldKeys(column - 1))).Value2 = "bad dimension"
                    End Select
                    Set definition = RectSetInputOutcome(settings, units, code, source, reason)
                    If CStr(valueKind) = "positive" Then
                        Check "rect.positive." & CStr(unit) & "." & CStr(fixture) & "." & CStr(column), code = 0
                        Check "rect.unit.once", definition.ConcreteDomain.Area = area
                    Else
                        Check "rect.reject.address." & CStr(unit) & "." & CStr(fixture) & "." & CStr(column) & "." & CStr(valueKind), _
                            code <> 0 And InStr(reason, expectedAddress) > 0
                    End If
                    result = result & InputContractRow(CStr(unit), CStr(fixture), CStr(column) & "." & CStr(valueKind), code, source, reason, 0#)
                Next valueKind
            Next column
            ' Для Rectangle нижние размеры и смещение не читаются; для LSection смещение также неактивно.
            range.Formula = block
            If CStr(fixture) = "Rectangle" Then
                FixtureInputCell(range, "General", "H2").Value2 = "inactive H2": FixtureInputCell(range, "General", "B2").Value2 = CVErr(xlErrValue)
            End If
            If CStr(fixture) <> "TwoRectangles" Then
                range.Cells(offsetRow, 3).Value2 = "inactive offset"
                Set definition = RectSetInputOutcome(settings, units, code, source, reason)
                Check "rect.inactive.fields." & CStr(unit) & "." & CStr(fixture), code = 0 And definition.ConcreteDomain.Area = area
                result = result & InputContractRow(CStr(unit), CStr(fixture), "inactive", code, source, reason, 0#)
            Else
                For Each pair In Array(-100#, 0#, 100#)
                    range.Formula = block: range.Cells(offsetRow, 3).Value2 = CDbl(pair) / factor
                    Set definition = RectSetInputOutcome(settings, units, code, source, reason)
                    Check "rect.signed.offset." & CStr(unit) & "." & CStr(pair), code = 0
                    result = result & InputContractRow(CStr(unit), CStr(fixture), "offset." & CStr(pair), code, source, reason, 0#)
                Next pair
            End If
            ' Во взаимосвязанных полях сначала проверяется, что второе поле числовое,
            ' и только затем — положительность первого.
            For Each pair In Array(1, 3)
                If CLng(pair) <= maxColumn Then
                    range.Formula = block: FixtureInputCell(range, "General", CStr(fieldKeys(CLng(pair) - 1))).Value2 = 0#
                    FixtureInputCell(range, "General", CStr(fieldKeys(CLng(pair)))).Value2 = "bad dimension"
                    Set definition = RectSetInputOutcome(settings, units, code, source, reason)
                    Check "rect.pair.numeric.before.positive." & CStr(unit) & "." & CStr(fixture) & "." & CStr(pair), _
                        code <> 0 And InStr(reason, keyed.SourceReference("General", CStr(fieldKeys(CLng(pair))))) > 0
                    result = result & InputContractRow(CStr(unit), CStr(fixture), "pair." & CStr(pair), code, source, reason, 0#)
                    Set context = reader.Capture(ThisWorkbook, Array("rngRectSetGeometry", "rngRebarInput"), units): Set table = context.Table("rngRectSetGeometry"): Set keyed = context.Reader("rngRectSetGeometry")
                    On Error Resume Next
                    number = keyed.PositiveLength("General", CStr(fieldKeys(CLng(pair) - 1)))
                    otherCode = Err.Number: otherReason = Err.Description: Err.Clear
                    On Error GoTo Failed
                    Check "rect.direct.helper.changes.first.address", otherCode <> 0 And _
                        InStr(otherReason, keyed.SourceReference("General", CStr(fieldKeys(CLng(pair) - 1)))) > 0 And otherReason <> reason
                End If
            Next pair
            range.Formula = block
            ' Составные условия остаются у рецепта и указывают соответствующее поле.
            If CStr(fixture) = "LSection" Then
                FixtureInputCell(range, "General", "B1").Value2 = FixtureInputCell(range, "General", "B2").Value2
                Set definition = RectSetInputOutcome(settings, units, code, source, reason)
                Check "rect.LSection.B1.less.B2", code <> 0 And InStr(reason, keyed.SourceReference("General", "B1")) > 0
                result = result & InputContractRow(CStr(unit), CStr(fixture), "incompatible.widths", code, source, reason, 0#)
            ElseIf CStr(fixture) = "TwoRectangles" Then
                range.Cells(offsetRow, 3).Value2 = FixtureInputCell(range, "General", "B2").Value2
                Set definition = RectSetInputOutcome(settings, units, code, source, reason)
                Check "rect.adjacency.address", code <> 0 And InStr(reason, table.SourceReference(offsetRow, 3)) > 0
                result = result & InputContractRow(CStr(unit), CStr(fixture), "nonadjacent", code, source, reason, 0#)
            End If
            For Each aliasName In Array("RectSet", "TwoRectangle", "TwoRectanglesOffset", "OffsetTwoRectangles")
                If (CStr(fixture) = "LSection" And CStr(aliasName) = "RectSet") Or _
                        (CStr(fixture) = "TwoRectangles" And CStr(aliasName) <> "RectSet") Then
                    range.Formula = block: range.Cells(shapeRow, 3).Value2 = CStr(aliasName)
                    Set definition = RectSetInputOutcome(settings, units, code, source, reason)
                    Check "rect.alias." & CStr(unit) & "." & CStr(aliasName), code = 0 And definition.ConcreteDomain.Area = area
                    result = result & InputContractRow(CStr(unit), CStr(fixture), CStr(aliasName), code, source, reason, 0#)
                End If
            Next aliasName
            range.Formula = block
        Next fixture
    Next unit
    GoTo Finished
Failed:
    Check "rect.runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    CloseoutRectSetInputMatrix = result
End Function

' Вызывает обычное чтение описания сечения и возвращает полученную ошибку.
' Ожидаемый номер и текст проверяет вызывающий тест.
Private Function RectSetInputOutcome(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByRef code As Long, ByRef source As String, ByRef reason As String) As CSectionDefinition
    code = 0: source = "": reason = ""
    On Error GoTo Invalid
    Set RectSetInputOutcome = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
    Exit Function
Invalid:
    code = Err.Number: source = Err.Source: reason = Err.Description
End Function

Private Function InputContractRow(ByVal unit As String, ByVal fixture As String, ByVal name As String, _
        ByVal code As Long, ByVal source As String, ByVal reason As String, ByVal area As Double) As String
    InputContractRow = unit & "|" & fixture & "|" & name & "|" & CStr(code) & "|" & source & "|" & reason & "|" & CStr(area) & vbCrLf
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

' ДЛЯ ТЕСТОВ: контактные последовательности, независимые настройки и сохранение состояния после ошибки.
Public Function RunRebarContactBuilderTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    TestContactSequence
    TestContactDefaults
    TestContactSigns
    TestContactFailures
    GoTo Finished
Failed:
    Check "contactBuilder.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunRebarContactBuilderTests = mReport & "TOTAL REBAR CONTACT BUILDER: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: пример 32 + 20 + 16 по нормали и 12 вдоль линии; все ряды относятся к одному Main.
Private Sub TestContactSequence()
    Dim contact As CRebarContactBuilder, defaults As CRebarRuleOptions, row As CRebarRuleOptions
    Set defaults = New CRebarRuleOptions
    Set contact = New CRebarContactBuilder
    contact.Initialize 32#, 1#, 1#, defaults
    Set row = contact.AppendRow(20#, "Stacked", "EverySecondBar", "ContactTest[1,2]")
    Check "contactBuilder.example.first", row.NormalOffset = 28.5 And row.TangentOffset = 0#
    Check "contactBuilder.example.binding", row.EverySecond
    Check "contactBuilder.example.source", row.SourceReference = "ContactTest[1,2]"
    Set row = contact.AppendRow(16#, "Stacked", "EachBar", "ContactTest[2,2]")
    Check "contactBuilder.example.second", row.NormalOffset = 49# And row.TangentOffset = 0#
    Check "contactBuilder.example.nominal", Not row.EverySecond
    Set row = contact.AppendRow(12#, "SideBySide", "EachBar", "ContactTest[3,2]")
    Check "contactBuilder.example.third", row.TangentOffset = 24.5 And row.NormalOffset = 0#
    Set row = contact.AppendRow(10#, "Stacked", "EverySecondBar", "ContactTest[4,2]")
    Check "contactBuilder.example.fourth", row.NormalOffset = 64.5 And row.TangentOffset = 0#
    Set row = contact.AppendRow(8#, "SideBySide", "EachBar", "ContactTest[5,2]")
    Check "contactBuilder.example.fifth", row.TangentOffset = 37# And row.NormalOffset = 0#
End Sub

' ДЛЯ ТЕСТОВ: defaults, его снимок и каждый результат независимы; индивидуальные поля можно менять.
Private Sub TestContactDefaults()
    Dim contact As CRebarContactBuilder, defaults As CRebarRuleOptions, first As CRebarRuleOptions, second As CRebarRuleOptions, copy As CRebarRuleOptions
    Set defaults = New CRebarRuleOptions
    defaults.StartSetback = 7#: defaults.EndSetback = 9#: defaults.EndMode = ExcludeBoth
    defaults.SourcePrefix = "Extra.": defaults.SourceSuffix = ".Default": defaults.Comment = "Общие настройки"
    defaults.AnnotationGroup = "ExtraGroup": defaults.Publication = PublishPerSourcePosition: defaults.Coincidence = KeepPublishedOwner
    defaults.NormalOffset = 90#: defaults.TangentOffset = -90#: defaults.EverySecond = True
    defaults.PlacementCheck = CheckFullDisk: defaults.IncomingTangentAtJoint = True
    defaults.DirectionX = 2#: defaults.DirectionY = -3#: defaults.SourceNormalSign = -1#
    defaults.SuccessfulProjectionOrdinal = True: defaults.NewPublicationRoot = True: defaults.SkipMissingProjection = True
    defaults.PositionOrder = FromPathMiddle: defaults.OverlapPolicy = RejectOverlap
    defaults.SourceReference = "DefaultsTest[1,2]"
    Set copy = defaults.Clone
    Check "contactBuilder.clone.StartSetback", copy.StartSetback = defaults.StartSetback
    Check "contactBuilder.clone.EndSetback", copy.EndSetback = defaults.EndSetback
    Check "contactBuilder.clone.EndMode", copy.EndMode = defaults.EndMode
    Check "contactBuilder.clone.DistributionScope", copy.DistributionScope = defaults.DistributionScope
    Check "contactBuilder.clone.RemainderPlacement", copy.RemainderPlacement = defaults.RemainderPlacement
    Check "contactBuilder.clone.TangentOffset", copy.TangentOffset = defaults.TangentOffset
    Check "contactBuilder.clone.NormalOffset", copy.NormalOffset = defaults.NormalOffset
    Check "contactBuilder.clone.EverySecond", copy.EverySecond = defaults.EverySecond
    Check "contactBuilder.clone.IncomingTangentAtJoint", copy.IncomingTangentAtJoint = defaults.IncomingTangentAtJoint
    Check "contactBuilder.clone.DirectionX", copy.DirectionX = defaults.DirectionX
    Check "contactBuilder.clone.DirectionY", copy.DirectionY = defaults.DirectionY
    Check "contactBuilder.clone.SourceNormalSign", copy.SourceNormalSign = defaults.SourceNormalSign
    Check "contactBuilder.clone.SuccessfulProjectionOrdinal", copy.SuccessfulProjectionOrdinal = defaults.SuccessfulProjectionOrdinal
    Check "contactBuilder.clone.NewPublicationRoot", copy.NewPublicationRoot = defaults.NewPublicationRoot
    Check "contactBuilder.clone.SkipMissingProjection", copy.SkipMissingProjection = defaults.SkipMissingProjection
    Check "contactBuilder.clone.Publication", copy.Publication = defaults.Publication
    Check "contactBuilder.clone.PositionOrder", copy.PositionOrder = defaults.PositionOrder
    Check "contactBuilder.clone.Coincidence", copy.Coincidence = defaults.Coincidence
    Check "contactBuilder.clone.PlacementCheck", copy.PlacementCheck = defaults.PlacementCheck
    Check "contactBuilder.clone.OverlapPolicy", copy.OverlapPolicy = defaults.OverlapPolicy
    Check "contactBuilder.clone.SourcePrefix", copy.SourcePrefix = defaults.SourcePrefix
    Check "contactBuilder.clone.SourceSuffix", copy.SourceSuffix = defaults.SourceSuffix
    Check "contactBuilder.clone.AnnotationGroup", copy.AnnotationGroup = defaults.AnnotationGroup
    Check "contactBuilder.clone.Comment", copy.Comment = defaults.Comment
    Check "contactBuilder.clone.SourceReference", copy.SourceReference = defaults.SourceReference
    copy.StartSetback = 111#: copy.AnnotationGroup = "CopyOnly"
    Check "contactBuilder.clone.independent", defaults.StartSetback = 7# And defaults.AnnotationGroup = "ExtraGroup"
    Set contact = New CRebarContactBuilder
    contact.Initialize 32#, 1#, -1#, defaults
    defaults.StartSetback = 999#: defaults.EndMode = ClosedUniform: defaults.SourcePrefix = "Changed."
    defaults.AnnotationGroup = "Changed": defaults.PlacementCheck = CheckCenter
    defaults.SourceReference = "ChangedDefaults"
    Set first = contact.AppendRow(20#, "Stacked", "EachBar", "FirstTest[1,2]")
    Check "contactBuilder.snapshot.setbacks", first.StartSetback = 7# And first.EndSetback = 9#
    Check "contactBuilder.snapshot.endMode", first.EndMode = ExcludeBoth
    Check "contactBuilder.snapshot.names", first.SourcePrefix = "Extra." And first.SourceSuffix = ".Default"
    Check "contactBuilder.snapshot.policies", first.AnnotationGroup = "ExtraGroup" And first.PlacementCheck = CheckFullDisk
    Check "contactBuilder.snapshot.source", first.SourceReference = "FirstTest[1,2]"
    Check "contactBuilder.snapshot.overrides", first.NormalOffset = 28.5 And first.TangentOffset = 0# And Not first.EverySecond
    first.StartSetback = 123#: first.AnnotationGroup = "FirstOnly": first.SourceSuffix = "_R2": first.Comment = "Индивидуальный ряд"
    first.NormalOffset = -1000#: first.PlacementCheck = CheckCenter
    Set second = contact.AppendRow(16#, "Stacked", "EverySecondBar", "SecondTest[1,2]")
    Check "contactBuilder.rows.independent", second.StartSetback = 7# And second.AnnotationGroup = "ExtraGroup"
    Check "contactBuilder.rows.names", second.SourceSuffix = ".Default" And second.Comment = "Общие настройки"
    Check "contactBuilder.rows.nominal", second.NormalOffset = 49# And second.TangentOffset = 0# And second.EverySecond
    Check "contactBuilder.rows.placement", second.PlacementCheck = CheckFullDisk
    Check "contactBuilder.rows.firstRetained", first.SourceSuffix = "_R2" And first.Comment = "Индивидуальный ряд" And first.NormalOffset = -1000#
End Sub

' ДЛЯ ТЕСТОВ: все сочетания знаков, перемежающиеся направления и EverySecondBar.
Private Sub TestContactSigns()
    Dim contact As CRebarContactBuilder, defaults As CRebarRuleOptions, row As CRebarRuleOptions, normal As Variant, tangent As Variant
    Set defaults = New CRebarRuleOptions
    For Each normal In Array(-1#, 1#)
        For Each tangent In Array(-1#, 1#)
            Set contact = New CRebarContactBuilder
            contact.Initialize 32#, CDbl(normal), CDbl(tangent), defaults
            Set row = contact.AppendRow(20#, "Stacked", "EachBar", "Signs[1,2]")
            Check "contactBuilder.sign.normal.first", row.NormalOffset = CDbl(normal) * 28.5 And row.TangentOffset = 0#
            Set row = contact.AppendRow(12#, "SideBySide", "EverySecondBar", "Signs[2,2]")
            Check "contactBuilder.sign.tangent.first", row.TangentOffset = CDbl(tangent) * 24.5 And row.NormalOffset = 0#
            Set row = contact.AppendRow(16#, "Stacked", "EverySecondBar", "Signs[3,2]")
            Check "contactBuilder.sign.normal.next", row.NormalOffset = CDbl(normal) * 49# And row.EverySecond
            Set row = contact.AppendRow(8#, "SideBySide", "EachBar", "Signs[4,2]")
            Check "contactBuilder.sign.tangent.next", row.TangentOffset = CDbl(tangent) * 37# And Not row.EverySecond
        Next tangent
    Next normal
End Sub

' ДЛЯ ТЕСТОВ: ошибки не расходуют диаметр и не меняют независимый накопитель второго направления.
Private Sub TestContactFailures()
    Dim contact As CRebarContactBuilder, defaults As CRebarRuleOptions, row As CRebarRuleOptions, scenario As Long
    Set defaults = New CRebarRuleOptions: defaults.SourceReference = "InitializeTest[1,2]"
    Set contact = New CRebarContactBuilder
    Check "contactBuilder.uninitialized", ContactAppendRejected(contact, 20#, "Stacked", "EachBar")
    For scenario = 0 To 4
        Set contact = New CRebarContactBuilder
        Check "contactBuilder.initialize.reject." & CStr(scenario), ContactInitializeRejected(contact, defaults, scenario)
        contact.Initialize 32#, 1#, -1#, defaults
        Set row = contact.AppendRow(20#, "Stacked", "EachBar", "AfterInitializeTest[1,2]")
        Check "contactBuilder.initialize.recover." & CStr(scenario), row.NormalOffset = 28.5
    Next scenario
    For scenario = 0 To 6
        Set contact = New CRebarContactBuilder
        contact.Initialize 32#, 1#, -1#, defaults
        Set row = contact.AppendRow(20#, "Stacked", "EachBar", "BeforeFailure[1,2]")
        Select Case scenario
            Case 0: Check "contactBuilder.reject.location", ContactAppendRejected(contact, 100#, "Unknown", "EachBar")
            Case 1: Check "contactBuilder.reject.emptyLocation", ContactAppendRejected(contact, 100#, "", "EachBar")
            Case 2: Check "contactBuilder.reject.binding", ContactAppendRejected(contact, 100#, "Stacked", "Unknown")
            Case 3: Check "contactBuilder.reject.emptyBinding", ContactAppendRejected(contact, 100#, "SideBySide", "")
            Case 4: Check "contactBuilder.reject.zero", ContactAppendRejected(contact, 0#, "Stacked", "EachBar")
            Case 5: Check "contactBuilder.reject.negative", ContactAppendRejected(contact, -10#, "SideBySide", "EachBar")
            Case 6: Check "contactBuilder.reject.reinitialize", ContactInitializeRejected(contact, defaults, 5)
        End Select
        Set row = contact.AppendRow(16#, "Stacked", "EverySecondBar", "AfterFailure[1,2]")
        Check "contactBuilder.failure.normal." & CStr(scenario), row.NormalOffset = 49#
        Set row = contact.AppendRow(12#, "SideBySide", "EachBar", "AfterFailure[2,2]")
        Check "contactBuilder.failure.tangent." & CStr(scenario), row.TangentOffset = -24.5
    Next scenario
    For scenario = 0 To 1
        Set contact = New CRebarContactBuilder
        contact.Initialize 1#, 1#, -1#, defaults
        Dim huge As Double
        huge = 1E+308: If scenario = 1 Then huge = 1.4E+308
        Set row = contact.AppendRow(huge, "Stacked", "EachBar", "HugeTest[1,2]")
        Check "contactBuilder.overflow.reject." & CStr(scenario), ContactAppendRejected(contact, huge, "Stacked", "EachBar")
        Set row = contact.AppendRow(1#, "Stacked", "EachBar", "HugeRecovery[1,2]")
        Check "contactBuilder.overflow.recover." & CStr(scenario), row.NormalOffset = huge
    Next scenario
End Sub

' ДЛЯ ТЕСТОВ: ожидает адресную ошибку AppendRow без регистрации правила и без размещения стержней.
Private Function ContactAppendRejected(ByVal contact As CRebarContactBuilder, ByVal diameter As Double, ByVal location As String, ByVal binding As String) As Boolean
    Dim row As CRebarRuleOptions
    On Error GoTo Rejected
    Set row = contact.AppendRow(diameter, location, binding, "FailureTest[1,2]")
    Exit Function
Rejected:
    ContactAppendRejected = Err.Number <> 0 And InStr(1, Err.Description, "FailureTest[1,2]", vbBinaryCompare) > 0
End Function

' ДЛЯ ТЕСТОВ: неверная и повторная инициализация не должны частично менять основной ряд.
Private Function ContactInitializeRejected(ByVal contact As CRebarContactBuilder, ByVal defaults As CRebarRuleOptions, ByVal scenario As Long) As Boolean
    On Error GoTo Rejected
    Select Case scenario
        Case 0: contact.Initialize 0#, 1#, 1#, defaults
        Case 1: contact.Initialize -10#, 1#, 1#, defaults
        Case 2: contact.Initialize 32#, 0#, 1#, defaults
        Case 3: contact.Initialize 32#, 1#, 2#, defaults
        Case 4: contact.Initialize 32#, 1#, 1#, Nothing
        Case 5: contact.Initialize 64#, -1#, 1#, defaults
    End Select
    Exit Function
Rejected:
    ContactInitializeRejected = Err.Number <> 0
    If scenario <> 4 Then ContactInitializeRejected = ContactInitializeRejected And InStr(1, Err.Description, defaults.SourceReference, vbBinaryCompare) > 0
End Function

' ДЛЯ ТЕСТОВ: реальные координаты, фильтр исходных позиций и запрет новых коллизий общим исполнителем.
Public Function RunIndependentRowOffsetGeometryTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim kind As Long, b2 As Variant, b3 As Variant, definition As CSectionDefinition, layout As CRebarLayout
    Dim executor As CRebarRuleExecutor, i As Long, secondCount As Long, thirdCount As Long, radius As Double
    For kind = 0 To 1
        For Each b2 In Array("EachBar", "EverySecondBar")
            For Each b3 In Array("EachBar", "EverySecondBar")
                Set definition = OffsetGeometryDefinition(kind, 100#, 100#, CStr(b2), CStr(b3))
                Set executor = New CRebarRuleExecutor
                Set layout = executor.Build(definition, FixtureGeometry(definition.ConcreteDomain))
                secondCount = 0: thirdCount = 0
                For i = 1 To executor.PositionCount
                    If executor.PositionRuleID(i) = "Row2" Then
                        secondCount = secondCount + 1
                        If kind = 0 Then radius = layout.Y(executor.PositionPhysicalOwner(i)) Else radius = Sqr(layout.X(executor.PositionPhysicalOwner(i)) ^ 2 + layout.Y(executor.PositionPhysicalOwner(i)) ^ 2)
                        Check "offset.geometry.row2." & CStr(kind) & "." & CStr(i), Abs(radius - IIf(kind = 0, 100#, 200#)) < 0.000001
                    ElseIf executor.PositionRuleID(i) = "Row3" Then
                        thirdCount = thirdCount + 1
                        If kind = 0 Then radius = layout.Y(executor.PositionPhysicalOwner(i)) Else radius = Sqr(layout.X(executor.PositionPhysicalOwner(i)) ^ 2 + layout.Y(executor.PositionPhysicalOwner(i)) ^ 2)
                        Check "offset.geometry.row3." & CStr(kind) & "." & CStr(i), Abs(radius - IIf(kind = 0, 120.5, 179.5)) < 0.000001
                    End If
                Next i
                Check "offset.geometry.count2", secondCount = IIf(CStr(b2) = "EachBar", 5, 3)
                Check "offset.geometry.count3", thirdCount = IIf(CStr(b3) = "EachBar", 5, 3)
            Next b3
        Next b2
    Next kind
    For kind = 0 To 1
        Set definition = OffsetGeometryDefinition(kind, 0#, 0#)
        Set executor = New CRebarRuleExecutor
        Set layout = executor.Build(definition, FixtureGeometry(definition.ConcreteDomain))
        For i = 1 To executor.PositionCount
            If executor.PositionRuleID(i) = "Row2" Or executor.PositionRuleID(i) = "Row3" Then
                If kind = 0 Then radius = layout.Y(executor.PositionPhysicalOwner(i)) Else radius = 300# - Sqr(layout.X(executor.PositionPhysicalOwner(i)) ^ 2 + layout.Y(executor.PositionPhysicalOwner(i)) ^ 2)
                If executor.PositionRuleID(i) = "Row2" Then
                    Check "offset.profile.zero.row2", Abs(radius - 28.5) < 0.000001 And layout.Diameter(executor.PositionPhysicalOwner(i)) = 20#
                Else
                    Check "offset.profile.zero.row3", Abs(radius - 49#) < 0.000001 And layout.Diameter(executor.PositionPhysicalOwner(i)) = 16#
                End If
            End If
        Next i
    Next kind
    Check "offset.overlap.rejected", OffsetGeometryRejected(OffsetGeometryDefinition(0, 100#, 105#), "Диски стержней")
    Set layout = FixtureLayout(OffsetGeometryDefinition(0, 100#, 118#))
    Check "offset.touch.allowed", layout.Count = 15
    Check "offset.disk.outside", OffsetGeometryRejected(OffsetGeometryDefinition(0, 1499#, 100#), "целиком")
    Check "offset.arc.center", OffsetGeometryRejected(OffsetGeometryDefinition(1, 300#, 0#), "центра")
    Check "offset.arc.inversion", OffsetGeometryRejected(OffsetGeometryDefinition(1, 301#, 0#), "центра")
    Check "offset.arc.contact.collapse", OffsetGeometryRejected(OffsetGeometryDefinition(1, 290#, 290#), "центра")
    For kind = 0 To 1
        Check "offset.coincidence.order." & CStr(kind), OffsetGeometryRejected(OffsetGeometryDefinition(0, 100#, 200#, "EachBar", "EachBar", kind + 1), "совпал")
        Check "offset.shared.upgrade.order." & CStr(kind), OffsetGeometryRejected(OffsetSharedJointDefinition(kind, True), "Диски стержней")
        Set layout = FixtureLayout(OffsetSharedJointDefinition(kind, False))
        Check "offset.shared.joint.allowed." & CStr(kind), layout.Count = 2
    Next kind
    GoTo Finished
Failed:
    Check "offset.geometry.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunIndependentRowOffsetGeometryTests = mReport & "TOTAL INDEPENDENT ROW OFFSET GEOMETRY: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: пять исходных позиций и два связанных с Main ряда; формулы координат выполняет рабочий код.
Private Function OffsetGeometryDefinition(ByVal kind As Long, ByVal a2 As Double, ByVal a3 As Double, _
        Optional ByVal b2 As String = "EachBar", Optional ByVal b3 As String = "EachBar", Optional ByVal intruder As Long = 0) As CSectionDefinition
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Dim contact As CRebarContactBuilder, sign As Double
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "OffsetGeometryTest"
    If kind = 0 Then Set path = facade.RectanglePath("Boundary", -500#, -500#, 2000#, 2000#) Else Set path = facade.CirclePath("Boundary", 0#, 0#, 500#)
    facade.Outer "OUTER", path.ID
    If intruder = 1 Then AddOffsetIntruder facade, 100#
    If kind = 0 Then Set path = facade.LinePath("Axis", 0#, 0#, 400#, 0#) Else Set path = facade.CirclePath("Axis", 0#, 0#, 300#)
    If kind = 1 Then path.ValidateNormalTransfer 0#, 10#
    Set options = New CRebarRuleOptions: options.Coincidence = KeepPublishedOwner
    If kind = 1 Then options.EndMode = ClosedUniform
    facade.RebarByCount "Main", path.ID, 5, 32#, options
    Set options = New CRebarRuleOptions: options.SourceReference = "OffsetTest[2,3]": options.Coincidence = KeepPublishedOwner
    Set contact = New CRebarContactBuilder: contact.Initialize 32#, 1#, 1#, options
    Set options = contact.AppendRow(20#, "Stacked", b2, "OffsetTest[2,3]", a2)
    facade.RebarDependent "Row2", "Main", 20#, options
    Set options = contact.AppendRow(16#, "Stacked", b3, "OffsetTest[3,3]", a3)
    facade.RebarDependent "Row3", "Main", 16#, options
    If intruder = 2 Then AddOffsetIntruder facade, 100#
    Set OffsetGeometryDefinition = facade.Finish
End Function

' ДЛЯ ТЕСТОВ: независимый стержень на месте нового ряда; его нельзя скрыть слиянием.
Private Sub AddOffsetIntruder(ByVal facade As CSectionDefinitionBuilder, ByVal y As Double)
    Dim path As CGeometryPath, options As CRebarRuleOptions
    Set path = facade.LinePath("IntruderPath", 0#, y, 400#, y)
    Set options = New CRebarRuleOptions: options.Coincidence = KeepPublishedOwner
    facade.RebarByCount "Intruder", path.ID, 5, 20#, options
End Sub

' ДЛЯ ТЕСТОВ: два Main делят конец соседних линий; новое участие должно усилить проверки общего стержня.
Private Function OffsetSharedJointDefinition(ByVal reversed As Long, ByVal neighbor As Boolean) As CSectionDefinition
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "SharedOffsetJoint"
    Set path = facade.RectanglePath("Boundary", -500#, -500#, 1000#, 1000#): facade.Outer "OUTER", path.ID
    If neighbor And reversed = 1 Then AddOffsetJointNeighbor facade
    Set path = facade.LinePath("FirstPath", 0#, 0#, 400#, 0#)
    Set options = New CRebarRuleOptions: options.Coincidence = KeepPublishedOwner
    facade.RebarAtStations "FirstMain", path.ID, Array(0#), 32#, options
    options.NormalOffset = 100#: options.SourceReference = "LegacyJoint[2,3]"
    facade.RebarDependent "LegacyJoint", "FirstMain", 20#, options
    Set path = facade.LinePath("SecondPath", -400#, 0#, 0#, 0#)
    Set options = New CRebarRuleOptions: options.Coincidence = KeepPublishedOwner
    facade.RebarAtStations "SecondMain", path.ID, Array(400#), 32#, options
    options.NormalOffset = 100#: options.ValidateRowOffset = True: options.SourceReference = "OffsetJoint[2,3]"
    facade.RebarDependent "OffsetJoint", "SecondMain", 20#, options
    If neighbor And reversed = 0 Then AddOffsetJointNeighbor facade
    Set OffsetSharedJointDefinition = facade.Finish
End Function

' ДЛЯ ТЕСТОВ: сосед пересекает общий дополнительный стержень, но не основные ряды.
Private Sub AddOffsetJointNeighbor(ByVal facade As CSectionDefinitionBuilder)
    Dim path As CGeometryPath, options As CRebarRuleOptions
    Set path = facade.LinePath("NeighborPath", 5#, 100#, 100#, 100#)
    Set options = New CRebarRuleOptions: options.Coincidence = KeepPublishedOwner
    facade.RebarAtStations "Neighbor", path.ID, Array(0#), 8#, options
End Sub

' ДЛЯ ТЕСТОВ: ожидает понятную адресную ошибку общего исполнителя.
Private Function OffsetGeometryRejected(ByVal definition As CSectionDefinition, ByVal expected As String) As Boolean
    Dim layout As CRebarLayout
    On Error GoTo Rejected
    Set layout = FixtureLayout(definition)
    Exit Function
Rejected:
    OffsetGeometryRejected = InStr(1, Err.Description, expected, vbTextCompare) > 0 And _
        (InStr(1, Err.Description, "OffsetTest[", vbBinaryCompare) > 0 Or InStr(1, Err.Description, "OffsetJoint[", vbBinaryCompare) > 0)
End Function

' ДЛЯ ТЕСТОВ: настоящий ввод четырёх рецептов, единицы, пустые и формульные отступы.
' Каждая попытка строит арматуру общим исполнителем; исходный Config восстанавливается.
Public Function RunIndependentRowOffsetWorkbookTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, calculation As Long, fixture As Variant, item As Variant
    Dim unitName As Variant, offsets As Variant, settings As CSystemSettingsReader, units As CUnitSystem
    Dim definition As CSectionDefinition, bars As CRebarLayout, expected As Variant, actual As Variant
    Dim target As Object, extra As Long, main As Long, cell As Object, bad As Variant, i As Long, inputTable As CSectionInputTable
    Dim description As String, number As Long, prefix As String, rowCount As Long
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    calculation = Application.Calculation: Application.Calculation = -4135
    offsets = Array(Array(0#, 0#), Array(100#, 0#), Array(0#, 200#), Array(100#, 200#), Array(100#, 100#))
    For Each fixture In Array("Circle", "Rectangle", "RoundedMixed", "HollowOffset")
        For Each item In offsets
            expected = Empty
            For Each unitName In Array("mm", "cm", "m")
                ConfigureOffsetWorkbook CStr(fixture), CStr(unitName), CDbl(item(0)), CDbl(item(1))
                Set definition = ReadOffsetWorkbookDefinition()
                Set bars = FixtureLayout(definition): actual = OffsetBarsSnapshot(bars)
                prefix = "offset.workbook." & CStr(fixture) & "." & CStr(item(0)) & "." & CStr(item(1)) & "." & CStr(unitName)
                Check prefix & ".three.rows", bars.Count >= 15
                If IsEmpty(expected) Then expected = actual Else Check prefix & ".units", OffsetBarsEqual(expected, actual)
            Next unitName
        Next item
        ConfigureOffsetWorkbook CStr(fixture), "mm", 100#, 118#
        Set target = OffsetWorkbookTable(CStr(fixture))
        OffsetWorkbookRows target, CStr(fixture), main, extra
        Set cell = OffsetRebarCell(extra, 4)
        For Each bad In Array(-1#, "abc", True, CVErr(2015))
            cell.Value2 = bad: number = 0: description = ""
            On Error Resume Next
            Set definition = ReadOffsetWorkbookDefinition(): number = Err.Number: description = Err.Description
            Err.Clear: On Error GoTo Failed
            Check "offset.input.reject." & CStr(fixture) & "." & CStr(VarType(bad)), number <> 0
            Check "offset.input.address." & CStr(fixture) & "." & CStr(VarType(bad)), InStr(description, cell.Address(False, False)) > 0 And InStr(description, "Config") > 0
        Next bad
        cell.ClearContents: Set definition = ReadOffsetWorkbookDefinition(): Set bars = FixtureLayout(definition)
        actual = OffsetBarsSnapshot(bars)
        cell.Value2 = 0#: Set definition = ReadOffsetWorkbookDefinition(): Set bars = FixtureLayout(definition)
        Check "offset.input.blank.zero." & CStr(fixture), OffsetBarsEqual(actual, OffsetBarsSnapshot(bars))
        OffsetRebarCell(extra, 3).Value2 = 0#: cell.Value2 = "inactive"
        Set definition = ReadOffsetWorkbookDefinition(): Set bars = FixtureLayout(definition)
        Check "offset.input.disabled.second." & CStr(fixture), bars.Count >= 10
        OffsetRebarCell(extra, 7).Value2 = 0#: OffsetRebarCell(extra, 8).Value2 = CVErr(2015)
        Set definition = ReadOffsetWorkbookDefinition(): Set bars = FixtureLayout(definition)
        Check "offset.input.disabled.extras." & CStr(fixture), bars.Count >= 5
    Next fixture
    ConfigureOffsetWorkbook "Circle", "mm", 100#, 200#
    Set units = New CUnitSystem: units.InitializeDefaults
    Set target = OffsetWorkbookTable("Circle")
    Set inputTable = New CSectionInputTable: inputTable.Initialize "rngRebarInput", ThisWorkbook.Names("rngRebarInput").RefersToRange.Value2
    Check "offset.input.direct.reader", inputTable.OptionalRowOffset(FixtureInput(ThisWorkbook.Names("rngRebarInput").RefersToRange.Value2).RowIndex("Extra", "Circle.Circle"), FixtureInput(ThisWorkbook.Names("rngRebarInput").RefersToRange.Value2).ColumnIndex("Extra", "a2"), units) = 100#
    Set cell = FixtureCircleFieldCell(ThisWorkbook, "Rebar.AxisOffset2")
    cell.Formula = "=" & FixtureCircleFieldCell(ThisWorkbook, "Circle.Diameter").Address() & "/10"
    FixtureCircleFieldCell(ThisWorkbook, "Circle.Diameter").Value2 = 1200#
    Set definition = ReadOffsetWorkbookDefinition(): Set bars = FixtureLayout(definition)
    actual = OffsetBarsSnapshot(bars)
    cell.Value2 = 120#: Set definition = ReadOffsetWorkbookDefinition(): Set bars = FixtureLayout(definition)
    Check "offset.input.formula.manual.recalculated", OffsetBarsEqual(actual, OffsetBarsSnapshot(bars))
    GoTo Finished
Failed:
    Check "offset.workbook.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation
    RunIndependentRowOffsetWorkbookTests = mReport & "TOTAL INDEPENDENT ROW OFFSET WORKBOOK: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: задаёт достаточно просторные сечения; отступы и диаметры даны в мм.
' Открытый ряд берём по верхней стороне; в пустотелом сечении добавляем ряд у отверстия.
Private Sub ConfigureOffsetWorkbook(ByVal fixture As String, ByVal unitName As String, ByVal a2 As Double, ByVal a3 As Double)
    Dim unitTable As Object, target As Object, factor As Double, row As Long, key As Variant
    Dim keyed As CSectionInputReader, active As Boolean
    Set unitTable = ThisWorkbook.Names("rngUnitSettings").RefersToRange
    For row = 1 To unitTable.Rows.Count
        If CStr(unitTable.Cells(row, 1).Value2) = "Length" Then unitTable.Cells(row, 2).Value2 = unitName
    Next row
    factor = 1#: If unitName = "cm" Then factor = 10#
    If unitName = "m" Then factor = 1000#
    modTestShapeArchitecture.ConfigureShapeFixture fixture
    Set target = OffsetWorkbookTable(fixture)
    Select Case fixture
        Case "Circle"
            FixtureInputCell(target, "General", "D").Value2 = 1000# / factor
            FixtureInputCell(target, "Main", "Circle", "as").Value2 = 60# / factor
            FixtureInputCell(target, "Main", "Circle", "n").Value2 = 12
            FixtureInputCell(target, "Main", "Circle", "d").Value2 = 32# / factor
        Case "Rectangle"
            FixtureInputCell(target, "General", "H1").Value2 = 1200# / factor: FixtureInputCell(target, "General", "B1").Value2 = 1200# / factor
        Case "RoundedMixed"
            FixtureInputCell(target, "General", "B").Value2 = 1600# / factor: FixtureInputCell(target, "General", "H").Value2 = 1600# / factor
            For Each key In Array("left", "right")
                FixtureInputCell(target, "Detail", "W", CStr(key)).Value2 = 400# / factor
                FixtureInputCell(target, "Detail", "R1", CStr(key)).Value2 = 300# / factor
                FixtureInputCell(target, "Detail", "R2", CStr(key)).Value2 = 300# / factor
            Next key
        Case Else
            FixtureInputCell(target, "General", "H").Value2 = 2400# / factor: FixtureInputCell(target, "General", "B").Value2 = 2400# / factor
            FixtureInputCell(target, "Detail", "Ho").Value2 = 800# / factor: FixtureInputCell(target, "Detail", "Bo").Value2 = 800# / factor
    End Select
    Select Case fixture
        Case "Circle": Set keyed = FixtureInput(FixtureRecipeValues("rngCircleGeometry"))
        Case "Rectangle": Set keyed = FixtureInput(FixtureRecipeValues("rngRectSetGeometry"))
        Case "RoundedMixed": Set keyed = FixtureInput(FixtureRecipeValues("rngRoundedRectangleGeometry"))
        Case Else: Set keyed = FixtureInput(FixtureRecipeValues("rngHollowRectangleGeometry"))
    End Select
    For Each key In keyed.RowKeys("Main")
        active = (CStr(key) = "Circle" Or CStr(key) = "Top" Or CStr(key) = "R1.Top" Or CStr(key) = "OpeningTop")
        If fixture <> "Circle" Then
            FixtureInputCell(target, "Main", CStr(key), "as").Value2 = 60# / factor
            FixtureInputCell(target, "Main", CStr(key), "d").Value2 = IIf(active, 32# / factor, 0#)
            If Left$(CStr(key), 7) <> "Opening" Then FixtureInputCell(target, "Main", CStr(key), "n").Value2 = IIf(active, 5, 0)
            If fixture = "Rectangle" Then
                FixtureInputCell(target, "Main", CStr(key), "s0").Value2 = 150# / factor
                FixtureInputCell(target, "Main", CStr(key), "s1").Value2 = 150# / factor
            End If
        End If
        FixtureInputCell(target, "Extra", CStr(key), "d2").Value2 = IIf(active, 20# / factor, 0#)
        FixtureInputCell(target, "Extra", CStr(key), "d3").Value2 = IIf(active, 16# / factor, 0#)
        FixtureInputCell(target, "Extra", CStr(key), "a2").Value2 = a2 / factor
        FixtureInputCell(target, "Extra", CStr(key), "a3").Value2 = a3 / factor
        FixtureInputCell(target, "Extra", CStr(key), "bind2").Value2 = "EachBar"
        FixtureInputCell(target, "Extra", CStr(key), "bind3").Value2 = "EachBar"
        FixtureInputCell(target, "Extra", CStr(key), "pos2").Value2 = "Stacked"
        FixtureInputCell(target, "Extra", CStr(key), "pos3").Value2 = "Stacked"
    Next key
End Sub

' ДЛЯ ТЕСТОВ: получает действующую таблицу выбранного рецепта по штатному имени.
Private Function OffsetWorkbookTable(ByVal fixture As String) As Object
    Dim name As String
    Select Case fixture
        Case "Circle": name = "rngCircleGeometry"
        Case "Rectangle": name = "rngRectSetGeometry"
        Case "RoundedMixed": name = "rngRoundedRectangleGeometry"
        Case Else: name = "rngHollowRectangleGeometry"
    End Select
    Set OffsetWorkbookTable = ThisWorkbook.Names(name).RefersToRange
End Function

' ДЛЯ ТЕСТОВ: строки основного и дополнительного армирования тестируемой верхней стороны.
Private Sub OffsetWorkbookRows(ByVal target As Object, ByVal fixture As String, ByRef main As Long, ByRef extra As Long)
    Dim keyed As CSectionInputReader, key As String, typeID As String
    Select Case fixture
        Case "Circle": typeID = "Circle": key = "Circle"
        Case "Rectangle": typeID = "RectSet": key = "R1.Top"
        Case "RoundedMixed": typeID = "RoundedRectangle": key = "Top"
        Case Else: typeID = "HollowRectangle": key = "Top"
    End Select
    FixtureSelectRecipeType ThisWorkbook, typeID
    ThisWorkbook.Names("rngRebarInput").RefersToRange.Calculate
    Set keyed = FixtureInput(ThisWorkbook.Names("rngRebarInput").RefersToRange.Value2)
    key = typeID & "." & key
    main = keyed.RowIndex("Main", key): extra = keyed.RowIndex("Extra", key)
End Sub

' ДЛЯ ТЕСТОВ: чтение тем же снимком входа, которым пользуется рабочий сборщик.
Private Function ReadOffsetWorkbookDefinition() As CSectionDefinition
    Dim settings As CSystemSettingsReader, units As CUnitSystem
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set ReadOffsetWorkbookDefinition = FixtureWorkbookDefinition(settings.GetString("Geometry.Type"), ThisWorkbook, units)
End Function

' ДЛЯ ТЕСТОВ: независимая численная запись координат и диаметров в порядке публикации.
Private Function OffsetBarsSnapshot(ByVal bars As CRebarLayout) As Variant
    Dim values() As Double, i As Long
    ReDim values(1 To bars.Count, 1 To 3)
    For i = 1 To bars.Count
        values(i, 1) = bars.X(i): values(i, 2) = bars.Y(i): values(i, 3) = bars.Diameter(i)
    Next i
    OffsetBarsSnapshot = values
End Function

' ДЛЯ ТЕСТОВ: допускает только погрешность пересчёта единиц, без округления входа.
Private Function OffsetBarsEqual(ByVal first As Variant, ByVal second As Variant) As Boolean
    Dim i As Long, j As Long
    If UBound(first, 1) <> UBound(second, 1) Then Exit Function
    For i = 1 To UBound(first, 1)
        For j = 1 To 3
            If Abs(first(i, j) - second(i, j)) > 0.00000001 Then Exit Function
        Next j
    Next i
    OffsetBarsEqual = True
End Function

' ДЛЯ ТЕСТОВ: настоящие схемы Results и расчётные снимки новых рядов, без изменения рабочей книги.
Public Function DumpIndependentRowOffsetExamples(ByVal outputDirectory As String) As String
    Dim config As Object, saved As Variant, result As String, fixtures As Variant, fixture As String
    Dim titles As Variant, a2 As Double, a3 As Double, index As Long, target As Object, main As Long, extra As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim calculation As Long, events As Boolean, updating As Boolean, plot As String
    On Error GoTo Failed
    calculation = Application.Calculation: events = Application.EnableEvents: updating = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = True
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    fixtures = Array("Circle", "Circle", "Rectangle", "RoundedMixed", "HollowOffset", "Rectangle", "Circle")
    titles = Array("Circle_TwoRows_100", "Circle_ThreeRows_100_100", "Rectangle_ThreeRows_100_200", _
        "Rounded_Sloped_60_100", "Hollow_Opening_100_100_SideBySide", "Rectangle_100_100_EverySecond", "Circle_Arc_ThreeRows_100_200")
    For index = 0 To UBound(fixtures)
        fixture = CStr(fixtures(index)): a2 = 100#: a3 = 100#
        If index = 2 Or index = 6 Then a3 = 200#
        If index = 3 Then a2 = 60#: a3 = 100#
        ConfigureOffsetWorkbook fixture, "mm", a2, a3
        Set target = OffsetWorkbookTable(fixture): OffsetWorkbookRows target, fixture, main, extra
        If index = 0 Then OffsetRebarCell(extra, 7).Value2 = 0#
        If index = 3 Then
            OffsetRebarCell(main, 4).Value2 = 0#: OffsetRebarCell(main, 5).Value2 = 0
            OffsetRebarCell(extra, 3).Value2 = 0#: OffsetRebarCell(extra, 7).Value2 = 0#
            FixtureInputCell(target, "Main", "Left", "d").Value2 = 32#: FixtureInputCell(target, "Main", "Left", "n").Value2 = 5
            FixtureInputCell(target, "Extra", "Left", "d2").Value2 = 20#: FixtureInputCell(target, "Extra", "Left", "d3").Value2 = 16#
        End If
        If index = 6 Then
            FixtureCircleFieldCell(ThisWorkbook, "Rebar.StartSetback").Value2 = 150#
            FixtureCircleFieldCell(ThisWorkbook, "Rebar.EndSetback").Value2 = 250#
        End If
        If index = 4 Or index = 5 Then
            OffsetRebarCell(extra, 10).Value2 = "SideBySide": OffsetRebarCell(extra, 9).Value2 = "EverySecondBar"
            If index = 4 Then
                OffsetRebarCell(extra + 4, 10).Value2 = "SideBySide": OffsetRebarCell(extra + 4, 9).Value2 = "EverySecondBar"
            End If
        End If
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
        result = result & "EXAMPLE: " & CStr(titles(index)) & "; bars=" & CStr(builder.Rebars.Count) & vbCrLf
        result = result & modTestShapeArchitecture.PreparedShapeSnapshot(CStr(titles(index)), model, builder.Mesh, builder.Rebars)
        If index = 1 Or index = 2 Or index = 4 Or index = 6 Then result = result & modTestShapeArchitecture.PreparedCalculationSnapshot(CStr(titles(index)), model)
        plot = modTestShapeArchitecture.PreparedPreviewSnapshot(CStr(titles(index)), model, outputDirectory)
        result = result & plot
        If index = 4 Then ExportOffsetPreviewLarge outputDirectory & Application.PathSeparator & "Hollow_Detail_4x.png"
    Next index
    GoTo Finished
Failed:
    result = result & "FAIL: offset.examples; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = updating
    DumpIndependentRowOffsetExamples = result
End Function

' ДЛЯ ТЕСТОВ: Excel повторно выводит векторный снимок готовой диаграммы в крупный PNG.
' Координаты стержней и фигуры исходной диаграммы не меняются.
Private Sub ExportOffsetPreviewLarge(ByVal path As String)
    Dim container As Object, scratch As Object, canvas As Object, picture As Object
    Dim width As Double, height As Double
    Set container = ThisWorkbook.Worksheets("Расчет").ChartObjects("chtNDMSectionPlot")
    width = container.Width * 4#: height = container.Height * 4#
    container.Activate: container.Chart.CopyPicture 1, -4147
    Set scratch = ThisWorkbook.Worksheets.Add
    On Error GoTo Failed
    Set canvas = scratch.ChartObjects.Add(0#, 0#, width, height)
    canvas.Activate: canvas.Chart.Paste
    Set picture = canvas.Chart.Shapes(1)
    picture.LockAspectRatio = 0: picture.Left = 0#: picture.Top = 0#: picture.Width = width: picture.Height = height
    If Not canvas.Chart.Export(path, "PNG") Then Err.Raise vbObjectError + 5294, "modTestAudit02Api", "Не сохранён крупный снимок схемы."
    Application.DisplayAlerts = False: scratch.Delete
    Exit Sub
Failed:
    Dim number As Long, reason As String
    number = Err.Number: reason = Err.Description
    Application.DisplayAlerts = False: scratch.Delete
    Err.Raise number, "modTestAudit02Api", reason
End Sub

' ДЛЯ ТЕСТОВ: смещение по касательной скругления и центр подписи асимметричного ряда у отверстия.
Public Function RunOffsetOpeningAnnotationTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, target As Object, main As Long, extra As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, builder As CSectionBuilder, model As CSectionModel
    Dim i As Long, source As Long, second As Long, third As Long, minX As Double, maxX As Double, maxY As Double
    Dim dx As Double, dy As Double, nx As Double, ny As Double, found As Boolean, annotations As CSectionAnnotations
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    ConfigureOffsetWorkbook "HollowOffset", "mm", 100#, 100#
    Set target = OffsetWorkbookTable("HollowOffset"): OffsetWorkbookRows target, "HollowOffset", main, extra
    OffsetRebarCell(extra + 4, 10).Value2 = "SideBySide": OffsetRebarCell(extra + 4, 9).Value2 = "EverySecondBar"
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set builder = New CSectionBuilder: Set model = BuildWorkbookSectionModel(ThisWorkbook, settings, units, builder)
    minX = 1E+100: maxX = -1E+100: maxY = -1E+100
    For i = 1 To model.RebarCount
        Select Case model.RebarSourceName(i)
            Case "Opening.B.Top.Main.2": source = i
            Case "Opening.B.Top.Row2.1": second = i
            Case "Opening.B.Top.Row3.1": third = i
        End Select
        If Left$(model.RebarSourceName(i), 19) = "Opening.B.Top.Main." Then
            minX = GeomMin(minX, model.RebarX(i)): maxX = GeomMax(maxX, model.RebarX(i))
            maxY = GeomMax(maxY, model.RebarY(i))
        End If
    Next i
    Check "opening.offset.source.positions", source > 0 And second > 0 And third > 0
    If source = 0 Or second = 0 Or third = 0 Then GoTo Finished
    nx = model.RebarX(second) - model.RebarX(source): ny = model.RebarY(second) - model.RebarY(source)
    dx = model.RebarX(third) - model.RebarX(second): dy = model.RebarY(third) - model.RebarY(second)
    Check "opening.offset.normal100", Abs(Sqr(nx * nx + ny * ny) - 100#) < 0.0000001
    Check "opening.offset.profile.gap2.5", Abs(Sqr(dx * dx + dy * dy) - 20.5) < 0.0000001
    Check "opening.offset.tangent.perpendicular", Abs(nx * dx + ny * dy) < 0.0000001
    Check "opening.offset.corner.components", Abs(dx - 17.0453280954174) < 0.0000001 And Abs(dy - 11.3888888888889) < 0.0000001
    Set annotations = model.Annotations
    For i = 1 To annotations.Count
        If annotations.AnnotationID(i) = "REBAR_Opening.B.Top" Then
            found = True
            Check "opening.annotation.full.row.start", Abs(annotations.StartX(i) - minX) < 0.0000001
            Check "opening.annotation.full.row.end", Abs(annotations.EndX(i) - maxX) < 0.0000001
            Check "opening.annotation.full.row.center", Abs((annotations.StartX(i) + annotations.EndX(i)) / 2#) < 0.0000001
            Check "opening.annotation.extreme.support", Abs(annotations.StartY(i) - maxY) < 0.0000001 And Abs(annotations.EndY(i) - maxY) < 0.0000001
        End If
    Next i
    Check "opening.annotation.found", found
    GoTo Finished
Failed:
    Check "opening.annotation.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    RunOffsetOpeningAnnotationTests = mReport & "TOTAL OFFSET OPENING ANNOTATION: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: полный круг и открытая дуга, полученная концевыми отступами;
' координаты проверяются по независимо заданным точкам полуокружности.
Public Function RunCircleSetbackTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, unitName As Variant, factor As Double, table As Object
    Dim bars As CRebarLayout, original As Variant, point As Variant, expected As Variant, i As Long, column As Variant
    Dim definition As CSectionDefinition, bad As Variant, number As Long, reason As String
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, x As Double, y As Double, tx As Double, ty As Double
    Dim executor As CRebarRuleExecutor, location As Variant, binding As Variant, secondCount As Long, thirdCount As Long, owner As Long
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "SliceTest"
    Set path = facade.CirclePath("Full", 0#, 0#, 100#)
    path.AtDistance 0#, x, y, tx, ty
    Check "circle.setback.generic.start.unchanged", Abs(x - 100#) < 0.0000001 And Abs(y) < 0.0000001 And ty > 0#
    Set path = facade.Slice("Half", path.ID, 50# * GEOM_PI, 150# * GEOM_PI)
    Check "circle.setback.shared.slice", Not path.IsClosed And Abs(path.Length - 100# * GEOM_PI) < 0.0000001
    expected = Array(Array(100#, 0#), Array(50#, -86.6025403784439), Array(-50#, -86.6025403784439), Array(-100#, 0#))
    For Each unitName In Array("mm", "cm", "m")
        ConfigureOffsetWorkbook "Circle", CStr(unitName), 0#, 0#
        factor = 1#: If CStr(unitName) = "cm" Then factor = 10#
        If CStr(unitName) = "m" Then factor = 1000#
        FixtureCircleFieldCell(ThisWorkbook, "Circle.Diameter").Value2 = 300# / factor
        FixtureCircleFieldCell(ThisWorkbook, "Rebar.AxisDistance").Value2 = 50# / factor
        FixtureCircleFieldCell(ThisWorkbook, "Rebar.Count").Value2 = 4
        FixtureCircleFieldCell(ThisWorkbook, "Rebar.Diameter").Value2 = 16# / factor
        FixtureCircleFieldCell(ThisWorkbook, "Rebar.Diameter2").Value2 = 0#
        FixtureCircleFieldCell(ThisWorkbook, "Rebar.Diameter3").Value2 = 0#
        FixtureCircleFieldCell(ThisWorkbook, "Rebar.StartSetback").Value2 = 0#
        FixtureCircleFieldCell(ThisWorkbook, "Rebar.EndSetback").Value2 = 0#
        Set definition = ReadOffsetWorkbookDefinition(): Set bars = FixtureLayout(definition)
        original = OffsetBarsSnapshot(bars)
        Check "circle.setback.full.closed", definition.Path("Circle.Axis").IsClosed And bars.Count = 4
        Check "circle.setback.full.first.top", Abs(bars.X(1)) < 0.0000001 And Abs(bars.Y(1) - 100#) < 0.0000001
        Check "circle.setback.full.clockwise", Abs(bars.X(2) - 100#) < 0.0000001 And Abs(bars.Y(2)) < 0.0000001
        Check "circle.setback.full.bottom.left", Abs(bars.X(3)) < 0.0000001 And Abs(bars.Y(3) + 100#) < 0.0000001 And Abs(bars.X(4) + 100#) < 0.0000001 And Abs(bars.Y(4)) < 0.0000001
        For Each column In Array("Rebar.StartSetback", "Rebar.EndSetback")
            FixtureCircleFieldCell(ThisWorkbook, CStr(column)).Value2 = 50# * GEOM_PI / factor
        Next column
        Set definition = ReadOffsetWorkbookDefinition(): Set bars = FixtureLayout(definition)
        Check "circle.setback.arc.open", Not definition.Path("Circle.Axis").IsClosed And bars.Count = 4
        For i = 1 To 4
            point = expected(i - 1)
            Check "circle.setback.arc.point." & CStr(unitName) & "." & CStr(i), Abs(bars.X(i) - CDbl(point(0))) < 0.0000001 And Abs(bars.Y(i) - CDbl(point(1))) < 0.0000001 And bars.Diameter(i) = 16#
        Next i
        FixtureCircleFieldCell(ThisWorkbook, "Rebar.EndSetback").Value2 = 0#
        Set definition = ReadOffsetWorkbookDefinition(): Set bars = FixtureLayout(definition)
        Check "circle.setback.start.only", Abs(bars.X(1) - 100#) < 0.0000001 And Abs(bars.Y(1)) < 0.0000001 And Abs(bars.Y(4) - 100#) < 0.0000001
        FixtureCircleFieldCell(ThisWorkbook, "Rebar.StartSetback").Value2 = 0#
        FixtureCircleFieldCell(ThisWorkbook, "Rebar.EndSetback").Value2 = 50# * GEOM_PI / factor
        Set definition = ReadOffsetWorkbookDefinition(): Set bars = FixtureLayout(definition)
        Check "circle.setback.end.only", Abs(bars.X(1)) < 0.0000001 And Abs(bars.Y(1) - 100#) < 0.0000001 And Abs(bars.X(4) + 100#) < 0.0000001
        FixtureCircleFieldCell(ThisWorkbook, "Rebar.StartSetback").Value2 = 0#
        FixtureCircleFieldCell(ThisWorkbook, "Rebar.EndSetback").Value2 = 0#
        Set definition = ReadOffsetWorkbookDefinition(): Set bars = FixtureLayout(definition)
        Check "circle.setback.return.zero", OffsetBarsEqual(original, OffsetBarsSnapshot(bars))
        For Each column In Array("Rebar.StartSetback", "Rebar.EndSetback")
            Set table = FixtureCircleFieldCell(ThisWorkbook, CStr(column))
            For Each bad In Array(-1#, "bad", True, CVErr(2015), 200# * GEOM_PI / factor)
                table.Value2 = bad: number = 0: reason = ""
                On Error Resume Next
                Set definition = ReadOffsetWorkbookDefinition()
                number = Err.Number: reason = Err.Description: Err.Clear: On Error GoTo Failed
                Check "circle.setback.invalid." & CStr(column), number <> 0 And InStr(reason, "Config!") > 0
            Next bad
            table.Value2 = 0#
        Next column
        FixtureCircleFieldCell(ThisWorkbook, "Rebar.Diameter2").Value2 = 12# / factor
        FixtureCircleFieldCell(ThisWorkbook, "Rebar.Diameter3").Value2 = 10# / factor
        FixtureCircleFieldCell(ThisWorkbook, "Rebar.AxisOffset2").Value2 = 20# / factor
        FixtureCircleFieldCell(ThisWorkbook, "Rebar.AxisOffset3").Value2 = 20# / factor
        For Each location In Array("Stacked", "SideBySide")
            FixtureCircleFieldCell(ThisWorkbook, "Rebar.Loc3row").Value2 = location
            For Each binding In Array("EachBar", "EverySecondBar")
                FixtureCircleFieldCell(ThisWorkbook, "Rebar.Bind3row").Value2 = binding
                Set definition = ReadOffsetWorkbookDefinition(): Set executor = New CRebarRuleExecutor
                Set bars = executor.Build(definition, FixtureGeometry(definition.ConcreteDomain))
                secondCount = 0: thirdCount = 0
                For i = 1 To executor.PositionCount
                    owner = executor.PositionPhysicalOwner(i)
                    If executor.PositionRuleID(i) = "Circle.Row2" Then
                        secondCount = secondCount + 1
                        Check "circle.clockwise.row2.inward", Abs(Sqr(bars.X(owner) ^ 2 + bars.Y(owner) ^ 2) - 80#) < 0.0000001
                    ElseIf executor.PositionRuleID(i) = "Circle.Row3" Then
                        thirdCount = thirdCount + 1
                        If thirdCount = 1 Then
                            Check "circle.clockwise.row3.top", Abs(bars.X(owner) - IIf(location = "Stacked", 0#, 13.5)) < 0.0000001 And Abs(bars.Y(owner) - IIf(location = "Stacked", 66.5, 80#)) < 0.0000001
                        ElseIf thirdCount = 2 And binding = "EverySecondBar" Then
                            Check "circle.clockwise.everySecond.bottom", Abs(bars.X(owner) + IIf(location = "Stacked", 0#, 13.5)) < 0.0000001 And Abs(bars.Y(owner) + IIf(location = "Stacked", 66.5, 80#)) < 0.0000001
                        End If
                    End If
                Next i
                Check "circle.clockwise.row.counts", secondCount = 4 And thirdCount = IIf(binding = "EachBar", 4, 2)
            Next binding
        Next location
    Next unitName
    GoTo Finished
Failed:
    Check "circle.setback.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    RunCircleSetbackTests = mReport & "TOTAL CIRCLE SETBACKS: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: встреча рядов наружной грани и отверстия в тонкой стенке;
' отдельно проверяет диск, центр которого в бетоне, а край заходит в отверстие.
Public Function RunOffsetWallCollisionTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    Dim config As Object, saved As Variant, target As Object, main As Long, extra As Long
    Dim definition As CSectionDefinition, bars As CRebarLayout, scenario As Long, number As Long, reason As String
    On Error GoTo Failed
    Set config = ThisWorkbook.Worksheets("Config").UsedRange: saved = config.Formula
    For scenario = 0 To 2
        ConfigureOffsetWorkbook "HollowOffset", "mm", 100#, 100#
        Set target = OffsetWorkbookTable("HollowOffset"): OffsetWorkbookRows target, "HollowOffset", main, extra
        FixtureInputCell(target, "General", "H", "val").Value2 = 1000#: FixtureInputCell(target, "Detail", "Ho", "val").Value2 = 324#: FixtureInputCell(target, "Detail", "Ro", "val").Value2 = 0#
        FixtureInputCell(target, "Detail", "Xo", "val").Value2 = 0#: FixtureInputCell(target, "Detail", "Yo", "val").Value2 = 0#
        OffsetRebarCell(extra, 7).Value2 = 0#: OffsetRebarCell(extra + 4, 7).Value2 = 0#
        OffsetRebarCell(extra + 4, 3).Value2 = 16#
        If scenario = 1 Then FixtureInputCell(target, "General", "H", "val").Value2 = 992#
        If scenario = 2 Then
            OffsetRebarCell(extra, 3).Value2 = 0#
            OffsetRebarCell(extra + 4, 3).Value2 = 200#: OffsetRebarCell(extra + 4, 4).Value2 = 1#
        End If
        Set bars = Nothing: number = 0: reason = ""
        On Error Resume Next
        Set definition = ReadOffsetWorkbookDefinition(): Set bars = FixtureLayout(definition)
        number = Err.Number: reason = Err.Description: Err.Clear: On Error GoTo Failed
        mReport = mReport & "OBSERVED wall." & CStr(scenario) & ": " & CStr(number) & "; " & reason & vbCrLf
        Select Case scenario
            Case 0
                Check "offset.wall.touch.allowed", number = 0
                If Not bars Is Nothing Then Check "offset.wall.touch.both.families", bars.Count = 16
            Case 1
                Check "offset.wall.overlap.rejected", number <> 0 And InStr(reason, "Диски стержней") > 0 And InStr(reason, "Config!") > 0
            Case 2
                Check "offset.wall.disk.opening.rejected", number <> 0 And InStr(reason, "целиком") > 0 And InStr(reason, "Config!") > 0
        End Select
    Next scenario
    GoTo Finished
Failed:
    Check "offset.wall.runtime; " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    If Not config Is Nothing And Not IsEmpty(saved) Then config.Formula = saved
    RunOffsetWallCollisionTests = mReport & "TOTAL OFFSET WALL COLLISION: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function

' ДЛЯ ТЕСТОВ: физическая ячейка дополнительного ряда в общей таблице армирования.
Private Function OffsetRebarCell(ByVal row As Long, ByVal column As Long) As Object
    Set OffsetRebarCell = ThisWorkbook.Names("rngRebarInput").RefersToRange.Cells(row, column)
End Function
