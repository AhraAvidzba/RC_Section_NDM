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
    options.AnnotationGroup = "Group" ' В одной подписи группы суммируем стержни каждого диаметра, включая все её ряды.
    options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
    options.Coincidence = KeepPublishedOwner ' При совпадении сохраняем первый опубликованный стержень.
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
    Set path = definition.Path("Shape"): path.CopySegment path.SegmentIndex("Edge.TR"), edge
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
    Dim shapePath As CGeometryPath, boundary As CGeometryPath, offsetPath As CGeometryPath, rightAxis As CGeometryPath
    Dim options As CRebarRuleOptions, b As Double, h As Double, w As Double
    b = 600#: h = 1000#: w = 200#
    Set polygon = New CPolygonSpec
    polygon.AddVertex "TL", -b / 2#, h / 2#, 80#
    polygon.AddVertex "TR", b / 2#, h / 2#, 120#
    polygon.AddVertex "RightNose", b / 2# + w, 0#, 80#
    polygon.AddVertex "BR", b / 2#, -h / 2#, 120#
    polygon.AddVertex "BL", -b / 2#, -h / 2#, 80#
    polygon.Seal
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Audit02Demo"
    Set shapePath = facade.RoundedPolygon("Shape", polygon)
    Set boundary = facade.Chain("Boundary", shapePath.ID, "Edge.BL", 0#, "Edge.BL", 0#, True, True)
    facade.Outer "OUTER", boundary.ID
    Set offsetPath = facade.OffsetRoundedPolygon("Offset.H.Right", polygon, 50#)
    Set rightAxis = facade.Chain("H.Right", offsetPath.ID, "Edge.TL", 1#, "Edge.BR", 0#)
    Set options = New CRebarRuleOptions
    options.AnnotationGroup = "H.Right" ' В одной подписи группы суммируем стержни каждого диаметра, включая все её ряды.
    options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
    options.Coincidence = KeepPublishedOwner ' При совпадении сохраняем первый опубликованный стержень.
    options.EndMode = ExcludeBoth: options.StartSetback = 100#: options.EndSetback = 150#
    options.IncomingTangentAtJoint = True
    facade.RebarByCount "H.Right.Main", rightAxis.ID, 15, 20#, options
    facade.RebarRowExtentLabel "H.Right", "H.Right.Main", 1#, 0#, True
    facade.PathAnchor "Slope.A", shapePath.ID, "Edge.TR", 0#
    facade.PathAnchor "Slope.B", shapePath.ID, "Edge.TR", 1#
    facade.Dimension "DIM_SLOPE", "Slope.A", "Slope.B", "L", "Aligned"
    facade.HorizontalDistance "DIM_B", "B", -b / 2#, b / 2#, h / 2#, 1#
    facade.VerticalExtent "DIM_H", boundary.ID, "H", -b / 2#, -1#
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
    options.Coincidence = KeepPublishedOwner ' При совпадении сохраняем первый опубликованный стержень.
    facade.RebarByCount "Main", "MainAxis", 5, 20#, options
    Set options = New CRebarRuleOptions
    options.AnnotationGroup = "Bottom" ' В одной подписи группы суммируем стержни каждого диаметра, включая все её ряды.
    options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
    options.Coincidence = KeepPublishedOwner ' При совпадении сохраняем первый опубликованный стержень.
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

' ДЛЯ ТЕСТОВ: нулевые диаметры запрещены внутри контактного helper, отказ не меняет options.
Public Function RunAudit02CloseoutContactTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim options As CRebarRuleOptions, cases As Variant, item As Variant, code As Long, reason As String
    Set options = New CRebarRuleOptions
    options.SourceReference = "Program.Contact": options.StartSetback = 100#: options.EndSetback = 150#
    options.AnnotationGroup = "Group" ' В одной подписи группы суммируем стержни каждого диаметра, включая все её ряды.
    options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
    options.Coincidence = KeepPublishedOwner ' При совпадении сохраняем первый опубликованный стержень.
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

' Снимок ошибок содержит номер, source и полный текст с реальным адресом, без маскирования различий.
Private Function CloseoutRectSetInputMatrix() As String
    Dim config As Object, saved As Variant, range As Object, block As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, definition As CSectionDefinition
    Dim reader As CSectionInputSnapshotReader, context As CSectionRecipeContext, table As CSectionInputTable
    Dim fixture As Variant, valueKind As Variant, unit As Variant, aliasName As Variant
    Dim row As Long, column As Long, maxColumn As Long, offsetRow As Long, shapeRow As Long
    Dim code As Long, source As String, reason As String, expectedAddress As String, area As Double
    Dim result As String, factor As Double, pair As Variant, number As Double, otherCode As Long, otherReason As String
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
            Set context = reader.Capture(ThisWorkbook, Array("rngRectSetGeometry"), units): Set table = context.Table("rngRectSetGeometry")
            row = table.RowByKey("H1") + 1: offsetRow = table.RowByKey("RectSet.UpperOffsetX")
            shapeRow = table.RowByKey("RectSet.SectionType")
            Set definition = RectSetInputOutcome(settings, units, code, source, reason)
            Check "rect.valid." & CStr(unit) & "." & CStr(fixture), code = 0
            If code <> 0 Then GoTo Failed
            area = definition.ConcreteDomain.Area
            result = result & InputContractRow(CStr(unit), CStr(fixture), "base", code, source, reason, area)
            maxColumn = 4: If CStr(fixture) = "Rectangle" Then maxColumn = 2
            For column = 1 To maxColumn
                expectedAddress = table.SourceReference(row, column)
                For Each valueKind In Array("positive", "zero", "negative", "empty", "text")
                    range.Formula = block
                    Select Case CStr(valueKind)
                        Case "zero": range.Cells(row, column).Value2 = 0#
                        Case "negative": range.Cells(row, column).Value2 = -1# / factor
                        Case "empty": range.Cells(row, column).ClearContents
                        Case "text": range.Cells(row, column).Value2 = "bad dimension"
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
                range.Cells(row, 3).Value2 = "inactive H2": range.Cells(row, 4).Value2 = CVErr(xlErrValue)
            End If
            If CStr(fixture) <> "TwoRectangles" Then
                range.Cells(offsetRow, 2).Value2 = "inactive offset"
                Set definition = RectSetInputOutcome(settings, units, code, source, reason)
                Check "rect.inactive.fields." & CStr(unit) & "." & CStr(fixture), code = 0 And definition.ConcreteDomain.Area = area
                result = result & InputContractRow(CStr(unit), CStr(fixture), "inactive", code, source, reason, 0#)
            Else
                For Each pair In Array(-100#, 0#, 100#)
                    range.Formula = block: range.Cells(offsetRow, 2).Value2 = CDbl(pair) / factor
                    Set definition = RectSetInputOutcome(settings, units, code, source, reason)
                    Check "rect.signed.offset." & CStr(unit) & "." & CStr(pair), code = 0
                    result = result & InputContractRow(CStr(unit), CStr(fixture), "offset." & CStr(pair), code, source, reason, 0#)
                Next pair
            End If
            ' Контракт пары: числовое чтение второго поля предшествует проверке положительности первого.
            For Each pair In Array(1, 3)
                If CLng(pair) <= maxColumn Then
                    range.Formula = block: range.Cells(row, CLng(pair)).Value2 = 0#
                    range.Cells(row, CLng(pair) + 1).Value2 = "bad dimension"
                    Set definition = RectSetInputOutcome(settings, units, code, source, reason)
                    Check "rect.pair.numeric.before.positive." & CStr(unit) & "." & CStr(fixture) & "." & CStr(pair), _
                        code <> 0 And InStr(reason, table.SourceReference(row, CLng(pair) + 1)) > 0
                    result = result & InputContractRow(CStr(unit), CStr(fixture), "pair." & CStr(pair), code, source, reason, 0#)
                    Set context = reader.Capture(ThisWorkbook, Array("rngRectSetGeometry"), units): Set table = context.Table("rngRectSetGeometry")
                    On Error Resume Next
                    number = table.PositiveLength(row, CLng(pair), units)
                    otherCode = Err.Number: otherReason = Err.Description: Err.Clear
                    On Error GoTo Failed
                    Check "rect.direct.helper.changes.first.address", otherCode <> 0 And _
                        InStr(otherReason, table.SourceReference(row, CLng(pair))) > 0 And otherReason <> reason
                End If
            Next pair
            range.Formula = block
            ' Составные условия остаются у рецепта и указывают соответствующее поле.
            If CStr(fixture) = "LSection" Then
                range.Cells(row, 2).Value2 = range.Cells(row, 4).Value2
                Set definition = RectSetInputOutcome(settings, units, code, source, reason)
                Check "rect.LSection.B1.less.B2", code <> 0 And InStr(reason, table.SourceReference(row, 2)) > 0
                result = result & InputContractRow(CStr(unit), CStr(fixture), "incompatible.widths", code, source, reason, 0#)
            ElseIf CStr(fixture) = "TwoRectangles" Then
                range.Cells(offsetRow, 2).Value2 = range.Cells(row, 4).Value2
                Set definition = RectSetInputOutcome(settings, units, code, source, reason)
                Check "rect.adjacency.address", code <> 0 And InStr(reason, table.SourceReference(offsetRow, 2)) > 0
                result = result & InputContractRow(CStr(unit), CStr(fixture), "nonadjacent", code, source, reason, 0#)
            End If
            For Each aliasName In Array("RectSet", "TwoRectangle", "TwoRectanglesOffset", "OffsetTwoRectangles")
                If (CStr(fixture) = "LSection" And CStr(aliasName) = "RectSet") Or _
                        (CStr(fixture) = "TwoRectangles" And CStr(aliasName) <> "RectSet") Then
                    range.Formula = block: range.Cells(shapeRow, 2).Value2 = CStr(aliasName)
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

' Перехватывает только результат штатного чтения definition; ожидания ошибки остаются у вызывающего теста.
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
