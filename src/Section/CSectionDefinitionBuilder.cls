VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionDefinitionBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Декларативный фасад рецепта. Общая математика выполняется query/path;
' обычный ряд задается коротким вызовом, сложный - типизированными options.
Private mDefinition As CSectionDefinition
Private mQuery As CSectionGeometryQuery
Private mFinished As Boolean

Public Sub Initialize(ByVal typeID As String, Optional ByVal schemaVersion As Long = 1)
    If Not mDefinition Is Nothing Then SectionRecipeError "CSectionDefinitionBuilder", "Фасад уже создан."
    Set mDefinition = New CSectionDefinition: mDefinition.Initialize typeID, schemaVersion
    Set mQuery = New CSectionGeometryQuery
End Sub

Public Function CirclePath(ByVal pathID As String, ByVal centerX As Double, ByVal centerY As Double, ByVal radius As Double) As CGeometryPath
    RequireOpen
    Dim result As CGeometryPath
    Set result = New CGeometryPath: result.Initialize pathID
    result.AddArc pathID, centerX, centerY, radius, 0#, 2# * GEOM_PI
    result.Seal True: mDefinition.AddPath result
    Set CirclePath = result
End Function

Public Function LinePath(ByVal pathID As String, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As CGeometryPath
    RequireOpen
    Dim result As CGeometryPath
    Set result = New CGeometryPath: result.Initialize pathID
    result.AddLine pathID, x1, y1, x2, y2
    result.Seal: mDefinition.AddPath result
    Set LinePath = result
End Function

Public Function ArcPath(ByVal pathID As String, ByVal centerX As Double, ByVal centerY As Double, _
        ByVal radius As Double, ByVal startAngle As Double, ByVal sweep As Double) As CGeometryPath
    RequireOpen
    Dim result As CGeometryPath
    Set result = New CGeometryPath: result.Initialize pathID
    result.AddArc pathID, centerX, centerY, radius, startAngle, sweep
    result.Seal (Abs(sweep) = 2# * GEOM_PI): mDefinition.AddPath result
    Set ArcPath = result
End Function

Public Function RoundedPolygon(ByVal pathID As String, ByVal vertices As Variant, _
        ByVal radii As Variant, ByVal vertexIDs As Variant) As CGeometryPath
    RequireOpen
    Set RoundedPolygon = mQuery.RoundedPolygon(pathID, vertices, radii, vertexIDs)
    mDefinition.AddPath RoundedPolygon
End Function

Public Function OffsetRoundedPolygon(ByVal pathID As String, ByVal vertices As Variant, _
        ByVal radii As Variant, ByVal vertexIDs As Variant, ByVal distance As Double) As CGeometryPath
    RequireOpen
    Set OffsetRoundedPolygon = mQuery.OffsetRoundedPolygon(pathID, vertices, radii, vertexIDs, distance)
    mDefinition.AddPath OffsetRoundedPolygon
End Function

Public Sub AddPath(ByVal path As CGeometryPath)
    RequireOpen
    mDefinition.AddPath path
End Sub

' Привязки принадлежат модели; пиксельные координаты в рецепте не используются.
Public Sub Anchor(ByVal anchorID As String, ByVal x As Double, ByVal y As Double)
    RequireOpen
    mDefinition.AddAnchor anchorID, x, y
End Sub

Public Sub PathAnchor(ByVal anchorID As String, ByVal pathID As String, ByVal segmentID As String, Optional ByVal parameter As Double = 0#)
    RequireOpen
    Dim path As CGeometryPath, x As Double, y As Double, tx As Double, ty As Double
    Set path = mDefinition.Path(pathID)
    path.AtDistance path.AnchorStation(segmentID, parameter), x, y, tx, ty
    mDefinition.AddAnchor anchorID, x, y
End Sub

Public Sub Source(ByVal entityID As String, ByVal sourceReference As String)
    RequireOpen
    mDefinition.AddSource entityID, sourceReference
End Sub

Public Sub TransformAnchor(ByVal anchorID As String, ByVal sourceAnchorID As String, ByVal shiftX As Double, ByVal shiftY As Double, _
        Optional ByVal rotation As Double = 0#, Optional ByVal reflectX As Boolean = False)
    RequireOpen
    Dim x As Double, y As Double
    mDefinition.AnchorCoordinates sourceAnchorID, x, y
    mQuery.TransformCoordinates x, y, shiftX, shiftY, rotation, reflectX
    mDefinition.AddAnchor anchorID, x, y
End Sub

Public Function Offset(ByVal pathID As String, ByVal sourcePathID As String, ByVal signedLeftDistance As Double) As CGeometryPath
    RequireOpen
    On Error GoTo Failed
    Set Offset = mQuery.OffsetPath(pathID, mDefinition.Path(sourcePathID), signedLeftDistance)
    mDefinition.AddPath Offset
    Exit Function
Failed:
    SectionRecipeGeometryError "CSectionDefinitionBuilder", mDefinition.SourceReference(pathID)
End Function

Public Function ConstrainLine(ByVal pathID As String, ByVal sourcePathID As String, _
        ByVal startNX As Double, ByVal startNY As Double, ByVal startBound As Double, _
        ByVal endNX As Double, ByVal endNY As Double, ByVal endBound As Double) As CGeometryPath
    RequireOpen
    On Error GoTo Failed
    Set ConstrainLine = mQuery.ConstrainLine(pathID, mDefinition.Path(sourcePathID), startNX, startNY, startBound, endNX, endNY, endBound)
    mDefinition.AddPath ConstrainLine
    Exit Function
Failed:
    SectionRecipeGeometryError "CSectionDefinitionBuilder", mDefinition.SourceReference(pathID)
End Function

Public Property Get LastGeometryDiagnostic() As CGeometryDiagnostic
    Dim raw As CGeometryDiagnostic, mapped As CGeometryDiagnostic
    If mQuery Is Nothing Then Exit Property
    Set raw = mQuery.LastDiagnostic
    If raw Is Nothing Then Exit Property
    Set mapped = New CGeometryDiagnostic
    mapped.Initialize raw.Code, raw.EntityID, raw.RelatedEntityID, raw.Reason, raw.Detail1, raw.Detail2, mDefinition.SourceReference(raw.EntityID)
    Set LastGeometryDiagnostic = mapped
End Property

Public Function Chain(ByVal pathID As String, ByVal sourcePathID As String, _
        ByVal startSegmentID As String, ByVal startParameter As Double, ByVal endSegmentID As String, ByVal endParameter As Double, _
        Optional ByVal forward As Boolean = True, Optional ByVal fullLoopIfSame As Boolean = False) As CGeometryPath
    RequireOpen
    Dim source As CGeometryPath
    Set source = mDefinition.Path(sourcePathID)
    Set Chain = source.BetweenAnchors(pathID, startSegmentID, startParameter, endSegmentID, endParameter, forward, fullLoopIfSame)
    mDefinition.AddPath Chain
End Function

Public Function Transform(ByVal pathID As String, ByVal sourcePathID As String, ByVal shiftX As Double, ByVal shiftY As Double, _
        Optional ByVal rotation As Double = 0#, Optional ByVal reflectX As Boolean = False) As CGeometryPath
    RequireOpen
    Dim source As CGeometryPath, result As CGeometryPath
    Set source = mDefinition.Path(sourcePathID): Set result = source.Transformed(shiftX, shiftY, rotation, reflectX)
    Set Transform = result.WithID(pathID): mDefinition.AddPath Transform
End Function

Public Function RectanglePath(ByVal pathID As String, ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal height As Double) As CGeometryPath
    RequireOpen
    Set RectanglePath = mQuery.AdjacentRectangles(pathID, Array(Array(x, y, width, height)))
    mDefinition.AddPath RectanglePath
End Function

Public Function AdjacentRectangles(ByVal pathID As String, ByVal rectangles As Variant) As CGeometryPath
    RequireOpen
    Set AdjacentRectangles = mQuery.AdjacentRectangles(pathID, rectangles)
    mDefinition.AddPath AdjacentRectangles
End Function

Public Sub PathGroup(ByVal groupID As String, ByVal pathIDs As Variant)
    RequireOpen
    mDefinition.AddGroup groupID, pathIDs
End Sub

Public Sub Outer(ByVal loopID As String, ByVal pathID As String, Optional ByVal comment As String = "", _
        Optional ByVal contourIDs As Variant, Optional ByVal arcComment As String = "")
    RequireOpen
    Dim value As TSectionLoop
    value.LoopID = loopID: value.PathID = pathID: value.Role = LoopOuter: value.Comment = comment: value.ArcComment = arcComment
    If Not IsMissing(contourIDs) Then value.ContourIDs = contourIDs
    mDefinition.AddLoop value
End Sub

Public Sub Opening(ByVal loopID As String, ByVal pathID As String, ByVal parentOuterID As String, _
        Optional ByVal comment As String = "", Optional ByVal contourIDs As Variant, Optional ByVal arcComment As String = "")
    RequireOpen
    Dim value As TSectionLoop
    value.LoopID = loopID: value.PathID = pathID: value.Role = LoopOpening
    value.ParentID = parentOuterID: value.Comment = comment: value.ArcComment = arcComment
    If Not IsMissing(contourIDs) Then value.ContourIDs = contourIDs
    mDefinition.AddLoop value
End Sub

Public Sub RebarByCount(ByVal ruleID As String, ByVal target As String, ByVal count As Long, ByVal diameter As Double, _
        Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal steelClass As String = "Rebar")
    Dim value As TRebarRule
    PrepareRule value, ruleID, target, ByCount, diameter, steelClass, options
    If count < 0 Then SectionRecipeError "CSectionDefinitionBuilder", "Количество стержней не может быть отрицательным.", value.SourceReference
    value.Count = count: value.Enabled = (count > 0 And diameter > 0#)
    mDefinition.AddRule value
End Sub

Public Sub RebarByMaxStep(ByVal ruleID As String, ByVal target As String, ByVal step As Double, ByVal diameter As Double, _
        Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal steelClass As String = "Rebar")
    Dim value As TRebarRule
    PrepareRule value, ruleID, target, ByMaxStep, diameter, steelClass, options
    value.Step = step: mDefinition.AddRule value
End Sub

Public Sub RebarByFixedStep(ByVal ruleID As String, ByVal target As String, ByVal step As Double, ByVal diameter As Double, _
        Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal steelClass As String = "Rebar")
    Dim value As TRebarRule
    PrepareRule value, ruleID, target, ByFixedStep, diameter, steelClass, options
    value.Step = step: mDefinition.AddRule value
End Sub

Public Sub RebarAtStations(ByVal ruleID As String, ByVal target As String, ByVal stations As Variant, ByVal diameter As Double, _
        Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal steelClass As String = "Rebar")
    Dim value As TRebarRule
    PrepareRule value, ruleID, target, ExplicitPositions, diameter, steelClass, options
    value.ExplicitStations = stations: mDefinition.AddRule value
End Sub

Public Sub RebarDependent(ByVal ruleID As String, ByVal sourceRuleID As String, ByVal diameter As Double, _
        ByVal options As CRebarRuleOptions, Optional ByVal steelClass As String = "Rebar")
    Dim value As TRebarRule
    PrepareRule value, ruleID, "", DependentOffset, diameter, steelClass, options
    value.SourceRuleID = sourceRuleID: mDefinition.AddRule value
End Sub

Public Sub RebarProjected(ByVal ruleID As String, ByVal target As String, ByVal sourceRuleID As String, ByVal diameter As Double, _
        ByVal options As CRebarRuleOptions, Optional ByVal steelClass As String = "Rebar")
    Dim value As TRebarRule
    PrepareRule value, ruleID, target, ProjectFrom, diameter, steelClass, options
    value.SourceRuleID = sourceRuleID: mDefinition.AddRule value
End Sub

Private Sub PrepareRule(ByRef value As TRebarRule, ByVal ruleID As String, ByVal target As String, ByVal mode As ERebarDistribution, _
        ByVal diameter As Double, ByVal steelClass As String, ByVal options As CRebarRuleOptions)
    RequireOpen
    value.RuleID = ruleID: value.Target = target: value.Mode = mode: value.Diameter = diameter
    value.SteelClass = steelClass: value.Enabled = (diameter > 0#)
    If options Is Nothing Then Set options = New CRebarRuleOptions
    options.CopyTo value
    If diameter < 0# Then SectionRecipeError "CSectionDefinitionBuilder", "Диаметр не может быть отрицательным.", value.SourceReference
End Sub

' Измерение выбирается по точной геометрии, размещение - отдельно.
Public Sub VerticalExtent(ByVal annotationID As String, ByVal target As String, ByVal name As String, _
        ByVal placementX As Double, Optional ByVal normalX As Double = -1#, Optional ByVal comment As String = "")
    RequireOpen
    Dim value As TSectionAnnotation
    value.AnnotationID = annotationID: value.Kind = "VerticalExtent": value.Target = target: value.Name = name
    value.StartX = placementX: value.EndX = placementX: value.NormalX = normalX: value.Comment = comment
    mDefinition.AddAnnotation value
End Sub

Public Sub HorizontalExtent(ByVal annotationID As String, ByVal target As String, ByVal name As String, _
        ByVal placementY As Double, Optional ByVal normalY As Double = 1#, Optional ByVal comment As String = "")
    RequireOpen
    Dim value As TSectionAnnotation
    value.AnnotationID = annotationID: value.Kind = "HorizontalExtent": value.Target = target: value.Name = name
    value.StartY = placementY: value.EndY = placementY: value.NormalY = normalY: value.Comment = comment
    mDefinition.AddAnnotation value
End Sub

Public Sub HorizontalDistance(ByVal annotationID As String, ByVal name As String, _
        ByVal startX As Double, ByVal endX As Double, ByVal placementY As Double, _
        Optional ByVal normalY As Double = 1#, Optional ByVal comment As String = "", Optional ByVal displayIncrement As Double = 0#)
    RequireOpen
    If displayIncrement < 0# Then SectionRecipeError "CSectionDefinitionBuilder", "Шаг округления подписи не может быть отрицательным."
    Dim value As TSectionAnnotation
    value.AnnotationID = annotationID: value.Kind = "BetweenPoints": value.Name = name
    value.StartX = startX: value.EndX = endX: value.StartY = placementY: value.EndY = placementY
    value.NormalY = normalY: value.Value = endX - startX: value.Comment = comment: value.DisplayIncrement = displayIncrement
    mDefinition.AddAnnotation value
End Sub

Public Sub VerticalDistance(ByVal annotationID As String, ByVal name As String, _
        ByVal startY As Double, ByVal endY As Double, ByVal placementX As Double, _
        Optional ByVal normalX As Double = -1#, Optional ByVal comment As String = "")
    RequireOpen
    Dim value As TSectionAnnotation
    value.AnnotationID = annotationID: value.Kind = "BetweenPoints": value.Name = name
    value.StartX = placementX: value.EndX = placementX: value.StartY = startY: value.EndY = endY
    value.NormalX = normalX: value.Value = Abs(endY - startY): value.Comment = comment
    mDefinition.AddAnnotation value
End Sub

' Размер между именованными привязками: Horizontal/Vertical/Aligned.
Public Sub Dimension(ByVal annotationID As String, ByVal startAnchorID As String, ByVal endAnchorID As String, _
        ByVal name As String, Optional ByVal measurement As String = "Aligned", _
        Optional ByVal normalX As Double = 0#, Optional ByVal normalY As Double = 0#, Optional ByVal comment As String = "")
    RequireOpen
    If measurement <> "Horizontal" And measurement <> "Vertical" And measurement <> "Aligned" Then _
        SectionRecipeError "CSectionDefinitionBuilder", "Measurement должен быть Horizontal, Vertical или Aligned."
    Dim value As TSectionAnnotation
    value.AnnotationID = annotationID: value.Kind = "Anchors." & measurement: value.Name = name
    value.StartAnchorID = startAnchorID: value.EndAnchorID = endAnchorID
    value.NormalX = normalX: value.NormalY = normalY: value.Comment = comment
    mDefinition.AddAnnotation value
End Sub

Public Sub PathLength(ByVal annotationID As String, ByVal target As String, ByVal placementAnchorID As String, _
        Optional ByVal name As String = "L", Optional ByVal comment As String = "")
    RequireOpen
    Dim value As TSectionAnnotation
    value.AnnotationID = annotationID: value.Kind = "PathLength": value.Target = target
    value.StartAnchorID = placementAnchorID: value.Name = name: value.Comment = comment
    mDefinition.AddAnnotation value
End Sub

Public Sub Radius(ByVal annotationID As String, ByVal pathID As String, ByVal segmentID As String, _
        Optional ByVal parameter As Double = 0.5, Optional ByVal name As String = "R", Optional ByVal comment As String = "")
    CircularDimension annotationID, pathID, segmentID, parameter, name, comment, False
End Sub

Public Sub Diameter(ByVal annotationID As String, ByVal pathID As String, ByVal segmentID As String, _
        Optional ByVal parameter As Double = 0.5, Optional ByVal name As String = "D", Optional ByVal comment As String = "")
    CircularDimension annotationID, pathID, segmentID, parameter, name, comment, True
End Sub

Private Sub CircularDimension(ByVal annotationID As String, ByVal pathID As String, ByVal segmentID As String, _
        ByVal parameter As Double, ByVal name As String, ByVal comment As String, ByVal diameter As Boolean)
    RequireOpen
    If parameter < 0# Or parameter > 1# Then SectionRecipeError "CSectionDefinitionBuilder", "Параметр точки дуги должен лежать в [0;1]."
    Dim value As TSectionAnnotation
    value.AnnotationID = annotationID: value.Kind = "Radius": If diameter Then value.Kind = "Diameter"
    value.Target = pathID: value.SegmentID = segmentID: value.Value = parameter
    value.Name = name: value.Comment = comment
    mDefinition.AddAnnotation value
End Sub

Public Sub Text(ByVal annotationID As String, ByVal anchorID As String, ByVal textValue As String, _
        Optional ByVal comment As String = "")
    RequireOpen
    Dim value As TSectionAnnotation
    value.AnnotationID = annotationID: value.Kind = "Text": value.StartAnchorID = anchorID
    value.Name = textValue: value.Comment = comment
    mDefinition.AddAnnotation value
End Sub

Public Sub RebarLabel(ByVal groupID As String, ByVal target As String, ByVal normalX As Double, ByVal normalY As Double, _
        Optional ByVal axisDistance As Double = 0#)
    RequireOpen
    Dim value As TSectionAnnotation
    value.AnnotationID = "REBAR_" & groupID: value.Kind = "RebarLabel": value.Target = target: value.Name = groupID
    value.NormalX = normalX: value.NormalY = normalY: value.Value = axisDistance
    mDefinition.AddAnnotation value
End Sub

Public Sub RebarRowLabel(ByVal groupID As String, ByVal ruleID As String, ByVal normalX As Double, ByVal normalY As Double, _
        Optional ByVal axisDistance As Double = 0#)
    RequireOpen
    Dim value As TSectionAnnotation
    value.AnnotationID = "REBAR_" & groupID: value.Kind = "RebarRowLabel": value.Target = ruleID: value.Name = groupID
    value.NormalX = normalX: value.NormalY = normalY: value.Value = axisDistance
    mDefinition.AddAnnotation value
End Sub

' Подпись по фактическим экстремумам основного ряда, а не по всей траектории.
Public Sub RebarRowExtentLabel(ByVal groupID As String, ByVal ruleID As String, ByVal normalX As Double, ByVal normalY As Double, _
        Optional ByVal axisDistance As Double = 0#, Optional ByVal reverseDirection As Boolean = False)
    RequireOpen
    Dim value As TSectionAnnotation
    value.AnnotationID = "REBAR_" & groupID: value.Kind = "RebarRowExtentLabel": value.Target = ruleID: value.Name = groupID
    value.NormalX = normalX: value.NormalY = normalY: value.Value = axisDistance: value.ReverseDirection = reverseDirection
    mDefinition.AddAnnotation value
End Sub

Public Function Finish() As CSectionDefinition
    RequireOpen
    mDefinition.Seal
    mFinished = True
    Set Finish = mDefinition
End Function

' Геометрическая сторона крайнего ряда и направление подписи независимы:
' у отверстия подпись обращена в пустоту, а ряд выбирается со стороны бетона.
Public Sub RebarRowExtremeLabel(ByVal groupID As String, ByVal ruleID As String, ByVal directionX As Double, ByVal directionY As Double, _
        ByVal normalX As Double, ByVal normalY As Double, Optional ByVal axisDistance As Double = 0#)
    RequireOpen
    If Abs(directionX * directionX + directionY * directionY - 1#) > GEOM_TOLERANCE Then _
        SectionRecipeError "CSectionDefinitionBuilder", "Направление крайнего ряда должно быть единичным."
    Dim value As TSectionAnnotation
    value.AnnotationID = "REBAR_" & groupID: value.Kind = "RebarRowExtremeLabel": value.Target = ruleID: value.Name = groupID
    value.DirectionX = directionX: value.DirectionY = directionY
    value.NormalX = normalX: value.NormalY = normalY: value.Value = axisDistance
    mDefinition.AddAnnotation value
End Sub

Private Sub RequireOpen()
    If mDefinition Is Nothing Or mFinished Then SectionRecipeError "CSectionDefinitionBuilder", "Фасад не открыт для описания сечения."
End Sub
