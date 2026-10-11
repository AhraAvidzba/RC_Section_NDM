VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionAnnotationBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Общая интерпретация модельных измерений и привязок. Оформление и положение
' в пикселях остаются в существующих layout/plotter; контуры не меняются.
Public Sub Build(ByVal definition As CSectionDefinition, ByVal rebars As CRebarLayout, ByVal annotations As CSectionAnnotations, _
        Optional ByVal rules As CRebarRuleExecutor = Nothing)
    Dim value As TSectionAnnotation, i As Long, minX As Double, maxX As Double, minY As Double, maxY As Double
    annotations.Clear
    For i = 1 To definition.AnnotationCount
        definition.CopyAnnotation i, value
        If value.Kind <> "RebarRowLabel" And value.Kind <> "BetweenPoints" Then TargetBounds definition, value.Target, minX, maxX, minY, maxY
        Select Case value.Kind
            Case "BetweenPoints"
                annotations.AddDimension value.AnnotationID, value.StartX, value.StartY, value.EndX, value.EndY, _
                    value.NormalX, value.NormalY, value.Name, value.Value, value.Comment
            Case "VerticalExtent"
                annotations.AddDimension value.AnnotationID, value.StartX, minY, value.EndX, maxY, _
                    value.NormalX, value.NormalY, value.Name, maxY - minY, value.Comment
            Case "HorizontalExtent"
                annotations.AddDimension value.AnnotationID, minX, value.StartY, maxX, value.EndY, _
                    value.NormalX, value.NormalY, value.Name, maxX - minX, value.Comment
            Case "RebarLabel"
                If Abs(value.NormalY) >= Abs(value.NormalX) Then
                    value.StartX = minX: value.EndX = maxX
                    If value.NormalY >= 0# Then value.StartY = maxY Else value.StartY = minY
                    value.EndY = value.StartY
                Else
                    value.StartY = minY: value.EndY = maxY
                    If value.NormalX >= 0# Then value.StartX = maxX Else value.StartX = minX
                    value.EndX = value.StartX
                End If
                rebars.AddAnnotationAnchor value.Name, value.StartX, value.StartY, value.EndX, value.EndY, _
                    value.NormalX, value.NormalY, value.Value
            Case "RebarRowLabel"
                If rules Is Nothing Then SectionRecipeError "CSectionAnnotationBuilder", "Для подписи ряда нужны исполненные правила."
                If rules.RuleEnds(value.Target, value.StartX, value.StartY, value.EndX, value.EndY) Then _
                    rebars.AddAnnotationAnchor value.Name, value.StartX, value.StartY, value.EndX, value.EndY, _
                        value.NormalX, value.NormalY, value.Value
            Case Else
                SectionRecipeError "CSectionAnnotationBuilder", "Неизвестное измерение '" & value.Kind & "'."
        End Select
    Next i
    Dim labels As CRebarGroupAnnotationBuilder
    Set labels = New CRebarGroupAnnotationBuilder
    labels.AddRebarLabels annotations, rebars
End Sub

Private Sub TargetBounds(ByVal definition As CSectionDefinition, ByVal target As String, _
        ByRef minX As Double, ByRef maxX As Double, ByRef minY As Double, ByRef maxY As Double)
    Dim ids As Variant, id As Variant, path As CGeometryPath, low As Double, high As Double
    ids = definition.TargetPaths(target)
    minX = 1E+100: maxX = -1E+100: minY = 1E+100: maxY = -1E+100
    For Each id In ids
        Set path = definition.Path(CStr(id))
        path.ProjectionBounds 1#, 0#, low, high: minX = GeomMin(minX, low): maxX = GeomMax(maxX, high)
        path.ProjectionBounds 0#, 1#, low, high: minY = GeomMin(minY, low): maxY = GeomMax(maxY, high)
    Next id
End Sub
