VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLShapeAnnotationBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Semantic-аннотации Г-образного сечения
' ==========================================================================
' Builder знает смысл граней LShape: где полная высота, общая ширина, стенка,
' полка и линии арматуры. Он формирует CPlotDimension и CPlotRebarLabel, но
' не решает, как эти подписи будут разведены по lanes на листе.

Public Sub Build(ByVal model As CSectionModel, ByVal geometry As ISectionGeometry, ByVal rebars As CRebarLayout)
    If model Is Nothing Then Err.Raise vbObjectError + 5300, "CLShapeAnnotationBuilder", "Section model is not provided."
    If geometry Is Nothing Then Err.Raise vbObjectError + 5301, "CLShapeAnnotationBuilder", "Geometry is not provided."

    Dim annotations As CSectionAnnotations
    Set annotations = model.Annotations
    annotations.Clear

    AddContour annotations, geometry
    AddDimensions annotations, geometry

    Dim rebarLabels As CRebarGroupAnnotationBuilder
    Set rebarLabels = New CRebarGroupAnnotationBuilder
    rebarLabels.AddRebarLabels annotations, rebars
End Sub

Private Sub AddContour(ByVal annotations As CSectionAnnotations, ByVal geometry As ISectionGeometry)
    Dim px() As Double
    Dim py() As Double
    geometry.GetExtremePoints px, py

    Dim i As Long
    Dim nextIndex As Long
    For i = LBound(px) To UBound(px)
        If i = UBound(px) Then
            nextIndex = LBound(px)
        Else
            nextIndex = i + 1
        End If
        annotations.AddContourLine "CONTOUR_LINE_" & CStr(i), px(i), py(i), px(nextIndex), py(nextIndex)
    Next i
End Sub

Private Sub AddDimensions(ByVal annotations As CSectionAnnotations, ByVal geometry As ISectionGeometry)
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    minX = geometry.MinX
    maxX = geometry.MaxX
    minY = geometry.MinY
    maxY = geometry.MaxY

    annotations.AddDimension "DIM_B", minX, minY, maxX, minY, 0#, -1#, "B", maxX - minX, "Overall concrete width."
    annotations.AddDimension "DIM_H", minX, minY, minX, maxY, -1#, 0#, "H", maxY - minY, "Overall concrete height."

    Dim px() As Double
    Dim py() As Double
    geometry.GetExtremePoints px, py
    If UBound(px) - LBound(px) + 1 >= 6 Then
        annotations.AddDimension "DIM_TOP_WIDTH", px(6), py(6), px(5), py(5), 0#, 1#, "B1", Abs(px(5) - px(6)), "Upper rectangle width."
        annotations.AddDimension "DIM_RIGHT_HEIGHT", px(2), py(2), px(3), py(3), 1#, 0#, "H2", Abs(py(3) - py(2)), "Lower rectangle height."
    End If
End Sub

