VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCircleAnnotationBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Semantic-аннотации круглого сечения
' ==========================================================================
' Builder знает, какие размеры и подписи имеют смысл именно для Circle:
' внешний контур круга, диаметр и групповая подпись арматуры по окружности.
' Он передает только model-координаты; Excel Chart здесь не участвует.

Public Sub Build(ByVal model As CSectionModel, ByVal geometry As ISectionGeometry, ByVal rebars As CRebarLayout)
    If model Is Nothing Then Err.Raise vbObjectError + 5295, "CCircleAnnotationBuilder", "Section model is not provided."
    If geometry Is Nothing Then Err.Raise vbObjectError + 5296, "CCircleAnnotationBuilder", "Geometry is not provided."

    Dim annotations As CSectionAnnotations
    Set annotations = model.Annotations
    annotations.Clear

    Dim centerX As Double
    Dim centerY As Double
    Dim radius As Double
    centerX = 0.5 * (geometry.MinX + geometry.MaxX)
    centerY = 0.5 * (geometry.MinY + geometry.MaxY)
    radius = 0.5 * GeomMin(geometry.MaxX - geometry.MinX, geometry.MaxY - geometry.MinY)

    annotations.AddContourCircle "CONTOUR_CIRCLE", centerX, centerY, radius
    annotations.AddDimension "DIM_CIRCLE_D", geometry.MinX, geometry.MinY, geometry.MinX, geometry.MaxY, _
        -1#, 0#, "D", geometry.MaxY - geometry.MinY, "Diameter dimension for generated circle."

    Dim rebarLabels As CRebarGroupAnnotationBuilder
    Set rebarLabels = New CRebarGroupAnnotationBuilder
    rebarLabels.AddRebarLabels annotations, rebars
End Sub

