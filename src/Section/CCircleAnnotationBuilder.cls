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
' диаметр и групповая подпись арматуры по окружности. Исходный контур
' уже создан геометрическим классом; builder меняет только Annotations.
' Передаются внутренние координаты; Excel Chart здесь не участвует.

' Заменяет только размерные аннотации и подписи арматуры круга.
' Координаты остаются внутренними; расчетная геометрия и НДС не изменяются.
Public Sub Build(ByVal model As CSectionModel, ByVal geometry As ISectionGeometry, ByVal rebars As CRebarLayout)
    If model Is Nothing Then Err.Raise vbObjectError + 5295, "CCircleAnnotationBuilder", "Модель сечения не передана."
    If geometry Is Nothing Then Err.Raise vbObjectError + 5296, "CCircleAnnotationBuilder", "Геометрия круглого сечения не передана."

    Dim annotations As CSectionAnnotations
    Set annotations = model.Annotations
    annotations.Clear
    annotations.AddDimension "DIM_CIRCLE_D", geometry.MinX, geometry.MinY, geometry.MinX, geometry.MaxY, _
        -1#, 0#, "D", geometry.MaxY - geometry.MinY, "Диаметр параметрического круглого сечения."

    Dim rebarLabels As CRebarGroupAnnotationBuilder
    Set rebarLabels = New CRebarGroupAnnotationBuilder
    rebarLabels.AddRebarLabels annotations, rebars
End Sub

