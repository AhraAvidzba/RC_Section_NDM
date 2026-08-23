VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRoundedRectAnnotationBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Semantic-аннотации прямоугольника со скруглениями
' ==========================================================================
' Builder формирует контур и основные размеры RoundedRectangle. Его результат
' должен быть пригоден для того же универсального Reader/Layout/Plotter, что
' и аннотации Circle или LShape.

Public Sub Build(ByVal model As CSectionModel, ByVal geometry As ISectionGeometry, ByVal rebars As CRebarLayout)
    If model Is Nothing Then Err.Raise vbObjectError + 5305, "CRoundedRectAnnotationBuilder", "Модель сечения не передана."
    If geometry Is Nothing Then Err.Raise vbObjectError + 5306, "CRoundedRectAnnotationBuilder", "Геометрия скругленного прямоугольника не передана."

    Dim annotations As CSectionAnnotations
    Set annotations = model.Annotations
    annotations.Clear

    annotations.AddContourLine "CONTOUR_LINE_1", geometry.MinX, geometry.MinY, geometry.MaxX, geometry.MinY
    annotations.AddContourLine "CONTOUR_LINE_2", geometry.MaxX, geometry.MinY, geometry.MaxX, geometry.MaxY
    annotations.AddContourLine "CONTOUR_LINE_3", geometry.MaxX, geometry.MaxY, geometry.MinX, geometry.MaxY
    annotations.AddContourLine "CONTOUR_LINE_4", geometry.MinX, geometry.MaxY, geometry.MinX, geometry.MinY
    annotations.AddDimension "DIM_B", geometry.MinX, geometry.MinY, geometry.MaxX, geometry.MinY, 0#, -1#, _
        "B", geometry.MaxX - geometry.MinX, "Overall concrete width."
    annotations.AddDimension "DIM_H", geometry.MinX, geometry.MinY, geometry.MinX, geometry.MaxY, -1#, 0#, _
        "H", geometry.MaxY - geometry.MinY, "Overall concrete height."

    Dim rebarLabels As CRebarGroupAnnotationBuilder
    Set rebarLabels = New CRebarGroupAnnotationBuilder
    rebarLabels.AddRebarLabels annotations, rebars
End Sub

