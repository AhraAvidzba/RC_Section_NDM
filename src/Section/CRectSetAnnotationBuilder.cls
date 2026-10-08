VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRectSetAnnotationBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Semantic-аннотации RectSet-сечения
' ==========================================================================
' Builder знает смысл верхнего и нижнего прямоугольников: общий габарит,
' размеры активных прямоугольников и линии арматуры. Он формирует
' только размерные/арматурные аннотации; исходный контур принадлежит геометрии.
' Визуальное размещение подписей на листе не относится к ответственности builder-а.

' Заменяет размеры активных прямоугольников и подписи арматуры.
' Исходные контуры, расчетное сечение и его НДС здесь не изменяются.
Public Sub Build(ByVal model As CSectionModel, ByVal geometry As ISectionGeometry, ByVal rebars As CRebarLayout)
    If model Is Nothing Then Err.Raise vbObjectError + 5300, "CRectSetAnnotationBuilder", "Модель сечения не передана."
    If geometry Is Nothing Then Err.Raise vbObjectError + 5301, "CRectSetAnnotationBuilder", "Геометрия RectSet-сечения не передана."

    Dim annotations As CSectionAnnotations
    Set annotations = model.Annotations
    annotations.Clear
    AddDimensions annotations, geometry

    Dim rebarLabels As CRebarGroupAnnotationBuilder
    Set rebarLabels = New CRebarGroupAnnotationBuilder
    rebarLabels.AddRebarLabels annotations, rebars
End Sub

' Добавляет основные размеры RectSet и, при необходимости, смещение верхнего прямоугольника.
Private Sub AddDimensions(ByVal annotations As CSectionAnnotations, ByVal geometry As ISectionGeometry)
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    minX = geometry.MinX
    maxX = geometry.MaxX
    minY = geometry.MinY
    maxY = geometry.MaxY

    annotations.AddDimension "DIM_B", minX, minY, maxX, minY, 0#, -1#, "B", maxX - minX, "Общая ширина RectSet-сечения."
    annotations.AddDimension "DIM_H", minX, minY, minX, maxY, -1#, 0#, "H", maxY - minY, "Общая высота RectSet-сечения."

    Dim geometryObject As Object
    Set geometryObject = geometry
    If TypeName(geometryObject) <> "CGeometryRectSet" Then Exit Sub

    Dim rectSet As CGeometryRectSet
    Set rectSet = geometryObject

    annotations.AddDimension "DIM_TOP_WIDTH", rectSet.UpperMinX, rectSet.UpperMaxY, _
        rectSet.UpperMaxX, rectSet.UpperMaxY, 0#, 1#, "B1", rectSet.B1, _
        "Ширина верхнего прямоугольника RectSet-сечения."

    If rectSet.HasLowerRectangle Then
        annotations.AddDimension "DIM_RIGHT_HEIGHT", rectSet.LowerMaxX, rectSet.LowerMinY, _
            rectSet.LowerMaxX, rectSet.LowerMaxY, 1#, 0#, "H2", rectSet.H2, _
            "Высота нижнего прямоугольника RectSet-сечения."

        If StrComp(rectSet.SectionType, "TwoRectangles", vbTextCompare) = 0 And Abs(rectSet.EffectiveUpperOffsetX) > GEOM_TOLERANCE Then
            annotations.AddDimension "DIM_UPPER_OFFSET_X", rectSet.LowerMinX, rectSet.LowerMaxY, _
                rectSet.UpperMinX, rectSet.LowerMaxY, 0#, 1#, ChrW$(&H394), rectSet.EffectiveUpperOffsetX, _
                "Горизонтальное смещение верхнего прямоугольника относительно базового положения."
        End If
    End If
End Sub
