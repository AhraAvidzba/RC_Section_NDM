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
' Semantic-аннотации скругленного прямоугольника
' ==========================================================================
' Builder формирует размеры B/H, при необходимости W левой/правой
' tapered-части и групповые подписи арматуры. Он не строит расчетную сетку и
' не принимает решений по NDM и не меняет точные контуры формы.
' Размеры и подписи относятся только к Annotations.

' Заменяет только размеры и подписи арматуры.
' Передает внутренние координаты визуальному слою, не изменяя сетку или НДС.
Public Sub Build(ByVal model As CSectionModel, ByVal geometry As ISectionGeometry, ByVal rebars As CRebarLayout)
    If model Is Nothing Then Err.Raise vbObjectError + 5305, "CRoundedRectAnnotationBuilder", "Модель сечения не передана."
    If geometry Is Nothing Then Err.Raise vbObjectError + 5306, "CRoundedRectAnnotationBuilder", "Геометрия скругленного прямоугольника не передана."

    Dim annotations As CSectionAnnotations
    Set annotations = model.Annotations
    annotations.Clear
    AddDimensions annotations, geometry

    Dim rebarLabels As CRebarGroupAnnotationBuilder
    Set rebarLabels = New CRebarGroupAnnotationBuilder
    rebarLabels.AddRebarLabels annotations, rebars
End Sub

' Добавляет размеры B/H и фактический W' для tapered-сторон RoundedRectangle.
Private Sub AddDimensions(ByVal annotations As CSectionAnnotations, ByVal geometry As ISectionGeometry)
    Dim rounded As CGeometryRoundedRectangle
    Dim geometryObject As Object
    Set geometryObject = geometry
    If TypeName(geometryObject) = "CGeometryRoundedRectangle" Then Set rounded = geometryObject

    If rounded Is Nothing Then
        annotations.AddDimension "DIM_B", geometry.MinX, geometry.MinY, geometry.MaxX, geometry.MinY, 0#, -1#, _
            "B", geometry.MaxX - geometry.MinX, "Габаритная ширина скругленного прямоугольника."
        annotations.AddDimension "DIM_H", geometry.MinX, geometry.MinY, geometry.MinX, geometry.MaxY, -1#, 0#, _
            "H", geometry.MaxY - geometry.MinY, "Габаритная высота скругленного прямоугольника."
        Exit Sub
    End If

    annotations.AddDimension "DIM_B", rounded.LeftBaseX, rounded.TopY, rounded.RightBaseX, rounded.TopY, 0#, 1#, _
        "B", rounded.B, "Размер горизонтальной базы RoundedRectangle между левым и правым торцом."
    annotations.AddDimension "DIM_H", rounded.MinX, rounded.BottomY, rounded.MinX, rounded.TopY, -1#, 0#, _
        "H", rounded.H, "Полная высота RoundedRectangle."

    Dim actualLeftW As Double
    actualLeftW = rounded.LeftBaseX - rounded.MinX
    If StrComp(rounded.LeftType, "Tapered", vbTextCompare) = 0 And actualLeftW > GEOM_TOLERANCE Then
        annotations.AddDimension "DIM_W_LEFT", rounded.MinX, rounded.TopY, _
            rounded.LeftBaseX, rounded.TopY, 0#, 1#, "W'", RoundedLengthForDisplay(actualLeftW), _
            "Фактический вынос левого края бетона относительно левого торца базы B."
    End If
    Dim actualRightW As Double
    actualRightW = rounded.MaxX - rounded.RightBaseX
    If StrComp(rounded.RightType, "Tapered", vbTextCompare) = 0 And actualRightW > GEOM_TOLERANCE Then
        annotations.AddDimension "DIM_W_RIGHT", rounded.RightBaseX, rounded.TopY, _
            rounded.MaxX, rounded.TopY, 0#, 1#, "W'", RoundedLengthForDisplay(actualRightW), _
            "Фактический вынос правого края бетона относительно правого торца базы B."
    End If
End Sub

' Округляет справочный размер W' до целых миллиметров. Сам контур сечения
' остается расчетным и строится по точным координатам; округляется только
' числовая подпись размерной линии.
Private Function RoundedLengthForDisplay(ByVal valueMm As Double) As Double
    If valueMm >= 0# Then
        RoundedLengthForDisplay = Int(valueMm + 0.5)
    Else
        RoundedLengthForDisplay = -Int(Abs(valueMm) + 0.5)
    End If
End Function
