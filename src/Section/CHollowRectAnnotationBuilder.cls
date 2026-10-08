VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHollowRectAnnotationBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Semantic-аннотации HollowRectangle
' ==========================================================================
' Builder создает только основные размерные линии и групповые подписи
' арматуры. Расчетные волокна, материалы и solver-ы здесь не создаются.
' Точные Outer/Opening уже получены от геометрии; они здесь не меняются.

' Заменяет размеры и подписи арматуры, не изменяя наружную границу,
' отверстия, бетонную сетку или расчетное НДС.
Public Sub Build(ByVal model As CSectionModel, ByVal geometry As ISectionGeometry, ByVal rebars As CRebarLayout)
    If model Is Nothing Then Err.Raise vbObjectError + 5315, "CHollowRectAnnotationBuilder", "Модель сечения не передана."
    If geometry Is Nothing Then Err.Raise vbObjectError + 5316, "CHollowRectAnnotationBuilder", "Геометрия HollowRectangle не передана."

    Dim annotations As CSectionAnnotations
    Set annotations = model.Annotations
    annotations.Clear
    ' Размеры наружной формы и проема берутся из физических параметров
    ' HollowRectangle; сама геометрия контуров уже подготовлена отдельно.
    Dim hollow As CGeometryHollowRectangle
    Dim geometryObject As Object
    Set geometryObject = geometry
    If TypeName(geometryObject) = "CGeometryHollowRectangle" Then Set hollow = geometryObject
    If hollow Is Nothing Then Err.Raise vbObjectError + 5317, "CHollowRectAnnotationBuilder", "Ожидалась геометрия CGeometryHollowRectangle."

    AddDimensions annotations, hollow

    Dim rebarLabels As CRebarGroupAnnotationBuilder
    Set rebarLabels = New CRebarGroupAnnotationBuilder
    rebarLabels.AddRebarLabels annotations, rebars
End Sub

' Размеры выводятся по физическим параметрам таблицы: наружные B/H и размеры
' Opening B_o/H_o с учетом фактического положения отверстия.
Private Sub AddDimensions(ByVal annotations As CSectionAnnotations, ByVal hollow As CGeometryHollowRectangle)
    annotations.AddDimension "DIM_B", hollow.MinX, hollow.MaxY, hollow.MaxX, hollow.MaxY, 0#, 1#, _
        "B", hollow.B, "Наружная ширина HollowRectangle."
    annotations.AddDimension "DIM_H", hollow.MinX, hollow.MinY, hollow.MinX, hollow.MaxY, -1#, 0#, _
        "H", hollow.H, "Наружная высота HollowRectangle."

    Dim bottomWall As Double
    bottomWall = hollow.OpeningMinY - hollow.MinY
    Dim rightWall As Double
    rightWall = hollow.MaxX - hollow.OpeningMaxX

    ' Для размеров Opening базовую линию выносим на толщину соответствующей
    ' стенки: затем обычный пользовательский Offset добавится уже от наружной
    ' стороны стенки и размерная линия не налезет на бетонное кольцо.
    annotations.AddDimension "DIM_B_OPENING", hollow.OpeningMinX, hollow.OpeningMinY - bottomWall, _
        hollow.OpeningMaxX, hollow.OpeningMinY - bottomWall, 0#, -1#, _
        "B_o", hollow.OpeningB, "Ширина внутреннего отверстия Opening."
    annotations.AddDimension "DIM_H_OPENING", hollow.OpeningMaxX + rightWall, hollow.OpeningMinY, _
        hollow.OpeningMaxX + rightWall, hollow.OpeningMaxY, 1#, 0#, _
        "H_o", hollow.OpeningH, "Высота внутреннего отверстия Opening."
End Sub
