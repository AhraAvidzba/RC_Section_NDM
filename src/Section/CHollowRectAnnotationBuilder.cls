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
' Builder записывает в Results точный внешний контур, контур Opening, основные
' размерные линии и групповые подписи арматуры. Расчетные волокна, материалы и
' solver-ы здесь не создаются: это только слой визуализации/AutoCAD export.

' Формирует все аннотации для последнего расчетного снимка.
Public Sub Build(ByVal model As CSectionModel, ByVal geometry As ISectionGeometry, ByVal rebars As CRebarLayout)
    If model Is Nothing Then Err.Raise vbObjectError + 5315, "CHollowRectAnnotationBuilder", "Модель сечения не передана."
    If geometry Is Nothing Then Err.Raise vbObjectError + 5316, "CHollowRectAnnotationBuilder", "Геометрия HollowRectangle не передана."

    Dim annotations As CSectionAnnotations
    Set annotations = model.Annotations
    annotations.Clear

    ' Аннотации HollowRectangle требуют именно concrete geometry этого типа:
    ' наружная и внутренняя петли контура нужны AutoCAD export и Excel-схеме.
    Dim hollow As CGeometryHollowRectangle
    Dim geometryObject As Object
    Set geometryObject = geometry
    If TypeName(geometryObject) = "CGeometryHollowRectangle" Then Set hollow = geometryObject
    If hollow Is Nothing Then Err.Raise vbObjectError + 5317, "CHollowRectAnnotationBuilder", "Ожидалась геометрия CGeometryHollowRectangle."

    ' Контуры пишутся первыми, затем размерные линии и подписи групп арматуры.
    ' Такой порядок удобен для чтения Results и не влияет на расчетную модель.
    AddContourLoop annotations, hollow, False, "OUTER", "наружного контура"
    AddContourLoop annotations, hollow, True, "OPENING", "внутреннего контура Opening"
    AddDimensions annotations, hollow

    Dim rebarLabels As CRebarGroupAnnotationBuilder
    Set rebarLabels = New CRebarGroupAnnotationBuilder
    rebarLabels.AddRebarLabels annotations, rebars
End Sub

' Добавляет одну замкнутую петлю контура. Суффикс OUTER/OPENING нужен AutoCAD
' export, чтобы не пытаться сшить наружный контур и отверстие одной полилинией.
Private Sub AddContourLoop(ByVal annotations As CSectionAnnotations, ByVal hollow As CGeometryHollowRectangle, _
        ByVal openingLoop As Boolean, ByVal loopName As String, ByVal commentPart As String)
    Dim segmentType() As String
    Dim startX() As Double, startY() As Double
    Dim endX() As Double, endY() As Double
    Dim centerX() As Double, centerY() As Double
    Dim radiusValue() As Double, sweepAngle() As Double
    Dim segmentCount As Long
    hollow.GetContourSegments openingLoop, segmentType, startX, startY, endX, endY, centerX, centerY, radiusValue, sweepAngle, segmentCount

    ' Каждая дуга и линия получает отдельную semantic-аннотацию. Позже plotter и
    ' AutoCAD export сами решают, как визуально отрисовать эти сегменты.
    Dim i As Long
    For i = 1 To segmentCount
        If StrComp(segmentType(i), "ARC", vbTextCompare) = 0 Then
            annotations.AddContourArc "CONTOUR_" & loopName & "_ARC_" & CStr(i), _
                startX(i), startY(i), endX(i), endY(i), centerX(i), centerY(i), radiusValue(i), sweepAngle(i), _
                "Дуга " & commentPart & " HollowRectangle."
        Else
            annotations.AddContourLine "CONTOUR_" & loopName & "_LINE_" & CStr(i), _
                startX(i), startY(i), endX(i), endY(i), _
                "Прямой участок " & commentPart & " HollowRectangle."
        End If
    Next i
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
