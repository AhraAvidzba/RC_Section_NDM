VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CGeometryDistance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Точные расстояния между ограниченными прямыми и круговыми дугами
' ==========================================================================
' Общий запрос для бетонной стенки и арматурного контура; координаты в мм.
' Сегменты должны быть невырожденными и подготовленными геометрическим ядром.

' Минимум расстояния между двумя аналитическими сегментами. При отсутствии
' пересечения проверяет концы и внутренние пары с общими нормалями.
Friend Function EdgeDistance(ByRef first As TRegionEdge, ByRef second As TRegionEdge) As Double
    Dim intersections() As Double, count As Long
    modGeometryPrimitives.EdgeIntersectionCuts first, second, intersections, count
    If count > 0 Then Exit Function ' При пересечении или касании зазор равен нулю.
    Dim parameter As Double, unused As Boolean, minimum As Double
    minimum = PointDistance(first, second.X1, second.Y1, parameter, unused)
    minimum = GeomMin(minimum, PointDistance(first, second.X2, second.Y2, parameter, unused))
    minimum = GeomMin(minimum, PointDistance(second, first.X1, first.Y1, parameter, unused))
    minimum = GeomMin(minimum, PointDistance(second, first.X2, first.Y2, parameter, unused))
    If IsArc(first) Then
        If IsArc(second) Then
            minimum = ArcNormalGap(first, second, second.CenterX - first.CenterX, second.CenterY - first.CenterY, minimum)
        Else
            minimum = ArcNormalGap(first, second, second.Y1 - second.Y2, second.X2 - second.X1, minimum)
        End If
    ElseIf IsArc(second) Then
        minimum = ArcNormalGap(second, first, first.Y1 - first.Y2, first.X2 - first.X1, minimum)
    End If
    EdgeDistance = minimum
End Function

' Проверяет внутренние экстремумы дуги по нормали прямой или по
' линии центров двух окружностей. Для концентрических дуг достаточно концов.
Private Function ArcNormalGap(ByRef arc As TRegionEdge, ByRef other As TRegionEdge, _
        ByVal nx As Double, ByVal ny As Double, ByVal minimum As Double) As Double
    ArcNormalGap = minimum
    Dim norm As Double: norm = DistanceForVector(nx, ny)
    If norm = 0# Then Exit Function
    Dim radius As Double: radius = DistanceForVector(arc.X1 - arc.CenterX, arc.Y1 - arc.CenterY)
    Dim side As Long, x As Double, y As Double, parameter As Double, unused As Boolean
    For side = -1 To 1 Step 2
        x = arc.CenterX + side * radius * (nx / norm): y = arc.CenterY + side * radius * (ny / norm)
        If modGeometryPrimitives.PointParameter(arc, x, y, parameter) Then _
            ArcNormalGap = GeomMin(ArcNormalGap, PointDistance(other, x, y, parameter, unused))
    Next side
End Function

' Проецирует точку на ограниченный сегмент и возвращает расстояние.
' Для дуги рассматривает радиальную проекцию и оба конца.
Friend Function PointDistance(ByRef edge As TRegionEdge, ByVal x As Double, ByVal y As Double, _
        ByRef parameter As Double, ByRef ambiguous As Boolean) As Double
    Dim dx As Double, dy As Double, nearX As Double, nearY As Double
    ambiguous = False
    If Not IsArc(edge) Then
        dx = edge.X2 - edge.X1: dy = edge.Y2 - edge.Y1
        parameter = GeomMax(0#, GeomMin(1#, ((x - edge.X1) * dx + (y - edge.Y1) * dy) / (dx * dx + dy * dy)))
        PointDistance = DistanceForVector(x - edge.X1 - parameter * dx, y - edge.Y1 - parameter * dy)
        Exit Function
    End If
    dx = x - edge.CenterX: dy = y - edge.CenterY
    Dim radial As Double, radius As Double
    radial = DistanceForVector(dx, dy): radius = DistanceForVector(edge.X1 - edge.CenterX, edge.Y1 - edge.CenterY)
    If radial = 0# Then
        parameter = 0#: ambiguous = True: PointDistance = radius
        Exit Function
    End If
    nearX = edge.CenterX + radius * (dx / radial): nearY = edge.CenterY + radius * (dy / radial)
    If modGeometryPrimitives.PointParameter(edge, nearX, nearY, parameter) Then
        PointDistance = Abs(radial - radius)
    Else
        Dim firstGap As Double, lastGap As Double
        firstGap = DistanceForVector(x - edge.X1, y - edge.Y1): lastGap = DistanceForVector(x - edge.X2, y - edge.Y2)
        ambiguous = Abs(firstGap - lastGap) <= GEOM_BOUNDARY_TOLERANCE
        If firstGap <= lastGap Then
            parameter = 0#: PointDistance = firstGap
        Else
            parameter = 1#: PointDistance = lastGap
        End If
    End If
End Function

' Возвращает длину плоского вектора, мм.
Private Function DistanceForVector(ByVal x As Double, ByVal y As Double) As Double
    DistanceForVector = Sqr(x * x + y * y)
End Function

' Согласует тип сегмента с аналитическим геометрическим ядром.
Private Function IsArc(ByRef edge As TRegionEdge) As Boolean
    IsArc = Abs(edge.Sweep) > GEOM_PARAMETER_TOLERANCE
End Function
