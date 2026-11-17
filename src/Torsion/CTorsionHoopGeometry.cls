VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTorsionHoopGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Геометрия одного замкнутого хомута
' ==========================================================================
' Хранит проверенную ось в плоскости сечения, её длину и охваченную площадь.
' Только точные отрезки и круговые дуги. Отверстия внутри оси допустимы,
' если стержень и заданный защитный слой целиком находятся в бетоне.
' Кинематику и применимость модели кручения этот класс не определяет.

Private mPath As CGeometryPath
Private mEnclosed As CConcreteRegion, mQuery As CSectionGeometryQuery ' Точная область внутри оси для отбора продольных стержней.
Private mEdges() As TRegionEdge, mEdgeLengths() As Double
Private mArea As Double, mLength As Double, mMinimumCover As Double
Private mBoundaryEdges() As TRegionEdge, mBoundaryIsOpening() As Boolean, mHasOpening As Boolean

' Строит ось внутрь единственного наружного контура. Защитный слой
' измерен до поверхности хомута, все линейные размеры в мм.
Public Sub InitializeAutomatic(ByVal section As CSectionModel, ByVal clearCover As Double, ByVal diameter As Double)
    RequireEmpty
    ValidateDimensions clearCover, diameter
    Dim query As New CSectionGeometryQuery, domain As CConcreteRegion, outerIndex As Long
    Set domain = ConcreteDomain(section, query, outerIndex)
    Dim boundary As CGeometryPath, axis As CGeometryPath
    Set boundary = PathFromLoop(domain, outerIndex, "Torsion.Outer")
    On Error GoTo OffsetFailed
    Set axis = query.OffsetPath("Torsion.Hoop", boundary, clearCover + 0.5 * diameter) ' Бетон лежит слева от ориентированного наружного контура.
    On Error GoTo 0
    StoreCheckedPath query, domain, axis, clearCover, diameter
    Exit Sub
OffsetFailed:
    If Err.Number = 18 Then Err.Raise Err.Number, Err.Source, Err.Description
    Dim detail As String: detail = Err.Description
    HoopError "Не удалось построить одну ось хомута по защитному слою. " & detail & _
        " При наличии готовой оси задайте её явно."
End Sub

' Принимает фактическую, уже замкнутую ось. Без заданного защитного
' слоя проверяется принадлежность бетону всего сечения стержня.
Public Sub InitializeExplicit(ByVal section As CSectionModel, ByVal centerline As CGeometryPath, _
        ByVal diameter As Double, Optional ByVal clearCover As Double = 0#)
    RequireEmpty
    ValidateDimensions clearCover, diameter
    If centerline Is Nothing Then HoopError "Укажите замкнутую ось хомута из отрезков и дуг."
    If Not centerline.IsSealed Then HoopError "Завершите формирование оси хомута."
    If Not centerline.IsClosed Then HoopError "Ось хомута должна быть замкнута. Проверьте совпадение её концов."
    Dim query As New CSectionGeometryQuery, domain As CConcreteRegion, outerIndex As Long
    Set domain = ConcreteDomain(section, query, outerIndex)
    StoreCheckedPath query, domain, centerline, clearCover, diameter
End Sub

' Принимает точную ось из импортированного 2D-контура без отверстий.
' Тот же путь проверки используется для явно заданных отрезков и дуг.
Public Sub InitializeExplicitRegion(ByVal section As CSectionModel, ByVal axisRegion As CConcreteRegion, _
        ByVal diameter As Double, Optional ByVal clearCover As Double = 0#)
    RequireEmpty
    If axisRegion Is Nothing Then HoopError "Импортированный контур оси хомута отсутствует."
    If axisRegion.LoopCount <> 1 Then HoopError "Импортируйте одну замкнутую ось хомута."
    If axisRegion.LoopIsOpening(1) Then HoopError "Импортированный контур оси хомута задан как внутреннее отверстие."
    Dim axis As CGeometryPath: Set axis = PathFromLoop(axisRegion, 1, "Torsion.Hoop")
    InitializeExplicit section, axis, diameter, clearCover
End Sub

' Возвращает неизменяемую проверенную ось. Начало соответствует нулевой станции.
Public Property Get Centerline() As CGeometryPath
    RequireReady
    Set Centerline = mPath
End Property

' Возвращает охваченную осью площадь, включая вложенные пустоты, мм2.
Public Property Get EnclosedArea() As Double
    RequireReady
    EnclosedArea = mArea
End Property

' Возвращает полный периметр оси, мм.
Public Property Get Length() As Double
    RequireReady
    Length = mLength
End Property

' Возвращает наименьший зазор от поверхности стержня до любой границы бетона, мм.
Public Property Get MinimumClearCover() As Double
    RequireReady
    MinimumClearCover = mMinimumCover
End Property

' Отличает продольный стержень внутри контура от точки на самой оси.
Friend Function ContainsInside(ByVal x As Double, ByVal y As Double) As Boolean
    RequireReady
    ContainsInside = mQuery.ContainsPoint(mEnclosed, x, y, False)
End Function

' Относит продольную арматуру к наружному ряду по ближайшей бетонной
' границе. У отверстия арматура не включается даже при широкой полосе отбора.
Friend Function IsOuterRowPoint(ByVal x As Double, ByVal y As Double) As Boolean
    RequireReady
    If Not mHasOpening Then IsOuterRowPoint = True: Exit Function
    Dim outerGap As Double, openingGap As Double, gap As Double, i As Long
    Dim parameter As Double, unused As Boolean
    outerGap = 1E+100: openingGap = 1E+100
    For i = LBound(mBoundaryEdges) To UBound(mBoundaryEdges)
        gap = PointGap(mBoundaryEdges(i), x, y, parameter, unused)
        If mBoundaryIsOpening(i) Then
            openingGap = GeomMin(openingGap, gap)
        Else
            outerGap = GeomMin(outerGap, gap)
        End If
    Next i
    IsOuterRowPoint = (outerGap + GEOM_BOUNDARY_TOLERANCE < openingGap)
End Function

' Проецирует точку на ближайшую станцию оси. Два равноудалённых следа
' с разных сторон общего угла заменяются серединой короткого участка.
' Неразрешённые совпадения отмечаются как неоднозначные.
Public Function TryProjectStation(ByVal x As Double, ByVal y As Double, ByRef station As Double, _
        ByRef distance As Double, Optional ByRef ambiguous As Boolean = False) As Boolean
    RequireReady
    station = 0#: distance = 0#: ambiguous = False
    If Not FiniteCoordinate(x) Or Not FiniteCoordinate(y) Then Exit Function
    Dim i As Long, endStation As Double, candidate As Double, gap As Double, parameter As Double
    Dim bestEdge As Long, cornerEdge As Long, cornerStation As Double
    Dim alternative As Boolean, stationGap As Double
    For i = LBound(mEdges) To UBound(mEdges)
        gap = PointGap(mEdges(i), x, y, parameter, alternative)
        candidate = endStation + parameter * mEdgeLengths(i)
        If candidate = mLength Then candidate = 0# ' На шве конец имеет ту же станцию, что и начало.
        If i = LBound(mEdges) Or gap < distance - GEOM_BOUNDARY_TOLERANCE Then
            distance = gap: station = candidate: ambiguous = alternative
            bestEdge = i: cornerEdge = 0
        ElseIf Abs(gap - distance) <= GEOM_BOUNDARY_TOLERANCE Then
            stationGap = Abs(candidate - station)
            If alternative Then ambiguous = True
            If GeomMin(stationGap, mLength - stationGap) > GEOM_BOUNDARY_TOLERANCE Then
                If cornerEdge = 0 Then
                    cornerEdge = i: cornerStation = candidate
                ElseIf GeomMin(Abs(candidate - cornerStation), mLength - Abs(candidate - cornerStation)) > GEOM_BOUNDARY_TOLERANCE Then
                    ambiguous = True
                End If
            End If
            If gap < distance Then distance = gap
        End If
        endStation = endStation + mEdgeLengths(i)
    Next i
    If cornerEdge > 0 And Not ambiguous Then
        ambiguous = Not ResolveCornerStation(bestEdge, cornerEdge, distance, cornerStation, station)
    End If
    TryProjectStation = True
End Function

' Устраняет зависимость углового стержня от начальной точки и направления
' обхода. Удалённые равные проекции остаются неоднозначными.
Private Function ResolveCornerStation(ByVal firstEdge As Long, ByVal secondEdge As Long, ByVal distance As Double, _
        ByVal secondStation As Double, ByRef firstStation As Double) As Boolean
    Dim edgeGap As Long, stationGap As Double
    edgeGap = secondEdge - firstEdge
    If edgeGap <> 1 And edgeGap <> UBound(mEdges) - 1 Then Exit Function
    stationGap = secondStation - firstStation
    If stationGap > 0.5 * mLength Then stationGap = stationGap - mLength
    If stationGap < -0.5 * mLength Then stationGap = stationGap + mLength
    If Abs(stationGap) > 4# * distance + GEOM_BOUNDARY_TOLERANCE Then Exit Function ' Средняя станция относится к одному локальному углу.
    firstStation = firstStation + 0.5 * stationGap
    If firstStation < 0# Then firstStation = firstStation + mLength
    If firstStation >= mLength Then firstStation = firstStation - mLength
    ResolveCornerStation = True
End Function

' Получает точный бетон с отверстиями из исходных контуров сечения.
' Один хомут обслуживает один наружный контур; сетка положение оси не задаёт.
Private Function ConcreteDomain(ByVal section As CSectionModel, ByVal query As CSectionGeometryQuery, _
        ByRef outerIndex As Long) As CConcreteRegion
    If section Is Nothing Then HoopError "Не передана геометрия сечения для хомута."
    Set ConcreteDomain = query.RegionFromContours(section.Contours)
    Dim i As Long
    For i = 1 To ConcreteDomain.LoopCount
        If Not ConcreteDomain.LoopIsOpening(i) Then
            If outerIndex > 0 Then HoopError "У сечения несколько отдельных наружных контуров. Для этой геометрии один хомут неприменим."
            outerIndex = i
        End If
    Next i
    If outerIndex = 0 Then HoopError "Для хомута нужен точный наружный контур бетона."
End Function

' Переносит точно ориентированные сегменты выбранной петли из бетонной области.
Private Function PathFromLoop(ByVal domain As CConcreteRegion, ByVal loopIndex As Long, ByVal pathID As String) As CGeometryPath
    Dim edges() As TRegionEdge, count As Long, i As Long
    domain.CopyEdges edges, count
    Dim path As New CGeometryPath: path.Initialize pathID
    For i = 1 To count
        If edges(i).LoopID = loopIndex Then path.Append pathID & "." & CStr(i), edges(i)
    Next i
    path.Seal True
    Set PathFromLoop = path
End Function

' Проверяет замыкание, простоту оси и точный минимальный зазор между всеми
' парами сегментов. Площадь, внутри которой допустимы отверстия, вычисляется отдельно.
Private Sub StoreCheckedPath(ByVal query As CSectionGeometryQuery, ByVal domain As CConcreteRegion, _
        ByVal axis As CGeometryPath, ByVal clearCover As Double, ByVal diameter As Double)
    Dim enclosed As CConcreteRegion: Set enclosed = query.RegionFromPath(axis)
    If enclosed.LoopCount <> 1 Then HoopError "Ось хомута должна образовывать один замкнутый контур без самопересечений."
    Dim boundaries() As TRegionEdge, count As Long, i As Long, j As Long
    domain.CopyEdges boundaries, count
    ReDim mBoundaryEdges(1 To count)
    ReDim mBoundaryIsOpening(1 To count)
    For i = 1 To count
        LSet mBoundaryEdges(i) = boundaries(i)
        mBoundaryIsOpening(i) = domain.LoopIsOpening(boundaries(i).LoopID)
        If mBoundaryIsOpening(i) Then mHasOpening = True
    Next i
    Dim edge As TRegionEdge, distance As Double, minimum As Double
    Dim limitingSegment As Long, limitingLoop As Long
    axis.CopySegment 1, edge
    If Not query.ContainsPoint(domain, edge.X1, edge.Y1, False) Then _
        HoopError "Ось хомута находится вне бетона или в отверстии."
    minimum = 1E+100
    For i = 1 To axis.SegmentCount
        axis.CopySegment i, edge
        For j = 1 To count
            distance = SegmentGap(edge, boundaries(j))
            If distance < minimum Then
                minimum = distance: limitingSegment = i: limitingLoop = boundaries(j).LoopID
            End If
        Next j
    Next i
    Dim requiredDistance As Double: requiredDistance = clearCover + 0.5 * diameter
    If minimum + GEOM_BOUNDARY_TOLERANCE < requiredDistance Then _
        HoopError "Участок оси " & CStr(limitingSegment) & " слишком близок к " & BoundaryCaption(domain, limitingLoop) & _
            ". Зазор от оси " & Format$(minimum, "0.###") & " мм, требуется " & Format$(requiredDistance, "0.###") & _
            " мм. Проверьте защитный слой, диаметр и положение оси хомута."
    ReDim mEdges(1 To axis.SegmentCount), mEdgeLengths(1 To axis.SegmentCount)
    For i = 1 To axis.SegmentCount
        axis.CopySegment i, mEdges(i)
        mEdgeLengths(i) = modGeometryPrimitives.EdgeLength(mEdges(i))
    Next i
    mArea = enclosed.Area: mLength = axis.Length: mMinimumCover = GeomMax(0#, minimum - 0.5 * diameter)
    Set mEnclosed = enclosed: Set mQuery = query
    Set mPath = axis
End Sub

' Минимум расстояния между двумя аналитическими сегментами. При отсутствии
' пересечения проверяет концы и внутренние пары с общими нормалями.
Private Function SegmentGap(ByRef first As TRegionEdge, ByRef second As TRegionEdge) As Double
    Dim intersections() As Double, count As Long
    modGeometryPrimitives.EdgeIntersectionCuts first, second, intersections, count
    If count > 0 Then Exit Function ' При пересечении или касании зазор равен нулю.
    Dim parameter As Double, unused As Boolean, minimum As Double
    minimum = PointGap(first, second.X1, second.Y1, parameter, unused)
    minimum = GeomMin(minimum, PointGap(first, second.X2, second.Y2, parameter, unused))
    minimum = GeomMin(minimum, PointGap(second, first.X1, first.Y1, parameter, unused))
    minimum = GeomMin(minimum, PointGap(second, first.X2, first.Y2, parameter, unused))
    If IsArc(first) Then
        If IsArc(second) Then
            minimum = ArcNormalGap(first, second, second.CenterX - first.CenterX, second.CenterY - first.CenterY, minimum)
        Else
            minimum = ArcNormalGap(first, second, second.Y1 - second.Y2, second.X2 - second.X1, minimum)
        End If
    ElseIf IsArc(second) Then
        minimum = ArcNormalGap(second, first, first.Y1 - first.Y2, first.X2 - first.X1, minimum)
    End If
    SegmentGap = minimum
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
            ArcNormalGap = GeomMin(ArcNormalGap, PointGap(other, x, y, parameter, unused))
    Next side
End Function

' Проецирует точку на ограниченный сегмент и возвращает расстояние.
' Для дуги рассматривает радиальную проекцию и оба конца.
Private Function PointGap(ByRef edge As TRegionEdge, ByVal x As Double, ByVal y As Double, _
        ByRef parameter As Double, ByRef ambiguous As Boolean) As Double
    Dim dx As Double, dy As Double, nearX As Double, nearY As Double
    ambiguous = False
    If Not IsArc(edge) Then
        dx = edge.X2 - edge.X1: dy = edge.Y2 - edge.Y1
        parameter = GeomMax(0#, GeomMin(1#, ((x - edge.X1) * dx + (y - edge.Y1) * dy) / (dx * dx + dy * dy)))
        PointGap = DistanceForVector(x - edge.X1 - parameter * dx, y - edge.Y1 - parameter * dy)
        Exit Function
    End If
    dx = x - edge.CenterX: dy = y - edge.CenterY
    Dim radial As Double, radius As Double
    radial = DistanceForVector(dx, dy): radius = DistanceForVector(edge.X1 - edge.CenterX, edge.Y1 - edge.CenterY)
    If radial = 0# Then
        parameter = 0#: ambiguous = True: PointGap = radius
        Exit Function
    End If
    nearX = edge.CenterX + radius * (dx / radial): nearY = edge.CenterY + radius * (dy / radial)
    If modGeometryPrimitives.PointParameter(edge, nearX, nearY, parameter) Then
        PointGap = Abs(radial - radius)
    Else
        Dim firstGap As Double, lastGap As Double
        firstGap = DistanceForVector(x - edge.X1, y - edge.Y1): lastGap = DistanceForVector(x - edge.X2, y - edge.Y2)
        ambiguous = Abs(firstGap - lastGap) <= GEOM_BOUNDARY_TOLERANCE
        If firstGap <= lastGap Then
            parameter = 0#: PointGap = firstGap
        Else
            parameter = 1#: PointGap = lastGap
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

' Называет границу, которая ограничивает защитный слой.
Private Function BoundaryCaption(ByVal region As CConcreteRegion, ByVal loopIndex As Long) As String
    If region.LoopIsOpening(loopIndex) Then
        BoundaryCaption = "границе отверстия №" & CStr(loopIndex)
    Else
        BoundaryCaption = "наружной границе"
    End If
End Function

' Проверяет размеры до смещения и вычисления зазоров.
Private Sub ValidateDimensions(ByVal clearCover As Double, ByVal diameter As Double)
    If Not FiniteCoordinate(clearCover) Or clearCover < 0# Then HoopError "Укажите конечный неотрицательный защитный слой хомута."
    If Not FiniteCoordinate(diameter) Or 0.5 * diameter <= GEOM_BOUNDARY_TOLERANCE Then _
        HoopError "Укажите положительный диаметр хомута больше геометрического допуска."
End Sub

' Ограничивает значения диапазоном, пригодным для квадратичных операций.
Private Function FiniteCoordinate(ByVal value As Double) As Boolean
    FiniteCoordinate = (value = value And Abs(value) < 1E+100)
End Function

' Проверяет, что экземпляр ещё можно инициализировать.
Private Sub RequireEmpty()
    If Not mPath Is Nothing Then HoopError "Геометрия хомута уже подготовлена."
End Sub

' Проверяет готовность расчётной геометрии.
Private Sub RequireReady()
    If mPath Is Nothing Then HoopError "Геометрия хомута ещё не подготовлена."
End Sub

' Передаёт вызывающему адресуемую ошибку геометрии без обращения к Excel.
Private Sub HoopError(ByVal reason As String)
    Err.Raise vbObjectError + 5660, "CTorsionHoopGeometry", reason
End Sub
