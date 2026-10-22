Attribute VB_Name = "modGeometryPrimitives"
Option Explicit
Option Private Module

' Вычисляет длины, точки, касательные и пересечения отрезков и дуг.
' CSectionGeometryQuery и CGeometryPath используют одну общую реализацию
' этих формул без обращения к Excel и конкретному сечению.
Public Const GEOM_BOUNDARY_TOLERANCE As Double = 0.000001 ' Допуск границ и совпадения координат, мм.
Public Const GEOM_PARAMETER_TOLERANCE As Double = 0.0000000001 ' Допуск безразмерного параметра сегмента.
Private Const QUERY_TOLERANCE As Double = GEOM_BOUNDARY_TOLERANCE
Private Const PARAM_TOLERANCE As Double = GEOM_PARAMETER_TOLERANCE

' Общие экстремумы аналитического сегмента для области и открытого пути.
' Направление уже единичное; проверяются также внутренние экстремумы дуги.
Public Sub AccumulateEdgeProjection(ByRef edge As TRegionEdge, ByVal nx As Double, ByVal ny As Double, _
        ByRef minimum As Double, ByRef maximum As Double)
    Dim radius As Double, u As Double, projection As Double
    projection = nx * edge.X1 + ny * edge.Y1
    minimum = GeomMin(minimum, projection): maximum = GeomMax(maximum, projection)
    projection = nx * edge.X2 + ny * edge.Y2
    minimum = GeomMin(minimum, projection): maximum = GeomMax(maximum, projection)
    If Abs(edge.Sweep) > PARAM_TOLERANCE Then
        radius = Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2)
        projection = nx * edge.CenterX + ny * edge.CenterY
        If PointParameter(edge, edge.CenterX + radius * nx, edge.CenterY + radius * ny, u) Then maximum = GeomMax(maximum, projection + radius)
        If PointParameter(edge, edge.CenterX - radius * nx, edge.CenterY - radius * ny, u) Then minimum = GeomMin(minimum, projection - radius)
    End If
End Sub

' Ближайшее пересечение луча с ограниченным сегментом, а не с полной прямой
' или окружностью. Используется тот же аналитический поиск, что при clipping.
Public Function EdgeRayHit(ByRef edge As TRegionEdge, ByVal x As Double, ByVal y As Double, _
        ByVal dx As Double, ByVal dy As Double, ByRef distance As Double, ByRef parameter As Double) As Boolean
    Dim norm As Double, cuts() As Double, count As Long, i As Long, px As Double, py As Double, candidate As Double
    norm = Sqr(dx * dx + dy * dy)
    If norm <= GEOM_TOLERANCE Then GeometryError "Направление луча имеет нулевую длину."
    dx = dx / norm: dy = dy / norm
    distance = 0#: parameter = 0#
    If PointParameter(edge, x, y, parameter) Then
        EdgeRayHit = True
        Exit Function
    End If
    HalfPlaneCuts edge, -dy, dx, -dy * x + dx * y, cuts, count ' Прямая луча задаётся нормалью (-dy, dx); отрицательные расстояния отсеиваются ниже.
    For i = 1 To count
        EdgePoint edge, cuts(i), px, py
        candidate = (px - x) * dx + (py - y) * dy
        If candidate >= -QUERY_TOLERANCE Then
            If Not EdgeRayHit Or candidate < distance Then
                distance = GeomMax(0#, candidate): parameter = cuts(i): EdgeRayHit = True
            End If
        End If
    Next i
End Function

' Добавляет параметры пересечения сегмента с бесконечной линией n*P=bound.
' Касание дает одну точку; совпадающий отрезок добавляет оба конца.
Public Sub HalfPlaneCuts(ByRef edge As TRegionEdge, ByVal nx As Double, ByVal ny As Double, _
        ByVal bound As Double, ByRef cuts() As Double, ByRef count As Long)
    Dim f1 As Double, f2 As Double, u As Double
    If Abs(edge.Sweep) <= PARAM_TOLERANCE Then
        f1 = nx * edge.X1 + ny * edge.Y1 - bound
        f2 = nx * edge.X2 + ny * edge.Y2 - bound
        If Abs(f1) <= QUERY_TOLERANCE Then AddScalar cuts, count, 0#
        If Abs(f2) <= QUERY_TOLERANCE Then AddScalar cuts, count, 1#
        If Abs(f1 - f2) > GEOM_TOLERANCE Then
            u = f1 / (f1 - f2)
            If u > PARAM_TOLERANCE And u < 1# - PARAM_TOLERANCE Then AddScalar cuts, count, u
        End If
    Else
        Dim radius As Double, norm As Double, distance As Double, bx As Double, by As Double, offset As Double
        Dim roundoff As Double
        radius = Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2)
        norm = Sqr(nx * nx + ny * ny)
        distance = (bound - nx * edge.CenterX - ny * edge.CenterY) / norm
        If Abs(distance) > radius + QUERY_TOLERANCE Then Exit Sub
        roundoff = 3.5527136788005E-15 * _
            ((Abs(bound) + Abs(nx * edge.CenterX) + Abs(ny * edge.CenterY)) / norm + radius + 1#)
        If Abs(Abs(distance) - radius) <= roundoff Then distance = Sgn(distance) * radius
        If Abs(distance) > radius Then distance = Sgn(distance) * radius ' Допуск допускает касание снаружи, но не схлопывает настоящую секущую изнутри: даже малая глубина дает два разных конца дуги.
        bx = edge.CenterX + nx / norm * distance: by = edge.CenterY + ny / norm * distance
        offset = Sqr(GeomMax((radius - Abs(distance)) * (radius + Abs(distance)), 0#))
        If PointParameter(edge, bx - ny / norm * offset, by + nx / norm * offset, u) Then AddScalar cuts, count, u
        If PointParameter(edge, bx + ny / norm * offset, by - nx / norm * offset, u) Then AddScalar cuts, count, u
    End If
End Sub

' Вычисляет пересечения пары аналитических сегментов, включая совпадающие
' линии/дуги. Параметры добавляются только на первом сегменте для его разбиения.
Public Sub EdgeIntersectionCuts(ByRef a As TRegionEdge, ByRef b As TRegionEdge, _
        ByRef cuts() As Double, ByRef count As Long)
    Dim u As Double, v As Double, dx As Double, dy As Double, denominator As Double, x As Double, y As Double
    If PointParameter(a, b.X1, b.Y1, u) Then AddScalar cuts, count, u
    If PointParameter(a, b.X2, b.Y2, u) Then AddScalar cuts, count, u
    If Abs(a.Sweep) <= PARAM_TOLERANCE And Abs(b.Sweep) <= PARAM_TOLERANCE Then
        dx = a.X2 - a.X1: dy = a.Y2 - a.Y1
        denominator = dx * (b.Y2 - b.Y1) - dy * (b.X2 - b.X1)
        If Abs(denominator) <= GEOM_TOLERANCE Then Exit Sub
        u = ((b.X1 - a.X1) * (b.Y2 - b.Y1) - (b.Y1 - a.Y1) * (b.X2 - b.X1)) / denominator
        If u < -PARAM_TOLERANCE Or u > 1# + PARAM_TOLERANCE Then Exit Sub
        EdgePoint a, u, x, y
        If PointParameter(b, x, y, v) Then AddScalar cuts, count, GeomMax(0#, GeomMin(1#, u))
    ElseIf Abs(a.Sweep) <= PARAM_TOLERANCE Or Abs(b.Sweep) <= PARAM_TOLERANCE Then
        Dim arc As TRegionEdge, line As TRegionEdge, arcCuts() As Double, arcCount As Long, i As Long
        If Abs(a.Sweep) > PARAM_TOLERANCE Then
            LSet arc = a: LSet line = b
        Else
            LSet arc = b: LSet line = a
        End If
        dx = line.X2 - line.X1: dy = line.Y2 - line.Y1
        HalfPlaneCuts arc, -dy, dx, -dy * line.X1 + dx * line.Y1, arcCuts, arcCount
        For i = 1 To arcCount
            EdgePoint arc, arcCuts(i), x, y
            If PointParameter(line, x, y, v) And PointParameter(a, x, y, u) Then AddScalar cuts, count, u
        Next i
    Else
        Dim r1 As Double, r2 As Double, distance As Double, along As Double, height As Double, bx As Double, by As Double
        r1 = Sqr((a.X1 - a.CenterX) ^ 2 + (a.Y1 - a.CenterY) ^ 2)
        r2 = Sqr((b.X1 - b.CenterX) ^ 2 + (b.Y1 - b.CenterY) ^ 2)
        dx = b.CenterX - a.CenterX: dy = b.CenterY - a.CenterY: distance = Sqr(dx * dx + dy * dy)
        If distance <= QUERY_TOLERANCE Then Exit Sub
        If distance > r1 + r2 + QUERY_TOLERANCE Or distance < Abs(r1 - r2) - QUERY_TOLERANCE Then Exit Sub
        along = (r1 * r1 - r2 * r2 + distance * distance) / (2# * distance) ' Расстояние от центра первой окружности до основания их общей хорды, мм.
        height = Sqr(GeomMax(r1 * r1 - along * along, 0#)) ' Половина длины общей хорды; отрицательный остаток округления при касании обнуляется.
        bx = a.CenterX + along * dx / distance: by = a.CenterY + along * dy / distance
        x = bx - height * dy / distance: y = by + height * dx / distance
        If PointParameter(a, x, y, u) And PointParameter(b, x, y, v) Then AddScalar cuts, count, u
        x = bx + height * dy / distance: y = by - height * dx / distance
        If PointParameter(a, x, y, u) And PointParameter(b, x, y, v) Then AddScalar cuts, count, u
    End If
End Sub

' Проверяет принадлежность точки аналитическому сегменту и возвращает u.
' У дуг angle-wrap учитывается по signed sweep, без зависимости от глобальных осей.
Public Function PointParameter(ByRef edge As TRegionEdge, ByVal x As Double, ByVal y As Double, ByRef u As Double) As Boolean
    Dim dx As Double, dy As Double, lengthSquared As Double, px As Double, py As Double
    If Abs(edge.Sweep) <= PARAM_TOLERANCE Then
        dx = edge.X2 - edge.X1: dy = edge.Y2 - edge.Y1: lengthSquared = dx * dx + dy * dy
        If lengthSquared <= QUERY_TOLERANCE * QUERY_TOLERANCE Then Exit Function
        u = ((x - edge.X1) * dx + (y - edge.Y1) * dy) / lengthSquared
        If u < -PARAM_TOLERANCE Or u > 1# + PARAM_TOLERANCE Then Exit Function
        px = edge.X1 + u * dx: py = edge.Y1 + u * dy
        PointParameter = SamePoint(px, py, x, y)
    Else
        Dim radius As Double, pointRadius As Double, startAngle As Double, pointAngle As Double, delta As Double
        radius = Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2)
        pointRadius = Sqr((x - edge.CenterX) ^ 2 + (y - edge.CenterY) ^ 2)
        If Abs(radius - pointRadius) > QUERY_TOLERANCE Then Exit Function
        startAngle = AngleOf(edge.X1 - edge.CenterX, edge.Y1 - edge.CenterY)
        pointAngle = AngleOf(x - edge.CenterX, y - edge.CenterY)
        delta = pointAngle - startAngle ' Угол от начала дуги затем приводится к направлению её знакового угла Sweep.
        If edge.Sweep > 0# Then
            Do While delta < -PARAM_TOLERANCE: delta = delta + 2# * GEOM_PI: Loop
        Else
            Do While delta > PARAM_TOLERANCE: delta = delta - 2# * GEOM_PI: Loop
        End If
        u = delta / edge.Sweep
        PointParameter = (u >= -PARAM_TOLERANCE And u <= 1# + PARAM_TOLERANCE)
    End If
    If PointParameter Then u = GeomMax(0#, GeomMin(1#, u))
End Function

' Возвращает точку аналитического сегмента по u от нуля до единицы.
Public Sub EdgePoint(ByRef edge As TRegionEdge, ByVal u As Double, ByRef x As Double, ByRef y As Double)
    If Abs(edge.Sweep) <= PARAM_TOLERANCE Then
        x = edge.X1 + u * (edge.X2 - edge.X1): y = edge.Y1 + u * (edge.Y2 - edge.Y1)
    Else
        Dim angle As Double, radius As Double
        angle = AngleOf(edge.X1 - edge.CenterX, edge.Y1 - edge.CenterY) + u * edge.Sweep
        radius = Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2)
        x = edge.CenterX + radius * Cos(angle): y = edge.CenterY + radius * Sin(angle)
    End If
End Sub

' Возвращает касательную в направлении обхода; ее длина не нормируется.
Public Sub EdgeTangent(ByRef edge As TRegionEdge, ByVal u As Double, ByRef tx As Double, ByRef ty As Double)
    If Abs(edge.Sweep) <= PARAM_TOLERANCE Then
        tx = edge.X2 - edge.X1: ty = edge.Y2 - edge.Y1
    Else
        Dim x As Double, y As Double
        EdgePoint edge, u, x, y
        tx = -(y - edge.CenterY) * edge.Sweep: ty = (x - edge.CenterX) * edge.Sweep
    End If
End Sub

' Выделяет точный подотрезок/поддугу после пересечения; центр и signed sweep
' остаются аналитическими, поэтому площадь не зависит от плотности рисунка.
Public Sub SubEdge(ByRef original As TRegionEdge, ByVal u1 As Double, ByVal u2 As Double, ByRef result As TRegionEdge)
    LSet result = original
    EdgePoint original, u1, result.X1, result.Y1
    EdgePoint original, u2, result.X2, result.Y2
    result.Sweep = original.Sweep * (u2 - u1): result.LoopID = 0 ' Новый участок сохраняет дугу, но ещё не принадлежит собранному замкнутому контуру.
End Sub

' Обращает сегмент вместе со знаком дуги, сохраняя ту же физическую границу.
Public Sub ReverseEdge(ByRef edge As TRegionEdge)
    Dim value As Double
    value = edge.X1: edge.X1 = edge.X2: edge.X2 = value
    value = edge.Y1: edge.Y1 = edge.Y2: edge.Y2 = value
    edge.Sweep = -edge.Sweep
End Sub

' Возвращает atan2 в диапазоне -pi..pi без обращения к Excel.
Public Function AngleOf(ByVal x As Double, ByVal y As Double) As Double
    If Abs(x) <= GEOM_TOLERANCE Then
        AngleOf = Sgn(y) * GEOM_PI / 2#
    Else
        AngleOf = Atn(y / x)
        If x < 0# Then
            If y >= 0# Then AngleOf = AngleOf + GEOM_PI Else AngleOf = AngleOf - GEOM_PI
        End If
    End If
End Function

' Возвращает угол линии для читаемой подписи, в радианах. Направление обхода
' не переворачивает текст вверх ногами; вертикальная подпись читается снизу вверх.
Public Function ReadableLineAngle(ByVal directionX As Double, ByVal directionY As Double) As Double
    Dim angle As Double
    angle = AngleOf(directionX, directionY)
    If angle > GEOM_PI / 2# Then angle = angle - GEOM_PI
    If angle <= -GEOM_PI / 2# Then angle = angle + GEOM_PI
    ReadableLineAngle = angle
End Function

' Возвращает физическую длину сегмента для исключения вырожденных ребер.
Public Function EdgeLength(ByRef edge As TRegionEdge) As Double
    If Abs(edge.Sweep) <= PARAM_TOLERANCE Then
        EdgeLength = Sqr((edge.X2 - edge.X1) ^ 2 + (edge.Y2 - edge.Y1) ^ 2)
    Else
        EdgeLength = Abs(edge.Sweep) * Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2)
    End If
End Function

' Добавляет численный параметр в список, если такого значения ещё нет.
' Значения сравниваются по абсолютному допуску параметра сегмента;
' размерный допуск координат проверяет вызывающая операция пересечения.
Public Sub AddScalar(ByRef values() As Double, ByRef count As Long, ByVal value As Double)
    Dim i As Long
    For i = 1 To count
        If Abs(values(i) - value) <= PARAM_TOLERANCE Then Exit Sub
    Next i
    If count = 0 Then ReDim values(1 To 8)
    If count = UBound(values) Then ReDim Preserve values(1 To UBound(values) * 2)
    count = count + 1: values(count) = value
End Sub

' Централизует допуск совпадения концов, вершин и принадлежности границе.
Public Function SamePoint(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Boolean
    SamePoint = (Abs(x1 - x2) <= QUERY_TOLERANCE And Abs(y1 - y2) <= QUERY_TOLERANCE)
End Function


' Отклоняет нулевое направление прямой. Команды линии и геометрические запросы
' обычно проверяют его заранее, но прямой вызов этой операции также защищён.
Private Sub GeometryError(ByVal reason As String)
    Err.Raise vbObjectError + 5401, "modGeometryPrimitives", reason
End Sub
