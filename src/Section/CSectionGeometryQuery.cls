VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionGeometryQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Общие геометрические операции над областью бетона
' ==========================================================================
' Работает с точными отрезками и дугами и с границами бетонных элементов.
' Удаляет общие внутренние рёбра и соединяет оставшиеся в наружные контуры
' и отверстия. Геометрия полученной области и её площадь создаются вместе.
' Расчёт трещин передаёт сюда геометрические ограничения; нормативные
' коэффициенты и группы арматуры определяются вызывающим расчётом.
' Прежний расчёт характеристик сетки для устойчивости сохраняется отдельно.

Private Const QUERY_TOLERANCE As Double = GEOM_BOUNDARY_TOLERANCE ' Допуск соединения соседних участков бетона вдоль одной прямой, мм.
Private Const PARAM_TOLERANCE As Double = GEOM_PARAMETER_TOLERANCE ' Допуск безразмерного положения точки внутри отрезка или дуги.
Private mSection As CSectionModel ' Модель текущего запроса; операции над переданными областями не требуют модели.
Private mRevision As Long ' Версия геометрии, для которой подготовлена бетонная область.
Private mDomain As CConcreteRegion ' Подготовленная область текущей модели; при выдаче создаётся независимая копия.
Private mDomainBuildCount As Long ' Число успешных построений области этим объектом, включая смену модели и её версии.
Private mDiagnostic As CGeometryDiagnostic ' Последняя ошибка общей геометрической операции с её исходной привязкой.

' Последняя машинная диагностика; Nothing при отсутствии записанного события.
Public Property Get LastDiagnostic() As CGeometryDiagnostic
    Set LastDiagnostic = mDiagnostic
End Property

' Отражает, поворачивает и переносит точку так же, как преобразуются сегменты и привязки линии.
Public Sub TransformCoordinates(ByRef x As Double, ByRef y As Double, ByVal shiftX As Double, ByVal shiftY As Double, _
        Optional ByVal rotation As Double = 0#, Optional ByVal reflectX As Boolean = False)
    Dim newX As Double, cosine As Double, sine As Double
    cosine = Cos(rotation): sine = Sin(rotation)
    If reflectX Then y = -y
    newX = cosine * x - sine * y + shiftX
    y = sine * x + cosine * y + shiftY: x = newX
End Sub

' Строит линию на заданном расстоянии от исходной: положительное смещение
' направлено влево по обходу. Гладкие стыки отрезков и дуг сохраняют касание,
' соседние прямые соединяются пересечением продолжений. Исчезнувшее скругление
' между касательными прямыми заменяется их стыком. Для негладкого стыка
' с дугой требуется явно заданный скруглённый многоугольник RoundedPolygon.
Public Function OffsetPath(ByVal pathID As String, ByVal source As CGeometryPath, ByVal distance As Double) As CGeometryPath
    If source Is Nothing Then GeometryError "Не передан исходный путь.", "MISSING_PATH", pathID
    Dim edges() As TRegionEdge, original As TRegionEdge, i As Long, following As Long, count As Long
    Dim radiusValue As Double, radialScale As Double, tx As Double, ty As Double, length As Double, result As CGeometryPath
    count = source.SegmentCount: ReDim edges(1 To count)
    Dim removed() As Boolean: ReDim removed(1 To count)
    For i = 1 To count
        source.CopySegment i, edges(i)
        If edges(i).Sweep = 0# Then
            length = modGeometryPrimitives.EdgeLength(edges(i)): tx = (edges(i).X2 - edges(i).X1) / length: ty = (edges(i).Y2 - edges(i).Y1) / length
            edges(i).X1 = edges(i).X1 - ty * distance: edges(i).X2 = edges(i).X2 - ty * distance
            edges(i).Y1 = edges(i).Y1 + tx * distance: edges(i).Y2 = edges(i).Y2 + tx * distance
        Else
            radiusValue = Sqr((edges(i).X1 - edges(i).CenterX) ^ 2 + (edges(i).Y1 - edges(i).CenterY) ^ 2)
            radialScale = (radiusValue - Sgn(edges(i).Sweep) * distance) / radiusValue
            If radialScale <= QUERY_TOLERANCE / radiusValue Then
                RequireOffsetCorner pathID, source, i, radiusValue, distance
                removed(i) = True
            End If
            edges(i).X1 = edges(i).CenterX + (edges(i).X1 - edges(i).CenterX) * radialScale
            edges(i).Y1 = edges(i).CenterY + (edges(i).Y1 - edges(i).CenterY) * radialScale
            edges(i).X2 = edges(i).CenterX + (edges(i).X2 - edges(i).CenterX) * radialScale
            edges(i).Y2 = edges(i).CenterY + (edges(i).Y2 - edges(i).CenterY) * radialScale
        End If
    Next i
    For i = 1 To count
        If Not removed(i) Then
            following = i + 1
            Do While following <= count
                If Not removed(following) Then Exit Do
                following = following + 1
            Loop
            If following > count Then
                If Not source.IsClosed Then Exit For
                following = 1
                Do While removed(following): following = following + 1: Loop
            End If
            JoinOffsetEdges pathID, source.SegmentID(i), edges(i), edges(following)
        End If
    Next i
    Dim sourceFirst As Double, sourceLast As Double, lengthSquared As Double, span As Double, stations As Variant
    ReDim stations(0 To count)
    Set result = New CGeometryPath: result.Initialize pathID
    For i = 1 To count
        If removed(i) Then
            result.AddCollapsedSegment source.SegmentID(i)
        Else
            source.CopySegment i, original
            sourceFirst = source.SourceStartParameter(i): sourceLast = source.SourceEndParameter(i)
            If edges(i).Sweep = 0# Then
                If (edges(i).X2 - edges(i).X1) * (original.X2 - original.X1) + (edges(i).Y2 - edges(i).Y1) * (original.Y2 - original.Y1) <= QUERY_TOLERANCE Then _
                    GeometryError "После смещения у прямого участка не осталось положительной длины. Уменьшите расстояние смещения или увеличьте размеры исходного контура.", "OFFSET_DISAPPEARED", pathID, source.SegmentID(i), distance
                ' После соединения продолжений длина отрезка может измениться. Сохраняем
                ' соответствие новых концов параметрам исходного отрезка для правильных привязок.
                tx = original.X2 - original.X1: ty = original.Y2 - original.Y1
                lengthSquared = tx * tx + ty * ty: span = sourceLast - sourceFirst
                sourceLast = sourceFirst + span * ((edges(i).X2 - original.X1) * tx + (edges(i).Y2 - original.Y1) * ty) / lengthSquared
                sourceFirst = sourceFirst + span * ((edges(i).X1 - original.X1) * tx + (edges(i).Y1 - original.Y1) * ty) / lengthSquared
            End If
            result.Append source.SegmentID(i), edges(i), source.SourcePathID(i), source.SourceSegmentID(i), sourceFirst, sourceLast
        End If
        stations(i) = result.AccumulatedLength
    Next i
    source.CopyCollapsedAnchors result, 0#, source.Length, 0#, segmentStations:=stations
    Dim boundary As CGeometryPath
    If Not source.IsClosed Then
        Set boundary = source.OffsetBoundary
        If Not boundary Is Nothing Then result.SetOffsetBoundary OffsetPath(pathID & ".Boundary", boundary, distance * source.BoundaryDirection), source.BoundaryDirection
    End If
    result.Seal source.IsClosed
    ' Проверяем замкнутость и отсутствие недопустимых пересечений всего контура.
    ' Выбирать только один длинный участок вместо исправления ошибки нельзя.
    Dim region As CConcreteRegion
    If result.IsClosed Then Set region = RegionFromPath(result)
    Set OffsetPath = result
End Function

' Исчезнуть может скругление между двумя касательными прямыми. Прямые после
' смещения соединяются в углу; отдельная дуга и полный круг такой замены не имеют.
Private Sub RequireOffsetCorner(ByVal pathID As String, ByVal source As CGeometryPath, ByVal index As Long, _
        ByVal radius As Double, ByVal distance As Double)
    Dim before As Long, after As Long, corner As TRegionEdge, first As TRegionEdge, last As TRegionEdge
    before = index - 1: after = index + 1
    If source.IsClosed Then
        If before < 1 Then before = source.SegmentCount
        If after > source.SegmentCount Then after = 1
    End If
    If before < 1 Or after > source.SegmentCount Or before = index Or after = index Then GoTo NoCorner
    source.CopySegment index, corner: source.CopySegment before, first: source.CopySegment after, last
    If first.Sweep <> 0# Or last.Sweep <> 0# Or Abs(corner.Sweep) >= GEOM_PI - PARAM_TOLERANCE Then GoTo NoCorner
    Dim ax As Double, ay As Double, bx As Double, by As Double
    modGeometryPrimitives.EdgeTangent first, 1#, ax, ay
    modGeometryPrimitives.EdgeTangent corner, 0#, bx, by
    If Not SameOffsetTangent(ax, ay, bx, by) Then GoTo NoCorner
    modGeometryPrimitives.EdgeTangent corner, 1#, ax, ay
    modGeometryPrimitives.EdgeTangent last, 0#, bx, by
    If Not SameOffsetTangent(ax, ay, bx, by) Then GoTo NoCorner
    Exit Sub
NoCorner:
    If source.IsClosed And source.SegmentCount = 1 Then source.ValidateNormalTransfer 0#, distance
    GeometryError "После смещения дуга исчезла, а соседних касательных прямых для образования угла нет. Уменьшите расстояние смещения.", _
        "OFFSET_DISAPPEARED", pathID, source.SegmentID(index), radius, distance
End Sub

' Сравнивает направления касательных без зависимости от длины исходных отрезков.
Private Function SameOffsetTangent(ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double) As Boolean
    SameOffsetTangent = ax * bx + ay * by > 0# And _
        Abs(ax * by - ay * bx) <= PARAM_TOLERANCE * Sqr((ax * ax + ay * ay) * (bx * bx + by * by))
End Function

' Стыкует смещённые отрезки по пересечению их продолжений.
Private Sub JoinOffsetEdges(ByVal pathID As String, ByVal segmentID As String, ByRef first As TRegionEdge, ByRef second As TRegionEdge)
    If Sqr((first.X2 - second.X1) ^ 2 + (first.Y2 - second.Y1) ^ 2) <= QUERY_TOLERANCE Then Exit Sub
    If first.Sweep <> 0# Or second.Sweep <> 0# Then GeometryError "Негладкий стык с дугой: задайте согласованное скругление.", "OFFSET_JOINT", pathID, segmentID
    Dim ax As Double, ay As Double, bx As Double, by As Double, cross As Double, parameter As Double
    ax = first.X2 - first.X1: ay = first.Y2 - first.Y1: bx = second.X2 - second.X1: by = second.Y2 - second.Y1
    cross = ax * by - ay * bx
    If Abs(cross) <= PARAM_TOLERANCE * Sqr((ax * ax + ay * ay) * (bx * bx + by * by)) Then _
        GeometryError "Невозможно состыковать параллельные смещенные линии.", "OFFSET_JOINT", pathID, segmentID
    parameter = ((second.X1 - first.X1) * by - (second.Y1 - first.Y1) * bx) / cross
    first.X2 = first.X1 + parameter * ax: first.Y2 = first.Y1 + parameter * ay
    second.X1 = first.X2: second.Y1 = first.Y2
End Sub

' Обрезает или продлевает линию до двух поперечных ограничений.
' Их положение определяется геометрией независимо от наличия стержней на концах.
Public Function ConstrainLine(ByVal pathID As String, ByVal source As CGeometryPath, _
        ByVal startNX As Double, ByVal startNY As Double, ByVal startBound As Double, _
        ByVal endNX As Double, ByVal endNY As Double, ByVal endBound As Double) As CGeometryPath
    If source Is Nothing Then GeometryError "Не передана ограничиваемая линия.", "MISSING_PATH", pathID
    Dim edge As TRegionEdge, tx As Double, ty As Double, length As Double, first As Double, last As Double, denom As Double
    If source.SegmentCount <> 1 Then GeometryError "Ограничение плоскостями требует одну прямую линию.", "CONSTRAINT_TARGET", pathID, source.ID
    source.CopySegment 1, edge
    If edge.Sweep <> 0# Then GeometryError "Ограничение плоскостями требует прямую линию.", "CONSTRAINT_TARGET", pathID, source.ID
    length = modGeometryPrimitives.EdgeLength(edge): tx = (edge.X2 - edge.X1) / length: ty = (edge.Y2 - edge.Y1) / length
    denom = startNX * tx + startNY * ty
    If Abs(denom) <= PARAM_TOLERANCE Then GeometryError "Начальный ограничитель параллелен линии.", "CONSTRAINT_PARALLEL", pathID, source.ID
    first = (startBound - startNX * edge.X1 - startNY * edge.Y1) / denom
    denom = endNX * tx + endNY * ty
    If Abs(denom) <= PARAM_TOLERANCE Then GeometryError "Конечный ограничитель параллелен линии.", "CONSTRAINT_PARALLEL", pathID, source.ID
    last = (endBound - endNX * edge.X1 - endNY * edge.Y1) / denom
    If last <= first + QUERY_TOLERANCE Then GeometryError "После ограничения концов у линии не осталось положительной длины. Проверьте положение начала и конца: конец должен находиться дальше начала по направлению линии.", "CONSTRAINT_DISAPPEARED", pathID, source.ID, first, last
    Dim result As CGeometryPath, startParameter As Double, endParameter As Double
    startParameter = source.SourceStartParameter(1) + (source.SourceEndParameter(1) - source.SourceStartParameter(1)) * first / length
    endParameter = source.SourceStartParameter(1) + (source.SourceEndParameter(1) - source.SourceStartParameter(1)) * last / length
    edge.X2 = edge.X1 + last * tx: edge.Y2 = edge.Y1 + last * ty
    edge.X1 = edge.X1 + first * tx: edge.Y1 = edge.Y1 + first * ty
    ' Координатные ограничения задают координату точно, без потери разрядов
    ' на вычитании и повторном сложении большого смещения начала линии.
    If startNY = 0# Then edge.X1 = startBound / startNX
    If startNX = 0# Then edge.Y1 = startBound / startNY
    If endNY = 0# Then edge.X2 = endBound / endNX
    If endNX = 0# Then edge.Y2 = endBound / endNY
    Set result = New CGeometryPath: result.Initialize pathID
    result.Append source.SegmentID(1), edge, source.SourcePathID(1), source.SourceSegmentID(1), startParameter, endParameter
    result.SetOffsetBoundary source.OffsetBoundary, source.BoundaryDirection
    result.Seal: Set ConstrainLine = result
End Function

' Подключает модель для геометрических запросов. При наличии точных контуров
' область бетона строится по ним; сетка используется, если наружного контура нет.
Public Sub Initialize(ByVal section As CSectionModel)
    If section Is Nothing Then GeometryError "Модель сечения не передана."
    If section.ConcreteCount = 0 Then GeometryError "В модели нет бетонных элементов."
    Set mSection = section
    mRevision = section.Revision
    Set mDomain = Nothing
End Sub

' Подключает модель и сохраняет уже подготовленную область, пока модель и версия
' её геометрии не изменились. Нагрузки отдельных сочетаний эту подготовку не меняют.
Friend Sub EnsureSection(ByVal section As CSectionModel)
    If section Is Nothing Then GeometryError "Модель сечения не передана."
    If Not section Is mSection Then
        Initialize section
    ElseIf section.Revision <> mRevision Then
        mRevision = section.Revision
        Set mDomain = Nothing
    End If
End Sub

' Возвращает независимую копию всей подготовленной бетонной области с отверстиями.
' После изменения геометрии модели область предварительно строится заново.
Public Function ConcreteDomain() As CConcreteRegion
    EnsureDomain
    Set ConcreteDomain = mDomain.Clone
End Function

' Возвращает наименьшую и наибольшую проекции точной бетонной области, учитывая дуги.
' Используется при построении геометрии и расчёте трещин; проекции бетонной сетки
' для устойчивости вычисляет отдельно CSectionPropertiesCalculator.
Public Sub ProjectionBounds(ByVal region As CConcreteRegion, ByVal nx As Double, ByVal ny As Double, _
        ByRef minimum As Double, ByRef maximum As Double)
    RequireRegion region
    Dim norm As Double
    norm = Sqr(nx * nx + ny * ny)
    If norm <= GEOM_TOLERANCE Then GeometryError "Направление проекции имеет нулевую длину."
    nx = nx / norm: ny = ny / norm
    If region.SegmentCount = 0 Then GeometryError "У пустой области нет крайних проекций."
    Dim edges() As TRegionEdge, count As Long, i As Long
    region.CopyEdges edges, count
    minimum = 1E+100: maximum = -1E+100
    For i = 1 To count
        modGeometryPrimitives.AccumulateEdgeProjection edges(i), nx, ny, minimum, maximum
    Next i
End Sub

' Возвращает точные габариты каждого явно классифицированного отверстия.
' Геометрический генератор сетки получает весь набор независимо от типа
' формы. Колонки массива: minX, maxX, minY, maxY во внутренних мм.
Public Sub OpeningBounds(ByVal contours As CSectionContours, ByRef bounds() As Double, ByRef openingCount As Long)
    Erase bounds: openingCount = 0
    If contours Is Nothing Then GeometryError "Контуры геометрии для определения отверстий не переданы."
    Dim outer() As TRegionEdge, outerCount As Long, holes() As TRegionEdge, holeCount As Long
    ReadMaterialContours contours, outer, outerCount, holes, holeCount
    If holeCount = 0 Then Exit Sub
    OrientLoops holes, holeCount, False
    Dim allOpenings As CConcreteRegion, oneOpening As CConcreteRegion
    Dim oneLoop() As TRegionEdge, oneCount As Long, loopID As Long, i As Long
    Set allOpenings = MakeRegion(holes, holeCount, "ExplicitOpeningBounds", 0)
    allOpenings.CopyEdges holes, holeCount
    openingCount = allOpenings.LoopCount
    ReDim bounds(1 To openingCount, 1 To 4)
    For loopID = 1 To openingCount
        Erase oneLoop: oneCount = 0
        For i = 1 To holeCount
            If holes(i).LoopID = loopID Then AppendEdge oneLoop, oneCount, holes(i)
        Next i
        Set oneOpening = MakeRegion(oneLoop, oneCount, "ExplicitOpeningBounds", 0)
        ProjectionBounds oneOpening, 1#, 0#, bounds(loopID, 1), bounds(loopID, 2)
        ProjectionBounds oneOpening, 0#, 1#, bounds(loopID, 3), bounds(loopID, 4)
    Next loopID
End Sub

' Создаёт бетонную область заданного замкнутого многоугольника. Координаты уже
' переданы в мм; общее ядро проверяет границы и выбирает правильное направление обхода.
Public Function PolygonRegion(ByVal coordinates As Variant) As CConcreteRegion
    Dim edges() As TRegionEdge, count As Long, i As Long, j As Long
    If Not IsArray(coordinates) Then GeometryError "Координаты контура должны быть массивом пар X/Y."
    For i = LBound(coordinates, 1) To UBound(coordinates, 1)
        j = i + 1
        If j > UBound(coordinates, 1) Then j = LBound(coordinates, 1)
        AddLine edges, count, CDbl(coordinates(i, 1)), CDbl(coordinates(i, 2)), _
            CDbl(coordinates(j, 1)), CDbl(coordinates(j, 2))
    Next i
    OrientLoops edges, count, False
    Set PolygonRegion = MakeRegion(edges, count, "ExplicitContour", 0)
End Function

' Создаёт точную круглую область четырьмя дугами. Окружность не заменяется
' многоугольником ни при расчёте площади, ни при дальнейшей обрезке области.
Public Function CircleRegion(ByVal centerX As Double, ByVal centerY As Double, ByVal radius As Double) As CConcreteRegion
    If radius <= QUERY_TOLERANCE Then GeometryError "Радиус контура должен быть положительным."
    Dim edges() As TRegionEdge, count As Long
    AddCircle edges, count, centerX, centerY, radius, False
    Set CircleRegion = MakeRegion(edges, count, "ExplicitContour", 0)
End Function

' Собирает область по неупорядоченным отрезкам и дугам точного контура.
' Проверяет замкнутость, пересечения и дубликаты. Наружные границы и отверстия
' различает по вложенности, поэтому направление обхода исходного объекта
' AutoCAD не определяет назначение контура.
Public Function BoundaryRegion(ByRef edges() As TRegionEdge, ByVal count As Long, _
        Optional ByVal source As String = "ValidatedContour") As CConcreteRegion
    If count <= 0 Then GeometryError "Заданный контур не содержит граничных сегментов."
    Dim i As Long, j As Long, ordered() As TRegionEdge, loops As Long
    For i = 1 To count
        If modGeometryPrimitives.EdgeLength(edges(i)) <= QUERY_TOLERANCE Then GeometryError "Заданный контур содержит ребро нулевой длины."
        If Abs(edges(i).Sweep) > PARAM_TOLERANCE Then ValidateArc edges(i)
    Next i
    StitchLoops edges, count, ordered, loops, True
    ValidateSimpleBoundary ordered, count

    Dim oneLoop() As TRegionEdge, oneCount As Long, loopID As Long, otherLoop As Long
    Dim x As Double, y As Double, depth As Long, area As Double
    For loopID = 1 To loops
        area = SignedLoopArea(ordered, count, loopID)
        If Abs(area) <= QUERY_TOLERANCE Then GeometryError "Заданная геометрия содержит замкнутый контур нулевой площади."
        For i = 1 To count
            If ordered(i).LoopID = loopID Then
                modGeometryPrimitives.EdgePoint ordered(i), 0.5, x, y
                Exit For
            End If
        Next i
        depth = 0
        For otherLoop = 1 To loops
            If otherLoop <> loopID Then
                Erase oneLoop: oneCount = 0
                For j = 1 To count
                    If ordered(j).LoopID = otherLoop Then AppendEdge oneLoop, oneCount, ordered(j)
                Next j
                ' Для внутренней проверки временно направляем обход контура против часовой
                ' стрелки.
                ' Исходное направление объекта на результат проверки не влияет.
                If SignedLoopArea(ordered, count, otherLoop) < 0# Then
                    For j = 1 To oneCount: modGeometryPrimitives.ReverseEdge oneLoop(j): Next j
                End If
                If ContainsRaw(oneLoop, oneCount, x, y, False) Then depth = depth + 1
            End If
        Next otherLoop
        If (area > 0# And depth Mod 2 = 1) Or (area < 0# And depth Mod 2 = 0) Then
            For i = 1 To count
                If ordered(i).LoopID = loopID Then modGeometryPrimitives.ReverseEdge ordered(i)
            Next i
        End If
    Next loopID
    Set BoundaryRegion = MakeRegion(ordered, count, source, 0)
End Function

' Проверяет отсутствие самопересечений каждого контура и касаний между контурами.
' При ошибке исходной геометрии останавливает импорт до создания новой модели.
Private Sub ValidateSimpleBoundary(ByRef edges() As TRegionEdge, ByVal count As Long)
    Dim i As Long, j As Long, k As Long, cuts() As Double, cutCount As Long
    Dim u As Double, x As Double, y As Double, adjacent As Boolean
    For i = 1 To count
        For j = i + 1 To count
            adjacent = False
            If edges(i).LoopID = edges(j).LoopID Then
                adjacent = modGeometryPrimitives.SamePoint(edges(i).X2, edges(i).Y2, edges(j).X1, edges(j).Y1) Or _
                    modGeometryPrimitives.SamePoint(edges(j).X2, edges(j).Y2, edges(i).X1, edges(i).Y1)
            End If
            Erase cuts: cutCount = 0
            modGeometryPrimitives.EdgeIntersectionCuts edges(i), edges(j), cuts, cutCount
            For k = 1 To cutCount
                If Not adjacent Or (cuts(k) > PARAM_TOLERANCE And cuts(k) < 1# - PARAM_TOLERANCE) Then _
                    GeometryError "Заданная геометрия содержит пересечение, касание независимых замкнутых контуров или дублирующее ребро."
            Next k
            Erase cuts: cutCount = 0
            modGeometryPrimitives.EdgeIntersectionCuts edges(j), edges(i), cuts, cutCount
            For k = 1 To cutCount
                If Not adjacent Or (cuts(k) > PARAM_TOLERANCE And cuts(k) < 1# - PARAM_TOLERANCE) Then _
                    GeometryError "Заданная геометрия содержит пересечение, касание независимых замкнутых контуров или дублирующее ребро."
            Next k
            ' Полностью совпадающие сегменты могут иметь только концевые
            ' пересечения. Проверяем их середины без аппроксимации дуг.
            modGeometryPrimitives.EdgePoint edges(i), 0.5, x, y
            If modGeometryPrimitives.PointParameter(edges(j), x, y, u) Then
                If u > PARAM_TOLERANCE And u < 1# - PARAM_TOLERANCE Then _
                    GeometryError "Заданный контур содержит совпадающие граничные сегменты."
            End If
        Next j
    Next i
End Sub

' Возвращает назначение проверенного контура: наружная граница или отверстие.
' Классы импорта и записи получают готовое назначение без собственных догадок по тексту ошибки.
Public Function LoopIsOpening(ByVal region As CConcreteRegion, ByVal loopID As Long) As Boolean
    RequireRegion region
    If loopID < 1 Or loopID > region.LoopCount Then GeometryError "Номер замкнутого контура вне допустимого диапазона."
    Dim edges() As TRegionEdge, count As Long
    region.CopyEdges edges, count
    LoopIsOpening = (SignedLoopArea(edges, count, loopID) < 0#)
End Function

' Выбирает связную часть материала, содержащую заданную точку. Сохраняет
' ее настоящие прямые/дуговые границы и отверстия, не продолжая вырезы
' за их геометрические пределы. Отдельная часть внутри отверстия относится
' к собственному наружному контуру и не присоединяется к выбранной части.
Public Function ConnectedRegionAtPoint(ByVal region As CConcreteRegion, ByVal x As Double, _
        ByVal y As Double) As CConcreteRegion
    RequireRegion region
    Dim edges() As TRegionEdge, count As Long, selected() As TRegionEdge, selectedCount As Long
    Dim outerLoop As Long, loopID As Long, i As Long, keepLoop As Boolean, sampleX As Double, sampleY As Double
    region.CopyEdges edges, count
    If ContainsRaw(edges, count, x, y, True) Then _
        outerLoop = ContainingOuterLoop(edges, count, region.LoopCount, x, y)
    If outerLoop > 0 Then
        For loopID = 1 To region.LoopCount
            keepLoop = (loopID = outerLoop)
            If SignedLoopArea(edges, count, loopID) < 0# Then
                For i = 1 To count
                    If edges(i).LoopID = loopID Then
                        modGeometryPrimitives.EdgePoint edges(i), 0.5, sampleX, sampleY
                        Exit For
                    End If
                Next i
                keepLoop = (ContainingOuterLoop(edges, count, region.LoopCount, sampleX, sampleY) = outerLoop)
            End If
            If keepLoop Then
                For i = 1 To count
                    If edges(i).LoopID = loopID Then AppendEdge selected, selectedCount, edges(i)
                Next i
            End If
        Next loopID
    End If
    Set ConnectedRegionAtPoint = MakeRegion(selected, selectedCount, region.Source, region.SectionRevision)
    ConnectedRegionAtPoint.SetOwner region.OwnerID, region.Standard, region.AnchorID
End Function

' Находит ближайшую наружную границу, охватывающую выбранный контур.
' Выбор наименьшей подходящей области позволяет различить отдельную бетонную часть
' внутри отверстия и окружающий бетон и правильно отнести к ним отверстия.
Private Function ContainingOuterLoop(ByRef edges() As TRegionEdge, ByVal count As Long, _
        ByVal loopCount As Long, ByVal x As Double, ByVal y As Double) As Long
    Dim loopID As Long, i As Long, area As Double, bestArea As Double
    Dim oneLoop() As TRegionEdge, oneCount As Long
    bestArea = 1E+100
    For loopID = 1 To loopCount
        area = SignedLoopArea(edges, count, loopID)
        If area > QUERY_TOLERANCE And area < bestArea Then
            Erase oneLoop: oneCount = 0
            For i = 1 To count
                If edges(i).LoopID = loopID Then AppendEdge oneLoop, oneCount, edges(i)
            Next i
            If ContainsRaw(oneLoop, oneCount, x, y, True) Then
                ContainingOuterLoop = loopID: bestArea = area
            End If
        End If
    Next loopID
End Function

' Объединяет раздельные бетонные области, сохраняя отверстия каждой.
' Наложение или полное совпадение исходных объектов считается ошибкой импорта;
' геометрия чертежа автоматически не исправляется.
Public Function MergeDisjointRegions(ByVal regions As Collection) As CConcreteRegion
    If regions Is Nothing Then GeometryError "Набор контуров не передан."
    If regions.Count = 0 Then GeometryError "На заданном слое не найдено замкнутых контуров."
    Dim raw() As TRegionEdge, count As Long, one() As TRegionEdge, oneCount As Long
    Dim i As Long, j As Long, k As Long, region As CConcreteRegion, other As CConcreteRegion
    For i = 1 To regions.Count
        Set region = regions(i)
        For j = 1 To i - 1
            Set other = regions(j)
            If BoundariesMeet(region, other) Or MaterialSampleInside(region, other) Or MaterialSampleInside(other, region) Then _
                GeometryError "Импортируемые контуры пересекаются, касаются или дублируют одну область."
        Next j
        region.CopyEdges one, oneCount
        For k = 1 To oneCount: AppendEdge raw, count, one(k): Next k
    Next i
    Set MergeDisjointRegions = MakeRegion(raw, count, "ValidatedAutoCADContours", 0)
End Function

' Вычитает отдельное явно заданное отверстие после проверки его расположения.
' Отверстие вне бетона или поверх существующей пустоты вызывает ошибку,
' а не незаметно обрезается по границам допустимой области.
Public Function SubtractContainedRegion(ByVal outer As CConcreteRegion, ByVal opening As CConcreteRegion) As CConcreteRegion
    ValidateContainedRegion outer, opening
    Set SubtractContainedRegion = SubtractRegion(outer, opening)
End Function

' Проверяет принадлежность отверстия, не создавая ненужную вычтенную копию родительской области.
' Те же проверки использует SubtractContainedRegion; отдельной геометрической трактовки нет.
Friend Sub ValidateContainedRegion(ByVal outer As CConcreteRegion, ByVal opening As CConcreteRegion)
    RequireRegion outer: RequireRegion opening
    If BoundariesMeet(outer, opening) Then GeometryError "Контур отверстия пересекает или касается границы материала либо другой пустоты."
    If Not AllMaterialSamplesInside(opening, outer) Then GeometryError "Заданное отверстие находится вне бетонной области."
    If AnyBoundarySampleInside(outer, opening) Then GeometryError "Заданное отверстие накрывает уже существующую пустоту или отдельную компоненту бетона."
End Sub

' Проверяет уже собранные замкнутые контуры: длины отрезков, соответствие дуг
' их центрам и отсутствие недопустимых пересечений. Повторно соединять сегменты
' и строить ту же область для этой проверки не требуется.
Friend Sub ValidatePreparedBoundary(ByVal region As CConcreteRegion)
    RequireRegion region
    Dim edges() As TRegionEdge, count As Long, i As Long
    region.CopyEdges edges, count
    If count <= 0 Then GeometryError "Заданный контур не содержит граничных сегментов."
    For i = 1 To count
        If modGeometryPrimitives.EdgeLength(edges(i)) <= QUERY_TOLERANCE Then GeometryError "Заданный контур содержит ребро нулевой длины."
        If Abs(edges(i).Sweep) > PARAM_TOLERANCE Then ValidateArc edges(i)
    Next i
    ValidateSimpleBoundary edges, count
End Sub

' Проверяет пересечение точных границ двух областей, включая совпавшие концы
' и наложенные сегменты. Все потребители используют эту общую проверку.
Private Function BoundariesMeet(ByVal first As CConcreteRegion, ByVal second As CConcreteRegion) As Boolean
    Dim a() As TRegionEdge, b() As TRegionEdge, na As Long, nb As Long, i As Long, j As Long
    Dim cuts() As Double, count As Long
    first.CopyEdges a, na: second.CopyEdges b, nb
    For i = 1 To na
        For j = 1 To nb
            Erase cuts: count = 0
            modGeometryPrimitives.EdgeIntersectionCuts a(i), b(j), cuts, count
            If count > 0 Then BoundariesMeet = True: Exit Function
        Next j
    Next i
End Function

' Проверяет, есть ли часть бетона первой области внутри второй.
' Берёт контрольные точки со стороны бетона у границ, включая отверстия;
' центр тяжести для этого не подходит, поскольку может лежать в пустоте.
Private Function MaterialSampleInside(ByVal first As CConcreteRegion, ByVal second As CConcreteRegion) As Boolean
    Dim edges() As TRegionEdge, count As Long, i As Long, x As Double, y As Double, tx As Double, ty As Double, length As Double
    first.CopyEdges edges, count
    For i = 1 To count
        modGeometryPrimitives.EdgePoint edges(i), 0.5, x, y: modGeometryPrimitives.EdgeTangent edges(i), 0.5, tx, ty
        length = Sqr(tx * tx + ty * ty)
        If ContainsPoint(second, x - ty / length * QUERY_TOLERANCE * 8#, y + tx / length * QUERY_TOLERANCE * 8#, False) Then _
            MaterialSampleInside = True: Exit Function
    Next i
End Function

' После проверки отсутствия пересечений подтверждает, что все контуры отверстия
' лежат в бетоне. Проверяет каждый сегмент, чтобы учесть раздельные части области
' и не полагаться на одну произвольно выбранную точку.
Private Function AllMaterialSamplesInside(ByVal first As CConcreteRegion, ByVal second As CConcreteRegion) As Boolean
    Dim edges() As TRegionEdge, count As Long, i As Long, x As Double, y As Double
    first.CopyEdges edges, count
    If count = 0 Then Exit Function
    For i = 1 To count
        modGeometryPrimitives.EdgePoint edges(i), 0.5, x, y
        If Not ContainsPoint(second, x, y, False) Then Exit Function
    Next i
    AllMaterialSamplesInside = True
End Function

' Проверяет, не накрывает ли новое отверстие уже существующую границу бетона.
' Так выявляются повторное отверстие и отверстие, заданное поверх пустоты.
Private Function AnyBoundarySampleInside(ByVal first As CConcreteRegion, ByVal second As CConcreteRegion) As Boolean
    Dim edges() As TRegionEdge, count As Long, i As Long, x As Double, y As Double
    first.CopyEdges edges, count
    For i = 1 To count
        modGeometryPrimitives.EdgePoint edges(i), 0.5, x, y
        If ContainsPoint(second, x, y, False) Then AnyBoundarySampleInside = True: Exit Function
    Next i
End Function

' Вычитает отверстия из переданной области. Пересекающиеся отверстия не
' вычитаются дважды: итоговая граница определяется принадлежностью материалу.
Public Function SubtractRegion(ByVal outer As CConcreteRegion, ByVal opening As CConcreteRegion) As CConcreteRegion
    RequireRegion outer
    RequireRegion opening
    Dim raw() As TRegionEdge, holes() As TRegionEdge, count As Long, holeCount As Long, i As Long
    outer.CopyEdges raw, count
    opening.CopyEdges holes, holeCount
    For i = 1 To holeCount
        modGeometryPrimitives.ReverseEdge holes(i)
        AppendEdge raw, count, holes(i)
    Next i
    Dim boundary() As TRegionEdge, boundaryCount As Long
    NormalizeBoundary raw, count, boundary, boundaryCount
    Set SubtractRegion = MakeRegion(boundary, boundaryCount, outer.Source, outer.SectionRevision)
    SubtractRegion.SetOwner outer.OwnerID, outer.Standard, outer.AnchorID
End Function

' Сравнивает готовые области по обеим геометрическим разностям, а не только
' по площади. Допуск площади выводится из общего допуска длины и периметра;
' направление обхода и разбиение одной дуги на сегменты не меняют ответ.
Public Function RegionsCoincide(ByVal first As CConcreteRegion, ByVal second As CConcreteRegion) As Boolean
    RequireRegion first: RequireRegion second
    Dim edges() As TRegionEdge, count As Long, i As Long, perimeter As Double
    Dim difference As CConcreteRegion, tolerance As Double
    first.CopyEdges edges, count
    For i = 1 To count: perimeter = perimeter + modGeometryPrimitives.EdgeLength(edges(i)): Next i
    second.CopyEdges edges, count
    For i = 1 To count: perimeter = perimeter + modGeometryPrimitives.EdgeLength(edges(i)): Next i
    tolerance = QUERY_TOLERANCE * GeomMax(QUERY_TOLERANCE, perimeter)
    If Abs(first.Area - second.Area) > tolerance Then Exit Function
    Set difference = SubtractRegion(first, second)
    If difference.Area > tolerance Then Exit Function
    Set difference = SubtractRegion(second, first)
    RegionsCoincide = (difference.Area <= tolerance)
End Function

' Оставляет часть области nx*x+ny*y >= bound. Нормаль допускается ненормированной,
' а bound масштабируется вместе с ней. Пустая область является корректным ответом.
Public Function ClipHalfPlane(ByVal region As CConcreteRegion, ByVal nx As Double, _
        ByVal ny As Double, ByVal bound As Double) As CConcreteRegion
    RequireRegion region
    Dim norm As Double
    norm = Sqr(nx * nx + ny * ny)
    If norm <= GEOM_TOLERANCE Then GeometryError "Нормаль отсечения имеет нулевую длину."
    nx = nx / norm: ny = ny / norm: bound = bound / norm
    If region.SegmentCount = 0 Then
        Set ClipHalfPlane = region.Clone
        Exit Function
    End If
    Dim minimum As Double, maximum As Double, roundoff As Double
    ProjectionBounds region, nx, ny, minimum, maximum
    ' Не разбиваем область при касании ее крайней точки. Здесь учитывается
    ' только ошибка Double при переносе координат, не геометрический допуск.
    roundoff = 3.5527136788005E-15 * (Abs(minimum) + Abs(maximum) + Abs(bound) + 1#)
    If bound <= minimum + roundoff Then
        Set ClipHalfPlane = region.Clone
        Exit Function
    ElseIf bound >= maximum - roundoff Then
        Dim emptyEdges() As TRegionEdge
        Set ClipHalfPlane = MakeRegion(emptyEdges, 0, region.Source, region.SectionRevision)
        ClipHalfPlane.SetOwner region.OwnerID, region.Standard, region.AnchorID
        Exit Function
    End If
    Dim raw() As TRegionEdge, rawCount As Long, output() As TRegionEdge, outputCount As Long
    region.CopyEdges raw, rawCount
    Dim cuts() As Double, cutCount As Long, lineCuts() As Double, lineCount As Long
    Dim i As Long, j As Long, x As Double, y As Double, piece As TRegionEdge
    For i = 1 To rawCount
        Erase cuts: cutCount = 0
        modGeometryPrimitives.AddScalar cuts, cutCount, 0#: modGeometryPrimitives.AddScalar cuts, cutCount, 1#
        modGeometryPrimitives.HalfPlaneCuts raw(i), nx, ny, bound, cuts, cutCount
        SortScalars cuts, cutCount
        For j = 1 To cutCount
            modGeometryPrimitives.EdgePoint raw(i), cuts(j), x, y
            If Abs(nx * x + ny * y - bound) <= QUERY_TOLERANCE Then _
                modGeometryPrimitives.AddScalar lineCuts, lineCount, ny * x - nx * y
        Next j
        For j = 1 To cutCount - 1
            modGeometryPrimitives.EdgePoint raw(i), 0.5 * (cuts(j) + cuts(j + 1)), x, y
            ' Короткая поддуга возле касания может целиком лежать внутри,
            ' но ближе допуска к прямой. Ее удаление разрывает контур.
            If nx * x + ny * y > bound Then
                modGeometryPrimitives.SubEdge raw(i), cuts(j), cuts(j + 1), piece
                ' Совпадающий с прямой отрезок создается ниже один раз.
                If Abs(piece.Sweep) > PARAM_TOLERANCE Or _
                        Abs(nx * piece.X1 + ny * piece.Y1 - bound) > QUERY_TOLERANCE Or _
                        Abs(nx * piece.X2 + ny * piece.Y2 - bound) > QUERY_TOLERANCE Then _
                    AppendEdge output, outputCount, piece
            End If
        Next j
    Next i

    ' Между соседними пересечениями принадлежность постоянна. Проверяем
    ' саму середину: фиксированное смещение может перескочить узкий сегмент.
    SortScalars lineCuts, lineCount
    Dim s1 As Double, s2 As Double, mid As Double
    For i = 1 To lineCount - 1
        s1 = lineCuts(i): s2 = lineCuts(i + 1): mid = 0.5 * (s1 + s2)
        x = nx * bound + ny * mid: y = ny * bound - nx * mid
        If CutLineHasMaterial(raw, rawCount, nx, ny, bound, x, y) Then
            AddLine output, outputCount, nx * bound + ny * s1, ny * bound - nx * s1, _
                nx * bound + ny * s2, ny * bound - nx * s2
        End If
    Next i
    Set ClipHalfPlane = MakeRegion(output, outputCount, region.Source, region.SectionRevision)
    ClipHalfPlane.SetOwner region.OwnerID, region.Standard, region.AnchorID
End Function

' Обрезает область четырьмя прямыми по границам nMin/nMax и tMin/tMax
' в повёрнутой системе координат. Положение границ и его нормативный смысл
' определяет вызывающий расчёт.
Public Function ClipWindow(ByVal region As CConcreteRegion, ByVal nx As Double, ByVal ny As Double, _
        ByVal nMin As Double, ByVal nMax As Double, ByVal tMin As Double, ByVal tMax As Double) As CConcreteRegion
    Dim result As CConcreteRegion, norm As Double
    norm = Sqr(nx * nx + ny * ny)
    If norm <= GEOM_TOLERANCE Then GeometryError "Направление окна имеет нулевую длину."
    nx = nx / norm: ny = ny / norm
    Set result = ClipHalfPlane(region, nx, ny, nMin)
    Set result = ClipHalfPlane(result, -nx, -ny, -nMax)
    Set result = ClipHalfPlane(result, -ny, nx, tMin)
    Set ClipWindow = ClipHalfPlane(result, ny, -nx, -tMax)
End Function

' Выбирает бетон, непрерывно связанный с опорной линией n = seedPlane
' вдоль заданного направления. За отверстием дальняя стенка в выбранную область
' не включается. Граница строится из точных отрезков и дуг и не зависит
' от количества контрольных прямых.
Public Function MaterialRegionThroughPlane(ByVal region As CConcreteRegion, ByVal nx As Double, _
        ByVal ny As Double, ByVal seedPlane As Double) As CConcreteRegion
    RequireRegion region
    Dim norm As Double
    norm = Sqr(nx * nx + ny * ny)
    If norm <= GEOM_TOLERANCE Then GeometryError "Направление локальной стенки имеет нулевую длину."
    nx = nx / norm: ny = ny / norm
    If region.SegmentCount = 0 Then
        Set MaterialRegionThroughPlane = region.Clone
        Exit Function
    End If
    Dim raw() As TRegionEdge, rawCount As Long, boundary() As TRegionEdge, boundaryCount As Long
    Dim transverseCuts() As Double, transverseCount As Long, cuts() As Double, cutCount As Long
    Dim i As Long, j As Long, k As Long, x As Double, y As Double, u As Double, radius As Double
    region.CopyEdges raw, rawCount
    For i = 1 To rawCount
        modGeometryPrimitives.AddScalar transverseCuts, transverseCount, -ny * raw(i).X1 + nx * raw(i).Y1
        modGeometryPrimitives.AddScalar transverseCuts, transverseCount, -ny * raw(i).X2 + nx * raw(i).Y2
        If Abs(raw(i).Sweep) > PARAM_TOLERANCE Then
            radius = Sqr((raw(i).X1 - raw(i).CenterX) ^ 2 + (raw(i).Y1 - raw(i).CenterY) ^ 2)
            If modGeometryPrimitives.PointParameter(raw(i), raw(i).CenterX - ny * radius, raw(i).CenterY + nx * radius, u) Then _
                modGeometryPrimitives.AddScalar transverseCuts, transverseCount, -ny * raw(i).CenterX + nx * raw(i).CenterY + radius
            If modGeometryPrimitives.PointParameter(raw(i), raw(i).CenterX + ny * radius, raw(i).CenterY - nx * radius, u) Then _
                modGeometryPrimitives.AddScalar transverseCuts, transverseCount, -ny * raw(i).CenterX + nx * raw(i).CenterY - radius
        End If
        Erase cuts: cutCount = 0
        modGeometryPrimitives.HalfPlaneCuts raw(i), nx, ny, seedPlane, cuts, cutCount
        For j = 1 To cutCount
            modGeometryPrimitives.EdgePoint raw(i), cuts(j), x, y
            modGeometryPrimitives.AddScalar transverseCuts, transverseCount, -ny * x + nx * y
        Next j
    Next i
    SortScalars transverseCuts, transverseCount
    Dim piece As TRegionEdge, tangentX As Double, tangentY As Double, length As Double
    Dim leftInside As Boolean, rightInside As Boolean, seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    ' Выбранный бетонный участок может измениться только у вершин, крайних точек дуг
    ' и пересечений с опорной линией. Поэтому проверяем эти положения,
    ' сохраняя сами дуги точными при вычислении границ и площади.
    For i = 1 To rawCount
        Erase cuts: cutCount = 0
        modGeometryPrimitives.AddScalar cuts, cutCount, 0#: modGeometryPrimitives.AddScalar cuts, cutCount, 1#
        For j = 1 To transverseCount
            modGeometryPrimitives.HalfPlaneCuts raw(i), -ny, nx, transverseCuts(j), cuts, cutCount
        Next j
        SortScalars cuts, cutCount
        For j = 1 To cutCount - 1
            modGeometryPrimitives.SubEdge raw(i), cuts(j), cuts(j + 1), piece
            If modGeometryPrimitives.EdgeLength(piece) > QUERY_TOLERANCE Then
                modGeometryPrimitives.EdgePoint piece, 0.5, x, y
                modGeometryPrimitives.EdgeTangent piece, 0.5, tangentX, tangentY
                length = Sqr(tangentX * tangentX + tangentY * tangentY)
                leftInside = InMaterialIntervalThroughPlane(region, nx, ny, seedPlane, _
                    x - tangentY / length * QUERY_TOLERANCE * 8#, y + tangentX / length * QUERY_TOLERANCE * 8#)
                rightInside = InMaterialIntervalThroughPlane(region, nx, ny, seedPlane, _
                    x + tangentY / length * QUERY_TOLERANCE * 8#, y - tangentX / length * QUERY_TOLERANCE * 8#)
                If leftInside <> rightInside Then
                    If rightInside Then modGeometryPrimitives.ReverseEdge piece
                    AppendUniqueBoundary boundary, boundaryCount, piece, seen
                End If
            End If
        Next j
    Next i
    ' Добавляем границы, за которыми отверстие прерывает связь с опорной линией.
    ' Эти отрезки отделяют доступную стенку от дальнего бетона за пустотой.
    Dim intervals As Variant, nCuts() As Double, nCount As Long, n1 As Double, n2 As Double, t As Double
    For i = 1 To transverseCount
        t = transverseCuts(i)
        intervals = LineIntervals(region, -ny * t, nx * t, nx, ny)
        If Not IsEmpty(intervals) Then
            Erase nCuts: nCount = 0
            For j = LBound(intervals, 1) To UBound(intervals, 1)
                modGeometryPrimitives.AddScalar nCuts, nCount, intervals(j, 1): modGeometryPrimitives.AddScalar nCuts, nCount, intervals(j, 2)
            Next j
            ' Если опорная линия совпала с гранью отверстия, сохраняем также реальные точки
            ' их пересечения. Иначе замыкающий отрезок можно ошибочно продлить за край
            ' отверстия.
            For k = 1 To rawCount
                Erase cuts: cutCount = 0
                modGeometryPrimitives.HalfPlaneCuts raw(k), -ny, nx, t, cuts, cutCount
                modGeometryPrimitives.AddScalar cuts, cutCount, 0#: modGeometryPrimitives.AddScalar cuts, cutCount, 1#
                For j = 1 To cutCount
                    modGeometryPrimitives.EdgePoint raw(k), cuts(j), x, y
                    If Abs(-ny * x + nx * y - t) <= QUERY_TOLERANCE Then _
                        modGeometryPrimitives.AddScalar nCuts, nCount, nx * x + ny * y
                Next j
            Next k
            modGeometryPrimitives.AddScalar nCuts, nCount, seedPlane
            SortScalars nCuts, nCount
            For j = 1 To nCount - 1
                n1 = nCuts(j): n2 = nCuts(j + 1)
                If n2 > n1 + QUERY_TOLERANCE Then
                    x = nx * (n1 + n2) / 2# - ny * t
                    y = ny * (n1 + n2) / 2# + nx * t
                    leftInside = InMaterialIntervalThroughPlane(region, nx, ny, seedPlane, _
                        x - ny * QUERY_TOLERANCE * 8#, y + nx * QUERY_TOLERANCE * 8#)
                    rightInside = InMaterialIntervalThroughPlane(region, nx, ny, seedPlane, _
                        x + ny * QUERY_TOLERANCE * 8#, y - nx * QUERY_TOLERANCE * 8#)
                    If leftInside <> rightInside Then
                        piece.X1 = nx * n1 - ny * t: piece.Y1 = ny * n1 + nx * t
                        piece.X2 = nx * n2 - ny * t: piece.Y2 = ny * n2 + nx * t
                        piece.CenterX = 0#: piece.CenterY = 0#: piece.Sweep = 0#: piece.LoopID = 0
                        If rightInside Then modGeometryPrimitives.ReverseEdge piece
                        AppendUniqueBoundary boundary, boundaryCount, piece, seen
                    End If
                End If
            Next j
        End If
    Next i
    Dim normalized() As TRegionEdge, normalizedCount As Long
    ' Новая граница выбранной стенки может совпасть с частью исходной грани.
    ' Разделяем наложенные участки и оставляем одну общую границу.
    NormalizeBoundary boundary, boundaryCount, normalized, normalizedCount
    Set MaterialRegionThroughPlane = MakeRegion(normalized, normalizedCount, region.Source, region.SectionRevision)
    MaterialRegionThroughPlane.SetOwner region.OwnerID, region.Standard, region.AnchorID
End Function

' Проверяет, принадлежит ли точка бетонному участку, связанному с опорной линией.
' Одного попадания в бетон недостаточно: дальняя стенка за пустотой исключается.
Private Function InMaterialIntervalThroughPlane(ByVal region As CConcreteRegion, ByVal nx As Double, _
        ByVal ny As Double, ByVal seedPlane As Double, ByVal x As Double, ByVal y As Double) As Boolean
    Dim t As Double, n As Double, intervals As Variant, i As Long
    t = -ny * x + nx * y: n = nx * x + ny * y - seedPlane
    intervals = LineIntervals(region, nx * seedPlane - ny * t, ny * seedPlane + nx * t, nx, ny)
    If IsEmpty(intervals) Then Exit Function
    For i = LBound(intervals, 1) To UBound(intervals, 1)
        If intervals(i, 1) <= QUERY_TOLERANCE And intervals(i, 2) >= -QUERY_TOLERANCE Then
            If n >= intervals(i, 1) - PARAM_TOLERANCE And n <= intervals(i, 2) + PARAM_TOLERANCE Then
                InMaterialIntervalThroughPlane = True
                Exit Function
            End If
        End If
    Next i
End Function

' Добавляет один направленный отрезок или дугу границы.
' Совпадение исходной и новой границ не должно добавлять второй участок и площадь.
Private Sub AppendUniqueBoundary(ByRef edges() As TRegionEdge, ByRef count As Long, _
        ByRef edge As TRegionEdge, ByVal seen As Object)
    Dim key As String, i As Long, u1 As Double, u2 As Double
    If Abs(edge.Sweep) <= PARAM_TOLERANCE Then
        For i = 1 To count
            If Abs(edges(i).Sweep) <= PARAM_TOLERANCE Then
                If modGeometryPrimitives.PointParameter(edges(i), edge.X1, edge.Y1, u1) And _
                        modGeometryPrimitives.PointParameter(edges(i), edge.X2, edge.Y2, u2) Then
                    If u2 > u1 Then Exit Sub
                End If
            End If
        Next i
    End If
    key = PointKey(edge.X1, edge.Y1) & "/" & PointKey(edge.X2, edge.Y2) & "/" & _
        PointKey(edge.CenterX, edge.CenterY) & "/" & Format$(edge.Sweep, "0.0000000000")
    If seen.Exists(key) Then Exit Sub
    seen.Add key, True
    AppendEdge edges, count, edge
End Sub

' Проверяет, принадлежит ли точка бетонной области. При Boundary = True
' точки на наружном контуре и контуре отверстия считаются принадлежащими области.
' Точки внутри отверстия всегда отклоняются.
Public Function ContainsPoint(ByVal region As CConcreteRegion, ByVal x As Double, ByVal y As Double, _
        Optional ByVal includeBoundary As Boolean = True) As Boolean
    RequireRegion region
    Dim edges() As TRegionEdge, count As Long
    region.CopyEdges edges, count
    ContainsPoint = ContainsRaw(edges, count, x, y, includeBoundary)
End Function

' Проверяет точку по заранее подготовленным границам; область уже проверена при создании массива.
' В отличие от сеточной проверки позволяет включать границы отверстий в проверочную область ряда.
Friend Function ContainsPreparedPoint(ByRef edges() As TRegionEdge, ByVal count As Long, ByVal x As Double, ByVal y As Double, _
        Optional ByVal includeBoundary As Boolean = True) As Boolean
    If count > 0 Then ContainsPreparedPoint = ContainsRaw(edges, count, x, y, includeBoundary)
End Function

' Опорную точку подписи переносим до бетонной границы по заданной нормали.
' Из бетона ищется выход вперёд, из внешней точки — ближайшая граница в обе стороны.
' Метод использует точные отрезки/дуги; сетка и расчётное сечение не меняются.
Friend Function LabelBoundaryDistance(ByVal region As CConcreteRegion, ByVal x As Double, ByVal y As Double, _
        ByVal normalX As Double, ByVal normalY As Double, ByRef signedDistance As Double) As Boolean
    RequireRegion region
    Dim norm As Double, edges() As TRegionEdge, count As Long, i As Long
    Dim distance As Double, parameter As Double, nearest As Double, inside As Boolean
    norm = Sqr(normalX ^ 2 + normalY ^ 2)
    If norm <= QUERY_TOLERANCE Then GeometryError "Не задано направление отступа подписи."
    normalX = normalX / norm: normalY = normalY / norm
    region.CopyEdges edges, count: inside = ContainsRaw(edges, count, x, y, True)
    nearest = 1E+100: signedDistance = 0#
    For i = 1 To count
        If modGeometryPrimitives.EdgeRayHit(edges(i), x, y, normalX, normalY, distance, parameter) Then
            If distance < nearest Then nearest = distance: signedDistance = distance: LabelBoundaryDistance = True
        End If
        If Not inside Then
            If modGeometryPrimitives.EdgeRayHit(edges(i), x, y, -normalX, -normalY, distance, parameter) Then
                If distance < nearest Then nearest = distance: signedDistance = -distance: LabelBoundaryDistance = True
            End If
        End If
    Next i
End Function

' Минимальное расстояние до фактической границы из отрезков и дуг.
' Проверка полного диска стержня дополняет, но не заменяет ContainsPoint.
Public Function BoundaryDistance(ByVal region As CConcreteRegion, ByVal x As Double, ByVal y As Double) As Double
    RequireRegion region
    Dim edges() As TRegionEdge, count As Long, i As Long, dx As Double, dy As Double
    Dim u As Double, px As Double, py As Double, norm As Double, radius As Double, distance As Double
    region.CopyEdges edges, count
    If count = 0 Then GeometryError "У пустой области нет расстояния до границы."
    BoundaryDistance = 1E+100
    For i = 1 To count
        distance = Sqr((x - edges(i).X1) ^ 2 + (y - edges(i).Y1) ^ 2)
        distance = GeomMin(distance, Sqr((x - edges(i).X2) ^ 2 + (y - edges(i).Y2) ^ 2))
        If edges(i).Sweep = 0# Then
            dx = edges(i).X2 - edges(i).X1: dy = edges(i).Y2 - edges(i).Y1
            u = ((x - edges(i).X1) * dx + (y - edges(i).Y1) * dy) / (dx * dx + dy * dy)
            u = GeomMax(0#, GeomMin(1#, u))
            modGeometryPrimitives.EdgePoint edges(i), u, px, py
            distance = Sqr((x - px) ^ 2 + (y - py) ^ 2)
        Else
            dx = x - edges(i).CenterX: dy = y - edges(i).CenterY
            norm = Sqr(dx * dx + dy * dy)
            radius = Sqr((edges(i).X1 - edges(i).CenterX) ^ 2 + (edges(i).Y1 - edges(i).CenterY) ^ 2)
            If norm > QUERY_TOLERANCE Then
                px = edges(i).CenterX + radius * dx / norm: py = edges(i).CenterY + radius * dy / norm
                If modGeometryPrimitives.PointParameter(edges(i), px, py, u) Then distance = GeomMin(distance, Abs(norm - radius))
            End If
        End If
        BoundaryDistance = GeomMin(BoundaryDistance, distance)
    Next i
End Function

' Строит скруглённый многоугольник по завершённому CPolygonSpec.
' Сохраняет исходные вершины и их порядок; у вогнутого угла меняет направление дуги.
' Заданные радиусы самостоятельно не уменьшает.
Public Function RoundedPolygon(ByVal pathID As String, ByVal polygon As CPolygonSpec) As CGeometryPath
    Dim vertices As Variant, radii As Variant, vertexIDs As Variant, edgeIDs As Variant, cornerIDs As Variant
    PolygonKernelData pathID, polygon, vertices, radii, vertexIDs, edgeIDs, cornerIDs
    Set RoundedPolygon = RoundedPolygonKernel(pathID, vertices, radii, vertexIDs, edgeIDs, cornerIDs)
End Function

' Преобразует вершины CPolygonSpec во внутренние массивы координат, радиусов и имён.
' Размеры массивов должны совпадать; рецепты сечения напрямую эти массивы не задают.
Private Function RoundedPolygonKernel(ByVal pathID As String, ByVal vertices As Variant, _
        ByVal radii As Variant, ByVal vertexIDs As Variant, ByVal edgeIDs As Variant, ByVal cornerIDs As Variant) As CGeometryPath
    Dim points() As TPoint2D, radius() As Double, names() As String, count As Long
    ReadPolygonInput pathID, vertices, radii, vertexIDs, points, radius, names, count
    Dim raw() As TRegionEdge, i As Long, previous As Long, following As Long
    Dim directions() As TPoint2D, lengths() As Double, distances() As Double, turns() As Double
    Dim first() As TPoint2D, last() As TPoint2D, centers() As TPoint2D
    ReDim raw(1 To count): ReDim directions(1 To count): ReDim lengths(1 To count)
    ReDim distances(1 To count): ReDim turns(1 To count)
    ReDim first(1 To count): ReDim last(1 To count): ReDim centers(1 To count)
    For i = 1 To count
        following = i Mod count + 1
        raw(i).X1 = points(i).X: raw(i).Y1 = points(i).Y
        raw(i).X2 = points(following).X: raw(i).Y2 = points(following).Y
        lengths(i) = modGeometryPrimitives.EdgeLength(raw(i))
        If lengths(i) <= QUERY_TOLERANCE Then GeometryError pathID & ": нулевая грань '" & edgeIDs(i - 1) & "'."
        directions(i).X = (raw(i).X2 - raw(i).X1) / lengths(i)
        directions(i).Y = (raw(i).Y2 - raw(i).Y1) / lengths(i)
    Next i
    Dim original As CConcreteRegion
    Set original = BoundaryRegion(raw, count)
    Dim cross As Double, dot As Double, side As Double
    For i = 1 To count
        previous = (i + count - 2) Mod count + 1
        cross = directions(previous).X * directions(i).Y - directions(previous).Y * directions(i).X
        dot = directions(previous).X * directions(i).X + directions(previous).Y * directions(i).Y
        turns(i) = modGeometryPrimitives.AngleOf(dot, cross)
        If radius(i) > 0# Then
            If Abs(turns(i)) <= PARAM_TOLERANCE Or Abs(Abs(turns(i)) - GEOM_PI) <= PARAM_TOLERANCE Then _
                GeometryError "Скругление у вершины '" & names(i) & "' невозможно: соседние участки почти лежат на одной прямой или развёрнуты навстречу друг другу. Увеличьте отклонение вершины от прямой или задайте её радиус равным 0.", "FILLET_ANGLE", pathID, CStr(cornerIDs(i - 1))
            distances(i) = radius(i) * Tan(Abs(turns(i)) / 2#)
            ' Концы очень короткой дуги нельзя надёжно различить с точностью геометрического ядра.
            If 2# * radius(i) * Abs(Sin(turns(i) / 2#)) <= QUERY_TOLERANCE Then _
                GeometryError "Скругление у вершины '" & names(i) & "' слишком мало для надёжного построения: концы дуги практически совпадают. Увеличьте отклонение вершины от прямой, измените радиус или задайте его равным 0.", "FILLET_TOO_SHORT", pathID, CStr(cornerIDs(i - 1))
        End If
        first(i).X = points(i).X - directions(previous).X * distances(i)
        first(i).Y = points(i).Y - directions(previous).Y * distances(i)
        last(i).X = points(i).X + directions(i).X * distances(i)
        last(i).Y = points(i).Y + directions(i).Y * distances(i)
        side = Sgn(turns(i))
        centers(i).X = first(i).X - directions(previous).Y * radius(i) * side
        centers(i).Y = first(i).Y + directions(previous).X * radius(i) * side
    Next i
    For i = 1 To count
        following = i Mod count + 1
        If distances(i) + distances(following) > lengths(i) + GEOM_TOLERANCE Then _
            GeometryError "Скругления соседних вершин пересекаются на участке '" & edgeIDs(i - 1) & "': для заданных размеров не хватает длины прямого участка. Уменьшите радиусы или увеличьте размеры сечения.", "FILLET_OVERLAP", pathID, CStr(edgeIDs(i - 1)), radius(i), radius(following)
    Next i
    Dim result As CGeometryPath, edge As TRegionEdge
    Set result = New CGeometryPath: result.Initialize pathID
    For i = 1 To count
        following = i Mod count + 1
        If radius(i) > 0# Then
            edge.X1 = first(i).X: edge.Y1 = first(i).Y: edge.X2 = last(i).X: edge.Y2 = last(i).Y
            edge.CenterX = centers(i).X: edge.CenterY = centers(i).Y: edge.Sweep = turns(i)
            ValidateArc edge
            result.Append CStr(cornerIDs(i - 1)), edge
        Else
            result.AddCollapsedSegment CStr(cornerIDs(i - 1))
        End If
        edge.X1 = last(i).X: edge.Y1 = last(i).Y: edge.X2 = first(following).X: edge.Y2 = first(following).Y
        edge.CenterX = 0#: edge.CenterY = 0#: edge.Sweep = 0#
        If modGeometryPrimitives.EdgeLength(edge) > QUERY_TOLERANCE Then
            result.Append CStr(edgeIDs(i - 1)), edge
        Else
            result.AddCollapsedSegment CStr(edgeIDs(i - 1))
        End If
    Next i
    result.Seal True
    ' Проверяем весь готовый контур, чтобы обнаружить пересечения скруглений
    ' не только у соседних вершин.
    Dim prepared As CConcreteRegion
    Set prepared = RegionFromPath(result)
    Set RoundedPolygonKernel = result
End Function

' Строит смещённый скруглённый многоугольник по тому же CPolygonSpec.
' Положительное расстояние направлено внутрь независимо от направления обхода,
' отрицательное — наружу, например от отверстия в бетон. Острые углы остаются
' острыми; дуга, радиус которой уменьшился до нуля, исчезает.
Public Function OffsetRoundedPolygon(ByVal pathID As String, ByVal polygon As CPolygonSpec, ByVal distance As Double) As CGeometryPath
    Dim vertices As Variant, radii As Variant, vertexIDs As Variant, edgeIDs As Variant, cornerIDs As Variant
    PolygonKernelData pathID, polygon, vertices, radii, vertexIDs, edgeIDs, cornerIDs
    Set OffsetRoundedPolygon = OffsetRoundedPolygonKernel(pathID, vertices, radii, vertexIDs, edgeIDs, cornerIDs, distance)
End Function

' Смещаем исходный многоугольник до построения скруглений.
' Знак здесь означает внутрь или наружу площади, а не влево по обходу готовой линии.
Private Function OffsetRoundedPolygonKernel(ByVal pathID As String, ByVal vertices As Variant, _
        ByVal radii As Variant, ByVal vertexIDs As Variant, ByVal edgeIDs As Variant, ByVal cornerIDs As Variant, ByVal distance As Double) As CGeometryPath
    Dim original As CGeometryPath
    Set original = RoundedPolygonKernel(pathID, vertices, radii, vertexIDs, edgeIDs, cornerIDs)
    If distance = 0# Then Set OffsetRoundedPolygonKernel = original: Exit Function

    Dim points() As TPoint2D, radius() As Double, names() As String, count As Long
    ReadPolygonInput pathID, vertices, radii, vertexIDs, points, radius, names, count
    Dim directions() As TPoint2D, shifted() As TPoint2D, i As Long, following As Long, previous As Long
    Dim length As Double, signedArea As Double, orientation As Double
    ReDim directions(1 To count): ReDim shifted(1 To count)
    For i = 1 To count
        following = i Mod count + 1
        directions(i).X = points(following).X - points(i).X
        directions(i).Y = points(following).Y - points(i).Y
        length = Sqr(directions(i).X ^ 2 + directions(i).Y ^ 2)
        directions(i).X = directions(i).X / length: directions(i).Y = directions(i).Y / length
        signedArea = signedArea + (points(i).X - points(1).X) * (points(following).Y - points(1).Y) - _
            (points(following).X - points(1).X) * (points(i).Y - points(1).Y)
    Next i
    orientation = Sgn(signedArea)
    Dim cross As Double, dot As Double, beforeX As Double, beforeY As Double, afterX As Double, afterY As Double, parameter As Double
    Dim shiftedVertices() As Variant, shiftedRadii() As Variant
    ReDim shiftedVertices(0 To count - 1): ReDim shiftedRadii(0 To count - 1)
    For i = 1 To count
        previous = (i + count - 2) Mod count + 1
        beforeX = points(i).X - orientation * directions(previous).Y * distance
        beforeY = points(i).Y + orientation * directions(previous).X * distance
        afterX = points(i).X - orientation * directions(i).Y * distance
        afterY = points(i).Y + orientation * directions(i).X * distance
        cross = directions(previous).X * directions(i).Y - directions(previous).Y * directions(i).X
        If Abs(cross) <= PARAM_TOLERANCE Then
            dot = directions(previous).X * directions(i).X + directions(previous).Y * directions(i).Y
            If dot <= 0# Then GeometryError pathID & ": смещение не определено в развороте вершины '" & names(i) & "'."
            shifted(i).X = afterX: shifted(i).Y = afterY
        Else
            parameter = ((afterX - beforeX) * directions(i).Y - (afterY - beforeY) * directions(i).X) / cross
            shifted(i).X = beforeX + parameter * directions(previous).X
            shifted(i).Y = beforeY + parameter * directions(previous).Y
        End If
        shiftedVertices(i - 1) = Array(shifted(i).X, shifted(i).Y)
        If radius(i) > 0# Then
            shiftedRadii(i - 1) = GeomMax(0#, radius(i) - distance * orientation * Sgn(cross))
        Else
            shiftedRadii(i - 1) = 0#
        End If
    Next i
    For i = 1 To count
        following = i Mod count + 1
        dot = directions(i).X * (shifted(following).X - shifted(i).X) + _
            directions(i).Y * (shifted(following).Y - shifted(i).Y)
        If dot <= QUERY_TOLERANCE Then GeometryError pathID & ": после смещения прямой участок '" & edgeIDs(i - 1) & "' исчез или изменил направление. Уменьшите расстояние смещения или увеличьте размеры исходного контура.", "OFFSET_DISAPPEARED", pathID, CStr(edgeIDs(i - 1)), distance, dot
    Next i
    ' Проверяем радиусы и самопересечения общими средствами проверки контура.
    Set OffsetRoundedPolygonKernel = RoundedPolygonKernel(pathID, shiftedVertices, shiftedRadii, vertexIDs, edgeIDs, cornerIDs)
End Function

' Копирует исходные вершины и радиусы из описания во внутренние массивы
' геометрического ядра, не меняя само описание.
Private Sub PolygonKernelData(ByVal pathID As String, ByVal polygon As CPolygonSpec, _
        ByRef vertices As Variant, ByRef radii As Variant, ByRef vertexIDs As Variant, ByRef edgeIDs As Variant, ByRef cornerIDs As Variant)
    If polygon Is Nothing Then GeometryError pathID & ": не передано описание исходных вершин."
    If Not polygon.IsSealed Then GeometryError pathID & ": завершите описание исходных вершин методом Seal."
    polygon.CopyKernelData vertices, radii, vertexIDs, edgeIDs, cornerIDs
End Sub

' Линия сохраняет исходное направление и имена сегментов. Для бетонной области
' отдельно приводим обход наружных контуров и отверстий к принятому направлению.
Public Function RegionFromPath(ByVal path As CGeometryPath) As CConcreteRegion
    If path Is Nothing Then GeometryError "Путь границы не передан."
    If Not path.IsClosed Then GeometryError "Бетонная граница должна быть замкнутой."
    Dim edges() As TRegionEdge, i As Long
    ReDim edges(1 To path.SegmentCount)
    For i = 1 To path.SegmentCount: path.CopySegment i, edges(i): Next i
    Set RegionFromPath = BoundaryRegion(edges, path.SegmentCount)
End Function

' Строит общую границу смежных прямоугольников без наложения их площадей.
' Фигуры с пересекающимися внутренними частями отклоняются; итог должен иметь
' одну наружную границу.
Public Function AdjacentRectangles(ByVal pathID As String, ByVal rectangles As Variant) As CGeometryPath
    If Not IsArray(rectangles) Then GeometryError "Прямоугольники должны быть массивом четверок x/y/B/H."
    Dim bounds() As Double, n As Long, i As Long, j As Long, rect As Variant, start As Long
    n = UBound(rectangles) - LBound(rectangles) + 1
    If n <= 0 Then GeometryError "Не задан ни один прямоугольник."
    ReDim bounds(1 To n, 1 To 4)
    For i = 1 To n
        rect = rectangles(LBound(rectangles) + i - 1)
        If Not IsArray(rect) Then GeometryError "Прямоугольник должен содержать x/y/B/H."
        start = LBound(rect)
        If UBound(rect) - start <> 3 Then GeometryError "Прямоугольник должен содержать четыре числа."
        bounds(i, 1) = CDbl(rect(start)): bounds(i, 2) = CDbl(rect(start + 1))
        If CDbl(rect(start + 2)) <= 0# Or CDbl(rect(start + 3)) <= 0# Then GeometryError "Ширина и высота прямоугольника должны быть положительными."
        bounds(i, 3) = bounds(i, 1) + CDbl(rect(start + 2))
        bounds(i, 4) = bounds(i, 2) + CDbl(rect(start + 3))
        For j = 1 To i - 1
            If GeomMin(bounds(i, 3), bounds(j, 3)) - GeomMax(bounds(i, 1), bounds(j, 1)) > QUERY_TOLERANCE And _
                GeomMin(bounds(i, 4), bounds(j, 4)) - GeomMax(bounds(i, 2), bounds(j, 2)) > QUERY_TOLERANCE Then _
                GeometryError "Объединение допускает только смежные прямоугольники без пересечения внутренних площадей."
        Next j
    Next i
    Dim raw() As TRegionEdge, count As Long, edge As TRegionEdge, boundary() As TRegionEdge, boundaryCount As Long
    Dim region As CConcreteRegion, result As CGeometryPath
    For i = 1 To n
        edge.Sweep = 0#: edge.LoopID = i
        edge.X1 = bounds(i, 1): edge.Y1 = bounds(i, 2): edge.X2 = bounds(i, 3): edge.Y2 = bounds(i, 2)
        AppendEdge raw, count, edge
        edge.X1 = bounds(i, 3): edge.Y1 = bounds(i, 2): edge.X2 = bounds(i, 3): edge.Y2 = bounds(i, 4)
        AppendEdge raw, count, edge
        edge.X1 = bounds(i, 3): edge.Y1 = bounds(i, 4): edge.X2 = bounds(i, 1): edge.Y2 = bounds(i, 4)
        AppendEdge raw, count, edge
        edge.X1 = bounds(i, 1): edge.Y1 = bounds(i, 4): edge.X2 = bounds(i, 1): edge.Y2 = bounds(i, 2)
        AppendEdge raw, count, edge
    Next i
    NormalizeBoundary raw, count, boundary, boundaryCount
    Set region = BoundaryRegion(boundary, boundaryCount)
    If region.LoopCount <> 1 Then GeometryError "Прямоугольники должны образовывать одну связную область без отверстий."
    region.CopyEdges boundary, boundaryCount
    Set result = New CGeometryPath: result.Initialize pathID
    For i = 1 To boundaryCount
        result.AddLine pathID & "." & CStr(i), boundary(i).X1, boundary(i).Y1, boundary(i).X2, boundary(i).Y2
    Next i
    result.Seal True
    Set AdjacentRectangles = result
End Function

' Разбирает компактную запись координат, переданную программно, без чтения Excel.
Private Sub ReadPolygonInput(ByVal pathID As String, ByVal vertices As Variant, ByVal radii As Variant, _
        ByVal vertexIDs As Variant, ByRef points() As TPoint2D, ByRef radius() As Double, ByRef names() As String, ByRef count As Long)
    On Error GoTo Invalid
    count = UBound(vertices) - LBound(vertices) + 1
    If count < 3 Or UBound(radii) - LBound(radii) + 1 <> count Or UBound(vertexIDs) - LBound(vertexIDs) + 1 <> count Then GoTo Invalid
    ReDim points(1 To count): ReDim radius(1 To count): ReDim names(1 To count)
    Dim i As Long, j As Long, vertex As Variant
    For i = 1 To count
        vertex = vertices(LBound(vertices) + i - 1)
        If UBound(vertex) - LBound(vertex) <> 1 Then GoTo Invalid
        points(i).X = CDbl(vertex(LBound(vertex))): points(i).Y = CDbl(vertex(LBound(vertex) + 1))
        radius(i) = CDbl(radii(LBound(radii) + i - 1))
        If radius(i) < 0# Then GoTo Invalid
        names(i) = CStr(vertexIDs(LBound(vertexIDs) + i - 1))
        If Len(Trim$(names(i))) = 0 Then GoTo Invalid
        For j = 1 To i - 1: If names(i) = names(j) Then GoTo Invalid
        Next j
    Next i
    Exit Sub
Invalid:
    GeometryError pathID & ": задайте не менее трех вершин Array(Array(x,y),...), столько же неотрицательных радиусов и уникальных имен вершин."
End Sub

' Для точки на исходном отрезке учитываем, с какой стороны находится бетон.
' На остальных участках между пересечениями проверяем середину: она лежит
' либо в бетоне, либо в пустоте, в том числе у дуг.
Private Function CutLineHasMaterial(ByRef edges() As TRegionEdge, ByVal count As Long, _
        ByVal nx As Double, ByVal ny As Double, ByVal bound As Double, _
        ByVal x As Double, ByVal y As Double) As Boolean
    Dim i As Long, u As Double
    For i = 1 To count
        If Abs(edges(i).Sweep) <= PARAM_TOLERANCE Then
            If Abs(nx * edges(i).X1 + ny * edges(i).Y1 - bound) <= QUERY_TOLERANCE And _
                    Abs(nx * edges(i).X2 + ny * edges(i).Y2 - bound) <= QUERY_TOLERANCE Then
                If modGeometryPrimitives.PointParameter(edges(i), x, y, u) Then
                    CutLineHasMaterial = nx * (edges(i).Y1 - edges(i).Y2) + _
                        ny * (edges(i).X2 - edges(i).X1) > 0#
                    Exit Function
                End If
            End If
        End If
    Next i
    CutLineHasMaterial = ContainsRaw(edges, count, x, y, False)
End Function

' При отборе точек сетки считает точку на наружном контуре принадлежащей бетону.
' Точку на границе отверстия отклоняет. Для всех проверок использует
' один подготовленный массив границ без повторного копирования.
Friend Function ContainsMeshPoint(ByVal region As CConcreteRegion, ByRef edges() As TRegionEdge, _
        ByVal count As Long, ByVal x As Double, ByVal y As Double) As Boolean
    Dim i As Long, u As Double
    For i = 1 To count
        If modGeometryPrimitives.PointParameter(edges(i), x, y, u) Then
            ContainsMeshPoint = Not region.LoopIsOpening(edges(i).LoopID)
            Exit Function
        End If
    Next i
    ContainsMeshPoint = ContainsRaw(edges, count, x, y, False)
End Function

' Проверяет, не проходит ли узкий вырез через ячейку с четырьмя углами внутри бетона.
' Генератор сетки получает точные границы один раз и не зависит от типа сечения.
Friend Sub CopyMeshOuterEdges(ByVal contours As CSectionContours, ByRef edges() As TRegionEdge, ByRef count As Long)
    Dim holes() As TRegionEdge, holeCount As Long
    Erase edges: count = 0
    ReadMaterialContours contours, edges, count, holes, holeCount
End Sub

' Возвращает реальные вершины и экстремумы дуг, а не углы габаритного
' прямоугольника, которые для произвольного сечения могут быть вне бетона.
Friend Sub MeshExtremePoints(ByRef edges() As TRegionEdge, ByVal count As Long, _
        ByRef pointX() As Double, ByRef pointY() As Double)
    Dim i As Long, j As Long, n As Long, radius As Double, x As Double, y As Double, u As Double
    ReDim pointX(1 To count * 6): ReDim pointY(1 To count * 6)
    For i = 1 To count
        n = n + 1: pointX(n) = edges(i).X1: pointY(n) = edges(i).Y1
        n = n + 1: pointX(n) = edges(i).X2: pointY(n) = edges(i).Y2
        If Abs(edges(i).Sweep) > PARAM_TOLERANCE Then
            radius = Sqr((edges(i).X1 - edges(i).CenterX) ^ 2 + (edges(i).Y1 - edges(i).CenterY) ^ 2)
            For j = 0 To 3
                x = edges(i).CenterX + radius * Cos(j * GEOM_PI / 2#)
                y = edges(i).CenterY + radius * Sin(j * GEOM_PI / 2#)
                If modGeometryPrimitives.PointParameter(edges(i), x, y, u) Then
                    n = n + 1: pointX(n) = x: pointY(n) = y
                End If
            Next j
        End If
    Next i
    ReDim Preserve pointX(1 To n): ReDim Preserve pointY(1 To n)
End Sub

' Проверяет заход наружного ребра в открытую ячейку. Совпадение с ее
' стороной не требует дробления. Линии и дуги используют прежние пересечения.
Friend Function MeshBoundaryCrossesCell(ByRef edges() As TRegionEdge, ByVal count As Long, _
        ByVal minX As Double, ByVal minY As Double, ByVal maxX As Double, ByVal maxY As Double) As Boolean
    Dim i As Long, j As Long, cuts() As Double, cutCount As Long, x As Double, y As Double
    For i = 1 To count
        If Abs(edges(i).Sweep) <= PARAM_TOLERANCE Then
            If GeomMax(edges(i).X1, edges(i).X2) <= minX + QUERY_TOLERANCE Or _
                    GeomMin(edges(i).X1, edges(i).X2) >= maxX - QUERY_TOLERANCE Or _
                    GeomMax(edges(i).Y1, edges(i).Y2) <= minY + QUERY_TOLERANCE Or _
                    GeomMin(edges(i).Y1, edges(i).Y2) >= maxY - QUERY_TOLERANCE Then GoTo NextEdge
        End If
        Erase cuts: cutCount = 0
        modGeometryPrimitives.AddScalar cuts, cutCount, 0#: modGeometryPrimitives.AddScalar cuts, cutCount, 1#
        modGeometryPrimitives.HalfPlaneCuts edges(i), 1#, 0#, minX, cuts, cutCount
        modGeometryPrimitives.HalfPlaneCuts edges(i), 1#, 0#, maxX, cuts, cutCount
        modGeometryPrimitives.HalfPlaneCuts edges(i), 0#, 1#, minY, cuts, cutCount
        modGeometryPrimitives.HalfPlaneCuts edges(i), 0#, 1#, maxY, cuts, cutCount
        SortScalars cuts, cutCount
        For j = 1 To cutCount - 1
            modGeometryPrimitives.EdgePoint edges(i), (cuts(j) + cuts(j + 1)) / 2#, x, y
            If x > minX + QUERY_TOLERANCE And x < maxX - QUERY_TOLERANCE And _
                    y > minY + QUERY_TOLERANCE And y < maxY - QUERY_TOLERANCE Then
                MeshBoundaryCrossesCell = True: Exit Function
            End If
        Next j
NextEdge:
    Next i
End Function

' Пересекает область параметрической прямой P+t*D. Интервалы ограничиваются
' явно переданным диапазоном t: прямая, луч и отрезок используют один алгоритм.
' D нормируется, поэтому t и границы интервалов измеряются в миллиметрах.
Public Function LineIntervals(ByVal region As CConcreteRegion, ByVal pointX As Double, _
        ByVal pointY As Double, ByVal directionX As Double, ByVal directionY As Double, _
        Optional ByVal parameterMin As Double = -1E+100, Optional ByVal parameterMax As Double = 1E+100) As Variant
    RequireRegion region
    Dim norm As Double
    norm = Sqr(directionX * directionX + directionY * directionY)
    If norm <= GEOM_TOLERANCE Then GeometryError "Направление пересечения имеет нулевую длину."
    If parameterMin > parameterMax Then GeometryError "Начало диапазона параметра больше его конца."
    directionX = directionX / norm: directionY = directionY / norm
    Dim edges() As TRegionEdge, count As Long, cuts() As Double, cutCount As Long
    region.CopyEdges edges, count
    Dim localCuts() As Double, localCount As Long, i As Long, j As Long, x As Double, y As Double, t As Double
    Dim lineNormalX As Double, lineNormalY As Double, lineBound As Double
    lineNormalX = -directionY: lineNormalY = directionX
    lineBound = lineNormalX * pointX + lineNormalY * pointY
    For i = 1 To count
        Erase localCuts: localCount = 0
        modGeometryPrimitives.HalfPlaneCuts edges(i), lineNormalX, lineNormalY, lineBound, localCuts, localCount
        modGeometryPrimitives.AddScalar localCuts, localCount, 0#: modGeometryPrimitives.AddScalar localCuts, localCount, 1#
        For j = 1 To localCount
            modGeometryPrimitives.EdgePoint edges(i), localCuts(j), x, y
            If Abs(lineNormalX * x + lineNormalY * y - lineBound) <= QUERY_TOLERANCE Then
                t = (x - pointX) * directionX + (y - pointY) * directionY
                modGeometryPrimitives.AddScalar cuts, cutCount, t
            End If
        Next j
    Next i
    If cutCount < 2 Then Exit Function
    SortScalars cuts, cutCount
    Dim answer() As Double, answerCount As Long, lo As Double, hi As Double
    ReDim answer(1 To cutCount, 1 To 2)
    For i = 1 To cutCount - 1
        lo = GeomMax(cuts(i), parameterMin): hi = GeomMin(cuts(i + 1), parameterMax)
        If hi > lo + QUERY_TOLERANCE Then
            t = 0.5 * (lo + hi)
            If ContainsRaw(edges, count, pointX + t * directionX, pointY + t * directionY, True) Then
                If answerCount > 0 Then
                    If Abs(answer(answerCount, 2) - lo) <= QUERY_TOLERANCE Then
                        answer(answerCount, 2) = hi
                        GoTo NextInterval
                    End If
                End If
                answerCount = answerCount + 1
                answer(answerCount, 1) = lo: answer(answerCount, 2) = hi
            End If
        End If
NextInterval:
    Next i
    If answerCount = 0 Then Exit Function
    Dim final() As Double
    ReDim final(1 To answerCount, 1 To 2)
    For i = 1 To answerCount
        final(i, 1) = answer(i, 1): final(i, 2) = answer(i, 2)
    Next i
    LineIntervals = final
End Function

' На нескольких равномерно расположенных прямых находит начало Q и конец P
' бетонного участка, содержащего опорную точку. Дальняя стенка за отверстием
' в этот участок не включается. Для каждой прямой сохраняет признак наличия участка
' и координаты его концов. Эти данные нужны для проверки; площадь определяется
' точными контурами, а не количеством контрольных прямых.
Public Function SampleMaterialIntervals(ByVal region As CConcreteRegion, ByVal nx As Double, ByVal ny As Double, _
        ByVal seedPlane As Double, ByVal tMin As Double, ByVal tMax As Double, ByVal probeCount As Long) As Variant
    RequireRegion region
    Dim norm As Double, data() As Variant, i As Long, j As Long, t As Double, x As Double, y As Double
    Dim intervals As Variant
    norm = Sqr(nx * nx + ny * ny)
    If norm <= GEOM_TOLERANCE Then GeometryError "Направление пробных прямых имеет нулевую длину."
    If probeCount < 2 Or tMin > tMax Then GeometryError "Некорректное число пробных прямых или границы полосы."
    nx = nx / norm: ny = ny / norm
    ReDim data(1 To probeCount, 1 To 5)
    For i = 1 To probeCount
        t = tMin + (tMax - tMin) * CDbl(i - 1) / CDbl(probeCount - 1)
        x = nx * seedPlane - ny * t: y = ny * seedPlane + nx * t
        data(i, 1) = False
        intervals = LineIntervals(region, x, y, nx, ny)
        If Not IsEmpty(intervals) Then
            For j = LBound(intervals, 1) To UBound(intervals, 1)
                If intervals(j, 1) <= QUERY_TOLERANCE And intervals(j, 2) >= -QUERY_TOLERANCE Then
                    data(i, 1) = True
                    data(i, 2) = x + nx * intervals(j, 1): data(i, 3) = y + ny * intervals(j, 1)
                    data(i, 4) = x + nx * intervals(j, 2): data(i, 5) = y + ny * intervals(j, 2)
                    Exit For
                End If
            Next j
        End If
    Next i
    SampleMaterialIntervals = data
End Function

' Возвращает расстояние от точки до первого выхода из бетона вдоль заданного луча.
' Дальняя стенка за пустотой не увеличивает защитный слой. Если исходная точка
' не принадлежит бетону, возвращает -1.
Public Function ConcreteCoverFromPointAlongDirection(ByVal pointX As Double, ByVal pointY As Double, _
        ByVal directionX As Double, ByVal directionY As Double) As Double
    EnsureDomain
    Dim intervals As Variant
    intervals = LineIntervals(mDomain, pointX, pointY, directionX, directionY, 0#)
    ConcreteCoverFromPointAlongDirection = -1#
    If IsEmpty(intervals) Then Exit Function
    If intervals(1, 1) <= QUERY_TOLERANCE Then ConcreteCoverFromPointAlongDirection = intervals(1, 2)
End Function

' Подготавливает и проверяет общую бетонную область модели. Если геометрия изменилась,
' строит её заново. Номер сочетания нагрузок на подготовку не влияет.
' При ошибке незавершённая область не сохраняется для последующих запросов.
Private Sub EnsureDomain()
    If mSection Is Nothing Then GeometryError "Сначала привяжите геометрический запрос к модели сечения."
    If mSection.Revision <> mRevision Then Set mDomain = Nothing
    If Not mDomain Is Nothing Then Exit Sub
    mRevision = mSection.Revision
    Dim outer() As TRegionEdge, outerCount As Long, holes() As TRegionEdge, holeCount As Long
    ReadMaterialContours mSection.Contours, outer, outerCount, holes, holeCount
    Dim mesh() As TRegionEdge, meshCount As Long
    If outerCount = 0 Then
        MeshBoundary mSection, mesh, meshCount
        If holeCount > 0 Then KeepOuterMeshContours mesh, meshCount
    End If
    Dim raw() As TRegionEdge, count As Long, i As Long
    If outerCount > 0 Then
        ' Известный наружный контур задает всю область. Учитываем только
        ' заданные отверстия; неизвестные пустоты по сетке не добавляются.
        OrientLoops outer, outerCount, False
        For i = 1 To outerCount
            AppendEdge raw, count, outer(i)
        Next i
    Else
        For i = 1 To meshCount
            AppendEdge raw, count, mesh(i)
        Next i
    End If
    If holeCount > 0 Then
        OrientLoops holes, holeCount, True
        For i = 1 To holeCount
            AppendEdge raw, count, holes(i)
        Next i
    End If
    Dim boundary() As TRegionEdge, boundaryCount As Long
    NormalizeBoundary raw, count, boundary, boundaryCount
    Dim source As String
    source = "ElementBoundaryLoops"
    If outerCount > 0 Then source = "AuthoritativeContour"
    If holeCount > 0 Then source = source & "+KnownOpening"
    Dim preparedDomain As CConcreteRegion
    Set preparedDomain = MakeRegion(boundary, boundaryCount, source, mRevision)
    If preparedDomain.Area <= QUERY_TOLERANCE Then GeometryError "Бетонная область пуста после учета внешнего контура и отверстий."
    Set mDomain = preparedDomain
    mDomainBuildCount = mDomainBuildCount + 1
End Sub

' Читает точные наружные границы и отверстия из модели сечения.
' Области взаимодействия, полученные при расчёте трещин, не используются вместо исходных
' контуров.
Private Sub ReadMaterialContours(ByVal contours As CSectionContours, ByRef outer() As TRegionEdge, _
        ByRef outerCount As Long, ByRef holes() As TRegionEdge, ByRef holeCount As Long)
    Dim i As Long, edge As TRegionEdge, kind As String, isHole As Boolean
    For i = 1 To contours.Count
        kind = contours.SegmentType(i)
        If Left$(kind, 8) = "CONTOUR_" Then
            isHole = (StrComp(contours.LoopRole(i), "Opening", vbTextCompare) = 0)
            If kind = "CONTOUR_CIRCLE" Then
                If isHole Then
                    AddCircle holes, holeCount, contours.StartX(i), contours.StartY(i), contours.Radius(i), True
                Else
                    AddCircle outer, outerCount, contours.StartX(i), contours.StartY(i), contours.Radius(i), False
                End If
            ElseIf kind = "CONTOUR_LINE" Or kind = "CONTOUR_ARC" Then
                edge.X1 = contours.StartX(i): edge.Y1 = contours.StartY(i)
                edge.X2 = contours.EndX(i): edge.Y2 = contours.EndY(i)
                edge.CenterX = 0#: edge.CenterY = 0#: edge.Sweep = 0#: edge.LoopID = 0
                If kind = "CONTOUR_ARC" Then
                    edge.CenterX = contours.CenterX(i): edge.CenterY = contours.CenterY(i)
                    edge.Sweep = contours.SweepAngle(i)
                    ValidateArc edge
                End If
                If isHole Then
                    AppendEdge holes, holeCount, edge
                Else
                    AppendEdge outer, outerCount, edge
                End If
            End If
        End If
    Next i
End Sub

' Восстанавливает границу по прямоугольным оболочкам бетонных элементов.
' Сразу удаляет полностью совпавшие внутренние рёбра, затем разделяет
' и удаляет частичные общие участки, включая Т-образные стыки.
Private Sub MeshBoundary(ByVal section As CSectionModel, ByRef boundary() As TRegionEdge, ByRef count As Long)
    Dim dictionary As Object
    Set dictionary = CreateObject("Scripting.Dictionary")
    Dim edges() As TRegionEdge, edgeCount As Long, i As Long, j As Long, k As Long
    Dim width As Double, height As Double, angle As Double, px(1 To 4) As Double, py(1 To 4) As Double
    Dim lx As Double, ly As Double, edge As TRegionEdge, key As String, reverseKey As String, oldIndex As Long
    For i = 1 To section.ConcreteCount
        If Not section.ConcreteBoundaryRectangle(i, width, height, angle) Then GeometryError "Не удалось получить границу бетонного элемента " & section.ConcreteID(i) & "."
        For j = 1 To 4
            Select Case j
                Case 1: lx = -width / 2#: ly = -height / 2#
                Case 2: lx = width / 2#: ly = -height / 2#
                Case 3: lx = width / 2#: ly = height / 2#
                Case 4: lx = -width / 2#: ly = height / 2#
            End Select
            px(j) = section.ConcreteX(i) + lx * Cos(angle) - ly * Sin(angle)
            py(j) = section.ConcreteY(i) + lx * Sin(angle) + ly * Cos(angle)
        Next j
        For j = 1 To 4
            k = j + 1: If k = 5 Then k = 1
            edge.X1 = px(j): edge.Y1 = py(j): edge.X2 = px(k): edge.Y2 = py(k)
            key = PointKey(edge.X1, edge.Y1) & "/" & PointKey(edge.X2, edge.Y2)
            reverseKey = PointKey(edge.X2, edge.Y2) & "/" & PointKey(edge.X1, edge.Y1)
            If dictionary.Exists(reverseKey) Then
                oldIndex = CLng(dictionary(reverseKey))
                edges(oldIndex).X2 = edges(oldIndex).X1: edges(oldIndex).Y2 = edges(oldIndex).Y1
                dictionary.Remove reverseKey
            ElseIf Not dictionary.Exists(key) Then
                AppendEdge edges, edgeCount, edge
                dictionary.Add key, edgeCount
            End If
        Next j
    Next i
    Dim remaining() As TRegionEdge, remainingCount As Long
    For i = 1 To edgeCount
        If modGeometryPrimitives.EdgeLength(edges(i)) > QUERY_TOLERANCE Then AppendEdge remaining, remainingCount, edges(i)
    Next i
    NormalizeBoundary remaining, remainingCount, boundary, count
    Dim ordered() As TRegionEdge, loopCount As Long
    StitchLoops boundary, count, ordered, loopCount, False
    CopyArray ordered, count, boundary
End Sub

' Готовит область непосредственно по классифицированным контурам формы,
' без искусственной сетки и привязки запроса к расчетной модели.
Public Function RegionFromContours(ByVal contours As CSectionContours) As CConcreteRegion
    If contours Is Nothing Then GeometryError "Контуры бетонной формы не переданы."
    Dim outer() As TRegionEdge, outerCount As Long, holes() As TRegionEdge, holeCount As Long
    ReadMaterialContours contours, outer, outerCount, holes, holeCount
    If outerCount = 0 Then GeometryError "У формы отсутствует наружный контур бетона."
    OrientLoops outer, outerCount, False
    Dim raw() As TRegionEdge, count As Long, i As Long
    CopyArray outer, outerCount, raw: count = outerCount
    If holeCount > 0 Then
        OrientLoops holes, holeCount, True
        For i = 1 To holeCount: AppendEdge raw, count, holes(i): Next i
    End If
    Dim boundary() As TRegionEdge, boundaryCount As Long
    NormalizeBoundary raw, count, boundary, boundaryCount
    Set RegionFromContours = MakeRegion(boundary, boundaryCount, "AuthoritativeContour", 0)
    If RegionFromContours.Area <= QUERY_TOLERANCE Then GeometryError "Бетонная область пуста после учета наружных границ и отверстий."
End Function

' Точные первые моменты площади по отрезкам и круговым дугам (формула Грина).
' Локальное начало снижает потерю точности при больших координатах модели.
Public Sub RegionCentroid(ByVal region As CConcreteRegion, ByRef x As Double, ByRef y As Double)
    RequireRegion region
    Dim edges() As TRegionEdge, count As Long, i As Long, edge As TRegionEdge
    region.CopyEdges edges, count
    If count = 0 Then GeometryError "Нельзя определить центр тяжести пустой области."
    Dim originX As Double, originY As Double, area As Double, momentX As Double, momentY As Double
    originX = edges(1).X1: originY = edges(1).Y1
    Dim radius As Double, a As Double, b As Double, sa As Double, sb As Double, ca As Double, cb As Double
    Dim integralCos2 As Double, integralSin2 As Double
    For i = 1 To count
        LSet edge = edges(i)
        edge.X1 = edge.X1 - originX: edge.X2 = edge.X2 - originX
        edge.Y1 = edge.Y1 - originY: edge.Y2 = edge.Y2 - originY
        edge.CenterX = edge.CenterX - originX: edge.CenterY = edge.CenterY - originY
        area = area + EdgeSignedArea(edge)
        If Abs(edge.Sweep) <= PARAM_TOLERANCE Then
            momentX = momentX + (edge.Y2 - edge.Y1) * (edge.X1 ^ 2 + edge.X1 * edge.X2 + edge.X2 ^ 2) / 6#
            momentY = momentY - (edge.X2 - edge.X1) * (edge.Y1 ^ 2 + edge.Y1 * edge.Y2 + edge.Y2 ^ 2) / 6#
        Else
            radius = Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2)
            a = modGeometryPrimitives.AngleOf(edge.X1 - edge.CenterX, edge.Y1 - edge.CenterY): b = a + edge.Sweep
            sa = Sin(a): sb = Sin(b): ca = Cos(a): cb = Cos(b)
            integralCos2 = edge.Sweep / 2# + (Sin(2# * b) - Sin(2# * a)) / 4#
            integralSin2 = edge.Sweep / 2# - (Sin(2# * b) - Sin(2# * a)) / 4#
            momentX = momentX + radius / 2# * (edge.CenterX ^ 2 * (sb - sa) + _
                2# * edge.CenterX * radius * integralCos2 + radius ^ 2 * (sb - sb ^ 3 / 3# - sa + sa ^ 3 / 3#))
            momentY = momentY + radius / 2# * (edge.CenterY ^ 2 * (ca - cb) + _
                2# * edge.CenterY * radius * integralSin2 + radius ^ 2 * (-cb + cb ^ 3 / 3# + ca - ca ^ 3 / 3#))
        End If
    Next i
    If area <= QUERY_TOLERANCE Then GeometryError "Нельзя определить центр тяжести области нулевой площади."
    x = originX + momentX / area: y = originY + momentY / area
End Sub

' При явно заданных отверстиях сетка задает только наружные границы.
' Ее внутренние пустоты не добавляются к достоверному списку Opening;
' отдельные наружные части и вогнутые участки границы сохраняются.
Private Sub KeepOuterMeshContours(ByRef edges() As TRegionEdge, ByRef count As Long)
    If count = 0 Then Exit Sub
    Dim loopAreas() As Double, outer() As TRegionEdge, outerCount As Long, i As Long
    ReDim loopAreas(1 To count)
    For i = 1 To count
        loopAreas(edges(i).LoopID) = loopAreas(edges(i).LoopID) + EdgeSignedArea(edges(i))
    Next i
    For i = 1 To count
        If loopAreas(edges(i).LoopID) > QUERY_TOLERANCE Then AppendEdge outer, outerCount, edges(i)
    Next i
    Dim boundary() As TRegionEdge, boundaryCount As Long
    NormalizeBoundary outer, outerCount, boundary, boundaryCount
    CopyArray boundary, boundaryCount, edges
    count = boundaryCount
End Sub

' Разделяет пересекающиеся и совпавшие рёбра и оставляет границы между бетоном
' и пустотой. Внутренние рёбра и повторные отверстия не добавляют площадь
' и не попадают в экспорт.
Private Sub NormalizeBoundary(ByRef raw() As TRegionEdge, ByVal rawCount As Long, _
        ByRef boundary() As TRegionEdge, ByRef boundaryCount As Long)
    Erase boundary: boundaryCount = 0
    Dim cuts() As Double, cutCount As Long, i As Long, j As Long, k As Long
    Dim part As TRegionEdge, x As Double, y As Double, tx As Double, ty As Double, length As Double
    Dim leftInside As Boolean, rightInside As Boolean, seen As Object, key As String
    Set seen = CreateObject("Scripting.Dictionary")
    For i = 1 To rawCount
        If modGeometryPrimitives.EdgeLength(raw(i)) <= QUERY_TOLERANCE Then GoTo NextEdge
        Erase cuts: cutCount = 0
        modGeometryPrimitives.AddScalar cuts, cutCount, 0#: modGeometryPrimitives.AddScalar cuts, cutCount, 1#
        For j = 1 To rawCount
            If i <> j Then modGeometryPrimitives.EdgeIntersectionCuts raw(i), raw(j), cuts, cutCount
        Next j
        SortScalars cuts, cutCount
        For k = 1 To cutCount - 1
            modGeometryPrimitives.SubEdge raw(i), cuts(k), cuts(k + 1), part
            If modGeometryPrimitives.EdgeLength(part) > QUERY_TOLERANCE Then
                modGeometryPrimitives.EdgePoint part, 0.5, x, y
                modGeometryPrimitives.EdgeTangent part, 0.5, tx, ty
                length = Sqr(tx * tx + ty * ty)
                leftInside = ContainsRaw(raw, rawCount, x - ty / length * QUERY_TOLERANCE * 8#, y + tx / length * QUERY_TOLERANCE * 8#, False)
                rightInside = ContainsRaw(raw, rawCount, x + ty / length * QUERY_TOLERANCE * 8#, y - tx / length * QUERY_TOLERANCE * 8#, False)
                If leftInside <> rightInside Then
                    If rightInside Then modGeometryPrimitives.ReverseEdge part
                    key = PointKey(part.X1, part.Y1) & "/" & PointKey(part.X2, part.Y2) & "/" & Format$(part.Sweep, "0.0000000000")
                    If Not seen.Exists(key) Then
                        AppendEdge boundary, boundaryCount, part
                        seen.Add key, True
                    End If
                End If
            End If
        Next k
NextEdge:
    Next i
End Sub

' Создаёт готовую область только из замкнутых границ. Вычисляет её площадь
' по формуле Грина с точным вкладом дуг и вычитанием отверстий.
Private Function MakeRegion(ByRef edges() As TRegionEdge, ByVal count As Long, _
        ByVal source As String, ByVal revision As Long) As CConcreteRegion
    Dim ordered() As TRegionEdge, loopCount As Long, area As Double, i As Long
    Dim loopAreas() As Double, loopOpening() As Boolean, originX() As Double, originY() As Double
    Dim seenLoop() As Boolean
    StitchLoops edges, count, ordered, loopCount, False
    If loopCount > 0 Then
        ReDim loopAreas(1 To loopCount)
        ReDim loopOpening(1 To loopCount)
        ReDim originX(1 To loopCount): ReDim originY(1 To loopCount): ReDim seenLoop(1 To loopCount)
    End If
    For i = 1 To count
        Dim contribution As Double
        If Not seenLoop(ordered(i).LoopID) Then
            originX(ordered(i).LoopID) = ordered(i).X1: originY(ordered(i).LoopID) = ordered(i).Y1
            seenLoop(ordered(i).LoopID) = True
        End If
        contribution = EdgeSignedArea(ordered(i), originX(ordered(i).LoopID), originY(ordered(i).LoopID))
        area = area + contribution
        loopAreas(ordered(i).LoopID) = loopAreas(ordered(i).LoopID) + contribution
    Next i
    For i = 1 To loopCount: loopOpening(i) = (loopAreas(i) < 0#): Next i
    If area < -QUERY_TOLERANCE Then GeometryError "Не удалось получить положительную площадь бетона. Проверьте наружные границы Outer и отверстия Opening, их положение и направление обхода."
    Dim result As CConcreteRegion
    Set result = New CConcreteRegion
    result.SetGeometry ordered, count, loopCount, GeomMax(area, 0#), source, revision, loopOpening
    Set MakeRegion = result
End Function

' Соединяет отрезки и дуги по их концам в замкнутые контуры. Разрыв вызывает ошибку.
' Касающиеся раздельные области обходит отдельно, сохраняя бетон слева по обходу.
Private Sub StitchLoops(ByRef edges() As TRegionEdge, ByVal count As Long, ByRef ordered() As TRegionEdge, _
        ByRef loopCount As Long, ByVal allowReverse As Boolean)
    Erase ordered: loopCount = 0
    If count = 0 Then Exit Sub
    ReDim ordered(1 To count)
    Dim used() As Boolean, i As Long, j As Long, position As Long, first As TRegionEdge, current As TRegionEdge
    Dim candidate As Long, reverseCandidate As Boolean, bestTurn As Double, turn As Double, testEdge As TRegionEdge
    ReDim used(1 To count)
    For i = 1 To count
        If Not used(i) Then
            loopCount = loopCount + 1
            LSet first = edges(i)
            LSet current = first
            used(i) = True
            Do
                position = position + 1: current.LoopID = loopCount
                LSet ordered(position) = current
                If modGeometryPrimitives.SamePoint(current.X2, current.Y2, first.X1, first.Y1) Then Exit Do
                candidate = 0: bestTurn = -1E+100: reverseCandidate = False
                For j = 1 To count
                    If Not used(j) Then
                        LSet testEdge = edges(j)
                        If modGeometryPrimitives.SamePoint(current.X2, current.Y2, testEdge.X1, testEdge.Y1) Then
                            turn = EdgeTurn(current, testEdge)
                            If turn > bestTurn Then candidate = j: bestTurn = turn: reverseCandidate = False
                        ElseIf allowReverse And modGeometryPrimitives.SamePoint(current.X2, current.Y2, testEdge.X2, testEdge.Y2) Then
                            modGeometryPrimitives.ReverseEdge testEdge
                            turn = EdgeTurn(current, testEdge)
                            If turn > bestTurn Then candidate = j: bestTurn = turn: reverseCandidate = True
                        End If
                    End If
                Next j
                If candidate = 0 Then GeometryError "Контур не замкнут возле точки X=" & CStr(current.X2) & "; Y=" & CStr(current.Y2) & "."
                LSet current = edges(candidate)
                If reverseCandidate Then modGeometryPrimitives.ReverseEdge current
                used(candidate) = True
                If position > count Then GeometryError "Не удалось собрать замкнутые контуры из заданных сегментов. Проверьте разрывы, пересечения и повторяющиеся участки границ."
            Loop
        End If
    Next i
End Sub

' Задаёт принятое направление обхода для известных наружных контуров и отверстий.
' Их назначение уже задано; менять его по знаку площади исходного контура нельзя.
Private Sub OrientLoops(ByRef edges() As TRegionEdge, ByVal count As Long, ByVal opening As Boolean)
    Dim ordered() As TRegionEdge, loops As Long, loopID As Long, i As Long, area As Double
    StitchLoops edges, count, ordered, loops, True
    For loopID = 1 To loops
        area = SignedLoopArea(ordered, count, loopID)
        If Abs(area) <= QUERY_TOLERANCE Then GeometryError "Заданный контур имеет нулевую площадь."
        If (area > 0# And opening) Or (area < 0# And Not opening) Then
            For i = 1 To count
                If ordered(i).LoopID = loopID Then modGeometryPrimitives.ReverseEdge ordered(i)
            Next i
        End If
    Next loopID
    CopyArray ordered, count, edges
End Sub

' Вычисляет площадь одного замкнутого контура со знаком его обхода,
' учитывая точный вклад дуг.
Private Function SignedLoopArea(ByRef edges() As TRegionEdge, ByVal count As Long, ByVal loopID As Long) As Double
    Dim i As Long, originX As Double, originY As Double, found As Boolean
    For i = 1 To count
        If edges(i).LoopID = loopID Then
            If Not found Then
                originX = edges(i).X1: originY = edges(i).Y1: found = True
            End If
            SignedLoopArea = SignedLoopArea + EdgeSignedArea(edges(i), originX, originY)
        End If
    Next i
End Function

' Интеграл (x*dy-y*dx)/2 для отрезка или дуги окружности. Радиусы/углы не
' переводятся в пользовательские единицы, поэтому результат всегда мм2.
Private Function EdgeSignedArea(ByRef edge As TRegionEdge, Optional ByVal originX As Double = 0#, _
        Optional ByVal originY As Double = 0#) As Double
    EdgeSignedArea = 0.5 * ((edge.X1 - originX) * (edge.Y2 - originY) - _
        (edge.X2 - originX) * (edge.Y1 - originY))
    If Abs(edge.Sweep) > PARAM_TOLERANCE Then
        Dim radiusSquared As Double, segmentAngle As Double, square As Double
        radiusSquared = (edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2
        ' Площадь между хордой и дугой без вычитания больших глобальных
        ' координат; ряд сохраняет точность при почти касательном отсечении.
        If Abs(edge.Sweep) < 0.01 Then
            square = edge.Sweep * edge.Sweep
            segmentAngle = edge.Sweep * square * (1# / 6# - square / 120# + square * square / 5040#)
        Else
            segmentAngle = edge.Sweep - Sin(edge.Sweep)
        End If
        EdgeSignedArea = EdgeSignedArea + 0.5 * radiusSquared * segmentAngle
    End If
End Function

' Проверяет попадание точки по числу обходов границы вокруг неё.
' Рабочий массив может содержать пропуски после удаления внутренних рёбер.
Private Function ContainsRaw(ByRef edges() As TRegionEdge, ByVal count As Long, _
        ByVal x As Double, ByVal y As Double, ByVal includeBoundary As Boolean) As Boolean
    Dim i As Long, u As Double, winding As Long, crossX As Double, dy As Double, f1 As Double, f2 As Double
    Dim cuts() As Double, cutCount As Long, j As Long, px As Double, py As Double, tx As Double, ty As Double
    Dim countedStart As Boolean, countedEnd As Boolean
    For i = 1 To count
        If includeBoundary Then
            If modGeometryPrimitives.PointParameter(edges(i), x, y, u) Then ContainsRaw = True: Exit Function
        End If
        If Abs(edges(i).Sweep) <= PARAM_TOLERANCE Then
            ' На стыке дуги с отрезком считаем общий конец одной вершиной.
            ' Малую погрешность sin/cos учитываем тем же линейным допуском,
            ' который применяется в остальных проверках дуг.
            f1 = edges(i).Y1 - y: f2 = edges(i).Y2 - y
            If Abs(f1) <= QUERY_TOLERANCE Then f1 = 0#
            If Abs(f2) <= QUERY_TOLERANCE Then f2 = 0#
            If (f1 <= 0# And f2 > 0#) Or (f2 <= 0# And f1 > 0#) Then
                crossX = edges(i).X1 - f1 * (edges(i).X2 - edges(i).X1) / (f2 - f1)
                If crossX > x Then winding = winding + Sgn(edges(i).Y2 - edges(i).Y1)
            End If
        Else
            Erase cuts: cutCount = 0
            modGeometryPrimitives.HalfPlaneCuts edges(i), 0#, 1#, y, cuts, cutCount
            ' Угловой параметр может выйти за [0,1] при луче в пределах
            ' линейного допуска от конца дуги. Сохраняем саму общую вершину.
            If Abs(edges(i).Y1 - y) <= QUERY_TOLERANCE Then modGeometryPrimitives.AddScalar cuts, cutCount, 0#
            If Abs(edges(i).Y2 - y) <= QUERY_TOLERANCE Then modGeometryPrimitives.AddScalar cuts, cutCount, 1#
            countedStart = False: countedEnd = False
            For j = 1 To cutCount
                u = cuts(j)
                If u >= 0# And u <= 1# Then
                    ' Если точка близка к общей вершине дуги и отрезка, проверяем её
                    ' по тому же линейному допуску, а не по отдельному угловому допуску.
                    If u * modGeometryPrimitives.EdgeLength(edges(i)) <= QUERY_TOLERANCE Then u = 0#
                    If (1# - u) * modGeometryPrimitives.EdgeLength(edges(i)) <= QUERY_TOLERANCE Then u = 1#
                    If u = 0# Then
                        If countedStart Then GoTo NextCrossing
                        countedStart = True
                    ElseIf u = 1# Then
                        If countedEnd Then GoTo NextCrossing
                        countedEnd = True
                    End If
                    modGeometryPrimitives.EdgePoint edges(i), u, px, py
                    modGeometryPrimitives.EdgeTangent edges(i), u, tx, ty
                    dy = ty
                    ' Как у отрезков, учитываем нижний конец и исключаем
                    ' верхний. Просто [0,1) для дуги неверно при стыке с
                    ' направленным вниз отрезком и может заполнить пустоту.
                    If px > x And Abs(dy) > QUERY_TOLERANCE Then
                        If (u > PARAM_TOLERANCE And u < 1# - PARAM_TOLERANCE) Or _
                                (u <= PARAM_TOLERANCE And dy > 0#) Or _
                                (u >= 1# - PARAM_TOLERANCE And dy < 0#) Then winding = winding + Sgn(dy)
                    End If
                End If
NextCrossing:
            Next j
        End If
    Next i
    ContainsRaw = (winding > 0)
End Function

' Возвращает угол поворота касательной при переходе между сегментами в общей вершине.
Private Function EdgeTurn(ByRef previous As TRegionEdge, ByRef following As TRegionEdge) As Double
    Dim ax As Double, ay As Double, bx As Double, by As Double
    modGeometryPrimitives.EdgeTangent previous, 1#, ax, ay: modGeometryPrimitives.EdgeTangent following, 0#, bx, by
    EdgeTurn = modGeometryPrimitives.AngleOf(ax * bx + ay * by, ax * by - ay * bx)
End Function

' Проверяет, что концы дуги соответствуют её центру и углу обхода со знаком.
' Недопустимый явно заданный контур вызывает ошибку, а не заменяется границами сетки.
Private Sub ValidateArc(ByRef edge As TRegionEdge)
    Dim x As Double, y As Double
    If modGeometryPrimitives.EdgeLength(edge) <= QUERY_TOLERANCE Then GeometryError "Заданная дуга имеет нулевую длину."
    modGeometryPrimitives.EdgePoint edge, 1#, x, y
    If Not modGeometryPrimitives.SamePoint(x, y, edge.X2, edge.Y2) Then GeometryError "Конец дуги не соответствует ее центру и углу."
End Sub

' Добавляет невырожденный отрезок в рабочий массив границы.
Private Sub AddLine(ByRef edges() As TRegionEdge, ByRef count As Long, ByVal x1 As Double, _
        ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
    Dim edge As TRegionEdge
    edge.X1 = x1: edge.Y1 = y1: edge.X2 = x2: edge.Y2 = y2
    If modGeometryPrimitives.EdgeLength(edge) > QUERY_TOLERANCE Then AppendEdge edges, count, edge
End Sub

' Добавляет четыре точных четверти окружности с заданной ролью замкнутого контура.
Private Sub AddCircle(ByRef edges() As TRegionEdge, ByRef count As Long, ByVal cx As Double, _
        ByVal cy As Double, ByVal radius As Double, ByVal opening As Boolean)
    If radius <= QUERY_TOLERANCE Then GeometryError "Радиус окружности контура должен быть положительным."
    Dim i As Long, edge As TRegionEdge, sign As Double
    sign = 1#: If opening Then sign = -1#
    For i = 0 To 3
        edge.CenterX = cx: edge.CenterY = cy: edge.Sweep = sign * GEOM_PI / 2#
        edge.X1 = cx + radius * Cos(sign * i * GEOM_PI / 2#)
        edge.Y1 = cy + radius * Sin(sign * i * GEOM_PI / 2#)
        edge.X2 = cx + radius * Cos(sign * (i + 1) * GEOM_PI / 2#)
        edge.Y2 = cy + radius * Sin(sign * (i + 1) * GEOM_PI / 2#)
        AppendEdge edges, count, edge
    Next i
End Sub

' Добавляет сегмент в массив, при нехватке места увеличивая его размер вдвое.
' Это исключает расширение массива на каждом ребре при обходе большой сетки.
Private Sub AppendEdge(ByRef edges() As TRegionEdge, ByRef count As Long, ByRef edge As TRegionEdge)
    If count = 0 Then
        ReDim edges(1 To 16)
    ElseIf count = UBound(edges) Then
        ReDim Preserve edges(1 To UBound(edges) * 2)
    End If
    count = count + 1
    LSet edges(count) = edge
End Sub

' Копирует только заполненную часть одномерного массива аналитических ребер.
Private Sub CopyArray(ByRef source() As TRegionEdge, ByVal count As Long, ByRef destination() As TRegionEdge)
    Erase destination
    If count = 0 Then Exit Sub
    ReDim destination(1 To count)
    Dim i As Long
    For i = 1 To count
        LSet destination(i) = source(i)
    Next i
End Sub

' Сортирует небольшой массив параметров точек пересечения без создания отдельного объекта на
' точку.
Private Sub SortScalars(ByRef values() As Double, ByVal count As Long)
    Dim i As Long, j As Long, value As Double
    For i = 2 To count
        value = values(i): j = i - 1
        Do While j >= 1
            If values(j) <= value Then Exit Do
            values(j + 1) = values(j): j = j - 1
        Loop
        values(j + 1) = value
    Next i
End Sub

' Строит ключ для быстрого поиска полностью совпавших внутренних рёбер.
' Попадание в одну группу поиска не заменяет точную проверку границ
' и их частичных совпадений.
Private Function PointKey(ByVal x As Double, ByVal y As Double) As String
    PointKey = Format$(x / QUERY_TOLERANCE, "0") & ":" & Format$(y / QUERY_TOLERANCE, "0")
End Function

' Проверяет, что объект области передан. Отсутствие объекта считается ошибкой,
' а не пустой бетонной областью.
Private Sub RequireRegion(ByVal region As CConcreteRegion)
    If region Is Nothing Then GeometryError "Готовая геометрическая область не передана."
End Sub

' Сообщает о недопустимых геометрических данных. Вызывающий инженерный расчёт
' сохраняет причину в ResultMeta и определяет внешний статус общими правилами.
Private Sub GeometryError(ByVal description As String, Optional ByVal code As String = "GEOMETRY_INVALID", _
        Optional ByVal entityID As String = "", Optional ByVal relatedEntityID As String = "", _
        Optional ByVal detail1 As Double = 0#, Optional ByVal detail2 As Double = 0#)
    Set mDiagnostic = New CGeometryDiagnostic
    mDiagnostic.Initialize code, entityID, relatedEntityID, description, detail1, detail2
    Err.Raise vbObjectError + 5401, "CSectionGeometryQuery", description
End Sub

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================

' ДЛЯ ТЕСТОВ: возвращает число фактических построений области, без учёта выдачи готовых копий.
Friend Property Get DomainBuildCount() As Long
    DomainBuildCount = mDomainBuildCount
End Property
