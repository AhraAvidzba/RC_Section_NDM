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
' Общие запросы к бетонной области, пересечения и топология
' ==========================================================================
' Движок работает с аналитическими линиями/дугами и с границами реальных
' бетонных элементов CSectionModel. Внутренние общие ребра удаляются, оставшаяся
' граница сшивается в замкнутые outer/opening loops. Нормативный потребитель
' передает только геометрические ограничения; здесь нет СП, арматурных групп,
' НДС, solver/search, result-status policy и Excel. Геометрия конечной области
' и ее площадь создаются вместе и далее не восстанавливаются из числа Ar.
' Устойчивость не переводится на этот движок: ее A/I/проекции сетки неизменны.

Private Const QUERY_TOLERANCE As Double = 0.000001 ' Прежний допуск склейки бетонных интервалов, мм.
Private Const PARAM_TOLERANCE As Double = 0.0000000001 ' Только безразмерный параметр аналитического сегмента.
Private mSection As CSectionModel
Private mRevision As Long
Private mDomain As CConcreteRegion ' Неизменная область одной модели/query-session; потребитель получает clone.
Private mDomainBuildCount As Long ' Число успешных подготовок за жизнь query, включая смену контекста.

' Привязывает query-session к модели. Достоверные контуры берутся из Contours
' генератора/импортера, а неизвестная граница восстанавливается по КЭ.
Public Sub Initialize(ByVal section As CSectionModel)
    If section Is Nothing Then GeometryError "Модель сечения не передана."
    If section.ConcreteCount = 0 Then GeometryError "В модели нет бетонных элементов."
    Set mSection = section
    mRevision = section.Revision
    Set mDomain = Nothing
End Sub

' Привязывает общий запрос к фактической модели, сохраняя только ее неизменную
' подготовку. Новый объект/Revision сбрасывает область; LC-данных здесь нет.
Friend Sub EnsureSection(ByVal section As CSectionModel)
    If section Is Nothing Then GeometryError "Модель сечения не передана."
    If Not section Is mSection Then
        Initialize section
    ElseIf section.Revision <> mRevision Then
        mRevision = section.Revision
        Set mDomain = Nothing
    End If
End Sub

' ДЛЯ ТЕСТОВ: фактические подготовки, а не количество чтений clone.
Friend Property Get DomainBuildCount() As Long
    DomainBuildCount = mDomainBuildCount
End Property

' Возвращает независимую копию всей бетонной области, включая отверстия.
' Кэш локален для session и перестраивается при изменении геометрии модели.
Public Function ConcreteDomain() As CConcreteRegion
    EnsureDomain
    Set ConcreteDomain = mDomain.Clone
End Function

' Возвращает точные крайние проекции готовой области, включая экстремумы дуг.
' Это API трещин; прежние сеточные проекции устойчивости остаются в properties.
Public Sub ProjectionBounds(ByVal region As CConcreteRegion, ByVal nx As Double, ByVal ny As Double, _
        ByRef minimum As Double, ByRef maximum As Double)
    RequireRegion region
    Dim norm As Double
    norm = Sqr(nx * nx + ny * ny)
    If norm <= GEOM_TOLERANCE Then GeometryError "Направление проекции имеет нулевую длину."
    nx = nx / norm: ny = ny / norm
    If region.SegmentCount = 0 Then GeometryError "У пустой области нет крайних проекций."
    Dim edges() As TRegionEdge, count As Long, i As Long, radius As Double, u As Double, projection As Double
    region.CopyEdges edges, count
    minimum = 1E+100: maximum = -1E+100
    For i = 1 To count
        projection = nx * edges(i).X1 + ny * edges(i).Y1
        minimum = GeomMin(minimum, projection): maximum = GeomMax(maximum, projection)
        projection = nx * edges(i).X2 + ny * edges(i).Y2
        minimum = GeomMin(minimum, projection): maximum = GeomMax(maximum, projection)
        If Abs(edges(i).Sweep) > PARAM_TOLERANCE Then
            radius = Sqr((edges(i).X1 - edges(i).CenterX) ^ 2 + (edges(i).Y1 - edges(i).CenterY) ^ 2)
            projection = nx * edges(i).CenterX + ny * edges(i).CenterY
            If PointParameter(edges(i), edges(i).CenterX + radius * nx, edges(i).CenterY + radius * ny, u) Then maximum = GeomMax(maximum, projection + radius)
            If PointParameter(edges(i), edges(i).CenterX - radius * nx, edges(i).CenterY - radius * ny, u) Then minimum = GeomMin(minimum, projection - radius)
        End If
    Next i
End Sub

' Строит область заданного замкнутого многоугольника без обращения к модели.
' Координаты уже внутренние; orientation нормализуется общим topology API.
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

' Строит точную окружность четырьмя дугами; промежуточная полигональная
' аппроксимация не используется ни для площади, ни для последующего clipping.
Public Function CircleRegion(ByVal centerX As Double, ByVal centerY As Double, ByVal radius As Double) As CConcreteRegion
    If radius <= QUERY_TOLERANCE Then GeometryError "Радиус контура должен быть положительным."
    Dim edges() As TRegionEdge, count As Long
    AddCircle edges, count, centerX, centerY, radius, False
    Set CircleRegion = MakeRegion(edges, count, "ExplicitContour", 0)
End Function

' Принимает неупорядоченные аналитические ребра достоверного контура.
' Проверяет замкнутость, вырождение, пересечения и дубли; роль каждого контура
' определяется вложенностью, а не направлением обхода исходного объекта.
' Поэтому Region с отверстиями и полилинии используют один topology-контракт.
Public Function BoundaryRegion(ByRef edges() As TRegionEdge, ByVal count As Long, _
        Optional ByVal source As String = "ValidatedContour") As CConcreteRegion
    If count <= 0 Then GeometryError "Заданный контур не содержит граничных сегментов."
    Dim i As Long, j As Long, ordered() As TRegionEdge, loops As Long
    For i = 1 To count
        If EdgeLength(edges(i)) <= QUERY_TOLERANCE Then GeometryError "Заданный контур содержит ребро нулевой длины."
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
                EdgePoint ordered(i), 0.5, x, y
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
                ' Временный замкнутый контур должен иметь положительный winding для
                ' геометрической проверки, независимо от исходного обхода.
                If SignedLoopArea(ordered, count, otherLoop) < 0# Then
                    For j = 1 To oneCount: ReverseEdge oneLoop(j): Next j
                End If
                If ContainsRaw(oneLoop, oneCount, x, y, False) Then depth = depth + 1
            End If
        Next otherLoop
        If (area > 0# And depth Mod 2 = 1) Or (area < 0# And depth Mod 2 = 0) Then
            For i = 1 To count
                If ordered(i).LoopID = loopID Then ReverseEdge ordered(i)
            Next i
        End If
    Next loopID
    Set BoundaryRegion = MakeRegion(ordered, count, source, 0)
End Function

' Проверяет, что отдельные аналитические контуры простые и не касаются друг
' друга. Такое касание не исправляется молча при импорте; пользователь
' получает ошибку исходного контура до публикации новой модели.
Private Sub ValidateSimpleBoundary(ByRef edges() As TRegionEdge, ByVal count As Long)
    Dim i As Long, j As Long, k As Long, cuts() As Double, cutCount As Long
    Dim u As Double, x As Double, y As Double, adjacent As Boolean
    For i = 1 To count
        For j = i + 1 To count
            adjacent = False
            If edges(i).LoopID = edges(j).LoopID Then
                adjacent = SamePoint(edges(i).X2, edges(i).Y2, edges(j).X1, edges(j).Y1) Or _
                    SamePoint(edges(j).X2, edges(j).Y2, edges(i).X1, edges(i).Y1)
            End If
            Erase cuts: cutCount = 0
            EdgeIntersectionCuts edges(i), edges(j), cuts, cutCount
            For k = 1 To cutCount
                If Not adjacent Or (cuts(k) > PARAM_TOLERANCE And cuts(k) < 1# - PARAM_TOLERANCE) Then _
                    GeometryError "Заданная геометрия содержит пересечение, касание независимых замкнутых контуров или дублирующее ребро."
            Next k
            Erase cuts: cutCount = 0
            EdgeIntersectionCuts edges(j), edges(i), cuts, cutCount
            For k = 1 To cutCount
                If Not adjacent Or (cuts(k) > PARAM_TOLERANCE And cuts(k) < 1# - PARAM_TOLERANCE) Then _
                    GeometryError "Заданная геометрия содержит пересечение, касание независимых замкнутых контуров или дублирующее ребро."
            Next k
            ' Полностью совпадающие сегменты могут иметь только концевые
            ' пересечения. Проверяем их середины без аппроксимации дуг.
            EdgePoint edges(i), 0.5, x, y
            If PointParameter(edges(j), x, y, u) Then
                If u > PARAM_TOLERANCE And u < 1# - PARAM_TOLERANCE Then _
                    GeometryError "Заданный контур содержит совпадающие граничные сегменты."
            End If
        Next j
    Next i
End Sub

' Возвращает семантическую роль готового ориентированного контура. Решение
' принимает общий движок; importer/writer не угадывают роль по тексту ошибки.
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
                        EdgePoint edges(i), 0.5, sampleX, sampleY
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

' Находит ближайший охватывающий наружный контур по вложенности. Наименьшая
' положительная площадь отличает отдельную бетонную часть внутри проема
' от окружающего бетона и правильно назначает владельца ее отверстиям.
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

' Объединяет независимые достоверные объекты материала, сохраняя отверстия
' каждого Region. Пересечение/дубли объектов являются ошибкой импорта,
' а не командой молча исправить чертеж через геометрический union.
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

' Вычитает только геометрически допустимое отдельное opening. В отличие
' от обычного clipping явно заданное отверстие за пределами материала или
' поверх другой пустоты не обрезается молча и не скрывается fallback-ом.
Public Function SubtractContainedRegion(ByVal outer As CConcreteRegion, ByVal opening As CConcreteRegion) As CConcreteRegion
    RequireRegion outer: RequireRegion opening
    If BoundariesMeet(outer, opening) Then GeometryError "Контур отверстия пересекает или касается границы материала либо другой пустоты."
    If Not AllMaterialSamplesInside(opening, outer) Then GeometryError "Заданное отверстие находится вне бетонной области."
    If AnyBoundarySampleInside(outer, opening) Then GeometryError "Заданное отверстие накрывает уже существующую пустоту или отдельную компоненту бетона."
    Set SubtractContainedRegion = SubtractRegion(outer, opening)
End Function

' Проверяет аналитические пересечения двух готовых границ, включая концы
' и совпадающие сегменты. Единый владелец операции - geometry-query.
Private Function BoundariesMeet(ByVal first As CConcreteRegion, ByVal second As CConcreteRegion) As Boolean
    Dim a() As TRegionEdge, b() As TRegionEdge, na As Long, nb As Long, i As Long, j As Long
    Dim cuts() As Double, count As Long
    first.CopyEdges a, na: second.CopyEdges b, nb
    For i = 1 To na
        For j = 1 To nb
            Erase cuts: count = 0
            EdgeIntersectionCuts a(i), b(j), cuts, count
            If count > 0 Then BoundariesMeet = True: Exit Function
        Next j
    Next i
End Function

' Возвращает True, если хотя бы один материальный участок первой области
' находится во второй. Проба слева от точного ребра учитывает также контуры
' пустот, не используя центр тяжести, который может лежать вне материала.
Private Function MaterialSampleInside(ByVal first As CConcreteRegion, ByVal second As CConcreteRegion) As Boolean
    Dim edges() As TRegionEdge, count As Long, i As Long, x As Double, y As Double, tx As Double, ty As Double, length As Double
    first.CopyEdges edges, count
    For i = 1 To count
        EdgePoint edges(i), 0.5, x, y: EdgeTangent edges(i), 0.5, tx, ty
        length = Sqr(tx * tx + ty * ty)
        If ContainsPoint(second, x - ty / length * QUERY_TOLERANCE * 8#, y + tx / length * QUERY_TOLERANCE * 8#, False) Then _
            MaterialSampleInside = True: Exit Function
    Next i
End Function

' После проверки отсутствия пересечений требует принадлежность всех замкнутых контуров
' opening материалу. Проверка каждого сегмента также охватывает несколько
' независимых компонент и не опирается на одну произвольную контрольную точку.
Private Function AllMaterialSamplesInside(ByVal first As CConcreteRegion, ByVal second As CConcreteRegion) As Boolean
    Dim edges() As TRegionEdge, count As Long, i As Long, x As Double, y As Double
    first.CopyEdges edges, count
    If count = 0 Then Exit Function
    For i = 1 To count
        EdgePoint edges(i), 0.5, x, y
        If Not ContainsPoint(second, x, y, False) Then Exit Function
    Next i
    AllMaterialSamplesInside = True
End Function

' Проверяет, что вычитаемая область не накрывает существующую материальную
' границу. Это отдельно выявляет дубли opening и отверстие поверх пустоты.
Private Function AnyBoundarySampleInside(ByVal first As CConcreteRegion, ByVal second As CConcreteRegion) As Boolean
    Dim edges() As TRegionEdge, count As Long, i As Long, x As Double, y As Double
    first.CopyEdges edges, count
    For i = 1 To count
        EdgePoint edges(i), 0.5, x, y
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
        ReverseEdge holes(i)
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
    For i = 1 To count: perimeter = perimeter + EdgeLength(edges(i)): Next i
    second.CopyEdges edges, count
    For i = 1 To count: perimeter = perimeter + EdgeLength(edges(i)): Next i
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
    Dim raw() As TRegionEdge, rawCount As Long, output() As TRegionEdge, outputCount As Long
    region.CopyEdges raw, rawCount
    Dim cuts() As Double, cutCount As Long, lineCuts() As Double, lineCount As Long
    Dim i As Long, j As Long, x As Double, y As Double, piece As TRegionEdge
    For i = 1 To rawCount
        Erase cuts: cutCount = 0
        AddScalar cuts, cutCount, 0#: AddScalar cuts, cutCount, 1#
        HalfPlaneCuts raw(i), nx, ny, bound, cuts, cutCount
        SortScalars cuts, cutCount
        For j = 1 To cutCount
            EdgePoint raw(i), cuts(j), x, y
            If Abs(nx * x + ny * y - bound) <= QUERY_TOLERANCE Then _
                AddScalar lineCuts, lineCount, ny * x - nx * y
        Next j
        For j = 1 To cutCount - 1
            EdgePoint raw(i), 0.5 * (cuts(j) + cuts(j + 1)), x, y
            If nx * x + ny * y > bound + QUERY_TOLERANCE Then
                SubEdge raw(i), cuts(j), cuts(j + 1), piece
                AppendEdge output, outputCount, piece
            End If
        Next j
    Next i

    ' Новые ребра линии отсечения замыкают материал слева. Проверка смещенной
    ' внутрь точки отличает касание от участка с ненулевой площадью.
    SortScalars lineCuts, lineCount
    Dim s1 As Double, s2 As Double, mid As Double
    For i = 1 To lineCount - 1
        s1 = lineCuts(i): s2 = lineCuts(i + 1): mid = 0.5 * (s1 + s2)
        x = nx * bound + ny * mid: y = ny * bound - nx * mid
        If ContainsRaw(raw, rawCount, x + nx * QUERY_TOLERANCE * 8#, y + ny * QUERY_TOLERANCE * 8#, False) Then
            AddLine output, outputCount, nx * bound + ny * s1, ny * bound - nx * s1, _
                nx * bound + ny * s2, ny * bound - nx * s2
        End If
    Next i
    Set ClipHalfPlane = MakeRegion(output, outputCount, region.Source, region.SectionRevision)
    ClipHalfPlane.SetOwner region.OwnerID, region.Standard, region.AnchorID
End Function

' Применяет один общий набор линейных ограничений: nMin <= n <= nMax,
' tMin <= t <= tMax. Нормативные значения и смысл границ задает потребитель.
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

' Оставляет только интервалы материала, проходящие через опорную линию
' n=seedPlane. Для каждой поперечной координаты t выбирается тот интервал,
' который содержит эту линию: перейти через пустоту к дальней стенке нельзя.
' Граница строится из точных исходных сегментов и линий смены интервала,
' а не из ограниченного набора пробных нормалей или прямоугольной оболочки.
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
        AddScalar transverseCuts, transverseCount, -ny * raw(i).X1 + nx * raw(i).Y1
        AddScalar transverseCuts, transverseCount, -ny * raw(i).X2 + nx * raw(i).Y2
        If Abs(raw(i).Sweep) > PARAM_TOLERANCE Then
            radius = Sqr((raw(i).X1 - raw(i).CenterX) ^ 2 + (raw(i).Y1 - raw(i).CenterY) ^ 2)
            If PointParameter(raw(i), raw(i).CenterX - ny * radius, raw(i).CenterY + nx * radius, u) Then _
                AddScalar transverseCuts, transverseCount, -ny * raw(i).CenterX + nx * raw(i).CenterY + radius
            If PointParameter(raw(i), raw(i).CenterX + ny * radius, raw(i).CenterY - nx * radius, u) Then _
                AddScalar transverseCuts, transverseCount, -ny * raw(i).CenterX + nx * raw(i).CenterY - radius
        End If
        Erase cuts: cutCount = 0
        HalfPlaneCuts raw(i), nx, ny, seedPlane, cuts, cutCount
        For j = 1 To cutCount
            EdgePoint raw(i), cuts(j), x, y
            AddScalar transverseCuts, transverseCount, -ny * x + nx * y
        Next j
    Next i
    SortScalars transverseCuts, transverseCount
    Dim piece As TRegionEdge, tangentX As Double, tangentY As Double, length As Double
    Dim leftInside As Boolean, rightInside As Boolean, seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    ' Смена выбранного интервала возможна только на проекциях вершин,
    ' экстремумов дуг и пересечений с опорной линией. Сами дуги не дробятся
    ' на линейную аппроксимацию и сохраняют аналитический вклад в площадь.
    For i = 1 To rawCount
        Erase cuts: cutCount = 0
        AddScalar cuts, cutCount, 0#: AddScalar cuts, cutCount, 1#
        For j = 1 To transverseCount
            HalfPlaneCuts raw(i), -ny, nx, transverseCuts(j), cuts, cutCount
        Next j
        SortScalars cuts, cutCount
        For j = 1 To cutCount - 1
            SubEdge raw(i), cuts(j), cuts(j + 1), piece
            If EdgeLength(piece) > QUERY_TOLERANCE Then
                EdgePoint piece, 0.5, x, y
                EdgeTangent piece, 0.5, tangentX, tangentY
                length = Sqr(tangentX * tangentX + tangentY * tangentY)
                leftInside = InMaterialIntervalThroughPlane(region, nx, ny, seedPlane, _
                    x - tangentY / length * QUERY_TOLERANCE * 8#, y + tangentX / length * QUERY_TOLERANCE * 8#)
                rightInside = InMaterialIntervalThroughPlane(region, nx, ny, seedPlane, _
                    x + tangentY / length * QUERY_TOLERANCE * 8#, y - tangentX / length * QUERY_TOLERANCE * 8#)
                If leftInside <> rightInside Then
                    If rightInside Then ReverseEdge piece
                    AppendUniqueBoundary boundary, boundaryCount, piece, seen
                End If
            End If
        Next j
    Next i
    ' Замыкаем границы тени пустот. Это настоящие линии перехода между
    ' доступной стенкой и недоступным материалом, а не мосты сквозь полость.
    Dim intervals As Variant, nCuts() As Double, nCount As Long, n1 As Double, n2 As Double, t As Double
    For i = 1 To transverseCount
        t = transverseCuts(i)
        intervals = LineIntervals(region, -ny * t, nx * t, nx, ny)
        If Not IsEmpty(intervals) Then
            Erase nCuts: nCount = 0
            For j = LBound(intervals, 1) To UBound(intervals, 1)
                AddScalar nCuts, nCount, intervals(j, 1): AddScalar nCuts, nCount, intervals(j, 2)
            Next j
            ' На совпавшей грани opening общий interval включает boundary.
            ' Сохраняем также все реальные узлы пересечения: иначе можно
            ' ошибочно продолжить замыкающий отрезок за край пустоты.
            For k = 1 To rawCount
                Erase cuts: cutCount = 0
                HalfPlaneCuts raw(k), -ny, nx, t, cuts, cutCount
                AddScalar cuts, cutCount, 0#: AddScalar cuts, cutCount, 1#
                For j = 1 To cutCount
                    EdgePoint raw(k), cuts(j), x, y
                    If Abs(-ny * x + nx * y - t) <= QUERY_TOLERANCE Then _
                        AddScalar nCuts, nCount, nx * x + ny * y
                Next j
            Next k
            AddScalar nCuts, nCount, seedPlane
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
                        If rightInside Then ReverseEdge piece
                        AppendUniqueBoundary boundary, boundaryCount, piece, seen
                    End If
                End If
            Next j
        End If
    Next i
    Dim normalized() As TRegionEdge, normalizedCount As Long
    ' Граница тени может частично совпасть с исходной вертикальной гранью.
    ' Общая нормализация разбивает такие наложения и оставляет одно ребро.
    NormalizeBoundary boundary, boundaryCount, normalized, normalizedCount
    Set MaterialRegionThroughPlane = MakeRegion(normalized, normalizedCount, region.Source, region.SectionRevision)
    MaterialRegionThroughPlane.SetOwner region.OwnerID, region.Standard, region.AnchorID
End Function

' Проверяет принадлежность именно интервалу опорной линии, а не просто
' всей бетонной области. Запрос одинаков для прямых и криволинейных пустот.
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

' Сохраняет одно ориентированное граничное ребро. Совпадение исходного
' ребра с новой границей не должно создавать двойной участок и площадь.
Private Sub AppendUniqueBoundary(ByRef edges() As TRegionEdge, ByRef count As Long, _
        ByRef edge As TRegionEdge, ByVal seen As Object)
    Dim key As String, i As Long, u1 As Double, u2 As Double
    If Abs(edge.Sweep) <= PARAM_TOLERANCE Then
        For i = 1 To count
            If Abs(edges(i).Sweep) <= PARAM_TOLERANCE Then
                If PointParameter(edges(i), edge.X1, edge.Y1, u1) And _
                        PointParameter(edges(i), edge.X2, edge.Y2, u2) Then
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

' Проверяет принадлежность точной области. Boundary=True включает точки на
' outer/opening границах; строгое попадание внутрь пустоты всегда исключается.
Public Function ContainsPoint(ByVal region As CConcreteRegion, ByVal x As Double, ByVal y As Double, _
        Optional ByVal includeBoundary As Boolean = True) As Boolean
    RequireRegion region
    Dim edges() As TRegionEdge, count As Long
    region.CopyEdges edges, count
    ContainsPoint = ContainsRaw(edges, count, x, y, includeBoundary)
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
        HalfPlaneCuts edges(i), lineNormalX, lineNormalY, lineBound, localCuts, localCount
        AddScalar localCuts, localCount, 0#: AddScalar localCuts, localCount, 1#
        For j = 1 To localCount
            EdgePoint edges(i), localCuts(j), x, y
            If Abs(lineNormalX * x + lineNormalY * y - lineBound) <= QUERY_TOLERANCE Then
                t = (x - pointX) * directionX + (y - pointY) * directionY
                AddScalar cuts, cutCount, t
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

' Запрашивает пары границ материала на равномерных прямых внутри заданной
' полосы. Выбирается только интервал, содержащий опорную плоскость: переход
' через opening к дальней стенке запрещен. Возвращает [найдено, Qx,Qy,Px,Py];
' число проб задает потребитель. Эти точки не заменяют точные контуры области.
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

' Возвращает первый выход из локальной бетонной стенки по направлению луча.
' Дальняя стенка за пустотой не увеличивает защитный слой; -1 означает, что
' исходная точка не принадлежит бетону. Этот API заменяет прежний ray clipping.
Public Function ConcreteCoverFromPointAlongDirection(ByVal pointX As Double, ByVal pointY As Double, _
        ByVal directionX As Double, ByVal directionY As Double) As Double
    EnsureDomain
    Dim intervals As Variant
    intervals = LineIntervals(mDomain, pointX, pointY, directionX, directionY, 0#)
    ConcreteCoverFromPointAlongDirection = -1#
    If IsEmpty(intervals) Then Exit Function
    If intervals(1, 1) <= QUERY_TOLERANCE Then ConcreteCoverFromPointAlongDirection = intervals(1, 2)
End Function

' Подготавливает и валидирует локальный domain. Изменение сетки сбрасывает
' снимок. При совместном использовании query ключ остается identity/Revision
' модели, а не LC; при ошибке незавершенная область не публикуется в кэш.
Private Sub EnsureDomain()
    If mSection Is Nothing Then GeometryError "Сначала привяжите геометрический запрос к модели сечения."
    If mSection.Revision <> mRevision Then Set mDomain = Nothing
    If Not mDomain Is Nothing Then Exit Sub
    mRevision = mSection.Revision
    Dim outer() As TRegionEdge, outerCount As Long, holes() As TRegionEdge, holeCount As Long
    ReadMaterialContours mSection.Contours, outer, outerCount, holes, holeCount
    Dim mesh() As TRegionEdge, meshCount As Long, meshRegion As CConcreteRegion
    If outerCount = 0 Or holeCount = 0 Then
        MeshBoundary mSection, mesh, meshCount
        Set meshRegion = MakeRegion(mesh, meshCount, "ElementBoundaryLoops", mRevision)
    End If
    Dim raw() As TRegionEdge, count As Long, i As Long, loopArea As Double, loopID As Long
    If outerCount > 0 Then
        OrientLoops outer, outerCount, False
        For i = 1 To outerCount
            AppendEdge raw, count, outer(i)
        Next i
        If holeCount = 0 Then
            ' Неизвестные opening не исчезают при наличии точного outer:
            ' внутренние петли фактической дискретной сетки остаются пустотами.
            For loopID = 1 To meshRegion.LoopCount
                loopArea = SignedLoopArea(mesh, meshCount, loopID)
                If loopArea < -QUERY_TOLERANCE Then
                    For i = 1 To meshCount
                        If mesh(i).LoopID = loopID Then AppendEdge raw, count, mesh(i)
                    Next i
                End If
            Next loopID
        End If
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

' Читает исходные материальные границы из геометрии модели. Расчетные области
' трещин находятся в другом слое и не могут стать новым outer/opening.
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

' Восстанавливает границу по реальным оболочкам элементов. Совпадающие
' внутренние ребра удаляются сразу hash-индексом; частичные общие ребра и
' T-соединения затем разбиваются и исключаются общим topology алгоритмом.
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
        If EdgeLength(edges(i)) > QUERY_TOLERANCE Then AppendEdge remaining, remainingCount, edges(i)
    Next i
    NormalizeBoundary remaining, remainingCount, boundary, count
    Dim ordered() As TRegionEdge, loopCount As Long
    StitchLoops boundary, count, ordered, loopCount, False
    CopyArray ordered, count, boundary
End Sub

' Разбивает пересекающиеся/совпадающие ребра и оставляет только границы
' между материалом и пустотой. Внутренние ребра, дубли отверстий и пересечения
' разных компонент не получают независимую площадь и не попадают в экспорт.
Private Sub NormalizeBoundary(ByRef raw() As TRegionEdge, ByVal rawCount As Long, _
        ByRef boundary() As TRegionEdge, ByRef boundaryCount As Long)
    Erase boundary: boundaryCount = 0
    Dim cuts() As Double, cutCount As Long, i As Long, j As Long, k As Long
    Dim part As TRegionEdge, x As Double, y As Double, tx As Double, ty As Double, length As Double
    Dim leftInside As Boolean, rightInside As Boolean, seen As Object, key As String
    Set seen = CreateObject("Scripting.Dictionary")
    For i = 1 To rawCount
        If EdgeLength(raw(i)) <= QUERY_TOLERANCE Then GoTo NextEdge
        Erase cuts: cutCount = 0
        AddScalar cuts, cutCount, 0#: AddScalar cuts, cutCount, 1#
        For j = 1 To rawCount
            If i <> j Then EdgeIntersectionCuts raw(i), raw(j), cuts, cutCount
        Next j
        SortScalars cuts, cutCount
        For k = 1 To cutCount - 1
            SubEdge raw(i), cuts(k), cuts(k + 1), part
            If EdgeLength(part) > QUERY_TOLERANCE Then
                EdgePoint part, 0.5, x, y
                EdgeTangent part, 0.5, tx, ty
                length = Sqr(tx * tx + ty * ty)
                leftInside = ContainsRaw(raw, rawCount, x - ty / length * QUERY_TOLERANCE * 8#, y + tx / length * QUERY_TOLERANCE * 8#, False)
                rightInside = ContainsRaw(raw, rawCount, x + ty / length * QUERY_TOLERANCE * 8#, y - tx / length * QUERY_TOLERANCE * 8#, False)
                If leftInside <> rightInside Then
                    If rightInside Then ReverseEdge part
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

' Упаковывает только замкнутую конечную границу. Площадь считается по формуле
' Грина с аналитическим вкладом дуг и signed вкладом внутренних контуров отверстий.
Private Function MakeRegion(ByRef edges() As TRegionEdge, ByVal count As Long, _
        ByVal source As String, ByVal revision As Long) As CConcreteRegion
    Dim ordered() As TRegionEdge, loopCount As Long, area As Double, i As Long
    Dim loopAreas() As Double, loopOpening() As Boolean
    StitchLoops edges, count, ordered, loopCount, False
    If loopCount > 0 Then
        ReDim loopAreas(1 To loopCount)
        ReDim loopOpening(1 To loopCount)
    End If
    For i = 1 To count
        Dim contribution As Double
        contribution = EdgeSignedArea(ordered(i))
        area = area + contribution
        loopAreas(ordered(i).LoopID) = loopAreas(ordered(i).LoopID) + contribution
    Next i
    For i = 1 To loopCount: loopOpening(i) = (loopAreas(i) < 0#): Next i
    If area < -QUERY_TOLERANCE Then GeometryError "Граница области имеет отрицательную площадь; проверьте outer/opening."
    Dim result As CConcreteRegion
    Set result = New CConcreteRegion
    result.SetGeometry ordered, count, loopCount, GeomMax(area, 0#), source, revision, loopOpening
    Set MakeRegion = result
End Function

' Сшивает сегменты по их концам. Разрыв явно является ошибкой геометрии, а
' касающиеся компоненты обходятся отдельно с сохранением материала слева.
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
                If SamePoint(current.X2, current.Y2, first.X1, first.Y1) Then Exit Do
                candidate = 0: bestTurn = -1E+100: reverseCandidate = False
                For j = 1 To count
                    If Not used(j) Then
                        LSet testEdge = edges(j)
                        If SamePoint(current.X2, current.Y2, testEdge.X1, testEdge.Y1) Then
                            turn = EdgeTurn(current, testEdge)
                            If turn > bestTurn Then candidate = j: bestTurn = turn: reverseCandidate = False
                        ElseIf allowReverse And SamePoint(current.X2, current.Y2, testEdge.X2, testEdge.Y2) Then
                            ReverseEdge testEdge
                            turn = EdgeTurn(current, testEdge)
                            If turn > bestTurn Then candidate = j: bestTurn = turn: reverseCandidate = True
                        End If
                    End If
                Next j
                If candidate = 0 Then GeometryError "Контур не замкнут возле точки X=" & CStr(current.X2) & "; Y=" & CStr(current.Y2) & "."
                LSet current = edges(candidate)
                If reverseCandidate Then ReverseEdge current
                used(candidate) = True
                If position > count Then GeometryError "Нарушена топология замкнутых контуров."
            Loop
        End If
    Next i
End Sub

' Нормализует направление явно заданных outer/opening loops. В отличие от
' mesh, семантическая роль известна и не выводится из пользовательских знаков.
Private Sub OrientLoops(ByRef edges() As TRegionEdge, ByVal count As Long, ByVal opening As Boolean)
    Dim ordered() As TRegionEdge, loops As Long, loopID As Long, i As Long, area As Double
    StitchLoops edges, count, ordered, loops, True
    For loopID = 1 To loops
        area = SignedLoopArea(ordered, count, loopID)
        If Abs(area) <= QUERY_TOLERANCE Then GeometryError "Заданный контур имеет нулевую площадь."
        If (area > 0# And opening) Or (area < 0# And Not opening) Then
            For i = 1 To count
                If ordered(i).LoopID = loopID Then ReverseEdge ordered(i)
            Next i
        End If
    Next loopID
    CopyArray ordered, count, edges
End Sub

' Возвращает signed площадь одного замкнутого контура, не теряя аналитический вклад дуг.
Private Function SignedLoopArea(ByRef edges() As TRegionEdge, ByVal count As Long, ByVal loopID As Long) As Double
    Dim i As Long
    For i = 1 To count
        If edges(i).LoopID = loopID Then SignedLoopArea = SignedLoopArea + EdgeSignedArea(edges(i))
    Next i
End Function

' Интеграл (x*dy-y*dx)/2 для отрезка или дуги окружности. Радиусы/углы не
' переводятся в пользовательские единицы, поэтому результат всегда мм2.
Private Function EdgeSignedArea(ByRef edge As TRegionEdge) As Double
    If Abs(edge.Sweep) <= PARAM_TOLERANCE Then
        EdgeSignedArea = 0.5 * (edge.X1 * edge.Y2 - edge.X2 * edge.Y1)
    Else
        Dim radiusSquared As Double
        radiusSquared = (edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2
        EdgeSignedArea = 0.5 * (edge.CenterX * (edge.Y2 - edge.Y1) - _
            edge.CenterY * (edge.X2 - edge.X1) + radiusSquared * edge.Sweep)
    End If
End Function

' Проверяет аналитическую принадлежность по winding number. Полупустые
' технические массивы допустимы при нормализации: внутренние ребра сокращаются.
Private Function ContainsRaw(ByRef edges() As TRegionEdge, ByVal count As Long, _
        ByVal x As Double, ByVal y As Double, ByVal includeBoundary As Boolean) As Boolean
    Dim i As Long, u As Double, winding As Long, crossX As Double, dy As Double, f1 As Double, f2 As Double
    Dim cuts() As Double, cutCount As Long, j As Long, px As Double, py As Double, tx As Double, ty As Double
    Dim countedStart As Boolean, countedEnd As Boolean
    For i = 1 To count
        If includeBoundary Then
            If PointParameter(edges(i), x, y, u) Then ContainsRaw = True: Exit Function
        End If
        If Abs(edges(i).Sweep) <= PARAM_TOLERANCE Then
            ' Стык с дугой должен считаться одной вершиной: координата конца
            ' после sin/cos может отличаться от уровня луча на машинный шум.
            ' Используем тот же линейный допуск, что аналитический arc query.
            f1 = edges(i).Y1 - y: f2 = edges(i).Y2 - y
            If Abs(f1) <= QUERY_TOLERANCE Then f1 = 0#
            If Abs(f2) <= QUERY_TOLERANCE Then f2 = 0#
            If (f1 <= 0# And f2 > 0#) Or (f2 <= 0# And f1 > 0#) Then
                crossX = edges(i).X1 - f1 * (edges(i).X2 - edges(i).X1) / (f2 - f1)
                If crossX > x Then winding = winding + Sgn(edges(i).Y2 - edges(i).Y1)
            End If
        Else
            Erase cuts: cutCount = 0
            HalfPlaneCuts edges(i), 0#, 1#, y, cuts, cutCount
            ' Угловой параметр может выйти за [0,1] при луче в пределах
            ' линейного допуска от конца дуги. Сохраняем саму общую вершину.
            If Abs(edges(i).Y1 - y) <= QUERY_TOLERANCE Then AddScalar cuts, cutCount, 0#
            If Abs(edges(i).Y2 - y) <= QUERY_TOLERANCE Then AddScalar cuts, cutCount, 1#
            countedStart = False: countedEnd = False
            For j = 1 To cutCount
                u = cuts(j)
                If u >= 0# And u <= 1# Then
                    ' Близкий к общей вершине arc-параметр согласуется с
                    ' прямым ребром по тому же линейному допуску, не по углу.
                    If u * EdgeLength(edges(i)) <= QUERY_TOLERANCE Then u = 0#
                    If (1# - u) * EdgeLength(edges(i)) <= QUERY_TOLERANCE Then u = 1#
                    If u = 0# Then
                        If countedStart Then GoTo NextCrossing
                        countedStart = True
                    ElseIf u = 1# Then
                        If countedEnd Then GoTo NextCrossing
                        countedEnd = True
                    End If
                    EdgePoint edges(i), u, px, py
                    EdgeTangent edges(i), u, tx, ty
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

' Добавляет параметры пересечения сегмента с бесконечной линией n*P=bound.
' Касание дает одну точку; совпадающий отрезок добавляет оба конца.
Private Sub HalfPlaneCuts(ByRef edge As TRegionEdge, ByVal nx As Double, ByVal ny As Double, _
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
        radius = Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2)
        norm = Sqr(nx * nx + ny * ny)
        distance = (bound - nx * edge.CenterX - ny * edge.CenterY) / norm
        If Abs(distance) > radius + QUERY_TOLERANCE Then Exit Sub
        ' При аналитическом касании вычитание перенесенных координат может
        ' дать две ложные близкие точки. В пределах линейного допуска это
        ' одна точка касания, одинаковая для clipping и material intervals.
        If Abs(Abs(distance) - radius) <= QUERY_TOLERANCE Then distance = Sgn(distance) * radius
        bx = edge.CenterX + nx / norm * distance: by = edge.CenterY + ny / norm * distance
        offset = Sqr(GeomMax(radius * radius - distance * distance, 0#))
        If PointParameter(edge, bx - ny / norm * offset, by + nx / norm * offset, u) Then AddScalar cuts, count, u
        If PointParameter(edge, bx + ny / norm * offset, by - nx / norm * offset, u) Then AddScalar cuts, count, u
    End If
End Sub

' Вычисляет пересечения пары аналитических сегментов, включая совпадающие
' линии/дуги. Параметры добавляются только на первом сегменте для его разбиения.
Private Sub EdgeIntersectionCuts(ByRef a As TRegionEdge, ByRef b As TRegionEdge, _
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
        along = (r1 * r1 - r2 * r2 + distance * distance) / (2# * distance)
        height = Sqr(GeomMax(r1 * r1 - along * along, 0#))
        bx = a.CenterX + along * dx / distance: by = a.CenterY + along * dy / distance
        x = bx - height * dy / distance: y = by + height * dx / distance
        If PointParameter(a, x, y, u) And PointParameter(b, x, y, v) Then AddScalar cuts, count, u
        x = bx + height * dy / distance: y = by - height * dx / distance
        If PointParameter(a, x, y, u) And PointParameter(b, x, y, v) Then AddScalar cuts, count, u
    End If
End Sub

' Проверяет принадлежность точки аналитическому сегменту и возвращает u.
' У дуг angle-wrap учитывается по signed sweep, без зависимости от глобальных осей.
Private Function PointParameter(ByRef edge As TRegionEdge, ByVal x As Double, ByVal y As Double, ByRef u As Double) As Boolean
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
        delta = pointAngle - startAngle
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
Private Sub EdgePoint(ByRef edge As TRegionEdge, ByVal u As Double, ByRef x As Double, ByRef y As Double)
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
Private Sub EdgeTangent(ByRef edge As TRegionEdge, ByVal u As Double, ByRef tx As Double, ByRef ty As Double)
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
Private Sub SubEdge(ByRef original As TRegionEdge, ByVal u1 As Double, ByVal u2 As Double, ByRef result As TRegionEdge)
    LSet result = original
    EdgePoint original, u1, result.X1, result.Y1
    EdgePoint original, u2, result.X2, result.Y2
    result.Sweep = original.Sweep * (u2 - u1): result.LoopID = 0
End Sub

' Обращает сегмент вместе со знаком дуги, сохраняя ту же физическую границу.
Private Sub ReverseEdge(ByRef edge As TRegionEdge)
    Dim value As Double
    value = edge.X1: edge.X1 = edge.X2: edge.X2 = value
    value = edge.Y1: edge.Y1 = edge.Y2: edge.Y2 = value
    edge.Sweep = -edge.Sweep
End Sub

' Возвращает угол поворота касательной при обходе узла topology.
Private Function EdgeTurn(ByRef previous As TRegionEdge, ByRef following As TRegionEdge) As Double
    Dim ax As Double, ay As Double, bx As Double, by As Double
    EdgeTangent previous, 1#, ax, ay: EdgeTangent following, 0#, bx, by
    EdgeTurn = AngleOf(ax * bx + ay * by, ax * by - ay * bx)
End Function

' Возвращает atan2 в диапазоне -pi..pi без обращения к Excel.
Private Function AngleOf(ByVal x As Double, ByVal y As Double) As Double
    If Abs(x) <= GEOM_TOLERANCE Then
        AngleOf = Sgn(y) * GEOM_PI / 2#
    Else
        AngleOf = Atn(y / x)
        If x < 0# Then
            If y >= 0# Then AngleOf = AngleOf + GEOM_PI Else AngleOf = AngleOf - GEOM_PI
        End If
    End If
End Function

' Проверяет соответствие концов дуги центру и signed sweep; неверный явный
' контур не подменяется fallback по сетке.
Private Sub ValidateArc(ByRef edge As TRegionEdge)
    Dim x As Double, y As Double
    If EdgeLength(edge) <= QUERY_TOLERANCE Then GeometryError "Заданная дуга имеет нулевую длину."
    EdgePoint edge, 1#, x, y
    If Not SamePoint(x, y, edge.X2, edge.Y2) Then GeometryError "Конец дуги не соответствует ее центру и углу."
End Sub

' Возвращает физическую длину сегмента для исключения вырожденных ребер.
Private Function EdgeLength(ByRef edge As TRegionEdge) As Double
    If Abs(edge.Sweep) <= PARAM_TOLERANCE Then
        EdgeLength = Sqr((edge.X2 - edge.X1) ^ 2 + (edge.Y2 - edge.Y1) ^ 2)
    Else
        EdgeLength = Abs(edge.Sweep) * Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2)
    End If
End Function

' Добавляет невырожденный отрезок в рабочий массив границы.
Private Sub AddLine(ByRef edges() As TRegionEdge, ByRef count As Long, ByVal x1 As Double, _
        ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
    Dim edge As TRegionEdge
    edge.X1 = x1: edge.Y1 = y1: edge.X2 = x2: edge.Y2 = y2
    If EdgeLength(edge) > QUERY_TOLERANCE Then AppendEdge edges, count, edge
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

' Добавляет сегмент с геометрическим ростом capacity, чтобы горячий обход
' сетки не выполнял ReDim Preserve на каждом бетонном ребре.
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

' Добавляет скалярный параметр без дублей; размерный допуск выбирается
' вызывающим пересечением, а здесь сравнивается относительная точность Double.
Private Sub AddScalar(ByRef values() As Double, ByRef count As Long, ByVal value As Double)
    Dim i As Long
    For i = 1 To count
        If Abs(values(i) - value) <= PARAM_TOLERANCE Then Exit Sub
    Next i
    If count = 0 Then ReDim values(1 To 8)
    If count = UBound(values) Then ReDim Preserve values(1 To UBound(values) * 2)
    count = count + 1: values(count) = value
End Sub

' Упорядочивает небольшие наборы параметров intersections без объектов на точку.
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

' Централизует допуск совпадения концов, вершин и принадлежности границе.
Private Function SamePoint(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Boolean
    SamePoint = (Abs(x1 - x2) <= QUERY_TOLERANCE And Abs(y1 - y2) <= QUERY_TOLERANCE)
End Function

' Hash-ключ только ускоряет удаление полностью совпадающих внутренних ребер.
' Геометрическая принадлежность и частичные совпадения затем проверяются точно.
Private Function PointKey(ByVal x As Double, ByVal y As Double) As String
    PointKey = Format$(x / QUERY_TOLERANCE, "0") & ":" & Format$(y / QUERY_TOLERANCE, "0")
End Function

' Отсутствующий объект не равен корректной пустой области.
Private Sub RequireRegion(ByVal region As CConcreteRegion)
    If region Is Nothing Then GeometryError "Готовая геометрическая область не передана."
End Sub

' Выдает ошибку геометрического контракта. Engineering-слой преобразует ее
' через общую ResultMeta/status policy; geometry не назначает display-статусы.
Private Sub GeometryError(ByVal description As String)
    Err.Raise vbObjectError + 5401, "CSectionGeometryQuery", description
End Sub
