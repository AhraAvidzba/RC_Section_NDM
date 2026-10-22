VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGeometryVariableOffset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Замкнутая проверочная граница с отдельным смещением каждого участка
' ==========================================================================
' Сохраняет точные прямые и концентрические дуги. Продолжения соседних прямых
' пересекает; разрыв при разных смещениях дуги и соседа соединяет отрезком.
' Исчезнувшее скругление удаляет; соседние прямые продолжает до пересечения.
' Такой стык может быть негладким: координаты стержней он не исправляет.
' Исходные контуры, прежние Offset и OffsetRoundedPolygon не изменяет.

Private Const TOL As Double = GEOM_BOUNDARY_TOLERANCE ' Геометрический допуск стыков и вырождения, мм.

' Строит один замкнутый контур и проверяет его до использования при фильтрации.
' distances задаёт расстояние влево по обходу каждого исходного сегмента, мм.
Friend Function Build(ByRef original() As TRegionEdge, ByRef distances() As Double, ByVal count As Long, _
        ByVal sourceReference As String, Optional ByVal segmentSources As Variant) As CConcreteRegion
    Dim shifted() As TRegionEdge, result() As TRegionEdge, sourceIndex() As Long, kept As Long, radius As Double
    Dim i As Long, following As Long, resultCount As Long
    Dim allReferences As String: allReferences = sourceReference
    Dim edge As TRegionEdge, region As CConcreteRegion, query As CSectionGeometryQuery, before As Double, after As Double
    If count < 1 Then Fail "Проверочный контур не содержит участков.", sourceReference
    ReDim shifted(1 To count): ReDim result(1 To 2 * count): ReDim sourceIndex(1 To count)
    For i = 1 To count
        If IsArray(segmentSources) Then sourceReference = OffsetSourceReferences(segmentSources, i, i, i)
        before = before + EdgeArea(original(i))
        If original(i).Sweep <> 0# Then
            radius = Sqr((original(i).X1 - original(i).CenterX) ^ 2 + (original(i).Y1 - original(i).CenterY) ^ 2)
            If radius - Sgn(original(i).Sweep) * distances(i) <= TOL Then GoTo NextSource
        End If
        kept = kept + 1: sourceIndex(kept) = i ' После удаления дуг сохраняем номера исходных участков для направлений и адресов ошибок.
        ShiftEdge original(i), distances(i), shifted(kept), sourceReference
NextSource:
    Next i
    If kept = 0 Then
        radius = Sqr((original(1).X1 - original(1).CenterX) ^ 2 + (original(1).Y1 - original(1).CenterY) ^ 2) ' Здесь исчезли все дуги; параметры первой поясняют, почему контур построить нельзя.
        Fail "При заданных отступах проверочный контур исчез целиком. " & _
            "Исходный радиус первой дуги: " & CStr(radius) & " мм; полный нормальный отступ внутрь этой дуги: " & _
            CStr(Sgn(original(1).Sweep) * distances(1)) & " мм. Уменьшите отступы рядов.", sourceReference
    End If
    For i = 1 To kept
        following = i Mod kept + 1 ' Следующий сохранившийся участок; последний соединяется с первым.
        If IsArray(segmentSources) Then sourceReference = OffsetSourceReferences(segmentSources, sourceIndex(i), sourceIndex(following), sourceIndex(i))
        If shifted(i).Sweep = 0# And shifted(following).Sweep = 0# Then JoinLines shifted(i), shifted(following)
    Next i
    For i = 1 To kept
        following = i Mod kept + 1 ' Следующий сохранившийся участок; последний соединяется с первым.
        If IsArray(segmentSources) Then sourceReference = OffsetSourceReferences(segmentSources, _
            sourceIndex((i + kept - 2) Mod kept + 1), sourceIndex(i), sourceIndex(following))
        If shifted(i).Sweep = 0# Then
            If (shifted(i).X2 - shifted(i).X1) * (original(sourceIndex(i)).X2 - original(sourceIndex(i)).X1) + _
                    (shifted(i).Y2 - shifted(i).Y1) * (original(sourceIndex(i)).Y2 - original(sourceIndex(i)).Y1) <= TOL * TOL Then _
                Fail "При заданных отступах прямой участок проверочного контура исчез или изменил направление. Уменьшите отступы рядов.", sourceReference
        End If
        resultCount = resultCount + 1: result(resultCount) = shifted(i)
        If Not modGeometryPrimitives.SamePoint(shifted(i).X2, shifted(i).Y2, shifted(following).X1, shifted(following).Y1) Then
            edge.X1 = shifted(i).X2: edge.Y1 = shifted(i).Y2
            edge.X2 = shifted(following).X1: edge.Y2 = shifted(following).Y1
            edge.CenterX = 0#: edge.CenterY = 0#: edge.Sweep = 0#: edge.LoopID = 0
            resultCount = resultCount + 1: result(resultCount) = edge
        End If
    Next i
    sourceReference = allReferences
    For i = 1 To resultCount: after = after + EdgeArea(result(i)): Next i
    If Abs(after) <= TOL Or before * after <= 0# Then _
        Fail "Проверочный контур выродился или изменил направление обхода. Уменьшите отступы дополнительных рядов.", sourceReference
    Set query = New CSectionGeometryQuery
    On Error GoTo InvalidContour
    Set region = query.BoundaryRegion(result, resultCount, "DependentRowsVariableOffset")
    query.ValidatePreparedBoundary region
    If region.LoopCount <> 1 Then Fail "Отступы образовали несколько раздельных проверочных контуров.", sourceReference
    Set Build = region
    Exit Function
InvalidContour:
    Dim reason As String
    reason = Err.Description
    Fail "Не удалось построить единый проверочный контур дополнительных рядов: " & reason & _
        " Уменьшите отступы или измените армирование соседних участков.", sourceReference
End Function

' Оставляет адреса проблемного участка и соседей, определяющих точки стыка.
' Каждый адрес выводится один раз, даже если связан с несколькими участками.
Private Function OffsetSourceReferences(ByVal sources As Variant, ByVal first As Long, ByVal middle As Long, ByVal last As Long) As String
    Dim unique As Object, index As Variant, address As Variant, token As String
    Set unique = CreateObject("Scripting.Dictionary")
    For Each index In Array(first, middle, last)
        For Each address In Split(CStr(sources(CLng(index))), ";")
            token = Trim$(CStr(address))
            If Len(token) > 0 Then
                If Not unique.Exists(token) Then unique.Add token, True
            End If
        Next address
    Next index
    If unique.Count > 0 Then OffsetSourceReferences = Join(unique.Keys, "; ")
End Function

' Переносит прямую или дугу точно, без разбиения дуги на короткие отрезки.
Friend Sub ShiftEdge(ByRef original As TRegionEdge, ByVal distance As Double, ByRef result As TRegionEdge, _
        ByVal sourceReference As String)
    Dim length As Double, radius As Double, radialScale As Double, tx As Double, ty As Double
    result = original
    If original.Sweep = 0# Then
        length = modGeometryPrimitives.EdgeLength(original)
        If length <= TOL Then Fail "В исходном проверочном контуре есть участок нулевой длины.", sourceReference
        tx = (original.X2 - original.X1) / length: ty = (original.Y2 - original.Y1) / length
        result.X1 = original.X1 - ty * distance: result.X2 = original.X2 - ty * distance
        result.Y1 = original.Y1 + tx * distance: result.Y2 = original.Y2 + tx * distance
    Else
        radius = Sqr((original.X1 - original.CenterX) ^ 2 + (original.Y1 - original.CenterY) ^ 2)
        radialScale = (radius - Sgn(original.Sweep) * distance) / radius ' Левая нормаль уменьшает радиус дуги против часовой стрелки и увеличивает радиус обратной дуги.
        If radialScale <= TOL / radius Then Fail "При заданном отступе радиус проверочной дуги становится нулевым или отрицательным. " & _
            "Исходный радиус: " & CStr(radius) & " мм; полный отступ внутрь дуги: " & CStr(Sgn(original.Sweep) * distance) & " мм. " & _
            "Уменьшите отступ дополнительного ряда или увеличьте радиус скругления.", sourceReference
        result.X1 = original.CenterX + (original.X1 - original.CenterX) * radialScale
        result.Y1 = original.CenterY + (original.Y1 - original.CenterY) * radialScale
        result.X2 = original.CenterX + (original.X2 - original.CenterX) * radialScale
        result.Y2 = original.CenterY + (original.Y2 - original.CenterY) * radialScale
    End If
End Sub

' Стыкует непараллельные прямые пересечением их продолжений.
' У параллельных участков с разными отступами разрыв соединяется в Build.
Private Sub JoinLines(ByRef first As TRegionEdge, ByRef second As TRegionEdge)
    Dim ax As Double, ay As Double, bx As Double, by As Double, cross As Double, parameter As Double
    ax = first.X2 - first.X1: ay = first.Y2 - first.Y1
    bx = second.X2 - second.X1: by = second.Y2 - second.Y1
    cross = ax * by - ay * bx ' Определитель направлений; близкий к нулю означает параллельность.
    If Abs(cross) <= GEOM_PARAMETER_TOLERANCE * Sqr((ax * ax + ay * ay) * (bx * bx + by * by)) Then Exit Sub
    parameter = ((second.X1 - first.X1) * by - (second.Y1 - first.Y1) * bx) / cross
    first.X2 = first.X1 + parameter * ax: first.Y2 = first.Y1 + parameter * ay
    second.X1 = first.X2: second.Y1 = first.Y2
End Sub

' Площадь под ориентированным отрезком или дугой; используется для контроля обхода.
Friend Function EdgeArea(ByRef edge As TRegionEdge) As Double
    Dim radiusSquared As Double
    EdgeArea = (edge.X1 * edge.Y2 - edge.X2 * edge.Y1) / 2#
    If edge.Sweep <> 0# Then
        radiusSquared = (edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2
        EdgeArea = EdgeArea + radiusSquared * (edge.Sweep - Sin(edge.Sweep)) / 2# ' К площади под хордой добавляется знаковая площадь кругового сегмента, мм2.
    End If
End Function

' Ошибка построения относится к исходным полям, а не превращается в пропуск всех стержней.
Private Sub Fail(ByVal reason As String, ByVal sourceReference As String)
    SectionRecipeError "CGeometryVariableOffset", reason, sourceReference
End Sub
