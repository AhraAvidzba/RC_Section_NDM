VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAutoCADContourReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Чтение достоверных аналитических границ AutoCAD
' ==========================================================================
' Excel/COM-адаптер извлекает реальные линии, круги и дуги из Region либо
' замкнутой полилинии. Координаты WCS трактуются как мм по общему контракту
' импорта. Замкнутость, вложенность и роли контуров определяет общий geometry-
' query. Здесь нет эквивалентных прямоугольников A/I, НДС или нормативных зон.
' Explode выполняется только над собственной копией; исходный объект
' не удаляется и не меняется, временные фрагменты очищаются при любом исходе.

Private Const PLANE_TOLERANCE As Double = 0.000001 ' Тот же линейный допуск, что у geometry-query, мм.
Private Const NORMAL_TOLERANCE As Double = 0.0000000001 ' Проверка направления нормали относительно WCS XY.

' Возвращает проверенную область фактического объекта с outer/opening loops.
' Поддерживаются Region, физически замкнутые LWPolyline/2D Polyline и Circle; кривые
' другого типа отклоняются явно, без полигональной или A/I-аппроксимации.
Public Function ReadRegion(ByVal entity As Object) As CConcreteRegion
    Dim temporary As Object, pieces As Variant, edges() As TRegionEdge, count As Long
    Dim kind As String, i As Long, query As CSectionGeometryQuery
    Dim errorNumber As Long, errorDescription As String, context As String, document As Object
    On Error GoTo Failed
    If entity Is Nothing Then Err.Raise vbObjectError + 4410, "CAutoCADContourReader", "Объект контура не передан."
    context = "Слой '" & CStr(entity.Layer) & "', объект " & CStr(entity.Handle) & ": "
    kind = CStr(entity.ObjectName)
    Select Case kind
        Case "AcDbRegion", "AcDbPolyline", "AcDb2dPolyline"
            ValidatePlaneNormal entity
            ' AutoCAD разрешает Copy/Explode на locked-слое, но запрещает Delete.
            ' Проверяем до первой копии, чтобы не оставить временные объекты.
            Set document = entity.Document
            If Not document Is Nothing Then
                If CBool(document.Layers.Item(CStr(entity.Layer)).Lock) Then _
                    Err.Raise vbObjectError + 4410, "CAutoCADContourReader", _
                        "Слой контура заблокирован. Разблокируйте этот слой в AutoCAD и повторите импорт."
            End If
            If kind <> "AcDbRegion" Then
                If Abs(CDbl(entity.Elevation)) > PLANE_TOLERANCE Then Err.Raise vbObjectError + 4410, "CAutoCADContourReader", "Контур должен лежать в расчетной плоскости WCS XY (Z = 0)."
            End If
            Set temporary = entity.Copy
            pieces = temporary.Explode
            If Not IsArray(pieces) Then Err.Raise vbObjectError + 4410, "CAutoCADContourReader", "AutoCAD не вернул граничные кривые объекта."
            For i = LBound(pieces) To UBound(pieces)
                ' Флаг Closed не заменяет проверку фактических концов. Повтор
                ' первой вершины допустим; его пустое прямое ребро не является границей.
                AppendCurve pieces(i), edges, count, (kind <> "AcDbRegion")
            Next i
        Case "AcDbCircle"
            AppendCurve entity, edges, count
        Case "AcDbLine", "AcDbArc"
            Err.Raise vbObjectError + 4410, "CAutoCADContourReader", _
                "На контурном слое найден отдельный отрезок или дуга, а не полилиния. " & _
                "Объедините границы командой PEDIT (ПОЛРЕД) в замкнутую полилинию или создайте Region. " & _
                "Если этот объект не относится к контуру, перенесите его на другой слой."
        Case Else
            Err.Raise vbObjectError + 4410, "CAutoCADContourReader", _
                "Тип " & kind & " не поддерживается для контура. Используйте Region, замкнутую 2D-полилинию или окружность с границами из прямых и круговых дуг."
    End Select
    Set query = New CSectionGeometryQuery
    Set ReadRegion = query.BoundaryRegion(edges, count, "AutoCAD:" & kind)
    CleanupTemporary temporary, pieces
    Exit Function
Failed:
    errorNumber = Err.Number: errorDescription = Err.Description
    CleanupTemporary temporary, pieces
    Err.Raise errorNumber, "CAutoCADContourReader", context & errorDescription
End Function

' Проверяет пустой самостоятельный отрезок на контурном слое. Импортер
' может пропустить его с предупреждением: такой объект не задает границу.
' Ненулевые отрезки и дуги здесь не объединяются и не скрываются.
Public Function IsEmptyStraightSegment(ByVal entity As Object) As Boolean
    Dim point As Variant, endpoint As Variant, context As String, dx As Double, dy As Double
    On Error GoTo Failed
    If entity Is Nothing Then Err.Raise vbObjectError + 4410, "CAutoCADContourReader", "Объект контура не передан."
    context = "Слой '" & CStr(entity.Layer) & "', объект " & CStr(entity.Handle) & ": "
    If StrComp(CStr(entity.ObjectName), "AcDbLine", vbTextCompare) <> 0 Then Exit Function
    point = entity.StartPoint: endpoint = entity.EndPoint
    ValidatePoint point: ValidatePoint endpoint
    dx = CDbl(point(LBound(point))) - CDbl(endpoint(LBound(endpoint)))
    dy = CDbl(point(LBound(point) + 1)) - CDbl(endpoint(LBound(endpoint) + 1))
    IsEmptyStraightSegment = (Abs(dx) <= PLANE_TOLERANCE And Abs(dy) <= PLANE_TOLERANCE)
    Exit Function
Failed:
    Err.Raise Err.Number, "CAutoCADContourReader", context & Err.Description
End Function

' Читает только аналитически поддержанные граничные кривые. EndPoint и
' StartPoint фрагментов Explode заданы в WCS, поэтому OCS-координаты
' полилинии и bulge не пересчитываются вручную и не теряют исходный поворот.
' Только в полилинии пропускаются пустые прямые сегменты повторных вершин;
' вырожденные дуги и поврежденные Region по-прежнему отклоняются.
Private Sub AppendCurve(ByVal curve As Object, ByRef edges() As TRegionEdge, ByRef count As Long, _
        Optional ByVal ignoreRepeatedPolylineVertex As Boolean = False)
    Dim kind As String, edge As TRegionEdge, point As Variant, endpoint As Variant
    Dim normal As Variant, radius As Double, sweep As Double, angle As Double, i As Long
    kind = CStr(curve.ObjectName)
    Select Case kind
        Case "AcDbLine"
            point = curve.StartPoint: endpoint = curve.EndPoint
            ValidatePoint point: ValidatePoint endpoint
            edge.X1 = CDbl(point(LBound(point))): edge.Y1 = CDbl(point(LBound(point) + 1))
            edge.X2 = CDbl(endpoint(LBound(endpoint))): edge.Y2 = CDbl(endpoint(LBound(endpoint) + 1))
            If ignoreRepeatedPolylineVertex Then
                If Abs(edge.X2 - edge.X1) <= PLANE_TOLERANCE And Abs(edge.Y2 - edge.Y1) <= PLANE_TOLERANCE Then Exit Sub
            End If
            AppendEdge edges, count, edge
        Case "AcDbArc"
            ValidatePlaneNormal curve
            point = curve.Center: ValidatePoint point
            edge.CenterX = CDbl(point(LBound(point))): edge.CenterY = CDbl(point(LBound(point) + 1))
            point = curve.StartPoint: endpoint = curve.EndPoint
            ValidatePoint point: ValidatePoint endpoint
            edge.X1 = CDbl(point(LBound(point))): edge.Y1 = CDbl(point(LBound(point) + 1))
            edge.X2 = CDbl(endpoint(LBound(endpoint))): edge.Y2 = CDbl(endpoint(LBound(endpoint) + 1))
            normal = curve.Normal
            sweep = CDbl(curve.TotalAngle)
            If sweep <= 0# Or sweep > 2# * GEOM_PI Then Err.Raise vbObjectError + 4410, "CAutoCADContourReader", "Круговая дуга имеет недопустимый угол."
            edge.Sweep = sweep * Sgn(CDbl(normal(LBound(normal) + 2)))
            AppendEdge edges, count, edge
        Case "AcDbCircle"
            ValidatePlaneNormal curve
            point = curve.Center: ValidatePoint point
            edge.CenterX = CDbl(point(LBound(point))): edge.CenterY = CDbl(point(LBound(point) + 1))
            radius = CDbl(curve.Radius)
            If radius <= PLANE_TOLERANCE Then Err.Raise vbObjectError + 4410, "CAutoCADContourReader", "Окружность контура имеет нулевой радиус."
            ' Четыре точные дуги сохраняют полную окружность без неоднозначного
            ' совпадения начала/конца одного сегмента.
            For i = 0 To 3
                angle = CDbl(i) * GEOM_PI / 2#
                edge.X1 = edge.CenterX + radius * Cos(angle): edge.Y1 = edge.CenterY + radius * Sin(angle)
                edge.X2 = edge.CenterX + radius * Cos(angle + GEOM_PI / 2#)
                edge.Y2 = edge.CenterY + radius * Sin(angle + GEOM_PI / 2#)
                edge.Sweep = GEOM_PI / 2#
                AppendEdge edges, count, edge
            Next i
        Case Else
            Err.Raise vbObjectError + 4410, "CAutoCADContourReader", _
                "Граница содержит неподдерживаемую кривую " & kind & ". Допустимы только прямые, круговые дуги и окружности; сплайны и эллипсы не аппроксимируются."
    End Select
End Sub

' Требует плоскую геометрию в расчетной WCS XY. Обратная нормаль допустима:
' направление обхода нормализует topology, а отрицательная нормаль дуги
' учитывается при чтении signed sweep.
Private Sub ValidatePlaneNormal(ByVal entity As Object)
    Dim normal As Variant
    normal = entity.Normal
    If Abs(CDbl(normal(LBound(normal)))) > NORMAL_TOLERANCE Or _
            Abs(CDbl(normal(LBound(normal) + 1))) > NORMAL_TOLERANCE Or _
            Abs(Abs(CDbl(normal(LBound(normal) + 2))) - 1#) > NORMAL_TOLERANCE Then _
        Err.Raise vbObjectError + 4410, "CAutoCADContourReader", "Плоскость контура не совпадает с расчетной WCS XY."
End Sub

' Проверяет фактические WCS-точки, а не только нормаль. Контур в параллельной
' плоскости с ненулевым Z нельзя молча спроецировать на сечение.
Private Sub ValidatePoint(ByVal point As Variant)
    If UBound(point) - LBound(point) <> 2 Then Err.Raise vbObjectError + 4410, "CAutoCADContourReader", "AutoCAD вернул некорректную трехмерную точку."
    If Abs(CDbl(point(LBound(point) + 2))) > PLANE_TOLERANCE Then _
        Err.Raise vbObjectError + 4410, "CAutoCADContourReader", "Контур должен лежать в расчетной плоскости WCS XY (Z = 0)."
End Sub

' Добавляет один аналитический сегмент в рабочий массив без изменения
' исходных COM-объектов или создания расчетного элемента на каждую точку.
Private Sub AppendEdge(ByRef edges() As TRegionEdge, ByRef count As Long, ByRef edge As TRegionEdge)
    count = count + 1
    ReDim Preserve edges(1 To count)
    LSet edges(count) = edge
End Sub

' Удаляет исключительно собственную Copy и фрагменты ее Explode, включая
' ошибочный импорт. Исходная сущность никогда не передается этому методу.
Private Sub CleanupTemporary(ByVal temporary As Object, ByVal pieces As Variant)
    On Error Resume Next
    Dim i As Long
    If IsArray(pieces) Then
        For i = LBound(pieces) To UBound(pieces): pieces(i).Delete: Next i
    End If
    If Not temporary Is Nothing Then temporary.Delete
    On Error GoTo 0
End Sub
