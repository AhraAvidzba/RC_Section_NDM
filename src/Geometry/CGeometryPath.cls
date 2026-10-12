VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGeometryPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Непрерывный аналитический путь из линий и дуг. После Seal он неизменяем.
' Длины и параметры готовятся один раз; общая математика остается в query.
Private Const PATH_TOLERANCE As Double = 0.000001 ' Допуск стыков, мм.
Private mID As String
Private mEdges() As TRegionEdge
Private mSegmentIDs() As String
Private mSourcePathIDs() As String, mSourceSegmentIDs() As String
Private mSourceStart() As Double, mSourceEnd() As Double
Private mStations() As Double
Private mCount As Long, mCapacity As Long
Private mSealed As Boolean, mClosed As Boolean
Private mQuery As CSectionGeometryQuery
Private mDiagnostic As CGeometryDiagnostic

Public Property Get LastDiagnostic() As CGeometryDiagnostic
    Set LastDiagnostic = mDiagnostic
End Property

Public Sub Initialize(ByVal pathID As String)
    If Len(mID) > 0 Then PathError "Путь уже инициализирован."
    If Len(Trim$(pathID)) = 0 Then PathError "Не задан ID пути."
    mID = pathID
    Set mQuery = New CSectionGeometryQuery
End Sub

Public Sub AddLine(ByVal segmentID As String, ByVal x1 As Double, ByVal y1 As Double, _
        ByVal x2 As Double, ByVal y2 As Double)
    Dim edge As TRegionEdge
    edge.X1 = x1: edge.Y1 = y1: edge.X2 = x2: edge.Y2 = y2
    Append segmentID, edge
End Sub

' Каноническая дуга: центр/радиус/начальный угол/signed sweep, углы в радианах.
' Полный оборот допустим; нулевая дуга и более одного оборота запрещены.
Public Sub AddArc(ByVal segmentID As String, ByVal centerX As Double, ByVal centerY As Double, _
        ByVal radius As Double, ByVal startAngle As Double, ByVal sweep As Double)
    RequireMutable
    If radius <= 0# Then PathError "Радиус дуги должен быть положительным.", "INVALID_RADIUS", segmentID, radius
    If Abs(sweep) <= 0.0000000001 Or Abs(sweep) > 2# * GEOM_PI Then PathError "Угол дуги должен быть ненулевым и не больше полного оборота.", "INVALID_SWEEP", segmentID, sweep
    Dim edge As TRegionEdge
    edge.CenterX = centerX: edge.CenterY = centerY: edge.Sweep = sweep
    edge.X1 = centerX + radius * Cos(startAngle): edge.Y1 = centerY + radius * Sin(startAngle)
    edge.X2 = centerX + radius * Cos(startAngle + sweep): edge.Y2 = centerY + radius * Sin(startAngle + sweep)
    If Abs(sweep) = 2# * GEOM_PI Then edge.X2 = edge.X1: edge.Y2 = edge.Y1
    Append segmentID, edge
End Sub

' Публикация закрывает возможность изменения; разрывы нельзя чинить
' фиктивной линией, а Closed не соединяет физически разные концы.
Public Sub Seal(Optional ByVal closed As Boolean = False)
    RequireMutable
    If mCount = 0 Then PathError "Путь не содержит сегментов."
    If closed Then
        If Not PointsCoincide(mEdges(1).X1, mEdges(1).Y1, mEdges(mCount).X2, mEdges(mCount).Y2) Then _
            PathError "Замкнутый путь имеет разные начало и конец."
    End If
    ReDim Preserve mEdges(1 To mCount)
    ReDim Preserve mSegmentIDs(1 To mCount)
    ReDim Preserve mSourcePathIDs(1 To mCount): ReDim Preserve mSourceSegmentIDs(1 To mCount)
    ReDim Preserve mSourceStart(1 To mCount): ReDim Preserve mSourceEnd(1 To mCount)
    ReDim Preserve mStations(0 To mCount)
    mCapacity = mCount: mClosed = closed: mSealed = True
End Sub

Public Property Get ID() As String: ID = mID: End Property
Public Property Get SegmentCount() As Long: SegmentCount = mCount: End Property
Public Property Get IsSealed() As Boolean: IsSealed = mSealed: End Property
Public Property Get IsClosed() As Boolean: RequireSealed: IsClosed = mClosed: End Property
Public Property Get Length() As Double: RequireSealed: Length = mStations(mCount): End Property
Public Property Get SegmentID(ByVal index As Long) As String
    CheckIndex index: SegmentID = mSegmentIDs(index)
End Property

Public Property Get SourcePathID(ByVal index As Long) As String
    CheckIndex index: SourcePathID = mSourcePathIDs(index)
End Property

Public Property Get SourceSegmentID(ByVal index As Long) As String
    CheckIndex index: SourceSegmentID = mSourceSegmentIDs(index)
End Property

Public Property Get SourceStartParameter(ByVal index As Long) As Double
    CheckIndex index: SourceStartParameter = mSourceStart(index)
End Property

Public Property Get SourceEndParameter(ByVal index As Long) As Double
    CheckIndex index: SourceEndParameter = mSourceEnd(index)
End Property

Public Function SegmentIndex(ByVal segmentID As String) As Long
    Dim i As Long
    For i = 1 To mCount
        If StrComp(mSegmentIDs(i), segmentID, vbBinaryCompare) = 0 Then SegmentIndex = i: Exit Function
    Next i
    PathError "Не найден сегмент '" & segmentID & "'.", "MISSING_SEGMENT", segmentID
End Function

' Привязка задается именем сегмента и долей его длины, а не номером в массиве.
Public Function AnchorStation(ByVal segmentID As String, Optional ByVal parameter As Double = 0#) As Double
    RequireSealed
    If parameter < 0# Or parameter > 1# Then PathError "Параметр привязки должен быть от 0 до 1."
    Dim index As Long
    index = SegmentIndex(segmentID)
    AnchorStation = mStations(index - 1) + parameter * (mStations(index) - mStations(index - 1))
End Function

Public Function WithID(ByVal pathID As String) As CGeometryPath
    RequireSealed
    Dim result As CGeometryPath, i As Long
    Set result = New CGeometryPath: result.Initialize pathID
    For i = 1 To mCount
        result.Append mSegmentIDs(i), mEdges(i), mSourcePathIDs(i), mSourceSegmentIDs(i), mSourceStart(i), mSourceEnd(i)
    Next i
    result.Seal mClosed: Set WithID = result
End Function

' На замкнутом контуре направление обхода задается явно. Обход через начало
' не добавляет хорду; части сегмента сохраняют исходные ID и интервалы.
Public Function BetweenAnchors(ByVal pathID As String, ByVal startSegmentID As String, ByVal startParameter As Double, _
        ByVal endSegmentID As String, ByVal endParameter As Double, Optional ByVal forward As Boolean = True, _
        Optional ByVal fullLoopIfSame As Boolean = False) As CGeometryPath
    RequireSealed
    Dim result As CGeometryPath, part As CGeometryPath, first As Double, last As Double
    If Not forward Then
        Set part = BetweenAnchors(pathID, endSegmentID, endParameter, startSegmentID, startParameter, True, fullLoopIfSame)
        Set BetweenAnchors = part.Reversed: Exit Function
    End If
    first = AnchorStation(startSegmentID, startParameter): last = AnchorStation(endSegmentID, endParameter)
    If mClosed Then
        If first >= mStations(mCount) - PATH_TOLERANCE Then first = 0#
        If last >= mStations(mCount) - PATH_TOLERANCE Then last = 0#
    End If
    If last > first + PATH_TOLERANCE Then
        Set part = Slice(first, last): Set BetweenAnchors = part.WithID(pathID): Exit Function
    End If
    If Not mClosed Then PathError "Привязки открытого пути не соответствуют выбранному направлению."
    If Abs(first - last) <= PATH_TOLERANCE And Not fullLoopIfSame Then PathError "Привязки совпадают; полный обход должен быть задан явно."
    Set result = New CGeometryPath: result.Initialize pathID
    AppendInterval result, first, mStations(mCount)
    AppendInterval result, 0#, last
    result.Seal (Abs(first - last) <= PATH_TOLERANCE)
    Set BetweenAnchors = result
End Function

' Параметр - пройденная длина в мм. По умолчанию на стыке выбирается
' исходящая касательная; incomingTangentAtJoint сохраняет входящее направление.
Public Sub AtDistance(ByVal distance As Double, ByRef x As Double, ByRef y As Double, _
        ByRef tangentX As Double, ByRef tangentY As Double, Optional ByVal incomingTangentAtJoint As Boolean = False)
    RequireSealed
    distance = CheckedDistance(distance)
    Dim lo As Long, hi As Long, middle As Long, parameter As Double, norm As Double
    lo = 1: hi = mCount
    Do While lo < hi
        middle = lo + (hi - lo) \ 2
        If distance < mStations(middle) Or (incomingTangentAtJoint And distance <= mStations(middle) + GEOM_TOLERANCE) Then
            hi = middle
        Else
            lo = middle + 1
        End If
    Loop
    parameter = (distance - mStations(lo - 1)) / (mStations(lo) - mStations(lo - 1))
    parameter = GeomMax(0#, GeomMin(1#, parameter))
    mQuery.EdgePoint mEdges(lo), parameter, x, y
    mQuery.EdgeTangent mEdges(lo), parameter, tangentX, tangentY
    norm = Sqr(tangentX * tangentX + tangentY * tangentY)
    tangentX = tangentX / norm: tangentY = tangentY / norm
End Sub

Public Sub ProjectionBounds(ByVal directionX As Double, ByVal directionY As Double, _
        ByRef minimum As Double, ByRef maximum As Double)
    RequireSealed
    Dim norm As Double, i As Long
    norm = Sqr(directionX * directionX + directionY * directionY)
    If norm <= GEOM_TOLERANCE Then PathError "Направление проекции имеет нулевую длину."
    directionX = directionX / norm: directionY = directionY / norm
    minimum = 1E+100: maximum = -1E+100
    For i = 1 To mCount
        mQuery.AccumulateEdgeProjection mEdges(i), directionX, directionY, minimum, maximum
    Next i
End Sub

' Ограничивает весь составной путь по длине без сброса шага на стыках.
' Полученные части сохраняют исходные смысловые ID и точные дуги.
Public Function Slice(ByVal startDistance As Double, ByVal endDistance As Double) As CGeometryPath
    RequireSealed
    startDistance = CheckedDistance(startDistance): endDistance = CheckedDistance(endDistance)
    If endDistance <= startDistance + PATH_TOLERANCE Then PathError "Рабочая длина пути после отсечения неположительна."
    Dim result As CGeometryPath, i As Long, first As Double, last As Double, segmentLength As Double, edge As TRegionEdge
    Set result = New CGeometryPath: result.Initialize mID
    For i = 1 To mCount
        first = GeomMax(startDistance, mStations(i - 1)): last = GeomMin(endDistance, mStations(i))
        If last > first + PATH_TOLERANCE Then
            segmentLength = mStations(i) - mStations(i - 1)
            mQuery.SubEdge mEdges(i), (first - mStations(i - 1)) / segmentLength, (last - mStations(i - 1)) / segmentLength, edge
            result.Append mSegmentIDs(i), edge, mSourcePathIDs(i), mSourceSegmentIDs(i), _
                OriginParameter(i, (first - mStations(i - 1)) / segmentLength), OriginParameter(i, (last - mStations(i - 1)) / segmentLength)
        End If
    Next i
    result.Seal (mClosed And startDistance = 0# And endDistance = mStations(mCount))
    Set Slice = result
End Function

Private Sub AppendInterval(ByVal result As CGeometryPath, ByVal first As Double, ByVal last As Double)
    If last <= first + PATH_TOLERANCE Then Exit Sub
    Dim part As CGeometryPath, i As Long, edge As TRegionEdge, name As String, suffix As Long
    Set part = Slice(first, last)
    For i = 1 To part.SegmentCount
        part.CopySegment i, edge: name = part.SegmentID(i): suffix = 1
        Do While result.HasSegmentID(name)
            suffix = suffix + 1: name = part.SegmentID(i) & ".Part" & CStr(suffix)
        Loop
        result.Append name, edge, part.SourcePathID(i), part.SourceSegmentID(i), part.SourceStartParameter(i), part.SourceEndParameter(i)
    Next i
End Sub

Private Function OriginParameter(ByVal index As Long, ByVal parameter As Double) As Double
    OriginParameter = mSourceStart(index) + parameter * (mSourceEnd(index) - mSourceStart(index))
End Function

Friend Function HasSegmentID(ByVal segmentID As String) As Boolean
    Dim i As Long
    For i = 1 To mCount
        If StrComp(mSegmentIDs(i), segmentID, vbBinaryCompare) = 0 Then HasSegmentID = True: Exit Function
    Next i
End Function

Public Function Reversed() As CGeometryPath
    RequireSealed
    Dim result As CGeometryPath, i As Long, edge As TRegionEdge
    Set result = New CGeometryPath: result.Initialize mID
    For i = mCount To 1 Step -1
        LSet edge = mEdges(i): mQuery.ReverseEdge edge
        result.Append mSegmentIDs(i), edge, mSourcePathIDs(i), mSourceSegmentIDs(i), mSourceEnd(i), mSourceStart(i)
    Next i
    result.Seal mClosed
    Set Reversed = result
End Function

' Изометрия: сначала отражение относительно локальной оси X, затем поворот
' вокруг (0,0) и перенос. Отражение меняет signed sweep, но не смысловые ID.
Public Function Transformed(ByVal shiftX As Double, ByVal shiftY As Double, _
        Optional ByVal rotation As Double = 0#, Optional ByVal reflectX As Boolean = False) As CGeometryPath
    RequireSealed
    Dim result As CGeometryPath, i As Long, edge As TRegionEdge, cosine As Double, sine As Double
    Set result = New CGeometryPath: result.Initialize mID
    cosine = Cos(rotation): sine = Sin(rotation)
    For i = 1 To mCount
        LSet edge = mEdges(i)
        TransformPoint edge.X1, edge.Y1, shiftX, shiftY, cosine, sine, reflectX
        TransformPoint edge.X2, edge.Y2, shiftX, shiftY, cosine, sine, reflectX
        If edge.Sweep <> 0# Then
            TransformPoint edge.CenterX, edge.CenterY, shiftX, shiftY, cosine, sine, reflectX
            If reflectX Then edge.Sweep = -edge.Sweep
        End If
        result.Append mSegmentIDs(i), edge, mSourcePathIDs(i), mSourceSegmentIDs(i), mSourceStart(i), mSourceEnd(i)
    Next i
    result.Seal mClosed
    Set Transformed = result
End Function

' ProjectFrom получает ближайший положительный hit на фактическом target.
' Отсутствие hit возвращает False, не переносит точку на бесконечную прямую.
Public Function RayHit(ByVal originX As Double, ByVal originY As Double, ByVal directionX As Double, _
        ByVal directionY As Double, ByRef pathDistance As Double, ByRef rayDistance As Double) As Boolean
    RequireSealed
    Dim i As Long, distance As Double, parameter As Double
    pathDistance = 0#: rayDistance = 0#
    If Sqr(directionX * directionX + directionY * directionY) <= GEOM_TOLERANCE Then PathError "Направление луча имеет нулевую длину."
    For i = 1 To mCount
        If mQuery.EdgeRayHit(mEdges(i), originX, originY, directionX, directionY, distance, parameter) Then
            If Not RayHit Or distance < rayDistance Then
                rayDistance = distance
                pathDistance = mStations(i - 1) + parameter * (mStations(i) - mStations(i - 1))
                RayHit = True
            End If
        End If
    Next i
End Function

Friend Sub CopySegment(ByVal index As Long, ByRef edge As TRegionEdge)
    RequireSealed
    CheckIndex index
    LSet edge = mEdges(index)
End Sub

Friend Sub Append(ByVal segmentID As String, ByRef edge As TRegionEdge, Optional ByVal sourcePathID As String = "", _
        Optional ByVal sourceSegmentID As String = "", Optional ByVal sourceStart As Double = 0#, Optional ByVal sourceEnd As Double = 1#)
    RequireMutable
    If Len(Trim$(segmentID)) = 0 Then PathError "Не задан ID сегмента."
    Dim i As Long, segmentLength As Double
    For i = 1 To mCount
        If mSegmentIDs(i) = segmentID Then PathError "Повторный ID сегмента '" & segmentID & "'."
    Next i
    segmentLength = mQuery.EdgeLength(edge)
    If segmentLength <= PATH_TOLERANCE Then PathError "Сегмент '" & segmentID & "' имеет нулевую длину."
    If mCount > 0 Then
        If Not PointsCoincide(mEdges(mCount).X2, mEdges(mCount).Y2, edge.X1, edge.Y1) Then _
            PathError "Разрыв перед сегментом '" & segmentID & "'. Несвязанные пути должны быть отдельными."
    End If
    If mCount = mCapacity Then
        mCapacity = GeomMax(4#, CDbl(mCapacity) * 2#)
        ReDim Preserve mEdges(1 To mCapacity)
        ReDim Preserve mSegmentIDs(1 To mCapacity)
        ReDim Preserve mSourcePathIDs(1 To mCapacity): ReDim Preserve mSourceSegmentIDs(1 To mCapacity)
        ReDim Preserve mSourceStart(1 To mCapacity): ReDim Preserve mSourceEnd(1 To mCapacity)
        ReDim Preserve mStations(0 To mCapacity)
    End If
    mCount = mCount + 1
    LSet mEdges(mCount) = edge: mSegmentIDs(mCount) = segmentID
    If Len(sourcePathID) = 0 Then sourcePathID = mID
    If Len(sourceSegmentID) = 0 Then sourceSegmentID = segmentID
    mSourcePathIDs(mCount) = sourcePathID: mSourceSegmentIDs(mCount) = sourceSegmentID
    mSourceStart(mCount) = sourceStart: mSourceEnd(mCount) = sourceEnd
    mStations(mCount) = mStations(mCount - 1) + segmentLength
End Sub

Private Sub TransformPoint(ByRef x As Double, ByRef y As Double, ByVal shiftX As Double, ByVal shiftY As Double, _
        ByVal cosine As Double, ByVal sine As Double, ByVal reflectX As Boolean)
    Dim newX As Double
    If reflectX Then y = -y
    newX = cosine * x - sine * y + shiftX
    y = sine * x + cosine * y + shiftY: x = newX
End Sub

Private Function PointsCoincide(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Boolean
    PointsCoincide = (Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2) <= PATH_TOLERANCE)
End Function

Private Function CheckedDistance(ByVal distance As Double) As Double
    If distance < -PATH_TOLERANCE Or distance > mStations(mCount) + PATH_TOLERANCE Then PathError "Позиция выходит за пределы длины пути."
    CheckedDistance = GeomMax(0#, GeomMin(mStations(mCount), distance))
End Function

Private Sub RequireMutable()
    If Len(mID) = 0 Then PathError "Путь не инициализирован."
    If mSealed Then PathError "Опубликованный путь нельзя изменять."
End Sub

Private Sub RequireSealed()
    If Not mSealed Then PathError "Путь еще не подготовлен и не опубликован."
End Sub

Private Sub CheckIndex(ByVal index As Long)
    If index < 1 Or index > mCount Then PathError "Номер сегмента вне допустимого диапазона."
End Sub

Private Sub PathError(ByVal description As String, Optional ByVal code As String = "PATH_INVALID", _
        Optional ByVal relatedEntityID As String = "", Optional ByVal detail1 As Double = 0#, Optional ByVal detail2 As Double = 0#)
    Set mDiagnostic = New CGeometryDiagnostic
    mDiagnostic.Initialize code, mID, relatedEntityID, description, detail1, detail2
    Err.Raise vbObjectError + 5520, "CGeometryPath", "Путь '" & mID & "': " & description
End Sub
