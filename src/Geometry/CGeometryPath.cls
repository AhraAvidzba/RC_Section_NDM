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

' ==========================================================================
' Непрерывная линия из отрезков и дуг
' ==========================================================================
' Описывает открытый или замкнутый контур для границы бетона, расстановки
' арматуры и привязки аннотаций. После Seal геометрия не изменяется.
' Длины и параметры подготавливаются один раз; формулы отдельных сегментов
' находятся в modGeometryPrimitives.

Private Const PATH_TOLERANCE As Double = 0.000001 ' Допуск стыков, мм.
Private mID As String ' Имя линии в описании сечения; используется для ссылок на неё из правил и аннотаций.
Private mEdges() As TRegionEdge ' Отрезки и дуги в порядке обхода линии; соединения между раздельными участками не добавляются.
Private mSegmentIDs() As String ' Имена отрезков и дуг этой линии для привязки правил и аннотаций.
Private mSourcePathIDs() As String, mSourceSegmentIDs() As String ' Имена исходных линий и сегментов для каждого полученного сегмента.
Private mSourceStart() As Double, mSourceEnd() As Double ' Начальный и конечный параметры исходного сегмента для переноса привязанных точек.
Private mStations() As Double ' Накопленные длины сегментов; позволяют искать точку без повторного суммирования.
Private mCollapsedIDs() As String, mCollapsedStations() As Double, mCollapsedCount As Long ' Имена и координаты привязки исчезнувших участков, число таких участков; новые отрезки вместо них не создаются.
Private mCount As Long, mCapacity As Long ' Число аналитических сегментов и ёмкость рабочего массива.
Private mSealed As Boolean, mClosed As Boolean ' Завершение построения и фактическая замкнутость линии.
Private mOffsetBoundary As CGeometryPath ' Полная замкнутая линия, из которой выбран участок; ссылка только на неизменяемую геометрию.
Private mBoundaryDirection As Double ' Направление участка относительно полной линии: 1 по её обходу, -1 против.
Private mVariableOffsetSource As CGeometryPath ' Исходный контур проверенного общего смещения; устанавливается только его построителем.
Private mVariableDistances As Variant, mVariableReferences As Variant ' Независимые снимки as по исходным сегментам и адресов; соединительные отрезки не становятся исходными гранями.
Private mDiagnostic As CGeometryDiagnostic ' Последняя ошибка построения или обращения к линии.

' Последняя машинная диагностика; Nothing при отсутствии записанного события.
Public Property Get LastDiagnostic() As CGeometryDiagnostic
    Set LastDiagnostic = mDiagnostic
End Property

' Позволяет исполнителю связать общий смещённый контур с исходными сегментами без угадывания соединительных отрезков.
Friend Sub SetVariableOffsetSource(ByVal source As CGeometryPath, ByVal distances As Variant, ByVal references As Variant)
    RequireMutable
    If Not source.IsClosed Then PathError "Исходная граница индивидуального смещения должна быть замкнутой."
    Set mVariableOffsetSource = source
    mVariableDistances = distances: mVariableReferences = references
End Sub

' Возвращает происхождение только у контура, подготовленного общим построителем индивидуальных смещений.
Friend Function VariableOffsetSource() As CGeometryPath
    RequireSealed
    Set VariableOffsetSource = mVariableOffsetSource
End Function

' Возвращает копию подписанных as исходных участков для следующего равномерного нормального переноса.
Friend Function VariableOffsetDistances() As Variant
    RequireSealed
    VariableOffsetDistances = mVariableDistances
End Function

' Возвращает копию адресов исходных as; последующее смещение сохраняет понятные причины вырождения соседних сторон.
Friend Function VariableOffsetReferences() As Variant
    RequireSealed
    VariableOffsetReferences = mVariableReferences
End Function

' Начинает описание линии с непустым ID. Допускается только один вызов для объекта.
Public Sub Initialize(ByVal pathID As String)
    If Len(mID) > 0 Then PathError "Путь уже инициализирован."
    If Len(Trim$(pathID)) = 0 Then PathError "Не задан ID пути."
    mID = pathID
End Sub

' Добавляет отрезок с заданным ID; его начало должно совпадать с концом предыдущего сегмента.
Public Sub AddLine(ByVal segmentID As String, ByVal x1 As Double, ByVal y1 As Double, _
        ByVal x2 As Double, ByVal y2 As Double)
    Dim edge As TRegionEdge
    edge.X1 = x1: edge.Y1 = y1: edge.X2 = x2: edge.Y2 = y2
    Append segmentID, edge
End Sub

' Добавляет дугу по центру, радиусу, начальному углу и углу обхода со знаком, в радианах.
' Полный оборот допустим; нулевая дуга и дуга больше одного оборота вызывают ошибку.
Public Sub AddArc(ByVal segmentID As String, ByVal centerX As Double, ByVal centerY As Double, _
        ByVal radius As Double, ByVal startAngle As Double, ByVal sweep As Double)
    RequireMutable
    If radius <= 0# Then PathError "Радиус дуги должен быть положительным.", "INVALID_RADIUS", segmentID, radius
    If Abs(sweep) <= 0.0000000001 Or Abs(sweep) > 2# * GEOM_PI Then PathError "Угол дуги должен быть ненулевым и не больше полного оборота.", "INVALID_SWEEP", segmentID, sweep
    Dim edge As TRegionEdge
    edge.CenterX = centerX: edge.CenterY = centerY: edge.Sweep = sweep
    edge.X1 = centerX + radius * Cos(startAngle): edge.Y1 = centerY + radius * Sin(startAngle)
    edge.X2 = centerX + radius * Cos(startAngle + sweep): edge.Y2 = centerY + radius * Sin(startAngle + sweep)
    If Abs(sweep) = 2# * GEOM_PI Then edge.X2 = edge.X1: edge.Y2 = edge.Y1 ' У полного оборота замыкание задаём точно, не по погрешным sin/cos.
    Append segmentID, edge
End Sub

' Завершает и проверяет линию, после чего менять её нельзя.
' Разрывы не соединяются дополнительными отрезками. Запрос Closed
' не замыкает разные концы автоматически: они уже должны совпадать.
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

' Возвращает имя линии, по которому на неё ссылаются правила и аннотации.
Public Property Get ID() As String: ID = mID: End Property
' Возвращает число отрезков и дуг линии.
Public Property Get SegmentCount() As Long: SegmentCount = mCount: End Property
' Показывает, завершено ли описание линии и запрещены ли дальнейшие изменения.
Public Property Get IsSealed() As Boolean: IsSealed = mSealed: End Property
' Замкнут ли фактический путь, без фиктивного соединения.
Public Property Get IsClosed() As Boolean: RequireSealed: IsClosed = mClosed: End Property
' Точная суммарная длина пути, мм.
Public Property Get Length() As Double: RequireSealed: Length = mStations(mCount): End Property
' Возвращает имя отрезка или дуги по порядковому номеру в линии.
Public Property Get SegmentID(ByVal index As Long) As String
    CheckIndex index: SegmentID = mSegmentIDs(index)
End Property

' Возвращает ID линии, из которой получен сегмент после обрезки, преобразования или смещения.
Public Property Get SourcePathID(ByVal index As Long) As String
    CheckIndex index: SourcePathID = mSourcePathIDs(index)
End Property

' Возвращает исходное имя сегмента, сохранённое при его разделении или переименовании.
Public Property Get SourceSegmentID(ByVal index As Long) As String
    CheckIndex index: SourceSegmentID = mSourceSegmentIDs(index)
End Property

' Начальный параметр исходного сегмента; может выйти за [0;1] при явном продлении прямой.
Public Property Get SourceStartParameter(ByVal index As Long) As Double
    CheckIndex index: SourceStartParameter = mSourceStart(index)
End Property

' Конечный параметр исходного сегмента, включая обращённый обход.
Public Property Get SourceEndParameter(ByVal index As Long) As Double
    CheckIndex index: SourceEndParameter = mSourceEnd(index)
End Property

' Находит номер сегмента по его ID. Если нужная дуга исчезла при построении,
' сообщает об ошибке MISSING_SEGMENT.
Public Function SegmentIndex(ByVal segmentID As String) As Long
    Dim i As Long
    For i = 1 To mCount
        If StrComp(mSegmentIDs(i), segmentID, vbBinaryCompare) = 0 Then SegmentIndex = i: Exit Function
    Next i
    PathError "Не найден сегмент '" & segmentID & "'.", "MISSING_SEGMENT", segmentID
End Function

' Находит привязку по имени сегмента и доле его длины. Если сегмент исчез
' при построении, обе его границы соответствуют одному сохранённому стыку.
' Неизвестное имя вызывает ошибку.
Public Function AnchorStation(ByVal segmentID As String, Optional ByVal parameter As Double = 0#) As Double
    RequireSealed
    If parameter < 0# Or parameter > 1# Then PathError "Параметр привязки должен быть от 0 до 1."
    Dim index As Long
    For index = 1 To mCollapsedCount
        If mCollapsedIDs(index) = segmentID Then AnchorStation = mCollapsedStations(index): Exit Function
    Next index
    index = SegmentIndex(segmentID)
    AnchorStation = mStations(index - 1) + parameter * (mStations(index) - mStations(index - 1))
End Function

' Находит существующий сегмент по точке на линии; параметр — доля его длины.
' Стык нескольких сегментов неоднозначен: для него задайте ID явно или выберите точку внутри участка.
Public Sub ResolveAnchor(ByVal x As Double, ByVal y As Double, _
        ByRef segmentID As String, ByRef parameter As Double)
    RequireSealed
    Dim i As Long, matched As Long, found As Long, candidate As Double, resolved As Double, matches As String
    For i = 1 To mCount
        If modGeometryPrimitives.PointParameter(mEdges(i), x, y, candidate) Then
            If matched > 0 Then matches = matches & ", "
            matches = matches & "'" & mSegmentIDs(i) & "'"
            matched = matched + 1: found = i: resolved = candidate
        End If
    Next i
    If matched = 0 Then PathError "Точка (" & CStr(x) & "; " & CStr(y) & ") не лежит на линии.", _
        "ANCHOR_NOT_ON_PATH", "", x, y
    If matched > 1 Then PathError "Точка (" & CStr(x) & "; " & CStr(y) & ") принадлежит нескольким сегментам: " & matches & _
        ". Задайте ID нужного сегмента явно или выберите точку внутри него.", "ANCHOR_AMBIGUOUS", matches, x, y
    segmentID = mSegmentIDs(found): parameter = resolved ' Выходные аргументы меняются только после успешного однозначного поиска.
End Sub

' Создаёт копию завершённой линии с новым ID, сохраняя имена сегментов
' и сведения о том, откуда они получены.
Public Function WithID(ByVal pathID As String) As CGeometryPath
    RequireSealed
    Dim result As CGeometryPath, i As Long
    Set result = New CGeometryPath: result.Initialize pathID
    For i = 1 To mCount
        result.Append mSegmentIDs(i), mEdges(i), mSourcePathIDs(i), mSourceSegmentIDs(i), mSourceStart(i), mSourceEnd(i)
    Next i
    CopyCollapsedAnchors result, 0#, Length, 0#
    CopyOffsetBoundary result
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
    If Abs(first - last) <= PATH_TOLERANCE And Not fullLoopIfSame Then
        If HasCollapsedSegmentID(startSegmentID) Or HasCollapsedSegmentID(endSegmentID) Then _
            PathError "После скругления или смещения контура между выбранными концами не осталось участка ненулевой длины. Измените размеры контура, радиусы скругления или величину смещения.", _
                "SEGMENT_DISAPPEARED", startSegmentID
    End If
    If Not mClosed Then PathError "Конец выбранного участка расположен раньше его начала. Измените направление обхода или поменяйте местами начало и конец.", "ANCHOR_DIRECTION", endSegmentID
    If Abs(first - last) <= PATH_TOLERANCE And Not fullLoopIfSame Then _
        PathError "Начало и конец выбранного участка совпадают. Укажите разные точки; для полного обхода замкнутого контура задайте fullLoopIfSame = True.", "ANCHORS_COINCIDE", startSegmentID
    Set result = New CGeometryPath: result.Initialize pathID
    AppendInterval result, first, mStations(mCount)
    AppendInterval result, 0#, last
    CopyOffsetBoundary result
    result.Seal (Abs(first - last) <= PATH_TOLERANCE)
    Set BetweenAnchors = result
End Function

' Находит точку по пройденному расстоянию от начала линии, мм, и направление
' касательной в ней. На стыке по умолчанию берёт следующий сегмент;
' incomingTangentAtJoint выбирает направление предыдущего.
Public Sub AtDistance(ByVal distance As Double, ByRef x As Double, ByRef y As Double, _
        ByRef tangentX As Double, ByRef tangentY As Double, Optional ByVal incomingTangentAtJoint As Boolean = False)
    RequireSealed
    distance = CheckedDistance(distance)
    Dim lo As Long, parameter As Double, norm As Double
    lo = SegmentAtDistance(distance, incomingTangentAtJoint)
    parameter = (distance - mStations(lo - 1)) / (mStations(lo) - mStations(lo - 1))
    parameter = GeomMax(0#, GeomMin(1#, parameter)) ' Исправляем только допустимую погрешность параметра у концов сегмента.
    modGeometryPrimitives.EdgePoint mEdges(lo), parameter, x, y
    modGeometryPrimitives.EdgeTangent mEdges(lo), parameter, tangentX, tangentY
    norm = Sqr(tangentX * tangentX + tangentY * tangentY)
    tangentX = tangentX / norm: tangentY = tangentY / norm
End Sub

' Проверяет перенос исходной позиции по нормали: используемая дуга не должна схлопнуться или вывернуться.
' Положение по длине и выбор касательной на стыке совпадают с AtDistance; новую раскладку не строит.
Public Sub ValidateNormalTransfer(ByVal distance As Double, ByVal normalOffset As Double, Optional ByVal incomingTangentAtJoint As Boolean = False)
    RequireSealed
    Dim segment As Long, radius As Double
    distance = CheckedDistance(distance): segment = SegmentAtDistance(distance, incomingTangentAtJoint)
    If Abs(mEdges(segment).Sweep) <= GEOM_TOLERANCE Then Exit Sub
    radius = Sqr((mEdges(segment).X1 - mEdges(segment).CenterX) ^ 2 + (mEdges(segment).Y1 - mEdges(segment).CenterY) ^ 2)
    If radius - Sgn(mEdges(segment).Sweep) * normalOffset <= GEOM_TOLERANCE Then _
        PathError "После смещения дополнительного ряда центры стержней достигают центра дуги или переходят на её другую сторону. " & _
            "Смещение ряда — " & CStr(Abs(normalOffset)) & " мм; радиус линии основного ряда — " & CStr(radius) & _
            " мм. Уменьшите смещение дополнительного ряда или диаметры стержней либо увеличьте радиус участка размещения.", "ROW_OFFSET_COLLAPSED", mSegmentIDs(segment), normalOffset, radius
End Sub


' Возвращает полную линию для совместного смещения всех её участков.
' Независимая открытая линия не имеет такой границы и возвращает Nothing.
Friend Function OffsetBoundary() As CGeometryPath
    RequireSealed
    If mClosed Then Set OffsetBoundary = Me Else Set OffsetBoundary = mOffsetBoundary
End Function

' Знак нормали участка относительно нормали полного контура.
Friend Property Get BoundaryDirection() As Double
    If mClosed Then BoundaryDirection = 1# Else BoundaryDirection = mBoundaryDirection
End Property

' Сохраняет связь до Seal. Полный контур уже завершён; циклических ссылок нет.
Friend Sub SetOffsetBoundary(ByVal boundary As CGeometryPath, ByVal direction As Double)
    RequireMutable
    If boundary Is Nothing Then Exit Sub
    If boundary Is Me Then PathError "Нельзя назначить линию собственной родительской границей."
    If Not boundary.IsClosed Then PathError "Полная граница должна быть замкнутой."
    If direction <> 1# And direction <> -1# Then PathError "Направление участка должно быть 1 или -1."
    Set mOffsetBoundary = boundary: mBoundaryDirection = direction
End Sub

' Передаёт неизменяемую полную границу при выделении участка или смене его имени.
Private Sub CopyOffsetBoundary(ByVal result As CGeometryPath, Optional ByVal reverse As Boolean = False)
    Dim boundary As CGeometryPath, direction As Double
    Set boundary = OffsetBoundary: direction = BoundaryDirection
    If reverse Then direction = -direction
    result.SetOffsetBoundary boundary, direction
End Sub

' Определяет, исчезнет ли при смещении хотя бы одна дуга этой линии.
' Проверку используют до подготовки смещённой линии, а не для изменения радиуса стержня.
Friend Function HasCollapsedNormalArc(ByVal normalOffset As Double) As Boolean
    RequireSealed
    Dim i As Long, radius As Double
    For i = 1 To mCount
        If mEdges(i).Sweep <> 0# Then
            radius = Sqr((mEdges(i).X1 - mEdges(i).CenterX) ^ 2 + (mEdges(i).Y1 - mEdges(i).CenterY) ^ 2)
            If radius - Sgn(mEdges(i).Sweep) * normalOffset <= GEOM_TOLERANCE Then HasCollapsedNormalArc = True: Exit Function
        End If
    Next i
End Function

' Проверяет попадание на любой фактический сегмент. Общий стык двух сегментов
' тоже принадлежит линии; однозначное имя сегмента здесь не требуется.
Friend Function ContainsPoint(ByVal x As Double, ByVal y As Double) As Boolean
    RequireSealed
    Dim i As Long, parameter As Double
    For i = 1 To mCount
        If modGeometryPrimitives.PointParameter(mEdges(i), x, y, parameter) Then ContainsPoint = True: Exit Function
    Next i
End Function

' Находит используемый сегмент для точки и проверки её нормального переноса одним правилом выбора стыка.
Private Function SegmentAtDistance(ByVal distance As Double, ByVal incomingTangentAtJoint As Boolean) As Long
    Dim lo As Long, hi As Long, middle As Long
    lo = 1: hi = mCount
    Do While lo < hi
        middle = lo + (hi - lo) \ 2 ' Двоичный поиск по накопленным длинам сегментов, без линейного обхода.
        If (Not incomingTangentAtJoint And distance < mStations(middle) - GEOM_TOLERANCE) Or _
                (incomingTangentAtJoint And distance <= mStations(middle) + GEOM_TOLERANCE) Then ' У численно близкого стыка сохраняем явно выбранную входящую или исходящую касательную.
            hi = middle
        Else
            lo = middle + 1
        End If
    Loop
    SegmentAtDistance = lo
End Function

' Возвращает наименьшую и наибольшую проекции линии на заданное направление.
' Учитывает крайние точки дуг, которые могут лежать между их концами.
Public Sub ProjectionBounds(ByVal directionX As Double, ByVal directionY As Double, _
        ByRef minimum As Double, ByRef maximum As Double)
    RequireSealed
    Dim norm As Double, i As Long
    norm = Sqr(directionX * directionX + directionY * directionY)
    If norm <= GEOM_TOLERANCE Then PathError "Направление проекции имеет нулевую длину."
    directionX = directionX / norm: directionY = directionY / norm
    minimum = 1E+100: maximum = -1E+100
    For i = 1 To mCount
        modGeometryPrimitives.AccumulateEdgeProjection mEdges(i), directionX, directionY, minimum, maximum
    Next i
End Sub

' Вырезает участок составной линии по расстояниям от её начала.
' Отсчёт длины на стыках не начинается заново; точные дуги и имена сегментов сохраняются.
Public Function Slice(ByVal startDistance As Double, ByVal endDistance As Double) As CGeometryPath
    RequireSealed
    startDistance = CheckedDistance(startDistance): endDistance = CheckedDistance(endDistance)
    If endDistance <= startDistance + PATH_TOLERANCE Then PathError "После обрезки не осталось участка линии. " & _
        "Начало участка: " & CStr(startDistance) & " мм; конец: " & CStr(endDistance) & " мм. " & _
        "Задайте конец дальше начала по направлению линии."
    Dim result As CGeometryPath, i As Long, first As Double, last As Double, segmentLength As Double, edge As TRegionEdge
    Set result = New CGeometryPath: result.Initialize mID
    For i = 1 To mCount
        first = GeomMax(startDistance, mStations(i - 1)): last = GeomMin(endDistance, mStations(i))
        If last > first + PATH_TOLERANCE Then
            segmentLength = mStations(i) - mStations(i - 1)
            modGeometryPrimitives.SubEdge mEdges(i), (first - mStations(i - 1)) / segmentLength, (last - mStations(i - 1)) / segmentLength, edge
            result.Append mSegmentIDs(i), edge, mSourcePathIDs(i), mSourceSegmentIDs(i), _
                OriginParameter(i, (first - mStations(i - 1)) / segmentLength), OriginParameter(i, (last - mStations(i - 1)) / segmentLength)
        End If
    Next i
    CopyCollapsedAnchors result, startDistance, endDistance, 0#
    CopyOffsetBoundary result
    result.Seal (mClosed And startDistance = 0# And endDistance = mStations(mCount))
    Set Slice = result
End Function

' Добавляет части линии между заданными расстояниями от её начала,
' сохраняя отрезки и дуги без замены приближённой ломаной.
Private Sub AppendInterval(ByVal result As CGeometryPath, ByVal first As Double, ByVal last As Double)
    If last <= first + PATH_TOLERANCE Then Exit Sub
    Dim part As CGeometryPath, i As Long, edge As TRegionEdge, name As String, suffix As Long, offset As Double
    offset = result.AccumulatedLength
    Set part = Slice(first, last)
    For i = 1 To part.SegmentCount
        part.CopySegment i, edge: name = part.SegmentID(i): suffix = 1
        Do While result.HasSegmentID(name) Or result.HasCollapsedSegmentID(name) Or HasCollapsedSegmentID(name) ' Суффикс не должен занять имя исчезнувшего участка, даже если его привязка копируется позже.
            suffix = suffix + 1: name = part.SegmentID(i) & ".Part" & CStr(suffix)
        Loop
        result.Append name, edge, part.SourcePathID(i), part.SourceSegmentID(i), part.SourceStartParameter(i), part.SourceEndParameter(i)
    Next i
    part.CopyCollapsedAnchors result, 0#, part.Length, offset
End Sub

' Сохраняет место исчезнувшего участка как именованную точку привязки.
' Сам участок не добавляется. Вызывать можно, когда построитель знает точное место его
' исчезновения.
Friend Sub AddCollapsedSegment(ByVal segmentID As String, Optional ByVal station As Double = -1#)
    RequireMutable
    If Len(Trim$(segmentID)) = 0 Then PathError "Не задан ID исчезнувшего сегмента."
    If HasSegmentID(segmentID) Then PathError "Повторный ID сегмента '" & segmentID & "'."
    Dim i As Long, currentLength As Double
    For i = 1 To mCollapsedCount
        If mCollapsedIDs(i) = segmentID Then PathError "Повторный ID привязки '" & segmentID & "'."
    Next i
    If mCount > 0 Then currentLength = mStations(mCount)
    If station = -1# Then station = currentLength ' Служебное -1 означает текущий конец ещё собираемой линии.
    If station < -PATH_TOLERANCE Or station > currentLength + PATH_TOLERANCE Then PathError "Привязка исчезнувшего сегмента вне пути."
    station = GeomMax(0#, GeomMin(currentLength, station))
    mCollapsedCount = mCollapsedCount + 1
    ReDim Preserve mCollapsedIDs(1 To mCollapsedCount): ReDim Preserve mCollapsedStations(1 To mCollapsedCount)
    mCollapsedIDs(mCollapsedCount) = segmentID: mCollapsedStations(mCollapsedCount) = station
End Sub

' Переносит только привязки внутри выбранного интервала. При обратном обходе
' отсчёт меняется вместе с геометрией; на замыкающем стыке достаточно одного имени.
' Для Offset передаются новые длины до концов сегментов: привязка остаётся на том же стыке.
Friend Sub CopyCollapsedAnchors(ByVal result As CGeometryPath, ByVal first As Double, ByVal last As Double, _
        ByVal offset As Double, Optional ByVal reverse As Boolean = False, Optional ByRef segmentStations As Variant)
    Dim i As Long, j As Long, station As Double, parameter As Double
    For i = 1 To mCollapsedCount
        station = mCollapsedStations(i)
        If station >= first And station <= last Then
            If Not result.HasCollapsedSegmentID(mCollapsedIDs(i)) Then
                If Not IsMissing(segmentStations) Then
                    For j = 1 To mCount
                        If station <= mStations(j) Then
                            parameter = (station - mStations(j - 1)) / (mStations(j) - mStations(j - 1))
                            station = segmentStations(j - 1) + parameter * (segmentStations(j) - segmentStations(j - 1))
                            Exit For
                        End If
                    Next j
                End If
                If reverse Then station = last - station Else station = station - first ' Привязка отсчитывается уже от начала нового участка.
                result.AddCollapsedSegment mCollapsedIDs(i), offset + station
            End If
        End If
    Next i
End Sub

' Возвращает длину уже добавленных сегментов, чтобы продолжить сборку линии до Seal.
Friend Property Get AccumulatedLength() As Double
    If mCount > 0 Then AccumulatedLength = mStations(mCount)
End Property

' Проверяет, сохранено ли имя как точка привязки исчезнувшего участка.
' Такая привязка отличается от существующего сегмента и от неизвестного имени.
Friend Function HasCollapsedSegmentID(ByVal segmentID As String) As Boolean
    Dim i As Long
    For i = 1 To mCollapsedCount
        If mCollapsedIDs(i) = segmentID Then HasCollapsedSegmentID = True: Exit Function
    Next i
End Function

' Переводит локальный параметр обрезанного сегмента в параметр исходного сегмента.
Private Function OriginParameter(ByVal index As Long, ByVal parameter As Double) As Double
    OriginParameter = mSourceStart(index) + parameter * (mSourceEnd(index) - mSourceStart(index))
End Function

' Проверяет, остался ли сегмент с указанным именем в линии.
' Отрезок, исчезнувший при скруглении, существующим сегментом не считается.
Friend Function HasSegmentID(ByVal segmentID As String) As Boolean
    Dim i As Long
    For i = 1 To mCount
        If StrComp(mSegmentIDs(i), segmentID, vbBinaryCompare) = 0 Then HasSegmentID = True: Exit Function
    Next i
End Function

' Создаёт копию линии с обратным направлением обхода и порядком сегментов.
' Меняет знаки углов дуг и направление отсчёта параметров исходных сегментов.
Public Function Reversed() As CGeometryPath
    RequireSealed
    Dim result As CGeometryPath, i As Long, edge As TRegionEdge
    Set result = New CGeometryPath: result.Initialize mID
    For i = mCount To 1 Step -1
        LSet edge = mEdges(i): modGeometryPrimitives.ReverseEdge edge
        result.Append mSegmentIDs(i), edge, mSourcePathIDs(i), mSourceSegmentIDs(i), mSourceEnd(i), mSourceStart(i)
    Next i
    CopyCollapsedAnchors result, 0#, Length, 0#, True
    CopyOffsetBoundary result, True
    result.Seal mClosed
    Set Reversed = result
End Function

' Создаёт преобразованную линию: отражает относительно локальной оси X,
' поворачивает вокруг (0, 0), затем переносит. При отражении меняет знак
' угла обхода дуг; имена сегментов сохраняет.
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
    CopyCollapsedAnchors result, 0#, Length, 0#
    If Not mOffsetBoundary Is Nothing Then result.SetOffsetBoundary mOffsetBoundary.Transformed(shiftX, shiftY, rotation, reflectX), mBoundaryDirection
    result.Seal mClosed
    Set Transformed = result
End Function

' Ищет ближайшее пересечение в положительном направлении проекции
' с фактическими отрезками и дугами целевой линии. Без пересечения возвращает False;
' продолжения линий до бесконечности не используются.
Public Function RayHit(ByVal originX As Double, ByVal originY As Double, ByVal directionX As Double, _
        ByVal directionY As Double, ByRef pathDistance As Double, ByRef rayDistance As Double) As Boolean
    RequireSealed
    Dim i As Long, distance As Double, parameter As Double
    pathDistance = 0#: rayDistance = 0#
    If Sqr(directionX * directionX + directionY * directionY) <= GEOM_TOLERANCE Then PathError "Направление луча имеет нулевую длину."
    For i = 1 To mCount
        If modGeometryPrimitives.EdgeRayHit(mEdges(i), originX, originY, directionX, directionY, distance, parameter) Then
            If Not RayHit Or distance < rayDistance Then
                rayDistance = distance
                pathDistance = mStations(i - 1) + parameter * (mStations(i) - mStations(i - 1))
                RayHit = True
            End If
        End If
    Next i
End Function

' Копирует данные аналитического сегмента в переданную структуру без изменения пути.
Friend Sub CopySegment(ByVal index As Long, ByRef edge As TRegionEdge)
    RequireSealed
    CheckIndex index
    LSet edge = mEdges(index)
End Sub

' Добавляет один отрезок или дугу с уникальным именем и проверкой непрерывности.
' Необязательные source-параметры сохраняют происхождение и интервал исходного сегмента.
Friend Sub Append(ByVal segmentID As String, ByRef edge As TRegionEdge, Optional ByVal sourcePathID As String = "", _
        Optional ByVal sourceSegmentID As String = "", Optional ByVal sourceStart As Double = 0#, Optional ByVal sourceEnd As Double = 1#)
    RequireMutable
    If Len(Trim$(segmentID)) = 0 Then PathError "Не задан ID сегмента."
    Dim i As Long, segmentLength As Double
    For i = 1 To mCount
        If mSegmentIDs(i) = segmentID Then PathError "Повторный ID сегмента '" & segmentID & "'."
    Next i
    For i = 1 To mCollapsedCount
        If mCollapsedIDs(i) = segmentID Then PathError "Повторный ID привязки '" & segmentID & "'."
    Next i
    segmentLength = modGeometryPrimitives.EdgeLength(edge)
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

' Отражает, поворачивает и переносит точку; координаты остаются в мм.
Private Sub TransformPoint(ByRef x As Double, ByRef y As Double, ByVal shiftX As Double, ByVal shiftY As Double, _
        ByVal cosine As Double, ByVal sine As Double, ByVal reflectX As Boolean)
    Dim newX As Double
    If reflectX Then y = -y
    newX = cosine * x - sine * y + shiftX
    y = sine * x + cosine * y + shiftY: x = newX
End Sub

' Сравнивает координаты двух точек в геометрическом допуске.
Private Function PointsCoincide(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Boolean
    PointsCoincide = (Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2) <= PATH_TOLERANCE)
End Function

' Проверяет расстояние от начала линии до точки. Исправляет только малую
' погрешность у концов; заведомо недопустимое расстояние вызывает ошибку.
Private Function CheckedDistance(ByVal distance As Double) As Double
    If distance < -PATH_TOLERANCE Or distance > mStations(mCount) + PATH_TOLERANCE Then PathError "Расстояние " & CStr(distance) & _
        " мм выходит за пределы линии длиной " & CStr(mStations(mCount)) & " мм. Задайте расстояние от 0 до полной длины линии."
    CheckedDistance = GeomMax(0#, GeomMin(mStations(mCount), distance))
End Function

' Запрещает изменение линии после её завершения методом Seal.
Private Sub RequireMutable()
    If Len(mID) = 0 Then PathError "Путь не инициализирован."
    If mSealed Then PathError "Опубликованный путь нельзя изменять."
End Sub

' Требует завершённую проверенную геометрию перед её использованием.
Private Sub RequireSealed()
    If Not mSealed Then PathError "Путь еще не подготовлен и не опубликован."
End Sub

' Проверяет индекс перед обращением к массивам геометрического объекта.
Private Sub CheckIndex(ByVal index As Long)
    If index < 1 Or index > mCount Then PathError "Номер сегмента вне допустимого диапазона." & " Запрошен номер " & CStr(index) & "; доступно элементов: " & CStr(mCount) & "."
End Sub

' Выдаёт геометрическую ошибку с ID пути и фактической причиной.
Private Sub PathError(ByVal description As String, Optional ByVal code As String = "PATH_INVALID", _
        Optional ByVal relatedEntityID As String = "", Optional ByVal detail1 As Double = 0#, Optional ByVal detail2 As Double = 0#)
    Set mDiagnostic = New CGeometryDiagnostic
    mDiagnostic.Initialize code, mID, relatedEntityID, description, detail1, detail2
    Err.Raise vbObjectError + 5520, "CGeometryPath", SectionRecipeErrorText(description, context:="Геометрическая линия '" & mID & "':")
End Sub
