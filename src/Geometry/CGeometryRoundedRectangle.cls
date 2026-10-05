VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGeometryRoundedRectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Математическая геометрия скругленного прямоугольника с Simple/Tapered торцами
' ==========================================================================
' Класс описывает бетонную область RoundedRectangle:
' центральный прямой участок B x H и независимые левую/правую боковые части.
' Боковая часть может быть Simple - вертикальная вставка с радиусами R1, либо
' Tapered - скругленное сопряжение R1, наклонные участки и носовой радиус R2.
'
' Для расчетного ядра класс остается обычной ISectionGeometry: дает габариты,
' ContainsPoint, расчетную площадь/центр по дискретизированному контуру и
' смысловые точки для арматуры/аннотаций. Сетка, NDM и solver не знают о
' внутреннем устройстве торцов.

Implements ISectionGeometry

Private Const SIDE_SIMPLE As String = "simple" ' Вертикальная боковая вставка с верхним/нижним R1.
Private Const SIDE_TAPERED As String = "tapered" ' Наклонная боковая часть с дополнительным носовым R2.
Private Const ARC_SEGMENTS As Long = 12 ' Число малых отрезков на дугу; достаточно для схемы и mesh boundary.

Private Type TRoundedRawPoint
    X As Double
    Y As Double
    Radius As Double
End Type

Private mB As Double ' Длина горизонтальной расчетной базы между левым и правым торцом.
Private mH As Double ' Полная высота сечения.
Private mLeftType As String
Private mRightType As String
Private mLeftW As Double ' Вынос tapered-носа влево; для Simple не влияет на форму.
Private mRightW As Double ' Вынос tapered-носа вправо; для Simple не влияет на форму.
Private mLeftR1 As Double ' Радиус верхнего/нижнего сопряжения слева.
Private mRightR1 As Double ' Радиус верхнего/нижнего сопряжения справа.
Private mLeftR2 As Double ' Носовой радиус слева для Tapered.
Private mRightR2 As Double ' Носовой радиус справа для Tapered.
Private mCenterX As Double
Private mCenterY As Double
Private mInitialized As Boolean
Private mValidationReady As Boolean ' Результат проверки относится к текущему набору параметров.
Private mGeometryValid As Boolean ' Горячие запросы точки не перестраивают и не проверяют контур.
Private mValidationMessage As String ' Сохраненная причина невалидности до следующего Initialize.

Private mPointX() As Double ' Дискретизированный внешний контур в координатах модели.
Private mPointY() As Double
Private mPointCount As Long
Private mArea As Double
Private mCentroidX As Double
Private mCentroidY As Double
Private mMinX As Double
Private mMaxX As Double
Private mMinY As Double
Private mMaxY As Double

' Возвращает внешний контур как набор прямых и дуг. Расчет использует те же
' raw-точки, offset-вершины и касательные, что и дискретизированный контур,
' но не дробит радиусы на ломаную для визуализации схемы.
Public Sub GetContourSegments(ByRef segmentType() As String, ByRef startX() As Double, ByRef startY() As Double, _
        ByRef endX() As Double, ByRef endY() As Double, ByRef centerX() As Double, ByRef centerY() As Double, _
        ByRef radiusValue() As Double, ByRef sweepAngle() As Double, ByRef segmentCount As Long)
    Erase segmentType
    Erase startX
    Erase startY
    Erase endX
    Erase endY
    Erase centerX
    Erase centerY
    Erase radiusValue
    Erase sweepAngle
    segmentCount = 0

    Dim raw() As TRoundedRawPoint
    Dim rawCount As Long
    BuildRawContour raw, rawCount
    If rawCount < 3 Then Exit Sub

    Dim offsetX() As Double
    Dim offsetY() As Double
    BuildOffsetVertices raw, rawCount, 0#, offsetX, offsetY

    Dim cornerStartX() As Double, cornerStartY() As Double
    Dim cornerEndX() As Double, cornerEndY() As Double
    Dim arcCenterX() As Double, arcCenterY() As Double
    Dim arcRadius() As Double, arcSweep() As Double
    Dim hasArc() As Boolean
    ReDim cornerStartX(1 To rawCount)
    ReDim cornerStartY(1 To rawCount)
    ReDim cornerEndX(1 To rawCount)
    ReDim cornerEndY(1 To rawCount)
    ReDim arcCenterX(1 To rawCount)
    ReDim arcCenterY(1 To rawCount)
    ReDim arcRadius(1 To rawCount)
    ReDim arcSweep(1 To rawCount)
    ReDim hasArc(1 To rawCount)

    Dim i As Long
    For i = 1 To rawCount
        GetRoundedCornerGeometry raw, offsetX, offsetY, rawCount, _
            PreviousIndex(i, rawCount), i, NextIndex(i, rawCount), 0#, _
            cornerStartX(i), cornerStartY(i), cornerEndX(i), cornerEndY(i), _
            arcCenterX(i), arcCenterY(i), arcRadius(i), arcSweep(i), hasArc(i)
    Next i

    For i = 1 To rawCount
        Dim prevIndex As Long
        prevIndex = PreviousIndex(i, rawCount)
        AddContourSegment segmentType, startX, startY, endX, endY, centerX, centerY, _
            radiusValue, sweepAngle, segmentCount, "LINE", _
            cornerEndX(prevIndex), cornerEndY(prevIndex), cornerStartX(i), cornerStartY(i), _
            0#, 0#, 0#, 0#
        If hasArc(i) Then
            AddContourSegment segmentType, startX, startY, endX, endY, centerX, centerY, _
                radiusValue, sweepAngle, segmentCount, "ARC", _
                cornerStartX(i), cornerStartY(i), cornerEndX(i), cornerEndY(i), _
                arcCenterX(i), arcCenterY(i), arcRadius(i), arcSweep(i)
        End If
    Next i
End Sub

' Краткая программная инициализация симметричного Simple/Simple-варианта.
' Основной путь Config использует InitializeSides и задает стороны явно.
Public Sub Initialize(ByVal widthMm As Double, ByVal heightMm As Double, _
        Optional ByVal radiusTLmm As Double = 0#, Optional ByVal radiusTRmm As Double = 0#, _
        Optional ByVal radiusBRmm As Double = 0#, Optional ByVal radiusBLmm As Double = 0#, _
        Optional ByVal centerXMm As Double = 0#, Optional ByVal centerYMm As Double = 0#)

    Dim leftR1 As Double
    Dim rightR1 As Double
    leftR1 = GeomMax(radiusTLmm, radiusBLmm)
    rightR1 = GeomMax(radiusTRmm, radiusBRmm)
    If radiusTLmm < 0# Or radiusBLmm < 0# Then leftR1 = -1#
    If radiusTRmm < 0# Or radiusBRmm < 0# Then rightR1 = -1#

    InitializeSides widthMm, heightMm, "Simple", "Simple", 0#, 0#, _
        leftR1, rightR1, 0#, 0#, _
        centerXMm, centerYMm
End Sub

' Инициализирует новую форму RoundedRectangle. B задает длину горизонтальной
' базы, H - высоту, а левая/правая боковые части управляются независимо.
Public Sub InitializeSides(ByVal bMm As Double, ByVal hMm As Double, _
        ByVal leftType As String, ByVal rightType As String, _
        ByVal leftWMm As Double, ByVal rightWMm As Double, _
        ByVal leftR1Mm As Double, ByVal rightR1Mm As Double, _
        ByVal leftR2Mm As Double, ByVal rightR2Mm As Double, _
        Optional ByVal centerXMm As Double = 0#, Optional ByVal centerYMm As Double = 0#)

    mValidationReady = False
    mGeometryValid = False
    mValidationMessage = vbNullString
    mB = bMm
    mH = hMm
    mLeftType = NormalizeSideType(leftType)
    mRightType = NormalizeSideType(rightType)
    mLeftW = leftWMm
    mRightW = rightWMm
    mLeftR1 = leftR1Mm
    mRightR1 = rightR1Mm
    mLeftR2 = leftR2Mm
    mRightR2 = rightR2Mm
    mCenterX = centerXMm
    mCenterY = centerYMm
    mInitialized = True
    RebuildContour
    Dim message As String
    mGeometryValid = IsValid(message)
End Sub

Public Property Get B() As Double
    B = mB
End Property

Public Property Get H() As Double
    H = mH
End Property

Public Property Get Width() As Double
    Width = mMaxX - mMinX
End Property

Public Property Get Height() As Double
    Height = mMaxY - mMinY
End Property

Public Property Get LeftType() As String
    LeftType = DisplaySideType(mLeftType)
End Property

Public Property Get RightType() As String
    RightType = DisplaySideType(mRightType)
End Property

Public Property Get LeftW() As Double
    LeftW = mLeftW
End Property

Public Property Get RightW() As Double
    RightW = mRightW
End Property

Public Property Get LeftR1() As Double
    LeftR1 = mLeftR1
End Property

Public Property Get RightR1() As Double
    RightR1 = mRightR1
End Property

Public Property Get LeftR2() As Double
    LeftR2 = mLeftR2
End Property

Public Property Get RightR2() As Double
    RightR2 = mRightR2
End Property

Public Property Get CenterX() As Double
    CenterX = mCenterX
End Property

Public Property Get CenterY() As Double
    CenterY = mCenterY
End Property

Public Property Get LeftBaseX() As Double
    LeftBaseX = mCenterX - mB / 2#
End Property

Public Property Get RightBaseX() As Double
    RightBaseX = mCenterX + mB / 2#
End Property

Public Property Get TopY() As Double
    TopY = mCenterY + mH / 2#
End Property

Public Property Get BottomY() As Double
    BottomY = mCenterY - mH / 2#
End Property

Public Property Get MinX() As Double
    MinX = mMinX
End Property

Public Property Get MaxX() As Double
    MaxX = mMaxX
End Property

Public Property Get MinY() As Double
    MinY = mMinY
End Property

Public Property Get MaxY() As Double
    MaxY = mMaxY
End Property

' Проверяет точку через быстрый bounding-box и точный polygon-contour.
Public Function ContainsPoint(ByVal X As Double, ByVal Y As Double) As Boolean
    Dim message As String
    If Not IsValid(message) Then Exit Function
    If X < mMinX - GEOM_TOLERANCE Or X > mMaxX + GEOM_TOLERANCE Then Exit Function
    If Y < mMinY - GEOM_TOLERANCE Or Y > mMaxY + GEOM_TOLERANCE Then Exit Function
    ContainsPoint = PointInContour(X, Y)
End Function

' Возвращает ту же проверку области через предусмотренное имя интерфейса.
Public Function IsPointInside(ByVal X As Double, ByVal Y As Double) As Boolean
    IsPointInside = ContainsPoint(X, Y)
End Function

' Возвращает проверку текущей неизменяемой геометрии. Initialize инвалидирует
' прежний результат и выполняет проверку один раз после подготовки контура.
Public Function IsValid(ByRef message As String) As Boolean
    If Not mValidationReady Then
        mGeometryValid = ValidateCurrentGeometry(mValidationMessage)
        mValidationReady = True
    End If
    message = mValidationMessage
    IsValid = mGeometryValid
End Function

' Проверяет размеры, сопряжения радиусов и невырожденность готового контура.
' Численная геометрия остается общей с ContainsPoint и построением арматуры.
Private Function ValidateCurrentGeometry(ByRef message As String) As Boolean
    message = vbNullString

    If Not mInitialized Then
        message = "Геометрия скругленного прямоугольника не инициализирована."
        Exit Function
    End If
    If mB <= 0# Then
        message = "Размер B скругленного прямоугольника должен быть больше нуля."
        Exit Function
    End If
    If mH <= 0# Then
        message = "Размер H скругленного прямоугольника должен быть больше нуля."
        Exit Function
    End If
    If Len(mLeftType) = 0 Or Len(mRightType) = 0 Then
        message = "Тип левой и правой стороны должен быть Simple или Tapered."
        Exit Function
    End If
    If mLeftW < 0# Or mRightW < 0# Then
        message = "Размер W для левой и правой стороны не должен быть отрицательным."
        Exit Function
    End If
    If mLeftR1 < 0# Or mRightR1 < 0# Or mLeftR2 < 0# Or mRightR2 < 0# Then
        message = "Радиусы R1 и R2 не должны быть отрицательными."
        Exit Function
    End If
    If mLeftType = SIDE_TAPERED And mLeftW <= GEOM_TOLERANCE Then
        message = "Для левой стороны Tapered размер W должен быть больше нуля."
        Exit Function
    End If
    If mRightType = SIDE_TAPERED And mRightW <= GEOM_TOLERANCE Then
        message = "Для правой стороны Tapered размер W должен быть больше нуля."
        Exit Function
    End If
    If Not RoundedCornersFit(message) Then
        Exit Function
    End If
    If mPointCount < 3 Or mArea <= GEOM_TOLERANCE Then
        RebuildContour
    End If
    If mPointCount < 3 Or mArea <= GEOM_TOLERANCE Then
        message = "Расчетный контур скругленного прямоугольника вырожден."
        Exit Function
    End If

    ValidateCurrentGeometry = True
End Function

' Проверяет, что касательные участки соседних скруглений помещаются на каждом
' прямом или наклонном сегменте исходной ломаной. Для Simple это соответствует
' ограничению 2*R <= H; для Tapered допустим R1 больше H/2,
' если касательные участки помещаются и соседние дуги не пересекаются.
Private Function RoundedCornersFit(ByRef message As String) As Boolean
    Dim raw() As TRoundedRawPoint
    Dim rawCount As Long
    BuildRawContour raw, rawCount
    If rawCount < 3 Then
        message = "Контур скругленного прямоугольника не задан."
        Exit Function
    End If

    Dim tangentDistance() As Double
    ReDim tangentDistance(1 To rawCount)

    Dim i As Long
    For i = 1 To rawCount
        tangentDistance(i) = CornerTangentDistance(raw, rawCount, PreviousIndex(i, rawCount), i, NextIndex(i, rawCount))
        If tangentDistance(i) < 0# Then
            message = "Не удалось проверить радиусы RoundedRectangle: есть вырожденный участок контура."
            Exit Function
        End If
    Next i

    For i = 1 To rawCount
        Dim j As Long
        j = NextIndex(i, rawCount)

        Dim edgeLength As Double
        edgeLength = DistanceBetweenRawPoints(raw, i, j)
        If edgeLength <= GEOM_TOLERANCE Then
            message = "Не удалось проверить радиусы RoundedRectangle: есть нулевой участок контура."
            Exit Function
        End If

        If tangentDistance(i) + tangentDistance(j) > edgeLength + GEOM_TOLERANCE Then
            message = "Радиусы RoundedRectangle пересекаются на участке " & RoundedEdgeLabel(i, j, rawCount) & _
                ". Уменьшите соседние R1/R2 или увеличьте соответствующий размер."
            Exit Function
        End If
    Next i

    RoundedCornersFit = True
End Function

' Возвращает длину участка исходной ломаной между двумя raw-вершинами.
Private Function DistanceBetweenRawPoints(ByRef raw() As TRoundedRawPoint, ByVal firstIndex As Long, ByVal secondIndex As Long) As Double
    Dim dx As Double
    Dim dy As Double
    dx = raw(secondIndex).X - raw(firstIndex).X
    dy = raw(secondIndex).Y - raw(firstIndex).Y
    DistanceBetweenRawPoints = Sqr(dx * dx + dy * dy)
End Function

' Считает расстояние от вершины до точки касания дуги вдоль каждого соседнего
' сегмента без защитного clamp-а. Именно это расстояние нужно для проверки,
' пересекаются ли соседние скругления на общей грани.
Private Function CornerTangentDistance(ByRef raw() As TRoundedRawPoint, ByVal rawCount As Long, _
        ByVal prevIndex As Long, ByVal cornerIndex As Long, ByVal nextIndexValue As Long) As Double
    If raw(cornerIndex).Radius <= GEOM_TOLERANCE Then Exit Function

    Dim vPrevX As Double
    Dim vPrevY As Double
    Dim vNextX As Double
    Dim vNextY As Double
    Dim prevLength As Double
    Dim nextLength As Double
    vPrevX = raw(prevIndex).X - raw(cornerIndex).X
    vPrevY = raw(prevIndex).Y - raw(cornerIndex).Y
    vNextX = raw(nextIndexValue).X - raw(cornerIndex).X
    vNextY = raw(nextIndexValue).Y - raw(cornerIndex).Y
    prevLength = Sqr(vPrevX * vPrevX + vPrevY * vPrevY)
    nextLength = Sqr(vNextX * vNextX + vNextY * vNextY)
    If prevLength <= GEOM_TOLERANCE Or nextLength <= GEOM_TOLERANCE Then
        CornerTangentDistance = -1#
        Exit Function
    End If

    vPrevX = vPrevX / prevLength
    vPrevY = vPrevY / prevLength
    vNextX = vNextX / nextLength
    vNextY = vNextY / nextLength

    Dim dotValue As Double
    dotValue = vPrevX * vNextX + vPrevY * vNextY
    If dotValue > 1# Then dotValue = 1#
    If dotValue < -1# Then dotValue = -1#

    Dim angleValue As Double
    angleValue = Atn2Local(Sqr(GeomMax(1# - dotValue * dotValue, 0#)), dotValue)
    If angleValue <= GEOM_TOLERANCE Or Abs(GEOM_PI - angleValue) <= GEOM_TOLERANCE Then
        CornerTangentDistance = -1#
    Else
        CornerTangentDistance = raw(cornerIndex).Radius / Tan(angleValue / 2#)
    End If
End Function

' Дает короткое пользовательское имя участка, на котором пересеклись дуги.
Private Function RoundedEdgeLabel(ByVal firstIndex As Long, ByVal secondIndex As Long, ByVal rawCount As Long) As String
    If rawCount = 4 Then
        Select Case firstIndex
            Case 1: RoundedEdgeLabel = "верхней грани B"
            Case 2: RoundedEdgeLabel = "правой грани H"
            Case 3: RoundedEdgeLabel = "нижней грани B"
            Case Else: RoundedEdgeLabel = "левой грани H"
        End Select
        Exit Function
    End If

    If rawCount = 5 Then
        Select Case firstIndex
            Case 1: RoundedEdgeLabel = "верхней грани B"
            Case 2: RoundedEdgeLabel = "наклонной грани Tapered"
            Case 3: RoundedEdgeLabel = "наклонной грани Tapered"
            Case 4: RoundedEdgeLabel = "нижней грани B"
            Case Else: RoundedEdgeLabel = "левой грани H"
        End Select
        Exit Function
    End If

    If firstIndex = 1 Or firstIndex = 4 Then
        RoundedEdgeLabel = "верхней/нижней грани B"
    Else
        RoundedEdgeLabel = "наклонной грани Tapered"
    End If
End Function

' Возвращает площадь расчетного polygon-представления скругленного контура.
Public Function AnalyticalArea(ByRef available As Boolean) As Double
    Dim message As String
    available = IsValid(message)
    If available Then AnalyticalArea = mArea
End Function

' Возвращает центр тяжести polygon-представления скругленного контура.
Public Sub AnalyticalCentroid(ByRef available As Boolean, ByRef X As Double, ByRef Y As Double)
    Dim message As String
    available = IsValid(message)
    If Not available Then Exit Sub
    X = mCentroidX
    Y = mCentroidY
End Sub

' Возвращает точки внешнего polygon-контура для сетки, габаритов и визуализации.
Public Sub GetExtremePoints(ByRef pointX() As Double, ByRef pointY() As Double)
    ReDim pointX(1 To mPointCount)
    ReDim pointY(1 To mPointCount)

    Dim i As Long
    For i = 1 To mPointCount
        pointX(i) = mPointX(i)
        pointY(i) = mPointY(i)
    Next i
End Sub

' Возвращает линию раскладки верхней или нижней горизонтальной грани B.
' Линия уже смещена внутрь на as и идет только по прямому участку между
' началом/концом скруглений R1.
Public Sub GetHorizontalRebarLine(ByVal isTop As Boolean, ByVal axisDistance As Double, _
        ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double, _
        ByRef normalX As Double, ByRef normalY As Double)
    Dim raw() As TRoundedRawPoint
    Dim rawCount As Long
    BuildRawContour raw, rawCount

    Dim offsetX() As Double
    Dim offsetY() As Double
    BuildOffsetVertices raw, rawCount, axisDistance, offsetX, offsetY

    Dim startX As Double, startY As Double, endX As Double, endY As Double
    If isTop Then
        CornerTangentPoints raw, offsetX, offsetY, rawCount, 1, 2, axisDistance, startX, startY, x1, y1
        CornerTangentPoints raw, offsetX, offsetY, 1, 2, NextIndex(2, rawCount), axisDistance, x2, y2, endX, endY
        normalX = 0#: normalY = 1#
    Else
        Dim bottomRight As Long
        Dim bottomLeft As Long
        bottomRight = RightBottomRawIndex()
        bottomLeft = bottomRight + 1
        CornerTangentPoints raw, offsetX, offsetY, PreviousIndex(bottomLeft, rawCount), bottomLeft, NextIndex(bottomLeft, rawCount), _
            axisDistance, x1, y1, endX, endY
        CornerTangentPoints raw, offsetX, offsetY, PreviousIndex(bottomRight, rawCount), bottomRight, NextIndex(bottomRight, rawCount), _
            axisDistance, startX, startY, x2, y2
        normalX = 0#: normalY = -1#
    End If
End Sub

' Возвращает боковую траекторию арматуры H, смещенную внутрь на as.
' Начальная точка берется в касательной точке верхней горизонтальной линии B,
' конечная - в касательной точке нижней горизонтальной линии B. Между ними
' путь идет по фактической боковой форме: Simple дает вертикальную вставку со
' скруглениями, Tapered - скругление R1, наклонную часть, носовой R2 и нижнюю
' наклонную часть. Builder дальше делит эту длину на n+1 равных интервалов.
Public Sub GetSideRebarPath(ByVal isLeft As Boolean, ByVal axisDistance As Double, _
        ByRef pointX() As Double, ByRef pointY() As Double)
    Erase pointX
    Erase pointY
    Dim count As Long

    Dim raw() As TRoundedRawPoint
    Dim rawCount As Long
    BuildRawContour raw, rawCount

    Dim offsetX() As Double
    Dim offsetY() As Double
    BuildOffsetVertices raw, rawCount, axisDistance, offsetX, offsetY

    Dim bottomRight As Long
    Dim bottomLeft As Long
    bottomRight = RightBottomRawIndex()
    bottomLeft = bottomRight + 1

    If isLeft Then
        Dim leftTopNext As Long ' Следующая вершина после верхнего левого угла вдоль H-линии.
        Dim leftBottomPrev As Long ' Предыдущая вершина перед нижним левым углом вдоль H-линии.
        If mLeftType = SIDE_TAPERED Then
            leftTopNext = rawCount
            leftBottomPrev = rawCount
        Else
            leftTopNext = bottomLeft
            leftBottomPrev = 1
        End If
        AppendCornerFromOffset raw, offsetX, offsetY, rawCount, 2, 1, leftTopNext, axisDistance, pointX, pointY, count
        If mLeftType = SIDE_TAPERED Then
            AppendCornerFromOffset raw, offsetX, offsetY, rawCount, 1, rawCount, bottomLeft, axisDistance, pointX, pointY, count
        End If
        AppendCornerFromOffset raw, offsetX, offsetY, rawCount, leftBottomPrev, bottomLeft, bottomRight, axisDistance, pointX, pointY, count
    Else
        Dim rightTopNext As Long ' Следующая вершина после верхнего правого угла вдоль H-линии.
        Dim rightBottomPrev As Long ' Предыдущая вершина перед нижним правым углом вдоль H-линии.
        If mRightType = SIDE_TAPERED Then
            rightTopNext = 3
            rightBottomPrev = 3
        Else
            rightTopNext = bottomRight
            rightBottomPrev = 2
        End If
        AppendCornerFromOffset raw, offsetX, offsetY, rawCount, 1, 2, rightTopNext, axisDistance, pointX, pointY, count
        If mRightType = SIDE_TAPERED Then
            AppendCornerFromOffset raw, offsetX, offsetY, rawCount, 2, 3, bottomRight, axisDistance, pointX, pointY, count
        End If
        AppendCornerFromOffset raw, offsetX, offsetY, rawCount, rightBottomPrev, bottomRight, bottomLeft, axisDistance, pointX, pointY, count
    End If
    If count > 0 Then
        ReDim Preserve pointX(1 To count)
        ReDim Preserve pointY(1 To count)
    End If
End Sub

Private Property Get ISectionGeometry_MinX() As Double
    ISectionGeometry_MinX = MinX
End Property

Private Property Get ISectionGeometry_MaxX() As Double
    ISectionGeometry_MaxX = MaxX
End Property

Private Property Get ISectionGeometry_MinY() As Double
    ISectionGeometry_MinY = MinY
End Property

Private Property Get ISectionGeometry_MaxY() As Double
    ISectionGeometry_MaxY = MaxY
End Property

' Делегирует проверку точки в реализацию класса для интерфейса ISectionGeometry.
Private Function ISectionGeometry_ContainsPoint(ByVal X As Double, ByVal Y As Double) As Boolean
    ISectionGeometry_ContainsPoint = ContainsPoint(X, Y)
End Function

' Делегирует IsPointInside в реализацию класса для интерфейса ISectionGeometry.
Private Function ISectionGeometry_IsPointInside(ByVal X As Double, ByVal Y As Double) As Boolean
    ISectionGeometry_IsPointInside = IsPointInside(X, Y)
End Function

' Делегирует валидацию в реализацию класса для интерфейса ISectionGeometry.
Private Function ISectionGeometry_IsValid(ByRef message As String) As Boolean
    ISectionGeometry_IsValid = IsValid(message)
End Function

' Делегирует аналитическую площадь в реализацию класса для интерфейса ISectionGeometry.
Private Function ISectionGeometry_AnalyticalArea(ByRef available As Boolean) As Double
    ISectionGeometry_AnalyticalArea = AnalyticalArea(available)
End Function

' Делегирует аналитический центр тяжести в реализацию класса для интерфейса ISectionGeometry.
Private Sub ISectionGeometry_AnalyticalCentroid(ByRef available As Boolean, ByRef X As Double, ByRef Y As Double)
    AnalyticalCentroid available, X, Y
End Sub

' Делегирует внешний контур в реализацию класса для интерфейса ISectionGeometry.
Private Sub ISectionGeometry_GetExtremePoints(ByRef pointX() As Double, ByRef pointY() As Double)
    GetExtremePoints pointX, pointY
End Sub

' Нормализует тип боковой стороны: Simple или Tapered.
Private Function NormalizeSideType(ByVal rawValue As String) As String
    Select Case LCase$(Trim$(rawValue))
        Case vbNullString, SIDE_SIMPLE
            NormalizeSideType = SIDE_SIMPLE
        Case SIDE_TAPERED
            NormalizeSideType = SIDE_TAPERED
        Case Else
            NormalizeSideType = vbNullString
    End Select
End Function

' Возвращает пользовательскую подпись нормализованного типа боковой стороны.
Private Function DisplaySideType(ByVal normalizedValue As String) As String
    If normalizedValue = SIDE_TAPERED Then
        DisplaySideType = "Tapered"
    Else
        DisplaySideType = "Simple"
    End If
End Function

' Полностью пересобирает расчетный контур: raw-ломаная, скругления, затем свойства polygon-а.
Private Sub RebuildContour()
    mPointCount = 0
    Erase mPointX
    Erase mPointY
    If Not mInitialized Then Exit Sub

    Dim raw() As TRoundedRawPoint
    Dim rawCount As Long
    BuildRawContour raw, rawCount
    BuildRoundedPath raw, rawCount, 0#, mPointX, mPointY, mPointCount
    RemoveClosingDuplicate
    RecalculatePolygonProperties
End Sub

' Собирает исходную ломаную форму до скруглений: базовые торцы B, вершины
' tapered-носов и радиусы, которые затем превращаются в касательные дуги.
Private Sub BuildRawContour(ByRef raw() As TRoundedRawPoint, ByRef rawCount As Long)
    rawCount = 0
    AddRawPoint raw, rawCount, LeftBaseX, TopY, mLeftR1
    AddRawPoint raw, rawCount, RightBaseX, TopY, mRightR1
    If mRightType = SIDE_TAPERED Then AddRawPoint raw, rawCount, RightBaseX + mRightW, mCenterY, mRightR2
    AddRawPoint raw, rawCount, RightBaseX, BottomY, mRightR1
    AddRawPoint raw, rawCount, LeftBaseX, BottomY, mLeftR1
    If mLeftType = SIDE_TAPERED Then AddRawPoint raw, rawCount, LeftBaseX - mLeftW, mCenterY, mLeftR2
End Sub

' Добавляет одну вершину исходной ломаной и связанный с ней радиус сопряжения.
Private Sub AddRawPoint(ByRef raw() As TRoundedRawPoint, ByRef rawCount As Long, _
        ByVal X As Double, ByVal Y As Double, ByVal radius As Double)
    rawCount = rawCount + 1
    If rawCount = 1 Then
        ReDim raw(1 To 8)
    ElseIf rawCount > UBound(raw) Then
        ReDim Preserve raw(1 To UBound(raw) + 8)
    End If
    raw(rawCount).X = X
    raw(rawCount).Y = Y
    raw(rawCount).Radius = radius
End Sub

' Строит дискретизированный контур по исходной ломаной с учетом внутреннего
' смещения. Один и тот же алгоритм используется для внешней границы и линий
' раскладки арматуры, чтобы их геометрия была согласованной.
Private Sub BuildRoundedPath(ByRef raw() As TRoundedRawPoint, ByVal rawCount As Long, ByVal offsetDistance As Double, _
        ByRef outX() As Double, ByRef outY() As Double, ByRef outCount As Long)
    outCount = 0
    Erase outX
    Erase outY
    If rawCount < 3 Then Exit Sub

    Dim offsetX() As Double
    Dim offsetY() As Double
    BuildOffsetVertices raw, rawCount, offsetDistance, offsetX, offsetY

    Dim i As Long
    For i = 1 To rawCount
        AppendCornerFromOffset raw, offsetX, offsetY, rawCount, PreviousIndex(i, rawCount), i, NextIndex(i, rawCount), _
            offsetDistance, outX, outY, outCount
    Next i
    RemoveOutputClosingDuplicate outX, outY, outCount
End Sub

' Получает вершины параллельной внутренней ломаной. Для каждой стороны
' сдвигается несущая прямая, затем соседние прямые пересекаются в новой вершине.
Private Sub BuildOffsetVertices(ByRef raw() As TRoundedRawPoint, ByVal rawCount As Long, ByVal offsetDistance As Double, _
        ByRef offsetX() As Double, ByRef offsetY() As Double)
    ReDim offsetX(1 To rawCount)
    ReDim offsetY(1 To rawCount)

    Dim dirX() As Double, dirY() As Double, lineX() As Double, lineY() As Double
    ReDim dirX(1 To rawCount)
    ReDim dirY(1 To rawCount)
    ReDim lineX(1 To rawCount)
    ReDim lineY(1 To rawCount)

    Dim signedArea As Double
    signedArea = RawSignedArea(raw, rawCount)

    Dim i As Long
    For i = 1 To rawCount
        Dim j As Long
        j = NextIndex(i, rawCount)
        Dim dx As Double, dy As Double, lengthValue As Double
        dx = raw(j).X - raw(i).X
        dy = raw(j).Y - raw(i).Y
        lengthValue = Sqr(dx * dx + dy * dy)
        If lengthValue <= GEOM_TOLERANCE Then lengthValue = 1#
        dirX(i) = dx / lengthValue
        dirY(i) = dy / lengthValue

        Dim nx As Double, ny As Double
        If signedArea < 0# Then
            nx = dirY(i): ny = -dirX(i)
        Else
            nx = -dirY(i): ny = dirX(i)
        End If
        lineX(i) = raw(i).X + nx * offsetDistance
        lineY(i) = raw(i).Y + ny * offsetDistance
    Next i

    For i = 1 To rawCount
        Dim p As Long
        p = PreviousIndex(i, rawCount)
        IntersectOffsetLines lineX(p), lineY(p), dirX(p), dirY(p), _
            lineX(i), lineY(i), dirX(i), dirY(i), raw(i).X, raw(i).Y, offsetX(i), offsetY(i)
    Next i
End Sub

' Возвращает знаковую площадь исходной ломаной, чтобы определить направление
' обхода и выбрать внутреннюю нормаль для offset-геометрии.
Private Function RawSignedArea(ByRef raw() As TRoundedRawPoint, ByVal rawCount As Long) As Double
    Dim result As Double
    Dim i As Long
    For i = 1 To rawCount
        Dim j As Long
        j = NextIndex(i, rawCount)
        result = result + raw(i).X * raw(j).Y - raw(j).X * raw(i).Y
    Next i
    RawSignedArea = result / 2#
End Function

' Пересекает две сдвинутые прямые. Если прямые почти параллельны, возвращает
' исходную вершину как устойчивый fallback для вырожденного локального угла.
Private Sub IntersectOffsetLines(ByVal x1 As Double, ByVal y1 As Double, ByVal dx1 As Double, ByVal dy1 As Double, _
        ByVal x2 As Double, ByVal y2 As Double, ByVal dx2 As Double, ByVal dy2 As Double, _
        ByVal fallbackX As Double, ByVal fallbackY As Double, ByRef resultX As Double, ByRef resultY As Double)
    Dim denom As Double
    denom = dx1 * dy2 - dy1 * dx2
    If Abs(denom) <= GEOM_TOLERANCE Then
        resultX = fallbackX
        resultY = fallbackY
        Exit Sub
    End If

    Dim t As Double
    t = ((x2 - x1) * dy2 - (y2 - y1) * dx2) / denom
    resultX = x1 + t * dx1
    resultY = y1 + t * dy1
End Sub

' Добавляет один скругленный угол по уже найденным вершинам offset-ломаной.
' Радиус уменьшается на величину offset, поэтому линии арматуры идут внутри
' бетонного контура на заданном расстоянии as.
Private Sub AppendCornerFromOffset(ByRef raw() As TRoundedRawPoint, ByRef offsetX() As Double, ByRef offsetY() As Double, _
        ByVal rawCount As Long, ByVal prevIndex As Long, ByVal cornerIndex As Long, ByVal nextIndexValue As Long, _
        ByVal offsetDistance As Double, ByRef outX() As Double, ByRef outY() As Double, ByRef outCount As Long)
    Dim cornerRadius As Double
    cornerRadius = GeomMax(raw(cornerIndex).Radius - offsetDistance, 0#)

    AppendRoundedCorner outX, outY, outCount, _
        offsetX(prevIndex), offsetY(prevIndex), offsetX(cornerIndex), offsetY(cornerIndex), _
        offsetX(nextIndexValue), offsetY(nextIndexValue), cornerRadius
End Sub

' Возвращает точную геометрию одного скругленного угла. Если радиус
' вырожден, начало и конец совпадают с вершиной, а hasArc=False.
Private Sub GetRoundedCornerGeometry(ByRef raw() As TRoundedRawPoint, ByRef offsetX() As Double, ByRef offsetY() As Double, _
        ByVal rawCount As Long, ByVal prevIndex As Long, ByVal cornerIndex As Long, ByVal nextIndexValue As Long, _
        ByVal offsetDistance As Double, ByRef startX As Double, ByRef startY As Double, _
        ByRef endX As Double, ByRef endY As Double, ByRef centerX As Double, ByRef centerY As Double, _
        ByRef radius As Double, ByRef sweepAngle As Double, ByRef hasArc As Boolean)
    radius = GeomMax(raw(cornerIndex).Radius - offsetDistance, 0#)
    GetCornerArcGeometry offsetX(prevIndex), offsetY(prevIndex), offsetX(cornerIndex), offsetY(cornerIndex), _
        offsetX(nextIndexValue), offsetY(nextIndexValue), radius, startX, startY, endX, endY, _
        centerX, centerY, radius, sweepAngle, hasArc
End Sub

' Рассчитывает касательные точки, центр и signed sweep для произвольного
' сопряжения. Это единый источник геометрии радиусов для контура и схемы.
Private Sub GetCornerArcGeometry(ByVal prevX As Double, ByVal prevY As Double, _
        ByVal cornerX As Double, ByVal cornerY As Double, ByVal nextX As Double, ByVal nextY As Double, _
        ByVal inputRadius As Double, ByRef startX As Double, ByRef startY As Double, _
        ByRef endX As Double, ByRef endY As Double, ByRef centerX As Double, ByRef centerY As Double, _
        ByRef radius As Double, ByRef sweepAngle As Double, ByRef hasArc As Boolean)
    Dim vPrevX As Double, vPrevY As Double, vNextX As Double, vNextY As Double
    Dim prevLength As Double, nextLength As Double
    vPrevX = prevX - cornerX: vPrevY = prevY - cornerY
    vNextX = nextX - cornerX: vNextY = nextY - cornerY
    prevLength = Sqr(vPrevX * vPrevX + vPrevY * vPrevY)
    nextLength = Sqr(vNextX * vNextX + vNextY * vNextY)
    If prevLength <= GEOM_TOLERANCE Or nextLength <= GEOM_TOLERANCE Or inputRadius <= GEOM_TOLERANCE Then
        startX = cornerX: startY = cornerY
        endX = cornerX: endY = cornerY
        radius = 0#: sweepAngle = 0#: hasArc = False
        Exit Sub
    End If

    vPrevX = vPrevX / prevLength: vPrevY = vPrevY / prevLength
    vNextX = vNextX / nextLength: vNextY = vNextY / nextLength

    Dim dotValue As Double
    dotValue = vPrevX * vNextX + vPrevY * vNextY
    If dotValue > 1# Then dotValue = 1#
    If dotValue < -1# Then dotValue = -1#

    Dim angleValue As Double
    angleValue = Atn2Local(Sqr(GeomMax(1# - dotValue * dotValue, 0#)), dotValue)
    If angleValue <= GEOM_TOLERANCE Or Abs(GEOM_PI - angleValue) <= GEOM_TOLERANCE Then
        startX = cornerX: startY = cornerY
        endX = cornerX: endY = cornerY
        radius = 0#: sweepAngle = 0#: hasArc = False
        Exit Sub
    End If

    Dim tangentDistance As Double
    tangentDistance = inputRadius / Tan(angleValue / 2#)
    Dim maxDistance As Double
    maxDistance = GeomMin(prevLength, nextLength) - GEOM_TOLERANCE
    If tangentDistance > maxDistance Then
        tangentDistance = maxDistance
        inputRadius = tangentDistance * Tan(angleValue / 2#)
    End If
    If tangentDistance <= GEOM_TOLERANCE Or inputRadius <= GEOM_TOLERANCE Then
        startX = cornerX: startY = cornerY
        endX = cornerX: endY = cornerY
        radius = 0#: sweepAngle = 0#: hasArc = False
        Exit Sub
    End If

    startX = cornerX + vPrevX * tangentDistance
    startY = cornerY + vPrevY * tangentDistance
    endX = cornerX + vNextX * tangentDistance
    endY = cornerY + vNextY * tangentDistance

    Dim bisectorX As Double, bisectorY As Double, bisectorLength As Double
    bisectorX = vPrevX + vNextX
    bisectorY = vPrevY + vNextY
    bisectorLength = Sqr(bisectorX * bisectorX + bisectorY * bisectorY)
    If bisectorLength <= GEOM_TOLERANCE Then
        startX = cornerX: startY = cornerY
        endX = cornerX: endY = cornerY
        radius = 0#: sweepAngle = 0#: hasArc = False
        Exit Sub
    End If
    bisectorX = bisectorX / bisectorLength
    bisectorY = bisectorY / bisectorLength

    Dim centerDistance As Double
    centerDistance = inputRadius / Sin(angleValue / 2#)
    centerX = cornerX + bisectorX * centerDistance
    centerY = cornerY + bisectorY * centerDistance

    Dim startAngle As Double, endAngle As Double
    startAngle = Atn2Local(startY - centerY, startX - centerX)
    endAngle = Atn2Local(endY - centerY, endX - centerX)

    Dim crossValue As Double
    crossValue = vPrevX * vNextY - vPrevY * vNextX
    If crossValue > 0# Then
        If endAngle > startAngle Then endAngle = endAngle - 2# * GEOM_PI
    Else
        If endAngle < startAngle Then endAngle = endAngle + 2# * GEOM_PI
    End If

    radius = inputRadius
    sweepAngle = endAngle - startAngle
    hasArc = True
End Sub

' Добавляет прямой или дуговой сегмент в выходной массив визуального контура.
Private Sub AddContourSegment(ByRef segmentType() As String, ByRef startX() As Double, ByRef startY() As Double, _
        ByRef endX() As Double, ByRef endY() As Double, ByRef centerX() As Double, ByRef centerY() As Double, _
        ByRef radiusValue() As Double, ByRef sweepAngle() As Double, ByRef segmentCount As Long, _
        ByVal kind As String, ByVal sx As Double, ByVal sy As Double, ByVal ex As Double, ByVal ey As Double, _
        ByVal cx As Double, ByVal cy As Double, ByVal radius As Double, ByVal sweep As Double)
    If Abs(sx - ex) <= GEOM_TOLERANCE And Abs(sy - ey) <= GEOM_TOLERANCE Then Exit Sub
    segmentCount = segmentCount + 1
    If segmentCount = 1 Then
        ReDim segmentType(1 To 16)
        ReDim startX(1 To 16)
        ReDim startY(1 To 16)
        ReDim endX(1 To 16)
        ReDim endY(1 To 16)
        ReDim centerX(1 To 16)
        ReDim centerY(1 To 16)
        ReDim radiusValue(1 To 16)
        ReDim sweepAngle(1 To 16)
    ElseIf segmentCount > UBound(segmentType) Then
        ReDim Preserve segmentType(1 To UBound(segmentType) + 16)
        ReDim Preserve startX(1 To UBound(startX) + 16)
        ReDim Preserve startY(1 To UBound(startY) + 16)
        ReDim Preserve endX(1 To UBound(endX) + 16)
        ReDim Preserve endY(1 To UBound(endY) + 16)
        ReDim Preserve centerX(1 To UBound(centerX) + 16)
        ReDim Preserve centerY(1 To UBound(centerY) + 16)
        ReDim Preserve radiusValue(1 To UBound(radiusValue) + 16)
        ReDim Preserve sweepAngle(1 To UBound(sweepAngle) + 16)
    End If
    segmentType(segmentCount) = kind
    startX(segmentCount) = sx
    startY(segmentCount) = sy
    endX(segmentCount) = ex
    endY(segmentCount) = ey
    centerX(segmentCount) = cx
    centerY(segmentCount) = cy
    radiusValue(segmentCount) = radius
    sweepAngle(segmentCount) = sweep
End Sub

' Рассчитывает касательные точки произвольного угла и добавляет дугу заданного
' радиуса. Здесь нет фиксированного угла: наклон tapered-части получается из
' фактических координат соседних отрезков.
Private Sub AppendRoundedCorner(ByRef outX() As Double, ByRef outY() As Double, ByRef outCount As Long, _
        ByVal prevX As Double, ByVal prevY As Double, ByVal cornerX As Double, ByVal cornerY As Double, _
        ByVal nextX As Double, ByVal nextY As Double, ByVal radius As Double)
    Dim startX As Double, startY As Double, endX As Double, endY As Double
    Dim centerX As Double, centerY As Double, sweepAngle As Double, hasArc As Boolean
    GetCornerArcGeometry prevX, prevY, cornerX, cornerY, nextX, nextY, radius, _
        startX, startY, endX, endY, centerX, centerY, radius, sweepAngle, hasArc
    If Not hasArc Then
        AddOutputPoint outX, outY, outCount, cornerX, cornerY
        Exit Sub
    End If

    AddOutputPoint outX, outY, outCount, startX, startY
    AddArcPointsToOutput outX, outY, outCount, centerX, centerY, radius, _
        Atn2Local(startY - centerY, startX - centerX), _
        Atn2Local(startY - centerY, startX - centerX) + sweepAngle
End Sub

' Дискретизирует дугу в общий массив точек контура или линии раскладки.
Private Sub AddArcPointsToOutput(ByRef outX() As Double, ByRef outY() As Double, ByRef outCount As Long, _
        ByVal centerX As Double, ByVal centerY As Double, ByVal radius As Double, _
        ByVal angleStart As Double, ByVal angleEnd As Double)
    If radius <= GEOM_TOLERANCE Then Exit Sub
    Dim i As Long
    For i = 1 To ARC_SEGMENTS
        Dim t As Double
        t = CDbl(i) / CDbl(ARC_SEGMENTS)
        AddOutputPoint outX, outY, outCount, centerX + radius * Cos(angleStart + (angleEnd - angleStart) * t), _
            centerY + radius * Sin(angleStart + (angleEnd - angleStart) * t)
    Next i
End Sub

' Возвращает две касательные точки конкретного скругления. Горизонтальные линии
' арматуры используют эти точки как начало и конец прямого участка B.
Private Sub CornerTangentPoints(ByRef raw() As TRoundedRawPoint, ByRef offsetX() As Double, ByRef offsetY() As Double, _
        ByVal prevIndex As Long, ByVal cornerIndex As Long, ByVal nextIndexValue As Long, ByVal offsetDistance As Double, _
        ByRef startX As Double, ByRef startY As Double, ByRef endX As Double, ByRef endY As Double)
    Dim radius As Double
    radius = GeomMax(raw(cornerIndex).Radius - offsetDistance, 0#)
    Dim cornerX As Double, cornerY As Double
    cornerX = offsetX(cornerIndex): cornerY = offsetY(cornerIndex)

    Dim vPrevX As Double, vPrevY As Double, vNextX As Double, vNextY As Double
    Dim prevLength As Double, nextLength As Double
    vPrevX = offsetX(prevIndex) - cornerX: vPrevY = offsetY(prevIndex) - cornerY
    vNextX = offsetX(nextIndexValue) - cornerX: vNextY = offsetY(nextIndexValue) - cornerY
    prevLength = Sqr(vPrevX * vPrevX + vPrevY * vPrevY)
    nextLength = Sqr(vNextX * vNextX + vNextY * vNextY)
    If prevLength <= GEOM_TOLERANCE Or nextLength <= GEOM_TOLERANCE Or radius <= GEOM_TOLERANCE Then
        startX = cornerX: startY = cornerY
        endX = cornerX: endY = cornerY
        Exit Sub
    End If

    vPrevX = vPrevX / prevLength: vPrevY = vPrevY / prevLength
    vNextX = vNextX / nextLength: vNextY = vNextY / nextLength
    Dim dotValue As Double
    dotValue = vPrevX * vNextX + vPrevY * vNextY
    If dotValue > 1# Then dotValue = 1#
    If dotValue < -1# Then dotValue = -1#

    Dim angleValue As Double
    angleValue = Atn2Local(Sqr(GeomMax(1# - dotValue * dotValue, 0#)), dotValue)
    Dim tangentDistance As Double
    If angleValue <= GEOM_TOLERANCE Or Abs(GEOM_PI - angleValue) <= GEOM_TOLERANCE Then
        tangentDistance = 0#
    Else
        tangentDistance = radius / Tan(angleValue / 2#)
    End If
    tangentDistance = GeomMin(tangentDistance, GeomMin(prevLength, nextLength) - GEOM_TOLERANCE)
    If tangentDistance <= GEOM_TOLERANCE Then
        startX = cornerX: startY = cornerY
        endX = cornerX: endY = cornerY
    Else
        startX = cornerX + vPrevX * tangentDistance
        startY = cornerY + vPrevY * tangentDistance
        endX = cornerX + vNextX * tangentDistance
        endY = cornerY + vNextY * tangentDistance
    End If
End Sub

' Возвращает индекс правой нижней базовой вершины в исходной ломаной: при
' Tapered перед ней появляется дополнительная носовая вершина.
Private Function RightBottomRawIndex() As Long
    If mRightType = SIDE_TAPERED Then
        RightBottomRawIndex = 4
    Else
        RightBottomRawIndex = 3
    End If
End Function

' Циклический предыдущий индекс для замкнутого контура.
Private Function PreviousIndex(ByVal index As Long, ByVal count As Long) As Long
    If index <= 1 Then
        PreviousIndex = count
    Else
        PreviousIndex = index - 1
    End If
End Function

' Циклический следующий индекс для замкнутого контура.
Private Function NextIndex(ByVal index As Long, ByVal count As Long) As Long
    If index >= count Then
        NextIndex = 1
    Else
        NextIndex = index + 1
    End If
End Function

' Добавляет точку в выходной массив и пропускает дубль предыдущей точки, чтобы
' контур не получал нулевых ребер на стыках прямых и дуг.
Private Sub AddOutputPoint(ByRef outX() As Double, ByRef outY() As Double, ByRef outCount As Long, ByVal X As Double, ByVal Y As Double)
    If outCount > 0 Then
        If Abs(outX(outCount) - X) <= GEOM_TOLERANCE And Abs(outY(outCount) - Y) <= GEOM_TOLERANCE Then Exit Sub
    End If
    outCount = outCount + 1
    If outCount = 1 Then
        ReDim outX(1 To 16)
        ReDim outY(1 To 16)
    ElseIf outCount > UBound(outX) Then
        ReDim Preserve outX(1 To UBound(outX) + 16)
        ReDim Preserve outY(1 To UBound(outY) + 16)
    End If
    outX(outCount) = X
    outY(outCount) = Y
End Sub

' Убирает последнюю точку, если она совпала с первой после замыкания обхода.
Private Sub RemoveOutputClosingDuplicate(ByRef outX() As Double, ByRef outY() As Double, ByRef outCount As Long)
    If outCount < 2 Then Exit Sub
    If Abs(outX(1) - outX(outCount)) <= GEOM_TOLERANCE And Abs(outY(1) - outY(outCount)) <= GEOM_TOLERANCE Then
        outCount = outCount - 1
    End If
    If outCount > 0 Then
        ReDim Preserve outX(1 To outCount)
        ReDim Preserve outY(1 To outCount)
    End If
End Sub

' Локальная версия atan2 для VBA: возвращает угол в радианах с учетом квадранта.
Private Function Atn2Local(ByVal Y As Double, ByVal X As Double) As Double
    If Abs(X) <= GEOM_TOLERANCE Then
        If Y >= 0# Then
            Atn2Local = GEOM_PI / 2#
        Else
            Atn2Local = -GEOM_PI / 2#
        End If
    Else
        Atn2Local = Atn(Y / X)
        If X < 0# Then
            If Y >= 0# Then
                Atn2Local = Atn2Local + GEOM_PI
            Else
                Atn2Local = Atn2Local - GEOM_PI
            End If
        End If
    End If
End Function

' Убирает последнюю точку, если она повторяет первую после замыкания контура.
Private Sub RemoveClosingDuplicate()
    If mPointCount < 2 Then Exit Sub
    If Abs(mPointX(1) - mPointX(mPointCount)) <= GEOM_TOLERANCE And _
            Abs(mPointY(1) - mPointY(mPointCount)) <= GEOM_TOLERANCE Then
        mPointCount = mPointCount - 1
    End If
    ReDim Preserve mPointX(1 To mPointCount)
    ReDim Preserve mPointY(1 To mPointCount)
End Sub

' Пересчитывает площадь, центр тяжести и bounds по текущему polygon-контуру.
Private Sub RecalculatePolygonProperties()
    mArea = 0#
    mCentroidX = 0#
    mCentroidY = 0#
    If mPointCount < 3 Then Exit Sub

    mMinX = mPointX(1): mMaxX = mPointX(1)
    mMinY = mPointY(1): mMaxY = mPointY(1)
    Dim twiceArea As Double
    Dim cxSum As Double
    Dim cySum As Double
    Dim i As Long
    For i = 1 To mPointCount
        Dim j As Long
        If i = mPointCount Then
            j = 1
        Else
            j = i + 1
        End If
        Dim crossValue As Double
        crossValue = mPointX(i) * mPointY(j) - mPointX(j) * mPointY(i)
        twiceArea = twiceArea + crossValue
        cxSum = cxSum + (mPointX(i) + mPointX(j)) * crossValue
        cySum = cySum + (mPointY(i) + mPointY(j)) * crossValue
        If mPointX(i) < mMinX Then mMinX = mPointX(i)
        If mPointX(i) > mMaxX Then mMaxX = mPointX(i)
        If mPointY(i) < mMinY Then mMinY = mPointY(i)
        If mPointY(i) > mMaxY Then mMaxY = mPointY(i)
    Next i

    If Abs(twiceArea) <= GEOM_TOLERANCE Then Exit Sub
    mArea = Abs(twiceArea) / 2#
    mCentroidX = cxSum / (3# * twiceArea)
    mCentroidY = cySum / (3# * twiceArea)
End Sub

' Выполняет ray-casting проверку попадания точки внутрь расчетного polygon-контура.
Private Function PointInContour(ByVal X As Double, ByVal Y As Double) As Boolean
    Dim inside As Boolean
    Dim i As Long
    Dim j As Long
    j = mPointCount
    For i = 1 To mPointCount
        If ((mPointY(i) > Y) <> (mPointY(j) > Y)) Then
            Dim xIntersection As Double
            xIntersection = (mPointX(j) - mPointX(i)) * (Y - mPointY(i)) / (mPointY(j) - mPointY(i)) + mPointX(i)
            If X < xIntersection + GEOM_TOLERANCE Then inside = Not inside
        End If
        j = i
    Next i
    PointInContour = inside
End Function
