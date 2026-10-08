VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGeometryHollowRectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Пустотелый скругленный прямоугольник
' ==========================================================================
' Класс описывает параметрическую бетонную область HollowRectangle: внешний
' скругленный прямоугольник минус внутреннее отверстие Opening. Расчетное ядро
' видит обычную ISectionGeometry: ContainsPoint, габариты, площадь и центр.
'
' Внешний и внутренний контуры построены на уже проверенной математике
' CGeometryRoundedRectangle. Этот класс только задает ответственность
' "наружный контур - отверстие" и предоставляет дополнительные методы для
' арматуры/аннотаций, которым нужны отдельные петли контура.

Implements ISectionGeometry

Private mOuter As CGeometryRoundedRectangle ' Наружная бетонная граница.
Private mOpening As CGeometryRoundedRectangle ' Граница пустоты; внутри нее бетона нет.
Private mOuterRotated As CGeometryRoundedRectangle ' Наружная граница в системе u=Y, v=-X для H/B-раскладки.
Private mOpeningRotated As CGeometryRoundedRectangle ' Opening в повернутой системе; используется только для линий арматуры.
Private mB As Double ' Полная наружная ширина B.
Private mH As Double ' Полная наружная высота H.
Private mR As Double ' Радиус наружных углов.
Private mOpeningB As Double ' Ширина внутреннего отверстия B_o.
Private mOpeningH As Double ' Высота внутреннего отверстия H_o.
Private mOpeningR As Double ' Радиус внутренних углов R_o.
Private mOpeningOffsetX As Double ' Смещение центра отверстия относительно центра наружного контура по X.
Private mOpeningOffsetY As Double ' Смещение центра отверстия относительно центра наружного контура по Y.
Private mArea As Double ' Площадь бетонного кольца по polygon-контурам наружной формы и отверстия.
Private mCentroidX As Double ' Центр тяжести бетонного кольца.
Private mCentroidY As Double
Private mInitialized As Boolean
Private mValidationReady As Boolean ' Сброс при каждом Initialize, включая ошибочные параметры.
Private mGeometryValid As Boolean ' Наружный контур и отверстие уже проверены совместно.
Private mValidationMessage As String ' Причина отказа текущей геометрии, а не предыдущего сечения.

' Инициализирует внешний контур и Opening в одной системе координат.
' OffsetX/OffsetY задают смещение центра отверстия относительно центра
' наружного прямоугольника; положительные значения идут по положительным X/Y.
Public Sub Initialize(ByVal widthMm As Double, ByVal heightMm As Double, ByVal radiusMm As Double, _
        ByVal openingWidthMm As Double, ByVal openingHeightMm As Double, ByVal openingRadiusMm As Double, _
        Optional ByVal openingOffsetXMm As Double = 0#, Optional ByVal openingOffsetYMm As Double = 0#)
    mValidationReady = False
    mGeometryValid = False
    mValidationMessage = vbNullString
    mArea = 0#
    mCentroidX = 0#
    mCentroidY = 0#
    mB = widthMm
    mH = heightMm
    mR = radiusMm
    mOpeningB = openingWidthMm
    mOpeningH = openingHeightMm
    mOpeningR = openingRadiusMm
    mOpeningOffsetX = openingOffsetXMm
    mOpeningOffsetY = openingOffsetYMm

    Set mOuter = New CGeometryRoundedRectangle
    mOuter.InitializeSides mB, mH, "Simple", "Simple", 0#, 0#, mR, mR, 0#, 0#, 0#, 0#

    Set mOpening = New CGeometryRoundedRectangle
    mOpening.InitializeSides mOpeningB, mOpeningH, "Simple", "Simple", 0#, 0#, _
        mOpeningR, mOpeningR, 0#, 0#, mOpeningOffsetX, mOpeningOffsetY

    ' В HollowRectangle смысл раскладки повернут относительно обычного RoundedRectangle:
    ' H - прямые боковые линии, B - верхняя/нижняя траектория с учетом радиусов.
    ' Чтобы не дублировать расчет касательных точек, строим ту же форму в системе
    ' u=Y, v=-X и затем возвращаем точки в исходные координаты.
    Set mOuterRotated = New CGeometryRoundedRectangle
    mOuterRotated.InitializeSides mH, mB, "Simple", "Simple", 0#, 0#, mR, mR, 0#, 0#, 0#, 0#

    Set mOpeningRotated = New CGeometryRoundedRectangle
    mOpeningRotated.InitializeSides mOpeningH, mOpeningB, "Simple", "Simple", 0#, 0#, _
        mOpeningR, mOpeningR, 0#, 0#, mOpeningOffsetY, -mOpeningOffsetX

    mInitialized = True
    RecalculateAnalyticalProperties
    Dim message As String
    mGeometryValid = IsValid(message)
End Sub

Public Property Get B() As Double
    B = mB
End Property

Public Property Get H() As Double
    H = mH
End Property

Public Property Get R() As Double
    R = mR
End Property

Public Property Get OpeningB() As Double
    OpeningB = mOpeningB
End Property

Public Property Get OpeningH() As Double
    OpeningH = mOpeningH
End Property

Public Property Get OpeningR() As Double
    OpeningR = mOpeningR
End Property

Public Property Get OpeningOffsetX() As Double
    OpeningOffsetX = mOpeningOffsetX
End Property

Public Property Get OpeningOffsetY() As Double
    OpeningOffsetY = mOpeningOffsetY
End Property

Public Property Get OpeningMinX() As Double
    OpeningMinX = mOpening.MinX
End Property

Public Property Get OpeningMaxX() As Double
    OpeningMaxX = mOpening.MaxX
End Property

Public Property Get OpeningMinY() As Double
    OpeningMinY = mOpening.MinY
End Property

Public Property Get OpeningMaxY() As Double
    OpeningMaxY = mOpening.MaxY
End Property

Public Property Get MinX() As Double
    MinX = mOuter.MinX
End Property

Public Property Get MaxX() As Double
    MaxX = mOuter.MaxX
End Property

Public Property Get MinY() As Double
    MinY = mOuter.MinY
End Property

Public Property Get MaxY() As Double
    MaxY = mOuter.MaxY
End Property

' Возвращает точный контур одной петли: наружной или внутренней.
' openingLoop=False - внешний контур, openingLoop=True - Opening.
Public Sub GetContourSegments(ByVal openingLoop As Boolean, ByRef segmentType() As String, _
        ByRef startX() As Double, ByRef startY() As Double, ByRef endX() As Double, ByRef endY() As Double, _
        ByRef centerX() As Double, ByRef centerY() As Double, ByRef radiusValue() As Double, _
        ByRef sweepAngle() As Double, ByRef segmentCount As Long)
    If openingLoop Then
        mOpening.GetContourSegments segmentType, startX, startY, endX, endY, centerX, centerY, radiusValue, sweepAngle, segmentCount
    Else
        mOuter.GetContourSegments segmentType, startX, startY, endX, endY, centerX, centerY, radiusValue, sweepAngle, segmentCount
    End If
End Sub

' Возвращает горизонтальную линию раскладки внешней или внутренней грани.
' Для Opening знак as инвертируется: линия уходит от пустоты в бетон.
Public Sub GetHorizontalRebarLine(ByVal openingFace As Boolean, ByVal isTop As Boolean, ByVal axisDistance As Double, _
        ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double, _
        ByRef normalX As Double, ByRef normalY As Double)
    If openingFace Then
        mOpening.GetHorizontalRebarLine isTop, -axisDistance, x1, y1, x2, y2, normalX, normalY
        normalX = -normalX
        normalY = -normalY
    Else
        mOuter.GetHorizontalRebarLine isTop, axisDistance, x1, y1, x2, y2, normalX, normalY
    End If
End Sub

' Возвращает боковую траекторию внешней или внутренней грани.
' Для внутренней грани используется отрицательный offset, потому что бетон
' находится снаружи Opening, а не внутри пустоты.
Public Sub GetSideRebarPath(ByVal openingFace As Boolean, ByVal isLeft As Boolean, ByVal axisDistance As Double, _
        ByRef pointX() As Double, ByRef pointY() As Double)
    If openingFace Then
        mOpening.GetSideRebarPath isLeft, -axisDistance, pointX, pointY
    Else
        mOuter.GetSideRebarPath isLeft, axisDistance, pointX, pointY
    End If
End Sub

' Возвращает прямую вертикальную линию раскладки для H-граней HollowRectangle.
' Внутри используется повернутая RoundedRectangle-геометрия: вертикальная H-грань
' исходного сечения соответствует горизонтальной линии в системе u=Y, v=-X.
Public Sub GetVerticalRebarLine(ByVal openingFace As Boolean, ByVal isLeft As Boolean, ByVal axisDistance As Double, _
        ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double, _
        ByRef normalX As Double, ByRef normalY As Double)
    Dim target As CGeometryRoundedRectangle
    If openingFace Then
        Set target = mOpeningRotated
    Else
        Set target = mOuterRotated
    End If

    Dim u1 As Double, v1 As Double, u2 As Double, v2 As Double
    Dim normalU As Double, normalV As Double
    target.GetHorizontalRebarLine isLeft, IIf(openingFace, -axisDistance, axisDistance), _
        u1, v1, u2, v2, normalU, normalV

    MapFromRotated u1, v1, x1, y1
    MapFromRotated u2, v2, x2, y2
    MapVectorFromRotated normalU, normalV, normalX, normalY

    If openingFace Then
        normalX = -normalX
        normalY = -normalY
    End If
End Sub

' Возвращает криволинейную верхнюю или нижнюю траекторию для B-граней.
' Это боковая траектория повернутой RoundedRectangle-геометрии, отображенная
' обратно в исходные координаты HollowRectangle.
Public Sub GetHorizontalRebarPath(ByVal openingFace As Boolean, ByVal isTop As Boolean, ByVal axisDistance As Double, _
        ByRef pointX() As Double, ByRef pointY() As Double)
    Dim target As CGeometryRoundedRectangle
    If openingFace Then
        Set target = mOpeningRotated
    Else
        Set target = mOuterRotated
    End If

    Dim pointU() As Double, pointV() As Double
    target.GetSideRebarPath Not isTop, IIf(openingFace, -axisDistance, axisDistance), pointU, pointV

    Dim i As Long
    ReDim pointX(LBound(pointU) To UBound(pointU))
    ReDim pointY(LBound(pointU) To UBound(pointU))
    For i = LBound(pointU) To UBound(pointU)
        MapFromRotated pointU(i), pointV(i), pointX(i), pointY(i)
    Next i
End Sub

' Проверяет попадание в бетон: точка должна быть внутри внешнего контура и
' одновременно вне внутреннего отверстия.
Public Function ContainsPoint(ByVal X As Double, ByVal Y As Double) As Boolean
    Dim message As String
    If Not IsValid(message) Then Exit Function
    ContainsPoint = (mOuter.ContainsPoint(X, Y) And Not mOpening.ContainsPoint(X, Y))
End Function

' Дублирует семантику ContainsPoint для общего интерфейса геометрий проекта:
' точка считается внутри, только если она попадает в бетонное кольцо, а не в пустоту.
Public Function IsPointInside(ByVal X As Double, ByVal Y As Double) As Boolean
    IsPointInside = ContainsPoint(X, Y)
End Function

' Возвращает сохраненную совместную проверку контуров. Геометрия меняется
' только через Initialize, поэтому массовые запросы точки не повторяют ее.
Public Function IsValid(ByRef message As String) As Boolean
    If Not mValidationReady Then
        mGeometryValid = ValidateCurrentGeometry(mValidationMessage)
        mValidationReady = True
    End If
    message = mValidationMessage
    IsValid = mGeometryValid
End Function

' Проверяет размеры, радиусы, размещение отверстия и площадь бетонного кольца.
' Вызывается один раз для каждого нового набора параметров геометрии.
Private Function ValidateCurrentGeometry(ByRef message As String) As Boolean
    message = vbNullString
    If Not mInitialized Then
        message = "Геометрия пустотелого прямоугольника не инициализирована."
        Exit Function
    End If

    If mB <= 0# Or mH <= 0# Then
        message = "Наружные размеры B и H пустотелого прямоугольника должны быть больше нуля."
        Exit Function
    End If
    If mOpeningB <= 0# Or mOpeningH <= 0# Then
        message = "Размеры отверстия B_o и H_o должны быть больше нуля."
        Exit Function
    End If
    If mR < 0# Or mOpeningR < 0# Then
        message = "Радиусы R и R_o не должны быть отрицательными."
        Exit Function
    End If

    If Not mOuter.IsValid(message) Then Exit Function
    If Not mOpening.IsValid(message) Then Exit Function

    If Not OpeningFitsInsideOuter() Then
        message = "Opening должен целиком находиться внутри наружного контура HollowRectangle с учетом InnerOffsetX/InnerOffsetY."
        Exit Function
    End If
    If mArea <= GEOM_TOLERANCE Then
        message = "Площадь бетонной части HollowRectangle должна быть больше нуля."
        Exit Function
    End If

    ValidateCurrentGeometry = True
End Function

' Возвращает площадь бетонного кольца по подготовленным polygon-контурам,
' если геометрия прошла валидацию. Дуги имеют ту же дискретизацию, что у
' RoundedRectangle; результат не является площадью точных круговых дуг.
Public Function AnalyticalArea(ByRef available As Boolean) As Double
    Dim message As String
    available = IsValid(message)
    If available Then AnalyticalArea = mArea
End Function

' Возвращает центр тяжести бетонной части по тем же polygon-контурам. При смещенном отверстии
' центр смещается в сторону, противоположную отверстию; этот результат нужен
' для согласованной схемы, аннотаций и контрольных свойств сечения.
Public Sub AnalyticalCentroid(ByRef available As Boolean, ByRef X As Double, ByRef Y As Double)
    Dim message As String
    available = IsValid(message)
    If Not available Then Exit Sub
    X = mCentroidX
    Y = mCentroidY
End Sub

' Для построения сетки наружный контур задает расчетные габариты и
' дискретизируется так же, как обычный RoundedRectangle.
Public Sub GetExtremePoints(ByRef pointX() As Double, ByRef pointY() As Double)
    mOuter.GetExtremePoints pointX, pointY
End Sub

' ----------
' Реализация интерфейса ISectionGeometry.
' Эти методы реализуют общий интерфейс через публичный API класса, чтобы
' HollowRectangle мог использоваться тем же pipeline сетки и расчета, что и
' остальные параметрические типы сечений.
' ----------
' Передает наружную границу и отверстие как собственные геометрические
' данные с явными ролями. Необязательные аннотации их не создают и не меняют.
Public Sub BuildContours(ByVal contours As CSectionContours)
    If contours Is Nothing Then Err.Raise vbObjectError + 5283, "CGeometryHollowRectangle", "Контейнер контуров сечения не передан."
    Dim message As String
    If Not IsValid(message) Then Err.Raise vbObjectError + 5283, "CGeometryHollowRectangle", message
    contours.Clear
    AddMaterialContour contours, False, "OUTER", "Outer", "наружного контура"
    AddMaterialContour contours, True, "OPENING", "Opening", "внутреннего контура Opening"
End Sub

' Сохраняет одну точную границу с явно заданными ролью и ключом.
' Потребители получают классификацию отверстия, а не выводят ее из имени.
Private Sub AddMaterialContour(ByVal contours As CSectionContours, ByVal openingLoop As Boolean, _
        ByVal loopName As String, ByVal loopRole As String, ByVal commentPart As String)
    Dim segmentType() As String, startX() As Double, startY() As Double
    Dim endX() As Double, endY() As Double, centerX() As Double, centerY() As Double
    Dim radiusValue() As Double, sweepAngle() As Double, segmentCount As Long, i As Long
    GetContourSegments openingLoop, segmentType, startX, startY, endX, endY, centerX, centerY, radiusValue, sweepAngle, segmentCount
    For i = 1 To segmentCount
        If StrComp(segmentType(i), "ARC", vbTextCompare) = 0 Then
            contours.AddContourArc "CONTOUR_" & loopName & "_ARC_" & CStr(i), _
                startX(i), startY(i), endX(i), endY(i), centerX(i), centerY(i), radiusValue(i), sweepAngle(i), _
                "Дуга " & commentPart & " HollowRectangle.", loopName, loopRole
        Else
            contours.AddContourLine "CONTOUR_" & loopName & "_LINE_" & CStr(i), _
                startX(i), startY(i), endX(i), endY(i), _
                "Прямой участок " & commentPart & " HollowRectangle.", loopName, loopRole
        End If
    Next i
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

' Передает интерфейсный запрос попадания точки в бетонную часть HollowRectangle.
Private Function ISectionGeometry_ContainsPoint(ByVal X As Double, ByVal Y As Double) As Boolean
    ISectionGeometry_ContainsPoint = ContainsPoint(X, Y)
End Function

' Проверяет попадание точки в бетон через альтернативный метод ISectionGeometry;
' внутренний проем, как и во внешнем ContainsPoint, исключается из области.
Private Function ISectionGeometry_IsPointInside(ByVal X As Double, ByVal Y As Double) As Boolean
    ISectionGeometry_IsPointInside = IsPointInside(X, Y)
End Function

' Передает интерфейсную валидацию в основной метод IsValid с теми же сообщениями.
Private Function ISectionGeometry_IsValid(ByRef message As String) As Boolean
    ISectionGeometry_IsValid = IsValid(message)
End Function

' Отдает аналитическую площадь через интерфейс ISectionGeometry.
Private Function ISectionGeometry_AnalyticalArea(ByRef available As Boolean) As Double
    ISectionGeometry_AnalyticalArea = AnalyticalArea(available)
End Function

' Отдает аналитический центр тяжести через интерфейс ISectionGeometry.
Private Sub ISectionGeometry_AnalyticalCentroid(ByRef available As Boolean, ByRef X As Double, ByRef Y As Double)
    AnalyticalCentroid available, X, Y
End Sub

' Отдает точки внешнего контура через интерфейс. Для сетки и габаритов берется
' именно наружная оболочка, а отверстие учитывается проверкой ContainsPoint.
Private Sub ISectionGeometry_GetExtremePoints(ByRef pointX() As Double, ByRef pointY() As Double)
    GetExtremePoints pointX, pointY
End Sub

' Передает классифицированные контуры через обязательный интерфейс формы.
Private Sub ISectionGeometry_BuildContours(ByVal contours As CSectionContours)
    BuildContours contours
End Sub

' Пересчитывает площадь и центр тяжести как разность наружного контура и
' Opening. Для смещенного отверстия центр бетонного кольца закономерно
' смещается в противоположную сторону.
Private Sub RecalculateAnalyticalProperties()
    Dim outerAvailable As Boolean
    Dim openingAvailable As Boolean
    Dim outerArea As Double
    Dim openingArea As Double
    Dim outerX As Double, outerY As Double
    Dim openingX As Double, openingY As Double

    outerArea = mOuter.AnalyticalArea(outerAvailable)
    openingArea = mOpening.AnalyticalArea(openingAvailable)
    mOuter.AnalyticalCentroid outerAvailable, outerX, outerY
    mOpening.AnalyticalCentroid openingAvailable, openingX, openingY

    If Not outerAvailable Or Not openingAvailable Then Exit Sub
    mArea = outerArea - openingArea
    If mArea <= GEOM_TOLERANCE Then Exit Sub

    mCentroidX = (outerArea * outerX - openingArea * openingX) / mArea
    mCentroidY = (outerArea * outerY - openingArea * openingY) / mArea
End Sub

' Opening считается допустимым, если все его дискретные точки лежат внутри
' наружного контура с небольшим допуском. Этого достаточно для axis-aligned
' скругленных прямоугольников и одновременно работает при смещении отверстия.
Private Function OpeningFitsInsideOuter() As Boolean
    Dim px() As Double
    Dim py() As Double
    mOpening.GetExtremePoints px, py

    Dim i As Long
    For i = LBound(px) To UBound(px)
        If Not mOuter.ContainsPoint(px(i), py(i)) Then Exit Function
    Next i
    OpeningFitsInsideOuter = True
End Function

' Преобразует точку из повернутой системы u=Y, v=-X обратно в исходные X/Y.
Private Sub MapFromRotated(ByVal uValue As Double, ByVal vValue As Double, _
        ByRef xValue As Double, ByRef yValue As Double)
    xValue = -vValue
    yValue = uValue
End Sub

' Преобразует вектор из повернутой системы u=Y, v=-X обратно в исходные X/Y.
Private Sub MapVectorFromRotated(ByVal uValue As Double, ByVal vValue As Double, _
        ByRef xValue As Double, ByRef yValue As Double)
    xValue = -vValue
    yValue = uValue
End Sub
