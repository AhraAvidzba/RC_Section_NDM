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
' Класс задает параметры и классифицированные контуры HollowRectangle:
' наружный скругленный прямоугольник и внутреннее отверстие Opening.
' Принадлежность точек и свойства бетонной области вычисляет общее ядро.
'
' Внешний и внутренний контуры построены на уже проверенной математике
' CGeometryRoundedRectangle. Этот класс только задает ответственность
' "наружный контур - отверстие" и предоставляет дополнительные методы для
' арматуры/аннотаций, которым нужны наружная и внутренняя границы.

Implements ISectionShape

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
    OpeningMinX = mOpeningOffsetX - mOpeningB / 2#
End Property

Public Property Get OpeningMaxX() As Double
    OpeningMaxX = mOpeningOffsetX + mOpeningB / 2#
End Property

Public Property Get OpeningMinY() As Double
    OpeningMinY = mOpeningOffsetY - mOpeningH / 2#
End Property

Public Property Get OpeningMaxY() As Double
    OpeningMaxY = mOpeningOffsetY + mOpeningH / 2#
End Property

Public Property Get MinX() As Double
    MinX = -mB / 2#
End Property

Public Property Get MaxX() As Double
    MaxX = mB / 2#
End Property

Public Property Get MinY() As Double
    MinY = -mH / 2#
End Property

Public Property Get MaxY() As Double
    MaxY = mH / 2#
End Property

' Возвращает точные сегменты наружной границы или отверстия.
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

' Проверяет размеры, радиусы и размещение отверстия внутри наружной границы.
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

    ValidateCurrentGeometry = True
End Function

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

' Совместная проверка точных границ использует общий topology/query-код.
Private Function OpeningFitsInsideOuter() As Boolean
    On Error GoTo InvalidOpening
    Dim outerContours As CSectionContours, openingContours As CSectionContours, query As CSectionGeometryQuery
    Set outerContours = New CSectionContours: mOuter.BuildContours outerContours
    Set openingContours = New CSectionContours: mOpening.BuildContours openingContours
    Set query = New CSectionGeometryQuery
    Dim material As CConcreteRegion
    Set material = query.SubtractContainedRegion(query.RegionFromContours(outerContours), query.RegionFromContours(openingContours))
    OpeningFitsInsideOuter = (material.Area > GEOM_TOLERANCE)
InvalidOpening:
End Function

Private Function ISectionShape_IsValid(ByRef message As String) As Boolean
    ISectionShape_IsValid = IsValid(message)
End Function

Private Sub ISectionShape_BuildContours(ByVal contours As CSectionContours)
    BuildContours contours
End Sub
