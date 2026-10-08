VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGeometryRectSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Математическая геометрия RectSet: одного или двух прямоугольников
' ==========================================================================
' Класс описывает бетон как один верхний прямоугольник либо как объединение
' верхнего и нижнего прямоугольников. H1/B1 - верхний прямоугольник, H2/B2 -
' нижний. В режиме LSection верхний прямоугольник стоит слева над нижним
' прямоугольником. В режиме TwoRectangles верхний прямоугольник можно
' смещать по X относительно этого базового положения.

Implements ISectionShape

Private Const SECTION_RECTANGLE As String = "rectangle" ' Только верхний прямоугольник B1 x H1.
Private Const SECTION_LSECTION As String = "lsection" ' Г-сечение: верхний прямоугольник слева над нижним.
Private Const SECTION_TWO_RECTANGLES As String = "tworectangles" ' Два прямоугольника, верхний смещается по X.

Private mB1 As Double
Private mH1 As Double
Private mB2 As Double
Private mH2 As Double
Private mOriginX As Double
Private mOriginY As Double
Private mUpperOffsetX As Double
Private mSectionType As String
Private mInitialized As Boolean

' Инициализирует RectSet в одном из трех режимов: Rectangle, LSection или TwoRectangles.
Public Sub Initialize(ByVal b1Mm As Double, ByVal h1Mm As Double, _
        ByVal b2Mm As Double, ByVal h2Mm As Double, _
        Optional ByVal originXMm As Double = 0#, Optional ByVal originYMm As Double = 0#, _
        Optional ByVal upperOffsetXMm As Double = 0#, Optional ByVal sectionType As String = "LSection")
    mB1 = b1Mm
    mH1 = h1Mm
    mB2 = b2Mm
    mH2 = h2Mm
    mOriginX = originXMm
    mOriginY = originYMm
    mUpperOffsetX = upperOffsetXMm
    mSectionType = NormalizeSectionType(sectionType)
    mInitialized = True
End Sub

Public Property Get B1() As Double
    B1 = mB1
End Property

Public Property Get H1() As Double
    H1 = mH1
End Property

Public Property Get B2() As Double
    B2 = mB2
End Property

Public Property Get H2() As Double
    H2 = mH2
End Property

Public Property Get OriginX() As Double
    OriginX = mOriginX
End Property

Public Property Get OriginY() As Double
    OriginY = mOriginY
End Property

' Преобразует проверенный внутренний код формы в ключ общего geometry-контракта.
Public Property Get SectionType() As String
    Select Case mSectionType
        Case SECTION_RECTANGLE: SectionType = "Rectangle"
        Case SECTION_TWO_RECTANGLES: SectionType = "TwoRectangles"
        Case Else: SectionType = "LSection"
    End Select
End Property

Public Property Get UpperOffsetX() As Double
    UpperOffsetX = mUpperOffsetX
End Property

Public Property Get EffectiveUpperOffsetX() As Double
    EffectiveUpperOffsetX = ActiveUpperOffsetX()
End Property

Public Property Get UpperMinX() As Double
    UpperMinX = ActiveUpperMinX()
End Property

Public Property Get UpperMaxX() As Double
    UpperMaxX = ActiveUpperMaxX()
End Property

Public Property Get UpperMinY() As Double
    UpperMinY = ActiveUpperMinY()
End Property

Public Property Get UpperMaxY() As Double
    UpperMaxY = ActiveUpperMaxY()
End Property

Public Property Get LowerMinX() As Double
    LowerMinX = mOriginX
End Property

Public Property Get LowerMaxX() As Double
    LowerMaxX = mOriginX + mB2
End Property

Public Property Get LowerMinY() As Double
    LowerMinY = mOriginY
End Property

Public Property Get LowerMaxY() As Double
    LowerMaxY = mOriginY + mH2
End Property

Public Property Get HasLowerRectangle() As Boolean
    HasLowerRectangle = LowerRectangleActive()
End Property

' Левая граница охватывает только активные прямоугольники и учитывает смещение верхнего.
Public Property Get MinX() As Double
    If LowerRectangleActive() Then
        MinX = GeomMin(mOriginX, ActiveUpperMinX())
    Else
        MinX = ActiveUpperMinX()
    End If
End Property

' Правая граница охватывает только активные прямоугольники и учитывает смещение верхнего.
Public Property Get MaxX() As Double
    If LowerRectangleActive() Then
        MaxX = GeomMax(mOriginX + mB2, ActiveUpperMaxX())
    Else
        MaxX = ActiveUpperMaxX()
    End If
End Property

Public Property Get MinY() As Double
    MinY = mOriginY
End Property

Public Property Get MaxY() As Double
    MaxY = ActiveUpperMaxY()
End Property

' Проверяет, что выбранный режим RectSet имеет положительные размеры и связную геометрию.
Public Function IsValid(ByRef message As String) As Boolean
    message = vbNullString
    If Not mInitialized Then
        message = "Геометрия RectSet не инициализирована."
    ElseIf Len(mSectionType) = 0 Then
        message = "RectSet.SectionType должен быть Rectangle, LSection или TwoRectangles."
    ElseIf mB1 <= 0# Then
        message = "RectSet.B1 должен быть положительным."
    ElseIf mH1 <= 0# Then
        message = "RectSet.H1 должен быть положительным."
    ElseIf LowerRectangleActive() And mB2 <= 0# Then
        message = "RectSet.B2 должен быть положительным."
    ElseIf LowerRectangleActive() And mH2 <= 0# Then
        message = "RectSet.H2 должен быть положительным."
    ElseIf mSectionType = SECTION_LSECTION And mB1 >= mB2 Then
        message = "Для RectSet.SectionType = LSection ширина верхнего прямоугольника RectSet.B1 должна быть меньше RectSet.B2."
    ElseIf mSectionType = SECTION_TWO_RECTANGLES And Not RectanglesOverlapByX() Then
        message = "RectSet.UpperOffsetX задан так, что верхний и нижний прямоугольники не пересекаются по X. Уменьшите смещение или измените B1/B2."
    Else
        IsValid = True
    End If
End Function

' Возвращает внешний контур активной RectSet-геометрии для аннотаций, схемы и AutoCAD export.
Private Sub BuildOutline(ByRef pointX() As Double, ByRef pointY() As Double)
    If Not LowerRectangleActive() Then
        ReDim pointX(1 To 4)
        ReDim pointY(1 To 4)
        pointX(1) = ActiveUpperMinX(): pointY(1) = ActiveUpperMinY()
        pointX(2) = ActiveUpperMaxX(): pointY(2) = ActiveUpperMinY()
        pointX(3) = ActiveUpperMaxX(): pointY(3) = ActiveUpperMaxY()
        pointX(4) = ActiveUpperMinX(): pointY(4) = ActiveUpperMaxY()
        Exit Sub
    End If

    Dim count As Long
    AddOutlinePoint pointX, pointY, count, mOriginX, mOriginY
    AddOutlinePoint pointX, pointY, count, mOriginX + mB2, mOriginY
    AddOutlinePoint pointX, pointY, count, mOriginX + mB2, mOriginY + mH2
    AddOutlinePoint pointX, pointY, count, ActiveUpperMaxX(), mOriginY + mH2
    AddOutlinePoint pointX, pointY, count, ActiveUpperMaxX(), ActiveUpperMaxY()
    AddOutlinePoint pointX, pointY, count, ActiveUpperMinX(), ActiveUpperMaxY()
    AddOutlinePoint pointX, pointY, count, ActiveUpperMinX(), mOriginY + mH2
    AddOutlinePoint pointX, pointY, count, mOriginX, mOriginY + mH2

    ReDim Preserve pointX(1 To count)
    ReDim Preserve pointY(1 To count)
End Sub

' Приводит тип сечения и поддерживаемые aliases к расчетному ключу.
' Пустое значение/RectSet выбирает LSection; неизвестное имя возвращает пустой ключ для валидации.
Private Function NormalizeSectionType(ByVal rawValue As String) As String
    Dim value As String
    value = LCase$(Trim$(rawValue))
    Select Case value
        Case vbNullString, SECTION_LSECTION, "rectset"
            NormalizeSectionType = SECTION_LSECTION
        Case SECTION_RECTANGLE
            NormalizeSectionType = SECTION_RECTANGLE
        Case SECTION_TWO_RECTANGLES, "tworectangle", "tworectanglesoffset", "offsettworectangles"
            NormalizeSectionType = SECTION_TWO_RECTANGLES
        Case Else
            NormalizeSectionType = vbNullString
    End Select
End Function

' Возвращает True, если нижний прямоугольник участвует в текущем режиме геометрии.
Private Function LowerRectangleActive() As Boolean
    LowerRectangleActive = (mSectionType <> SECTION_RECTANGLE)
End Function

' Возвращает фактическое смещение верхнего прямоугольника: только TwoRectangles использует Offset.
Private Function ActiveUpperOffsetX() As Double
    If mSectionType = SECTION_TWO_RECTANGLES Then ActiveUpperOffsetX = mUpperOffsetX
End Function

' Возвращает левую X-границу активного верхнего прямоугольника.
Private Function ActiveUpperMinX() As Double
    ActiveUpperMinX = mOriginX + ActiveUpperOffsetX()
End Function

' Возвращает правую X-границу активного верхнего прямоугольника.
Private Function ActiveUpperMaxX() As Double
    ActiveUpperMaxX = ActiveUpperMinX() + mB1
End Function

' Возвращает нижнюю Y-границу активного верхнего прямоугольника.
Private Function ActiveUpperMinY() As Double
    If LowerRectangleActive() Then
        ActiveUpperMinY = mOriginY + mH2
    Else
        ActiveUpperMinY = mOriginY
    End If
End Function

' Возвращает верхнюю Y-границу активного верхнего прямоугольника.
Private Function ActiveUpperMaxY() As Double
    ActiveUpperMaxY = ActiveUpperMinY() + mH1
End Function

' Проверяет горизонтальное пересечение прямоугольников, чтобы TwoRectangles оставался связным сечением.
Private Function RectanglesOverlapByX() As Boolean
    RectanglesOverlapByX = (ActiveUpperMinX() < mOriginX + mB2 - GEOM_TOLERANCE _
        And ActiveUpperMaxX() > mOriginX + GEOM_TOLERANCE)
End Function

' Добавляет точку контура, пропуская дубли на стыках прямоугольников.
Private Sub AddOutlinePoint(ByRef pointX() As Double, ByRef pointY() As Double, ByRef count As Long, _
        ByVal X As Double, ByVal Y As Double)
    If count = 0 Then
        ReDim pointX(1 To 8)
        ReDim pointY(1 To 8)
    ElseIf Abs(pointX(count) - X) <= GEOM_TOLERANCE And Abs(pointY(count) - Y) <= GEOM_TOLERANCE Then
        Exit Sub
    ElseIf count >= UBound(pointX) Then
        ReDim Preserve pointX(1 To UBound(pointX) + 4)
        ReDim Preserve pointY(1 To UBound(pointY) + 4)
    End If

    count = count + 1
    pointX(count) = X
    pointY(count) = Y
End Sub

' Передает точную наружную ломаную в общий контейнер геометрии.
' Роль Outer задается здесь и не зависит от размеров или подписей формы.
Public Sub BuildContours(ByVal contours As CSectionContours)
    If contours Is Nothing Then Err.Raise vbObjectError + 5283, "CGeometryRectSet", "Контейнер контуров сечения не передан."
    Dim message As String
    If Not IsValid(message) Then Err.Raise vbObjectError + 5283, "CGeometryRectSet", message
    Dim px() As Double, py() As Double, i As Long, nextIndex As Long
    BuildOutline px, py
    contours.Clear
    For i = LBound(px) To UBound(px)
        nextIndex = i + 1
        If i = UBound(px) Then nextIndex = LBound(px)
        contours.AddContourLine "CONTOUR_LINE_" & CStr(i), px(i), py(i), px(nextIndex), py(nextIndex), _
            "Сегмент контура сечения.", "OUTER", "Outer"
    Next i
End Sub


Private Function ISectionShape_IsValid(ByRef message As String) As Boolean
    ISectionShape_IsValid = IsValid(message)
End Function

Private Sub ISectionShape_BuildContours(ByVal contours As CSectionContours)
    BuildContours contours
End Sub
