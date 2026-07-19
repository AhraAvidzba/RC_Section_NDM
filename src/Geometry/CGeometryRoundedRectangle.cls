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

Implements ISectionGeometry

Private mWidth As Double
Private mHeight As Double
Private mRadiusTL As Double
Private mRadiusTR As Double
Private mRadiusBR As Double
Private mRadiusBL As Double
Private mCenterX As Double
Private mCenterY As Double
Private mInitialized As Boolean

Public Sub Initialize(ByVal widthMm As Double, ByVal heightMm As Double, _
        ByVal radiusTLmm As Double, ByVal radiusTRmm As Double, _
        ByVal radiusBRmm As Double, ByVal radiusBLmm As Double, _
        Optional ByVal centerXMm As Double = 0#, Optional ByVal centerYMm As Double = 0#)

    mWidth = widthMm
    mHeight = heightMm
    mRadiusTL = radiusTLmm
    mRadiusTR = radiusTRmm
    mRadiusBR = radiusBRmm
    mRadiusBL = radiusBLmm
    mCenterX = centerXMm
    mCenterY = centerYMm
    mInitialized = True
End Sub

Public Property Get Width() As Double
    Width = mWidth
End Property

Public Property Get Height() As Double
    Height = mHeight
End Property

Public Property Get RadiusTL() As Double
    RadiusTL = mRadiusTL
End Property

Public Property Get RadiusTR() As Double
    RadiusTR = mRadiusTR
End Property

Public Property Get RadiusBR() As Double
    RadiusBR = mRadiusBR
End Property

Public Property Get RadiusBL() As Double
    RadiusBL = mRadiusBL
End Property

Public Property Get CenterX() As Double
    CenterX = mCenterX
End Property

Public Property Get CenterY() As Double
    CenterY = mCenterY
End Property

Public Property Get MinX() As Double
    MinX = mCenterX - mWidth / 2#
End Property

Public Property Get MaxX() As Double
    MaxX = mCenterX + mWidth / 2#
End Property

Public Property Get MinY() As Double
    MinY = mCenterY - mHeight / 2#
End Property

Public Property Get MaxY() As Double
    MaxY = mCenterY + mHeight / 2#
End Property

Public Function ContainsPoint(ByVal X As Double, ByVal Y As Double) As Boolean
    Dim message As String
    If Not IsValid(message) Then
        ContainsPoint = False
        Exit Function
    End If

    If X < MinX - GEOM_TOLERANCE Or X > MaxX + GEOM_TOLERANCE Then Exit Function
    If Y < MinY - GEOM_TOLERANCE Or Y > MaxY + GEOM_TOLERANCE Then Exit Function

    If IsOutsideCorner(X, Y, MinX, MaxY, mRadiusTL, 1, -1) Then Exit Function
    If IsOutsideCorner(X, Y, MaxX, MaxY, mRadiusTR, -1, -1) Then Exit Function
    If IsOutsideCorner(X, Y, MaxX, MinY, mRadiusBR, -1, 1) Then Exit Function
    If IsOutsideCorner(X, Y, MinX, MinY, mRadiusBL, 1, 1) Then Exit Function

    ContainsPoint = True
End Function

Public Function IsPointInside(ByVal X As Double, ByVal Y As Double) As Boolean
    IsPointInside = ContainsPoint(X, Y)
End Function

Public Function IsValid(ByRef message As String) As Boolean
    message = vbNullString

    If Not mInitialized Then
        message = "????????? ?? ????????????????."
        Exit Function
    End If

    If mWidth <= 0# Then
        message = "?????? ?????? ???? ?????? ????."
        Exit Function
    End If

    If mHeight <= 0# Then
        message = "?????? ?????? ???? ?????? ????."
        Exit Function
    End If

    If mRadiusTL < 0# Or mRadiusTR < 0# Or mRadiusBR < 0# Or mRadiusBL < 0# Then
        message = "??????? ?? ?????? ???? ??????????????."
        Exit Function
    End If

    If mRadiusTL > GeomMin(mWidth, mHeight) Or _
       mRadiusTR > GeomMin(mWidth, mHeight) Or _
       mRadiusBR > GeomMin(mWidth, mHeight) Or _
       mRadiusBL > GeomMin(mWidth, mHeight) Then
        message = "?????? ???? ?????? ???????????? ??????? ???????."
        Exit Function
    End If

    If mRadiusTL + mRadiusTR > mWidth + GEOM_TOLERANCE Then
        message = "????? ??????? ???????? ?????? ??????."
        Exit Function
    End If

    If mRadiusBL + mRadiusBR > mWidth + GEOM_TOLERANCE Then
        message = "????? ?????? ???????? ?????? ??????."
        Exit Function
    End If

    If mRadiusTL + mRadiusBL > mHeight + GEOM_TOLERANCE Then
        message = "????? ????? ???????? ?????? ??????."
        Exit Function
    End If

    If mRadiusTR + mRadiusBR > mHeight + GEOM_TOLERANCE Then
        message = "????? ?????? ???????? ?????? ??????."
        Exit Function
    End If

    If AnalyticalAreaInternal() <= GEOM_TOLERANCE Then
        message = "??????? ??????? ?????????."
        Exit Function
    End If

    IsValid = True
End Function

Public Function AnalyticalArea(ByRef available As Boolean) As Double
    Dim message As String
    available = IsValid(message)
    If available Then AnalyticalArea = AnalyticalAreaInternal()
End Function

Public Sub AnalyticalCentroid(ByRef available As Boolean, ByRef X As Double, ByRef Y As Double)
    Dim message As String
    available = IsValid(message)
    If Not available Then Exit Sub

    Dim baseArea As Double
    Dim sx As Double
    Dim sy As Double
    Dim area As Double

    baseArea = mWidth * mHeight
    area = baseArea
    sx = baseArea * mCenterX
    sy = baseArea * mCenterY

    SubtractCorner area, sx, sy, mRadiusTL, MinX, MaxY, 1, -1
    SubtractCorner area, sx, sy, mRadiusTR, MaxX, MaxY, -1, -1
    SubtractCorner area, sx, sy, mRadiusBR, MaxX, MinY, -1, 1
    SubtractCorner area, sx, sy, mRadiusBL, MinX, MinY, 1, 1

    X = sx / area
    Y = sy / area
End Sub

Public Sub GetExtremePoints(ByRef pointX() As Double, ByRef pointY() As Double)
    ReDim pointX(1 To 8)
    ReDim pointY(1 To 8)

    pointX(1) = MinX + mRadiusTL: pointY(1) = MaxY
    pointX(2) = MaxX - mRadiusTR: pointY(2) = MaxY
    pointX(3) = MaxX: pointY(3) = MaxY - mRadiusTR
    pointX(4) = MaxX: pointY(4) = MinY + mRadiusBR
    pointX(5) = MaxX - mRadiusBR: pointY(5) = MinY
    pointX(6) = MinX + mRadiusBL: pointY(6) = MinY
    pointX(7) = MinX: pointY(7) = MinY + mRadiusBL
    pointX(8) = MinX: pointY(8) = MaxY - mRadiusTL
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

Private Function ISectionGeometry_ContainsPoint(ByVal X As Double, ByVal Y As Double) As Boolean
    ISectionGeometry_ContainsPoint = ContainsPoint(X, Y)
End Function

Private Function ISectionGeometry_IsPointInside(ByVal X As Double, ByVal Y As Double) As Boolean
    ISectionGeometry_IsPointInside = IsPointInside(X, Y)
End Function

Private Function ISectionGeometry_IsValid(ByRef message As String) As Boolean
    ISectionGeometry_IsValid = IsValid(message)
End Function

Private Function ISectionGeometry_AnalyticalArea(ByRef available As Boolean) As Double
    ISectionGeometry_AnalyticalArea = AnalyticalArea(available)
End Function

Private Sub ISectionGeometry_AnalyticalCentroid(ByRef available As Boolean, ByRef X As Double, ByRef Y As Double)
    AnalyticalCentroid available, X, Y
End Sub

Private Sub ISectionGeometry_GetExtremePoints(ByRef pointX() As Double, ByRef pointY() As Double)
    GetExtremePoints pointX, pointY
End Sub

Private Function IsOutsideCorner(ByVal X As Double, ByVal Y As Double, _
        ByVal cornerX As Double, ByVal cornerY As Double, ByVal radius As Double, _
        ByVal signX As Double, ByVal signY As Double) As Boolean

    If radius <= GEOM_TOLERANCE Then Exit Function

    Dim centerX As Double
    Dim centerY As Double
    centerX = cornerX + signX * radius
    centerY = cornerY + signY * radius

    If signX * (X - centerX) <= GEOM_TOLERANCE And _
       signY * (Y - centerY) <= GEOM_TOLERANCE Then

        Dim dx As Double
        Dim dy As Double
        dx = X - centerX
        dy = Y - centerY
        IsOutsideCorner = (dx * dx + dy * dy > radius * radius + GEOM_TOLERANCE)
    End If
End Function

Private Function AnalyticalAreaInternal() As Double
    Dim k As Double
    k = 1# - GEOM_PI / 4#
    AnalyticalAreaInternal = mWidth * mHeight - k * (mRadiusTL ^ 2 + mRadiusTR ^ 2 + mRadiusBR ^ 2 + mRadiusBL ^ 2)
End Function

Private Sub SubtractCorner(ByRef area As Double, ByRef sx As Double, ByRef sy As Double, _
        ByVal radius As Double, ByVal cornerX As Double, ByVal cornerY As Double, _
        ByVal signX As Double, ByVal signY As Double)

    If radius <= GEOM_TOLERANCE Then Exit Sub

    Dim removedArea As Double
    removedArea = radius * radius * (1# - GEOM_PI / 4#)
    If removedArea <= GEOM_TOLERANCE Then Exit Sub

    Dim squareArea As Double
    Dim quarterArea As Double
    Dim squareX As Double
    Dim squareY As Double
    Dim quarterX As Double
    Dim quarterY As Double
    Dim removedX As Double
    Dim removedY As Double
    Dim centerX As Double
    Dim centerY As Double

    squareArea = radius * radius
    quarterArea = GEOM_PI * radius * radius / 4#
    centerX = cornerX + signX * radius
    centerY = cornerY + signY * radius
    squareX = cornerX + signX * radius / 2#
    squareY = cornerY + signY * radius / 2#
    quarterX = centerX - signX * 4# * radius / (3# * GEOM_PI)
    quarterY = centerY - signY * 4# * radius / (3# * GEOM_PI)

    removedX = (squareArea * squareX - quarterArea * quarterX) / removedArea
    removedY = (squareArea * squareY - quarterArea * quarterY) / removedArea

    area = area - removedArea
    sx = sx - removedArea * removedX
    sy = sy - removedArea * removedY
End Sub










