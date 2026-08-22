VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGeometryLShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Математическая геометрия Г-образного сечения
' ==========================================================================
' Класс описывает бетон как объединение верхнего и нижнего прямоугольников:
' H1/B1 - верхняя ветвь, H2/B2 - нижняя полка. Он задает форму и контур, но
' не занимается раскладкой стержней и не решает, какие размеры подписывать.

Implements ISectionGeometry

Private mB1 As Double
Private mH1 As Double
Private mB2 As Double
Private mH2 As Double
Private mOriginX As Double
Private mOriginY As Double
Private mInitialized As Boolean

Public Sub Initialize(ByVal b1Mm As Double, ByVal h1Mm As Double, _
        ByVal b2Mm As Double, ByVal h2Mm As Double, _
        Optional ByVal originXMm As Double = 0#, Optional ByVal originYMm As Double = 0#)
    mB1 = b1Mm
    mH1 = h1Mm
    mB2 = b2Mm
    mH2 = h2Mm
    mOriginX = originXMm
    mOriginY = originYMm
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

Public Property Get MinX() As Double
    MinX = mOriginX
End Property

Public Property Get MaxX() As Double
    MaxX = mOriginX + mB2
End Property

Public Property Get MinY() As Double
    MinY = mOriginY
End Property

Public Property Get MaxY() As Double
    MaxY = mOriginY + mH2 + mH1
End Property

Public Function ContainsPoint(ByVal X As Double, ByVal Y As Double) As Boolean
    Dim message As String
    If Not IsValid(message) Then Exit Function

    Dim insideLower As Boolean
    Dim insideVertical As Boolean
    insideLower = (X >= mOriginX - GEOM_TOLERANCE And X <= mOriginX + mB2 + GEOM_TOLERANCE _
        And Y >= mOriginY - GEOM_TOLERANCE And Y <= mOriginY + mH2 + GEOM_TOLERANCE)
    insideVertical = (X >= mOriginX - GEOM_TOLERANCE And X <= mOriginX + mB1 + GEOM_TOLERANCE _
        And Y >= mOriginY + mH2 - GEOM_TOLERANCE And Y <= mOriginY + mH2 + mH1 + GEOM_TOLERANCE)
    ContainsPoint = insideLower Or insideVertical
End Function

Public Function IsPointInside(ByVal X As Double, ByVal Y As Double) As Boolean
    IsPointInside = ContainsPoint(X, Y)
End Function

Public Function IsValid(ByRef message As String) As Boolean
    message = vbNullString
    If Not mInitialized Then
        message = "L-shape geometry is not initialized."
    ElseIf mB1 <= 0# Then
        message = "LShape.B1 must be positive."
    ElseIf mH1 <= 0# Then
        message = "LShape.H1 must be positive."
    ElseIf mB2 <= 0# Then
        message = "LShape.B2 must be positive."
    ElseIf mH2 <= 0# Then
        message = "LShape.H2 must be positive."
    ElseIf mB1 >= mB2 Then
        message = "LShape.B1 must be less than LShape.B2."
    Else
        IsValid = True
    End If
End Function

Public Function AnalyticalArea(ByRef available As Boolean) As Double
    Dim message As String
    available = IsValid(message)
    If available Then AnalyticalArea = mB2 * mH2 + mB1 * mH1
End Function

Public Sub AnalyticalCentroid(ByRef available As Boolean, ByRef X As Double, ByRef Y As Double)
    Dim message As String
    available = IsValid(message)
    If Not available Then Exit Sub

    Dim areaLower As Double
    Dim areaUpper As Double
    areaLower = mB2 * mH2
    areaUpper = mB1 * mH1
    X = (areaLower * (mOriginX + mB2 / 2#) + areaUpper * (mOriginX + mB1 / 2#)) / (areaLower + areaUpper)
    Y = (areaLower * (mOriginY + mH2 / 2#) + areaUpper * (mOriginY + mH2 + mH1 / 2#)) / (areaLower + areaUpper)
End Sub

Public Sub GetExtremePoints(ByRef pointX() As Double, ByRef pointY() As Double)
    ReDim pointX(1 To 6)
    ReDim pointY(1 To 6)
    pointX(1) = mOriginX: pointY(1) = mOriginY
    pointX(2) = mOriginX + mB2: pointY(2) = mOriginY
    pointX(3) = mOriginX + mB2: pointY(3) = mOriginY + mH2
    pointX(4) = mOriginX + mB1: pointY(4) = mOriginY + mH2
    pointX(5) = mOriginX + mB1: pointY(5) = mOriginY + mH2 + mH1
    pointX(6) = mOriginX: pointY(6) = mOriginY + mH2 + mH1
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

