VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGeometryCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Математическая геометрия круглого сечения
' ==========================================================================
' Класс хранит диаметр круга и отвечает на геометрические вопросы: габариты,
' принадлежность точки области и контур для визуального вывода. Арматура и
' подписи рядов формируются отдельными builder-ами.

Implements ISectionGeometry

Private mRadius As Double
Private mCenterX As Double
Private mCenterY As Double
Private mInitialized As Boolean

Public Sub InitializeByRadius(ByVal radiusMm As Double, _
        Optional ByVal centerXMm As Double = 0#, Optional ByVal centerYMm As Double = 0#)
    mRadius = radiusMm
    mCenterX = centerXMm
    mCenterY = centerYMm
    mInitialized = True
End Sub

Public Sub InitializeByDiameter(ByVal diameterMm As Double, _
        Optional ByVal centerXMm As Double = 0#, Optional ByVal centerYMm As Double = 0#)
    InitializeByRadius diameterMm / 2#, centerXMm, centerYMm
End Sub

Public Property Get Radius() As Double
    Radius = mRadius
End Property

Public Property Get Diameter() As Double
    Diameter = 2# * mRadius
End Property

Public Property Get CenterX() As Double
    CenterX = mCenterX
End Property

Public Property Get CenterY() As Double
    CenterY = mCenterY
End Property

Public Property Get MinX() As Double
    MinX = mCenterX - mRadius
End Property

Public Property Get MaxX() As Double
    MaxX = mCenterX + mRadius
End Property

Public Property Get MinY() As Double
    MinY = mCenterY - mRadius
End Property

Public Property Get MaxY() As Double
    MaxY = mCenterY + mRadius
End Property

Public Function ContainsPoint(ByVal X As Double, ByVal Y As Double) As Boolean
    Dim message As String
    If Not IsValid(message) Then Exit Function

    Dim dx As Double
    Dim dy As Double
    dx = X - mCenterX
    dy = Y - mCenterY
    ContainsPoint = (dx * dx + dy * dy <= mRadius * mRadius + GEOM_TOLERANCE)
End Function

Public Function IsPointInside(ByVal X As Double, ByVal Y As Double) As Boolean
    IsPointInside = ContainsPoint(X, Y)
End Function

Public Function IsValid(ByRef message As String) As Boolean
    message = vbNullString
    If Not mInitialized Then
        message = "Геометрия круга не инициализирована."
    ElseIf mRadius <= 0# Then
        message = "Радиус круга должен быть положительным."
    Else
        IsValid = True
    End If
End Function

Public Function AnalyticalArea(ByRef available As Boolean) As Double
    Dim message As String
    available = IsValid(message)
    If available Then AnalyticalArea = GEOM_PI * mRadius * mRadius
End Function

Public Sub AnalyticalCentroid(ByRef available As Boolean, ByRef X As Double, ByRef Y As Double)
    Dim message As String
    available = IsValid(message)
    If Not available Then Exit Sub
    X = mCenterX
    Y = mCenterY
End Sub

Public Sub GetExtremePoints(ByRef pointX() As Double, ByRef pointY() As Double)
    ReDim pointX(1 To 4)
    ReDim pointY(1 To 4)
    pointX(1) = mCenterX + mRadius: pointY(1) = mCenterY
    pointX(2) = mCenterX: pointY(2) = mCenterY + mRadius
    pointX(3) = mCenterX - mRadius: pointY(3) = mCenterY
    pointX(4) = mCenterX: pointY(4) = mCenterY - mRadius
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


