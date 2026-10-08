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
' Параметры и контур круглого сечения
' ==========================================================================
' Класс хранит размеры и положение круга, задает точный контур с ролью Outer.
' Принадлежность точек и свойства бетона вычисляются общим геометрическим ядром.
' Арматура и необязательные подписи формируются отдельными builder-ами.

Implements ISectionShape

Private mRadius As Double
Private mCenterX As Double
Private mCenterY As Double
Private mInitialized As Boolean

' Инициализирует круг через радиус и центр во внутренних миллиметрах.
Public Sub InitializeByRadius(ByVal radiusMm As Double, _
        Optional ByVal centerXMm As Double = 0#, Optional ByVal centerYMm As Double = 0#)
    mRadius = radiusMm
    mCenterX = centerXMm
    mCenterY = centerYMm
    mInitialized = True
End Sub

' Удобная оболочка для пользовательского ввода диаметра: расчетная геометрия
' дальше хранит только радиус.
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

' Проверяет, что круг был инициализирован и радиус физически положителен.
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

' Передает точную окружность как наружную границу в общий контейнер.
' Аннотации и видимость схемы не участвуют в создании геометрии модели.
Public Sub BuildContours(ByVal contours As CSectionContours)
    If contours Is Nothing Then Err.Raise vbObjectError + 5283, "CGeometryCircle", "Контейнер контуров сечения не передан."
    Dim message As String
    If Not IsValid(message) Then Err.Raise vbObjectError + 5283, "CGeometryCircle", message
    contours.Clear
    contours.AddContourCircle "CONTOUR_CIRCLE", 0.5 * (MinX + MaxX), 0.5 * (MinY + MaxY), _
        0.5 * GeomMin(MaxX - MinX, MaxY - MinY), "Круговой контур сечения.", "OUTER", "Outer"
End Sub


Private Function ISectionShape_IsValid(ByRef message As String) As Boolean
    ISectionShape_IsValid = IsValid(message)
End Function

Private Sub ISectionShape_BuildContours(ByVal contours As CSectionContours)
    BuildContours contours
End Sub
