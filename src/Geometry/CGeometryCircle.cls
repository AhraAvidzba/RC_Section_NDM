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
' принадлежность точки области и достоверный контур с явной ролью Outer.
' Арматура и необязательные подписи формируются отдельными builder-ами.

Implements ISectionGeometry

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

' Проверяет принадлежность точки замкнутой области круга с малым
' геометрическим допуском на границе.
Public Function ContainsPoint(ByVal X As Double, ByVal Y As Double) As Boolean
    Dim message As String
    If Not IsValid(message) Then Exit Function

    Dim dx As Double
    Dim dy As Double
    dx = X - mCenterX
    dy = Y - mCenterY
    ContainsPoint = (dx * dx + dy * dy <= mRadius * mRadius + GEOM_TOLERANCE)
End Function

' Явный алиас для контракта ISectionGeometry: точка на контуре считается
' находящейся внутри расчетной бетонной области.
Public Function IsPointInside(ByVal X As Double, ByVal Y As Double) As Boolean
    IsPointInside = ContainsPoint(X, Y)
End Function

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

' Возвращает аналитическую площадь круга, когда геометрия валидна.
Public Function AnalyticalArea(ByRef available As Boolean) As Double
    Dim message As String
    available = IsValid(message)
    If available Then AnalyticalArea = GEOM_PI * mRadius * mRadius
End Function

' Возвращает аналитический центр тяжести круга.
Public Sub AnalyticalCentroid(ByRef available As Boolean, ByRef X As Double, ByRef Y As Double)
    Dim message As String
    available = IsValid(message)
    If Not available Then Exit Sub
    X = mCenterX
    Y = mCenterY
End Sub

' Возвращает четыре опорные точки контура по главным направлениям.
' Их достаточно для габаритов, аннотаций и первичного масштабирования схемы.
Public Sub GetExtremePoints(ByRef pointX() As Double, ByRef pointY() As Double)
    ReDim pointX(1 To 4)
    ReDim pointY(1 To 4)
    pointX(1) = mCenterX + mRadius: pointY(1) = mCenterY
    pointX(2) = mCenterX: pointY(2) = mCenterY + mRadius
    pointX(3) = mCenterX - mRadius: pointY(3) = mCenterY
    pointX(4) = mCenterX: pointY(4) = mCenterY - mRadius
End Sub

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

' Реализация интерфейса ISectionGeometry: свойства и методы ниже только
' делегируют публичной математике круга.
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

' Делегирует проверку точки в публичный ContainsPoint.
Private Function ISectionGeometry_ContainsPoint(ByVal X As Double, ByVal Y As Double) As Boolean
    ISectionGeometry_ContainsPoint = ContainsPoint(X, Y)
End Function

' Делегирует проверку точки в публичный IsPointInside.
Private Function ISectionGeometry_IsPointInside(ByVal X As Double, ByVal Y As Double) As Boolean
    ISectionGeometry_IsPointInside = IsPointInside(X, Y)
End Function

' Делегирует валидацию в публичный IsValid.
Private Function ISectionGeometry_IsValid(ByRef message As String) As Boolean
    ISectionGeometry_IsValid = IsValid(message)
End Function

' Делегирует аналитическую площадь в публичный AnalyticalArea.
Private Function ISectionGeometry_AnalyticalArea(ByRef available As Boolean) As Double
    ISectionGeometry_AnalyticalArea = AnalyticalArea(available)
End Function

' Делегирует аналитический центр тяжести в публичный AnalyticalCentroid.
Private Sub ISectionGeometry_AnalyticalCentroid(ByRef available As Boolean, ByRef X As Double, ByRef Y As Double)
    AnalyticalCentroid available, X, Y
End Sub

' Делегирует опорные точки контура в публичный GetExtremePoints.
Private Sub ISectionGeometry_GetExtremePoints(ByRef pointX() As Double, ByRef pointY() As Double)
    GetExtremePoints pointX, pointY
End Sub

' Передает классифицированные контуры через обязательный интерфейс формы.
Private Sub ISectionGeometry_BuildContours(ByVal contours As CSectionContours)
    BuildContours contours
End Sub


