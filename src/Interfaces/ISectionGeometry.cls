VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISectionGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Контракт бетонной геометрии
' ==========================================================================
' Интерфейс описывает только математическую форму бетонной части сечения:
' попадает ли точка внутрь, какие габариты у модели и как дискретизировать
' контур. Точные наружные границы и отверстия передаются с явными ролями
' через общий контейнер, независимо от необязательных аннотаций.
' Он не знает об Excel, арматуре, Results и расчетных solver-ах.

' Минимальная X-координата габаритного прямоугольника бетонной формы.
Public Property Get MinX() As Double
End Property

' Максимальная X-координата габаритного прямоугольника бетонной формы.
Public Property Get MaxX() As Double
End Property

' Минимальная Y-координата габаритного прямоугольника бетонной формы.
Public Property Get MinY() As Double
End Property

' Максимальная Y-координата габаритного прямоугольника бетонной формы.
Public Property Get MaxY() As Double
End Property

' Проверяет, принадлежит ли точка расчетной бетонной области.
Public Function ContainsPoint(ByVal X As Double, ByVal Y As Double) As Boolean
End Function

' Явное имя для проверок, где важно читать код как "точка внутри".
Public Function IsPointInside(ByVal X As Double, ByVal Y As Double) As Boolean
End Function

' Проверяет корректность параметрической геометрии и возвращает сообщение
' пользователю при ошибке.
Public Function IsValid(ByRef message As String) As Boolean
End Function

' Возвращает аналитическую площадь, если конкретная форма умеет считать ее
' без дискретной сетки.
Public Function AnalyticalArea(ByRef available As Boolean) As Double
End Function

' Возвращает аналитический центр тяжести, если он доступен для формы.
Public Sub AnalyticalCentroid(ByRef available As Boolean, ByRef X As Double, ByRef Y As Double)
End Sub

' Возвращает опорные точки внешнего контура для масштабирования и аннотаций.
Public Sub GetExtremePoints(ByRef pointX() As Double, ByRef pointY() As Double)
End Sub

' Заменяет достоверные контуры формы во внутренних мм. Каждый замкнутый
' контур получает собственный LoopID и роль Outer либо Opening сразу здесь,
' чтобы все потребители одинаково обрабатывали любую зарегистрированную форму.
Public Sub BuildContours(ByVal contours As CSectionContours)
End Sub










