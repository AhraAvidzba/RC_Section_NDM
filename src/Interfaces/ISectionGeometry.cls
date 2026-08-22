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
' контур. Он не знает об Excel, арматуре, Results и расчетных solver-ах.

Public Property Get MinX() As Double
End Property

Public Property Get MaxX() As Double
End Property

Public Property Get MinY() As Double
End Property

Public Property Get MaxY() As Double
End Property

Public Function ContainsPoint(ByVal X As Double, ByVal Y As Double) As Boolean
End Function

Public Function IsPointInside(ByVal X As Double, ByVal Y As Double) As Boolean
End Function

Public Function IsValid(ByRef message As String) As Boolean
End Function

Public Function AnalyticalArea(ByRef available As Boolean) As Double
End Function

Public Sub AnalyticalCentroid(ByRef available As Boolean, ByRef X As Double, ByRef Y As Double)
End Sub

Public Sub GetExtremePoints(ByRef pointX() As Double, ByRef pointY() As Double)
End Sub










