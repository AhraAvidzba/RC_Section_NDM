VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFakeMultiOpeningGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================
' ДЛЯ ТЕСТОВ: новая геометрия с несколькими отверстиями без annotation builder
' ======================================================================
' Реализует только обязательный ISectionGeometry. Наружная форма 100 x 100 мм
' содержит заданное число круглых отверстий радиусом 3 мм. Их ID намеренно
' не содержат слова Opening: потребители должны использовать явную роль.
' Fixture не регистрируется в пользовательском списке типов сечения.

Implements ISectionGeometry

Private mOuter As CGeometryRectSet
Private mOpeningCount As Long

' ДЛЯ ТЕСТОВ: подготавливает симметричную форму с двумя или пятью отверстиями.
Public Sub Initialize(ByVal openingCount As Long)
    If openingCount < 1 Or openingCount > 5 Then Err.Raise vbObjectError + 5283, "CFakeMultiOpeningGeometry", "Неверное число отверстий теста."
    Set mOuter = New CGeometryRectSet
    mOuter.Initialize 100#, 100#, 0#, 0#, 0#, 0#, 0#, "Rectangle"
    mOpeningCount = openingCount
End Sub

' ДЛЯ ТЕСТОВ: равномерно распределяет центры отверстий внутри наружной формы.
Public Function OpeningX(ByVal index As Long) As Double
    OpeningX = 100# * index / (mOpeningCount + 1)
End Function

' ДЛЯ ТЕСТОВ: возвращает точную площадь независимо от query и дискретной сетки.
Public Property Get ExpectedArea() As Double
    ExpectedArea = 10000# - mOpeningCount * GEOM_PI * 9#
End Property

' ДЛЯ ТЕСТОВ: габариты принадлежат наружной форме, а не отдельным отверстиям.
Private Property Get ISectionGeometry_MinX() As Double
    ISectionGeometry_MinX = mOuter.MinX
End Property
Private Property Get ISectionGeometry_MaxX() As Double
    ISectionGeometry_MaxX = mOuter.MaxX
End Property
Private Property Get ISectionGeometry_MinY() As Double
    ISectionGeometry_MinY = mOuter.MinY
End Property
Private Property Get ISectionGeometry_MaxY() As Double
    ISectionGeometry_MaxY = mOuter.MaxY
End Property

' ДЛЯ ТЕСТОВ: исключает все отверстия из принадлежности точки бетону.
Private Function ISectionGeometry_ContainsPoint(ByVal X As Double, ByVal Y As Double) As Boolean
    If Not mOuter.ContainsPoint(X, Y) Then Exit Function
    Dim i As Long
    For i = 1 To mOpeningCount
        If (X - OpeningX(i)) ^ 2 + (Y - 50#) ^ 2 < 9# Then Exit Function
    Next i
    ISectionGeometry_ContainsPoint = True
End Function

' ДЛЯ ТЕСТОВ: обе проверки точки используют одну математическую форму.
Private Function ISectionGeometry_IsPointInside(ByVal X As Double, ByVal Y As Double) As Boolean
    ISectionGeometry_IsPointInside = ISectionGeometry_ContainsPoint(X, Y)
End Function

' ДЛЯ ТЕСТОВ: проверяет подготовку fixture до генерации сетки и контуров.
Private Function ISectionGeometry_IsValid(ByRef message As String) As Boolean
    If mOuter Is Nothing Then
        message = "Тестовая геометрия не инициализирована."
        Exit Function
    End If
    ISectionGeometry_IsValid = mOuter.IsValid(message)
End Function

' ДЛЯ ТЕСТОВ: площадь и симметричный центр имеют независимый точный эталон.
Private Function ISectionGeometry_AnalyticalArea(ByRef available As Boolean) As Double
    available = (mOpeningCount > 0)
    If available Then ISectionGeometry_AnalyticalArea = ExpectedArea
End Function
Private Sub ISectionGeometry_AnalyticalCentroid(ByRef available As Boolean, ByRef X As Double, ByRef Y As Double)
    available = (mOpeningCount > 0)
    X = 50#: Y = 50#
End Sub

' ДЛЯ ТЕСТОВ: опорные точки сетки берутся только по наружному контуру.
Private Sub ISectionGeometry_GetExtremePoints(ByRef pointX() As Double, ByRef pointY() As Double)
    mOuter.GetExtremePoints pointX, pointY
End Sub

' ДЛЯ ТЕСТОВ: передает весь набор границ через обязательный общий API.
' Ни класса аннотаций, ни списка отверстий в потребителях для этого нет.
Private Sub ISectionGeometry_BuildContours(ByVal contours As CSectionContours)
    mOuter.BuildContours contours
    Dim i As Long
    For i = 1 To mOpeningCount
        contours.AddContourCircle "VOID_SEGMENT_" & CStr(i), OpeningX(i), 50#, 3#, _
            "Тестовое отверстие геометрии.", "VOID_" & CStr(i), "Opening"
    Next i
End Sub
