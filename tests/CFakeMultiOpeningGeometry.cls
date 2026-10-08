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
' Реализует только обязательный ISectionShape. Наружная форма 100 x 100 мм
' содержит заданное число круглых отверстий радиусом 3 мм. Их ID намеренно
' не содержат слова Opening: потребители должны использовать явную роль.
' Fixture не регистрируется в пользовательском списке типов сечения.

Implements ISectionShape

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

' ДЛЯ ТЕСТОВ: проверяет подготовку fixture до генерации сетки и контуров.
Private Function ISectionShape_IsValid(ByRef message As String) As Boolean
    If mOuter Is Nothing Then
        message = "Тестовая геометрия не инициализирована."
        Exit Function
    End If
    ISectionShape_IsValid = mOuter.IsValid(message)
End Function

' ДЛЯ ТЕСТОВ: передает весь набор границ через обязательный общий API.
' Ни класса аннотаций, ни списка отверстий в потребителях для этого нет.
Private Sub ISectionShape_BuildContours(ByVal contours As CSectionContours)
    mOuter.BuildContours contours
    Dim i As Long
    For i = 1 To mOpeningCount
        contours.AddContourCircle "VOID_SEGMENT_" & CStr(i), OpeningX(i), 50#, 3#, _
            "Тестовое отверстие геометрии.", "VOID_" & CStr(i), "Opening"
    Next i
End Sub
