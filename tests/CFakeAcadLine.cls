VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFakeAcadLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Тестовая имитация AutoCAD Line
' ==========================================================================
' Класс нужен только unit-тестам AutoCAD importer-а. Он дает минимальный COM-
' контракт прямой грани Region: ObjectName, StartPoint, EndPoint и Delete.

Private mStartPoint(0 To 2) As Double ' Начальная точка тестовой линии в единицах теста.
Private mEndPoint(0 To 2) As Double ' Конечная точка тестовой линии в единицах теста.
Private mLayer As String ' ДЛЯ ТЕСТОВ: слой линии в fixture очистки AutoCAD.
Private mDeleted As Boolean ' ДЛЯ ТЕСТОВ: факт вызова Delete, без изменения чужих COM-объектов.

' Задает две точки тестовой линии.
Public Sub Initialize(ByVal x1 As Double, ByVal y1 As Double, _
        ByVal x2 As Double, ByVal y2 As Double, Optional ByVal layerName As String = "")
    mStartPoint(0) = x1
    mStartPoint(1) = y1
    mStartPoint(2) = 0#
    mEndPoint(0) = x2
    mEndPoint(1) = y2
    mEndPoint(2) = 0#
    mLayer = layerName
    mDeleted = False
End Sub

' Возвращает имя AutoCAD-типа, по которому importer распознает прямую грань.
Public Property Get ObjectName() As String
    ObjectName = "AcDbLine"
End Property

' Возвращает начальную точку в формате, похожем на AutoCAD COM Variant-array.
Public Property Get StartPoint() As Variant
    StartPoint = mStartPoint
End Property

' Возвращает конечную точку в формате, похожем на AutoCAD COM Variant-array.
Public Property Get EndPoint() As Variant
    EndPoint = mEndPoint
End Property

' Имитирует удаление временного AutoCAD-объекта после Explode.
Public Sub Delete()
    mDeleted = True
End Sub

' ============================== ДЛЯ ТЕСТОВ ==============================
' Слой и факт Delete нужны общему алгоритму очистки; координаты и тип Line
' по-прежнему используются importer-ом при чтении грани Region.
Public Property Get Layer() As String
    Layer = mLayer
End Property

Public Property Get Deleted() As Boolean
    Deleted = mDeleted
End Property
