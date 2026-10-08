VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGeometryRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Математическая форма по готовому точному контуру бетона
' ==========================================================================
' Адаптер ISectionGeometry для обычного CFiberMeshBuilder. Внешние границы,
' дуги и несколько отверстий уже проверены общим geometry-query. Здесь нет
' AutoCAD COM, настроек Excel, арматуры или расчетных алгоритмов.

Implements ISectionGeometry

Private mRegion As CConcreteRegion
Private mQuery As CSectionGeometryQuery
Private mEdges() As TRegionEdge
Private mEdgeCount As Long
Private mMinX As Double, mMaxX As Double, mMinY As Double, mMaxY As Double

' Снимок исключает изменение формы во время генерации; ребра копируются один
' раз, а не на каждую точку регулярной сетки.
Public Sub Initialize(ByVal region As CConcreteRegion)
    Set mRegion = Nothing: Set mQuery = Nothing: Erase mEdges: mEdgeCount = 0
    mMinX = 0#: mMaxX = 0#: mMinY = 0#: mMaxY = 0#
    If region Is Nothing Then Err.Raise vbObjectError + 5290, "CGeometryRegion", "Точный контур бетона не передан."
    If region.SegmentCount = 0 Or region.Area <= GEOM_TOLERANCE Then _
        Err.Raise vbObjectError + 5290, "CGeometryRegion", "Точный контур бетона пуст или имеет нулевую площадь."
    Set mRegion = region.Clone
    Set mQuery = New CSectionGeometryQuery
    mRegion.CopyEdges mEdges, mEdgeCount
    mQuery.ProjectionBounds mRegion, 1#, 0#, mMinX, mMaxX
    mQuery.ProjectionBounds mRegion, 0#, 1#, mMinY, mMaxY
End Sub

Public Property Get MinX() As Double
    MinX = mMinX
End Property
Public Property Get MaxX() As Double
    MaxX = mMaxX
End Property
Public Property Get MinY() As Double
    MinY = mMinY
End Property
Public Property Get MaxY() As Double
    MaxY = mMaxY
End Property

' Как у параметрических форм, наружная граница входит в бетон,
' а граница отверстия исключается. Бетонные части внутри проемов сохраняются.
Public Function ContainsPoint(ByVal X As Double, ByVal Y As Double) As Boolean
    If mRegion Is Nothing Then Exit Function
    ContainsPoint = mQuery.ContainsMeshPoint(mRegion, mEdges, mEdgeCount, X, Y)
End Function

Public Function IsPointInside(ByVal X As Double, ByVal Y As Double) As Boolean
    IsPointInside = ContainsPoint(X, Y)
End Function

Public Function IsValid(ByRef message As String) As Boolean
    message = vbNullString
    If mRegion Is Nothing Then
        message = "Точный контур бетона не инициализирован."
    ElseIf mMaxX <= mMinX + GEOM_TOLERANCE Or mMaxY <= mMinY + GEOM_TOLERANCE Then
        message = "Габариты точного контура бетона вырождены."
    Else
        IsValid = True
    End If
End Function

Public Function AnalyticalArea(ByRef available As Boolean) As Double
    available = Not mRegion Is Nothing
    If available Then AnalyticalArea = mRegion.Area
End Function

' Аналитический центр query не вычисляет. Потребитель получает явный
' признак недоступности и использует обычные свойства сформированной сетки.
Public Sub AnalyticalCentroid(ByRef available As Boolean, ByRef X As Double, ByRef Y As Double)
    available = False: X = 0#: Y = 0#
End Sub

Public Sub GetExtremePoints(ByRef pointX() As Double, ByRef pointY() As Double)
    If mRegion Is Nothing Then Err.Raise vbObjectError + 5290, "CGeometryRegion", "Точный контур бетона не инициализирован."
    mQuery.MeshExtremePoints mEdges, mEdgeCount, pointX, pointY
End Sub

Public Sub BuildContours(ByVal contours As CSectionContours)
    If contours Is Nothing Then Err.Raise vbObjectError + 5290, "CGeometryRegion", "Контейнер контуров не передан."
    Dim message As String
    If Not IsValid(message) Then Err.Raise vbObjectError + 5290, "CGeometryRegion", message
    contours.Clear
    contours.AddMaterialRegion mRegion, "REGION", "Достоверный контур бетонной области."
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
Private Sub ISectionGeometry_BuildContours(ByVal contours As CSectionContours)
    BuildContours contours
End Sub
