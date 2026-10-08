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
' Единая подготовленная бетонная область для сетки и проверок арматуры.
' Точные отрезки, дуги и отверстия обрабатывает CSectionGeometryQuery.
' Параметрические формы и импорт передают сюда одинаковые данные.

Private mRegion As CConcreteRegion
Private mQuery As CSectionGeometryQuery
Private mEdges() As TRegionEdge
Private mEdgeCount As Long
Private mMinX As Double, mMaxX As Double, mMinY As Double, mMaxY As Double
Private mCentroidX As Double, mCentroidY As Double

' Снимок исключает изменение формы во время генерации; ребра копируются один
' раз, а не на каждую точку регулярной сетки.
Public Sub Initialize(ByVal region As CConcreteRegion)
    Set mRegion = Nothing: Set mQuery = Nothing: Erase mEdges: mEdgeCount = 0
    mMinX = 0#: mMaxX = 0#: mMinY = 0#: mMaxY = 0#
    mCentroidX = 0#: mCentroidY = 0#
    If region Is Nothing Then Err.Raise vbObjectError + 5290, "CGeometryRegion", "Точный контур бетона не передан."
    If region.SegmentCount = 0 Or region.Area <= GEOM_TOLERANCE Then _
        Err.Raise vbObjectError + 5290, "CGeometryRegion", "Точный контур бетона пуст или имеет нулевую площадь."
    Set mRegion = region.Clone
    Set mQuery = New CSectionGeometryQuery
    mRegion.CopyEdges mEdges, mEdgeCount
    mQuery.ProjectionBounds mRegion, 1#, 0#, mMinX, mMaxX
    mQuery.ProjectionBounds mRegion, 0#, 1#, mMinY, mMaxY
    mQuery.RegionCentroid mRegion, mCentroidX, mCentroidY
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
Public Property Get Width() As Double
    Width = mMaxX - mMinX
End Property
Public Property Get Height() As Double
    Height = mMaxY - mMinY
End Property

' Как у параметрических форм, наружная граница входит в бетон,
' а граница отверстия исключается. Бетонные части внутри проемов сохраняются.
Public Function ContainsPoint(ByVal X As Double, ByVal Y As Double) As Boolean
    If mRegion Is Nothing Then Exit Function
    ContainsPoint = mQuery.ContainsMeshPoint(mRegion, mEdges, mEdgeCount, X, Y)
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

' Площадь и центр относятся к точным границам, а не к бетонной сетке.
Public Sub AnalyticalCentroid(ByRef available As Boolean, ByRef X As Double, ByRef Y As Double)
    available = Not mRegion Is Nothing
    X = mCentroidX: Y = mCentroidY
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
