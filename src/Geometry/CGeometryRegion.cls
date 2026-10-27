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
' Бетонная область для построения сетки и проверки арматуры
' ==========================================================================
' Предоставляет общую область бетона по точным наружным контурам и отверстиям.
' Геометрические операции с отрезками и дугами выполняет CSectionGeometryQuery.
' Параметрические сечения и импорт передают сюда данные одинакового вида.

Private mRegion As CConcreteRegion ' Отдельная копия бетонной области для построения сетки.
Private mQuery As CSectionGeometryQuery ' Общие геометрические проверки и вычисления по точным границам.
Private mEdges() As TRegionEdge ' Подготовленные отрезки и дуги для проверки точек сетки без повторного копирования.
Private mEdgeCount As Long ' Количество отрезков и дуг в подготовленном массиве границ.
Private mMinX As Double, mMaxX As Double, mMinY As Double, mMaxY As Double ' Крайние координаты области по X и Y, мм.
Private mCentroidX As Double, mCentroidY As Double ' Координаты центра тяжести точной бетонной области, мм.

' Копирует границы один раз перед построением сетки, чтобы дальнейшие изменения
' исходной формы не влияли на расчёт. При проверке каждой точки сетки границы не копируются.
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
' Минимальная координата X геометрической области, мм.
Public Property Get MinX() As Double
    MinX = mMinX
End Property
' Максимальная координата X геометрической области, мм.
Public Property Get MaxX() As Double
    MaxX = mMaxX
End Property
' Минимальная координата Y геометрической области, мм.
Public Property Get MinY() As Double
    MinY = mMinY
End Property
' Максимальная координата Y геометрической области, мм.
Public Property Get MaxY() As Double
    MaxY = mMaxY
End Property
' Полная ширина геометрической области по X, мм.
Public Property Get Width() As Double
    Width = mMaxX - mMinX
End Property
' Полная высота геометрической области по Y, мм.
Public Property Get Height() As Double
    Height = mMaxY - mMinY
End Property

' Точка на наружном контуре считается принадлежащей бетону.
' Точка на границе отверстия бетону не принадлежит. Если внутри отверстия
' задан отдельный бетонный участок, он проверяется по своему контуру.
Public Function ContainsPoint(ByVal X As Double, ByVal Y As Double) As Boolean
    If mRegion Is Nothing Then Exit Function
    ContainsPoint = mQuery.ContainsMeshPoint(mRegion, mEdges, mEdgeCount, X, Y)
End Function


' Проверяет все подготовленные границы, включая отверстие целиком внутри
' ячейки и отдельный остров бетона. Касание стороны без захода внутрь не дробит.
Friend Function MeshBoundaryCrossesCell(ByVal minX As Double, ByVal minY As Double, _
        ByVal maxX As Double, ByVal maxY As Double) As Boolean
    If mRegion Is Nothing Then Exit Function
    MeshBoundaryCrossesCell = mQuery.MeshBoundaryCrossesCell(mEdges, mEdgeCount, minX, minY, maxX, maxY)
End Function

' Возвращает все связные части пересечения бетонной области с исходной ячейкой.
' Отверстия и дуги сохраняются аналитически; центр ячейки не определяет отбор.
Friend Function ClipCell(ByVal minX As Double, ByVal minY As Double, _
        ByVal maxX As Double, ByVal maxY As Double) As Collection
    If mRegion Is Nothing Then Err.Raise vbObjectError + 5290, "CGeometryRegion", "Точный контур бетонной области ещё не подготовлен."
    If maxX <= minX Or maxY <= minY Then Err.Raise vbObjectError + 5290, "CGeometryRegion", "Прямоугольная ячейка должна иметь положительные размеры."
    Dim clipped As CConcreteRegion
    Set clipped = mQuery.ClipWindow(mRegion, 1#, 0#, minX, maxX, minY, maxY)
    Set ClipCell = mQuery.ConnectedRegions(clipped)
End Function

' Проверяет, создана ли область и превышают ли её ширина и высота геометрический допуск.
' При ошибке возвращает False и понятную причину в message.
Public Function IsValid(ByRef message As String) As Boolean
    message = vbNullString
    If mRegion Is Nothing Then
        message = "Точный контур бетонной области ещё не подготовлен."
    ElseIf mMaxX <= mMinX + GEOM_TOLERANCE Or mMaxY <= mMinY + GEOM_TOLERANCE Then
        message = "Ширина или высота наружного контура равна нулю. Проверьте координаты и размеры бетонной области."
    Else
        IsValid = True
    End If
End Function

' Возвращает площадь по аналитическим контурам с учётом отверстий, мм^2.
Public Function AnalyticalArea(ByRef available As Boolean) As Double
    available = Not mRegion Is Nothing
    If available Then AnalyticalArea = mRegion.Area
End Function

' Площадь и центр относятся к точным границам, а не к бетонной сетке.
Public Sub AnalyticalCentroid(ByRef available As Boolean, ByRef X As Double, ByRef Y As Double)
    available = Not mRegion Is Nothing
    X = mCentroidX: Y = mCentroidY
End Sub

' Возвращает вершины границы и крайние точки дуг по осям X/Y.
' Для произвольного направления точные габариты вычисляет ProjectionBounds общего запроса.
Public Sub GetExtremePoints(ByRef pointX() As Double, ByRef pointY() As Double)
    If mRegion Is Nothing Then Err.Raise vbObjectError + 5290, "CGeometryRegion", "Точный контур бетонной области ещё не подготовлен."
    mQuery.MeshExtremePoints mEdges, mEdgeCount, pointX, pointY
End Sub

' Передаёт точные наружные контуры и отверстия в общий контейнер границ модели.
Public Sub BuildContours(ByVal contours As CSectionContours)
    If contours Is Nothing Then Err.Raise vbObjectError + 5290, "CGeometryRegion", "Контейнер контуров не передан."
    Dim message As String
    If Not IsValid(message) Then Err.Raise vbObjectError + 5290, "CGeometryRegion", message
    contours.Clear
    contours.AddMaterialRegion mRegion, "REGION", "Достоверный контур бетонной области."
End Sub
