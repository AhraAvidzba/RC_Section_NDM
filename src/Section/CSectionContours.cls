VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionContours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Достоверные наружные границы и отверстия геометрии сечения
' ==========================================================================
' Контейнер принадлежит CSectionModel и хранит точные линии, дуги и окружности
' во внутренних мм. Роль и принадлежность замкнутому контуру являются геометрическими
' данными, а не аннотацией результата. Контейнер не вычисляет расчетную область
' трещин: она остается отдельным LC-зависимым результатом geometry-query.

Private mKind() As String
Private mContourID() As String
Private mEdge() As TRegionEdge
Private mRadius() As Double
Private mLoopID() As String
Private mRole() As String
Private mSource() As String
Private mComment() As String
Private mCount As Long
Private mRevision As Long ' Версия исходных границ для инвалидирования geometry-query.

' Удаляет исходные границы, сохраняя монотонную версию контейнера.
Public Sub Clear()
    Erase mKind: Erase mContourID: Erase mEdge: Erase mRadius
    Erase mLoopID: Erase mRole: Erase mSource: Erase mComment
    mCount = 0
    mRevision = mRevision + 1
End Sub

' Добавляет прямой сегмент известного наружного контура или отверстия.
Public Sub AddContourLine(ByVal contourID As String, ByVal startX As Double, ByVal startY As Double, _
        ByVal endX As Double, ByVal endY As Double, Optional ByVal commentValue As String = "Сегмент контура сечения.", _
        Optional ByVal loopID As String = "OUTER", Optional ByVal loopRole As String = "Outer", _
        Optional ByVal sourceID As String = "Generated")
    AddSegment "CONTOUR_LINE", contourID, startX, startY, endX, endY, 0#, 0#, 0#, 0#, loopID, loopRole, sourceID, commentValue
End Sub

' Добавляет окружность без замены аналитической границы многоугольником.
Public Sub AddContourCircle(ByVal contourID As String, ByVal centerX As Double, ByVal centerY As Double, _
        ByVal radius As Double, Optional ByVal commentValue As String = "Круговой контур сечения.", _
        Optional ByVal loopID As String = "OUTER", Optional ByVal loopRole As String = "Outer", _
        Optional ByVal sourceID As String = "Generated")
    If radius <= 0# Then Err.Raise vbObjectError + 5284, "CSectionContours", "Радиус контура должен быть положительным."
    AddSegment "CONTOUR_CIRCLE", contourID, centerX, centerY, radius, 0#, centerX, centerY, radius, 0#, loopID, loopRole, sourceID, commentValue
End Sub

' Сохраняет фактическую дугу с центром, радиусом и знаковым углом в радианах.
Public Sub AddContourArc(ByVal contourID As String, ByVal startX As Double, ByVal startY As Double, _
        ByVal endX As Double, ByVal endY As Double, ByVal centerX As Double, ByVal centerY As Double, _
        ByVal radius As Double, ByVal sweepAngle As Double, Optional ByVal commentValue As String = "Дуга контура сечения.", _
        Optional ByVal loopID As String = "OUTER", Optional ByVal loopRole As String = "Outer", _
        Optional ByVal sourceID As String = "Generated")
    If radius <= 0# Then Err.Raise vbObjectError + 5284, "CSectionContours", "Радиус дуги контура должен быть положительным."
    AddSegment "CONTOUR_ARC", contourID, startX, startY, endX, endY, centerX, centerY, radius, sweepAngle, loopID, loopRole, sourceID, commentValue
End Sub

' Переносит проверенные замкнутые контуры импортированной области без изменения примитивов.
' Каждое отверстие имеет собственный ключ; внешняя Region отверстия меняет роль.
Public Sub AddMaterialRegion(ByVal region As CConcreteRegion, ByVal sourceID As String, ByVal commentValue As String, _
        Optional ByVal asOpening As Boolean = False)
    If region Is Nothing Then Err.Raise vbObjectError + 5283, "CSectionContours", "Область материального контура не передана."
    Dim i As Long, loopNumber As Long, role As String, key As String, contourID As String
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, cx As Double, cy As Double, sweep As Double, radius As Double
    For i = 1 To region.SegmentCount
        region.GetSegment i, x1, y1, x2, y2, cx, cy, sweep, loopNumber
        role = "Outer"
        If region.LoopIsOpening(loopNumber) Xor asOpening Then role = "Opening"
        key = sourceID & "_" & CStr(loopNumber)
        contourID = "CAD_" & UCase$(role) & "_" & key & "_" & CStr(i)
        If Abs(sweep) <= 0.0000000001 Then
            AddContourLine contourID, x1, y1, x2, y2, commentValue, key, role, sourceID
        Else
            radius = Sqr((x1 - cx) ^ 2 + (y1 - cy) ^ 2)
            AddContourArc contourID, x1, y1, x2, y2, cx, cy, radius, sweep, commentValue, key, role, sourceID
        End If
    Next i
End Sub

' Записывает геометрическую запись и увеличивает версию только после проверки metadata.
Private Sub AddSegment(ByVal kind As String, ByVal contourID As String, ByVal x1 As Double, ByVal y1 As Double, _
        ByVal x2 As Double, ByVal y2 As Double, ByVal cx As Double, ByVal cy As Double, ByVal radius As Double, _
        ByVal sweep As Double, ByVal loopID As String, ByVal role As String, ByVal source As String, ByVal comment As String)
    If Len(Trim$(contourID)) = 0 Or Len(Trim$(loopID)) = 0 Then Err.Raise vbObjectError + 5281, "CSectionContours", "Не заданы ID сегмента или замкнутого контура."
    If StrComp(role, "Outer", vbTextCompare) <> 0 And StrComp(role, "Opening", vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 5281, "CSectionContours", "Роль замкнутого контура должна быть Outer (наружная граница) или Opening (граница отверстия)."
    mCount = mCount + 1
    ReDim Preserve mKind(1 To mCount): ReDim Preserve mContourID(1 To mCount)
    ReDim Preserve mEdge(1 To mCount): ReDim Preserve mRadius(1 To mCount)
    ReDim Preserve mLoopID(1 To mCount): ReDim Preserve mRole(1 To mCount)
    ReDim Preserve mSource(1 To mCount): ReDim Preserve mComment(1 To mCount)
    mKind(mCount) = kind: mContourID(mCount) = contourID
    With mEdge(mCount)
        .X1 = x1: .Y1 = y1: .X2 = x2: .Y2 = y2
        .CenterX = cx: .CenterY = cy: .Sweep = sweep
    End With
    mRadius(mCount) = radius: mLoopID(mCount) = loopID: mRole(mCount) = role
    mSource(mCount) = source: mComment(mCount) = comment
    mRevision = mRevision + 1
End Sub

' Чтение сохраненных геометрических полей не изменяет контейнер или его версию.
Public Property Get Count() As Long
    Count = mCount
End Property
Public Property Get Revision() As Long
    Revision = mRevision
End Property
Public Property Get SegmentType(ByVal index As Long) As String
    CheckIndex index: SegmentType = mKind(index)
End Property
Public Property Get ContourID(ByVal index As Long) As String
    CheckIndex index: ContourID = mContourID(index)
End Property
Public Property Get StartX(ByVal index As Long) As Double
    CheckIndex index: StartX = mEdge(index).X1
End Property
Public Property Get StartY(ByVal index As Long) As Double
    CheckIndex index: StartY = mEdge(index).Y1
End Property
Public Property Get EndX(ByVal index As Long) As Double
    CheckIndex index: EndX = mEdge(index).X2
End Property
Public Property Get EndY(ByVal index As Long) As Double
    CheckIndex index: EndY = mEdge(index).Y2
End Property
Public Property Get CenterX(ByVal index As Long) As Double
    CheckIndex index: CenterX = mEdge(index).CenterX
End Property
Public Property Get CenterY(ByVal index As Long) As Double
    CheckIndex index: CenterY = mEdge(index).CenterY
End Property
Public Property Get Radius(ByVal index As Long) As Double
    CheckIndex index: Radius = mRadius(index)
End Property
Public Property Get SweepAngle(ByVal index As Long) As Double
    CheckIndex index: SweepAngle = mEdge(index).Sweep
End Property
Public Property Get LoopID(ByVal index As Long) As String
    CheckIndex index: LoopID = mLoopID(index)
End Property
Public Property Get LoopRole(ByVal index As Long) As String
    CheckIndex index: LoopRole = mRole(index)
End Property
Public Property Get SourceID(ByVal index As Long) As String
    CheckIndex index: SourceID = mSource(index)
End Property
Public Property Get Comment(ByVal index As Long) As String
    CheckIndex index: Comment = mComment(index)
End Property

' Проверяет индекс перед доступом к геометрическим массивам.
Private Sub CheckIndex(ByVal index As Long)
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 5282, "CSectionContours", "Индекс сегмента контура вне допустимого диапазона."
End Sub
