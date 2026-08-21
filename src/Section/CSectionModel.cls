VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSourceType As String

Private mConcreteID() As String
Private mConcreteSourceName() As String
Private mConcreteSourceHandle() As String
Private mConcreteX() As Double
Private mConcreteY() As Double
Private mConcreteArea() As Double
Private mConcreteMaterialID() As Long
Private mConcreteShapeType() As String
Private mConcreteWidth() As Double
Private mConcreteHeight() As Double
Private mConcreteRotation() As Double
Private mConcreteLocalIx() As Double
Private mConcreteLocalIy() As Double
Private mConcreteLocalIxy() As Double
Private mConcreteComment() As String
Private mConcreteCount As Long

Private mRebarID() As String
Private mRebarSourceName() As String
Private mRebarSourceHandle() As String
Private mRebarX() As Double
Private mRebarY() As Double
Private mRebarArea() As Double
Private mRebarDiameter() As Double
Private mRebarSteelClass() As String
Private mRebarMaterialID() As Long
Private mRebarComment() As String
Private mRebarCount As Long

Private mRebarAnnotationGroupName() As String
Private mRebarAnnotationStartX() As Double
Private mRebarAnnotationStartY() As Double
Private mRebarAnnotationEndX() As Double
Private mRebarAnnotationEndY() As Double
Private mRebarAnnotationNormalX() As Double
Private mRebarAnnotationNormalY() As Double
Private mRebarAnnotationAxisDistance() As Double
Private mRebarAnnotationCount As Long

Public Sub Clear()
    mSourceType = vbNullString

    Erase mConcreteID
    Erase mConcreteSourceName
    Erase mConcreteSourceHandle
    Erase mConcreteX
    Erase mConcreteY
    Erase mConcreteArea
    Erase mConcreteMaterialID
    Erase mConcreteShapeType
    Erase mConcreteWidth
    Erase mConcreteHeight
    Erase mConcreteRotation
    Erase mConcreteLocalIx
    Erase mConcreteLocalIy
    Erase mConcreteLocalIxy
    Erase mConcreteComment
    mConcreteCount = 0

    Erase mRebarID
    Erase mRebarSourceName
    Erase mRebarSourceHandle
    Erase mRebarX
    Erase mRebarY
    Erase mRebarArea
    Erase mRebarDiameter
    Erase mRebarSteelClass
    Erase mRebarMaterialID
    Erase mRebarComment
    mRebarCount = 0

    Erase mRebarAnnotationGroupName
    Erase mRebarAnnotationStartX
    Erase mRebarAnnotationStartY
    Erase mRebarAnnotationEndX
    Erase mRebarAnnotationEndY
    Erase mRebarAnnotationNormalX
    Erase mRebarAnnotationNormalY
    Erase mRebarAnnotationAxisDistance
    mRebarAnnotationCount = 0
End Sub

Public Property Let SourceType(ByVal value As String)
    mSourceType = Trim$(value)
End Property

Public Property Get SourceType() As String
    SourceType = mSourceType
End Property

Public Sub AddConcreteElement(ByVal xCoord As Double, ByVal yCoord As Double, ByVal area As Double, _
        Optional ByVal materialID As Long = 1, Optional ByVal sourceName As String = vbNullString, _
        Optional ByVal sourceHandle As String = vbNullString, Optional ByVal shapeType As String = "Rectangle", _
        Optional ByVal width As Double = 0#, Optional ByVal height As Double = 0#, _
        Optional ByVal rotation As Double = 0#, Optional ByVal comment As String = vbNullString, _
        Optional ByVal localIx As Double = 0#, Optional ByVal localIy As Double = 0#, _
        Optional ByVal localIxy As Double = 0#)
    If area <= 0# Then Err.Raise vbObjectError + 5200, "CSectionModel", "Площадь бетонного элемента должна быть положительной."

    mConcreteCount = mConcreteCount + 1
    ReDim Preserve mConcreteID(1 To mConcreteCount)
    ReDim Preserve mConcreteSourceName(1 To mConcreteCount)
    ReDim Preserve mConcreteSourceHandle(1 To mConcreteCount)
    ReDim Preserve mConcreteX(1 To mConcreteCount)
    ReDim Preserve mConcreteY(1 To mConcreteCount)
    ReDim Preserve mConcreteArea(1 To mConcreteCount)
    ReDim Preserve mConcreteMaterialID(1 To mConcreteCount)
    ReDim Preserve mConcreteShapeType(1 To mConcreteCount)
    ReDim Preserve mConcreteWidth(1 To mConcreteCount)
    ReDim Preserve mConcreteHeight(1 To mConcreteCount)
    ReDim Preserve mConcreteRotation(1 To mConcreteCount)
    ReDim Preserve mConcreteLocalIx(1 To mConcreteCount)
    ReDim Preserve mConcreteLocalIy(1 To mConcreteCount)
    ReDim Preserve mConcreteLocalIxy(1 To mConcreteCount)
    ReDim Preserve mConcreteComment(1 To mConcreteCount)

    mConcreteID(mConcreteCount) = "C" & CStr(mConcreteCount)
    mConcreteSourceName(mConcreteCount) = sourceName
    mConcreteSourceHandle(mConcreteCount) = sourceHandle
    mConcreteX(mConcreteCount) = xCoord
    mConcreteY(mConcreteCount) = yCoord
    mConcreteArea(mConcreteCount) = area
    mConcreteMaterialID(mConcreteCount) = materialID
    mConcreteShapeType(mConcreteCount) = shapeType
    mConcreteWidth(mConcreteCount) = width
    mConcreteHeight(mConcreteCount) = height
    mConcreteRotation(mConcreteCount) = rotation
    mConcreteLocalIx(mConcreteCount) = localIx
    mConcreteLocalIy(mConcreteCount) = localIy
    mConcreteLocalIxy(mConcreteCount) = localIxy
    mConcreteComment(mConcreteCount) = comment
End Sub

Public Sub AddRebarElement(ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal diameter As Double, ByVal area As Double, ByVal steelClass As String, _
        Optional ByVal materialID As Long = 1, Optional ByVal sourceName As String = vbNullString, _
        Optional ByVal sourceHandle As String = vbNullString, Optional ByVal comment As String = vbNullString)
    If diameter <= 0# Then Err.Raise vbObjectError + 5210, "CSectionModel", "Диаметр арматуры должен быть положительным."

    Dim actualArea As Double
    If area <= 0# Then
        actualArea = GEOM_PI * diameter * diameter / 4#
    Else
        actualArea = area
    End If
    If actualArea <= 0# Then Err.Raise vbObjectError + 5211, "CSectionModel", "Площадь арматуры должна быть положительной."

    mRebarCount = mRebarCount + 1
    ReDim Preserve mRebarID(1 To mRebarCount)
    ReDim Preserve mRebarSourceName(1 To mRebarCount)
    ReDim Preserve mRebarSourceHandle(1 To mRebarCount)
    ReDim Preserve mRebarX(1 To mRebarCount)
    ReDim Preserve mRebarY(1 To mRebarCount)
    ReDim Preserve mRebarArea(1 To mRebarCount)
    ReDim Preserve mRebarDiameter(1 To mRebarCount)
    ReDim Preserve mRebarSteelClass(1 To mRebarCount)
    ReDim Preserve mRebarMaterialID(1 To mRebarCount)
    ReDim Preserve mRebarComment(1 To mRebarCount)

    mRebarID(mRebarCount) = "R" & CStr(mRebarCount)
    mRebarSourceName(mRebarCount) = sourceName
    mRebarSourceHandle(mRebarCount) = sourceHandle
    mRebarX(mRebarCount) = xCoord
    mRebarY(mRebarCount) = yCoord
    mRebarArea(mRebarCount) = actualArea
    mRebarDiameter(mRebarCount) = diameter
    mRebarSteelClass(mRebarCount) = steelClass
    mRebarMaterialID(mRebarCount) = materialID
    mRebarComment(mRebarCount) = comment
End Sub

Public Property Get ConcreteCount() As Long
    ConcreteCount = mConcreteCount
End Property

Public Property Get RebarCount() As Long
    RebarCount = mRebarCount
End Property

Public Property Get ConcreteID(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteID = mConcreteID(index)
End Property

Public Property Get ConcreteSourceName(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteSourceName = mConcreteSourceName(index)
End Property

Public Property Get ConcreteSourceHandle(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteSourceHandle = mConcreteSourceHandle(index)
End Property

Public Property Get ConcreteX(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteX = mConcreteX(index)
End Property

Public Property Get ConcreteY(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteY = mConcreteY(index)
End Property

Public Property Get ConcreteArea(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteArea = mConcreteArea(index)
End Property

Public Property Get ConcreteMaterialID(ByVal index As Long) As Long
    CheckConcreteIndex index
    ConcreteMaterialID = mConcreteMaterialID(index)
End Property

Public Property Get ConcreteShapeType(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteShapeType = mConcreteShapeType(index)
End Property

Public Property Get ConcreteWidth(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteWidth = mConcreteWidth(index)
End Property

Public Property Get ConcreteHeight(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteHeight = mConcreteHeight(index)
End Property

Public Property Get ConcreteRotation(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteRotation = mConcreteRotation(index)
End Property

Public Property Get ConcreteLocalIx(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteLocalIx = mConcreteLocalIx(index)
End Property

Public Property Get ConcreteLocalIy(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteLocalIy = mConcreteLocalIy(index)
End Property

Public Property Get ConcreteLocalIxy(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteLocalIxy = mConcreteLocalIxy(index)
End Property

Public Property Get ConcreteComment(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteComment = mConcreteComment(index)
End Property

Public Property Get RebarID(ByVal index As Long) As String
    CheckRebarIndex index
    RebarID = mRebarID(index)
End Property

Public Property Get RebarSourceName(ByVal index As Long) As String
    CheckRebarIndex index
    RebarSourceName = mRebarSourceName(index)
End Property

Public Property Get RebarSourceHandle(ByVal index As Long) As String
    CheckRebarIndex index
    RebarSourceHandle = mRebarSourceHandle(index)
End Property

Public Property Get RebarX(ByVal index As Long) As Double
    CheckRebarIndex index
    RebarX = mRebarX(index)
End Property

Public Property Get RebarY(ByVal index As Long) As Double
    CheckRebarIndex index
    RebarY = mRebarY(index)
End Property

Public Property Get RebarArea(ByVal index As Long) As Double
    CheckRebarIndex index
    RebarArea = mRebarArea(index)
End Property

Public Property Get RebarDiameter(ByVal index As Long) As Double
    CheckRebarIndex index
    RebarDiameter = mRebarDiameter(index)
End Property

Public Property Get RebarSteelClass(ByVal index As Long) As String
    CheckRebarIndex index
    RebarSteelClass = mRebarSteelClass(index)
End Property

Public Property Get RebarMaterialID(ByVal index As Long) As Long
    CheckRebarIndex index
    RebarMaterialID = mRebarMaterialID(index)
End Property

Public Property Get RebarComment(ByVal index As Long) As String
    CheckRebarIndex index
    RebarComment = mRebarComment(index)
End Property

Public Sub AddRebarAnnotationAnchor(ByVal groupName As String, ByVal startX As Double, ByVal startY As Double, _
        ByVal endX As Double, ByVal endY As Double, ByVal normalX As Double, ByVal normalY As Double, _
        Optional ByVal axisDistance As Double = 0#)
    If Len(Trim$(groupName)) = 0 Then Exit Sub

    mRebarAnnotationCount = mRebarAnnotationCount + 1
    ReDim Preserve mRebarAnnotationGroupName(1 To mRebarAnnotationCount)
    ReDim Preserve mRebarAnnotationStartX(1 To mRebarAnnotationCount)
    ReDim Preserve mRebarAnnotationStartY(1 To mRebarAnnotationCount)
    ReDim Preserve mRebarAnnotationEndX(1 To mRebarAnnotationCount)
    ReDim Preserve mRebarAnnotationEndY(1 To mRebarAnnotationCount)
    ReDim Preserve mRebarAnnotationNormalX(1 To mRebarAnnotationCount)
    ReDim Preserve mRebarAnnotationNormalY(1 To mRebarAnnotationCount)
    ReDim Preserve mRebarAnnotationAxisDistance(1 To mRebarAnnotationCount)

    mRebarAnnotationGroupName(mRebarAnnotationCount) = Trim$(groupName)
    mRebarAnnotationStartX(mRebarAnnotationCount) = startX
    mRebarAnnotationStartY(mRebarAnnotationCount) = startY
    mRebarAnnotationEndX(mRebarAnnotationCount) = endX
    mRebarAnnotationEndY(mRebarAnnotationCount) = endY
    mRebarAnnotationNormalX(mRebarAnnotationCount) = normalX
    mRebarAnnotationNormalY(mRebarAnnotationCount) = normalY
    mRebarAnnotationAxisDistance(mRebarAnnotationCount) = axisDistance
End Sub

Public Property Get RebarAnnotationCount() As Long
    RebarAnnotationCount = mRebarAnnotationCount
End Property

Public Property Get RebarAnnotationGroupName(ByVal index As Long) As String
    CheckRebarAnnotationIndex index
    RebarAnnotationGroupName = mRebarAnnotationGroupName(index)
End Property

Public Property Get RebarAnnotationStartX(ByVal index As Long) As Double
    CheckRebarAnnotationIndex index
    RebarAnnotationStartX = mRebarAnnotationStartX(index)
End Property

Public Property Get RebarAnnotationStartY(ByVal index As Long) As Double
    CheckRebarAnnotationIndex index
    RebarAnnotationStartY = mRebarAnnotationStartY(index)
End Property

Public Property Get RebarAnnotationEndX(ByVal index As Long) As Double
    CheckRebarAnnotationIndex index
    RebarAnnotationEndX = mRebarAnnotationEndX(index)
End Property

Public Property Get RebarAnnotationEndY(ByVal index As Long) As Double
    CheckRebarAnnotationIndex index
    RebarAnnotationEndY = mRebarAnnotationEndY(index)
End Property

Public Property Get RebarAnnotationNormalX(ByVal index As Long) As Double
    CheckRebarAnnotationIndex index
    RebarAnnotationNormalX = mRebarAnnotationNormalX(index)
End Property

Public Property Get RebarAnnotationNormalY(ByVal index As Long) As Double
    CheckRebarAnnotationIndex index
    RebarAnnotationNormalY = mRebarAnnotationNormalY(index)
End Property

Public Property Get RebarAnnotationAxisDistance(ByVal index As Long) As Double
    CheckRebarAnnotationIndex index
    RebarAnnotationAxisDistance = mRebarAnnotationAxisDistance(index)
End Property

Private Sub CheckConcreteIndex(ByVal index As Long)
    If index < 1 Or index > mConcreteCount Then Err.Raise vbObjectError + 5220, "CSectionModel", "Индекс бетонного элемента вне диапазона."
End Sub

Private Sub CheckRebarIndex(ByVal index As Long)
    If index < 1 Or index > mRebarCount Then Err.Raise vbObjectError + 5221, "CSectionModel", "Индекс стержня вне диапазона."
End Sub

Private Sub CheckRebarAnnotationIndex(ByVal index As Long)
    If index < 1 Or index > mRebarAnnotationCount Then Err.Raise vbObjectError + 5222, "CSectionModel", "Индекс аннотации арматуры вне диапазона."
End Sub
