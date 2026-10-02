VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRebarLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Промежуточный список стержней до CSectionModel
' ==========================================================================
' CRebarLayout хранит координаты, диаметры и диагностические подписи стержней,
' сформированных раскладчиком. Затем CSectionModelBuilder переносит эти данные
' в единую расчетную модель и присваивает расчетные имена R1/R2/....

Private mBarID() As String
Private mX() As Double
Private mY() As Double
Private mDiameter() As Double
Private mArea() As Double
Private mSteelClass() As String
Private mComment() As String
Private mBarAnnotationGroupName() As String
Private mCount As Long

Private mAnnotationGroupName() As String
Private mAnnotationStartX() As Double
Private mAnnotationStartY() As Double
Private mAnnotationEndX() As Double
Private mAnnotationEndY() As Double
Private mAnnotationNormalX() As Double
Private mAnnotationNormalY() As Double
Private mAnnotationAxisDistance() As Double
Private mAnnotationCount As Long

' Удаляет промежуточные стержни и аннотации линий раскладки целиком.
' Оба счетчика сбрасываются вместе с массивами, чтобы следующая генерация
' не смешивала координаты или подписи с предыдущей геометрией.
Public Sub Clear()
    Erase mBarID
    Erase mX
    Erase mY
    Erase mDiameter
    Erase mArea
    Erase mSteelClass
    Erase mComment
    Erase mBarAnnotationGroupName
    mCount = 0

    Erase mAnnotationGroupName
    Erase mAnnotationStartX
    Erase mAnnotationStartY
    Erase mAnnotationEndX
    Erase mAnnotationEndY
    Erase mAnnotationNormalX
    Erase mAnnotationNormalY
    Erase mAnnotationAxisDistance
    mAnnotationCount = 0
End Sub

' Добавляет стержень в промежуточную раскладку с проверкой ID, площади, диаметра и положения в бетоне.
Public Sub AddBar(ByVal barName As String, ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal diameter As Double, ByVal area As Double, ByVal steelClass As String, _
        Optional ByVal comment As String = vbNullString, Optional ByVal geometry As ISectionGeometry = Nothing, _
        Optional ByVal annotationGroupName As String = vbNullString)

    If Len(Trim$(barName)) = 0 Then Err.Raise vbObjectError + 3100, "CRebarLayout", "ID стержня не заполнен."
    If diameter <= 0# Then Err.Raise vbObjectError + 3101, "CRebarLayout", "Диаметр арматуры должен быть положительным."

    Dim actualArea As Double
    If area <= 0# Then
        actualArea = GEOM_PI * diameter * diameter / 4#
    Else
        actualArea = area
    End If
    If actualArea <= 0# Then Err.Raise vbObjectError + 3102, "CRebarLayout", "Площадь арматуры должна быть положительной."

    Dim i As Long
    For i = 1 To mCount
        If StrComp(mBarID(i), barName, vbTextCompare) = 0 Then Err.Raise vbObjectError + 3103, "CRebarLayout", "Дублирующийся ID стержня: " & barName
        If Abs(mX(i) - xCoord) <= GEOM_TOLERANCE And Abs(mY(i) - yCoord) <= GEOM_TOLERANCE Then
            Err.Raise vbObjectError + 3104, "CRebarLayout", "Дублирующиеся координаты стержней."
        End If
    Next i

    If Not geometry Is Nothing Then
        If Not geometry.ContainsPoint(xCoord, yCoord) Then Err.Raise vbObjectError + 3105, "CRebarLayout", "Центр стержня находится вне бетонного сечения."
    End If

    mCount = mCount + 1
    ReDim Preserve mBarID(1 To mCount)
    ReDim Preserve mX(1 To mCount)
    ReDim Preserve mY(1 To mCount)
    ReDim Preserve mDiameter(1 To mCount)
    ReDim Preserve mArea(1 To mCount)
    ReDim Preserve mSteelClass(1 To mCount)
    ReDim Preserve mComment(1 To mCount)
    ReDim Preserve mBarAnnotationGroupName(1 To mCount)

    mBarID(mCount) = barName
    mX(mCount) = xCoord
    mY(mCount) = yCoord
    mDiameter(mCount) = diameter
    mArea(mCount) = actualArea
    mSteelClass(mCount) = steelClass
    mComment(mCount) = comment
    mBarAnnotationGroupName(mCount) = Trim$(annotationGroupName)
End Sub

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Get BarID(ByVal index As Long) As String
    CheckIndex index
    BarID = mBarID(index)
End Property

Public Property Get X(ByVal index As Long) As Double
    CheckIndex index
    X = mX(index)
End Property

Public Property Get Y(ByVal index As Long) As Double
    CheckIndex index
    Y = mY(index)
End Property

Public Property Get Diameter(ByVal index As Long) As Double
    CheckIndex index
    Diameter = mDiameter(index)
End Property

Public Property Get Area(ByVal index As Long) As Double
    CheckIndex index
    Area = mArea(index)
End Property

Public Property Get SteelClass(ByVal index As Long) As String
    CheckIndex index
    SteelClass = mSteelClass(index)
End Property

Public Property Get Comment(ByVal index As Long) As String
    CheckIndex index
    Comment = mComment(index)
End Property

Public Property Get BarAnnotationGroupName(ByVal index As Long) As String
    CheckIndex index
    BarAnnotationGroupName = mBarAnnotationGroupName(index)
End Property

' Добавляет опорную линию группы арматуры, по которой потом строится групповая подпись на схеме.
Public Sub AddAnnotationAnchor(ByVal groupName As String, ByVal startX As Double, ByVal startY As Double, _
        ByVal endX As Double, ByVal endY As Double, ByVal normalX As Double, ByVal normalY As Double, _
        Optional ByVal axisDistance As Double = 0#)
    If Len(Trim$(groupName)) = 0 Then Exit Sub

    mAnnotationCount = mAnnotationCount + 1
    ReDim Preserve mAnnotationGroupName(1 To mAnnotationCount)
    ReDim Preserve mAnnotationStartX(1 To mAnnotationCount)
    ReDim Preserve mAnnotationStartY(1 To mAnnotationCount)
    ReDim Preserve mAnnotationEndX(1 To mAnnotationCount)
    ReDim Preserve mAnnotationEndY(1 To mAnnotationCount)
    ReDim Preserve mAnnotationNormalX(1 To mAnnotationCount)
    ReDim Preserve mAnnotationNormalY(1 To mAnnotationCount)
    ReDim Preserve mAnnotationAxisDistance(1 To mAnnotationCount)

    mAnnotationGroupName(mAnnotationCount) = Trim$(groupName)
    mAnnotationStartX(mAnnotationCount) = startX
    mAnnotationStartY(mAnnotationCount) = startY
    mAnnotationEndX(mAnnotationCount) = endX
    mAnnotationEndY(mAnnotationCount) = endY
    mAnnotationNormalX(mAnnotationCount) = normalX
    mAnnotationNormalY(mAnnotationCount) = normalY
    mAnnotationAxisDistance(mAnnotationCount) = axisDistance
End Sub

Public Property Get AnnotationCount() As Long
    AnnotationCount = mAnnotationCount
End Property

Public Property Get AnnotationGroupName(ByVal index As Long) As String
    CheckAnnotationIndex index
    AnnotationGroupName = mAnnotationGroupName(index)
End Property

Public Property Get AnnotationStartX(ByVal index As Long) As Double
    CheckAnnotationIndex index
    AnnotationStartX = mAnnotationStartX(index)
End Property

Public Property Get AnnotationStartY(ByVal index As Long) As Double
    CheckAnnotationIndex index
    AnnotationStartY = mAnnotationStartY(index)
End Property

Public Property Get AnnotationEndX(ByVal index As Long) As Double
    CheckAnnotationIndex index
    AnnotationEndX = mAnnotationEndX(index)
End Property

Public Property Get AnnotationEndY(ByVal index As Long) As Double
    CheckAnnotationIndex index
    AnnotationEndY = mAnnotationEndY(index)
End Property

Public Property Get AnnotationNormalX(ByVal index As Long) As Double
    CheckAnnotationIndex index
    AnnotationNormalX = mAnnotationNormalX(index)
End Property

Public Property Get AnnotationNormalY(ByVal index As Long) As Double
    CheckAnnotationIndex index
    AnnotationNormalY = mAnnotationNormalY(index)
End Property

Public Property Get AnnotationAxisDistance(ByVal index As Long) As Double
    CheckAnnotationIndex index
    AnnotationAxisDistance = mAnnotationAxisDistance(index)
End Property

' Проверяет индекс стержня перед доступом к массивам раскладки.
Private Sub CheckIndex(ByVal index As Long)
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 3106, "CRebarLayout", "Индекс стержня вне допустимого диапазона."
End Sub

' Проверяет индекс аннотационной линии перед доступом к массивам раскладки.
Private Sub CheckAnnotationIndex(ByVal index As Long)
    If index < 1 Or index > mAnnotationCount Then Err.Raise vbObjectError + 3107, "CRebarLayout", "Индекс аннотации арматуры вне допустимого диапазона."
End Sub








