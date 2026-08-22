VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionAnnotations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Контейнер semantic-аннотаций рассчитанного сечения
' ==========================================================================
' Здесь хранятся не расчетные результаты, а смысловые подсказки для вывода:
' контур, характерные размеры и групповые подписи арматуры. Аннотации задаются
' в координатах модели и позже сериализуются в rngNDMSectionAnnotations.

Private mType() As String
Private mAnnotationID() As String
Private mStartX() As Double
Private mStartY() As Double
Private mEndX() As Double
Private mEndY() As Double
Private mNormalX() As Double
Private mNormalY() As Double
Private mOffset() As Double
Private mText() As String
Private mValue() As Double
Private mUnit() As String
Private mComment() As String
Private mCount As Long

' Очищает накопленное состояние перед новым расчетом или повторным формированием вывода.
Public Sub Clear()
    Erase mType
    Erase mAnnotationID
    Erase mStartX
    Erase mStartY
    Erase mEndX
    Erase mEndY
    Erase mNormalX
    Erase mNormalY
    Erase mOffset
    Erase mText
    Erase mValue
    Erase mUnit
    Erase mComment
    mCount = 0
End Sub

Public Sub AddAnnotation(ByVal annotationType As String, ByVal annotationID As String, _
        ByVal startX As Double, ByVal startY As Double, ByVal endX As Double, ByVal endY As Double, _
        ByVal normalX As Double, ByVal normalY As Double, ByVal offset As Double, _
        ByVal textValue As String, ByVal numericValue As Double, ByVal unitText As String, _
        ByVal commentValue As String)
    If Len(Trim$(annotationType)) = 0 Then Err.Raise vbObjectError + 5280, "CSectionAnnotations", "Annotation type is empty."
    If Len(Trim$(annotationID)) = 0 Then Err.Raise vbObjectError + 5281, "CSectionAnnotations", "Annotation ID is empty."

    mCount = mCount + 1
    ReDim Preserve mType(1 To mCount)
    ReDim Preserve mAnnotationID(1 To mCount)
    ReDim Preserve mStartX(1 To mCount)
    ReDim Preserve mStartY(1 To mCount)
    ReDim Preserve mEndX(1 To mCount)
    ReDim Preserve mEndY(1 To mCount)
    ReDim Preserve mNormalX(1 To mCount)
    ReDim Preserve mNormalY(1 To mCount)
    ReDim Preserve mOffset(1 To mCount)
    ReDim Preserve mText(1 To mCount)
    ReDim Preserve mValue(1 To mCount)
    ReDim Preserve mUnit(1 To mCount)
    ReDim Preserve mComment(1 To mCount)

    mType(mCount) = Trim$(annotationType)
    mAnnotationID(mCount) = Trim$(annotationID)
    mStartX(mCount) = startX
    mStartY(mCount) = startY
    mEndX(mCount) = endX
    mEndY(mCount) = endY
    mNormalX(mCount) = normalX
    mNormalY(mCount) = normalY
    mOffset(mCount) = offset
    mText(mCount) = textValue
    mValue(mCount) = numericValue
    mUnit(mCount) = unitText
    mComment(mCount) = commentValue
End Sub

Public Sub AddContourLine(ByVal annotationID As String, ByVal startX As Double, ByVal startY As Double, _
        ByVal endX As Double, ByVal endY As Double, Optional ByVal commentValue As String = "Generated section contour segment.")
    AddAnnotation "CONTOUR_LINE", annotationID, startX, startY, endX, endY, 0#, 0#, 0#, vbNullString, 0#, vbNullString, commentValue
End Sub

Public Sub AddContourCircle(ByVal annotationID As String, ByVal centerX As Double, ByVal centerY As Double, _
        ByVal radius As Double, Optional ByVal commentValue As String = "Generated circle contour.")
    AddAnnotation "CONTOUR_CIRCLE", annotationID, centerX, centerY, radius, 0#, 0#, 0#, 0#, vbNullString, radius, vbNullString, commentValue
End Sub

Public Sub AddDimension(ByVal annotationID As String, ByVal startX As Double, ByVal startY As Double, _
        ByVal endX As Double, ByVal endY As Double, ByVal normalX As Double, ByVal normalY As Double, _
        ByVal labelText As String, ByVal numericValue As Double, Optional ByVal commentValue As String = "Section dimension.")
    AddAnnotation "DIMENSION", annotationID, startX, startY, endX, endY, normalX, normalY, 0#, labelText, numericValue, "Length", commentValue
End Sub

Public Sub AddRebarLabel(ByVal annotationID As String, ByVal startX As Double, ByVal startY As Double, _
        ByVal endX As Double, ByVal endY As Double, ByVal normalX As Double, ByVal normalY As Double, _
        ByVal axisDistance As Double, ByVal labelText As String, Optional ByVal commentValue As String = "Grouped rebar annotation.")
    If Len(Trim$(labelText)) = 0 Then Exit Sub
    AddAnnotation "REBAR_ANNOTATION", annotationID, startX, startY, endX, endY, normalX, normalY, axisDistance, labelText, 0#, vbNullString, commentValue
End Sub

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Get AnnotationType(ByVal index As Long) As String
    CheckIndex index
    AnnotationType = mType(index)
End Property

Public Property Get AnnotationID(ByVal index As Long) As String
    CheckIndex index
    AnnotationID = mAnnotationID(index)
End Property

Public Property Get StartX(ByVal index As Long) As Double
    CheckIndex index
    StartX = mStartX(index)
End Property

Public Property Get StartY(ByVal index As Long) As Double
    CheckIndex index
    StartY = mStartY(index)
End Property

Public Property Get EndX(ByVal index As Long) As Double
    CheckIndex index
    EndX = mEndX(index)
End Property

Public Property Get EndY(ByVal index As Long) As Double
    CheckIndex index
    EndY = mEndY(index)
End Property

Public Property Get NormalX(ByVal index As Long) As Double
    CheckIndex index
    NormalX = mNormalX(index)
End Property

Public Property Get NormalY(ByVal index As Long) As Double
    CheckIndex index
    NormalY = mNormalY(index)
End Property

Public Property Get Offset(ByVal index As Long) As Double
    CheckIndex index
    Offset = mOffset(index)
End Property

Public Property Get Text(ByVal index As Long) As String
    CheckIndex index
    Text = mText(index)
End Property

Public Property Get Value(ByVal index As Long) As Double
    CheckIndex index
    Value = mValue(index)
End Property

Public Property Get Unit(ByVal index As Long) As String
    CheckIndex index
    Unit = mUnit(index)
End Property

Public Property Get Comment(ByVal index As Long) As String
    CheckIndex index
    Comment = mComment(index)
End Property

Private Sub CheckIndex(ByVal index As Long)
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 5282, "CSectionAnnotations", "Annotation index is out of range."
End Sub

