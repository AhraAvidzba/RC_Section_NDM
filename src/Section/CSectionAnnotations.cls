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
' Здесь хранятся размеры и групповые подписи арматуры. Writer дополняет их
' LC-зависимыми расчетными областями трещин в rngNDMSectionAnnotations.
' Исходные outer/opening принадлежат CSectionContours и геометрическому
' снимку, поэтому не смешиваются с оформлением или результатами проверок.

Private mType() As String
Private mAnnotationID() As String
Private mStartX() As Double
Private mStartY() As Double
Private mEndX() As Double
Private mEndY() As Double
Private mNormalX() As Double
Private mNormalY() As Double
Private mText() As String
Private mValue() As Double
Private mUnit() As String
Private mComment() As String
Private mCount As Long

' Удаляет все координаты, размерные значения и подписи semantic-аннотаций.
' Контейнер после очистки не содержит записей для вывода; геометрия модели
' и уже сохраненные Excel-таблицы этим методом не изменяются.
Public Sub Clear()
    Erase mType
    Erase mAnnotationID
    Erase mStartX
    Erase mStartY
    Erase mEndX
    Erase mEndY
    Erase mNormalX
    Erase mNormalY
    Erase mText
    Erase mValue
    Erase mUnit
    Erase mComment
    mCount = 0
End Sub

' Добавляет произвольную semantic-аннотацию в единый контейнер для последующей записи в Results.
Public Sub AddAnnotation(ByVal annotationType As String, ByVal annotationID As String, _
        ByVal startX As Double, ByVal startY As Double, ByVal endX As Double, ByVal endY As Double, _
        ByVal normalX As Double, ByVal normalY As Double, ByVal textValue As String, _
        ByVal numericValue As Double, ByVal unitText As String, _
        ByVal commentValue As String)
    If Len(Trim$(annotationType)) = 0 Then Err.Raise vbObjectError + 5280, "CSectionAnnotations", "Тип аннотации не задан."
    If Len(Trim$(annotationID)) = 0 Then Err.Raise vbObjectError + 5281, "CSectionAnnotations", "ID аннотации не задан."
    If Left$(UCase$(Trim$(annotationType)), 8) = "CONTOUR_" Then Err.Raise vbObjectError + 5280, _
        "CSectionAnnotations", "Исходный контур должен храниться в CSectionModel.Contours, а не в аннотациях."

    mCount = mCount + 1
    ReDim Preserve mType(1 To mCount)
    ReDim Preserve mAnnotationID(1 To mCount)
    ReDim Preserve mStartX(1 To mCount)
    ReDim Preserve mStartY(1 To mCount)
    ReDim Preserve mEndX(1 To mCount)
    ReDim Preserve mEndY(1 To mCount)
    ReDim Preserve mNormalX(1 To mCount)
    ReDim Preserve mNormalY(1 To mCount)
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
    mText(mCount) = textValue
    mValue(mCount) = numericValue
    mUnit(mCount) = unitText
    mComment(mCount) = commentValue
End Sub



' Добавляет размерную линию с числовым значением в единицах длины модели.
Public Sub AddDimension(ByVal annotationID As String, ByVal startX As Double, ByVal startY As Double, _
        ByVal endX As Double, ByVal endY As Double, ByVal normalX As Double, ByVal normalY As Double, _
        ByVal labelText As String, ByVal numericValue As Double, Optional ByVal commentValue As String = "Размер сечения для отображения на схеме.")
    AddAnnotation "DIMENSION", annotationID, startX, startY, endX, endY, normalX, normalY, labelText, numericValue, "Length", commentValue
End Sub

' Добавляет групповую подпись арматуры; пустой текст означает, что подпись выводить не нужно.
Public Sub AddRebarLabel(ByVal annotationID As String, ByVal startX As Double, ByVal startY As Double, _
        ByVal endX As Double, ByVal endY As Double, ByVal normalX As Double, ByVal normalY As Double, _
        ByVal labelText As String, Optional ByVal commentValue As String = "Групповая подпись арматуры по смысловой грани.")
    If Len(Trim$(labelText)) = 0 Then Exit Sub
    AddAnnotation "REBAR_ANNOTATION", annotationID, startX, startY, endX, endY, normalX, normalY, labelText, 0#, vbNullString, commentValue
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

' Проверяет индекс semantic-аннотации перед чтением массивов контейнера.
Private Sub CheckIndex(ByVal index As Long)
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 5282, "CSectionAnnotations", "Индекс аннотации вне допустимого диапазона."
End Sub
