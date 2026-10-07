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

Private Const AUTOCAD_REBAR_LABEL_EXTRA_OFFSET_MM As Double = 5# ' Прежний зазор сводной подписи над бетонным габаритом.

' Для AutoCAD-сетки у программы нет смысловых граней, но есть согласованная
' прямоугольная оболочка бетонных элементов. Ее строит CSectionModel:
' фактический прямоугольник, эквивалентный прямоугольник по A/I или
' последний эквивалентный квадрат по площади. Реальные A/I при этом остаются
' расчетными характеристиками элемента. Здесь добавляются только общие
' semantic-аннотации импорта: точные габариты при известном наружном контуре,
' иначе приблизительные габариты бетонной сетки, а также
' сводная подпись всей распознанной арматуры.
Public Sub PrepareImportedSection(ByVal section As CSectionModel, ByVal bounds As Variant)
    If section Is Nothing Then Exit Sub
    If StrComp(section.SourceType, "AutoCADImport", vbTextCompare) <> 0 Then Exit Sub
    If section.ConcreteCount <= 0 Then Exit Sub

    Dim minX As Double
    Dim maxX As Double
    Dim minY As Double
    Dim maxY As Double
    minX = bounds(1): maxX = bounds(2): minY = bounds(3): maxY = bounds(4)

    Dim knownOuter As Boolean, i As Long, query As CSectionGeometryQuery, domain As CConcreteRegion
    For i = 1 To section.Contours.Count
        If StrComp(section.Contours.LoopRole(i), "Outer", vbTextCompare) = 0 Then knownOuter = True: Exit For
    Next i
    If knownOuter Then
        Set query = New CSectionGeometryQuery: query.Initialize section
        Set domain = query.ConcreteDomain
        query.ProjectionBounds domain, 1#, 0#, minX, maxX
        query.ProjectionBounds domain, 0#, 1#, minY, maxY
    End If

    Dim annotations As CSectionAnnotations
    Set annotations = section.Annotations

    If Not AnnotationExists(annotations, "DIM_AUTO_BOUNDS_B") Then
        If knownOuter Then
            annotations.AddDimension "DIM_AUTO_BOUNDS_B", minX, minY, maxX, minY, 0#, -1#, _
                "B", maxX - minX, "Общая ширина по достоверному наружному контуру импортированного сечения."
        Else
            annotations.AddDimension "DIM_AUTO_BOUNDS_B", minX, minY, maxX, minY, 0#, -1#, _
                ApproxSign(), maxX - minX, "Приблизительная общая ширина по расчетным границам импортированного бетона."
        End If
    End If
    If Not AnnotationExists(annotations, "DIM_AUTO_BOUNDS_H") Then
        If knownOuter Then
            annotations.AddDimension "DIM_AUTO_BOUNDS_H", minX, minY, minX, maxY, -1#, 0#, _
                "H", maxY - minY, "Общая высота по достоверному наружному контуру импортированного сечения."
        Else
            annotations.AddDimension "DIM_AUTO_BOUNDS_H", minX, minY, minX, maxY, -1#, 0#, _
                ApproxSign(), maxY - minY, "Приблизительная общая высота по расчетным границам импортированного бетона."
        End If
    End If

    If section.RebarCount > 0 And Not AnnotationExists(annotations, "REBAR_AUTO_ALL") Then
        Dim labelText As String
        labelText = AutoCADRebarSummaryText(section)
        If Len(labelText) > 0 Then
            Dim labelY As Double
            labelY = maxY + AUTOCAD_REBAR_LABEL_EXTRA_OFFSET_MM
            annotations.AddRebarLabel "REBAR_AUTO_ALL", minX, labelY, maxX, labelY, 0#, 1#, _
                labelText, "Сводная подпись всей арматуры, распознанной при импорте AutoCAD."
        End If
    End If
End Sub

' Проверяет, есть ли semantic-аннотация с таким ID. Это защищает повторную
' запись одного и того же CSectionModel от дублирования размеров и подписей.
Private Function AnnotationExists(ByVal annotations As CSectionAnnotations, ByVal annotationID As String) As Boolean
    If annotations Is Nothing Then Exit Function

    Dim i As Long
    For i = 1 To annotations.Count
        If StrComp(annotations.AnnotationID(i), annotationID, vbTextCompare) = 0 Then
            AnnotationExists = True
            Exit Function
        End If
    Next i
End Function

' Собирает одну подпись для всей арматуры AutoCAD import. Диаметры группируются
' по видимому инженерному значению, сортируются по убыванию и выводятся без
' пробелов вокруг плюса: ALL: 3∅32+23∅16.
Private Function AutoCADRebarSummaryText(ByVal section As CSectionModel) As String
    If section Is Nothing Then Exit Function
    If section.RebarCount <= 0 Then Exit Function

    Dim diameterTexts() As String
    Dim diameterValues() As Double
    Dim counts() As Long
    Dim groupCount As Long

    Dim i As Long
    For i = 1 To section.RebarCount
        Dim diameterValue As Double
        diameterValue = section.RebarDiameter(i)
        If diameterValue > 0# Then
            AddRebarDiameterGroup diameterTexts, diameterValues, counts, groupCount, _
                FormatRebarDiameterForAnnotation(diameterValue), diameterValue
        End If
    Next i
    If groupCount <= 0 Then Exit Function

    SortRebarDiameterGroupsDescending diameterTexts, diameterValues, counts, groupCount

    Dim resultText As String
    resultText = "ALL: "
    For i = 1 To groupCount
        If i > 1 Then resultText = resultText & "+"
        resultText = resultText & CStr(counts(i)) & RebarDiameterSymbol() & diameterTexts(i)
    Next i
    AutoCADRebarSummaryText = resultText
End Function

' Добавляет диаметр в существующую группу или создает новую группу. Группировка
' идет по уже отформатированному тексту, чтобы 32 и 32.0000001 попали в одну
' пользовательскую запись ∅32.
Private Sub AddRebarDiameterGroup(ByRef diameterTexts() As String, ByRef diameterValues() As Double, _
        ByRef counts() As Long, ByRef groupCount As Long, ByVal diameterText As String, ByVal diameterValue As Double)
    Dim i As Long
    For i = 1 To groupCount
        If StrComp(diameterTexts(i), diameterText, vbTextCompare) = 0 Then
            counts(i) = counts(i) + 1
            Exit Sub
        End If
    Next i

    groupCount = groupCount + 1
    ReDim Preserve diameterTexts(1 To groupCount)
    ReDim Preserve diameterValues(1 To groupCount)
    ReDim Preserve counts(1 To groupCount)
    diameterTexts(groupCount) = diameterText
    diameterValues(groupCount) = diameterValue
    counts(groupCount) = 1
End Sub

' Стабилизирует порядок сводной подписи: сначала идут стержни большего
' диаметра. Для типичного импорта это дает привычный вид 3∅32+23∅16.
Private Sub SortRebarDiameterGroupsDescending(ByRef diameterTexts() As String, ByRef diameterValues() As Double, _
        ByRef counts() As Long, ByVal groupCount As Long)
    Dim i As Long
    Dim j As Long
    For i = 1 To groupCount - 1
        For j = i + 1 To groupCount
            If diameterValues(j) > diameterValues(i) Then
                SwapRebarDiameterGroups diameterTexts, diameterValues, counts, i, j
            End If
        Next j
    Next i
End Sub

' Меняет местами две группы диаметров при простой сортировке сводной подписи.
Private Sub SwapRebarDiameterGroups(ByRef diameterTexts() As String, ByRef diameterValues() As Double, _
        ByRef counts() As Long, ByVal firstIndex As Long, ByVal secondIndex As Long)
    Dim textBuffer As String
    Dim valueBuffer As Double
    Dim countBuffer As Long

    textBuffer = diameterTexts(firstIndex)
    valueBuffer = diameterValues(firstIndex)
    countBuffer = counts(firstIndex)

    diameterTexts(firstIndex) = diameterTexts(secondIndex)
    diameterValues(firstIndex) = diameterValues(secondIndex)
    counts(firstIndex) = counts(secondIndex)

    diameterTexts(secondIndex) = textBuffer
    diameterValues(secondIndex) = valueBuffer
    counts(secondIndex) = countBuffer
End Sub

' Форматирует диаметр для подписи арматуры без зависимости от десятичной
' запятой Windows и без незначащих хвостовых нулей.
Private Function FormatRebarDiameterForAnnotation(ByVal value As Double) As String
    Dim textValue As String
    textValue = Replace$(Format$(value, "0.###"), ",", ".")
    Do While InStr(1, textValue, ".", vbTextCompare) > 0 And Right$(textValue, 1) = "0"
        textValue = Left$(textValue, Len(textValue) - 1)
    Loop
    If Right$(textValue, 1) = "." Then textValue = Left$(textValue, Len(textValue) - 1)
    FormatRebarDiameterForAnnotation = textValue
End Function

' Возвращает символ диаметра для AutoCAD-сводки через ChrW$, чтобы модуль не
' зависел от кодировки исходного .cls-файла.
Private Function RebarDiameterSymbol() As String
    RebarDiameterSymbol = ChrW$(&H2205)
End Function

' Возвращает знак приблизительного равенства без литерала в исходнике. Так VBA-модуль
' не зависит от кодировки файла, а на лист Results попадает настоящий знак приблизительного равенства.
Private Function ApproxSign() As String
    ApproxSign = ChrW$(&H2248)
End Function

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
