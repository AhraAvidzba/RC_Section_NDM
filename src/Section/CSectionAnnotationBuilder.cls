VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionAnnotationBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const AUTOCAD_REBAR_LABEL_EXTRA_OFFSET_MM As Double = 5# ' Прежний зазор сводной подписи над бетонным габаритом.

' Общая интерпретация модельных измерений и привязок. Оформление и положение
' в пикселях остаются в существующих layout/plotter; контуры не меняются.
Public Sub Build(ByVal definition As CSectionDefinition, ByVal rebars As CRebarLayout, ByVal annotations As CSectionAnnotations, _
        Optional ByVal rules As CRebarRuleExecutor = Nothing)
    Dim value As TSectionAnnotation, i As Long, minX As Double, maxX As Double, minY As Double, maxY As Double
    annotations.Clear
    For i = 1 To definition.AnnotationCount
        definition.CopyAnnotation i, value
        If value.Kind <> "RebarRowLabel" And value.Kind <> "RebarRowExtentLabel" And value.Kind <> "RebarRowExtremeLabel" And value.Kind <> "BetweenPoints" Then _
            TargetBounds definition, value.Target, minX, maxX, minY, maxY
        Select Case value.Kind
            Case "BetweenPoints"
                annotations.AddDimension value.AnnotationID, value.StartX, value.StartY, value.EndX, value.EndY, _
                    value.NormalX, value.NormalY, value.Name, DisplayValue(value.Value, value.DisplayIncrement), value.Comment
            Case "VerticalExtent"
                annotations.AddDimension value.AnnotationID, value.StartX, minY, value.EndX, maxY, _
                    value.NormalX, value.NormalY, value.Name, maxY - minY, value.Comment
            Case "HorizontalExtent"
                annotations.AddDimension value.AnnotationID, minX, value.StartY, maxX, value.EndY, _
                    value.NormalX, value.NormalY, value.Name, maxX - minX, value.Comment
            Case "RebarLabel"
                If Abs(value.NormalY) >= Abs(value.NormalX) Then
                    value.StartX = minX: value.EndX = maxX
                    If value.NormalY >= 0# Then value.StartY = maxY Else value.StartY = minY
                    value.EndY = value.StartY
                Else
                    value.StartY = minY: value.EndY = maxY
                    If value.NormalX >= 0# Then value.StartX = maxX Else value.StartX = minX
                    value.EndX = value.StartX
                End If
                rebars.AddAnnotationAnchor value.Name, value.StartX, value.StartY, value.EndX, value.EndY, _
                    value.NormalX, value.NormalY, value.Value
            Case "RebarRowLabel"
                If rules Is Nothing Then SectionRecipeError "CSectionAnnotationBuilder", "Для подписи ряда нужны исполненные правила."
                If rules.RuleEnds(value.Target, value.StartX, value.StartY, value.EndX, value.EndY) Then _
                    rebars.AddAnnotationAnchor value.Name, value.StartX, value.StartY, value.EndX, value.EndY, _
                        value.NormalX, value.NormalY, value.Value
            Case "RebarRowExtentLabel"
                If rules Is Nothing Then SectionRecipeError "CSectionAnnotationBuilder", "Для подписи ряда нужны исполненные правила."
                If rules.RuleBounds(value.Target, minX, maxX, minY, maxY) Then
                    If Abs(value.NormalY) >= Abs(value.NormalX) Then
                        value.StartX = minX: value.EndX = maxX
                        If value.NormalY >= 0# Then value.StartY = maxY Else value.StartY = minY
                        value.EndY = value.StartY
                    Else
                        value.StartY = minY: value.EndY = maxY
                        If value.NormalX >= 0# Then value.StartX = maxX Else value.StartX = minX
                        value.EndX = value.StartX
                    End If
                    If value.ReverseDirection Then
                        Dim swap As Double
                        swap = value.StartX: value.StartX = value.EndX: value.EndX = swap
                        swap = value.StartY: value.StartY = value.EndY: value.EndY = swap
                    End If
                    rebars.AddAnnotationAnchor value.Name, value.StartX, value.StartY, value.EndX, value.EndY, _
                        value.NormalX, value.NormalY, value.Value
                End If
            Case "RebarRowExtremeLabel"
                If rules Is Nothing Then SectionRecipeError "CSectionAnnotationBuilder", "Для подписи ряда нужны исполненные правила."
                If rules.RuleExtremeEnds(value.Target, value.DirectionX, value.DirectionY, value.StartX, value.StartY, value.EndX, value.EndY) Then
                    If value.DirectionY < 0# Then
                        swap = value.StartX: value.StartX = value.EndX: value.EndX = swap
                        swap = value.StartY: value.StartY = value.EndY: value.EndY = swap
                    End If
                    rebars.AddAnnotationAnchor value.Name, value.StartX, value.StartY, value.EndX, value.EndY, _
                        value.NormalX, value.NormalY, value.Value
                End If
            Case Else
                SectionRecipeError "CSectionAnnotationBuilder", "Неизвестное измерение '" & value.Kind & "'."
        End Select
    Next i
    Dim labels As CRebarGroupAnnotationBuilder
    Set labels = New CRebarGroupAnnotationBuilder
    labels.AddRebarLabels annotations, rebars
End Sub

' Округляется только справочная подпись. Концы размерной линии и расчетная
' геометрия сохраняют исходную точность, включая размеры меньше шага.
Private Function DisplayValue(ByVal value As Double, ByVal increment As Double) As Double
    If increment < 0# Then SectionRecipeError "CSectionAnnotationBuilder", "Шаг округления подписи не может быть отрицательным."
    If increment = 0# Then DisplayValue = value Else DisplayValue = Sgn(value) * Int(Abs(value) / increment + 0.5) * increment
End Function

Private Sub TargetBounds(ByVal definition As CSectionDefinition, ByVal target As String, _
        ByRef minX As Double, ByRef maxX As Double, ByRef minY As Double, ByRef maxY As Double)
    Dim ids As Variant, id As Variant, path As CGeometryPath, low As Double, high As Double
    ids = definition.TargetPaths(target)
    minX = 1E+100: maxX = -1E+100: minY = 1E+100: maxY = -1E+100
    For Each id In ids
        Set path = definition.Path(CStr(id))
        path.ProjectionBounds 1#, 0#, low, high: minX = GeomMin(minX, low): maxX = GeomMax(maxX, high)
        path.ProjectionBounds 0#, 1#, low, high: minY = GeomMin(minY, low): maxY = GeomMax(maxY, high)
    Next id
End Sub

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
' пробелов вокруг плюса; каждый диаметр сопровождается своим количеством.
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
' пользовательскую запись диаметра 32 мм.
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
' диаметра. Для типичного импорта это дает привычную запись количества и диаметров.
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
