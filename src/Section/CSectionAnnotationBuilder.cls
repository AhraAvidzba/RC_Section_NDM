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

' ==========================================================================
' Подготовка размеров и подписей арматуры по геометрии сечения
' ==========================================================================
' По описанию рецепта и готовой арматуре вычисляет значения размеров
' и привязки подписей. Результат сохраняется в CSectionAnnotations
' в координатах модели; оформление диаграммы выполняется отдельно.

Private Const AUTOCAD_REBAR_LABEL_EXTRA_OFFSET_MM As Double = 5# ' Расстояние сводной подписи арматуры над верхней границей бетонной сетки.

' По описанию сечения вычисляет размеры и привязки подписей в координатах модели.
' Шрифт и положение на диаграмме задаются при отрисовке. Каждый вызов готовит
' привязки заново; готовая арматура используется только для чтения.
Public Sub Build(ByVal definition As CSectionDefinition, ByVal rebars As CRebarLayout, ByVal annotations As CSectionAnnotations, _
        Optional ByVal rules As CRebarRuleExecutor = Nothing)
    If definition Is Nothing Or rebars Is Nothing Or annotations Is Nothing Then _
        SectionRecipeError "CSectionAnnotationBuilder", "Определение, арматура и контейнер аннотаций обязательны."
    If Not definition.IsSealed Then SectionRecipeError "CSectionAnnotationBuilder", "Перед созданием размеров и подписей завершите описание сечения методом Finish."
    Dim value As TSectionAnnotation, i As Long, minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim anchors() As TSectionAnnotation, anchorCount As Long
    If definition.AnnotationCount > 0 Then ReDim anchors(1 To definition.AnnotationCount)
    annotations.Clear
    For i = 1 To definition.AnnotationCount
        definition.CopyAnnotation i, value
        Select Case value.Kind
            Case "VerticalExtent", "HorizontalExtent"
                TargetBounds definition, value.Target, minX, maxX, minY, maxY
        End Select
        Select Case value.Kind
            Case "Anchors.Horizontal", "Anchors.Vertical", "Anchors.Aligned"
                AnchorDimension definition, value, annotations
            Case "Radius", "Diameter"
                ArcDimension definition, value, annotations
            Case "PathLength", "Text"
                PointLabel definition, value, annotations
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
                TargetLabelAnchor definition, value
                AppendLabelAnchor anchors, anchorCount, value
            Case "RebarRowLabel"
                If rules Is Nothing Then SectionRecipeError "CSectionAnnotationBuilder", "Для подписи арматуры сначала требуется построить стержни по заданным правилам."
                If rules.RuleEnds(value.Target, value.StartX, value.StartY, value.EndX, value.EndY) Then _
                    AppendLabelAnchor anchors, anchorCount, value
            Case "RebarRowExtentLabel"
                If rules Is Nothing Then SectionRecipeError "CSectionAnnotationBuilder", "Для подписи арматуры сначала требуется построить стержни по заданным правилам."
                Dim nx As Double, ny As Double, tx As Double, ty As Double, support As Double
                LabelDirections value.NormalX, value.NormalY, nx, ny, tx, ty
                If rules.RuleBounds(value.Target, minX, maxX, minY, maxY, tx, ty) Then
                    If -ty * nx + tx * ny >= 0# Then support = maxY Else support = -minY
                    SetLabelAnchor value, nx, ny, tx, ty, minX, maxX, support
                    If value.ReverseDirection Then
                        Dim swap As Double
                        swap = value.StartX: value.StartX = value.EndX: value.EndX = swap
                        swap = value.StartY: value.StartY = value.EndY: value.EndY = swap
                    End If
                    AppendLabelAnchor anchors, anchorCount, value
                End If
            Case "RebarRowExtremeLabel"
                If rules Is Nothing Then SectionRecipeError "CSectionAnnotationBuilder", "Для подписи арматуры сначала требуется построить стержни по заданным правилам."
                If rules.RuleExtremeEnds(value.Target, value.DirectionX, value.DirectionY, value.StartX, value.StartY, value.EndX, value.EndY) Then
                    If value.DirectionY < 0# Then
                        swap = value.StartX: value.StartX = value.EndX: value.EndX = swap
                        swap = value.StartY: value.StartY = value.EndY: value.EndY = swap
                    End If
                    AppendLabelAnchor anchors, anchorCount, value
                End If
            Case Else
                SectionRecipeError "CSectionAnnotationBuilder", "Не распознан способ построения размера '" & value.Kind & "'."
        End Select
    Next i
    Dim labels As CRebarGroupAnnotationBuilder
    Set labels = New CRebarGroupAnnotationBuilder
    labels.AddRebarLabels annotations, rebars, anchors, anchorCount
End Sub

' Находит линию, относительно которой размещается подпись, по точным отрезкам и дугам.
' Учитывает проекции вдоль подписи и поперёк неё, поэтому наклонная грань
' не подменяется горизонтальным или вертикальным габаритом. Перенос к краю
' бетона выполняет CSectionPlotDataReader.
Private Sub TargetLabelAnchor(ByVal definition As CSectionDefinition, ByRef value As TSectionAnnotation)
    Dim nx As Double, ny As Double, tx As Double, ty As Double, ids As Variant, id As Variant
    Dim path As CGeometryPath, low As Double, high As Double, first As Double, last As Double, support As Double
    LabelDirections value.NormalX, value.NormalY, nx, ny, tx, ty
    ids = definition.TargetPaths(value.Target)
    first = 1E+100: last = -1E+100: support = -1E+100
    For Each id In ids
        Set path = definition.Path(CStr(id))
        path.ProjectionBounds tx, ty, low, high
        first = GeomMin(first, low): last = GeomMax(last, high)
        path.ProjectionBounds nx, ny, low, high
        support = GeomMax(support, high)
    Next id
    SetLabelAnchor value, nx, ny, tx, ty, first, last, support
End Sub

' Готовит единичную нормаль и касательную. Порядок концов по X, затем по Y,
' сохраняет прежние горизонтальные и вертикальные подписи при обеих нормалях.
Private Sub LabelDirections(ByVal normalX As Double, ByVal normalY As Double, ByRef nx As Double, ByRef ny As Double, _
        ByRef tx As Double, ByRef ty As Double)
    Dim norm As Double
    norm = Sqr(normalX ^ 2 + normalY ^ 2)
    If norm <= GEOM_TOLERANCE Then nx = 0#: ny = 1# Else nx = normalX / norm: ny = normalY / norm
    tx = ny: ty = -nx
    If tx < 0# Or (tx = 0# And ty < 0#) Then tx = -tx: ty = -ty
End Sub

' Переводит проекционные границы в две точки опорной линии подписи.
Private Sub SetLabelAnchor(ByRef value As TSectionAnnotation, ByVal nx As Double, ByVal ny As Double, _
        ByVal tx As Double, ByVal ty As Double, ByVal first As Double, ByVal last As Double, ByVal support As Double)
    value.StartX = first * tx + support * nx: value.StartY = first * ty + support * ny
    value.EndX = last * tx + support * nx: value.EndY = last * ty + support * ny
End Sub

' Сохраняет координаты и направления привязок для текущего построения аннотаций.
' Готовые координаты стержней не меняются.
Private Sub AppendLabelAnchor(ByRef anchors() As TSectionAnnotation, ByRef count As Long, ByRef value As TSectionAnnotation)
    If Len(Trim$(value.Name)) = 0 Then Exit Sub
    count = count + 1
    anchors(count) = value
    anchors(count).Name = Trim$(value.Name)
End Sub

' Вычисляет горизонтальный, вертикальный размер или расстояние между двумя точками
' в зависимости от заданного вида. Для проекционного размера положение
' размерной линии определяется первой точкой.
Private Sub AnchorDimension(ByVal definition As CSectionDefinition, ByRef value As TSectionAnnotation, ByVal annotations As CSectionAnnotations)
    definition.AnchorCoordinates value.StartAnchorID, value.StartX, value.StartY
    definition.AnchorCoordinates value.EndAnchorID, value.EndX, value.EndY
    Select Case value.Kind
        Case "Anchors.Horizontal"
            value.Value = Abs(value.EndX - value.StartX): value.EndY = value.StartY
        Case "Anchors.Vertical"
            value.Value = Abs(value.EndY - value.StartY): value.EndX = value.StartX
        Case "Anchors.Aligned"
            value.Value = Sqr((value.EndX - value.StartX) ^ 2 + (value.EndY - value.StartY) ^ 2)
    End Select
    If value.Value <= GEOM_TOLERANCE Then SectionRecipeError "CSectionAnnotationBuilder", "Размер '" & value.AnnotationID & "' имеет нулевую длину.", definition.SourceReference(value.AnnotationID)
    If value.NormalX = 0# And value.NormalY = 0# Then
        value.NormalX = -(value.EndY - value.StartY) / value.Value
        value.NormalY = (value.EndX - value.StartX) / value.Value
    End If
    annotations.AddDimension value.AnnotationID, value.StartX, value.StartY, value.EndX, value.EndY, _
        value.NormalX, value.NormalY, value.Name, value.Value, value.Comment
End Sub

' Измеряет выбранную дугу или окружность по её точной геометрии,
' а не по общему габариту сечения или соседнему участку.
Private Sub ArcDimension(ByVal definition As CSectionDefinition, ByRef value As TSectionAnnotation, ByVal annotations As CSectionAnnotations)
    Dim path As CGeometryPath, edge As TRegionEdge, tx As Double, ty As Double, radiusValue As Double
    Set path = definition.Path(value.Target)
    path.CopySegment path.SegmentIndex(value.SegmentID), edge
    If edge.Sweep = 0# Then SectionRecipeError "CSectionAnnotationBuilder", "Размер '" & value.AnnotationID & "' требует дугу или окружность.", definition.SourceReference(value.AnnotationID)
    path.AtDistance path.AnchorStation(value.SegmentID, value.Value), value.EndX, value.EndY, tx, ty
    radiusValue = Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2)
    value.StartX = edge.CenterX: value.StartY = edge.CenterY: value.Value = radiusValue
    If value.Kind = "Diameter" Then
        value.StartX = 2# * edge.CenterX - value.EndX: value.StartY = 2# * edge.CenterY - value.EndY
        value.Value = 2# * radiusValue
    End If
    annotations.AddDimension value.AnnotationID, value.StartX, value.StartY, value.EndX, value.EndY, tx, ty, value.Name, value.Value, value.Comment
End Sub

' Строит точечную привязку подписи по аналитическому пути и заданному направлению.
Private Sub PointLabel(ByVal definition As CSectionDefinition, ByRef value As TSectionAnnotation, ByVal annotations As CSectionAnnotations)
    definition.AnchorCoordinates value.StartAnchorID, value.StartX, value.StartY
    If value.Kind = "Text" Then
        annotations.AddAnnotation "TEXT_ANNOTATION", value.AnnotationID, value.StartX, value.StartY, value.StartX, value.StartY, _
            0#, 0#, value.Name, 0#, vbNullString, value.Comment
    Else
        Dim ids As Variant, id As Variant, path As CGeometryPath
        ids = definition.TargetPaths(value.Target): value.Value = 0#
        For Each id In ids
            Set path = definition.Path(CStr(id)): value.Value = value.Value + path.Length
        Next id
        annotations.AddAnnotation "MEASUREMENT_LABEL", value.AnnotationID, value.StartX, value.StartY, value.StartX, value.StartY, _
            0#, 0#, value.Name, value.Value, "Length", value.Comment
    End If
End Sub

' Округляется только справочная подпись. Концы размерной линии и расчетная
' геометрия сохраняют исходную точность, включая размеры меньше шага.
Private Function DisplayValue(ByVal value As Double, ByVal increment As Double) As Double
    If increment < 0# Then SectionRecipeError "CSectionAnnotationBuilder", "Шаг округления подписи не может быть отрицательным."
    If increment = 0# Then DisplayValue = value Else DisplayValue = Sgn(value) * Int(Abs(value) / increment + 0.5) * increment
End Function

' Определяет проекционные границы именованного пути или группы для размещения размеров.
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

' Добавляет общие аннотации импортированного сечения: горизонтальный и вертикальный
' габариты и общую подпись арматуры. При наличии точного наружного контура размеры
' точные, иначе вычисляются приблизительно по границам бетонных элементов.
' Эти границы заранее готовит CSectionModel по фактическому прямоугольнику
' или по эквивалентным размерам элемента; его расчётные площадь и инерции не меняются.
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

' Проверяет, есть ли аннотация с указанным ID. При повторной записи
' той же модели это предотвращает дублирование размеров и подписей.
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

' Собирает общую подпись всей арматуры, импортированной из AutoCAD.
' Объединяет одинаковые видимые диаметры, сортирует от большего к меньшему
' и выводит количество стержней каждого диаметра через знак плюса без пробелов.
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
