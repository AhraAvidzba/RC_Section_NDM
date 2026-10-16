VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionRecipeRoundedRectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Бетонное сечение со скруглёнными боковыми сторонами
' ==========================================================================
' Класс описывает сечение с прямой горизонтальной базой и двумя боковыми сторонами.
' Каждая сторона может быть обычной или иметь выступ; радиусы задаются независимо.
' Линии центров арматуры получаются смещением того же многоугольника, что задаёт бетон.
'
' Рецепт описывает границу бетона, расположение арматуры вдоль сторон и аннотации
' схемы (размеры сечения, количество и диаметры стержней). Общие геометрические
' классы строят дуги и смещения, а CSectionBuilder
' создаёт стержни и бетонную сетку.

Implements ISectionRecipe

Private Type TRoundedInput
    B As Double ' Ширина наружного контура, мм.
    H As Double ' Высота наружного контура, мм.
    SideTypes(1 To 2) As String ' Вид левой и правой боковых сторон.
    Widths(1 To 2) As Double ' Ширины левого и правого выступов, мм.
    R1(1 To 2) As Double ' Радиусы первых скруглений левой и правой сторон, мм.
    R2(1 To 2) As Double ' Радиусы вторых скруглений левой и правой сторон, мм.
    MainHeader As Long ' Строка заголовка основной таблицы армирования во входных данных.
    ExtraHeader As Long ' Строка заголовка дополнительных рядов арматуры во входных данных.
    GeometrySource As String ' Адрес исходных размеров для пояснения геометрических ошибок.
End Type

' Читает форму сечения со скруглёнными сторонами и параметры армирования. Формирует
' описание границы бетона, размещения стержней вдоль прямых и скруглённых сторон,
' а также аннотаций схемы (размеры сечения, количество и диаметры стержней).
Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
    Dim table As CSectionInputTable, units As CUnitSystem, facade As CSectionDefinitionBuilder
    Dim sizes As TRoundedInput, outer As CGeometryPath, polygon As CPolygonSpec
    Set table = context.Table("rngRoundedRectangleGeometry"): Set units = context.Units ' Берём копию входной таблицы и единицы измерения.
    ReadDimensions table, units, sizes ' Читаем и проверяем размеры сечения.
    Set polygon = DescribePolygon(sizes) ' Общее описание вершин для бетона и всех смещённых линий.
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "RoundedRectangle" ' Создаём сборщик описания сечения.
    Set outer = AddGeometry(facade, sizes, polygon) ' Получаем наружный контур по подготовленному многоугольнику.
    AddReinforcement facade, table, units, sizes, polygon ' Задаём ряды и их общие подписи по параметрам таблицы.
    AddDimensions facade, sizes, outer ' Добавляем габариты и размеры сечения.
    Set ISectionRecipe_CreateDefinition = facade.Finish ' Завершаем и возвращаем проверенное описание сечения.
End Function

' Заполняет sizes размерами в миллиметрах и проверяет параметры обеих сторон.
' Для обычной стороны не читает ширину выступа и его радиус; сохраняет адреса активных полей.
Private Sub ReadDimensions(ByVal table As CSectionInputTable, ByVal units As CUnitSystem, ByRef sizes As TRoundedInput)
    Dim row As Long, sides As Long, side As Long, column As Long
    row = table.RowByKey("B", 2) + 1
    sizes.B = table.PositiveLength(row, 2, units): sizes.H = table.PositiveLength(row, 3, units)
    sizes.GeometrySource = table.SourceReference(row, 2) & "; " & table.SourceReference(row, 3)
    sides = table.RowByKey("слева", 2)
    For side = 1 To 2
        column = side + 1
        sizes.SideTypes(side) = table.Choice(sides + 1, column, Array("Simple", "Tapered"))
        sizes.R1(side) = table.NonNegativeLength(sides + 3, column, units)
        sizes.GeometrySource = sizes.GeometrySource & "; " & table.SourceReference(sides + 3, column)
        If sizes.SideTypes(side) = "Tapered" Then
            sizes.Widths(side) = table.Length(sides + 2, column, units)
            sizes.R2(side) = table.NonNegativeLength(sides + 4, column, units)
            sizes.GeometrySource = sizes.GeometrySource & "; " & table.SourceReference(sides + 2, column) & "; " & table.SourceReference(sides + 4, column)
            Require table, sides + 2, column, sizes.Widths(side) > GEOM_TOLERANCE, "Для Tapered размер W должен быть положительным."
        End If
    Next side
    sizes.MainHeader = table.RowByKey("as", 2): sizes.ExtraHeader = table.RowByKey("d2", 2)
End Sub

' Создаёт наружный контур по заданным вершинам и радиусам и сразу назначает его границей бетона.
' Ошибку связывает с исходными размерами.
Private Function AddGeometry(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRoundedInput, ByVal polygon As CPolygonSpec) As CGeometryPath
    On Error GoTo InvalidGeometry
    Dim outer As CGeometryPath
    Set outer = facade.RoundedPolygon("Boundary", polygon) ' Создаём готовый замкнутый контур по исходным вершинам и радиусам.
    ' Назначаем эту замкнутую линию наружной границей бетона.
    facade.Outer "OUTER", outer.ID, "Прямой участок контура скругленного прямоугольника.", arcComment:="Дуга контура скругленного прямоугольника.", namingPolicy:=SequentialByKind
    Set AddGeometry = outer
    Exit Function
InvalidGeometry:
    SectionRecipeGeometryError "CSectionRecipeRoundedRectangle", sizes.GeometrySource
End Function

' Возвращает завершённое описание вершин и радиусов: четыре основных угла и вершины
' включённых выступов. Этот же объект используется для контура бетона и линий арматуры.
Private Function DescribePolygon(ByRef sizes As TRoundedInput) As CPolygonSpec
    Dim polygon As CPolygonSpec
    Set polygon = New CPolygonSpec
    polygon.AddVertex "TL", -sizes.B / 2#, sizes.H / 2#, sizes.R1(1), sizes.GeometrySource, outgoingEdgeID:="Edge.TL" ' Верхний левый угол; Edge.TL — верхняя грань, размеры в мм.
    polygon.AddVertex "TR", sizes.B / 2#, sizes.H / 2#, sizes.R1(2), sizes.GeometrySource ' Верхний правый угол до скругления; координаты и радиус в мм.
    ' Вершина правого выступа до скругления; координаты и радиус в мм.
    If sizes.SideTypes(2) = "Tapered" Then polygon.AddVertex "RightNose", sizes.B / 2# + sizes.Widths(2), 0#, sizes.R2(2), sizes.GeometrySource
    polygon.AddVertex "BR", sizes.B / 2#, -sizes.H / 2#, sizes.R1(2), sizes.GeometrySource, outgoingEdgeID:="Edge.BR" ' Нижний правый угол; Edge.BR — нижняя грань, размеры в мм.
    polygon.AddVertex "BL", -sizes.B / 2#, -sizes.H / 2#, sizes.R1(1), sizes.GeometrySource, outgoingEdgeID:="Edge.BL" ' Нижний левый угол; Edge.BL — участок к следующей вершине, размеры в мм.
    ' Вершина левого выступа до скругления; координаты и радиус в мм.
    If sizes.SideTypes(1) = "Tapered" Then polygon.AddVertex "LeftNose", -sizes.B / 2# - sizes.Widths(1), 0#, sizes.R2(1), sizes.GeometrySource, outgoingEdgeID:="Edge.LeftNose"
    polygon.Seal ' Проверяем вершины и завершаем их описание.
    Set DescribePolygon = polygon ' Возвращаем подготовленный многоугольник.
End Function

' Задаёт ряды арматуры вдоль четырёх сторон и объединяет их в общую подпись по каждой стороне.
' Порядок сторон сохраняет принадлежность общего углового стержня первому из заданных рядов.
Private Sub AddReinforcement(ByVal facade As CSectionDefinitionBuilder, ByVal table As CSectionInputTable, ByVal units As CUnitSystem, ByRef sizes As TRoundedInput, ByVal polygon As CPolygonSpec)
    Dim faces As Variant, captions As Variant, index As Long, row As Long, extra As Long, count As Long
    Dim groupID As String, diameter As Double, axisDistance As Double, options As CRebarRuleOptions, path As CGeometryPath
    Dim nx As Double, ny As Double, normalSign As Double, diameters(2 To 3) As Double, locations(2 To 3) As String, distance As Double
    Dim startColumn As Long, endColumn As Long, rowNumber As Long, column As Long, binding As String, keys As Variant
    faces = Array("B.Top", "B.Bottom", "H.Left", "H.Right")
    captions = Array("B - верхняя", "B - нижняя", "H - левая", "H - правая")
    For index = 0 To 3
        groupID = CStr(faces(index))
        keys = Array(captions(index), Replace$(groupID, ".", "-"), groupID)
        row = table.RowByAnyKey(keys, 1, sizes.MainHeader + 1, sizes.MainHeader + 4)
        count = table.Count(row, 4): diameter = table.OptionalDiameter(row, 3, units)
        If count = 0 Or diameter = 0# Then GoTo NextFace
        axisDistance = table.Length(row, 2, units)
        Require table, row, 2, axisDistance > diameter / 2#, "Расстояние до оси должно быть больше половины диаметра стержня."
        Set path = AxisPath(facade, polygon, groupID, axisDistance, table.SourceReference(row, 2)) ' Получаем линию центров стержней выбранной стороны с отступом as.
        nx = 0#: ny = 0#
        Select Case groupID
            Case "B.Top": ny = 1#: normalSign = -1#
            Case "B.Bottom": ny = -1#: normalSign = 1#
            Case "H.Left": nx = -1#: normalSign = 1#
            Case "H.Right": nx = 1#: normalSign = -1#
        End Select
        Set options = New CRebarRuleOptions ' Создаём настройки размещения этого ряда.
        If index >= 2 Then options.EndMode = ExcludeBoth ' Исключаем стержни на обоих концах линии.
        If startColumn = 0 Then
            startColumn = table.ColumnByCaption(sizes.MainHeader, "StartSetback")
            endColumn = table.ColumnByCaption(sizes.MainHeader, "EndSetback")
        End If
        options.StartSetback = table.NonNegativeLength(row, startColumn, units) ' Отступ от начала линии, мм.
        options.EndSetback = table.NonNegativeLength(row, endColumn, units) ' Отступ от конца линии, мм.
        options.AnnotationGroup = groupID ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой стороны.
        options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
        options.Coincidence = KeepPublishedOwner ' При совпадении центров сохраняем стержень, который был добавлен раньше.
        options.IncomingTangentAtJoint = True ' На стыке используем направление предыдущего участка.
        options.Comment = "RoundedRectangle: " & CStr(captions(index)) & " грань, ряд 1"
        ' Адреса полей для сообщения об ошибке размещения.
        options.SourceReference = table.SourceReference(row, 2) & "; " & table.SourceReference(row, startColumn) & "; " & table.SourceReference(row, endColumn)
        facade.RebarByCount groupID & ".Main", path.ID, count, diameter, options ' Задаём основной ряд по числу стержней и диаметру.
        extra = table.RowByAnyKey(keys, 1, sizes.ExtraHeader + 1, sizes.ExtraHeader + 4)
        diameters(2) = table.OptionalDiameter(extra, 2, units)
        diameters(3) = table.OptionalDiameter(extra, 5, units)
        For rowNumber = 2 To 3
            If diameters(rowNumber) > 0# Then
                column = 2 + (rowNumber - 2) * 3 ' В таблице на каждый ряд отведены диаметр, расположение и привязка.
                locations(rowNumber) = table.Choice(extra, column + 1, Array("Stacked", "SideBySide"))
                distance = 0#
                If rowNumber = 3 And locations(2) = locations(3) Then distance = diameters(2) ' Обходим номинальный второй ряд в том же направлении.
                binding = table.Choice(extra, column + 2, Array("EachBar", "EverySecondBar"))
                Set options = RecipeContactOptions(table.SourceReference(extra, column), locations(rowNumber), binding, _
                    diameter, diameters(rowNumber), normalSign, interveningDiameterSum:=distance)
                options.AnnotationGroup = groupID ' Одна подпись для всех рядов этой стороны.
                options.Publication = PublishPerSourcePosition ' Сначала основной стержень, затем связанные с ним.
                options.Coincidence = KeepPublishedOwner ' В совпавшей позиции сохраняем ранее добавленный стержень.
                options.Comment = "RoundedRectangle: " & CStr(captions(index)) & " грань, ряд " & CStr(rowNumber)
                facade.RebarDependent groupID & ".Row" & CStr(rowNumber), groupID & ".Main", diameters(rowNumber), options ' Оба дополнительных ряда связаны с основным.
            End If
        Next rowNumber
        If index >= 2 Then
            facade.RebarRowExtentLabel groupID, groupID & ".Main", nx, ny, True ' Привязываем подпись к крайним стержням основного ряда.
        Else
            facade.RebarLabel groupID, path.ID, nx, ny ' Общую подпись рядов привязываем к границе их линии.
        End If
NextFace:
    Next index
End Sub

' Возвращает линию центров стержней выбранной стороны после смещения контура на as.
' Для верхней и нижней сторон берёт прямой участок, для боковых — также скругления;
' обе боковые линии направлены сверху вниз.
Private Function AxisPath(ByVal facade As CSectionDefinitionBuilder, ByVal polygon As CPolygonSpec, _
        ByVal groupID As String, ByVal axisDistance As Double, ByVal sourceReference As String) As CGeometryPath
    On Error GoTo InvalidPath
    Dim offset As CGeometryPath
    Set offset = facade.OffsetRoundedPolygon("Offset." & groupID, polygon, axisDistance) ' Смещаем тот же многоугольник внутрь бетона на as, мм.
    Select Case groupID
        Case "B.Top": Set AxisPath = facade.Chain(groupID, offset.ID, "Edge.TL", 0#, "Edge.TL", 1#) ' Верхняя прямая грань слева направо.
        Case "B.Bottom": Set AxisPath = facade.Chain(groupID, offset.ID, "Edge.BR", 1#, "Edge.BR", 0#, False) ' Нижняя прямая грань справа налево.
        Case "H.Left"
            Set AxisPath = facade.Chain(groupID, offset.ID, "Edge.TL", 0#, "Edge.BR", 1#, False) ' Боковая сторона со скруглениями, сверху вниз.
        Case "H.Right"
            Set AxisPath = facade.Chain(groupID, offset.ID, "Edge.TL", 1#, "Edge.BR", 0#) ' Боковая сторона со скруглениями, сверху вниз.
    End Select
    Exit Function
InvalidPath:
    SectionRecipeGeometryError "CSectionRecipeRoundedRectangle", sourceReference
End Function

' Добавляет ширину прямой базы B, полную высоту H и фактические размеры включённых боковых выступов.
Private Sub AddDimensions(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRoundedInput, ByVal outer As CGeometryPath)
    Dim minX As Double, maxX As Double
    outer.ProjectionBounds 1#, 0#, minX, maxX
    ' Ширина прямой базы B; размер сверху.
    facade.HorizontalDistance "DIM_B", "B", -sizes.B / 2#, sizes.B / 2#, sizes.H / 2#, 1#, "Размер горизонтальной базы RoundedRectangle между левым и правым торцом."
    facade.VerticalExtent "DIM_H", outer.ID, "H", minX, -1#, "Полная высота RoundedRectangle." ' Полная высота сечения H; размер вынесен влево.
    ' Фактическая ширина левого выступа; подпись округляется до 1 мм.
    If sizes.SideTypes(1) = "Tapered" And -sizes.B / 2# - minX > GEOM_TOLERANCE Then facade.HorizontalDistance "DIM_W_LEFT", "W'", minX, -sizes.B / 2#, sizes.H / 2#, 1#, "Фактический вынос левого края бетона относительно левого торца базы B.", 1#
    ' Фактическая ширина правого выступа; подпись округляется до 1 мм.
    If sizes.SideTypes(2) = "Tapered" And maxX - sizes.B / 2# > GEOM_TOLERANCE Then facade.HorizontalDistance "DIM_W_RIGHT", "W'", sizes.B / 2#, maxX, sizes.H / 2#, 1#, "Фактический вынос правого края бетона относительно правого торца базы B.", 1#
End Sub

' Проверяет условие для указанной ячейки. При нарушении сообщает причину и её фактический адрес.
Private Sub Require(ByVal table As CSectionInputTable, ByVal row As Long, ByVal column As Long, ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then SectionRecipeError "CSectionRecipeRoundedRectangle", reason, table.SourceReference(row, column)
End Sub
