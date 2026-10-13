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
' Декларативный рецепт прямоугольника с независимыми скруглёнными сторонами
' ==========================================================================
Implements ISectionRecipe

' Рецепт RoundedRectangle: независимые Simple/Tapered-стороны, четыре
' смысловых ряда и размеры. Все дуги, смещения и позиции строит общее ядро.
Private Type TRoundedInput
    B As Double
    H As Double
    SideTypes(1 To 2) As String
    Widths(1 To 2) As Double
    R1(1 To 2) As Double
    R2(1 To 2) As Double
    Vertices As Variant
    Radii As Variant
    VertexIDs As Variant
    MainHeader As Long
    ExtraHeader As Long
    GeometrySource As String
End Type

' Описывает контуры, арматурные правила и размеры выбранного рецепта. Сетку и физические стержни
' строит общий CSectionBuilder.
Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
    Dim table As CSectionInputTable, units As CUnitSystem, facade As CSectionDefinitionBuilder  ' Объявляем сборщик описания сечения.
    Dim sizes As TRoundedInput, outer As CGeometryPath                                  ' Объявляем линии и параметры геометрии.
    Set table = context.Table("rngRoundedRectangleGeometry"): Set units = context.Units
    ReadDimensions table, units, sizes
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "RoundedRectangle"    ' Создаём описание скруглённого сечения.
    Set outer = AddGeometry(facade, sizes)                                              ' Описываем наружную границу и геометрию сечения.
    AddReinforcement facade, table, units, sizes                                        ' Описываем армирование всех активных сторон.
    AddDimensions facade, sizes, outer                                                  ' Описываем размеры сечения.
    Set ISectionRecipe_CreateDefinition = facade.Finish                                 ' Проверяем и возвращаем готовое описание сечения.
End Function

' Неактивные W/R2 стороны Simple не читаются; ошибки относятся к исходной
' ячейке, единицы переводятся один раз через CUnitSystem.
Private Sub ReadDimensions(ByVal table As CSectionInputTable, ByVal units As CUnitSystem, ByRef sizes As TRoundedInput)
    Dim row As Long, sides As Long, side As Long, column As Long
    row = table.RowByKey("B", 2) + 1
    sizes.B = table.Length(row, 2, units): sizes.H = table.Length(row, 3, units)
    sizes.GeometrySource = table.SourceReference(row, 2) & "; " & table.SourceReference(row, 3)
    Require table, row, 2, sizes.B > 0#, "Размер B должен быть положительным."
    Require table, row, 3, sizes.H > 0#, "Размер H должен быть положительным."
    sides = table.RowByKey("слева", 2)
    For side = 1 To 2
        column = side + 1
        sizes.SideTypes(side) = table.Choice(sides + 1, column, Array("Simple", "Tapered"))
        sizes.R1(side) = table.Length(sides + 3, column, units)
        sizes.GeometrySource = sizes.GeometrySource & "; " & table.SourceReference(sides + 3, column)
        Require table, sides + 3, column, sizes.R1(side) >= 0#, "Радиус R1 не может быть отрицательным."
        If sizes.SideTypes(side) = "Tapered" Then
            sizes.Widths(side) = table.Length(sides + 2, column, units)
            sizes.R2(side) = table.Length(sides + 4, column, units)
            sizes.GeometrySource = sizes.GeometrySource & "; " & table.SourceReference(sides + 2, column) & "; " & table.SourceReference(sides + 4, column)
            Require table, sides + 2, column, sizes.Widths(side) > GEOM_TOLERANCE, "Для Tapered размер W должен быть положительным."
            Require table, sides + 4, column, sizes.R2(side) >= 0#, "Радиус R2 не может быть отрицательным."
        End If
    Next side
    sizes.MainHeader = table.RowByKey("as", 2): sizes.ExtraHeader = table.RowByKey("d2", 2)
End Sub

' Рецепт описывает только исходные вершины и назначенные радиусы.
' Начало опубликованного контура задано смысловой гранью, а не индексом дуги.
Private Function AddGeometry(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRoundedInput) As CGeometryPath
    On Error GoTo InvalidGeometry
    Dim points(0 To 5) As Variant, radii(0 To 5) As Variant, names(0 To 5) As Variant, count As Long
    AddCorner points, radii, names, count, "TL", -sizes.B / 2#, sizes.H / 2#, sizes.R1(1)  ' Добавляем левую верхнюю вершину и её радиус.
    AddCorner points, radii, names, count, "TR", sizes.B / 2#, sizes.H / 2#, sizes.R1(2)  ' Добавляем правую верхнюю вершину и её радиус.
    If sizes.SideTypes(2) = "Tapered" Then AddCorner points, radii, names, count, "RightNose", sizes.B / 2# + sizes.Widths(2), 0#, sizes.R2(2)  ' Добавляем вершину правого выступа и её радиус.
    AddCorner points, radii, names, count, "BR", sizes.B / 2#, -sizes.H / 2#, sizes.R1(2)  ' Добавляем правую нижнюю вершину и её радиус.
    AddCorner points, radii, names, count, "BL", -sizes.B / 2#, -sizes.H / 2#, sizes.R1(1)  ' Добавляем левую нижнюю вершину и её радиус.
    If sizes.SideTypes(1) = "Tapered" Then AddCorner points, radii, names, count, "LeftNose", -sizes.B / 2# - sizes.Widths(1), 0#, sizes.R2(1)  ' Добавляем вершину левого выступа и её радиус.
    Dim vertices() As Variant, assignedRadii() As Variant, vertexIDs() As Variant, i As Long
    ReDim vertices(0 To count - 1): ReDim assignedRadii(0 To count - 1): ReDim vertexIDs(0 To count - 1)  ' Готовим массивы вершин, радиусов и имён.
    For i = 0 To count - 1: vertices(i) = points(i): assignedRadii(i) = radii(i): vertexIDs(i) = names(i): Next i  ' Копируем подготовленные вершины, радиусы и имена.
    sizes.Vertices = vertices: sizes.Radii = assignedRadii: sizes.VertexIDs = vertexIDs  ' Сохраняем вершины, радиусы и имена контура.
    Dim original As CGeometryPath, outer As CGeometryPath, first As String, edge As TRegionEdge, contourIDs() As String  ' Объявляем линии и параметры геометрии.
    Set original = facade.RoundedPolygon("Shape", sizes.Vertices, sizes.Radii, sizes.VertexIDs)  ' Создаём замкнутый контур с заданными скруглениями.
    If sizes.SideTypes(1) = "Tapered" Then first = "Edge.LeftNose" Else first = "Edge.BL"
    Set outer = facade.Chain("Boundary", original.ID, first, 0#, first, 0#, True, True)  ' Замыкаем контур с началом на выбранном участке.
    ReDim contourIDs(0 To outer.SegmentCount - 1)                                       ' Готовим имена участков контура.
    For i = 1 To outer.SegmentCount
        outer.CopySegment i, edge                                                       ' Получаем участок контура для назначения имени.
        If edge.Sweep = 0# Then contourIDs(i - 1) = "CONTOUR_LINE_" & CStr(i) Else contourIDs(i - 1) = "CONTOUR_ARC_" & CStr(i)  ' Назначаем имена прямым и дуговым участкам контура.
    Next i
    facade.Outer "OUTER", outer.ID, "Прямой участок контура скругленного прямоугольника.", contourIDs, "Дуга контура скругленного прямоугольника."  ' Назначаем наружную границу бетона.
    Set AddGeometry = outer                                                             ' Возвращаем готовую наружную границу.
    Exit Function
InvalidGeometry:
    SectionRecipeGeometryError "CSectionRecipeRoundedRectangle", sizes.GeometrySource
End Function

' Добавляет исходную вершину, её смысловой ID и назначенный радиус скругления.
Private Sub AddCorner(ByRef points() As Variant, ByRef radii() As Variant, ByRef names() As Variant, _
        ByRef count As Long, ByVal name As String, ByVal x As Double, ByVal y As Double, ByVal radius As Double)
    points(count) = Array(x, y): radii(count) = radius: names(count) = name: count = count + 1  ' Записываем координаты, радиус и имя новой вершины.
End Sub

' Порядок граней сохраняет владельца совпавшего стержня. Дополнительные
' ряды исполняются вместе со своей исходной позицией, без повторной раскладки.
Private Sub AddReinforcement(ByVal facade As CSectionDefinitionBuilder, ByVal table As CSectionInputTable, _
        ByVal units As CUnitSystem, ByRef sizes As TRoundedInput)
    Dim faces As Variant, captions As Variant, index As Long, row As Long, extra As Long, count As Long
    Dim groupID As String, diameter As Double, axisDistance As Double, options As CRebarRuleOptions, path As CGeometryPath  ' Объявляем линии геометрии и условия армирования.
    Dim nx As Double, ny As Double, normalSign As Double, d2 As Double, d3 As Double, loc2 As String, loc3 As String, distance As Double
    Dim startColumn As Long, endColumn As Long
    faces = Array("B.Top", "B.Bottom", "H.Left", "H.Right")
    captions = Array("B - верхняя", "B - нижняя", "H - левая", "H - правая")
    For index = 0 To 3
        groupID = CStr(faces(index))
        row = table.RowByAnyKey(Array(captions(index), Replace$(groupID, ".", "-"), groupID), 1, sizes.MainHeader + 1, sizes.MainHeader + 4)
        count = table.Count(row, 4): diameter = table.OptionalDiameter(row, 3, units)
        If count = 0 Or diameter = 0# Then GoTo NextFace
        axisDistance = table.Length(row, 2, units)
        Require table, row, 2, axisDistance > diameter / 2#, "Расстояние до оси должно быть больше половины диаметра стержня."
        Set path = AxisPath(facade, sizes, groupID, axisDistance, table.SourceReference(row, 2))  ' Получаем линию центров стержней выбранной стороны.
        nx = 0#: ny = 0#
        Select Case groupID
            Case "B.Top": ny = 1#: normalSign = -1#
            Case "B.Bottom": ny = -1#: normalSign = 1#
            Case "H.Left": nx = -1#: normalSign = 1#
            Case "H.Right": nx = 1#: normalSign = -1#
        End Select
        Set options = New CRebarRuleOptions                                             ' Создаём условия размещения арматуры.
        If index >= 2 Then options.EndMode = ExcludeBoth                                ' Исключаем стержни на обоих концах линии.
        startColumn = table.ColumnByCaption(sizes.MainHeader, "StartSetback")
        endColumn = table.ColumnByCaption(sizes.MainHeader, "EndSetback")
        options.StartSetback = table.Length(row, startColumn, units)                    ' Задаём отступ начала ряда от начала грани.
        options.EndSetback = table.Length(row, endColumn, units)                        ' Задаём отступ конца ряда от конца грани.
        Require table, row, startColumn, options.StartSetback >= 0#, "StartSetback не может быть отрицательным."
        Require table, row, endColumn, options.EndSetback >= 0#, "EndSetback не может быть отрицательным."
        options.Publication = PublishPerSourcePosition: options.Coincidence = KeepPublishedOwner  ' Сохраняем имена и владельца совпавшего стержня.
        options.IncomingTangentAtJoint = True                                           ' На стыках берём входящую касательную; задаём подпись.
        options.AnnotationGroup = groupID: options.Comment = "RoundedRectangle: " & CStr(captions(index)) & " грань, ряд 1"  ' Задаём общую группу подписи стержней.
        options.SourceReference = table.SourceReference(row, 2) & "; " & table.SourceReference(row, startColumn) & "; " & table.SourceReference(row, endColumn)  ' Привязываем ошибки размещения к ячейкам ввода.
        facade.RebarByCount groupID & ".Main", path.ID, count, diameter, options        ' Задаём основной ряд по числу и диаметру стержней.
        extra = table.RowByAnyKey(Array(captions(index), Replace$(groupID, ".", "-"), groupID), 1, sizes.ExtraHeader + 1, sizes.ExtraHeader + 4)
        d2 = table.OptionalDiameter(extra, 2, units): d3 = table.OptionalDiameter(extra, 5, units)
        If d2 > 0# Then
            loc2 = table.Choice(extra, 3, Array("Stacked", "SideBySide"))
            Set options = AdditionalOptions(table, extra, 2, groupID, CStr(captions(index)), 2, loc2, diameter / 2# + d2 / 2#, normalSign)  ' Получаем смещение и привязку дополнительного ряда.
            facade.RebarDependent groupID & ".Row2", groupID & ".Main", d2, options     ' Задаём второй ряд относительно основного.
        End If
        If d3 > 0# Then
            loc3 = table.Choice(extra, 6, Array("Stacked", "SideBySide"))
            distance = diameter / 2# + d3 / 2#
            If d2 > 0# And loc2 = loc3 Then distance = distance + d2
            Set options = AdditionalOptions(table, extra, 5, groupID, CStr(captions(index)), 3, loc3, distance, normalSign)  ' Получаем смещение и привязку дополнительного ряда.
            facade.RebarDependent groupID & ".Row3", groupID & ".Main", d3, options     ' Задаём третий ряд относительно основного.
        End If
        If index >= 2 Then
            facade.RebarRowExtentLabel groupID, groupID & ".Main", nx, ny, axisDistance, True  ' Привязываем подпись к крайним стержням основного ряда.
        Else
            facade.RebarLabel groupID, path.ID, nx, ny, axisDistance                    ' Объявляем общую подпись количества и диаметров группы.
        End If
NextFace:
    Next index
End Sub

' B занимает только прямой участок. H включает все боковые сопряжения;
' направление у обеих сторон сверху вниз, поэтому SideBySide однозначен.
Private Function AxisPath(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRoundedInput, _
        ByVal groupID As String, ByVal axisDistance As Double, ByVal sourceReference As String) As CGeometryPath
    On Error GoTo InvalidPath
    Dim offset As CGeometryPath                                                         ' Объявляем линии и параметры геометрии.
    Set offset = facade.OffsetRoundedPolygon("Offset." & groupID, sizes.Vertices, sizes.Radii, sizes.VertexIDs, axisDistance)  ' Смещаем контур к линии центров стержней.
    Select Case groupID
        Case "B.Top": Set AxisPath = facade.Chain(groupID, offset.ID, "Edge.TL", 0#, "Edge.TL", 1#)  ' Выделяем прямую линию верхнего ряда.
        Case "B.Bottom": Set AxisPath = facade.Chain(groupID, offset.ID, "Edge.BR", 1#, "Edge.BR", 0#, False)  ' Выделяем прямую линию нижнего ряда.
        Case "H.Left"
            Set AxisPath = facade.Chain(groupID, offset.ID, "Edge.TL", 0#, "Edge.BR", 1#, False)  ' Выделяем левую линию ряда со скруглениями.
        Case "H.Right"
            Set AxisPath = facade.Chain(groupID, offset.ID, "Edge.TL", 1#, "Edge.BR", 0#)  ' Выделяем правую линию ряда со скруглениями.
    End Select
    Exit Function
InvalidPath:
    SectionRecipeGeometryError "CSectionRecipeRoundedRectangle", sourceReference
End Function

' Описывает смещение и привязку дополнительного ряда относительно основного, сохраняя смысловую
' группу.
Private Function AdditionalOptions(ByVal table As CSectionInputTable, ByVal row As Long, ByVal column As Long, _
        ByVal groupID As String, ByVal caption As String, ByVal rowNumber As Long, ByVal location As String, _
        ByVal distance As Double, ByVal normalSign As Double) As CRebarRuleOptions
    Dim options As CRebarRuleOptions                                                    ' Объявляем условия размещения арматуры.
    Set options = New CRebarRuleOptions                                                 ' Создаём условия размещения арматуры.
    If location = "Stacked" Then options.NormalOffset = normalSign * distance Else options.TangentOffset = distance  ' Выбираем смещение поперёк ряда или вдоль него.
    options.EverySecond = (table.Choice(row, column + 2, Array("EachBar", "EverySecondBar")) = "EverySecondBar")  ' Выбираем все исходные позиции или через одну.
    options.Publication = PublishPerSourcePosition: options.Coincidence = KeepPublishedOwner: options.AnnotationGroup = groupID  ' Задаём имена стержней, владельца и общую подпись.
    options.Comment = "RoundedRectangle: " & caption & " грань, ряд " & CStr(rowNumber)  ' Задаём пояснение к стержням этого ряда.
    options.SourceReference = table.SourceReference(row, column)                        ' Привязываем ошибки размещения к ячейкам ввода.
    Set AdditionalOptions = options                                                     ' Возвращаем условия размещения дополнительного ряда.
End Function

' Описывает габариты и размерные привязки рецепта отдельно от расчётных контуров.
Private Sub AddDimensions(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRoundedInput, ByVal outer As CGeometryPath)
    Dim minX As Double, maxX As Double
    outer.ProjectionBounds 1#, 0#, minX, maxX
    facade.HorizontalDistance "DIM_B", "B", -sizes.B / 2#, sizes.B / 2#, sizes.H / 2#, 1#, "Размер горизонтальной базы RoundedRectangle между левым и правым торцом."  ' Добавляем горизонтальный размер ширины B.
    facade.VerticalExtent "DIM_H", outer.ID, "H", minX, -1#, "Полная высота RoundedRectangle."  ' Добавляем размер полной высоты сечения.
    If sizes.SideTypes(1) = "Tapered" And -sizes.B / 2# - minX > GEOM_TOLERANCE Then facade.HorizontalDistance "DIM_W_LEFT", "W'", minX, -sizes.B / 2#, sizes.H / 2#, 1#, "Фактический вынос левого края бетона относительно левого торца базы B.", 1#  ' Добавляем размер левого выступа.
    If sizes.SideTypes(2) = "Tapered" And maxX - sizes.B / 2# > GEOM_TOLERANCE Then facade.HorizontalDistance "DIM_W_RIGHT", "W'", sizes.B / 2#, maxX, sizes.H / 2#, 1#, "Фактический вынос правого края бетона относительно правого торца базы B.", 1#  ' Добавляем размер правого выступа.
End Sub

' Проверяет ограничение рецепта; ошибка содержит фактический адрес параметра во входной таблице.
Private Sub Require(ByVal table As CSectionInputTable, ByVal row As Long, ByVal column As Long, ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then SectionRecipeError "CSectionRecipeRoundedRectangle", reason, table.SourceReference(row, column)
End Sub
