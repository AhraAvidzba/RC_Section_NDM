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
    Dim sizes As TRoundedInput, outer As CGeometryPath  ' Объявляем линии и параметры геометрии.
    Set table = context.Table("rngRoundedRectangleGeometry"): Set units = context.Units
    ReadDimensions table, units, sizes
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "RoundedRectangle"  ' "RoundedRectangle" — имя типа сечения.
    Set outer = AddGeometry(facade, sizes)  ' facade — сборщик описания сечения; sizes — подготовленные размеры и радиусы.
    AddReinforcement facade, table, units, sizes  ' facade — сборщик описания сечения; table — снимок таблицы армирования; units — перевод исходных размеров в мм; sizes — размеры и данные формы.
    AddDimensions facade, sizes, outer  ' facade — сборщик описания сечения; sizes — размеры и данные формы; outer — наружная граница для габаритов.
    Set ISectionRecipe_CreateDefinition = facade.Finish  ' Проверяем и возвращаем готовое описание сечения.
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
    ' points — массив координат вершин; radii — массив радиусов; names — массив имён вершин; count —
    ' число уже заданных вершин; увеличивается на одну; "TL" — имя добавляемой вершины; -sizes.B / 2 —
    ' X вершины, мм; sizes.H / 2 — Y вершины, мм; sizes.R1(1) — радиус скругления, мм.
    AddCorner points, radii, names, count, "TL", -sizes.B / 2#, sizes.H / 2#, sizes.R1(1)
    ' points — массив координат вершин; radii — массив радиусов; names — массив имён вершин; count —
    ' число уже заданных вершин; увеличивается на одну; "TR" — имя добавляемой вершины; sizes.B / 2 —
    ' X вершины, мм; sizes.H / 2 — Y вершины, мм; sizes.R1(2) — радиус скругления, мм.
    AddCorner points, radii, names, count, "TR", sizes.B / 2#, sizes.H / 2#, sizes.R1(2)
    ' points — массив координат вершин; radii — массив радиусов; names — массив имён вершин; count —
    ' число уже заданных вершин; увеличивается на одну; "RightNose" — имя добавляемой вершины; sizes.B
    ' / 2 + sizes.Widths(2) — X вершины, мм; 0 — Y вершины, мм; sizes.R2(2) — радиус скругления, мм.
    If sizes.SideTypes(2) = "Tapered" Then AddCorner points, radii, names, count, "RightNose", sizes.B / 2# + sizes.Widths(2), 0#, sizes.R2(2)
    ' points — массив координат вершин; radii — массив радиусов; names — массив имён вершин; count —
    ' число уже заданных вершин; увеличивается на одну; "BR" — имя добавляемой вершины; sizes.B / 2 —
    ' X вершины, мм; -sizes.H / 2 — Y вершины, мм; sizes.R1(2) — радиус скругления, мм.
    AddCorner points, radii, names, count, "BR", sizes.B / 2#, -sizes.H / 2#, sizes.R1(2)
    ' points — массив координат вершин; radii — массив радиусов; names — массив имён вершин; count —
    ' число уже заданных вершин; увеличивается на одну; "BL" — имя добавляемой вершины; -sizes.B / 2 —
    ' X вершины, мм; -sizes.H / 2 — Y вершины, мм; sizes.R1(1) — радиус скругления, мм.
    AddCorner points, radii, names, count, "BL", -sizes.B / 2#, -sizes.H / 2#, sizes.R1(1)
    ' points — массив координат вершин; radii — массив радиусов; names — массив имён вершин; count —
    ' число уже заданных вершин; увеличивается на одну; "LeftNose" — имя добавляемой вершины; -sizes.B
    ' / 2 - sizes.Widths(1) — X вершины, мм; 0 — Y вершины, мм; sizes.R2(1) — радиус скругления, мм.
    If sizes.SideTypes(1) = "Tapered" Then AddCorner points, radii, names, count, "LeftNose", -sizes.B / 2# - sizes.Widths(1), 0#, sizes.R2(1)
    Dim vertices() As Variant, assignedRadii() As Variant, vertexIDs() As Variant, i As Long
    ReDim vertices(0 To count - 1): ReDim assignedRadii(0 To count - 1): ReDim vertexIDs(0 To count - 1)  ' Готовим массивы вершин, радиусов и имён.
    For i = 0 To count - 1: vertices(i) = points(i): assignedRadii(i) = radii(i): vertexIDs(i) = names(i): Next i  ' Копируем подготовленные вершины, радиусы и имена.
    sizes.Vertices = vertices: sizes.Radii = assignedRadii: sizes.VertexIDs = vertexIDs  ' Сохраняем вершины, радиусы и имена контура.
    Dim original As CGeometryPath, outer As CGeometryPath, first As String, edge As TRegionEdge, contourIDs() As String  ' Объявляем линии и параметры геометрии.
    ' "Shape" — имя контура; sizes.Vertices — координаты вершин (X,Y), мм; sizes.Radii — радиус каждой
    ' вершины, мм; sizes.VertexIDs — имена вершин для выбора участков.
    Set original = facade.RoundedPolygon("Shape", sizes.Vertices, sizes.Radii, sizes.VertexIDs)
    If sizes.SideTypes(1) = "Tapered" Then first = "Edge.LeftNose" Else first = "Edge.BL"
    ' "Boundary" — имя новой линии; original.ID — исходный контур; first — участок начала; 0 — точка
    ' начала участка: 0 начало, 1 конец; first — участок конца; 0 — точка конца участка: 0 начало, 1
    ' конец; True — по обходу исходной линии; True — полный замкнутый контур при совпавших концах.
    Set outer = facade.Chain("Boundary", original.ID, first, 0#, first, 0#, True, True)
    ReDim contourIDs(0 To outer.SegmentCount - 1)  ' Готовим имена участков контура.
    For i = 1 To outer.SegmentCount
        outer.CopySegment i, edge  ' Получаем участок контура для назначения имени.
        If edge.Sweep = 0# Then contourIDs(i - 1) = "CONTOUR_LINE_" & CStr(i) Else contourIDs(i - 1) = "CONTOUR_ARC_" & CStr(i)  ' Назначаем имена прямым и дуговым участкам контура.
    Next i
    ' "OUTER" — имя наружной границы бетона; outer.ID — замкнутый контур; comment — пояснение к
    ' границе; contourIDs — имена отдельных участков контура; arcComment — отдельное пояснение к
    ' дугам.
    facade.Outer "OUTER", outer.ID, "Прямой участок контура скругленного прямоугольника.", contourIDs, "Дуга контура скругленного прямоугольника."
    Set AddGeometry = outer  ' Возвращаем готовую наружную границу.
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
        ' facade — сборщик описания сечения; sizes — размеры, вершины и радиусы контура; groupID — имя
        ' стороны; axisDistance — as до центров стержней, мм; table.SourceReference(row, 2) — адрес
        ' исходной ячейки для диагностики.
        Set path = AxisPath(facade, sizes, groupID, axisDistance, table.SourceReference(row, 2))
        nx = 0#: ny = 0#
        Select Case groupID
            Case "B.Top": ny = 1#: normalSign = -1#
            Case "B.Bottom": ny = -1#: normalSign = 1#
            Case "H.Left": nx = -1#: normalSign = 1#
            Case "H.Right": nx = 1#: normalSign = -1#
        End Select
        Set options = New CRebarRuleOptions  ' Создаём условия размещения арматуры.
        If index >= 2 Then options.EndMode = ExcludeBoth  ' EndMode — ExcludeBoth исключает точки обоих концов линии.
        startColumn = table.ColumnByCaption(sizes.MainHeader, "StartSetback")
        endColumn = table.ColumnByCaption(sizes.MainHeader, "EndSetback")
        options.StartSetback = table.Length(row, startColumn, units)  ' StartSetback — отступ начала ряда от начала грани, мм.
        options.EndSetback = table.Length(row, endColumn, units)  ' EndSetback — отступ конца ряда от конца грани, мм.
        Require table, row, startColumn, options.StartSetback >= 0#, "StartSetback не может быть отрицательным."
        Require table, row, endColumn, options.EndSetback >= 0#, "EndSetback не может быть отрицательным."
        ' Publication — PublishPerSourcePosition сохраняет имена по исходным позициям; Coincidence —
        ' KeepPublishedOwner сохраняет первого владельца совпавшего стержня.
        options.Publication = PublishPerSourcePosition: options.Coincidence = KeepPublishedOwner
        options.IncomingTangentAtJoint = True  ' IncomingTangentAtJoint — True использует входящую касательную на стыке участков.
        options.AnnotationGroup = groupID: options.Comment = "RoundedRectangle: " & CStr(captions(index)) & " грань, ряд 1"  ' AnnotationGroup — имя общей подписи для всех рядов этой группы; Comment — пояснение к стержням данного ряда.
        options.SourceReference = table.SourceReference(row, 2) & "; " & table.SourceReference(row, startColumn) & "; " & table.SourceReference(row, endColumn)  ' SourceReference — адреса ячеек, которые надо исправить при ошибке размещения.
        ' groupID & ".Main" — имя нового ряда; path.ID — линия центров стержней; count — заданное
        ' число стержней; diameter — диаметр стержня, мм; options — отступы, нумерация и условия
        ' размещения.
        facade.RebarByCount groupID & ".Main", path.ID, count, diameter, options
        extra = table.RowByAnyKey(Array(captions(index), Replace$(groupID, ".", "-"), groupID), 1, sizes.ExtraHeader + 1, sizes.ExtraHeader + 4)
        d2 = table.OptionalDiameter(extra, 2, units): d3 = table.OptionalDiameter(extra, 5, units)
        If d2 > 0# Then
            loc2 = table.Choice(extra, 3, Array("Stacked", "SideBySide"))
            ' table — снимок таблицы армирования; extra — строка дополнительных рядов; 2 — столбец их
            ' диаметра; groupID — общая группа подписи; CStr(captions(index)) — название стороны; 2 —
            ' номер дополнительного ряда; loc2 — Stacked поперёк ряда, SideBySide вдоль; diameter / 2
            ' + d2 / 2 — расстояние между центрами стержней, мм; normalSign — знак смещения по
            ' нормали.
            Set options = AdditionalOptions(table, extra, 2, groupID, CStr(captions(index)), 2, loc2, diameter / 2# + d2 / 2#, normalSign)
            ' groupID & ".Row2" — имя дополнительного ряда; groupID & ".Main" — исходный ряд, чьи
            ' позиции используются; d2 — диаметр дополнительного стержня, мм; options — смещение и
            ' выбор позиций исходного ряда.
            facade.RebarDependent groupID & ".Row2", groupID & ".Main", d2, options
        End If
        If d3 > 0# Then
            loc3 = table.Choice(extra, 6, Array("Stacked", "SideBySide"))
            distance = diameter / 2# + d3 / 2#
            If d2 > 0# And loc2 = loc3 Then distance = distance + d2
            ' table — снимок таблицы армирования; extra — строка дополнительных рядов; 5 — столбец их
            ' диаметра; groupID — общая группа подписи; CStr(captions(index)) — название стороны; 3 —
            ' номер дополнительного ряда; loc3 — Stacked поперёк ряда, SideBySide вдоль; distance —
            ' расстояние между центрами стержней, мм; normalSign — знак смещения по нормали.
            Set options = AdditionalOptions(table, extra, 5, groupID, CStr(captions(index)), 3, loc3, distance, normalSign)
            ' groupID & ".Row3" — имя дополнительного ряда; groupID & ".Main" — исходный ряд, чьи
            ' позиции используются; d3 — диаметр дополнительного стержня, мм; options — смещение и
            ' выбор позиций исходного ряда.
            facade.RebarDependent groupID & ".Row3", groupID & ".Main", d3, options
        End If
        If index >= 2 Then
            ' groupID — группа общей подписи; groupID & ".Main" — ряд для определения крайних
            ' стержней; nx — X-компонента направления подписи; ny — Y-компонента направления подписи;
            ' axisDistance — as до осей стержней, мм; отступ подписи задаёт плоттер; True — меняем
            ' местами концы опорной линии подписи.
            facade.RebarRowExtentLabel groupID, groupID & ".Main", nx, ny, axisDistance, True
        Else
            ' groupID — группа общей подписи всех её рядов; path.ID — линия, по габаритам которой
            ' привязывается подпись; nx — X-компонента направления подписи; ny — Y-компонента
            ' направления подписи; axisDistance — as до осей стержней, мм; отступ подписи задаёт
            ' плоттер.
            facade.RebarLabel groupID, path.ID, nx, ny, axisDistance
        End If
NextFace:
    Next index
End Sub

' B занимает только прямой участок. H включает все боковые сопряжения;
' направление у обеих сторон сверху вниз, поэтому SideBySide однозначен.
Private Function AxisPath(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRoundedInput, _
        ByVal groupID As String, ByVal axisDistance As Double, ByVal sourceReference As String) As CGeometryPath
    On Error GoTo InvalidPath
    Dim offset As CGeometryPath  ' Объявляем линии и параметры геометрии.
    ' "Offset." & groupID — имя линии центров стержней; sizes.Vertices — исходные вершины (X,Y), мм;
    ' sizes.Radii — исходные радиусы вершин, мм; sizes.VertexIDs — имена исходных вершин; axisDistance
    ' — смещение, мм: положительное внутрь, отрицательное наружу.
    Set offset = facade.OffsetRoundedPolygon("Offset." & groupID, sizes.Vertices, sizes.Radii, sizes.VertexIDs, axisDistance)
    Select Case groupID
        ' groupID — имя новой линии; offset.ID — исходный контур; "Edge.TL" — участок начала; 0 —
        ' точка начала участка: 0 начало, 1 конец; "Edge.TL" — участок конца; 1 — точка конца участка:
        ' 0 начало, 1 конец.
        Case "B.Top": Set AxisPath = facade.Chain(groupID, offset.ID, "Edge.TL", 0#, "Edge.TL", 1#)
        ' groupID — имя новой линии; offset.ID — исходный контур; "Edge.BR" — участок начала; 1 —
        ' точка начала участка: 0 начало, 1 конец; "Edge.BR" — участок конца; 0 — точка конца участка:
        ' 0 начало, 1 конец; False — против обхода исходной линии.
        Case "B.Bottom": Set AxisPath = facade.Chain(groupID, offset.ID, "Edge.BR", 1#, "Edge.BR", 0#, False)
        Case "H.Left"
            ' groupID — имя новой линии; offset.ID — исходный контур; "Edge.TL" — участок начала; 0 —
            ' точка начала участка: 0 начало, 1 конец; "Edge.BR" — участок конца; 1 — точка конца
            ' участка: 0 начало, 1 конец; False — против обхода исходной линии.
            Set AxisPath = facade.Chain(groupID, offset.ID, "Edge.TL", 0#, "Edge.BR", 1#, False)
        Case "H.Right"
            ' groupID — имя новой линии; offset.ID — исходный контур; "Edge.TL" — участок начала; 1 —
            ' точка начала участка: 0 начало, 1 конец; "Edge.BR" — участок конца; 0 — точка конца
            ' участка: 0 начало, 1 конец.
            Set AxisPath = facade.Chain(groupID, offset.ID, "Edge.TL", 1#, "Edge.BR", 0#)
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
    Dim options As CRebarRuleOptions  ' Объявляем условия размещения арматуры.
    Set options = New CRebarRuleOptions  ' Создаём условия размещения арматуры.
    ' NormalOffset — знаковое смещение по нормали к исходному ряду, мм; TangentOffset — знаковое
    ' смещение вдоль исходного ряда, мм.
    If location = "Stacked" Then options.NormalOffset = normalSign * distance Else options.TangentOffset = distance
    options.EverySecond = (table.Choice(row, column + 2, Array("EachBar", "EverySecondBar")) = "EverySecondBar")  ' EverySecond — True берёт первый, третий и следующие нечётные стержни; False все.
    ' Publication — PublishPerSourcePosition сохраняет имена по исходным позициям; Coincidence —
    ' KeepPublishedOwner сохраняет первого владельца совпавшего стержня; AnnotationGroup — имя общей
    ' подписи для всех рядов этой группы.
    options.Publication = PublishPerSourcePosition: options.Coincidence = KeepPublishedOwner: options.AnnotationGroup = groupID
    options.Comment = "RoundedRectangle: " & caption & " грань, ряд " & CStr(rowNumber)  ' Comment — пояснение к стержням данного ряда.
    options.SourceReference = table.SourceReference(row, column)  ' SourceReference — адреса ячеек, которые надо исправить при ошибке размещения.
    Set AdditionalOptions = options  ' Возвращаем условия размещения дополнительного ряда.
End Function

' Описывает габариты и размерные привязки рецепта отдельно от расчётных контуров.
Private Sub AddDimensions(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRoundedInput, ByVal outer As CGeometryPath)
    Dim minX As Double, maxX As Double
    outer.ProjectionBounds 1#, 0#, minX, maxX
    ' "DIM_B" — имя размера; "B" — обозначение на схеме; -sizes.B / 2 — X первой точки, мм; sizes.B /
    ' 2 — X второй точки, мм; sizes.H / 2 — уровень исходной размерной линии Y, мм; 1 — вынос размера
    ' вверх; comment — пояснение к размеру.
    facade.HorizontalDistance "DIM_B", "B", -sizes.B / 2#, sizes.B / 2#, sizes.H / 2#, 1#, "Размер горизонтальной базы RoundedRectangle между левым и правым торцом."
    ' "DIM_H" — имя размера; outer.ID — контур, по которому измеряется высота; "H" — обозначение на
    ' схеме; minX — положение исходной размерной линии X, мм; -1 — вынос размера влево; comment —
    ' пояснение к размеру.
    facade.VerticalExtent "DIM_H", outer.ID, "H", minX, -1#, "Полная высота RoundedRectangle."
    ' "DIM_W_LEFT" — имя размера; "W'" — обозначение на схеме; minX — X первой точки, мм; -sizes.B / 2
    ' — X второй точки, мм; sizes.H / 2 — уровень исходной размерной линии Y, мм; 1 — вынос размера
    ' вверх; comment — пояснение к размеру; 1 — шаг округления только подписи, мм.
    If sizes.SideTypes(1) = "Tapered" And -sizes.B / 2# - minX > GEOM_TOLERANCE Then facade.HorizontalDistance "DIM_W_LEFT", "W'", minX, -sizes.B / 2#, sizes.H / 2#, 1#, "Фактический вынос левого края бетона относительно левого торца базы B.", 1#
    ' "DIM_W_RIGHT" — имя размера; "W'" — обозначение на схеме; sizes.B / 2 — X первой точки, мм; maxX
    ' — X второй точки, мм; sizes.H / 2 — уровень исходной размерной линии Y, мм; 1 — вынос размера
    ' вверх; comment — пояснение к размеру; 1 — шаг округления только подписи, мм.
    If sizes.SideTypes(2) = "Tapered" And maxX - sizes.B / 2# > GEOM_TOLERANCE Then facade.HorizontalDistance "DIM_W_RIGHT", "W'", sizes.B / 2#, maxX, sizes.H / 2#, 1#, "Фактический вынос правого края бетона относительно правого торца базы B.", 1#
End Sub

' Проверяет ограничение рецепта; ошибка содержит фактический адрес параметра во входной таблице.
Private Sub Require(ByVal table As CSectionInputTable, ByVal row As Long, ByVal column As Long, ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then SectionRecipeError "CSectionRecipeRoundedRectangle", reason, table.SourceReference(row, column)
End Sub
