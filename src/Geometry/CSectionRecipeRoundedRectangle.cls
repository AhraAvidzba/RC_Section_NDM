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
    MainHeader As Long
    ExtraHeader As Long
    GeometrySource As String
End Type

' Описывает контуры, арматурные правила и размеры выбранного рецепта. Сетку и физические стержни
' строит общий CSectionBuilder.
Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
    Dim table As CSectionInputTable, units As CUnitSystem, facade As CSectionDefinitionBuilder  ' Объявляем сборщик описания сечения.
    Dim sizes As TRoundedInput, outer As CGeometryPath, polygon As CPolygonSpec  ' Объявляем линии и параметры геометрии.
    Set table = context.Table("rngRoundedRectangleGeometry"): Set units = context.Units
    ReadDimensions table, units, sizes
    Set polygon = DescribePolygon(sizes) ' Один исходный многоугольник для контура и всех осевых смещений.
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "RoundedRectangle"  ' "RoundedRectangle" — имя типа сечения.
    Set outer = AddGeometry(facade, sizes, polygon)  ' facade — сборщик описания сечения; sizes — подготовленные размеры и параметры формы; polygon — завершённые исходные вершины с радиусами.
    ' facade — сборщик описания сечения; table — снимок таблицы армирования; units — перевод исходных
    ' размеров в мм; sizes — размеры и данные формы; polygon — завершённые исходные вершины с
    ' радиусами.
    AddReinforcement facade, table, units, sizes, polygon
    AddDimensions facade, sizes, outer  ' facade — сборщик описания сечения; sizes — размеры и данные формы; outer — наружная граница для габаритов.
    Set ISectionRecipe_CreateDefinition = facade.Finish  ' Проверяем и возвращаем готовое описание сечения.
End Function

' Неактивные W/R2 стороны Simple не читаются; ошибки относятся к исходной
' ячейке, единицы переводятся один раз через CUnitSystem.
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

' Создаёт и публикует контур из одного описателя исходных вершин.
' Начало опубликованного контура задано смысловой гранью, а не индексом дуги.
Private Function AddGeometry(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRoundedInput, _
        ByVal polygon As CPolygonSpec) As CGeometryPath
    On Error GoTo InvalidGeometry
    Dim original As CGeometryPath, outer As CGeometryPath, first As String
    Set original = facade.RoundedPolygon("Shape", polygon)  ' "Shape" — имя контура; polygon — завершённое описание исходных вершин и радиусов, мм.
    If sizes.SideTypes(1) = "Tapered" Then first = "Edge.LeftNose" Else first = "Edge.BL"
    Set outer = facade.Chain("Boundary", original.ID, first, 0#, first, 0#, True, True) ' Полный обход от прежней грани.
    ' "OUTER" — имя наружной границы бетона; outer.ID — замкнутый контур; comment — пояснение к
    ' границе; arcComment — отдельное пояснение к дугам; SequentialByKind — способ публикации имён
    ' участков.
    facade.Outer "OUTER", outer.ID, "Прямой участок контура скругленного прямоугольника.", _
        arcComment:="Дуга контура скругленного прямоугольника.", namingPolicy:=SequentialByKind
    Set AddGeometry = outer
    Exit Function
InvalidGeometry:
    SectionRecipeGeometryError "CSectionRecipeRoundedRectangle", sizes.GeometrySource
End Function

' Добавляет исходную вершину, её смысловой ID и назначенный радиус скругления.
Private Function DescribePolygon(ByRef sizes As TRoundedInput) As CPolygonSpec
    Dim polygon As CPolygonSpec
    Set polygon = New CPolygonSpec
    polygon.AddVertex "TL", -sizes.B / 2#, sizes.H / 2#, sizes.R1(1), sizes.GeometrySource ' Верхний левый угол до скругления.
    polygon.AddVertex "TR", sizes.B / 2#, sizes.H / 2#, sizes.R1(2), sizes.GeometrySource ' Верхний правый угол до скругления.
    If sizes.SideTypes(2) = "Tapered" Then _
        polygon.AddVertex "RightNose", sizes.B / 2# + sizes.Widths(2), 0#, sizes.R2(2), sizes.GeometrySource
    polygon.AddVertex "BR", sizes.B / 2#, -sizes.H / 2#, sizes.R1(2), sizes.GeometrySource ' Нижний правый угол до скругления.
    polygon.AddVertex "BL", -sizes.B / 2#, -sizes.H / 2#, sizes.R1(1), sizes.GeometrySource ' Нижний левый угол до скругления.
    If sizes.SideTypes(1) = "Tapered" Then _
        polygon.AddVertex "LeftNose", -sizes.B / 2# - sizes.Widths(1), 0#, sizes.R2(1), sizes.GeometrySource
    polygon.Seal
    Set DescribePolygon = polygon
End Function

' Порядок граней сохраняет владельца совпавшего стержня. Дополнительные
' ряды исполняются вместе со своей исходной позицией, без повторной раскладки.
Private Sub AddReinforcement(ByVal facade As CSectionDefinitionBuilder, ByVal table As CSectionInputTable, _
        ByVal units As CUnitSystem, ByRef sizes As TRoundedInput, ByVal polygon As CPolygonSpec)
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
        ' facade — сборщик описания сечения; polygon — завершённые исходные вершины с радиусами;
        ' groupID — имя стороны; axisDistance — as до центров стержней, мм; table.SourceReference(row,
        ' 2) — адрес исходной ячейки для диагностики.
        Set path = AxisPath(facade, polygon, groupID, axisDistance, table.SourceReference(row, 2))
        nx = 0#: ny = 0#
        Select Case groupID
            Case "B.Top": ny = 1#: normalSign = -1#
            Case "B.Bottom": ny = -1#: normalSign = 1#
            Case "H.Left": nx = -1#: normalSign = 1#
            Case "H.Right": nx = 1#: normalSign = -1#
        End Select
        Set options = New CRebarRuleOptions  ' Создаём условия размещения арматуры.
        If index >= 2 Then options.EndMode = ExcludeBoth  ' EndMode — ExcludeBoth исключает точки обоих концов линии.
        If startColumn = 0 Then
            startColumn = table.ColumnByCaption(sizes.MainHeader, "StartSetback")
            endColumn = table.ColumnByCaption(sizes.MainHeader, "EndSetback")
        End If
        options.StartSetback = table.NonNegativeLength(row, startColumn, units)  ' StartSetback — отступ начала ряда от начала грани, мм.
        options.EndSetback = table.NonNegativeLength(row, endColumn, units)  ' EndSetback — отступ конца ряда от конца грани, мм.
        ' Publication — PublishPerSourcePosition сохраняет имена по исходным позициям; Coincidence —
        ' KeepPublishedOwner сохраняет первого владельца совпавшего стержня.
        options.SetPublication groupID, PublishPerSourcePosition, KeepPublishedOwner
        options.IncomingTangentAtJoint = True  ' IncomingTangentAtJoint — True использует входящую касательную на стыке участков.
        options.Comment = "RoundedRectangle: " & CStr(captions(index)) & " грань, ряд 1"  ' AnnotationGroup — имя общей подписи для всех рядов этой группы; Comment — пояснение к стержням данного ряда.
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
            ' номер дополнительного ряда; loc2 — Stacked поперёк ряда, SideBySide вдоль; diameter —
            ' диаметр основного стержня, мм; d2 — диаметр добавляемого стержня, мм; normalSign — знак
            ' смещения по нормали.
            Set options = AdditionalOptions(table, extra, 2, groupID, CStr(captions(index)), 2, loc2, diameter, d2, normalSign)
            ' groupID & ".Row2" — имя дополнительного ряда; groupID & ".Main" — исходный ряд, чьи
            ' позиции используются; d2 — диаметр дополнительного стержня, мм; options — смещение и
            ' выбор позиций исходного ряда.
            facade.RebarDependent groupID & ".Row2", groupID & ".Main", d2, options
        End If
        If d3 > 0# Then
            loc3 = table.Choice(extra, 6, Array("Stacked", "SideBySide"))
            distance = 0#
            If d2 > 0# And loc2 = loc3 Then distance = d2
            ' table — снимок таблицы армирования; extra — строка дополнительных рядов; 5 — столбец их
            ' диаметра; groupID — общая группа подписи; CStr(captions(index)) — название стороны; 3 —
            ' номер дополнительного ряда; loc3 — Stacked поперёк ряда, SideBySide вдоль; diameter —
            ' диаметр основного стержня, мм; d3 — диаметр добавляемого стержня, мм; normalSign — знак
            ' смещения по нормали; distance — сумма промежуточных номинальных диаметров, мм.
            Set options = AdditionalOptions(table, extra, 5, groupID, CStr(captions(index)), 3, loc3, diameter, d3, normalSign, distance)
            ' groupID & ".Row3" — имя дополнительного ряда; groupID & ".Main" — исходный ряд, чьи
            ' позиции используются; d3 — диаметр дополнительного стержня, мм; options — смещение и
            ' выбор позиций исходного ряда.
            facade.RebarDependent groupID & ".Row3", groupID & ".Main", d3, options
        End If
        If index >= 2 Then
            ' groupID — группа общей подписи; groupID & ".Main" — ряд для определения крайних
            ' стержней; nx — X-компонента направления подписи; ny — Y-компонента направления подписи;
            ' True — меняем местами концы опорной линии подписи.
            facade.RebarRowExtentLabel groupID, groupID & ".Main", nx, ny, True
        Else
            ' groupID — группа общей подписи всех её рядов; path.ID — линия, по габаритам которой
            ' привязывается подпись; nx — X-компонента направления подписи; ny — Y-компонента
            ' направления подписи.
            facade.RebarLabel groupID, path.ID, nx, ny
        End If
NextFace:
    Next index
End Sub

' Исходные вершины хранятся отдельно от точек касания и привязок аннотаций.
' B занимает только прямой участок. H включает все боковые сопряжения;
' направление у обеих сторон сверху вниз, поэтому SideBySide однозначен.
Private Function AxisPath(ByVal facade As CSectionDefinitionBuilder, ByVal polygon As CPolygonSpec, _
        ByVal groupID As String, ByVal axisDistance As Double, ByVal sourceReference As String) As CGeometryPath
    On Error GoTo InvalidPath
    Dim offset As CGeometryPath  ' Объявляем линии и параметры геометрии.
    ' "Offset." & groupID — имя линии центров стержней; polygon — то же завершённое описание исходных
    ' вершин; axisDistance — смещение, мм: положительное внутрь, отрицательное наружу.
    Set offset = facade.OffsetRoundedPolygon("Offset." & groupID, polygon, axisDistance)
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
        ByVal baseDiameter As Double, ByVal diameter As Double, ByVal normalSign As Double, _
        Optional ByVal interveningDiameterSum As Double = 0#) As CRebarRuleOptions
    Dim options As CRebarRuleOptions  ' Объявляем условия размещения арматуры.
    Set options = New CRebarRuleOptions  ' Создаём условия размещения арматуры.
    ' NormalOffset — знаковое смещение по нормали к исходному ряду, мм; TangentOffset — знаковое
    ' смещение вдоль исходного ряда, мм.
    Dim axis As ERebarOffsetAxis, directionSign As Double, everySecond As Boolean
    axis = RowNormal: directionSign = normalSign
    If location = "SideBySide" Then axis = RowTangent: directionSign = 1#
    everySecond = (table.Choice(row, column + 2, Array("EachBar", "EverySecondBar")) = "EverySecondBar")
    options.SourceReference = table.SourceReference(row, column)
    options.ConfigureDependentContact baseDiameter, diameter, axis, directionSign, everySecond, interveningDiameterSum
    options.SetPublication groupID, PublishPerSourcePosition, KeepPublishedOwner
    options.Comment = "RoundedRectangle: " & caption & " грань, ряд " & CStr(rowNumber)  ' Comment — пояснение к стержням данного ряда.
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
