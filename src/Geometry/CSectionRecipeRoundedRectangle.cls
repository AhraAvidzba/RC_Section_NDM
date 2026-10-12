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
    Dim table As CSectionInputTable, units As CUnitSystem, facade As CSectionDefinitionBuilder
    Dim sizes As TRoundedInput, outer As CGeometryPath
    Set table = context.Table("rngRoundedRectangleGeometry"): Set units = context.Units
    ReadDimensions table, units, sizes
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "RoundedRectangle"
    Set outer = AddGeometry(facade, sizes)
    AddReinforcement facade, table, units, sizes
    AddDimensions facade, sizes, outer
    Set ISectionRecipe_CreateDefinition = facade.Finish
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
    AddCorner points, radii, names, count, "TL", -sizes.B / 2#, sizes.H / 2#, sizes.R1(1)
    AddCorner points, radii, names, count, "TR", sizes.B / 2#, sizes.H / 2#, sizes.R1(2)
    If sizes.SideTypes(2) = "Tapered" Then AddCorner points, radii, names, count, "RightNose", sizes.B / 2# + sizes.Widths(2), 0#, sizes.R2(2)
    AddCorner points, radii, names, count, "BR", sizes.B / 2#, -sizes.H / 2#, sizes.R1(2)
    AddCorner points, radii, names, count, "BL", -sizes.B / 2#, -sizes.H / 2#, sizes.R1(1)
    If sizes.SideTypes(1) = "Tapered" Then AddCorner points, radii, names, count, "LeftNose", -sizes.B / 2# - sizes.Widths(1), 0#, sizes.R2(1)
    Dim vertices() As Variant, assignedRadii() As Variant, vertexIDs() As Variant, i As Long
    ReDim vertices(0 To count - 1): ReDim assignedRadii(0 To count - 1): ReDim vertexIDs(0 To count - 1)
    For i = 0 To count - 1: vertices(i) = points(i): assignedRadii(i) = radii(i): vertexIDs(i) = names(i): Next i
    sizes.Vertices = vertices: sizes.Radii = assignedRadii: sizes.VertexIDs = vertexIDs
    Dim original As CGeometryPath, outer As CGeometryPath, first As String, edge As TRegionEdge, contourIDs() As String
    Set original = facade.RoundedPolygon("Shape", sizes.Vertices, sizes.Radii, sizes.VertexIDs)
    If sizes.SideTypes(1) = "Tapered" Then first = "Edge.LeftNose" Else first = "Edge.BL"
    If Not original.HasSegmentID(first) Then first = "Corner.TL"
    Set outer = facade.Chain("Boundary", original.ID, first, 0#, first, 0#, True, True)
    ReDim contourIDs(0 To outer.SegmentCount - 1)
    For i = 1 To outer.SegmentCount
        outer.CopySegment i, edge
        If edge.Sweep = 0# Then contourIDs(i - 1) = "CONTOUR_LINE_" & CStr(i) Else contourIDs(i - 1) = "CONTOUR_ARC_" & CStr(i)
    Next i
    facade.Outer "OUTER", outer.ID, "Прямой участок контура скругленного прямоугольника.", contourIDs, _
        "Дуга контура скругленного прямоугольника."
    Set AddGeometry = outer
    Exit Function
InvalidGeometry:
    SectionRecipeGeometryError "CSectionRecipeRoundedRectangle", sizes.GeometrySource
End Function

' Добавляет исходную вершину, её смысловой ID и назначенный радиус скругления.
Private Sub AddCorner(ByRef points() As Variant, ByRef radii() As Variant, ByRef names() As Variant, _
        ByRef count As Long, ByVal name As String, ByVal x As Double, ByVal y As Double, ByVal radius As Double)
    points(count) = Array(x, y): radii(count) = radius: names(count) = name: count = count + 1
End Sub

' Порядок граней сохраняет владельца совпавшего стержня. Дополнительные
' ряды исполняются вместе со своей исходной позицией, без повторной раскладки.
Private Sub AddReinforcement(ByVal facade As CSectionDefinitionBuilder, ByVal table As CSectionInputTable, _
        ByVal units As CUnitSystem, ByRef sizes As TRoundedInput)
    Dim faces As Variant, captions As Variant, index As Long, row As Long, extra As Long, count As Long
    Dim groupID As String, diameter As Double, axisDistance As Double, options As CRebarRuleOptions, path As CGeometryPath
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
        Set path = AxisPath(facade, sizes, groupID, axisDistance, table.SourceReference(row, 2))
        nx = 0#: ny = 0#
        Select Case groupID
            Case "B.Top": ny = 1#: normalSign = -1#
            Case "B.Bottom": ny = -1#: normalSign = 1#
            Case "H.Left": nx = -1#: normalSign = 1#
            Case "H.Right": nx = 1#: normalSign = -1#
        End Select
        Set options = New CRebarRuleOptions
        If index >= 2 Then options.EndMode = ExcludeBoth
        startColumn = table.ColumnByCaption(sizes.MainHeader, "StartSetback")
        endColumn = table.ColumnByCaption(sizes.MainHeader, "EndSetback")
        options.StartSetback = table.Length(row, startColumn, units)
        options.EndSetback = table.Length(row, endColumn, units)
        Require table, row, startColumn, options.StartSetback >= 0#, "StartSetback не может быть отрицательным."
        Require table, row, endColumn, options.EndSetback >= 0#, "EndSetback не может быть отрицательным."
        options.Publication = PublishPerSourcePosition: options.Coincidence = KeepPublishedOwner
        options.IncomingTangentAtJoint = True
        options.AnnotationGroup = groupID: options.Comment = "RoundedRectangle: " & CStr(captions(index)) & " грань, ряд 1"
        options.SourceReference = table.SourceReference(row, 2) & "; " & table.SourceReference(row, startColumn) & "; " & table.SourceReference(row, endColumn)
        facade.RebarByCount groupID & ".Main", path.ID, count, diameter, options
        extra = table.RowByAnyKey(Array(captions(index), Replace$(groupID, ".", "-"), groupID), 1, sizes.ExtraHeader + 1, sizes.ExtraHeader + 4)
        d2 = table.OptionalDiameter(extra, 2, units): d3 = table.OptionalDiameter(extra, 5, units)
        If d2 > 0# Then
            loc2 = table.Choice(extra, 3, Array("Stacked", "SideBySide"))
            Set options = AdditionalOptions(table, extra, 2, groupID, CStr(captions(index)), 2, loc2, diameter / 2# + d2 / 2#, normalSign)
            facade.RebarDependent groupID & ".Row2", groupID & ".Main", d2, options
        End If
        If d3 > 0# Then
            loc3 = table.Choice(extra, 6, Array("Stacked", "SideBySide"))
            distance = diameter / 2# + d3 / 2#
            If d2 > 0# And loc2 = loc3 Then distance = distance + d2
            Set options = AdditionalOptions(table, extra, 5, groupID, CStr(captions(index)), 3, loc3, distance, normalSign)
            facade.RebarDependent groupID & ".Row3", groupID & ".Main", d3, options
        End If
        If index >= 2 Then
            facade.RebarRowExtentLabel groupID, groupID & ".Main", nx, ny, axisDistance, True
        Else
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
    Dim offset As CGeometryPath, startID As String, endID As String, startAt As Double, endAt As Double
    Set offset = facade.OffsetRoundedPolygon("Offset." & groupID, sizes.Vertices, sizes.Radii, sizes.VertexIDs, axisDistance)
    Select Case groupID
        Case "B.Top": Set AxisPath = facade.Chain(groupID, offset.ID, "Edge.TL", 0#, "Edge.TL", 1#)
        Case "B.Bottom": Set AxisPath = facade.Chain(groupID, offset.ID, "Edge.BR", 1#, "Edge.BR", 0#, False)
        Case "H.Left"
            startID = "Edge.TL": startAt = 0#: endID = "Edge.BR": endAt = 1#
            If Not offset.HasSegmentID(startID) Then startID = "Corner.TL": startAt = 1#
            If Not offset.HasSegmentID(endID) Then endID = "Corner.BL": endAt = 0#
            Set AxisPath = facade.Chain(groupID, offset.ID, startID, startAt, endID, endAt, False)
        Case "H.Right"
            startID = "Edge.TL": startAt = 1#: endID = "Edge.BR": endAt = 0#
            If Not offset.HasSegmentID(startID) Then startID = "Corner.TR": startAt = 0#
            If Not offset.HasSegmentID(endID) Then endID = "Corner.BR": endAt = 1#
            Set AxisPath = facade.Chain(groupID, offset.ID, startID, startAt, endID, endAt)
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
    Dim options As CRebarRuleOptions
    Set options = New CRebarRuleOptions
    If location = "Stacked" Then options.NormalOffset = normalSign * distance Else options.TangentOffset = distance
    options.EverySecond = (table.Choice(row, column + 2, Array("EachBar", "EverySecondBar")) = "EverySecondBar")
    options.Publication = PublishPerSourcePosition: options.Coincidence = KeepPublishedOwner: options.AnnotationGroup = groupID
    options.Comment = "RoundedRectangle: " & caption & " грань, ряд " & CStr(rowNumber)
    options.SourceReference = table.SourceReference(row, column)
    Set AdditionalOptions = options
End Function

' Описывает габариты и размерные привязки рецепта отдельно от расчётных контуров.
Private Sub AddDimensions(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRoundedInput, ByVal outer As CGeometryPath)
    Dim minX As Double, maxX As Double
    outer.ProjectionBounds 1#, 0#, minX, maxX
    facade.HorizontalDistance "DIM_B", "B", -sizes.B / 2#, sizes.B / 2#, sizes.H / 2#, 1#, _
        "Размер горизонтальной базы RoundedRectangle между левым и правым торцом."
    facade.VerticalExtent "DIM_H", outer.ID, "H", minX, -1#, "Полная высота RoundedRectangle."
    If sizes.SideTypes(1) = "Tapered" And -sizes.B / 2# - minX > GEOM_TOLERANCE Then _
        facade.HorizontalDistance "DIM_W_LEFT", "W'", minX, -sizes.B / 2#, sizes.H / 2#, 1#, _
            "Фактический вынос левого края бетона относительно левого торца базы B.", 1#
    If sizes.SideTypes(2) = "Tapered" And maxX - sizes.B / 2# > GEOM_TOLERANCE Then _
        facade.HorizontalDistance "DIM_W_RIGHT", "W'", sizes.B / 2#, maxX, sizes.H / 2#, 1#, _
            "Фактический вынос правого края бетона относительно правого торца базы B.", 1#
End Sub

' Проверяет ограничение рецепта; ошибка содержит фактический адрес параметра во входной таблице.
Private Sub Require(ByVal table As CSectionInputTable, ByVal row As Long, ByVal column As Long, ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then SectionRecipeError "CSectionRecipeRoundedRectangle", reason, table.SourceReference(row, column)
End Sub
