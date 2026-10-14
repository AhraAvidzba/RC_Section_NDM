Attribute VB_Name = "modTestGeometryFixtures"
Option Explicit

' ДЛЯ ТЕСТОВ: численные области во внутренних мм через актуальную общую геометрию.
' Не участвует в добавлении производственных типов сечений; workbook snapshot
' ниже читает только зарегистрированные таблицы через штатный reader.
Public Function FixtureGeometry(ByVal region As CConcreteRegion) As CGeometryRegion
    Set FixtureGeometry = New CGeometryRegion
    FixtureGeometry.Initialize region
End Function

' ДЛЯ ТЕСТОВ: перенос уже подготовленных ребер с сохранением дуг и отверстий.
Public Function FixtureMoveGeometry(ByVal definition As CSectionDefinition, ByVal x As Double, ByVal y As Double) As CGeometryRegion
    Dim edges() As TRegionEdge, count As Long, i As Long, query As CSectionGeometryQuery
    definition.ConcreteDomain.CopyEdges edges, count
    For i = 1 To count
        edges(i).X1 = edges(i).X1 + x: edges(i).X2 = edges(i).X2 + x
        edges(i).Y1 = edges(i).Y1 + y: edges(i).Y2 = edges(i).Y2 + y
        edges(i).CenterX = edges(i).CenterX + x: edges(i).CenterY = edges(i).CenterY + y
    Next i
    Set query = New CSectionGeometryQuery
    Set FixtureMoveGeometry = FixtureGeometry(query.BoundaryRegion(edges, count))
End Function

Public Sub FixturePathEnds(ByVal path As CGeometryPath, ByRef x1 As Double, ByRef y1 As Double, _
        ByRef x2 As Double, ByRef y2 As Double, ByRef tangentX As Double, ByRef tangentY As Double)
    path.AtDistance 0#, x1, y1, tangentX, tangentY
    path.AtDistance path.Length, x2, y2, tangentX, tangentY
End Sub

' ДЛЯ ТЕСТОВ: тот же исполнитель и annotation builder, что в общей сборке,
' без сетки; используется для проверки физических рядов и их привязок.
Public Function FixtureAnnotatedLayout(ByVal definition As CSectionDefinition, _
        Optional ByVal executor As CRebarRuleExecutor = Nothing, Optional ByRef annotations As CSectionAnnotations = Nothing) As CRebarLayout
    If executor Is Nothing Then Set executor = New CRebarRuleExecutor
    Dim builder As CSectionAnnotationBuilder, all As CSectionAnnotations, layout As CRebarLayout, i As Long
    Set layout = executor.Build(definition, FixtureGeometry(definition.ConcreteDomain))
    Set builder = New CSectionAnnotationBuilder: Set all = New CSectionAnnotations
    builder.Build definition, layout, all, executor
    Set annotations = New CSectionAnnotations
    For i = 1 To all.Count
        If all.AnnotationType(i) = "REBAR_ANNOTATION" Then FixtureCopyAnnotation all, i, annotations, 0#, 0#
    Next i
    Set FixtureAnnotatedLayout = layout
End Function

' ДЛЯ ТЕСТОВ: только перенос готовой раскладки, без распределения позиций.
Public Function FixtureMoveLayout(ByVal layout As CRebarLayout, ByVal x As Double, ByVal y As Double) As CRebarLayout
    If x = 0# And y = 0# Then Set FixtureMoveLayout = layout: Exit Function
    Dim result As CRebarLayout, i As Long: Set result = New CRebarLayout
    For i = 1 To layout.Count
        result.AddBar layout.BarID(i), layout.X(i) + x, layout.Y(i) + y, layout.Diameter(i), layout.Area(i), _
            layout.SteelClass(i), layout.Comment(i), annotationGroupName:=layout.BarAnnotationGroupName(i)
    Next i

    Set FixtureMoveLayout = result
End Function

' ДЛЯ ТЕСТОВ: перенос уже вычисленных аннотаций, без изменения стержней и вычисления привязок.
Public Sub FixtureMoveAnnotations(ByRef annotations As CSectionAnnotations, ByVal x As Double, ByVal y As Double)
    Dim result As CSectionAnnotations, i As Long
    Set result = New CSectionAnnotations
    For i = 1 To annotations.Count: FixtureCopyAnnotation annotations, i, result, x, y: Next i
    Set annotations = result
End Sub

' Копирует полную запись аннотации и переносит только её геометрические точки.
Private Sub FixtureCopyAnnotation(ByVal source As CSectionAnnotations, ByVal i As Long, _
        ByVal result As CSectionAnnotations, ByVal x As Double, ByVal y As Double)
    result.AddAnnotation source.AnnotationType(i), source.AnnotationID(i), source.StartX(i) + x, source.StartY(i) + y, _
        source.EndX(i) + x, source.EndY(i) + y, source.NormalX(i), source.NormalY(i), source.Text(i), _
        source.Value(i), source.Unit(i), source.Comment(i)
End Sub

' ДЛЯ ТЕСТОВ: форматирование численных исходных данных рецептов. Здесь нет
' вычисления позиций, пересечений или формы; всю математику выполняет рецепт.
' Circle: [D, as, n, d, d2, d3, loc2, loc3, bind2, bind3].
Public Function FixtureCircleDefinition(ByVal parameters As Variant) As CSectionDefinition
    Dim values As Variant, table As CSectionInputTable, recipe As ISectionRecipe, keys As Variant, i As Long
    values = ThisWorkbook.Names("rngCircleGeometry").RefersToRange.Value2
    Set table = New CSectionInputTable: table.Initialize "rngCircleGeometry", values
    keys = Array("Circle.Diameter", "Rebar.AxisDistance", "Rebar.Count", "Rebar.Diameter", _
        "Rebar.Diameter2", "Rebar.Diameter3", "Rebar.Loc2row", "Rebar.Loc3row", "Rebar.Bind2row", "Rebar.Bind3row")
    For i = 0 To 9
        If i <= UBound(parameters) Then
            values(table.RowByKey(CStr(keys(i))), 2) = parameters(i)
        ElseIf i < 6 Then
            values(table.RowByKey(CStr(keys(i))), 2) = 0#
        ElseIf i < 8 Then
            values(table.RowByKey(CStr(keys(i))), 2) = "Stacked"
        Else
            values(table.RowByKey(CStr(keys(i))), 2) = "EachBar"
        End If
    Next i
    Set recipe = New CSectionRecipeCircle
    Set FixtureCircleDefinition = FixtureRecipeDefinition(recipe, "rngCircleGeometry", values)
End Function

' RectSet: [B1, H1, B2, H2, offset, subtype]; faces = [H1, H2, B1, B2].
' Каждая пара: [as1,as2,d1,d2,n1,n2,Start1,End1,Start2,End2,
' row2d1,row2d2,row3d1,row3d2,loc2,loc3,bind2,bind3,loc2right,loc3right,bind2right,bind3right].
Public Function FixtureRectSetDefinition(ByVal dimensions As Variant, Optional ByVal faces As Variant) As CSectionDefinition
    Dim values As Variant, pair As Variant, rows As Variant, extras As Variant, index As Long, side As Long
    Dim row As Long, extra As Long, recipe As ISectionRecipe
    values = ThisWorkbook.Names("rngRectSetGeometry").RefersToRange.Value2
    values(3, 2) = "LSection": values(4, 2) = 0#
    If UBound(dimensions) >= 4 Then values(4, 2) = dimensions(4)
    If UBound(dimensions) >= 5 Then values(3, 2) = dimensions(5)
    values(8, 1) = dimensions(1): values(8, 2) = dimensions(0)
    values(8, 3) = dimensions(3): values(8, 4) = dimensions(2)
    rows = Array(11, 15, 13, 17): extras = Array(21, 25, 23, 27)
    For index = 0 To 3
        pair = Array(0#, 0#, 0#, 0#, 0, 0, 0#, 0#, 0#, 0#)
        If Not IsMissing(faces) Then pair = faces(index)
        For side = 0 To 1
            row = CLng(rows(index)) + side: extra = CLng(extras(index)) + side
            values(row, 2) = pair(side): values(row, 3) = pair(2 + side): values(row, 4) = pair(4 + side)
            values(row, 5) = pair(6 + side * 2): values(row, 6) = pair(7 + side * 2)
            values(extra, 2) = 0#: values(extra, 5) = 0#
            values(extra, 3) = "Stacked": values(extra, 6) = "Stacked"
            values(extra, 4) = "EachBar": values(extra, 7) = "EachBar"
            If UBound(pair) >= 13 Then
                values(extra, 2) = pair(10 + side): values(extra, 5) = pair(12 + side)
                values(extra, 3) = pair(14): values(extra, 6) = pair(15)
            End If
            If UBound(pair) >= 17 Then values(extra, 4) = pair(16): values(extra, 7) = pair(17)
            If side = 1 And UBound(pair) >= 21 Then
                values(extra, 3) = pair(18): values(extra, 6) = pair(19)
                values(extra, 4) = pair(20): values(extra, 7) = pair(21)
            End If
        Next side
    Next index
    Set recipe = New CSectionRecipeRectSet
    Set FixtureRectSetDefinition = FixtureRecipeDefinition(recipe, "rngRectSetGeometry", values)
End Function

' Rounded: [B,H,leftType,rightType,Wleft,Wright,R1left,R1right,R2left,R2right].
' faces = [H.Left,H.Right,B.Top,B.Bottom], каждый [as,d,n,d2,d3,loc2,loc3,bind2,bind3].
Public Function FixtureRoundedDefinition(ByVal dimensions As Variant, Optional ByVal faces As Variant) As CSectionDefinition
    Dim values As Variant, recipe As ISectionRecipe
    values = ThisWorkbook.Names("rngRoundedRectangleGeometry").RefersToRange.Value2
    values(3, 2) = dimensions(0): values(3, 3) = dimensions(1)
    values(7, 2) = dimensions(2): values(7, 3) = dimensions(3)
    values(8, 2) = dimensions(4): values(8, 3) = dimensions(5)
    values(9, 2) = dimensions(6): values(9, 3) = dimensions(7)
    values(10, 2) = dimensions(8): values(10, 3) = dimensions(9)
    If IsMissing(faces) Then
        FixtureFaceTable values, Array(Empty, Empty, Empty, Empty)
    Else
        FixtureFaceTable values, faces
    End If
    Set recipe = New CSectionRecipeRoundedRectangle
    Set FixtureRoundedDefinition = FixtureRecipeDefinition(recipe, "rngRoundedRectangleGeometry", values)
End Function

' Hollow: [B,H,R,Bo,Ho,Ro,xo,yo]; faces = наружные H/H/B/B, затем внутренние H/H/B/B.
Public Function FixtureHollowDefinition(ByVal dimensions As Variant, Optional ByVal faces As Variant) As CSectionDefinition
    Dim values As Variant, table As CSectionInputTable, recipe As ISectionRecipe, row As Long
    values = ThisWorkbook.Names("rngHollowRectangleGeometry").RefersToRange.Value2
    Set table = New CSectionInputTable: table.Initialize "rngHollowRectangleGeometry", values
    row = table.RowByKey("H", 1) + 1
    values(row, 1) = dimensions(1): values(row, 2) = dimensions(0): values(row, 3) = dimensions(2)
    values(row, 4) = dimensions(4): values(row, 5) = dimensions(3): values(row, 6) = dimensions(5)
    values(table.RowByKey("HollowRectangle.InnerOffsetX"), 2) = 0#
    values(table.RowByKey("HollowRectangle.InnerOffsetY"), 2) = 0#
    If UBound(dimensions) >= 6 Then values(table.RowByKey("HollowRectangle.InnerOffsetX"), 2) = dimensions(6)
    If UBound(dimensions) >= 7 Then values(table.RowByKey("HollowRectangle.InnerOffsetY"), 2) = dimensions(7)
    If IsMissing(faces) Then
        FixtureFaceTable values, Array(Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty)
    Else
        FixtureFaceTable values, faces
    End If
    Set recipe = New CSectionRecipeHollowRectangle
    Set FixtureHollowDefinition = FixtureRecipeDefinition(recipe, "rngHollowRectangleGeometry", values)
End Function

' Общий formatter массива граней: только перенос значений в строки ввода.
Private Sub FixtureFaceTable(ByRef values As Variant, ByVal faces As Variant)
    Dim table As CSectionInputTable, main As Long, extra As Long, i As Long, row As Long, pair As Variant
    Set table = New CSectionInputTable: table.Initialize "FixtureFaces", values
    main = table.RowByKey("as", 2): extra = table.RowByKey("d2", 2)
    For i = 0 To UBound(faces)
        row = main + i + 1: pair = faces(i)
        If IsEmpty(pair) Then pair = Array(40#, 0#, 0, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")
        values(row, 2) = pair(0): values(row, 3) = pair(1): values(row, 4) = pair(2)
        values(extra + i + 1, 2) = pair(3): values(extra + i + 1, 5) = pair(4)
        values(extra + i + 1, 3) = pair(5): values(extra + i + 1, 6) = pair(6)
        values(extra + i + 1, 4) = pair(7): values(extra + i + 1, 7) = pair(8)
    Next i
End Sub

' ДЛЯ ТЕСТОВ: неизменяемый табличный ввод настоящего рецепта без Excel-ссылок.
Public Function FixtureRecipeDefinition(ByVal recipe As ISectionRecipe, ByVal tableName As String, _
        ByVal values As Variant, Optional ByVal units As CUnitSystem = Nothing) As CSectionDefinition
    If units Is Nothing Then
        Set units = New CUnitSystem: units.InitializeDefaults
    End If
    Dim context As CSectionRecipeContext, table As CSectionInputTable
    Set table = New CSectionInputTable: table.Initialize tableName, values
    Set context = New CSectionRecipeContext: context.Initialize units
    context.AddTable table: context.Seal
    Set FixtureRecipeDefinition = recipe.CreateDefinition(context)
End Function

' ДЛЯ ТЕСТОВ: выбранный тип проходит тот же registered snapshot, что и программа.
Public Function FixtureWorkbookDefinition(ByVal typeID As String, ByVal workbook As Object, _
        ByVal units As CUnitSystem) As CSectionDefinition
    Dim registry As CSectionTypeRegistry, reader As CSectionInputSnapshotReader, context As CSectionRecipeContext
    Set registry = New CSectionTypeRegistry: Set reader = New CSectionInputSnapshotReader
    Set context = reader.Capture(workbook, registry.RecipeTableNames(typeID), units)
    Set FixtureWorkbookDefinition = registry.CreateDefinition(typeID, context)
End Function

' ДЛЯ ТЕСТОВ: исполняет только арматурные правила готового определения.
Public Function FixtureLayout(ByVal definition As CSectionDefinition) As CRebarLayout
    Dim executor As CRebarRuleExecutor: Set executor = New CRebarRuleExecutor
    Set FixtureLayout = executor.Build(definition, FixtureGeometry(definition.ConcreteDomain))
End Function

' ДЛЯ ТЕСТОВ: равномерный ряд по явно переданному замкнутому пути.
Public Function FixtureClosedCountRow(ByVal boundary As CGeometryPath, ByVal axis As CGeometryPath, _
        ByVal count As Long, ByVal diameter As Double, Optional ByVal steelClass As String = "A400", _
        Optional ByVal group As String = "FixtureRow", Optional ByVal comment As String = "") As CRebarLayout
    Dim facade As CSectionDefinitionBuilder, options As CRebarRuleOptions
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "ClosedRowFixture"
    facade.AddPath boundary: facade.AddPath axis: facade.Outer "OUTER", boundary.ID
    Set options = New CRebarRuleOptions: options.EndMode = ClosedUniform
    options.SourcePrefix = "B": options.Publication = PublishPerSourcePosition
    options.AnnotationGroup = group: options.Comment = comment
    facade.RebarByCount "ClosedRow", axis.ID, count, diameter, options, steelClass
    Set FixtureClosedCountRow = FixtureLayout(facade.Finish)
End Function

' ДЛЯ ТЕСТОВ: явные численные ряды [ID, x1, y1, x2, y2, n, d, StartSetback, EndSetback].
' Никакие размеры граней или параметры типа сечения здесь не интерпретируются.
Public Function FixtureStraightRows(ByVal rectangles As Variant, ByVal rows As Variant, _
        Optional ByVal steelClass As String = "A400") As CRebarLayout
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, row As Variant, options As CRebarRuleOptions
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "StraightRowFixture"
    Set path = facade.AdjacentRectangles("Boundary", rectangles): facade.Outer "OUTER", path.ID
    For Each row In rows
        Set path = facade.LinePath(CStr(row(0)), CDbl(row(1)), CDbl(row(2)), CDbl(row(3)), CDbl(row(4)))
        Set options = New CRebarRuleOptions
        options.StartSetback = CDbl(row(7)): options.EndSetback = CDbl(row(8))
        options.PositionOrder = FromPathMiddle: options.AnnotationGroup = path.ID
        options.Publication = PublishPerSourcePosition: options.Coincidence = KeepPublishedOwner
        facade.RebarByCount path.ID & ".rule", path.ID, CLng(row(5)), CDbl(row(6)), options, steelClass
    Next row
    Set FixtureStraightRows = FixtureLayout(facade.Finish)
End Function

Public Function FixtureDisk(ByVal diameter As Double, Optional ByVal centerX As Double = 0#, Optional ByVal centerY As Double = 0#) As CGeometryRegion
    Dim query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    Set FixtureDisk = FixtureGeometry(query.CircleRegion(centerX, centerY, diameter / 2#))
End Function

Public Function FixtureBox(ByVal width As Double, ByVal height As Double, _
        Optional ByVal centerX As Double = 0#, Optional ByVal centerY As Double = 0#) As CGeometryRegion
    Set FixtureBox = FixtureBoxes(Array(Array(centerX - width / 2#, centerY - height / 2#, width, height)))
End Function

Public Function FixtureBoxes(ByVal rectangles As Variant) As CGeometryRegion
    Dim query As CSectionGeometryQuery, path As CGeometryPath
    Set query = New CSectionGeometryQuery
    Set path = query.AdjacentRectangles("FixtureBoundary", rectangles)
    Set FixtureBoxes = FixtureGeometry(query.RegionFromPath(path))
End Function

Public Function FixtureFillets(ByVal polygon As CPolygonSpec) As CGeometryRegion
    Dim polygonSpec As CPolygonSpec, polygonPoint As Variant
    Dim query As CSectionGeometryQuery, path As CGeometryPath
    Set query = New CSectionGeometryQuery
    Set path = query.RoundedPolygon("FixtureBoundary", polygon)
    Set FixtureFillets = FixtureGeometry(query.RegionFromPath(path))
End Function
