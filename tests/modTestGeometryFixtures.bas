Attribute VB_Name = "modTestGeometryFixtures"
Option Explicit

' ДЛЯ ТЕСТОВ: создаёт контрольные области общими геометрическими классами в мм.
' Новые рабочие типы не регистрирует. Входные данные книги читает обычным
' CSectionInputSnapshotReader только из зарегистрированных таблиц.
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
    Dim values As Variant, reader As CSectionInputReader, recipe As ISectionRecipe, keys As Variant, blocks As Variant
    Dim i As Long, value As Variant, row As String
    values = FixtureRecipeValues("rngCircleGeometry")
    Set reader = FixtureInput(values)
    keys = Array("val", "as", "n", "d", "d2", "d3", "pos2", "pos3", "bind2", "bind3", "a2", "a3")
    blocks = Array("General", "Main", "Main", "Main", "Extra", "Extra", "Extra", "Extra", "Extra", "Extra", "Extra", "Extra")
    For i = 0 To 11
        If i <= UBound(parameters) Then
            value = parameters(i)
        ElseIf i = 6 Or i = 7 Then
            value = "Stacked"
        ElseIf i = 8 Or i = 9 Then
            value = "EachBar"
        Else
            value = 0#
        End If
        If i = 0 Then row = "D" Else row = "Circle"
        values(reader.RowIndex(CStr(blocks(i)), row), reader.ColumnIndex(CStr(blocks(i)), CStr(keys(i)))) = value
    Next i
    Set recipe = New CSectionRecipeCircle
    Set FixtureCircleDefinition = FixtureRecipeDefinition(recipe, "rngCircleGeometry", values)
End Function

' Rectangles: [B1, H1, B2, H2, offset, subtype]; faces = [H1, H2, B1, B2].
' Каждая пара: [as1,as2,d1,d2,n1,n2,Start1,End1,Start2,End2,
' row2d1,row2d2,row3d1,row3d2,loc2,loc3,bind2,bind3,loc2right,loc3right,bind2right,bind3right].
Public Function FixtureRectanglesDefinition(ByVal dimensions As Variant, Optional ByVal faces As Variant) As CSectionDefinition
    Dim values As Variant, pair As Variant, keys As Variant, fields As Variant, index As Long, side As Long
    Dim key As String, recipe As ISectionRecipe, reader As CSectionInputReader, row As Long, extra As Long
    values = FixtureRecipeValues("rngRectanglesGeometry")
    Set reader = FixtureInput(values)
    FixtureSetInput values, "General", "Shape", "val", "TwoRect"
    FixtureSetInput values, "Detail", "XOffset", "val", 0#
    If UBound(dimensions) >= 4 Then FixtureSetInput values, "Detail", "XOffset", "val", dimensions(4)
    If UBound(dimensions) >= 5 Then FixtureSetInput values, "General", "Shape", "val", dimensions(5)
    fields = Array("B1", "H1", "B2", "H2")
    For index = 0 To 3: FixtureSetInput values, "General", CStr(fields(index)), "val", dimensions(index): Next index
    keys = Array("R1.Left", "R2.Left", "R1.Top", "R2.Top")
    For index = 0 To 3
        pair = Array(0#, 0#, 0#, 0#, 0, 0, 0#, 0#, 0#, 0#)
        If Not IsMissing(faces) Then pair = faces(index)
        For side = 0 To 1
            key = CStr(keys(index))
            If side = 1 Then key = Replace$(Replace$(key, "Left", "Right"), "Top", "Bottom")
            row = reader.RowIndex("Main", key): extra = reader.RowIndex("Extra", key)
            values(row, reader.ColumnIndex("Main", "as")) = pair(side)
            values(row, reader.ColumnIndex("Main", "d")) = pair(2 + side)
            values(row, reader.ColumnIndex("Main", "n")) = pair(4 + side)
            values(row, reader.ColumnIndex("Main", "s0")) = pair(6 + side * 2)
            values(row, reader.ColumnIndex("Main", "s1")) = pair(7 + side * 2)
            For Each fields In Array("d2", "d3", "a2", "a3"): values(extra, reader.ColumnIndex("Extra", CStr(fields))) = 0#: Next fields
            For Each fields In Array("pos2", "pos3"): values(extra, reader.ColumnIndex("Extra", CStr(fields))) = "Stacked": Next fields
            For Each fields In Array("bind2", "bind3"): values(extra, reader.ColumnIndex("Extra", CStr(fields))) = "EachBar": Next fields
            If UBound(pair) >= 13 Then
                values(extra, reader.ColumnIndex("Extra", "d2")) = pair(10 + side): values(extra, reader.ColumnIndex("Extra", "d3")) = pair(12 + side)
                values(extra, reader.ColumnIndex("Extra", "pos2")) = pair(14): values(extra, reader.ColumnIndex("Extra", "pos3")) = pair(15)
            End If
            If UBound(pair) >= 17 Then values(extra, reader.ColumnIndex("Extra", "bind2")) = pair(16): values(extra, reader.ColumnIndex("Extra", "bind3")) = pair(17)
            If side = 1 And UBound(pair) >= 21 Then
                values(extra, reader.ColumnIndex("Extra", "pos2")) = pair(18): values(extra, reader.ColumnIndex("Extra", "pos3")) = pair(19)
                values(extra, reader.ColumnIndex("Extra", "bind2")) = pair(20): values(extra, reader.ColumnIndex("Extra", "bind3")) = pair(21)
            End If
        Next side
    Next index
    Set recipe = New CSectionRecipeRectangles
    Set FixtureRectanglesDefinition = FixtureRecipeDefinition(recipe, "rngRectanglesGeometry", values)
End Function

' Rounded: [B,H,Wleft,Wright,R1left,R1right,R2left,R2right].
' faces = [H.Left,H.Right,B.Top,B.Bottom], каждый [as,d,n,d2,d3,loc2,loc3,bind2,bind3].
Public Function FixtureRoundedDefinition(ByVal dimensions As Variant, Optional ByVal faces As Variant, Optional ByVal readSetbacks As Boolean = False) As CSectionDefinition
    Dim values As Variant, recipe As ISectionRecipe, index As Long, fields As Variant, side As Long
    values = FixtureRecipeValues("rngPointedSectionGeometry", readSetbacks)
    FixtureSetInput values, "General", "B", "val", dimensions(0): FixtureSetInput values, "General", "H", "val", dimensions(1)
    fields = Array("W", "R1", "R2")
    For index = 0 To 2
        FixtureSetInput values, "Detail", CStr(fields(index)), "left", dimensions(2 + index * 2)
        FixtureSetInput values, "Detail", CStr(fields(index)), "right", dimensions(3 + index * 2)
    Next index
    If IsMissing(faces) Then
        FixtureFaceTable values, Array(Empty, Empty, Empty, Empty)
    Else
        FixtureFaceTable values, faces
    End If
    Set recipe = New CSectionRecipePointedSection
    Set FixtureRoundedDefinition = FixtureRecipeDefinition(recipe, "rngPointedSectionGeometry", values)
End Function

' Hollow: [B,H,R,Bo,Ho,Ro,xo,yo]; faces = наружные H/H/B/B, затем внутренние H/H/B/B.
Public Function FixtureHollowDefinition(ByVal dimensions As Variant, Optional ByVal faces As Variant, Optional ByVal readSetbacks As Boolean = False) As CSectionDefinition
    Dim values As Variant, recipe As ISectionRecipe, index As Long, keys As Variant
    values = FixtureRecipeValues("rngHollowRectangleGeometry", readSetbacks)
    keys = Array("B", "H", "R", "Bo", "Ho", "Ro")
    For index = 0 To 5
        If index < 3 Then
            FixtureSetInput values, "General", CStr(keys(index)), "val", dimensions(index)
        Else
            FixtureSetInput values, "Detail", CStr(keys(index)), "val", dimensions(index)
        End If
    Next index
    FixtureSetInput values, "Detail", "Xo", "val", 0#: FixtureSetInput values, "Detail", "Yo", "val", 0#
    If UBound(dimensions) >= 6 Then FixtureSetInput values, "Detail", "Xo", "val", dimensions(6)
    If UBound(dimensions) >= 7 Then FixtureSetInput values, "Detail", "Yo", "val", dimensions(7)
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
    Dim reader As CSectionInputReader, keys As Variant, fields As Variant, i As Long, j As Long, key As String, pair As Variant
    Set reader = FixtureInput(values)
    keys = Array("Left", "Right", "Top", "Bottom", "OpeningLeft", "OpeningRight", "OpeningTop", "OpeningBottom")
    fields = Array("as", "d", "n", "d2", "d3", "pos2", "pos3", "bind2", "bind3")
    For i = 0 To UBound(faces)
        key = CStr(keys(i)): pair = faces(i)
        If IsEmpty(pair) Then pair = Array(40#, 0#, 0, 0#, 0#, "Stacked", "Stacked", "EachBar", "EachBar")
        For j = 0 To 8
            If j < 3 Then
                FixtureSetInput values, "Main", key, CStr(fields(j)), pair(j)
            Else
                FixtureSetInput values, "Extra", key, CStr(fields(j)), pair(j)
            End If
        Next j
        FixtureSetInput values, "Extra", key, "a2", 0#: FixtureSetInput values, "Extra", key, "a3", 0#
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

' ДЛЯ ТЕСТОВ: читает входные данные выбранного рецепта тем же способом, что рабочая программа.
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


' ДЛЯ ТЕСТОВ: смысловое имя поля Circle переводится в единственный действующий источник новой таблицы.
Public Function FixtureCircleFieldCell(ByVal workbook As Object, ByVal key As String, Optional ByVal sourceTable As Object = Nothing) As Object
    Dim area As Object, column As String, block As String, row As String
    If sourceTable Is Nothing Then Set area = workbook.Names("rngCircleGeometry").RefersToRange Else Set area = sourceTable
    block = "Main": row = "Circle"
    Select Case key
        Case "Circle.Diameter": block = "General": row = "D": column = "val"
        Case "Rebar.AxisDistance": column = "as"
        Case "Rebar.Diameter": column = "d"
        Case "Rebar.Count": column = "n"
        Case "Rebar.StartSetback": column = "s0"
        Case "Rebar.EndSetback": column = "s1"
        Case "Rebar.Diameter2": block = "Extra": column = "d2"
        Case "Rebar.AxisOffset2": block = "Extra": column = "a2"
        Case "Rebar.Bind2row": block = "Extra": column = "bind2"
        Case "Rebar.Loc2row": block = "Extra": column = "pos2"
        Case "Rebar.Diameter3": block = "Extra": column = "d3"
        Case "Rebar.AxisOffset3": block = "Extra": column = "a3"
        Case "Rebar.Bind3row": block = "Extra": column = "bind3"
        Case "Rebar.Loc3row": block = "Extra": column = "pos3"
        Case Else: Err.Raise vbObjectError + 5590, , "Неизвестное контрольное поле круга: " & key
    End Select
    Set FixtureCircleFieldCell = FixtureInputCell(area, block, row, column)
End Function


' ДЛЯ ТЕСТОВ: индексирует массив по реальным ключам без ссылок на Excel.
Public Function FixtureInput(ByVal values As Variant) As CSectionInputReader
    Dim table As CSectionInputTable: Set table = New CSectionInputTable
    table.Initialize "Fixture", values
    Set FixtureInput = New CSectionInputReader: FixtureInput.Initialize table
End Function

' ДЛЯ ТЕСТОВ: меняет только указанное поле снимка, не вычисляя геометрию.
Public Sub FixtureSetInput(ByRef values As Variant, ByVal block As String, ByVal row As String, ByVal column As String, ByVal value As Variant)
    FixtureNormalizeField block, row, column
    Dim reader As CSectionInputReader: Set reader = FixtureInput(values)
    values(reader.RowIndex(block, row), reader.ColumnIndex(block, column)) = value
End Sub

' ДЛЯ ТЕСТОВ: возвращает настоящую ячейку поля после изменения размещения таблицы.
Public Function FixtureInputCell(ByVal area As Object, ByVal block As String, ByVal row As String, Optional ByVal column As String = "val") As Object
    FixtureNormalizeField block, row, column
    Dim reader As CSectionInputReader
    Set reader = FixtureInput(area.Value2)
    Set FixtureInputCell = area.Cells(reader.RowIndex(block, row), reader.ColumnIndex(block, column))
End Function

' ДЛЯ ТЕСТОВ: выбирает тип до обращения к строкам общей таблицы с формульными ключами.
Public Sub FixtureSelectRecipeType(ByVal workbook As Object, ByVal typeID As String)
    Dim settings As Object, row As Long
    Set settings = workbook.Names("rngSystemSettings").RefersToRange
    For row = 2 To settings.Rows.Count
        If CStr(settings.Cells(row, 1).Value2) = "Geometry.Type" Then settings.Cells(row, 2).Value2 = typeID: Exit Sub
    Next row
    Err.Raise vbObjectError + 5590, "modTestGeometryFixtures", "Не найдена настройка Geometry.Type."
End Sub

' ДЛЯ ТЕСТОВ: смысловое имя контрольного параметра находит в действующих блоках.
' Размеры читаются из General, армирование — из общей таблицы выбранного типа.
Public Function FixtureConfigFieldCell(ByVal key As String) As Object
    Dim parts As Variant, typeID As String, field As String, side As String, block As String
    Dim area As Object, rowKey As String, columnKey As String, index As Long, suffix As String
    If Left$(key, 6) = "Rebar." Or Left$(key, 7) = "Circle." Then
        Set FixtureConfigFieldCell = FixtureCircleFieldCell(ThisWorkbook, key): Exit Function
    End If
    parts = Split(key, "."): typeID = CStr(parts(0))
    If typeID <> "Rectangles" And typeID <> "PointedSection" And typeID <> "HollowRectangle" Then Exit Function
    Set area = ThisWorkbook.Names("rng" & typeID & "Geometry").RefersToRange
    field = Mid$(key, Len(typeID) + 2): block = "General": columnKey = "val"
    Select Case typeID
        Case "Rectangles"
            Select Case field
                Case "SectionType": rowKey = "Shape"
                Case "UpperOffsetX": rowKey = "XOffset"
                Case "H1", "B1", "H2", "B2": rowKey = field
                Case Else
                    side = CStr(parts(1)): field = CStr(parts(2))
                    index = CLng(Right$(field, 1)): field = Left$(field, Len(field) - 1)
                    rowKey = "R" & Right$(side, 1) & "."
                    If Left$(side, 1) = "H" Then
                        rowKey = rowKey & IIf(index = 1, "Left", "Right")
                    Else
                        rowKey = rowKey & IIf(index = 1, "Top", "Bottom")
                    End If
                    Select Case field
                        Case "as_": columnKey = "as"
                        Case "d_": columnKey = "d"
                        Case "n_": columnKey = "n"
                        Case "StartOffset": columnKey = "s0"
                        Case "EndOffset": columnKey = "s1"
                        Case "d_2row_": columnKey = "d2"
                        Case "d_3row_": columnKey = "d3"
                        Case "loc_2row_": columnKey = "pos2"
                        Case "loc_3row_": columnKey = "pos3"
                        Case "bind_2row_": columnKey = "bind2"
                        Case "bind_3row_": columnKey = "bind3"
                        Case Else: Err.Raise 5, , "Неизвестное контрольное поле: " & key
                    End Select
                    block = "Main": If InStr(columnKey, "2") Or InStr(columnKey, "3") Then block = "Extra"
            End Select
        Case Else
            Select Case field
                Case "InnerOffsetX": rowKey = "Xo"
                Case "InnerOffsetY": rowKey = "Yo"
                Case "OpeningB": rowKey = "Bo"
                Case "OpeningH": rowKey = "Ho"
                Case "OpeningR": rowKey = "Ro"
                Case "B", "H", "R": rowKey = field
                Case "Left.W", "Left.R1", "Left.R2", "Right.W", "Right.R1", "Right.R2": rowKey = field
                Case Else
                    index = UBound(parts): field = CStr(parts(index)): side = CStr(parts(index - 1))
                    rowKey = side: If CStr(parts(1)) = "Opening" Then rowKey = "Opening" & side
                    block = "Main": columnKey = field
                    Select Case field
                        Case "d_2row": block = "Extra": columnKey = "d2"
                        Case "d_3row": block = "Extra": columnKey = "d3"
                        Case "loc_2row": block = "Extra": columnKey = "pos2"
                        Case "loc_3row": block = "Extra": columnKey = "pos3"
                        Case "bind_2row": block = "Extra": columnKey = "bind2"
                        Case "bind_3row": block = "Extra": columnKey = "bind3"
                    End Select
            End Select
    End Select
    Set FixtureConfigFieldCell = FixtureInputCell(area, block, rowKey, columnKey)
End Function

' Приводит обозначения параметров в контрольных наборах к двум блокам новой таблицы.
Private Sub FixtureNormalizeField(ByRef block As String, ByRef row As String, ByRef column As String)
    If block = "Detail" Then
        block = "General"
        If column = "left" Or column = "right" Then row = UCase$(Left$(column, 1)) & Mid$(column, 2) & "." & row: column = "val"
    End If
End Sub

' ДЛЯ ТЕСТОВ: объединяет два источника в независимый массив для подготовки контрольного ввода.
' Позиции стержней и контуры здесь не вычисляются; размеры и ключи берутся из действующей книги.
' Обычно s0/s1 равны нулю; readSetbacks включает их чтение из общей таблицы для проверки отступов.
Public Function FixtureRecipeValues(ByVal tableName As String, Optional ByVal readSetbacks As Boolean = False) As Variant
    Dim values As Variant, reader As CSectionInputReader, key As Variant
    values = ThisWorkbook.Names(tableName).RefersToRange.Value2
    If Not readSetbacks Then
        Set reader = FixtureInput(values)
        For Each key In reader.RowKeys("Main")
            If CStr(key) <> "Units" Then
                values(reader.RowIndex("Main", CStr(key)), reader.ColumnIndex("Main", "s0")) = 0#
                values(reader.RowIndex("Main", CStr(key)), reader.ColumnIndex("Main", "s1")) = 0#
            End If
        Next key
    End If
    FixtureRecipeValues = values
End Function

' ДЛЯ ТЕСТОВ: задаёт исходное армирование конкретного типа, не меняя выбранный тип.
' Выбор типа в рабочей книге ничего не сбрасывает; только контрольный набор готовит свои значения.
Public Sub FixtureResetReinforcement(ByVal typeID As String)
    Dim rows As Variant, item As Variant, fields As Variant, index As Long, table As Object, value As Variant
    Dim settings As CSystemSettingsReader, units As CUnitSystem, factor As Double
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings: factor = units.InputLengthToInternal(1#)
    Select Case typeID
        Case "Circle"
            rows = Array( _
                Array("Circle", 40#, 20#, 8#, 0#, 0#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"))
        Case "Rectangles"
            rows = Array( _
                Array("R1.Top", 40#, 32#, 2#, 80#, 80#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"), _
                Array("R1.Bottom", 40#, 32#, 2#, 80#, 80#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"), _
                Array("R1.Left", 40#, 32#, 5#, 80#, 80#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"), _
                Array("R1.Right", 40#, 32#, 5#, 80#, 80#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"), _
                Array("R2.Top", 40#, 32#, 5#, 80#, 80#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"), _
                Array("R2.Bottom", 40#, 32#, 5#, 80#, 80#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"), _
                Array("R2.Left", 40#, 32#, 2#, 80#, 80#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"), _
                Array("R2.Right", 40#, 32#, 2#, 80#, 80#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"))
        Case "PointedSection"
            rows = Array( _
                Array("Top", 40#, 32#, 8#, 0#, 0#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"), _
                Array("Bottom", 40#, 32#, 8#, 0#, 0#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"), _
                Array("Left", 40#, 32#, 5#, 0#, 0#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"), _
                Array("Right", 40#, 32#, 5#, 0#, 0#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"))
        Case "HollowRectangle"
            rows = Array( _
                Array("Top", 40#, 20#, 8#, 0#, 0#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"), _
                Array("Bottom", 40#, 20#, 8#, 0#, 0#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"), _
                Array("Left", 40#, 20#, 6#, 0#, 0#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"), _
                Array("Right", 40#, 20#, 6#, 0#, 0#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"), _
                Array("OpeningTop", 40#, 20#, "—", 0#, 0#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"), _
                Array("OpeningBottom", 40#, 20#, "—", 0#, 0#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"), _
                Array("OpeningLeft", 40#, 20#, "—", 0#, 0#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"), _
                Array("OpeningRight", 40#, 20#, "—", 0#, 0#, Empty, 0#, "EachBar", "Stacked", Empty, 0#, "EachBar", "Stacked"))
    End Select
    Set table = ThisWorkbook.Names("rng" & typeID & "Geometry").RefersToRange
    fields = Array("as", "d", "n", "s0", "s1", "d2", "a2", "bind2", "pos2", "d3", "a3", "bind3", "pos3")
    For Each item In rows
        For index = 0 To UBound(fields)
            value = item(index + 1)
            If IsNumeric(value) And CStr(fields(index)) <> "n" Then value = CDbl(value) / factor
            If CStr(fields(index)) <> "n" Or CStr(value) <> "—" Then
                If index < 5 Then
                    FixtureInputCell(table, "Main", CStr(item(0)), CStr(fields(index))).Value2 = value
                Else
                    FixtureInputCell(table, "Extra", CStr(item(0)), CStr(fields(index))).Value2 = value
                End If
            End If
        Next index
    Next item
End Sub

' ДЛЯ ТЕСТОВ: сохраняет изменяемый диапазон Config и независимые таблицы сечений.
' Снимок нужен после переноса геометрии на отдельный лист.
Public Function FixtureSaveInputs(ByVal config As Object) As Variant
    Dim sections As Object
    Set sections = config.Worksheet.Parent.Worksheets("Сечения").UsedRange
    FixtureSaveInputs = Array(config.Formula, sections.Address, sections.Formula)
End Function

' ДЛЯ ТЕСТОВ: возвращает исходные формулы и значения обоих листов после сценария.
Public Sub FixtureRestoreInputs(ByVal config As Object, ByVal saved As Variant)
    config.Formula = saved(0)
    config.Worksheet.Parent.Worksheets("Сечения").Range(CStr(saved(1))).Formula = saved(2)
End Sub

' ДЛЯ ТЕСТОВ: возвращает выбранный тип, не меняя исходные данные сечений.
Public Function FixtureSelectedType() As String
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    FixtureSelectedType = settings.GetRequiredString("Geometry.Type")
End Function

' ДЛЯ ТЕСТОВ: таблица текущего типа содержит и размеры, и армирование.
Public Function FixtureSelectedInput() As Object
    Set FixtureSelectedInput = ThisWorkbook.Names("rng" & FixtureSelectedType() & "Geometry").RefersToRange
End Function

' ДЛЯ ТЕСТОВ
' Повторяет две пользовательские команды для сценариев, проверяющих новый
' ввод генератора. Сохранённый импорт AutoCAD повторно не получает.
Public Function FixtureRunSectionWorkflow(ByVal workbook As Object, Optional ByVal showMessages As Boolean = False) As String
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook workbook
    If StrComp(settings.GetRawString("Geometry.Source", ""), "AutoCAD", vbTextCompare) <> 0 Then
        Dim formation As String: formation = FormSectionForWorkbook(workbook)
    End If
    FixtureRunSectionWorkflow = RunSectionCalculationForWorkbook(workbook, showMessages)
End Function
