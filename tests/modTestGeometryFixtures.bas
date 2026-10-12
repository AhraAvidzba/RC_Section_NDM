Attribute VB_Name = "modTestGeometryFixtures"
Option Explicit

' ДЛЯ ТЕСТОВ: численные области во внутренних мм через актуальную общую геометрию.
' Не участвует в добавлении производственных типов сечений; workbook snapshot
' ниже читает только зарегистрированные таблицы через штатный reader.
Public Function FixtureGeometry(ByVal region As CConcreteRegion) As CGeometryRegion
    Set FixtureGeometry = New CGeometryRegion
    FixtureGeometry.Initialize region
End Function

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

Public Function FixtureFillets(ByVal vertices As Variant, ByVal radii As Variant) As CGeometryRegion
    Dim query As CSectionGeometryQuery, path As CGeometryPath, ids() As String, i As Long
    ReDim ids(LBound(vertices) To UBound(vertices))
    For i = LBound(vertices) To UBound(vertices): ids(i) = CStr(i): Next i
    Set query = New CSectionGeometryQuery
    Set path = query.RoundedPolygon("FixtureBoundary", vertices, radii, ids)
    Set FixtureFillets = FixtureGeometry(query.RegionFromPath(path))
End Function
