VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionRecipeRectSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ISectionRecipe

' Один рецепт Rectangle/LSection/TwoRectangles: свой табличный ввод,
' смежные прямоугольники, восемь смысловых граней и зависимые ряды.
Private Type TRectSetDimensions
    Shape As String
    Header As Long
    B1 As Double
    H1 As Double
    B2 As Double
    H2 As Double
    Offset As Double
End Type

Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
    Dim table As CSectionInputTable, units As CUnitSystem, facade As CSectionDefinitionBuilder
    Dim sizes As TRectSetDimensions, upper As CGeometryPath, lower As CGeometryPath, outer As CGeometryPath
    Set table = context.Table("rngRectSetGeometry")
    Set units = context.Units
    ReadDimensions table, units, sizes
    Set facade = New CSectionDefinitionBuilder
    facade.Initialize "RectSet"
    AddGeometry facade, sizes, upper, lower, outer
    AddRectangleRows facade, table, units, sizes.Header, 0, sizes.Offset, sizes.H2, sizes.B1, sizes.H1
    If sizes.Shape <> "Rectangle" Then AddRectangleRows facade, table, units, sizes.Header, 4, 0#, 0#, sizes.B2, sizes.H2
    AddDimensions facade, sizes, upper, lower, outer
    Set ISectionRecipe_CreateDefinition = facade.Finish
End Function

' Каждый вариант читает только свои активные размеры; отсутствующая нижняя
' часть Rectangle не требует заполнения полей остальных двух вариантов.
Private Sub ReadDimensions(ByVal table As CSectionInputTable, ByVal units As CUnitSystem, ByRef sizes As TRectSetDimensions)
    sizes.Shape = table.Choice(table.RowByKey("RectSet.SectionType"), 2, _
        Array("Rectangle", "LSection", "TwoRectangles", "RectSet", "TwoRectangle", "TwoRectanglesOffset", "OffsetTwoRectangles"))
    If sizes.Shape = "RectSet" Then sizes.Shape = "LSection"
    If sizes.Shape <> "Rectangle" And sizes.Shape <> "LSection" Then sizes.Shape = "TwoRectangles"
    sizes.Header = table.RowByKey("H1")
    sizes.H1 = table.Length(sizes.Header + 1, 1, units)
    sizes.B1 = table.Length(sizes.Header + 1, 2, units)
    Require table, sizes.Header + 1, 1, sizes.H1 > 0#, "Высота верхнего прямоугольника должна быть положительной."
    Require table, sizes.Header + 1, 2, sizes.B1 > 0#, "Ширина верхнего прямоугольника должна быть положительной."
    If sizes.Shape <> "Rectangle" Then
        sizes.H2 = table.Length(sizes.Header + 1, 3, units)
        sizes.B2 = table.Length(sizes.Header + 1, 4, units)
        Require table, sizes.Header + 1, 3, sizes.H2 > 0#, "Высота нижнего прямоугольника должна быть положительной."
        Require table, sizes.Header + 1, 4, sizes.B2 > 0#, "Ширина нижнего прямоугольника должна быть положительной."
        If sizes.Shape = "LSection" Then Require table, sizes.Header + 1, 2, sizes.B1 < sizes.B2, _
            "Для LSection ширина верхнего прямоугольника должна быть меньше нижнего."
    End If
    If sizes.Shape = "TwoRectangles" Then
        Dim row As Long
        row = table.RowByKey("RectSet.UpperOffsetX")
        sizes.Offset = table.Length(row, 2, units)
        Require table, row, 2, sizes.Offset < sizes.B2 - GEOM_TOLERANCE And sizes.Offset + sizes.B1 > GEOM_TOLERANCE, _
            "Верхний и нижний прямоугольники должны иметь общий участок по X."
    End If
End Sub

' LSection и TwoRectangles имеют одно построение; различаются размеры и
' допустимое смещение, а не общий алгоритм объединения прямоугольников.
Private Sub AddGeometry(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRectSetDimensions, _
        ByRef upper As CGeometryPath, ByRef lower As CGeometryPath, ByRef outer As CGeometryPath)
    Set upper = facade.RectanglePath("Upper", sizes.Offset, sizes.H2, sizes.B1, sizes.H1)
    If sizes.Shape = "Rectangle" Then
        Set outer = upper
    Else
        Set lower = facade.RectanglePath("Lower", 0#, 0#, sizes.B2, sizes.H2)
        Set outer = facade.AdjacentRectangles("Boundary", Array(Array(0#, 0#, sizes.B2, sizes.H2), _
            Array(sizes.Offset, sizes.H2, sizes.B1, sizes.H1)))
    End If
    Dim contourIDs() As String, i As Long
    ReDim contourIDs(0 To outer.SegmentCount - 1)
    For i = 1 To outer.SegmentCount
        contourIDs(i - 1) = "CONTOUR_LINE_" & CStr(i)
    Next i
    facade.Outer "OUTER", outer.ID, "Сегмент контура сечения.", contourIDs
End Sub

Private Sub AddDimensions(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRectSetDimensions, _
        ByVal upper As CGeometryPath, ByVal lower As CGeometryPath, ByVal outer As CGeometryPath)
    facade.HorizontalExtent "DIM_B", outer.ID, "B", 0#, -1#, "Общая ширина RectSet-сечения."
    facade.VerticalExtent "DIM_H", outer.ID, "H", GeomMin(0#, sizes.Offset), -1#, "Общая высота RectSet-сечения."
    facade.HorizontalExtent "DIM_TOP_WIDTH", upper.ID, "B1", sizes.H1 + sizes.H2, 1#, "Ширина верхнего прямоугольника RectSet-сечения."
    If sizes.Shape <> "Rectangle" Then
        facade.VerticalExtent "DIM_RIGHT_HEIGHT", lower.ID, "H2", sizes.B2, 1#, "Высота нижнего прямоугольника RectSet-сечения."
        If Abs(sizes.Offset) > GEOM_TOLERANCE Then facade.HorizontalDistance "DIM_UPPER_OFFSET_X", ChrW$(&H394), 0#, sizes.Offset, sizes.H2, 1#, _
            "Горизонтальное смещение верхнего прямоугольника относительно базового положения."
    End If
End Sub

' База отступов - полная длина грани, не линия, укороченная соседним as.
' Направления обхода и касательного SideBySide заданы явно в рецепте.
Private Sub AddRectangleRows(ByVal facade As CSectionDefinitionBuilder, ByVal table As CSectionInputTable, _
        ByVal units As CUnitSystem, ByVal header As Long, ByVal base As Long, _
        ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal height As Double)
    Dim side As Long, row As Long, extra As Long, axisDistance As Double, diameter As Double, count As Long
    Dim path As CGeometryPath, groupID As String, label As String, options As CRebarRuleOptions
    Dim nx As Double, ny As Double, sideSign As Double, d2 As Double, d3 As Double, loc2 As String, loc3 As String
    Dim prefix As String
    If base = 0 Then prefix = "1" Else prefix = "2"
    For side = 1 To 4
        row = header + 3 + base + side
        extra = header + 13 + base + side
        count = table.Count(row, 4)
        diameter = table.OptionalDiameter(row, 3, units)
        If count = 0 Or diameter = 0# Then GoTo NextSide
        axisDistance = table.Length(row, 2, units)
        Require table, row, 2, axisDistance > diameter / 2#, "Расстояние до оси должно быть больше половины диаметра стержня."
        Set options = New CRebarRuleOptions
        options.StartSetback = table.Length(row, 5, units)
        options.EndSetback = table.Length(row, 6, units)
        options.PositionOrder = FromPathMiddle
        options.Publication = PublishPerSourcePosition
        options.Coincidence = KeepPublishedOwner
        options.SourceReference = table.SourceReference(row, 2)
        label = table.Text(row, 1) & " грань"
        Select Case side
            Case 1
                groupID = "H" & prefix & ".as_1": nx = -1#: ny = 0#: sideSign = -1#
                Set path = facade.LinePath(groupID, x + axisDistance, y, x + axisDistance, y + height)
            Case 2
                groupID = "H" & prefix & ".as_2": nx = 1#: ny = 0#: sideSign = 1#
                Set path = facade.LinePath(groupID, x + width - axisDistance, y + height, x + width - axisDistance, y)
            Case 3
                groupID = "B" & prefix & ".as_1": nx = 0#: ny = 1#: sideSign = 1#
                Set path = facade.LinePath(groupID, x, y + height - axisDistance, x + width, y + height - axisDistance)
            Case 4
                groupID = "B" & prefix & ".as_2": nx = 0#: ny = -1#: sideSign = -1#
                Set path = facade.LinePath(groupID, x + width, y + axisDistance, x, y + axisDistance)
        End Select
        options.AnnotationGroup = groupID
        options.Comment = "RectSet: " & label & ", ряд 1"
        facade.RebarByCount groupID & ".Main", path.ID, count, diameter, options
        d2 = table.OptionalDiameter(extra, 2, units)
        d3 = table.OptionalDiameter(extra, 5, units)
        If d2 > 0# Then
            loc2 = table.Choice(extra, 3, Array("Stacked", "SideBySide"))
            Set options = AdditionalOptions(table, extra, 2, groupID, label, 2, loc2, diameter / 2# + d2 / 2#, sideSign)
            facade.RebarDependent groupID & ".Row2", groupID & ".Main", d2, options
        End If
        If d3 > 0# Then
            loc3 = table.Choice(extra, 6, Array("Stacked", "SideBySide"))
            Dim distance As Double
            distance = diameter / 2# + d3 / 2#
            If d2 > 0# And loc2 = loc3 Then distance = distance + d2
            Set options = AdditionalOptions(table, extra, 5, groupID, label, 3, loc3, distance, sideSign)
            facade.RebarDependent groupID & ".Row3", groupID & ".Main", d3, options
        End If
        facade.RebarRowLabel groupID, groupID & ".Main", nx, ny, axisDistance
NextSide:
    Next side
End Sub

Private Function AdditionalOptions(ByVal table As CSectionInputTable, ByVal row As Long, ByVal column As Long, _
        ByVal groupID As String, ByVal label As String, ByVal rowNumber As Long, ByVal location As String, _
        ByVal distance As Double, ByVal sideSign As Double) As CRebarRuleOptions
    Dim options As CRebarRuleOptions
    Set options = New CRebarRuleOptions
    If location = "Stacked" Then options.NormalOffset = -distance Else options.TangentOffset = sideSign * distance
    options.EverySecond = (table.Choice(row, column + 2, Array("EachBar", "EverySecondBar")) = "EverySecondBar")
    options.Publication = PublishPerSourcePosition
    options.Coincidence = KeepPublishedOwner
    options.AnnotationGroup = groupID
    options.Comment = "RectSet: " & label & ", ряд " & CStr(rowNumber)
    options.SourceReference = table.SourceReference(row, column)
    Set AdditionalOptions = options
End Function

Private Sub Require(ByVal table As CSectionInputTable, ByVal row As Long, ByVal column As Long, ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then SectionRecipeError "CSectionRecipeRectSet", reason, table.SourceReference(row, column)
End Sub
