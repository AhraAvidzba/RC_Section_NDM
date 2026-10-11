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
Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
    Dim table As CSectionInputTable, units As CUnitSystem, facade As CSectionDefinitionBuilder
    Dim shape As String, header As Long, b1 As Double, h1 As Double, b2 As Double, h2 As Double, offset As Double
    Dim upper As CGeometryPath, lower As CGeometryPath, outer As CGeometryPath, contourIDs() As String, i As Long
    Set table = context.Table("rngRectSetGeometry")
    Set units = context.Units
    shape = table.Choice(table.RowByKey("RectSet.SectionType"), 2, _
        Array("Rectangle", "LSection", "TwoRectangles", "RectSet", "TwoRectangle", "TwoRectanglesOffset", "OffsetTwoRectangles"))
    If shape = "RectSet" Then shape = "LSection"
    If shape <> "Rectangle" And shape <> "LSection" Then shape = "TwoRectangles"
    header = table.RowByKey("H1")
    h1 = table.Length(header + 1, 1, units)
    b1 = table.Length(header + 1, 2, units)
    Require table, header + 1, 1, h1 > 0#, "Высота верхнего прямоугольника должна быть положительной."
    Require table, header + 1, 2, b1 > 0#, "Ширина верхнего прямоугольника должна быть положительной."
    If shape <> "Rectangle" Then
        h2 = table.Length(header + 1, 3, units)
        b2 = table.Length(header + 1, 4, units)
        Require table, header + 1, 3, h2 > 0#, "Высота нижнего прямоугольника должна быть положительной."
        Require table, header + 1, 4, b2 > 0#, "Ширина нижнего прямоугольника должна быть положительной."
        If shape = "LSection" Then Require table, header + 1, 2, b1 < b2, "Для LSection ширина верхнего прямоугольника должна быть меньше нижнего."
    End If
    If shape = "TwoRectangles" Then
        i = table.RowByKey("RectSet.UpperOffsetX")
        offset = table.Length(i, 2, units)
        Require table, i, 2, offset < b2 - GEOM_TOLERANCE And offset + b1 > GEOM_TOLERANCE, _
            "Верхний и нижний прямоугольники должны иметь общий участок по X."
    End If
    Set facade = New CSectionDefinitionBuilder
    facade.Initialize "RectSet"
    Set upper = facade.RectanglePath("Upper", offset, h2, b1, h1)
    If shape = "Rectangle" Then
        Set outer = upper
    Else
        Set lower = facade.RectanglePath("Lower", 0#, 0#, b2, h2)
        Set outer = facade.AdjacentRectangles("Boundary", Array(Array(0#, 0#, b2, h2), Array(offset, h2, b1, h1)))
    End If
    ReDim contourIDs(0 To outer.SegmentCount - 1)
    For i = 1 To outer.SegmentCount
        contourIDs(i - 1) = "CONTOUR_LINE_" & CStr(i)
    Next i
    facade.Outer "OUTER", outer.ID, "Сегмент контура сечения.", contourIDs
    AddRectangleRows facade, table, units, header, 0, offset, h2, b1, h1
    If shape <> "Rectangle" Then AddRectangleRows facade, table, units, header, 4, 0#, 0#, b2, h2
    facade.HorizontalExtent "DIM_B", outer.ID, "B", 0#, -1#, "Общая ширина RectSet-сечения."
    facade.VerticalExtent "DIM_H", outer.ID, "H", GeomMin(0#, offset), -1#, "Общая высота RectSet-сечения."
    facade.HorizontalExtent "DIM_TOP_WIDTH", upper.ID, "B1", h1 + h2, 1#, "Ширина верхнего прямоугольника RectSet-сечения."
    If shape <> "Rectangle" Then
        facade.VerticalExtent "DIM_RIGHT_HEIGHT", lower.ID, "H2", b2, 1#, "Высота нижнего прямоугольника RectSet-сечения."
        If Abs(offset) > GEOM_TOLERANCE Then facade.HorizontalDistance "DIM_UPPER_OFFSET_X", ChrW$(&H394), 0#, offset, h2, 1#, _
            "Горизонтальное смещение верхнего прямоугольника относительно базового положения."
    End If
    Set ISectionRecipe_CreateDefinition = facade.Finish
End Function

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
