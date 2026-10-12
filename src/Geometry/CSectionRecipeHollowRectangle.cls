VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionRecipeHollowRectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Декларативный рецепт полого прямоугольника с конечными проекциями арматуры
' ==========================================================================
Implements ISectionRecipe

' Рецепт скругленной рамки: Outer/Opening, наружные ряды и их проекции.
' Математика дуг, пересечений, позиций и сетки остается в общем ядре.
Private Type THollowInput
    B As Double
    H As Double
    R As Double
    Bo As Double
    Ho As Double
    Ro As Double
    Xo As Double
    Yo As Double
    MainHeader As Long
    ExtraHeader As Long
    GeometrySource As String
End Type

Private Type THollowFace
    ID As String
    Caption As String
    AxisDistance As Double
    Diameter As Double
    Count As Long
    StartSetback As Double
    EndSetback As Double
    D2 As Double
    D3 As Double
    Location2 As String
    Location3 As String
    Binding2 As String
    Binding3 As String
    Source As String
End Type

' Описывает контуры, арматурные правила и размеры выбранного рецепта. Сетку и физические стержни
' строит общий CSectionBuilder.
Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
    On Error GoTo InvalidGeometry
    Dim table As CSectionInputTable, units As CUnitSystem, facade As CSectionDefinitionBuilder
    Dim sizes As THollowInput, faces(1 To 8) As THollowFace
    Set table = context.Table("rngHollowRectangleGeometry"): Set units = context.Units
    ReadDimensions table, units, sizes
    ReadFaces table, units, sizes, faces
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "HollowRectangle"
    AddGeometry facade, sizes
    AddReinforcement facade, sizes, faces
    AddDimensions facade, sizes
    Set ISectionRecipe_CreateDefinition = facade.Finish
    Exit Function
InvalidGeometry:
    SectionRecipeGeometryError "CSectionRecipeHollowRectangle", sizes.GeometrySource
End Function

' Читает активные размеры рецепта, проверяет допустимость и переводит длины в мм через
' CUnitSystem.
Private Sub ReadDimensions(ByVal table As CSectionInputTable, ByVal units As CUnitSystem, ByRef sizes As THollowInput)
    Dim row As Long, column As Long
    row = table.RowByKey("H", 1) + 1
    sizes.H = table.Length(row, 1, units): sizes.B = table.Length(row, 2, units): sizes.R = table.Length(row, 3, units)
    sizes.Ho = table.Length(row, 4, units): sizes.Bo = table.Length(row, 5, units): sizes.Ro = table.Length(row, 6, units)
    For column = 1 To 6
        If column <> 3 And column <> 6 Then Require table, row, column, table.Length(row, column, units) > 0#, "Размер должен быть положительным."
        sizes.GeometrySource = sizes.GeometrySource & table.SourceReference(row, column) & "; "
    Next column
    Require table, row, 3, sizes.R >= 0#, "Радиус R не может быть отрицательным."
    Require table, row, 6, sizes.Ro >= 0#, "Радиус R_o не может быть отрицательным."
    row = table.RowByKey("HollowRectangle.InnerOffsetX")
    sizes.Xo = table.Length(row, 2, units): sizes.GeometrySource = sizes.GeometrySource & table.SourceReference(row, 2) & "; "
    row = table.RowByKey("HollowRectangle.InnerOffsetY")
    sizes.Yo = table.Length(row, 2, units): sizes.GeometrySource = sizes.GeometrySource & table.SourceReference(row, 2)
    sizes.MainHeader = table.RowByKey("as", 2): sizes.ExtraHeader = table.RowByKey("d2", 2)
End Sub

' Отступы всех граней отверстия нужны даже при выключенном армировании:
' при R_o=0 они ограничивают соседние проекции. Остальные неактивные поля не читаются.
Private Sub ReadFaces(ByVal table As CSectionInputTable, ByVal units As CUnitSystem, ByRef sizes As THollowInput, ByRef faces() As THollowFace)
    Dim ids As Variant, captions As Variant, aliases As Variant, i As Long, row As Long, extra As Long, active As Boolean
    Dim startColumn As Long, endColumn As Long
    ids = Array("B.Top", "B.Bottom", "H.Left", "H.Right", "Opening.B.Top", "Opening.B.Bottom", "Opening.H.Left", "Opening.H.Right")
    captions = Array("B - верхняя", "B - нижняя", "H - левая", "H - правая", "B_o - верхняя", "B_o - нижняя", "H_o - левая", "H_o - правая")
    aliases = Array("B-top", "B-bottom", "H-left", "H-right", "bo-top", "bo-bottom", "ho-left", "ho-right")
    For i = 1 To 8
        With faces(i)
            .ID = CStr(ids(i - 1)): .Caption = CStr(captions(i - 1))
            row = table.RowByAnyKey(Array(.Caption, .ID, aliases(i - 1)), 1, sizes.MainHeader + 1, sizes.MainHeader + 8)
            .Source = table.SourceReference(row, 2)
            If i <= 4 Then .Count = table.Count(row, 4)
            .Diameter = table.OptionalDiameter(row, 3, units)
            active = (.Diameter > 0# And (i > 4 Or .Count > 0))
            If i > 4 Then active = active And faces(i - 4).Count > 0 And faces(i - 4).Diameter > 0#
            If i > 4 Or active Then
                .AxisDistance = table.Length(row, 2, units)
                Require table, row, 2, .AxisDistance >= 0#, "Расстояние до оси не может быть отрицательным."
            End If
            If active Then
                Require table, row, 2, .AxisDistance > .Diameter / 2#, "Расстояние до оси должно быть больше половины диаметра."
                startColumn = table.ColumnByCaption(sizes.MainHeader, "StartSetback")
                endColumn = table.ColumnByCaption(sizes.MainHeader, "EndSetback")
                .StartSetback = table.Length(row, startColumn, units)
                .EndSetback = table.Length(row, endColumn, units)
                Require table, row, startColumn, .StartSetback >= 0#, "StartSetback не может быть отрицательным."
                Require table, row, endColumn, .EndSetback >= 0#, "EndSetback не может быть отрицательным."
                .Source = .Source & "; " & table.SourceReference(row, startColumn) & "; " & table.SourceReference(row, endColumn)
                extra = table.RowByAnyKey(Array(.Caption, .ID, aliases(i - 1)), 1, sizes.ExtraHeader + 1, sizes.ExtraHeader + 8)
                .D2 = table.OptionalDiameter(extra, 2, units): .D3 = table.OptionalDiameter(extra, 5, units)
                If .D2 > 0# Then
                    .Location2 = table.Choice(extra, 3, Array("Stacked", "SideBySide"))
                    .Binding2 = table.Choice(extra, 4, Array("EachBar", "EverySecondBar"))
                End If
                If .D3 > 0# Then
                    .Location3 = table.Choice(extra, 6, Array("Stacked", "SideBySide"))
                    .Binding3 = table.Choice(extra, 7, Array("EachBar", "EverySecondBar"))
                End If
            End If
        End With
    Next i
End Sub

' Описывает аналитическую бетонную границу рецепта через общий фасад, без построения расчётной
' сетки.
Private Sub AddGeometry(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As THollowInput)
    AddBoundary facade, "OUTER", "Outer.Shape", sizes.B, sizes.H, sizes.R, 0#, 0#, False
    AddBoundary facade, "OPENING", "Opening.Shape", sizes.Bo, sizes.Ho, sizes.Ro, sizes.Xo, sizes.Yo, True
End Sub

' Добавляет именованную бетонную границу и её сегменты, сохраняя назначенные скругления.
Private Sub AddBoundary(ByVal facade As CSectionDefinitionBuilder, ByVal loopID As String, ByVal shapeID As String, _
        ByVal B As Double, ByVal H As Double, ByVal radius As Double, ByVal x As Double, ByVal y As Double, ByVal opening As Boolean)
    Dim shape As CGeometryPath, boundary As CGeometryPath, edge As TRegionEdge, ids() As String, i As Long, part As String, first As String
    Set shape = facade.RoundedPolygon(shapeID, Corners(B, H, x, y), Array(radius, radius, radius, radius), Array("TL", "TR", "BR", "BL"))
    first = "Edge.BL": If Not shape.HasSegmentID(first) Then first = "Corner.TL"
    Set boundary = facade.Chain(loopID & ".Boundary", shape.ID, first, 0#, first, 0#, True, True)
    ReDim ids(0 To boundary.SegmentCount - 1)
    For i = 1 To boundary.SegmentCount
        boundary.CopySegment i, edge
        If edge.Sweep = 0# Then ids(i - 1) = "CONTOUR_" & loopID & "_LINE_" & CStr(i) Else ids(i - 1) = "CONTOUR_" & loopID & "_ARC_" & CStr(i)
    Next i
    If opening Then part = "внутреннего контура Opening" Else part = "наружного контура"
    If opening Then
        facade.Opening loopID, boundary.ID, "OUTER", "Прямой участок " & part & " HollowRectangle.", ids, "Дуга " & part & " HollowRectangle."
    Else
        facade.Outer loopID, boundary.ID, "Прямой участок " & part & " HollowRectangle.", ids, "Дуга " & part & " HollowRectangle."
    End If
End Sub

' Возвращает исходные вершины прямоугольной границы в порядке обхода.
Private Function Corners(ByVal B As Double, ByVal H As Double, ByVal x As Double, ByVal y As Double) As Variant
    Corners = Array(Array(x - B / 2#, y + H / 2#), Array(x + B / 2#, y + H / 2#), _
        Array(x + B / 2#, y - H / 2#), Array(x - B / 2#, y - H / 2#))
End Function

' Описывает активные ряды и их зависимости; отключённая грань не создаёт стержней и подписи.
Private Sub AddReinforcement(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As THollowInput, ByRef faces() As THollowFace)
    Dim i As Long, base As Long, path As CGeometryPath, options As CRebarRuleOptions
    Dim normalSign As Double, nx As Double, ny As Double, distance As Double, annotationSign As Double
    For i = 1 To 8
        base = (i - 1) Mod 4 + 1
        If faces(i).Diameter = 0# Or faces(base).Count = 0 Or faces(base).Diameter = 0# Then GoTo NextFace
        Set path = AxisPath(facade, sizes, faces, i)
        Set options = New CRebarRuleOptions
        options.StartSetback = faces(i).StartSetback: options.EndSetback = faces(i).EndSetback
        options.Publication = PublishPerSourcePosition: options.Coincidence = KeepPublishedOwner
        options.IncomingTangentAtJoint = True: options.AnnotationGroup = faces(i).ID
        options.SourceReference = faces(i).Source
        options.Comment = "HollowRectangle: " & faces(i).Caption & IIf(i > 4, " грань Opening", " наружная грань") & ", ряд 1"
        normalSign = 1#: If base = 1 Or base = 3 Then normalSign = -1#
        If i <= 4 Then
            If base <= 2 Then options.EndMode = ExcludeBoth
            facade.RebarByCount faces(i).ID & ".Main", path.ID, faces(i).Count, faces(i).Diameter, options
        Else
            options.SourceNormalSign = normalSign: options.SkipMissingProjection = True
            options.SuccessfulProjectionOrdinal = True: options.NewPublicationRoot = True
            facade.RebarProjected faces(i).ID & ".Main", path.ID, faces(base).ID & ".Main", faces(i).Diameter, options
            normalSign = -normalSign
        End If
        If faces(i).D2 > 0# Then
            Set options = ExtraOptions(faces(i), 2, faces(i).Location2, faces(i).Binding2, faces(i).Diameter / 2# + faces(i).D2 / 2#, normalSign)
            facade.RebarDependent faces(i).ID & ".Row2", faces(i).ID & ".Main", faces(i).D2, options
        End If
        If faces(i).D3 > 0# Then
            distance = faces(i).Diameter / 2# + faces(i).D3 / 2#
            If faces(i).D2 > 0# And faces(i).Location2 = faces(i).Location3 Then distance = distance + faces(i).D2
            Set options = ExtraOptions(faces(i), 3, faces(i).Location3, faces(i).Binding3, distance, normalSign)
            facade.RebarDependent faces(i).ID & ".Row3", faces(i).ID & ".Main", faces(i).D3, options
        End If
        nx = 0#: ny = 0#: annotationSign = 1#: If i > 4 Then annotationSign = -1#
        Select Case base
            Case 1: ny = annotationSign
            Case 2: ny = -annotationSign
            Case 3: nx = -annotationSign
            Case 4: nx = annotationSign
        End Select
        If base <= 2 Then
            facade.RebarRowExtremeLabel faces(i).ID, faces(i).ID & ".Main", 0#, IIf(base = 1, 1#, -1#), nx, ny, faces(i).AxisDistance
        Else
            facade.RebarLabel faces(i).ID, path.ID, nx, ny, faces(i).AxisDistance
        End If
NextFace:
    Next i
End Sub

' B охватывает обе дуги, H только прямую. При R_o=0 соседние as задают
' концы целевых прямых независимо от наличия арматуры на соседней грани.
Private Function AxisPath(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As THollowInput, ByRef faces() As THollowFace, ByVal index As Long) As CGeometryPath
    On Error GoTo InvalidPath
    Dim B As Double, H As Double, radius As Double, x As Double, y As Double, offset As Double, base As Long, shape As CGeometryPath
    Dim first As String, last As String, firstParameter As Double, lastParameter As Double
    B = sizes.B: H = sizes.H: radius = sizes.R: offset = faces(index).AxisDistance
    base = (index - 1) Mod 4 + 1
    If index > 4 Then
        B = sizes.Bo: H = sizes.Ho: radius = sizes.Ro: x = sizes.Xo: y = sizes.Yo: offset = -offset
        If radius <= GEOM_TOLERANCE Then
            facade.Source faces(index).ID, faces(index).Source
            If base <= 2 Then
                y = y + IIf(base = 1, 1#, -1#) * (H / 2# - offset)
                Set shape = facade.LinePath("Unlimited." & faces(index).ID, x - B / 2#, y, x + B / 2#, y)
                Set AxisPath = facade.ConstrainLine(faces(index).ID, shape.ID, 1#, 0#, x - B / 2# - faces(7).AxisDistance, _
                    1#, 0#, x + B / 2# + faces(8).AxisDistance)
            Else
                x = x + IIf(base = 3, -1#, 1#) * (B / 2# - offset)
                Set shape = facade.LinePath("Unlimited." & faces(index).ID, x, y - H / 2#, x, y + H / 2#)
                Set AxisPath = facade.ConstrainLine(faces(index).ID, shape.ID, 0#, 1#, y - H / 2# - faces(6).AxisDistance, _
                    0#, 1#, y + H / 2# + faces(5).AxisDistance)
            End If
            Exit Function
        End If
    End If
    Set shape = facade.OffsetRoundedPolygon("Offset." & faces(index).ID, Corners(B, H, x, y), Array(radius, radius, radius, radius), Array("TL", "TR", "BR", "BL"), offset)
    Select Case base
        Case 1, 2
            first = "Edge.BL": last = "Edge.TR"
            If base = 1 Then firstParameter = 1# Else lastParameter = 1#
            If Not shape.HasSegmentID(first) Then
                If base = 1 Then first = "Corner.TL": firstParameter = 0# Else first = "Corner.BL": firstParameter = 1#
            End If
            If Not shape.HasSegmentID(last) Then
                If base = 1 Then last = "Corner.TR": lastParameter = 1# Else last = "Corner.BR": lastParameter = 0#
            End If
            Set AxisPath = facade.Chain(faces(index).ID, shape.ID, first, firstParameter, last, lastParameter, base = 1)
        Case 3: Set AxisPath = facade.Chain(faces(index).ID, shape.ID, "Edge.BL", 0#, "Edge.BL", 1#)
        Case 4: Set AxisPath = facade.Chain(faces(index).ID, shape.ID, "Edge.TR", 1#, "Edge.TR", 0#, False)
    End Select
    Exit Function
InvalidPath:
    SectionRecipeGeometryError "CSectionRecipeHollowRectangle", faces(index).Source
End Function

' Задаёт привязку и смещение дополнительного ряда без повторной раскладки исходных позиций.
Private Function ExtraOptions(ByRef face As THollowFace, ByVal rowNumber As Long, ByVal location As String, ByVal binding As String, ByVal distance As Double, ByVal normalSign As Double) As CRebarRuleOptions
    Dim options As CRebarRuleOptions
    Set options = New CRebarRuleOptions
    If location = "Stacked" Then options.NormalOffset = normalSign * distance Else options.TangentOffset = distance
    options.EverySecond = (binding = "EverySecondBar"): options.Publication = PublishPerSourcePosition
    options.Coincidence = KeepPublishedOwner: options.AnnotationGroup = face.ID: options.SourceReference = face.Source
    options.Comment = "HollowRectangle: " & face.Caption & IIf(Left$(face.ID, 8) = "Opening.", " грань Opening", " наружная грань") & ", ряд " & CStr(rowNumber)
    Set ExtraOptions = options
End Function

' Описывает габариты и размерные привязки рецепта отдельно от расчётных контуров.
Private Sub AddDimensions(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As THollowInput)
    facade.HorizontalDistance "DIM_B", "B", -sizes.B / 2#, sizes.B / 2#, sizes.H / 2#, 1#, "Наружная ширина HollowRectangle."
    facade.VerticalExtent "DIM_H", "OUTER.Boundary", "H", -sizes.B / 2#, -1#, "Наружная высота HollowRectangle."
    facade.HorizontalDistance "DIM_B_OPENING", "B_o", sizes.Xo - sizes.Bo / 2#, sizes.Xo + sizes.Bo / 2#, -sizes.H / 2#, -1#, "Ширина внутреннего отверстия Opening."
    facade.VerticalExtent "DIM_H_OPENING", "OPENING.Boundary", "H_o", sizes.B / 2#, 1#, "Высота внутреннего отверстия Opening."
End Sub

' Проверяет ограничение рецепта; ошибка содержит фактический адрес параметра во входной таблице.
Private Sub Require(ByVal table As CSectionInputTable, ByVal row As Long, ByVal column As Long, ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then SectionRecipeError "CSectionRecipeHollowRectangle", reason, table.SourceReference(row, column)
End Sub
