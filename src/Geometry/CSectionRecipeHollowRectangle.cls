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
    Dim table As CSectionInputTable, units As CUnitSystem, facade As CSectionDefinitionBuilder  ' Объявляем сборщик описания сечения.
    Dim sizes As THollowInput, faces(1 To 8) As THollowFace
    Dim outerPolygon As CPolygonSpec, openingPolygon As CPolygonSpec
    Set table = context.Table("rngHollowRectangleGeometry"): Set units = context.Units
    ReadDimensions table, units, sizes
    ReadFaces table, units, sizes, faces
    Set outerPolygon = DescribePolygon(sizes.B, sizes.H, 0#, 0#, sizes.R, sizes.GeometrySource)
    Set openingPolygon = DescribePolygon(sizes.Bo, sizes.Ho, sizes.Xo, sizes.Yo, sizes.Ro, sizes.GeometrySource)
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "HollowRectangle"  ' "HollowRectangle" — имя типа сечения.
    AddGeometry facade, outerPolygon, openingPolygon  ' facade — сборщик описания сечения; sizes — подготовленные размеры и радиусы.
    AddReinforcement facade, sizes, faces, outerPolygon, openingPolygon  ' facade — сборщик описания сечения; sizes — размеры и данные формы; faces — параметры всех наружных и внутренних сторон.
    AddDimensions facade, sizes  ' facade — сборщик описания сечения; sizes — размеры и данные формы.
    Set ISectionRecipe_CreateDefinition = facade.Finish  ' Проверяем и возвращаем готовое описание сечения.
    Exit Function
InvalidGeometry:
    SectionRecipeGeometryError "CSectionRecipeHollowRectangle", sizes.GeometrySource
End Function

' Читает активные размеры рецепта, проверяет допустимость и переводит длины в мм через
' CUnitSystem.
Private Sub ReadDimensions(ByVal table As CSectionInputTable, ByVal units As CUnitSystem, ByRef sizes As THollowInput)
    Dim row As Long, column As Long
    row = table.RowByKey("H", 1) + 1
    sizes.H = table.PositiveLength(row, 1, units): sizes.B = table.PositiveLength(row, 2, units): sizes.R = table.NonNegativeLength(row, 3, units)
    sizes.Ho = table.PositiveLength(row, 4, units): sizes.Bo = table.PositiveLength(row, 5, units): sizes.Ro = table.NonNegativeLength(row, 6, units)
    For column = 1 To 6
        sizes.GeometrySource = sizes.GeometrySource & table.SourceReference(row, column) & "; "
    Next column
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
                .AxisDistance = table.NonNegativeLength(row, 2, units)
            End If
            If active Then
                Require table, row, 2, .AxisDistance > .Diameter / 2#, "Расстояние до оси должно быть больше половины диаметра."
                If startColumn = 0 Then
                    startColumn = table.ColumnByCaption(sizes.MainHeader, "StartSetback")
                    endColumn = table.ColumnByCaption(sizes.MainHeader, "EndSetback")
                End If
                .StartSetback = table.NonNegativeLength(row, startColumn, units)
                .EndSetback = table.NonNegativeLength(row, endColumn, units)
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
Private Sub AddGeometry(ByVal facade As CSectionDefinitionBuilder, ByVal outerPolygon As CPolygonSpec, _
        ByVal openingPolygon As CPolygonSpec)
    AddBoundary facade, "OUTER", "Shape", outerPolygon, False
    AddBoundary facade, "OPENING", "Opening.Shape", openingPolygon, True
End Sub

' Добавляет именованную бетонную границу и её сегменты, сохраняя назначенные скругления.
Private Sub AddBoundary(ByVal facade As CSectionDefinitionBuilder, ByVal loopID As String, ByVal shapeID As String, _
        ByVal polygon As CPolygonSpec, ByVal opening As Boolean)
    Dim shape As CGeometryPath, boundary As CGeometryPath, part As String, first As String
    Set shape = facade.RoundedPolygon(shapeID, polygon) ' Линия регистрируется сразу; роль назначается ниже.
    first = "Edge.BL": If Not shape.HasSegmentID(first) Then first = "Corner.TL"
    Set boundary = facade.Chain(loopID & ".Boundary", shape.ID, first, 0#, first, 0#, True, True)
    If opening Then part = "внутреннего контура Opening" Else part = "наружного контура"
    If opening Then
        facade.Opening loopID, boundary.ID, "OUTER", "Прямой участок " & part & " HollowRectangle.", _
            arcComment:="Дуга " & part & " HollowRectangle.", namingPolicy:=SequentialByKind, namingPrefix:="CONTOUR_" & loopID
    Else
        facade.Outer loopID, boundary.ID, "Прямой участок " & part & " HollowRectangle.", _
            arcComment:="Дуга " & part & " HollowRectangle.", namingPolicy:=SequentialByKind, namingPrefix:="CONTOUR_" & loopID
    End If
End Sub

' Возвращает исходные вершины прямоугольной границы в порядке обхода.
Private Function DescribePolygon(ByVal B As Double, ByVal H As Double, ByVal x As Double, ByVal y As Double, _
        ByVal radius As Double, ByVal sourceReference As String) As CPolygonSpec
    Dim polygon As CPolygonSpec
    Set polygon = New CPolygonSpec
    polygon.AddVertex "TL", x - B / 2#, y + H / 2#, radius, sourceReference
    polygon.AddVertex "TR", x + B / 2#, y + H / 2#, radius, sourceReference
    polygon.AddVertex "BR", x + B / 2#, y - H / 2#, radius, sourceReference
    polygon.AddVertex "BL", x - B / 2#, y - H / 2#, radius, sourceReference
    polygon.Seal
    Set DescribePolygon = polygon
End Function

' Описывает активные ряды и их зависимости; отключённая грань не создаёт стержней и подписи.
Private Sub AddReinforcement(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As THollowInput, ByRef faces() As THollowFace, _
        ByVal outerPolygon As CPolygonSpec, ByVal openingPolygon As CPolygonSpec)
    Dim i As Long, base As Long, path As CGeometryPath, options As CRebarRuleOptions  ' Объявляем линии геометрии и условия армирования.
    Dim normalSign As Double, nx As Double, ny As Double, distance As Double, annotationSign As Double
    For i = 1 To 8
        base = (i - 1) Mod 4 + 1
        If faces(i).Diameter = 0# Or faces(base).Count = 0 Or faces(base).Diameter = 0# Then GoTo NextFace
        Set path = AxisPath(facade, sizes, faces, i, outerPolygon, openingPolygon) ' Выбирает наружный/внутренний исходный набор по грани.
        Set options = New CRebarRuleOptions  ' Создаём условия размещения арматуры.
        options.StartSetback = faces(i).StartSetback: options.EndSetback = faces(i).EndSetback  ' StartSetback — отступ начала ряда от начала грани, мм; EndSetback — отступ конца ряда от конца грани, мм.
        ' Publication — PublishPerSourcePosition сохраняет имена по исходным позициям; Coincidence —
        ' KeepPublishedOwner сохраняет первого владельца совпавшего стержня.
        options.SetPublication faces(i).ID, PublishPerSourcePosition, KeepPublishedOwner
        options.IncomingTangentAtJoint = True  ' IncomingTangentAtJoint — True использует входящую касательную на стыке участков; AnnotationGroup — имя общей подписи для всех рядов этой группы.
        options.SourceReference = faces(i).Source  ' SourceReference — адреса ячеек, которые надо исправить при ошибке размещения.
        options.Comment = "HollowRectangle: " & faces(i).Caption & IIf(i > 4, " грань Opening", " наружная грань") & ", ряд 1"  ' Comment — пояснение к стержням данного ряда.
        normalSign = 1#: If base = 1 Or base = 3 Then normalSign = -1#
        If i <= 4 Then
            If base <= 2 Then options.EndMode = ExcludeBoth  ' EndMode — ExcludeBoth исключает точки обоих концов линии.
            ' faces(i).ID & ".Main" — имя нового ряда; path.ID — линия центров стержней;
            ' faces(i).Count — заданное число стержней; faces(i).Diameter — диаметр стержня, мм;
            ' options — отступы, нумерация и условия размещения.
            facade.RebarByCount faces(i).ID & ".Main", path.ID, faces(i).Count, faces(i).Diameter, options
        Else
            ' SourceNormalSign — знак нормали исходного ряда для направления проекции;
            ' SkipMissingProjection — True пропускает позиции без пересечения с целевой линией.
            options.SourceNormalSign = normalSign: options.SkipMissingProjection = True
            ' SuccessfulProjectionOrdinal — True нумерует только успешно спроецированные позиции;
            ' NewPublicationRoot — True создаёт отдельные имена стержней внутреннего ряда.
            options.SuccessfulProjectionOrdinal = True: options.NewPublicationRoot = True
            ' faces(i).ID & ".Main" — имя внутреннего ряда; path.ID — линия, на которую проецируются
            ' позиции; faces(base).ID & ".Main" — исходный наружный ряд; faces(i).Diameter — диаметр
            ' внутреннего стержня, мм; options — сторона проекции, отступы и обработка пропусков.
            facade.RebarProjected faces(i).ID & ".Main", path.ID, faces(base).ID & ".Main", faces(i).Diameter, options
            normalSign = -normalSign
        End If
        If faces(i).D2 > 0# Then
            ' faces(i) — параметры данной стороны; 2 — номер дополнительного ряда; faces(i).Location2
            ' — Stacked поперёк ряда, SideBySide вдоль; faces(i).Binding2 — EachBar все позиции,
            ' EverySecondBar через одну; faces(i).Diameter / 2 + faces(i).D2 / 2 — межосевое
            ' расстояние, мм; normalSign — знак смещения по нормали.
            Set options = ExtraOptions(faces(i), 2, faces(i).Location2, faces(i).Binding2, faces(i).Diameter, faces(i).D2, normalSign)
            ' faces(i).ID & ".Row2" — имя дополнительного ряда; faces(i).ID & ".Main" — исходный ряд,
            ' чьи позиции используются; faces(i).D2 — диаметр дополнительного стержня, мм; options —
            ' смещение и выбор позиций исходного ряда.
            facade.RebarDependent faces(i).ID & ".Row2", faces(i).ID & ".Main", faces(i).D2, options
        End If
        If faces(i).D3 > 0# Then
            distance = 0#
            If faces(i).D2 > 0# And faces(i).Location2 = faces(i).Location3 Then distance = faces(i).D2
            ' faces(i) — параметры данной стороны; 3 — номер дополнительного ряда; faces(i).Location3
            ' — Stacked поперёк ряда, SideBySide вдоль; faces(i).Binding3 — EachBar все позиции,
            ' EverySecondBar через одну; distance — межосевое расстояние, мм; normalSign — знак
            ' смещения по нормали.
            Set options = ExtraOptions(faces(i), 3, faces(i).Location3, faces(i).Binding3, faces(i).Diameter, faces(i).D3, normalSign, distance)
            ' faces(i).ID & ".Row3" — имя дополнительного ряда; faces(i).ID & ".Main" — исходный ряд,
            ' чьи позиции используются; faces(i).D3 — диаметр дополнительного стержня, мм; options —
            ' смещение и выбор позиций исходного ряда.
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
            ' faces(i).ID — группа общей подписи; faces(i).ID & ".Main" — исходный ряд для выбора
            ' крайней линии; 0 — X-компонента направления выбора крайнего ряда; IIf(base = 1, 1, -1) —
            ' Y-компонента направления выбора крайнего ряда; nx — X-компонента направления подписи; ny
            ' — Y-компонента направления подписи.
            facade.RebarRowExtremeLabel faces(i).ID, faces(i).ID & ".Main", 0#, IIf(base = 1, 1#, -1#), nx, ny
        Else
            ' faces(i).ID — группа общей подписи всех её рядов; path.ID — линия, по габаритам которой
            ' привязывается подпись; nx — X-компонента направления подписи; ny — Y-компонента
            ' направления подписи.
            facade.RebarLabel faces(i).ID, path.ID, nx, ny
        End If
NextFace:
    Next i
End Sub

' B охватывает обе дуги, H только прямую. При R_o=0 соседние as задают
' концы целевых прямых независимо от наличия арматуры на соседней грани.
Private Function AxisPath(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As THollowInput, ByRef faces() As THollowFace, _
        ByVal index As Long, ByVal outerPolygon As CPolygonSpec, ByVal openingPolygon As CPolygonSpec) As CGeometryPath
    On Error GoTo InvalidPath
    Dim B As Double, H As Double, radius As Double, x As Double, y As Double, offset As Double, base As Long, shape As CGeometryPath  ' Объявляем линии и параметры геометрии.
    Dim first As String, last As String, firstParameter As Double, lastParameter As Double
    Dim polygon As CPolygonSpec
    Set polygon = outerPolygon
    B = sizes.B: H = sizes.H: radius = sizes.R: offset = faces(index).AxisDistance
    base = (index - 1) Mod 4 + 1
    If index > 4 Then
        Set polygon = openingPolygon
        B = sizes.Bo: H = sizes.Ho: radius = sizes.Ro: x = sizes.Xo: y = sizes.Yo: offset = -offset
        If radius <= GEOM_TOLERANCE Then
            facade.Source faces(index).ID, faces(index).Source  ' faces(index).ID — имя геометрии или ряда; faces(index).Source — адреса исходных ячеек для сообщений об ошибках.
            If base <= 2 Then
                y = y + IIf(base = 1, 1#, -1#) * (H / 2# - offset)
                Set shape = facade.LinePath("Unlimited." & faces(index).ID, x - B / 2#, y, x + B / 2#, y)  ' "Unlimited." & faces(index).ID — имя линии; x - B / 2 — X начала, мм; y — Y начала, мм; x + B / 2 — X конца, мм; y — Y конца, мм.
                ' faces(index).ID — имя ограниченной линии; shape.ID — исходная прямая; 1 —
                ' X-компонента нормали первой границы; 0 — Y-компонента нормали первой границы; x - B
                ' / 2 - faces(7).AxisDistance — первая граница: nX*X+nY*Y=bound, мм; 1 — X-компонента
                ' нормали второй границы; 0 — Y-компонента нормали второй границы; x + B / 2 +
                ' faces(8).AxisDistance — вторая граница: nX*X+nY*Y=bound, мм.
                Set AxisPath = facade.ConstrainLine(faces(index).ID, shape.ID, 1#, 0#, x - B / 2# - faces(7).AxisDistance, 1#, 0#, x + B / 2# + faces(8).AxisDistance)
            Else
                x = x + IIf(base = 3, -1#, 1#) * (B / 2# - offset)
                Set shape = facade.LinePath("Unlimited." & faces(index).ID, x, y - H / 2#, x, y + H / 2#)  ' "Unlimited." & faces(index).ID — имя линии; x — X начала, мм; y - H / 2 — Y начала, мм; x — X конца, мм; y + H / 2 — Y конца, мм.
                ' faces(index).ID — имя ограниченной линии; shape.ID — исходная прямая; 0 —
                ' X-компонента нормали первой границы; 1 — Y-компонента нормали первой границы; y - H
                ' / 2 - faces(6).AxisDistance — первая граница: nX*X+nY*Y=bound, мм; 0 — X-компонента
                ' нормали второй границы; 1 — Y-компонента нормали второй границы; y + H / 2 +
                ' faces(5).AxisDistance — вторая граница: nX*X+nY*Y=bound, мм.
                Set AxisPath = facade.ConstrainLine(faces(index).ID, shape.ID, 0#, 1#, y - H / 2# - faces(6).AxisDistance, 0#, 1#, y + H / 2# + faces(5).AxisDistance)
            End If
            Exit Function
        End If
    End If
    ' "Offset." & faces(index).ID — имя линии центров стержней; Corners(B, H, x, y) — исходные вершины
    ' (X,Y), мм; Array(radius, radius, radius, radius) — исходные радиусы вершин, мм; Array("TL",
    ' "TR", "BR", "BL") — имена исходных вершин; offset — смещение, мм: положительное внутрь,
    ' отрицательное наружу.
    Set shape = facade.OffsetRoundedPolygon("Offset." & faces(index).ID, polygon, offset)
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
            ' faces(index).ID — имя новой линии; shape.ID — исходный контур; first — участок начала;
            ' firstParameter — точка начала участка: 0 начало, 1 конец; last — участок конца;
            ' lastParameter — точка конца участка: 0 начало, 1 конец; base = 1 — True по обходу
            ' исходной линии, False против.
            Set AxisPath = facade.Chain(faces(index).ID, shape.ID, first, firstParameter, last, lastParameter, base = 1)
        ' faces(index).ID — имя новой линии; shape.ID — исходный контур; "Edge.BL" — участок начала; 0
        ' — точка начала участка: 0 начало, 1 конец; "Edge.BL" — участок конца; 1 — точка конца
        ' участка: 0 начало, 1 конец.
        Case 3: Set AxisPath = facade.Chain(faces(index).ID, shape.ID, "Edge.BL", 0#, "Edge.BL", 1#)
        ' faces(index).ID — имя новой линии; shape.ID — исходный контур; "Edge.TR" — участок начала; 1
        ' — точка начала участка: 0 начало, 1 конец; "Edge.TR" — участок конца; 0 — точка конца
        ' участка: 0 начало, 1 конец; False — против обхода исходной линии.
        Case 4: Set AxisPath = facade.Chain(faces(index).ID, shape.ID, "Edge.TR", 1#, "Edge.TR", 0#, False)
    End Select
    Exit Function
InvalidPath:
    SectionRecipeGeometryError "CSectionRecipeHollowRectangle", faces(index).Source
End Function

' Задаёт привязку и смещение дополнительного ряда без повторной раскладки исходных позиций.
Private Function ExtraOptions(ByRef face As THollowFace, ByVal rowNumber As Long, ByVal location As String, ByVal binding As String, ByVal baseDiameter As Double, ByVal diameter As Double, ByVal normalSign As Double, _
        Optional ByVal interveningDiameterSum As Double = 0#) As CRebarRuleOptions
    Dim options As CRebarRuleOptions  ' Объявляем условия размещения арматуры.
    Set options = New CRebarRuleOptions  ' Создаём условия размещения арматуры.
    ' NormalOffset — знаковое смещение по нормали к исходному ряду, мм; TangentOffset — знаковое
    ' смещение вдоль исходного ряда, мм.
    Dim axis As ERebarOffsetAxis, directionSign As Double
    axis = RowNormal: directionSign = normalSign
    If location = "SideBySide" Then axis = RowTangent: directionSign = 1#
    options.SourceReference = face.Source
    options.ConfigureDependentContact baseDiameter, diameter, axis, directionSign, (binding = "EverySecondBar"), interveningDiameterSum
    options.SetPublication face.ID, PublishPerSourcePosition, KeepPublishedOwner
    options.Comment = "HollowRectangle: " & face.Caption & IIf(Left$(face.ID, 8) = "Opening.", " грань Opening", " наружная грань") & ", ряд " & CStr(rowNumber)  ' Comment — пояснение к стержням данного ряда.
    Set ExtraOptions = options  ' Возвращаем условия размещения дополнительного ряда.
End Function

' Описывает габариты и размерные привязки рецепта отдельно от расчётных контуров.
Private Sub AddDimensions(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As THollowInput)
    ' "DIM_B" — имя размера; "B" — обозначение на схеме; -sizes.B / 2 — X первой точки, мм; sizes.B /
    ' 2 — X второй точки, мм; sizes.H / 2 — уровень исходной размерной линии Y, мм; 1 — вынос размера
    ' вверх; comment — пояснение к размеру.
    facade.HorizontalDistance "DIM_B", "B", -sizes.B / 2#, sizes.B / 2#, sizes.H / 2#, 1#, "Наружная ширина HollowRectangle."
    ' "DIM_H" — имя размера; "OUTER.Boundary" — контур, по которому измеряется высота; "H" —
    ' обозначение на схеме; -sizes.B / 2 — положение исходной размерной линии X, мм; -1 — вынос
    ' размера влево; comment — пояснение к размеру.
    facade.VerticalExtent "DIM_H", "OUTER.Boundary", "H", -sizes.B / 2#, -1#, "Наружная высота HollowRectangle."
    ' "DIM_B_OPENING" — имя размера; "B_o" — обозначение на схеме; sizes.Xo - sizes.Bo / 2 — X первой
    ' точки, мм; sizes.Xo + sizes.Bo / 2 — X второй точки, мм; -sizes.H / 2 — уровень исходной
    ' размерной линии Y, мм; -1 — вынос размера вниз; comment — пояснение к размеру.
    facade.HorizontalDistance "DIM_B_OPENING", "B_o", sizes.Xo - sizes.Bo / 2#, sizes.Xo + sizes.Bo / 2#, -sizes.H / 2#, -1#, "Ширина внутреннего отверстия Opening."
    ' "DIM_H_OPENING" — имя размера; "OPENING.Boundary" — контур, по которому измеряется высота; "H_o"
    ' — обозначение на схеме; sizes.B / 2 — положение исходной размерной линии X, мм; 1 — вынос
    ' размера вправо; comment — пояснение к размеру.
    facade.VerticalExtent "DIM_H_OPENING", "OPENING.Boundary", "H_o", sizes.B / 2#, 1#, "Высота внутреннего отверстия Opening."
End Sub

' Проверяет ограничение рецепта; ошибка содержит фактический адрес параметра во входной таблице.
Private Sub Require(ByVal table As CSectionInputTable, ByVal row As Long, ByVal column As Long, ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then SectionRecipeError "CSectionRecipeHollowRectangle", reason, table.SourceReference(row, column)
End Sub
