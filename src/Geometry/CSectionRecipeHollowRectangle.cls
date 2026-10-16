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
' Прямоугольное бетонное сечение с отверстием
' ==========================================================================
' Класс описывает наружный прямоугольник и внутреннее отверстие со скруглёнными углами.
' Отверстие может быть смещено относительно центра наружного контура.
' Арматура вдоль наружных сторон задаётся по числу стержней. Положения стержней
' у отверстия получают проекцией наружной арматуры; по каждой стороне могут быть
' ещё два дополнительных ряда арматуры.
'
' Рецепт описывает границы бетона и отверстия, расположение арматуры и аннотации
' схемы (размеры сечения и отверстия, количество и диаметры стержней). Построение дуг,
' проекцию стержней и создание бетонной сетки выполняют общие классы.

Implements ISectionRecipe

Private Type THollowInput
    B As Double ' Ширина наружного контура, мм.
    H As Double ' Высота наружного контура, мм.
    R As Double ' Радиус скругления наружного контура, мм.
    Bo As Double ' Ширина отверстия, мм.
    Ho As Double ' Высота отверстия, мм.
    Ro As Double ' Радиус скругления отверстия, мм.
    Xo As Double ' Смещение центра отверстия по X, мм.
    Yo As Double ' Смещение центра отверстия по Y, мм.
    MainHeader As Long ' Строка заголовка основной таблицы армирования во входных данных.
    ExtraHeader As Long ' Строка заголовка дополнительных рядов арматуры во входных данных.
    GeometrySource As String ' Адрес исходных размеров для пояснения геометрических ошибок.
End Type

Private Type THollowFace
    ID As String ' Имя стороны для привязки линий арматуры и подписей.
    Caption As String ' Понятное название стороны для сообщений об ошибках.
    AxisDistance As Double ' Расстояние от бетонной границы до центров основного ряда арматуры, мм.
    Diameter As Double ' Диаметр стержней основного ряда, мм.
    Count As Long ' Число стержней основного ряда.
    StartSetback As Double ' Отступ первого стержня от начала стороны, мм.
    EndSetback As Double ' Отступ последнего стержня от конца стороны, мм.
    Diameters(2 To 3) As Double ' Диаметры второго и третьего рядов арматуры, мм.
    Locations(2 To 3) As String ' Расположение каждого дополнительного ряда относительно основного.
    Bindings(2 To 3) As String ' Выбор основных стержней для второго и третьего рядов.
    Source As String ' Адрес исходных данных для пояснения ошибок.
End Type

' Читает размеры сечения и отверстия, параметры арматуры вдоль их сторон. Формирует
' описание границ бетона, размещения стержней и аннотаций схемы (размеры сечения
' и отверстия, суммарное количество стержней каждого диаметра по каждой стороне).
Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
    On Error GoTo InvalidGeometry
    Dim table As CSectionInputTable, units As CUnitSystem, facade As CSectionDefinitionBuilder
    Dim sizes As THollowInput, faces(1 To 8) As THollowFace
    Dim outerPolygon As CPolygonSpec, openingPolygon As CPolygonSpec
    Set table = context.Table("rngHollowRectangleGeometry"): Set units = context.Units ' Берём копию входной таблицы и единицы измерения.
    ReadDimensions table, units, sizes ' Читаем и проверяем размеры сечения.
    ReadFaces table, units, sizes, faces ' Читаем армирование наружных и внутренних сторон.
    Set outerPolygon = DescribePolygon(sizes.B, sizes.H, 0#, 0#, sizes.R, sizes.GeometrySource) ' Наружный прямоугольник с центром в начале координат.
    Set openingPolygon = DescribePolygon(sizes.Bo, sizes.Ho, sizes.Xo, sizes.Yo, sizes.Ro, sizes.GeometrySource) ' Отверстие с собственными размерами, центром и радиусом.
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "HollowRectangle" ' Создаём сборщик описания сечения.
    AddGeometry facade, outerPolygon, openingPolygon ' Добавляем подготовленные границы бетона.
    AddReinforcement facade, sizes, faces, outerPolygon, openingPolygon ' Задаём ряды и их общие подписи по параметрам таблицы.
    AddDimensions facade, sizes ' Добавляем габариты и размеры сечения.
    Set ISectionRecipe_CreateDefinition = facade.Finish ' Завершаем и возвращаем проверенное описание сечения.
    Exit Function
InvalidGeometry:
    SectionRecipeGeometryError "CSectionRecipeHollowRectangle", sizes.GeometrySource
End Function

' Заполняет sizes наружными размерами, размерами и смещением отверстия в миллиметрах.
' Проверяет размеры и радиусы, сохраняет адреса полей для сообщений об ошибках.
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

' Заполняет faces параметрами четырёх наружных и четырёх внутренних сторон.
' У выключенных рядов не читает ненужные параметры. Внутренние as читает всегда:
' они ограничивают линии у отверстия даже при отсутствии арматуры на соседней стороне.
Private Sub ReadFaces(ByVal table As CSectionInputTable, ByVal units As CUnitSystem, ByRef sizes As THollowInput, ByRef faces() As THollowFace)
    Dim ids As Variant, captions As Variant, aliases As Variant, i As Long, row As Long, extra As Long, active As Boolean
    Dim startColumn As Long, endColumn As Long, rowNumber As Long, column As Long, keys As Variant
    ids = Array("B.Top", "B.Bottom", "H.Left", "H.Right", "Opening.B.Top", "Opening.B.Bottom", "Opening.H.Left", "Opening.H.Right")
    captions = Array("B - верхняя", "B - нижняя", "H - левая", "H - правая", "B_o - верхняя", "B_o - нижняя", "H_o - левая", "H_o - правая")
    aliases = Array("B-top", "B-bottom", "H-left", "H-right", "bo-top", "bo-bottom", "ho-left", "ho-right")
    For i = 1 To 8
        With faces(i)
            .ID = CStr(ids(i - 1)): .Caption = CStr(captions(i - 1))
            keys = Array(.Caption, .ID, aliases(i - 1))
            row = table.RowByAnyKey(keys, 1, sizes.MainHeader + 1, sizes.MainHeader + 8)
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
                extra = table.RowByAnyKey(keys, 1, sizes.ExtraHeader + 1, sizes.ExtraHeader + 8)
                .Diameters(2) = table.OptionalDiameter(extra, 2, units)
                .Diameters(3) = table.OptionalDiameter(extra, 5, units)
                For rowNumber = 2 To 3
                    If .Diameters(rowNumber) > 0# Then
                        column = 2 + (rowNumber - 2) * 3 ' В таблице на каждый ряд отведены диаметр, расположение и привязка.
                        .Locations(rowNumber) = table.Choice(extra, column + 1, Array("Stacked", "SideBySide"))
                        .Bindings(rowNumber) = table.Choice(extra, column + 2, Array("EachBar", "EverySecondBar"))
                    End If
                Next rowNumber
            End If
        End With
    Next i
End Sub

' Добавляет в описание сечения наружную границу и отверстие по двум подготовленным многоугольникам.
Private Sub AddGeometry(ByVal facade As CSectionDefinitionBuilder, ByVal outerPolygon As CPolygonSpec, ByVal openingPolygon As CPolygonSpec)
    AddBoundary facade, "OUTER", "Shape", outerPolygon, False
    AddBoundary facade, "OPENING", "Opening.Shape", openingPolygon, True
End Sub

' Создаёт замкнутую линию по вершинам и радиусам, задаёт порядок обхода и имена её участков.
' Назначает её наружной границей или отверстием внутри OUTER, в зависимости от opening.
Private Sub AddBoundary(ByVal facade As CSectionDefinitionBuilder, ByVal loopID As String, ByVal shapeID As String, ByVal polygon As CPolygonSpec, ByVal opening As Boolean)
    Dim shape As CGeometryPath, boundary As CGeometryPath, part As String, first As String
    Set shape = facade.RoundedPolygon(shapeID, polygon) ' Создаём контур по подготовленным вершинам и радиусам.
    first = "Edge.BL": If Not shape.HasSegmentID(first) Then first = "Corner.TL"
    Set boundary = facade.Chain(loopID & ".Boundary", shape.ID, first, 0#, first, 0#, True, True) ' Полный обход контура от выбранного начального участка.
    If opening Then part = "внутреннего контура Opening" Else part = "наружного контура"
    If opening Then
        facade.Opening loopID, boundary.ID, "OUTER", "Прямой участок " & part & " HollowRectangle.", _
            arcComment:="Дуга " & part & " HollowRectangle.", namingPolicy:=SequentialByKind, namingPrefix:="CONTOUR_" & loopID ' Назначаем отверстие внутри наружной границы OUTER.
    Else
        facade.Outer loopID, boundary.ID, "Прямой участок " & part & " HollowRectangle.", _
            arcComment:="Дуга " & part & " HollowRectangle.", namingPolicy:=SequentialByKind, namingPrefix:="CONTOUR_" & loopID ' Назначаем эту замкнутую линию наружной границей бетона.
    End If
End Sub

' Возвращает завершённое описание четырёх углов прямоугольника с заданным радиусом.
' B и H задают габариты, x и y — центр; нулевой радиус оставляет углы без скругления.
Private Function DescribePolygon(ByVal B As Double, ByVal H As Double, ByVal x As Double, ByVal y As Double, ByVal radius As Double, ByVal sourceReference As String) As CPolygonSpec
    Dim polygon As CPolygonSpec
    Set polygon = New CPolygonSpec
    polygon.AddVertex "TL", x - B / 2#, y + H / 2#, radius, sourceReference, cornerID:="Corner.TL" ' Верхний левый угол и имя его скругления; размеры в мм.
    polygon.AddVertex "TR", x + B / 2#, y + H / 2#, radius, sourceReference, "Edge.TR", "Corner.TR" ' Правая грань и верхнее правое скругление; размеры в мм.
    polygon.AddVertex "BR", x + B / 2#, y - H / 2#, radius, sourceReference, cornerID:="Corner.BR" ' Нижний правый угол и имя его скругления; размеры в мм.
    polygon.AddVertex "BL", x - B / 2#, y - H / 2#, radius, sourceReference, "Edge.BL", "Corner.BL" ' Левая грань и нижнее левое скругление; размеры в мм.
    polygon.Seal ' Проверяем вершины и завершаем их описание.
    Set DescribePolygon = polygon ' Возвращаем подготовленный многоугольник.
End Function

' Задаёт арматуру вдоль наружных сторон по числу стержней, а вдоль отверстия — проекцией наружной арматуры.
' Добавляет связанные ряды и общую подпись каждой стороны по фактически созданной арматуре.
Private Sub AddReinforcement(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As THollowInput, ByRef faces() As THollowFace, _
        ByVal outerPolygon As CPolygonSpec, ByVal openingPolygon As CPolygonSpec)
    Dim i As Long, base As Long, path As CGeometryPath, options As CRebarRuleOptions, rowNumber As Long, rowComment As String
    Dim normalSign As Double, nx As Double, ny As Double, distance As Double, annotationSign As Double
    For i = 1 To 8
        base = (i - 1) Mod 4 + 1
        If faces(i).Diameter = 0# Or faces(base).Count = 0 Or faces(base).Diameter = 0# Then GoTo NextFace
        Set path = AxisPath(facade, sizes, faces, i, outerPolygon, openingPolygon) ' Получаем линию центров стержней выбранной стороны с отступом as.
        Set options = New CRebarRuleOptions ' Создаём настройки размещения этого ряда.
        options.StartSetback = faces(i).StartSetback: options.EndSetback = faces(i).EndSetback ' Отступ от начала и, если задан, от конца линии, мм.
        options.AnnotationGroup = faces(i).ID ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой стороны.
        options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
        options.Coincidence = KeepPublishedOwner ' При совпадении центров сохраняем стержень, который был добавлен раньше.
        options.IncomingTangentAtJoint = True ' На стыке используем направление предыдущего участка.
        options.SourceReference = faces(i).Source ' Адреса полей для сообщения об ошибке размещения.
        rowComment = "HollowRectangle: " & faces(i).Caption & IIf(i > 4, " грань Opening", " наружная грань") & ", ряд "
        options.Comment = rowComment & "1"
        normalSign = 1#: If base = 1 Or base = 3 Then normalSign = -1#
        If i <= 4 Then
            If base <= 2 Then options.EndMode = ExcludeBoth ' Исключаем стержни на обоих концах линии.
            facade.RebarByCount faces(i).ID & ".Main", path.ID, faces(i).Count, faces(i).Diameter, options ' Задаём основной ряд по числу стержней и диаметру.
        Else
            options.SourceNormalSign = normalSign: options.SkipMissingProjection = True ' Проекция идёт к отверстию; пропускаем лучи без пересечения.
            options.SuccessfulProjectionOrdinal = True: options.NewPublicationRoot = True ' Нумеруем только найденные проекции как самостоятельный ряд.
            facade.RebarProjected faces(i).ID & ".Main", path.ID, faces(base).ID & ".Main", faces(i).Diameter, options ' Проецируем стержни наружного ряда на линию у отверстия.
            normalSign = -normalSign
        End If
        For rowNumber = 2 To 3
            If faces(i).Diameters(rowNumber) > 0# Then
                distance = 0#
                If rowNumber = 3 And faces(i).Locations(2) = faces(i).Locations(3) Then distance = faces(i).Diameters(2) ' Обходим номинальный второй ряд в том же направлении.
                Set options = RecipeContactOptions(faces(i).Source, faces(i).Locations(rowNumber), faces(i).Bindings(rowNumber), _
                    faces(i).Diameter, faces(i).Diameters(rowNumber), normalSign, interveningDiameterSum:=distance)
                options.AnnotationGroup = faces(i).ID ' Одна подпись для всех рядов этой стороны.
                options.Publication = PublishPerSourcePosition ' Сначала основной стержень, затем связанные с ним.
                options.Coincidence = KeepPublishedOwner ' В совпавшей позиции сохраняем ранее добавленный стержень.
                options.Comment = rowComment & CStr(rowNumber)
                facade.RebarDependent faces(i).ID & ".Row" & CStr(rowNumber), faces(i).ID & ".Main", faces(i).Diameters(rowNumber), options ' Оба дополнительных ряда связаны с основным.
            End If
        Next rowNumber
        nx = 0#: ny = 0#: annotationSign = 1#: If i > 4 Then annotationSign = -1#
        Select Case base
            Case 1: ny = annotationSign
            Case 2: ny = -annotationSign
            Case 3: nx = -annotationSign
            Case 4: nx = annotationSign
        End Select
        If base <= 2 Then
            facade.RebarRowExtremeLabel faces(i).ID, faces(i).ID & ".Main", 0#, IIf(base = 1, 1#, -1#), nx, ny ' Для горизонтальной стороны подпись берём от крайней линии стержней.
        Else
            facade.RebarLabel faces(i).ID, path.ID, nx, ny ' Общую подпись рядов привязываем к границе их линии.
        End If
NextFace:
    Next i
End Sub

' Возвращает линию центров стержней у выбранной наружной стороны или стороны отверстия.
' Горизонтальные линии включают скругления, вертикальные используют прямую грань.
' При нескруглённом отверстии концы ограничиваются отступами соседних внутренних сторон.
Private Function AxisPath(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As THollowInput, ByRef faces() As THollowFace, _
        ByVal index As Long, ByVal outerPolygon As CPolygonSpec, ByVal openingPolygon As CPolygonSpec) As CGeometryPath
    On Error GoTo InvalidPath
    Dim B As Double, H As Double, radius As Double, x As Double, y As Double, offset As Double, base As Long, shape As CGeometryPath
    Dim first As String, last As String, firstParameter As Double, lastParameter As Double
    Dim polygon As CPolygonSpec
    Set polygon = outerPolygon
    B = sizes.B: H = sizes.H: radius = sizes.R: offset = faces(index).AxisDistance
    base = (index - 1) Mod 4 + 1
    If index > 4 Then
        Set polygon = openingPolygon
        B = sizes.Bo: H = sizes.Ho: radius = sizes.Ro: x = sizes.Xo: y = sizes.Yo: offset = -offset
        If radius <= GEOM_TOLERANCE Then
            facade.Source faces(index).ID, faces(index).Source ' Связываем ошибки этой линии с её исходными полями.
            If base <= 2 Then
                y = y + IIf(base = 1, 1#, -1#) * (H / 2# - offset)
                Set shape = facade.LinePath("Unlimited." & faces(index).ID, x - B / 2#, y, x + B / 2#, y) ' Прямая у отверстия до ограничения соседними отступами.
                ' Ограничиваем линию отступами двух соседних внутренних сторон.
                Set AxisPath = facade.ConstrainLine(faces(index).ID, shape.ID, 1#, 0#, x - B / 2# - faces(7).AxisDistance, 1#, 0#, x + B / 2# + faces(8).AxisDistance)
            Else
                x = x + IIf(base = 3, -1#, 1#) * (B / 2# - offset)
                Set shape = facade.LinePath("Unlimited." & faces(index).ID, x, y - H / 2#, x, y + H / 2#) ' Прямая у отверстия до ограничения соседними отступами.
                ' Ограничиваем линию отступами двух соседних внутренних сторон.
                Set AxisPath = facade.ConstrainLine(faces(index).ID, shape.ID, 0#, 1#, y - H / 2# - faces(6).AxisDistance, 0#, 1#, y + H / 2# + faces(5).AxisDistance)
            End If
            Exit Function
        End If
    End If
    Set shape = facade.OffsetRoundedPolygon("Offset." & faces(index).ID, polygon, offset) ' Смещаем наружный контур внутрь, а контур отверстия — в бетон.
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
            Set AxisPath = facade.Chain(faces(index).ID, shape.ID, first, firstParameter, last, lastParameter, base = 1) ' Горизонтальная сторона вместе с прилегающими скруглениями.
        Case 3: Set AxisPath = facade.Chain(faces(index).ID, shape.ID, "Edge.BL", 0#, "Edge.BL", 1#) ' Левая вертикальная грань.
        Case 4: Set AxisPath = facade.Chain(faces(index).ID, shape.ID, "Edge.TR", 1#, "Edge.TR", 0#, False) ' Правая вертикальная грань с обратным обходом.
    End Select
    Exit Function
InvalidPath:
    SectionRecipeGeometryError "CSectionRecipeHollowRectangle", faces(index).Source
End Function

' Добавляет наружные ширину и высоту сечения, а также ширину и высоту внутреннего отверстия.
Private Sub AddDimensions(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As THollowInput)
    facade.HorizontalDistance "DIM_B", "B", -sizes.B / 2#, sizes.B / 2#, sizes.H / 2#, 1#, "Наружная ширина HollowRectangle." ' Общая ширина сечения B.
    facade.VerticalExtent "DIM_H", "OUTER.Boundary", "H", -sizes.B / 2#, -1#, "Наружная высота HollowRectangle." ' Полная высота сечения H; размер вынесен влево.
    ' Ширина отверстия; размер под сечением.
    facade.HorizontalDistance "DIM_B_OPENING", "B_o", sizes.Xo - sizes.Bo / 2#, sizes.Xo + sizes.Bo / 2#, -sizes.H / 2#, -1#, "Ширина внутреннего отверстия Opening."
    facade.VerticalExtent "DIM_H_OPENING", "OPENING.Boundary", "H_o", sizes.B / 2#, 1#, "Высота внутреннего отверстия Opening." ' Высота отверстия; размер справа от сечения.
End Sub

' Проверяет условие для указанной ячейки. При нарушении сообщает причину и её фактический адрес.
Private Sub Require(ByVal table As CSectionInputTable, ByVal row As Long, ByVal column As Long, ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then SectionRecipeError "CSectionRecipeHollowRectangle", reason, table.SourceReference(row, column)
End Sub
