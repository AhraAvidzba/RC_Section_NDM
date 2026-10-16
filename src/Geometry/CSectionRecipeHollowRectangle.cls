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
    GeometrySource As String ' Адрес исходных размеров для пояснения геометрических ошибок.
End Type

Private Type THollowFace
    Active As Boolean ' Нужно создать основной и включённые дополнительные ряды этой стороны.
    ID As String ' Имя стороны для привязки линий арматуры и подписей.
    RowComment As String ' Общая часть пояснения для основного и дополнительных рядов стороны.
    Caption As String ' Понятное название стороны для сообщений об ошибках.
    AxisDistance As Double ' Расстояние от бетонной границы до центров основного ряда арматуры, мм.
    Diameter As Double ' Диаметр стержней основного ряда, мм.
    Count As Long ' Заданное число наружных стержней; у отверстия число определяет проекция.
    StartSetback As Double ' Отступ начала участка размещения от начала линии, мм.
    EndSetback As Double ' Отступ конца участка размещения от конца линии, мм.
    Diameters(2 To 3) As Double ' Диаметры второго и третьего рядов арматуры, мм.
    ExtraOptions(2 To 3) As CRebarRuleOptions ' Готовые независимые настройки включённых дополнительных рядов.
    NormalSign As Double ' Направление смещения связанных рядов относительно линии центров.
    LabelX As Double ' Компонента X направления выноса подписи.
    LabelY As Double ' Компонента Y направления выноса подписи.
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
    ReadGeometryInputs table, units, sizes ' Читаем и проверяем размеры сечения.
    ReadRebarInputs table, units, faces ' Читаем армирование наружных и внутренних сторон.
    Set outerPolygon = DescribePolygon(sizes.B, sizes.H, 0#, 0#, sizes.R, sizes.GeometrySource) ' Наружный прямоугольник с центром в начале координат.
    Set openingPolygon = DescribePolygon(sizes.Bo, sizes.Ho, sizes.Xo, sizes.Yo, sizes.Ro, sizes.GeometrySource) ' Отверстие с собственными размерами, центром и радиусом.
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "HollowRectangle" ' Создаём сборщик описания сечения.
    AddGeometry facade, outerPolygon, openingPolygon ' Добавляем подготовленные границы бетона.
    AddReinforcement facade, sizes, faces, outerPolygon, openingPolygon ' Создаём линии арматуры и задаём правила размещения стержней.
    AddAnnotations facade, sizes, faces ' Добавляем подписи арматуры, размеры сечения и отверстия.
    Set ISectionRecipe_CreateDefinition = facade.Finish ' Завершаем и возвращаем проверенное описание сечения.
    Exit Function
InvalidGeometry:
    SectionRecipeGeometryError "CSectionRecipeHollowRectangle", sizes.GeometrySource
End Function

' Заполняет sizes наружными размерами, размерами и смещением отверстия в миллиметрах.
' Проверяет размеры и радиусы, сохраняет адреса полей для сообщений об ошибках.
Private Sub ReadGeometryInputs(ByVal table As CSectionInputTable, ByVal units As CUnitSystem, ByRef sizes As THollowInput)
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
End Sub

' Заполняет faces параметрами четырёх наружных и четырёх внутренних сторон.
' У выключенных рядов не читает ненужные параметры. Внутренние as читает всегда:
' они ограничивают линии у отверстия даже при отсутствии арматуры на соседней стороне.
Private Sub ReadRebarInputs(ByVal table As CSectionInputTable, ByVal units As CUnitSystem, ByRef faces() As THollowFace)
    Dim contact As CRebarContactBuilder, defaults As CRebarRuleOptions, rowOptions As CRebarRuleOptions
    Dim locations(2 To 3) As String, bindings(2 To 3) As String, sources(2 To 3) As String
    Dim ids As Variant, captions As Variant, i As Long, row As Long, extra As Long, base As Long, annotationSign As Double
    Dim startColumn As Long, endColumn As Long, rowNumber As Long, column As Long, mainHeader As Long, extraHeader As Long
    ids = Array("B.Top", "B.Bottom", "H.Left", "H.Right", "Opening.B.Top", "Opening.B.Bottom", "Opening.H.Left", "Opening.H.Right")
    captions = Array("B - верхняя", "B - нижняя", "H - левая", "H - правая", "B_o - верхняя", "B_o - нижняя", "H_o - левая", "H_o - правая")
    mainHeader = table.RowByKey("as", 2): extraHeader = table.RowByKey("d2", 2)
    For i = 1 To 8
        With faces(i)
            .ID = CStr(ids(i - 1)): .Caption = CStr(captions(i - 1))
            .RowComment = "HollowRectangle: " & .Caption & IIf(i > 4, " грань Opening", " наружная грань") & ", ряд "
            row = table.RowByKey(.Caption, 1, mainHeader + 1, mainHeader + 8)
            .Source = table.SourceReference(row, 2)
            If i <= 4 Then .Count = table.Count(row, 4)
            .Diameter = table.OptionalDiameter(row, 3, units)
            .Active = (.Diameter > 0# And (i > 4 Or .Count > 0))
            If i > 4 Then .Active = .Active And faces(i - 4).Count > 0 And faces(i - 4).Diameter > 0#
            If i > 4 Or .Active Then
                .AxisDistance = table.NonNegativeLength(row, 2, units)
            End If
            base = (i - 1) Mod 4 + 1
            .NormalSign = 1#: If base = 1 Or base = 3 Then .NormalSign = -1#
            annotationSign = 1#
            If i > 4 Then .NormalSign = -.NormalSign: annotationSign = -1#
            Select Case base
                Case 1: .LabelY = annotationSign
                Case 2: .LabelY = -annotationSign
                Case 3: .LabelX = -annotationSign
                Case 4: .LabelX = annotationSign
            End Select
            If .Active Then
                Require table, row, 2, .AxisDistance > .Diameter / 2#, "Расстояние до оси должно быть больше половины диаметра."
                If startColumn = 0 Then
                    startColumn = table.ColumnByCaption(mainHeader, "StartSetback")
                    endColumn = table.ColumnByCaption(mainHeader, "EndSetback")
                End If
                .StartSetback = table.NonNegativeLength(row, startColumn, units)
                .EndSetback = table.NonNegativeLength(row, endColumn, units)
                .Source = .Source & "; " & table.SourceReference(row, startColumn) & "; " & table.SourceReference(row, endColumn)
                extra = table.RowByKey(.Caption, 1, extraHeader + 1, extraHeader + 8)
                .Diameters(2) = table.OptionalDiameter(extra, 2, units)
                .Diameters(3) = table.OptionalDiameter(extra, 5, units)
                For rowNumber = 2 To 3
                    If .Diameters(rowNumber) > 0# Then
                        column = 2 + (rowNumber - 2) * 3 ' В таблице на каждый ряд отведены диаметр, расположение и привязка.
                        locations(rowNumber) = table.Choice(extra, column + 1, Array("Stacked", "SideBySide"))
                        bindings(rowNumber) = table.Choice(extra, column + 2, Array("EachBar", "EverySecondBar"))
                    End If
                Next rowNumber
                Set defaults = New CRebarRuleOptions ' Общие настройки только дополнительных рядов этой стороны.
                defaults.AnnotationGroup = .ID ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой стороны.
                defaults.Publication = PublishPerSourcePosition ' Сначала основной стержень, затем связанные с ним.
                defaults.Coincidence = KeepPublishedOwner ' В совпавшей позиции сохраняем ранее добавленный стержень.
                Set contact = New CRebarContactBuilder
                contact.Initialize .Diameter, .NormalSign, 1#, defaults
                For rowNumber = 2 To 3
                    If .Diameters(rowNumber) > 0# Then
                        Set rowOptions = contact.AppendRow(.Diameters(rowNumber), locations(rowNumber), bindings(rowNumber), .Source)
                        rowOptions.Comment = .RowComment & CStr(rowNumber)
                        Set .ExtraOptions(rowNumber) = rowOptions
                    End If
                Next rowNumber
            End If
        End With
    Next i
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

' Добавляет в описание сечения наружную границу и отверстие по двум подготовленным многоугольникам.
Private Sub AddGeometry(ByVal facade As CSectionDefinitionBuilder, ByVal outerPolygon As CPolygonSpec, ByVal openingPolygon As CPolygonSpec)
    AddBoundary facade, "OUTER", "OUTER.Boundary", outerPolygon, False
    AddBoundary facade, "OPENING", "OPENING.Boundary", openingPolygon, True
End Sub

' Создаёт готовый замкнутый контур по вершинам и радиусам без повторного обхода.
' Назначает контур наружной границей или отверстием внутри OUTER, в зависимости от opening.
Private Sub AddBoundary(ByVal facade As CSectionDefinitionBuilder, ByVal loopID As String, ByVal pathID As String, ByVal polygon As CPolygonSpec, ByVal opening As Boolean)
    Dim boundary As CGeometryPath, part As String
    Set boundary = facade.RoundedPolygon(pathID, polygon) ' Создаём контур, который сразу используется как наружная граница или отверстие.
    If opening Then part = "внутреннего контура Opening" Else part = "наружного контура"
    If opening Then
        facade.Opening loopID, boundary.ID, "OUTER", "Прямой участок " & part & " HollowRectangle.", _
            arcComment:="Дуга " & part & " HollowRectangle.", namingPolicy:=SequentialByKind, namingPrefix:="CONTOUR_" & loopID ' Назначаем отверстие внутри наружной границы OUTER.
    Else
        facade.Outer loopID, boundary.ID, "Прямой участок " & part & " HollowRectangle.", _
            arcComment:="Дуга " & part & " HollowRectangle.", namingPolicy:=SequentialByKind, namingPrefix:="CONTOUR_" & loopID ' Назначаем эту замкнутую линию наружной границей бетона.
    End If
End Sub

' Задаёт арматуру вдоль наружных сторон по числу стержней, а вдоль отверстия — проекцией наружной арматуры.
' Задаёт связанные ряды по подготовленным параметрам каждой стороны.
Private Sub AddReinforcement(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As THollowInput, ByRef faces() As THollowFace, _
        ByVal outerPolygon As CPolygonSpec, ByVal openingPolygon As CPolygonSpec)
    Dim i As Long, base As Long, path As CGeometryPath, options As CRebarRuleOptions, rowNumber As Long
    For i = 1 To 8
        base = (i - 1) Mod 4 + 1
        If Not faces(i).Active Then GoTo NextFace
        Set path = AxisPath(facade, sizes, faces, i, outerPolygon, openingPolygon) ' Получаем линию центров стержней выбранной стороны с отступом as.
        Set options = New CRebarRuleOptions ' Создаём настройки размещения этого ряда.
        options.StartSetback = faces(i).StartSetback: options.EndSetback = faces(i).EndSetback ' Отступ от начала и, если задан, от конца линии, мм.
        options.AnnotationGroup = faces(i).ID ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой стороны.
        options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
        options.Coincidence = KeepPublishedOwner ' При совпадении центров сохраняем стержень, который был добавлен раньше.
        options.IncomingTangentAtJoint = True ' На стыке используем направление предыдущего участка.
        options.SourceReference = faces(i).Source ' Адреса полей для сообщения об ошибке размещения.
        options.Comment = faces(i).RowComment & "1"
        If i <= 4 Then
            If base <= 2 Then options.EndMode = ExcludeBoth ' Не ставим стержни на концах участка после концевых отступов.
            facade.RebarByCount faces(i).ID & ".Main", path.ID, faces(i).Count, faces(i).Diameter, options ' Задаём основной ряд по числу стержней и диаметру.
        Else
            options.SourceNormalSign = -faces(i).NormalSign: options.SkipMissingProjection = True ' Проекция идёт к отверстию; пропускаем лучи без пересечения.
            options.SuccessfulProjectionOrdinal = True: options.NewPublicationRoot = True ' Нумеруем только найденные проекции как самостоятельный ряд.
            facade.RebarProjected faces(i).ID & ".Main", path.ID, faces(base).ID & ".Main", faces(i).Diameter, options ' Проецируем стержни наружного ряда на линию у отверстия.
        End If
        For rowNumber = 2 To 3
            If faces(i).Diameters(rowNumber) > 0# Then
                facade.RebarDependent faces(i).ID & ".Row" & CStr(rowNumber), faces(i).ID & ".Main", faces(i).Diameters(rowNumber), faces(i).ExtraOptions(rowNumber) ' Оба дополнительных ряда связаны с основным.
            End If
        Next rowNumber
NextFace:
    Next i
End Sub

' Возвращает линию центров стержней у выбранной наружной стороны или стороны отверстия.
' Горизонтальные линии идут слева направо и включают скругления, если они есть; вертикальные — снизу вверх.
' При нескруглённом отверстии концы задаются отступами соседних внутренних сторон.
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
                Set shape = facade.LinePath("Unlimited." & faces(index).ID, x - B / 2#, y, x + B / 2#, y) ' Исходная прямая вдоль стороны отверстия.
                ' Задаём концы линии по отступам соседних внутренних сторон; при необходимости продлеваем её.
                Set AxisPath = facade.ConstrainLine(faces(index).ID, shape.ID, 1#, 0#, x - B / 2# - faces(7).AxisDistance, 1#, 0#, x + B / 2# + faces(8).AxisDistance)
            Else
                x = x + IIf(base = 3, -1#, 1#) * (B / 2# - offset)
                Set shape = facade.LinePath("Unlimited." & faces(index).ID, x, y - H / 2#, x, y + H / 2#) ' Исходная прямая вдоль стороны отверстия.
                ' Задаём концы линии по отступам соседних внутренних сторон; при необходимости продлеваем её.
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
            Set AxisPath = facade.Chain(faces(index).ID, shape.ID, first, firstParameter, last, lastParameter, base = 1) ' Линия центров вдоль горизонтальной стороны со скруглениями, слева направо.
        Case 3
            Set AxisPath = facade.Chain(faces(index).ID, shape.ID, "Edge.BL", 0#, "Edge.BL", 1#) ' Линия центров у левой грани, снизу вверх.
        Case 4
            Set AxisPath = facade.Chain(faces(index).ID, shape.ID, "Edge.TR", 1#, "Edge.TR", 0#, False) ' Линия центров у правой грани, снизу вверх; обход контура обратный.
    End Select
    Exit Function
InvalidPath:
    SectionRecipeGeometryError "CSectionRecipeHollowRectangle", faces(index).Source
End Function

' Добавляет подписи арматуры, наружные ширину и высоту сечения, а также размеры внутреннего отверстия.
Private Sub AddAnnotations(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As THollowInput, ByRef faces() As THollowFace)
    Dim i As Long, base As Long
    For i = 1 To 8
        If faces(i).Active Then
            base = (i - 1) Mod 4 + 1
            If base <= 2 Then
                facade.RebarRowExtremeLabel faces(i).ID, faces(i).ID & ".Main", 0#, IIf(base = 1, 1#, -1#), faces(i).LabelX, faces(i).LabelY ' Для горизонтальной стороны подпись берём от крайней линии стержней.
            Else
                facade.RebarLabel faces(i).ID, faces(i).ID, faces(i).LabelX, faces(i).LabelY ' Общую подпись рядов привязываем к границе их линии.
            End If
        End If
    Next i
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
