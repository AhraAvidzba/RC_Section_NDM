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
    OuterSource As String ' Адреса ширины, высоты и радиуса наружного контура.
    OpeningSource As String ' Адреса ширины, высоты и радиуса отверстия.
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
    BaseOffsets(2 To 3) As Double ' Базовые отступы второго и третьего рядов от основного, мм.
    Diameters(2 To 3) As Double ' Диаметры второго и третьего рядов арматуры, мм.
    ExtraSources(2 To 3) As String ' Адреса диаметра, отступа, привязки и положения каждого дополнительного ряда.
    Locations(2 To 3) As String ' Прочитанное расположение каждого дополнительного ряда относительно основного.
    Bindings(2 To 3) As String ' Прочитанный способ привязки дополнительных рядов к основному.
    LabelX As Double ' Компонента X направления выноса подписи.
    LabelY As Double ' Компонента Y направления выноса подписи.
    Source As String ' Адрес исходных данных для пояснения ошибок.
    AxisSource As String ' Адреса as, диаметра и заданного количества для ошибки построения линии арматуры.
End Type

' Читает размеры сечения и отверстия, параметры арматуры вдоль их сторон. Формирует
' описание границ бетона, размещения стержней и аннотаций схемы (размеры сечения
' и отверстия, суммарное количество стержней каждого диаметра по каждой стороне).
Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
    On Error GoTo InvalidGeometry
    Dim reader As CSectionInputReader, facade As CSectionDefinitionBuilder
    Dim sizes As THollowInput, faces(1 To 8) As THollowFace
    Dim outerPolygon As CPolygonSpec, openingPolygon As CPolygonSpec
    Set reader = context.Reader("rngHollowRectangleGeometry") ' Находим поля таблицы по ключам и используем единицы ввода.
    ReadGeometryInputs reader, sizes ' Читаем и проверяем размеры сечения.
    ReadRebarInputs reader, faces ' Читаем армирование наружных и внутренних сторон.
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
Private Sub ReadGeometryInputs(ByVal reader As CSectionInputReader, ByRef sizes As THollowInput)
    sizes.H = reader.PositiveLength("General", "H"): sizes.B = reader.PositiveLength("General", "B")
    sizes.R = reader.NonNegativeLength("General", "R")
    sizes.Ho = reader.PositiveLength("General", "Ho"): sizes.Bo = reader.PositiveLength("General", "Bo")
    sizes.Ro = reader.NonNegativeLength("General", "Ro")
    sizes.Xo = reader.Length("General", "Xo"): sizes.Yo = reader.Length("General", "Yo")
    sizes.OuterSource = reader.SourceReference("General", "H") & "; " & reader.SourceReference("General", "B") & "; " & reader.SourceReference("General", "R")
    sizes.OpeningSource = reader.SourceReference("General", "Ho") & "; " & reader.SourceReference("General", "Bo") & "; " & reader.SourceReference("General", "Ro")
    sizes.GeometrySource = sizes.OuterSource & "; " & sizes.OpeningSource & "; " & reader.SourceReference("General", "Xo") & "; " & reader.SourceReference("General", "Yo")
End Sub

' Заполняет faces параметрами четырёх наружных и четырёх внутренних сторон.
' У выключенных рядов не читает ненужные параметры. Внутренние as читает всегда:
' они ограничивают линии у отверстия даже при отсутствии арматуры на соседней стороне.
Private Sub ReadRebarInputs(ByVal reader As CSectionInputReader, ByRef faces() As THollowFace)
    Dim ids As Variant, captions As Variant, keys As Variant, key As String, i As Long, base As Long, annotationSign As Double
    Dim rowNumber As Long, suffix As String
    keys = Array("Top", "Bottom", "Left", "Right", "OpeningTop", "OpeningBottom", "OpeningLeft", "OpeningRight")
    ids = Array("B.Top", "B.Bottom", "H.Left", "H.Right", "Opening.B.Top", "Opening.B.Bottom", "Opening.H.Left", "Opening.H.Right")
    captions = Array("у верхней наружной стороны", "у нижней наружной стороны", "у левой наружной стороны", "у правой наружной стороны", _
        "у верхней стороны отверстия", "у нижней стороны отверстия", "у левой стороны отверстия", "у правой стороны отверстия")
    For i = 1 To 8
        With faces(i)
            .ID = CStr(ids(i - 1)): .Caption = CStr(captions(i - 1))
            .RowComment = "Пустотелый прямоугольник: арматура " & .Caption & ", ряд "
            key = CStr(keys(i - 1))
            .Source = reader.SourceReference("Main", key, "as")
            If i <= 4 Then .Count = reader.Count("Main", key, "n")
            .Diameter = reader.OptionalDiameter("Main", key, "d")
            .Active = (.Diameter > 0# And (i > 4 Or .Count > 0))
            If i > 4 Then .Active = .Active And faces(i - 4).Count > 0 And faces(i - 4).Diameter > 0#
            If i > 4 Or .Active Then
                .AxisDistance = reader.NonNegativeLength("Main", key, "as")
            End If
            base = (i - 1) Mod 4 + 1
            annotationSign = 1#
            If i > 4 Then annotationSign = -1#
            Select Case base
                Case 1: .LabelY = annotationSign
                Case 2: .LabelY = -annotationSign
                Case 3: .LabelX = -annotationSign
                Case 4: .LabelX = annotationSign
            End Select
            If .Active Then
                .AxisSource = reader.SourceReferences("Main", key, Array("as", "d"))
                If i <= 4 Then .AxisSource = .AxisSource & "; " & reader.SourceReference("Main", key, "n")
                reader.Require "Main", key, "as", .AxisDistance > .Diameter / 2#, "Расстояние до оси должно быть больше половины диаметра."
                .StartSetback = reader.NonNegativeLength("Main", key, "s0")
                .EndSetback = reader.NonNegativeLength("Main", key, "s1")
                .Source = reader.SourceReferences("Main", key, Array("as", "s0", "s1"))
                .Diameters(2) = reader.OptionalDiameter("Extra", key, "d2")
                .Diameters(3) = reader.OptionalDiameter("Extra", key, "d3")
                For rowNumber = 2 To 3
                    If .Diameters(rowNumber) > 0# Then
                        suffix = CStr(rowNumber)
                        .BaseOffsets(rowNumber) = reader.OptionalRowOffset("Extra", key, "a" & suffix)
                        .Locations(rowNumber) = reader.Choice("Extra", key, Array("Stacked", "SideBySide"), "pos" & suffix)
                        .Bindings(rowNumber) = reader.Choice("Extra", key, Array("EachBar", "EverySecondBar"), "bind" & suffix)
                        .ExtraSources(rowNumber) = reader.SourceReferences("Extra", key, Array("d" & suffix, "a" & suffix, "bind" & suffix, "pos" & suffix))
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
    Dim contact As CRebarContactBuilder, extraOptions As CRebarRuleOptions
    Dim normalSign As Double, i As Long, base As Long, path As CGeometryPath, options As CRebarRuleOptions, rowNumber As Long
    For i = 1 To 8
        base = (i - 1) Mod 4 + 1
        If Not faces(i).Active Then GoTo NextFace
        normalSign = 1#: If base = 1 Or base = 3 Then normalSign = -1#
        If i > 4 Then normalSign = -normalSign
        Set path = AxisPath(facade, sizes, faces, i, outerPolygon, openingPolygon) ' Получаем линию центров стержней выбранной стороны с отступом as.
        Set options = New CRebarRuleOptions ' Создаём настройки размещения этого ряда.
        options.StartSetback = faces(i).StartSetback: options.EndSetback = faces(i).EndSetback ' Отступ от начала и, если задан, от конца линии, мм.
        options.AnnotationGroup = faces(i).ID ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой стороны.
        options.Publication = PublishPerSourcePosition ' Нумеруем R1, R2, ...: за каждым основным стержнем идут связанные с ним стержни дополнительных рядов.
        options.Coincidence = KeepPublishedOwner ' При совпадении центров сохраняем стержень, который был добавлен раньше.
        options.IncomingTangentAtJoint = True ' На стыке используем направление предыдущего участка.
        options.SourceReference = faces(i).Source ' Адреса полей для сообщения об ошибке размещения.
        options.Comment = faces(i).RowComment & "1"
        If i <= 4 Then
            If base <= 2 Then options.EndMode = ExcludeBoth ' Не ставим стержни на концах участка после концевых отступов.
            facade.RebarByCount faces(i).ID & ".Main", path.ID, faces(i).Count, faces(i).Diameter, options ' Задаём основной ряд по числу стержней и диаметру.
        Else
            options.SourceNormalSign = -normalSign: options.SkipMissingProjection = True ' Проекция идёт к отверстию; пропускаем лучи без пересечения.
            options.SuccessfulProjectionOrdinal = True: options.NewPublicationRoot = True ' Нумеруем только найденные проекции как самостоятельный ряд.
            facade.RebarProjected faces(i).ID & ".Main", path.ID, faces(base).ID & ".Main", faces(i).Diameter, options ' Проецируем стержни наружного ряда на линию у отверстия.
        End If
        Set extraOptions = New CRebarRuleOptions ' Общие настройки только дополнительных рядов этой стороны.
        extraOptions.AnnotationGroup = faces(i).ID ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой стороны.
        extraOptions.Publication = PublishPerSourcePosition ' Нумеруем R1, R2, ...: за каждым основным стержнем идут связанные с ним стержни дополнительных рядов.
        extraOptions.Coincidence = KeepPublishedOwner ' В совпавшей позиции сохраняем ранее добавленный стержень.
        Set contact = New CRebarContactBuilder
        contact.Initialize faces(i).Diameter, normalSign, 1#, extraOptions
        For rowNumber = 2 To 3
            If faces(i).Diameters(rowNumber) > 0# Then
                Set options = contact.AppendRow(faces(i).Diameters(rowNumber), faces(i).Locations(rowNumber), faces(i).Bindings(rowNumber), faces(i).ExtraSources(rowNumber), faces(i).BaseOffsets(rowNumber))
                options.Comment = faces(i).RowComment & CStr(rowNumber)
                facade.RebarDependent faces(i).ID & ".Row" & CStr(rowNumber), faces(i).ID & ".Main", faces(i).Diameters(rowNumber), options ' Оба дополнительных ряда связаны с основным.
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
    Dim polygon As CPolygonSpec, geometrySource As String
    Set polygon = outerPolygon
    geometrySource = sizes.OuterSource
    B = sizes.B: H = sizes.H: radius = sizes.R: offset = faces(index).AxisDistance
    base = (index - 1) Mod 4 + 1
    If index > 4 Then
        Set polygon = openingPolygon
        geometrySource = sizes.OpeningSource
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
    SectionRecipeGeometryError "CSectionRecipeHollowRectangle", geometrySource & "; " & faces(index).AxisSource, "Арматура " & faces(index).Caption, shape
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
    facade.HorizontalExtent "DIM_B", "OUTER.Boundary", "B", sizes.H / 2#, 1#, "Наружная ширина HollowRectangle." ' Полная ширина по наружному контуру; размер вынесен вверх.
    facade.VerticalExtent "DIM_H", "OUTER.Boundary", "H", -sizes.B / 2#, -1#, "Наружная высота HollowRectangle." ' Полная высота сечения H; размер вынесен влево.
    ' Полная ширина по контуру отверстия; размер под сечением.
    facade.HorizontalExtent "DIM_B_OPENING", "OPENING.Boundary", "B_o", -sizes.H / 2#, -1#, "Ширина внутреннего отверстия Opening."
    facade.VerticalExtent "DIM_H_OPENING", "OPENING.Boundary", "H_o", sizes.B / 2#, 1#, "Высота внутреннего отверстия Opening." ' Высота отверстия; размер справа от сечения.
End Sub
