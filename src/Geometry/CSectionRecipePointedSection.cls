VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionRecipePointedSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Клиновидное бетонное сечение
' ==========================================================================
' Класс описывает сечение с прямой горизонтальной базой и двумя боковыми сторонами.
' Каждая сторона может быть обычной или иметь выступ; радиусы задаются независимо.
' Линии центров арматуры получаются смещением того же многоугольника, что задаёт бетон.
'
' Рецепт описывает границу бетона, расположение арматуры вдоль сторон и аннотации
' схемы (размеры сечения, количество и диаметры стержней). Общие геометрические
' классы строят дуги и смещения, а CSectionBuilder
' создаёт стержни и бетонную сетку.

Implements ISectionRecipe

Private Type TRoundedInput
    B As Double ' Ширина прямоугольной основы без боковых выступов, мм.
    H As Double ' Высота наружного контура, мм.
    Widths(1 To 2) As Double ' Выносы вершин левого и правого выступов до скругления, мм.
    R1(1 To 2) As Double ' Радиусы верхнего и нижнего углов каждой боковой стороны, мм.
    R2(1 To 2) As Double ' Радиусы скругления вершин левого и правого выступов, мм.
    GeometrySource As String ' Адрес исходных размеров для пояснения геометрических ошибок.
End Type

Private Type TRoundedFace
    ID As String ' Имя стороны для привязки линий арматуры и подписей.
    Caption As String ' Единственное название стороны в утверждённой таблице на листе «Сечения».
    Count As Long ' Число стержней основного ряда.
    Diameter As Double ' Диаметр стержней основного ряда, мм.
    AxisDistance As Double ' Расстояние от края бетона до осей основного ряда, мм.
    EndMode As ERebarEndMode ' Явный способ расстановки основного ряда на подготовленном пути.
    StartSetback As Double ' Отступ начала участка размещения от начала линии, мм.
    EndSetback As Double ' Отступ конца участка размещения от конца линии, мм.
    LabelX As Double ' Компонента X направления выноса подписи.
    LabelY As Double ' Компонента Y направления выноса подписи.
    BaseOffsets(2 To 3) As Double ' Базовые отступы второго и третьего рядов от основного, мм.
    Diameters(2 To 3) As Double ' Диаметры второго и третьего рядов арматуры, мм.
    Locations(2 To 3) As String ' Прочитанное расположение каждого дополнительного ряда относительно основного.
    Bindings(2 To 3) As String ' Прочитанный способ привязки дополнительных рядов к основному.
    ExtraSources(2 To 3) As String ' Адреса исходных полей дополнительных рядов для сообщения об ошибке.
    Source As String ' Адреса отступов основного ряда для сообщения об ошибке размещения.
    AxisSource As String ' Адреса as, диаметра и количества для пояснения ошибки построения линии арматуры.
End Type

' Читает форму клиновидного сечения и параметры армирования. Формирует
' описание границы бетона, размещения стержней вдоль прямых и скруглённых сторон,
' а также аннотаций схемы (размеры сечения, количество и диаметры стержней).
Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
    Dim reader As CSectionInputReader, facade As CSectionDefinitionBuilder
    Dim sizes As TRoundedInput, faces(1 To 4) As TRoundedFace, outer As CGeometryPath, polygon As CPolygonSpec
    Set reader = context.Reader("rngPointedSectionGeometry") ' Находим поля таблицы по ключам и используем единицы ввода.
    ReadGeometryInputs reader, sizes ' Читаем и проверяем размеры сечения.
    ReadRebarInputs reader, faces ' Читаем параметры сторон, отступы и привязки включённых рядов.
    Set polygon = DescribePolygon(sizes) ' Общее описание вершин для бетона и всех смещённых линий.
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "PointedSection" ' Создаём сборщик описания сечения.
    Set outer = AddGeometry(facade, sizes, polygon) ' Получаем наружный контур по подготовленному многоугольнику.
    AddReinforcement facade, faces, polygon ' Создаём линии арматуры и задаём правила размещения стержней.
    AddAnnotations facade, sizes, faces, outer ' Добавляем подписи арматуры, габариты и размеры сечения.
    Set ISectionRecipe_CreateDefinition = facade.Finish ' Завершаем и возвращаем проверенное описание сечения.
End Function

' Заполняет sizes размерами в миллиметрах и проверяет параметры обеих сторон.
' W = 0 задаёт сторону без выступа; её радиус R2 не используется. Сохраняет адреса активных полей.
Private Sub ReadGeometryInputs(ByVal reader As CSectionInputReader, ByRef sizes As TRoundedInput)
    Dim side As Long, key As String
    sizes.B = reader.PositiveLength("General", "B"): sizes.H = reader.PositiveLength("General", "H")
    sizes.GeometrySource = reader.SourceReference("General", "B") & "; " & reader.SourceReference("General", "H")
    For side = 1 To 2
        If side = 1 Then key = "Left." Else key = "Right."
        sizes.Widths(side) = reader.NonNegativeLength("General", key & "W")
        sizes.R1(side) = reader.NonNegativeLength("General", key & "R1")
        sizes.GeometrySource = sizes.GeometrySource & "; " & reader.SourceReference("General", key & "W") & "; " & reader.SourceReference("General", key & "R1")
        If sizes.Widths(side) > 0# Then
            sizes.R2(side) = reader.NonNegativeLength("General", key & "R2")
            sizes.GeometrySource = sizes.GeometrySource & "; " & reader.SourceReference("General", key & "R2")
        End If
    Next side
End Sub

' Читает четыре стороны по коротким ключам таблицы сечения и готовит параметры армирования.
' Выключенная сторона не требует as, отступов и дополнительных рядов.
Private Sub ReadRebarInputs(ByVal reader As CSectionInputReader, ByRef faces() As TRoundedFace)
    Dim ids As Variant, captions As Variant, keys As Variant, key As String, i As Long, rowNumber As Long, suffix As String
    keys = Array("Top", "Bottom", "Left", "Right")
    ids = Array("B.Top", "B.Bottom", "H.Left", "H.Right")
    captions = Array("B - верхняя", "B - нижняя", "H - левая", "H - правая")
    For i = 1 To 4
        With faces(i)
            .ID = CStr(ids(i - 1)): .Caption = CStr(captions(i - 1)) ' Номер структуры начинается с 1, массивы соответствий Array — с 0; это не номер ячейки.
            key = CStr(keys(i - 1)) ' Короткий ключ строки Main/Extra выбирается из фиксированного списка сторон.
            .Count = reader.Count("Main", key, "n"): .Diameter = reader.OptionalDiameter("Main", key, "d")
            .AxisSource = reader.SourceReference("Main", key, "as")
            If .Count = 0 Or .Diameter = 0# Then
                .AxisDistance = reader.OptionalRowOffset("Main", key, "as")
                GoTo NextFace
            End If
            .AxisDistance = reader.Length("Main", key, "as")
            reader.Require "Main", key, "as", .AxisDistance > .Diameter / 2#, "Расстояние до оси должно быть больше половины диаметра стержня."
            .AxisSource = reader.SourceReferences("Main", key, Array("as", "d", "n"))
            Select Case .ID
                Case "B.Top": .LabelY = 1#
                Case "B.Bottom": .LabelY = -1#
                Case "H.Left": .LabelX = -1#
                Case "H.Right": .LabelX = 1#
            End Select
            .StartSetback = reader.OptionalRowOffset("Main", key, "s0")
            .EndSetback = reader.OptionalRowOffset("Main", key, "s1")
            .EndMode = reader.RebarEndMode("Main", key)
            .Source = reader.SourceReferences("Main", key, Array("as", "s0", "s1", "endMode"))
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
        End With
NextFace:
    Next i
End Sub

' Возвращает завершённое описание вершин и радиусов: четыре основных угла и вершины
' включённых выступов. Этот же объект используется для контура бетона и линий арматуры.
Private Function DescribePolygon(ByRef sizes As TRoundedInput) As CPolygonSpec
    Dim polygon As CPolygonSpec
    Set polygon = New CPolygonSpec
    polygon.AddVertex "TL", -sizes.B / 2#, sizes.H / 2#, sizes.R1(1), sizes.GeometrySource, outgoingEdgeID:="Edge.TL" ' Верхний левый угол; Edge.TL — верхняя грань, размеры в мм.
    polygon.AddVertex "TR", sizes.B / 2#, sizes.H / 2#, sizes.R1(2), sizes.GeometrySource ' Верхний правый угол до скругления; координаты и радиус в мм.
    If sizes.Widths(2) > 0# Then polygon.AddVertex "RightNose", sizes.B / 2# + sizes.Widths(2), 0#, sizes.R2(2), sizes.GeometrySource ' Вершина правого выступа до скругления; координаты и радиус в мм.
    polygon.AddVertex "BR", sizes.B / 2#, -sizes.H / 2#, sizes.R1(2), sizes.GeometrySource, outgoingEdgeID:="Edge.BR" ' Нижний правый угол; Edge.BR — нижняя грань, размеры в мм.
    polygon.AddVertex "BL", -sizes.B / 2#, -sizes.H / 2#, sizes.R1(1), sizes.GeometrySource, outgoingEdgeID:="Edge.BL" ' Нижний левый угол; Edge.BL — участок к следующей вершине, размеры в мм.
    If sizes.Widths(1) > 0# Then polygon.AddVertex "LeftNose", -sizes.B / 2# - sizes.Widths(1), 0#, sizes.R2(1), sizes.GeometrySource, outgoingEdgeID:="Edge.LeftNose" ' Вершина левого выступа до скругления; координаты и радиус в мм.
    polygon.Seal ' Проверяем вершины и завершаем их описание.
    Set DescribePolygon = polygon ' Возвращаем подготовленный многоугольник.
End Function

' Создаёт наружный контур по заданным вершинам и радиусам и сразу назначает его границей бетона.
' Ошибку связывает с исходными размерами.
Private Function AddGeometry(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRoundedInput, ByVal polygon As CPolygonSpec) As CGeometryPath
    On Error GoTo InvalidGeometry
    Dim outer As CGeometryPath
    Set outer = facade.RoundedPolygon("Boundary", polygon) ' Создаём готовый замкнутый контур по исходным вершинам и радиусам.
    facade.Outer "OUTER", outer.ID, "Прямой участок контура клиновидного сечения.", arcComment:="Дуга контура клиновидного сечения.", namingPolicy:=SequentialByKind ' Назначаем эту замкнутую линию наружной границей бетона.
    Set AddGeometry = outer
    Exit Function
InvalidGeometry:
    SectionRecipeGeometryError "CSectionRecipePointedSection", sizes.GeometrySource, "Клиновидное сечение: наружный контур."
End Function

' Создаёт линии арматуры и задаёт ряды вдоль четырёх сторон по подготовленным параметрам.
' Порядок сторон сохраняет принадлежность общего углового стержня первому из заданных рядов.
Private Sub AddReinforcement(ByVal facade As CSectionDefinitionBuilder, ByRef faces() As TRoundedFace, ByVal polygon As CPolygonSpec)
    Dim contact As CRebarContactBuilder, extraOptions As CRebarRuleOptions
    Dim normalSign As Double, i As Long, rowNumber As Long, path As CGeometryPath, options As CRebarRuleOptions
    Dim contour As CGeometryPath, groups As Variant, leftGroup As String, rightGroup As String
    leftGroup = "Corner.TL|Edge.BL|Corner.BL": rightGroup = "Corner.TR|Edge.TR|Corner.BR"
    For i = 1 To polygon.Count
        If polygon.VertexID(i) = "LeftNose" Then leftGroup = leftGroup & "|Edge.LeftNose|Corner.LeftNose"
        If polygon.VertexID(i) = "RightNose" Then rightGroup = rightGroup & "|Edge.RightNose|Corner.RightNose"
    Next i
    groups = Array("Edge.TL", "Edge.BR", leftGroup, rightGroup)
    For i = 1 To 4
        If faces(i).Count = 0 Or faces(i).Diameter = 0# Then facade.RebarContourBaseline "OUTER", "Boundary", Split(CStr(groups(i - 1)), "|"), faces(i).AxisDistance, faces(i).AxisSource
    Next i
    Set contour = facade.OffsetContour("Rebar.Boundary", "Boundary", groups, _
        Array(faces(1).AxisDistance, faces(2).AxisDistance, faces(3).AxisDistance, faces(4).AxisDistance), _
        Array(faces(1).AxisSource, faces(2).AxisSource, faces(3).AxisSource, faces(4).AxisSource))
    For i = 1 To 4
        With faces(i)
            If .Count = 0 Or .Diameter = 0# Then GoTo NextFace
            Set path = AxisPath(facade, polygon, contour, .ID, .AxisSource, "Арматура у стороны «" & .Caption & "»") ' Выбираем сторону из общего контура с индивидуальными as.
            Set options = New CRebarRuleOptions ' Создаём настройки размещения этого ряда.
            options.EndMode = .EndMode ' Применяем выбранное правило после концевых отступов.
            options.StartSetback = .StartSetback: options.EndSetback = .EndSetback ' Отступы от начала и конца линии, мм.
            options.AnnotationGroup = .ID ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой стороны.
            options.Publication = PublishPerSourcePosition ' Нумеруем R1, R2, ...: за каждым основным стержнем идут связанные с ним стержни дополнительных рядов.
            options.Coincidence = KeepPublishedOwner ' Первый добавленный физический стержень сохраняет свои диаметр и материал.
            options.WarnOnCoincidence = True ' Каждый пропуск совпавшего Main фиксируется предупреждением, включая одинаковые диаметры.
            options.IncomingTangentAtJoint = True ' На стыке используем направление предыдущего участка.
            options.Comment = "Клиновидное сечение: " & .Caption & " грань, ряд 1"
            options.SourceReference = .Source ' Адреса полей для сообщения об ошибке размещения.
            facade.RebarByCount .ID & ".Main", path.ID, .Count, .Diameter, options ' Задаём основной ряд по числу стержней и диаметру.
            Select Case .ID
                Case "B.Top", "H.Right": normalSign = -1#
                Case "B.Bottom", "H.Left": normalSign = 1#
            End Select
            facade.RebarContour .ID & ".Main", "OUTER", "Boundary", .AxisDistance, .ID & ".Row2", .ID & ".Row3", sourceSegments:=Split(CStr(groups(i - 1)), "|"), normalSign:=normalSign ' Явные исходные участки включают дуги и не зависят от соединительных отрезков общего контура.
            Set extraOptions = New CRebarRuleOptions ' Общие настройки только дополнительных рядов этой стороны.
            extraOptions.AnnotationGroup = .ID ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой стороны.
            extraOptions.Publication = PublishPerSourcePosition ' Нумеруем R1, R2, ...: за каждым основным стержнем идут связанные с ним стержни дополнительных рядов.
            extraOptions.Coincidence = KeepPublishedOwner ' В совпавшей позиции сохраняем ранее добавленный стержень.
            Set contact = New CRebarContactBuilder
            contact.Initialize .Diameter, normalSign, 1#, extraOptions
            For rowNumber = 2 To 3
                If .Diameters(rowNumber) > 0# Then
                    Set options = contact.AppendRow(.Diameters(rowNumber), .Locations(rowNumber), .Bindings(rowNumber), .ExtraSources(rowNumber), .BaseOffsets(rowNumber))
                    options.Comment = "Клиновидное сечение: " & .Caption & " грань, ряд " & CStr(rowNumber)
                    facade.RebarDependent .ID & ".Row" & CStr(rowNumber), .ID & ".Main", .Diameters(rowNumber), options ' Оба дополнительных ряда связаны с основным.
                End If
            Next rowNumber
        End With
NextFace:
    Next i
End Sub

' Возвращает линию центров стержней выбранной стороны после смещения контура на as.
' Горизонтальные линии направлены слева направо; боковые включают скругления и идут сверху вниз.
' Параметры 0 и 1 задают концы сегмента, а forward выбирает обход исходного контура между ними.
Private Function AxisPath(ByVal facade As CSectionDefinitionBuilder, ByVal polygon As CPolygonSpec, _
        ByVal offset As CGeometryPath, ByVal groupID As String, ByVal sourceReference As String, ByVal rebarCaption As String) As CGeometryPath
    On Error GoTo InvalidPath
    Select Case groupID
        Case "B.Top"
            Set AxisPath = facade.Chain(groupID, offset.ID, "Edge.TL", 0#, "Edge.TL", 1#) ' Линия центров верхнего ряда, слева направо.
        Case "B.Bottom"
            Set AxisPath = facade.Chain(groupID, offset.ID, "Edge.BR", 1#, "Edge.BR", 0#, False) ' Линия центров нижнего ряда, слева направо.
        Case "H.Left"
            Set AxisPath = facade.Chain(groupID, offset.ID, "Edge.TL", 0#, "Edge.BR", 1#, False) ' Линия центров вдоль боковой стороны со скруглениями, сверху вниз.
        Case "H.Right"
            Set AxisPath = facade.Chain(groupID, offset.ID, "Edge.TL", 1#, "Edge.BR", 0#) ' Линия центров вдоль боковой стороны со скруглениями, сверху вниз.
    End Select
    Exit Function
InvalidPath:
    SectionRecipeGeometryError "CSectionRecipePointedSection", polygon.SourceReference(1) & "; " & sourceReference, rebarCaption, offset
End Function

' Добавляет подписи арматуры, ширину основы B, полную высоту H и размеры включённых боковых выступов.
Private Sub AddAnnotations(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRoundedInput, ByRef faces() As TRoundedFace, ByVal outer As CGeometryPath)
    Dim minX As Double, maxX As Double, i As Long
    For i = 1 To 4
        With faces(i)
            If .Count > 0 And .Diameter > 0# Then
                If i >= 3 Then
                    facade.RebarRowExtentLabel .ID, .ID & ".Main", .LabelX, .LabelY, True ' Привязываем подпись к крайним стержням основного ряда.
                Else
                    facade.RebarLabel .ID, .ID, .LabelX, .LabelY ' Общую подпись рядов привязываем к границе их линии.
                End If
            End If
        End With
    Next i
    outer.ProjectionBounds 1#, 0#, minX, maxX
    facade.HorizontalDistance "DIM_B", "B", -sizes.B / 2#, sizes.B / 2#, sizes.H / 2#, 1#, "Размер горизонтальной базы PointedSection между левым и правым торцом." ' Ширина основы B до скругления углов; размер сверху.
    facade.VerticalExtent "DIM_H", outer.ID, "H", minX, -1#, "Полная высота PointedSection." ' Полная высота сечения H; размер вынесен влево.
    If sizes.Widths(1) > 0# And -sizes.B / 2# - minX > GEOM_TOLERANCE Then facade.HorizontalDistance "DIM_W_LEFT", "W'", minX, -sizes.B / 2#, sizes.H / 2#, 1#, "Фактический вынос левого края бетона относительно левого торца базы B.", 1# ' Фактическая ширина левого выступа; подпись округляется до 1 мм.
    If sizes.Widths(2) > 0# And maxX - sizes.B / 2# > GEOM_TOLERANCE Then facade.HorizontalDistance "DIM_W_RIGHT", "W'", sizes.B / 2#, maxX, sizes.H / 2#, 1#, "Фактический вынос правого края бетона относительно правого торца базы B.", 1# ' Фактическая ширина правого выступа; подпись округляется до 1 мм.
End Sub
