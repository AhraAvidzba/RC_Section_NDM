VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionRecipeRoundedRectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Бетонное сечение со скруглёнными боковыми сторонами
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
    SideTypes(1 To 2) As String ' Вид левой и правой боковых сторон.
    Widths(1 To 2) As Double ' Выносы вершин левого и правого выступов до скругления, мм.
    R1(1 To 2) As Double ' Радиусы верхнего и нижнего углов каждой боковой стороны, мм.
    R2(1 To 2) As Double ' Радиусы скругления вершин левого и правого выступов, мм.
    GeometrySource As String ' Адрес исходных размеров для пояснения геометрических ошибок.
End Type

Private Type TRoundedFace
    ID As String ' Имя стороны для привязки линий арматуры и подписей.
    Caption As String ' Единственное название стороны в утверждённой таблице Config.
    Count As Long ' Число стержней основного ряда.
    Diameter As Double ' Диаметр стержней основного ряда, мм.
    AxisDistance As Double ' Расстояние от края бетона до осей основного ряда, мм.
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
    AxisSource As String ' Адрес расстояния до оси для пояснения ошибки построения линии.
End Type

' Читает форму сечения со скруглёнными сторонами и параметры армирования. Формирует
' описание границы бетона, размещения стержней вдоль прямых и скруглённых сторон,
' а также аннотаций схемы (размеры сечения, количество и диаметры стержней).
Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
    Dim table As CSectionInputTable, units As CUnitSystem, facade As CSectionDefinitionBuilder
    Dim sizes As TRoundedInput, faces(1 To 4) As TRoundedFace, outer As CGeometryPath, polygon As CPolygonSpec
    Set table = context.Table("rngRoundedRectangleGeometry"): Set units = context.Units ' Берём копию входной таблицы и единицы измерения.
    ReadGeometryInputs table, units, sizes ' Читаем и проверяем размеры сечения.
    ReadRebarInputs table, units, faces ' Читаем параметры сторон, отступы и привязки включённых рядов.
    Set polygon = DescribePolygon(sizes) ' Общее описание вершин для бетона и всех смещённых линий.
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "RoundedRectangle" ' Создаём сборщик описания сечения.
    Set outer = AddGeometry(facade, sizes, polygon) ' Получаем наружный контур по подготовленному многоугольнику.
    AddReinforcement facade, faces, polygon ' Создаём линии арматуры и задаём правила размещения стержней.
    AddAnnotations facade, sizes, faces, outer ' Добавляем подписи арматуры, габариты и размеры сечения.
    Set ISectionRecipe_CreateDefinition = facade.Finish ' Завершаем и возвращаем проверенное описание сечения.
End Function

' Заполняет sizes размерами в миллиметрах и проверяет параметры обеих сторон.
' Для обычной стороны не читает ширину выступа и его радиус; сохраняет адреса активных полей.
Private Sub ReadGeometryInputs(ByVal table As CSectionInputTable, ByVal units As CUnitSystem, ByRef sizes As TRoundedInput)
    Dim row As Long, sides As Long, side As Long, column As Long
    row = table.RowByKey("B", 2) + 1
    sizes.B = table.PositiveLength(row, 2, units): sizes.H = table.PositiveLength(row, 3, units)
    sizes.GeometrySource = table.SourceReference(row, 2) & "; " & table.SourceReference(row, 3)
    sides = table.RowByKey("слева", 2)
    For side = 1 To 2
        column = side + 1
        sizes.SideTypes(side) = table.Choice(sides + 1, column, Array("Simple", "Tapered"))
        sizes.R1(side) = table.NonNegativeLength(sides + 3, column, units)
        sizes.GeometrySource = sizes.GeometrySource & "; " & table.SourceReference(sides + 3, column)
        If sizes.SideTypes(side) = "Tapered" Then
            sizes.Widths(side) = table.Length(sides + 2, column, units)
            sizes.R2(side) = table.NonNegativeLength(sides + 4, column, units)
            sizes.GeometrySource = sizes.GeometrySource & "; " & table.SourceReference(sides + 2, column) & "; " & table.SourceReference(sides + 4, column)
            Require table, sides + 2, column, sizes.Widths(side) > GEOM_TOLERANCE, "Для Tapered размер W должен быть положительным."
        End If
    Next side
End Sub

' Читает четыре стороны по утверждённым названиям Config и готовит параметры армирования.
' Выключенная сторона не требует as, отступов и дополнительных рядов.
Private Sub ReadRebarInputs(ByVal table As CSectionInputTable, ByVal units As CUnitSystem, ByRef faces() As TRoundedFace)
    Dim ids As Variant, captions As Variant, i As Long, row As Long, extra As Long, mainHeader As Long, extraHeader As Long
    Dim startColumn As Long, endColumn As Long, rowNumber As Long, column As Long
    ids = Array("B.Top", "B.Bottom", "H.Left", "H.Right")
    captions = Array("B - верхняя", "B - нижняя", "H - левая", "H - правая")
    mainHeader = table.RowByKey("as", 2): extraHeader = table.AdditionalRowsHeader
    For i = 1 To 4
        With faces(i)
            .ID = CStr(ids(i - 1)): .Caption = CStr(captions(i - 1))
            row = table.RowByKey(.Caption, 1, mainHeader + 1, mainHeader + 4)
            .Count = table.Count(row, 4): .Diameter = table.OptionalDiameter(row, 3, units)
            If .Count = 0 Or .Diameter = 0# Then GoTo NextFace
            .AxisDistance = table.Length(row, 2, units)
            Require table, row, 2, .AxisDistance > .Diameter / 2#, "Расстояние до оси должно быть больше половины диаметра стержня."
            .AxisSource = table.SourceReference(row, 2)
            Select Case .ID
                Case "B.Top": .LabelY = 1#
                Case "B.Bottom": .LabelY = -1#
                Case "H.Left": .LabelX = -1#
                Case "H.Right": .LabelX = 1#
            End Select
            If startColumn = 0 Then
                startColumn = table.ColumnByCaption(mainHeader, "StartSetback")
                endColumn = table.ColumnByCaption(mainHeader, "EndSetback")
            End If
            .StartSetback = table.NonNegativeLength(row, startColumn, units)
            .EndSetback = table.NonNegativeLength(row, endColumn, units)
            .Source = .AxisSource & "; " & table.SourceReference(row, startColumn) & "; " & table.SourceReference(row, endColumn)
            extra = table.RowByKey(.Caption, 1, extraHeader + 1, extraHeader + 4)
            .Diameters(2) = table.OptionalDiameter(extra, 2, units)
            .Diameters(3) = table.OptionalDiameter(extra, 6, units)
            For rowNumber = 2 To 3
                If .Diameters(rowNumber) > 0# Then
                    column = 2 + (rowNumber - 2) * 4 ' На ряд отведены диаметр, отступ, привязка и положение.
                    .BaseOffsets(rowNumber) = table.OptionalRowOffset(extra, column + 1, units)
                    .Locations(rowNumber) = table.Choice(extra, column + 3, Array("Stacked", "SideBySide"))
                    .Bindings(rowNumber) = table.Choice(extra, column + 2, Array("EachBar", "EverySecondBar"))
                    .ExtraSources(rowNumber) = table.SourceReference(extra, column) & "; " & table.SourceReference(extra, column + 1) & "; " & table.SourceReference(extra, column + 2) & "; " & table.SourceReference(extra, column + 3)
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
    ' Вершина правого выступа до скругления; координаты и радиус в мм.
    If sizes.SideTypes(2) = "Tapered" Then polygon.AddVertex "RightNose", sizes.B / 2# + sizes.Widths(2), 0#, sizes.R2(2), sizes.GeometrySource
    polygon.AddVertex "BR", sizes.B / 2#, -sizes.H / 2#, sizes.R1(2), sizes.GeometrySource, outgoingEdgeID:="Edge.BR" ' Нижний правый угол; Edge.BR — нижняя грань, размеры в мм.
    polygon.AddVertex "BL", -sizes.B / 2#, -sizes.H / 2#, sizes.R1(1), sizes.GeometrySource, outgoingEdgeID:="Edge.BL" ' Нижний левый угол; Edge.BL — участок к следующей вершине, размеры в мм.
    ' Вершина левого выступа до скругления; координаты и радиус в мм.
    If sizes.SideTypes(1) = "Tapered" Then polygon.AddVertex "LeftNose", -sizes.B / 2# - sizes.Widths(1), 0#, sizes.R2(1), sizes.GeometrySource, outgoingEdgeID:="Edge.LeftNose"
    polygon.Seal ' Проверяем вершины и завершаем их описание.
    Set DescribePolygon = polygon ' Возвращаем подготовленный многоугольник.
End Function

' Создаёт наружный контур по заданным вершинам и радиусам и сразу назначает его границей бетона.
' Ошибку связывает с исходными размерами.
Private Function AddGeometry(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRoundedInput, ByVal polygon As CPolygonSpec) As CGeometryPath
    On Error GoTo InvalidGeometry
    Dim outer As CGeometryPath
    Set outer = facade.RoundedPolygon("Boundary", polygon) ' Создаём готовый замкнутый контур по исходным вершинам и радиусам.
    ' Назначаем эту замкнутую линию наружной границей бетона.
    facade.Outer "OUTER", outer.ID, "Прямой участок контура скругленного прямоугольника.", arcComment:="Дуга контура скругленного прямоугольника.", namingPolicy:=SequentialByKind
    Set AddGeometry = outer
    Exit Function
InvalidGeometry:
    SectionRecipeGeometryError "CSectionRecipeRoundedRectangle", sizes.GeometrySource
End Function

' Создаёт линии арматуры и задаёт ряды вдоль четырёх сторон по подготовленным параметрам.
' Порядок сторон сохраняет принадлежность общего углового стержня первому из заданных рядов.
Private Sub AddReinforcement(ByVal facade As CSectionDefinitionBuilder, ByRef faces() As TRoundedFace, ByVal polygon As CPolygonSpec)
    Dim contact As CRebarContactBuilder, defaults As CRebarRuleOptions
    Dim normalSign As Double, i As Long, rowNumber As Long, path As CGeometryPath, options As CRebarRuleOptions
    For i = 1 To 4
        With faces(i)
            If .Count = 0 Or .Diameter = 0# Then GoTo NextFace
            Set path = AxisPath(facade, polygon, .ID, .AxisDistance, .AxisSource) ' Получаем линию центров стержней выбранной стороны с отступом as.
            Set options = New CRebarRuleOptions ' Создаём настройки размещения этого ряда.
            If i >= 3 Then options.EndMode = ExcludeBoth ' Не ставим стержни на концах участка после концевых отступов.
            options.StartSetback = .StartSetback: options.EndSetback = .EndSetback ' Отступы от начала и конца линии, мм.
            options.AnnotationGroup = .ID ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой стороны.
            options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
            options.Coincidence = KeepPublishedOwner ' При совпадении центров сохраняем стержень, который был добавлен раньше.
            options.IncomingTangentAtJoint = True ' На стыке используем направление предыдущего участка.
            options.Comment = "RoundedRectangle: " & .Caption & " грань, ряд 1"
            options.SourceReference = .Source ' Адреса полей для сообщения об ошибке размещения.
            facade.RebarByCount .ID & ".Main", path.ID, .Count, .Diameter, options ' Задаём основной ряд по числу стержней и диаметру.
            Select Case .ID
                Case "B.Top", "H.Right": normalSign = -1#
                Case "B.Bottom", "H.Left": normalSign = 1#
            End Select
            Set defaults = New CRebarRuleOptions ' Общие настройки только дополнительных рядов этой стороны.
            defaults.AnnotationGroup = .ID ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой стороны.
            defaults.Publication = PublishPerSourcePosition ' Сначала основной стержень, затем связанные с ним.
            defaults.Coincidence = KeepPublishedOwner ' В совпавшей позиции сохраняем ранее добавленный стержень.
            Set contact = New CRebarContactBuilder
            contact.Initialize .Diameter, normalSign, 1#, defaults
            For rowNumber = 2 To 3
                If .Diameters(rowNumber) > 0# Then
                    Set options = contact.AppendRow(.Diameters(rowNumber), .Locations(rowNumber), .Bindings(rowNumber), .ExtraSources(rowNumber), .BaseOffsets(rowNumber))
                    options.Comment = "RoundedRectangle: " & .Caption & " грань, ряд " & CStr(rowNumber)
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
        ByVal groupID As String, ByVal axisDistance As Double, ByVal sourceReference As String) As CGeometryPath
    On Error GoTo InvalidPath
    Dim offset As CGeometryPath
    Set offset = facade.OffsetRoundedPolygon("Offset." & groupID, polygon, axisDistance) ' Смещаем тот же многоугольник внутрь бетона на as, мм.
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
    SectionRecipeGeometryError "CSectionRecipeRoundedRectangle", sourceReference
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
    ' Ширина основы B до скругления углов; размер сверху.
    facade.HorizontalDistance "DIM_B", "B", -sizes.B / 2#, sizes.B / 2#, sizes.H / 2#, 1#, "Размер горизонтальной базы RoundedRectangle между левым и правым торцом."
    facade.VerticalExtent "DIM_H", outer.ID, "H", minX, -1#, "Полная высота RoundedRectangle." ' Полная высота сечения H; размер вынесен влево.
    ' Фактическая ширина левого выступа; подпись округляется до 1 мм.
    If sizes.SideTypes(1) = "Tapered" And -sizes.B / 2# - minX > GEOM_TOLERANCE Then facade.HorizontalDistance "DIM_W_LEFT", "W'", minX, -sizes.B / 2#, sizes.H / 2#, 1#, "Фактический вынос левого края бетона относительно левого торца базы B.", 1#
    ' Фактическая ширина правого выступа; подпись округляется до 1 мм.
    If sizes.SideTypes(2) = "Tapered" And maxX - sizes.B / 2# > GEOM_TOLERANCE Then facade.HorizontalDistance "DIM_W_RIGHT", "W'", sizes.B / 2#, maxX, sizes.H / 2#, 1#, "Фактический вынос правого края бетона относительно правого торца базы B.", 1#
End Sub

' Проверяет условие для указанной ячейки. При нарушении сообщает причину и её фактический адрес.
Private Sub Require(ByVal table As CSectionInputTable, ByVal row As Long, ByVal column As Long, ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then SectionRecipeError "CSectionRecipeRoundedRectangle", reason, table.SourceReference(row, column)
End Sub
