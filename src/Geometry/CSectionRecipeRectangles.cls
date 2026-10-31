VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionRecipeRectangles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Прямоугольное и составное бетонное сечение
' ==========================================================================
' Класс описывает один или два смежных прямоугольника, заданных размерами B1/H1 и B2/H2.
' Прямоугольник отсутствует, если хотя бы один его размер равен нулю; его армирование не читается.
' В составном сечении верхняя часть может быть смещена относительно нижней.
' По каждой стороне свои отступы и армирование; дополнительные ряды арматуры связаны с основным.
'
' Параметры читаются из таблицы выбранного варианта. Граница бетона, расположение
' арматуры и аннотации схемы (размеры сечения, количество и диаметры стержней) передаются
' в CSectionDefinitionBuilder, а стержни и бетонную сетку создаёт общий CSectionBuilder.

Implements ISectionRecipe

Private Type TRectanglesDimensions
    HasUpper As Boolean ' Первый прямоугольник имеет положительные ширину и высоту.
    HasLower As Boolean ' Второй прямоугольник имеет положительные ширину и высоту.
    B1 As Double ' Ширина первого прямоугольника, мм.
    H1 As Double ' Высота первого прямоугольника, мм.
    B2 As Double ' Ширина второго прямоугольника, мм.
    H2 As Double ' Высота второго прямоугольника, мм.
    Offset As Double ' Смещение верхнего прямоугольника по X относительно нижнего, мм.
    R As Double ' Общий радиус всех четырёх углов каждого заданного прямоугольника, мм.
    GeometrySource As String ' Фактические адреса размеров, радиуса и смещения для ошибок геометрического ядра.
End Type

Private Type TRectangleFace
    ID As String ' Имя стороны для привязки линий арматуры и подписей.
    Caption As String ' Название стороны из утверждённой таблицы на листе «Сечения».
    Count As Long ' Число стержней основного ряда.
    Diameter As Double ' Диаметр стержней основного ряда, мм.
    AxisDistance As Double ' Заданный as для явной связи линии арматуры с исходным прямоугольником, мм.
    ContourSource As String ' Исходный верхний или нижний прямоугольник, относительно которого задан as.
    LabelX As Double ' Компонента X направления выноса подписи.
    LabelY As Double ' Компонента Y направления выноса подписи.
    EndMode As ERebarEndMode ' Явный способ расстановки основного ряда на подготовленном пути.
    StartSetback As Double ' Отступ начала участка размещения от начала линии, мм.
    EndSetback As Double ' Отступ конца участка размещения от конца линии, мм.
    BaseOffsets(2 To 3) As Double ' Базовые отступы второго и третьего рядов от основного, мм.
    Diameters(2 To 3) As Double ' Диаметры второго и третьего рядов арматуры, мм.
    Locations(2 To 3) As String ' Прочитанное расположение каждого дополнительного ряда относительно основного.
    Bindings(2 To 3) As String ' Прочитанный способ привязки дополнительных рядов к основному.
    ExtraSources(2 To 3) As String ' Адреса исходных полей дополнительных рядов для сообщения об ошибке.
    Source As String ' Адреса отступов основного ряда для сообщения об ошибке размещения.
End Type

' Читает форму сечения и параметры армирования. Формирует описание наружной границы
' бетона, размещения стержней вдоль сторон и аннотаций схемы (размеры сечения,
' суммарное количество стержней каждого диаметра по каждой стороне) для общего сборщика.
Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
    On Error GoTo InvalidGeometry
    Dim reader As CSectionInputReader, facade As CSectionDefinitionBuilder
    Dim sizes As TRectanglesDimensions, faces() As TRectangleFace, upper As CGeometryPath, lower As CGeometryPath, outer As CGeometryPath
    Set reader = context.Reader("rngRectanglesGeometry") ' Находим поля таблицы по ключам и используем единицы ввода.
    ReadGeometryInputs reader, sizes ' Читаем и проверяем размеры сечения.
    ReadRebarInputs reader, sizes, faces ' Читаем параметры и готовим координаты линий армирования.
    Set facade = New CSectionDefinitionBuilder ' Создаём сборщик описания сечения.
    facade.Initialize "Rectangles" ' Начинаем описание выбранного типа.
    AddGeometry facade, sizes, upper, lower, outer ' Добавляем подготовленные границы бетона.
    AddReinforcement facade, faces ' Создаём линии арматуры и задаём правила размещения стержней.
    AddAnnotations facade, sizes, faces, upper, lower, outer ' Добавляем подписи арматуры, габариты и размеры сечения.
    Set ISectionRecipe_CreateDefinition = facade.Finish ' Завершаем и возвращаем проверенное описание сечения.
    Exit Function
InvalidGeometry:
    SectionRecipeGeometryError "CSectionRecipeRectangles", sizes.GeometrySource
End Function

' Читает четыре неотрицательных размера в мм; ноль любой стороны исключает её прямоугольник.
' Неактивные габариты обнуляются внутри описания; XOffset читается только при двух прямоугольниках.
Private Sub ReadGeometryInputs(ByVal reader As CSectionInputReader, ByRef sizes As TRectanglesDimensions)
    sizes.GeometrySource = reader.SourceReference("General", "B1") & "; " & reader.SourceReference("General", "H1") & "; " & reader.SourceReference("General", "B2") & "; " & reader.SourceReference("General", "H2") & "; " & reader.SourceReference("General", "R")
    sizes.H1 = reader.NonNegativeLength("General", "H1"): sizes.B1 = reader.NonNegativeLength("General", "B1")
    sizes.H2 = reader.NonNegativeLength("General", "H2"): sizes.B2 = reader.NonNegativeLength("General", "B2")
    sizes.HasUpper = (sizes.H1 > 0# And sizes.B1 > 0#)
    sizes.HasLower = (sizes.H2 > 0# And sizes.B2 > 0#)
    If Not sizes.HasUpper And Not sizes.HasLower Then SectionRecipeError "CSectionRecipeRectangles", "Не задано ни одного прямоугольника: хотя бы у одного из них B и H должны быть больше нуля.", sizes.GeometrySource
    If Not sizes.HasUpper Then sizes.H1 = 0#: sizes.B1 = 0#
    If Not sizes.HasLower Then sizes.H2 = 0#: sizes.B2 = 0#
    sizes.R = reader.NonNegativeLength("General", "R")
    If sizes.HasUpper And sizes.HasLower Then
        sizes.Offset = reader.Length("General", "XOffset")
        sizes.GeometrySource = sizes.GeometrySource & "; " & reader.SourceReference("General", "XOffset")
        reader.Require "General", "XOffset", "val", sizes.Offset < sizes.B2 - GEOM_TOLERANCE And sizes.Offset + sizes.B1 > GEOM_TOLERANCE, "Верхний и нижний прямоугольники должны иметь общий участок по X."
    End If
End Sub

' Читает параметры четырёх сторон каждого включённого прямоугольника.
' Строки Main/Extra отсутствующего прямоугольника вообще не читаются, независимо от их содержимого.
Private Sub ReadRebarInputs(ByVal reader As CSectionInputReader, ByRef sizes As TRectanglesDimensions, ByRef faces() As TRectangleFace)
    If sizes.HasUpper And sizes.HasLower Then ReDim faces(1 To 8) Else ReDim faces(1 To 4)
    If sizes.HasUpper Then ReadRectangleFaces reader, 1, 0, faces
    If sizes.HasLower Then ReadRectangleFaces reader, 2, IIf(sizes.HasUpper, 4, 0), faces
End Sub

' Сопоставляет ключи строк таблицы сечения со сторонами одного прямоугольника и готовит их параметры в мм.
' as всех сторон формирует общий контур осей; s0/s1 отсчитываются по его выбранному участку.
Private Sub ReadRectangleFaces(ByVal reader As CSectionInputReader, _
        ByVal part As Long, ByVal base As Long, ByRef faces() As TRectangleFace)
    Dim side As Long, axisDistance As Double, rowNumber As Long, ids As Variant, keys As Variant
    Dim key As String, prefix As String, suffix As String, captions As Variant
    keys = Array("Left", "Right", "Top", "Bottom")
    captions = Array(" - левая грань", " - правая грань", " - верхняя грань", " - нижняя грань")
    If part = 1 Then ids = Array("H1.as_1", "H1.as_2", "B1.as_1", "B1.as_2"): prefix = "R1." Else ids = Array("H2.as_1", "H2.as_2", "B2.as_1", "B2.as_2"): prefix = "R2."
    For side = 1 To 4
        key = prefix & CStr(keys(side - 1)) ' Ключ R1./R2. и имя стороны; side с 1 сопоставляется с Array, начинающимся с 0.
        With faces(base + side) ' В массив входят только активные части; part сохраняет исходные ключи R1/R2.
            .ID = CStr(ids(side - 1))
            .Count = reader.Count("Main", key, "n")
            .Diameter = reader.OptionalDiameter("Main", key, "d")
            .Source = reader.SourceReference("Main", key, "as")
            If part = 1 Then .ContourSource = "Upper" Else .ContourSource = "Lower"
            If .Count = 0 Or .Diameter = 0# Then
                .AxisDistance = reader.OptionalRowOffset("Main", key, "as")
                GoTo NextSide
            End If
            axisDistance = reader.Length("Main", key, "as")
            .AxisDistance = axisDistance
            reader.Require "Main", key, "as", axisDistance > .Diameter / 2#, "Расстояние до оси должно быть больше половины диаметра стержня."
            .StartSetback = reader.OptionalRowOffset("Main", key, "s0")
            .EndSetback = reader.OptionalRowOffset("Main", key, "s1")
            .EndMode = reader.RebarEndMode("Main", key)
            .Source = reader.SourceReferences("Main", key, Array("as", "s0", "s1", "endMode"))
            .Caption = Left$(.ID, 2) & CStr(captions(side - 1))
            Select Case side
                Case 1 ' Линия арматуры у левой грани: отступ as, сверху вниз.
                    .LabelX = -1#
                Case 2 ' Линия арматуры у правой грани: отступ as, сверху вниз.
                    .LabelX = 1#
                Case 3 ' Линия арматуры у верхней грани: отступ as, слева направо.
                    .LabelY = 1#
                Case 4 ' Линия арматуры у нижней грани: отступ as, слева направо.
                    .LabelY = -1#
            End Select
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
NextSide:
    Next side
End Sub

' Декларативно добавляет только заданные прямоугольники и их общую наружную границу.
' Одиночная часть начинается в (0, 0); для двух частей верхняя стоит над нижней с XOffset.
Private Sub AddGeometry(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRectanglesDimensions, ByRef upper As CGeometryPath, ByRef lower As CGeometryPath, ByRef outer As CGeometryPath)
    If sizes.HasUpper Then Set upper = facade.RoundedPolygon("Upper", DescribeRectangle(sizes.Offset, sizes.H2, sizes.B1, sizes.H1, sizes.R, sizes.GeometrySource))
    If sizes.HasLower Then Set lower = facade.RoundedPolygon("Lower", DescribeRectangle(0#, 0#, sizes.B2, sizes.H2, sizes.R, sizes.GeometrySource))
    If sizes.HasUpper And sizes.HasLower Then
        Set outer = facade.JoinedPaths("Boundary", Array("Lower", "Upper")) ' Общая граница получается из уже скруглённых частей, без повторной расстановки дуг в рецепте.
    ElseIf sizes.HasUpper Then
        Set outer = upper
    Else
        Set outer = lower
    End If
    facade.Outer "OUTER", outer.ID, "Сегмент контура сечения.", namingPolicy:=SequentialByKind ' Назначаем эту замкнутую линию наружной границей бетона.

End Sub

' Описывает четыре вершины отдельного прямоугольника; сопряжения строит общий RoundedPolygon.
' Имена сторон и углов используются и бетоном, и общим смещённым контуром арматуры.
Private Function DescribeRectangle(ByVal x As Double, ByVal y As Double, ByVal B As Double, ByVal H As Double, ByVal radius As Double, ByVal source As String) As CPolygonSpec
    Dim polygon As CPolygonSpec: Set polygon = New CPolygonSpec
    polygon.AddVertex "BL", x, y, radius, source, "Bottom", "Corner.BL"
    polygon.AddVertex "BR", x + B, y, radius, source, "Right", "Corner.BR"
    polygon.AddVertex "TR", x + B, y + H, radius, source, "Top", "Corner.TR"
    polygon.AddVertex "TL", x, y + H, radius, source, "Left", "Corner.TL"
    polygon.Seal
    Set DescribeRectangle = polygon
End Function

' Создаёт линии центров арматуры и задаёт основной и дополнительные ряды по подготовленным сторонам.
Private Sub AddReinforcement(ByVal facade As CSectionDefinitionBuilder, ByRef faces() As TRectangleFace)
    Dim contact As CRebarContactBuilder, extraOptions As CRebarRuleOptions
    Dim normalSign As Double, forward As Boolean, i As Long, rowNumber As Long, path As CGeometryPath, options As CRebarRuleOptions
    Dim contours(1 To 2) As CGeometryPath, groups As Variant, start As Long, side As Long, part As Long
    groups = Array("Left", "Right", "Corner.TL|Top|Corner.TR", "Corner.BL|Bottom|Corner.BR")
    For part = 1 To UBound(faces) \ 4
        start = (part - 1) * 4
        Set contours(part) = facade.OffsetContour("Rebar." & faces(start + 1).ContourSource, faces(start + 1).ContourSource, groups, _
            Array(faces(start + 1).AxisDistance, faces(start + 2).AxisDistance, faces(start + 3).AxisDistance, faces(start + 4).AxisDistance), _
            Array(faces(start + 1).Source, faces(start + 2).Source, faces(start + 3).Source, faces(start + 4).Source))
        For side = 1 To 4
            If faces(start + side).Count = 0 Or faces(start + side).Diameter = 0# Then facade.RebarContourBaseline "OUTER", faces(start + side).ContourSource, Split(CStr(groups(side - 1)), "|"), faces(start + side).AxisDistance, faces(start + side).Source
        Next side
    Next part
    For i = 1 To UBound(faces)
        With faces(i)
            If .Count = 0 Or .Diameter = 0# Then GoTo NextFace
            side = (i - 1) Mod 4 + 1 ' Для обоих прямоугольников: Left/Right сверху вниз, Top/Bottom слева направо.
            forward = (side = 1 Or side = 4) ' Left и Bottom следуют исходному контуру; Right и Top проходят его в обратную сторону.
            normalSign = -1#: If forward Then normalSign = 1# ' Знак левой нормали выбранного пути, направленной внутрь соответствующего прямоугольника.
            Set path = AxisPath(facade, .ID, contours((i - 1) \ 4 + 1), side) ' Горизонтальные ряды включают обе дуги, вертикальные идут по прямой между скруглениями.
            Set options = New CRebarRuleOptions ' Создаём настройки размещения этого ряда.
            options.StartSetback = .StartSetback: options.EndSetback = .EndSetback ' Отступы от начала и конца линии, мм.
            options.EndMode = .EndMode ' Применяем выбранное правило после концевых отступов.
            options.PositionOrder = FromPathMiddle ' Нумеруем стержни от середины линии.
            options.SourceReference = .Source ' Адреса полей для сообщения об ошибке размещения.
            options.AnnotationGroup = .ID ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой стороны.
            options.Publication = PublishPerSourcePosition ' Нумеруем R1, R2, ...: за каждым основным стержнем идут связанные с ним стержни дополнительных рядов.
            options.Coincidence = KeepPublishedOwner ' Первый добавленный физический стержень сохраняет свои диаметр и материал.
            options.WarnOnCoincidence = True ' Каждый пропуск совпавшего Main фиксируется предупреждением, включая одинаковые диаметры.
            options.IncomingTangentAtJoint = True ' На стыке прямой и дуги используется касательная предыдущего участка, как у HollowRectangle.
            options.Comment = "Набор прямоугольников: " & .Caption & ", ряд 1"
            facade.RebarByCount .ID & ".Main", path.ID, .Count, .Diameter, options ' Задаём основной ряд по числу стержней и диаметру.
            facade.RebarContour .ID & ".Main", "OUTER", .ContourSource, .AxisDistance, .ID & ".Row2", .ID & ".Row3", sourceSegments:=Split(CStr(groups(side - 1)), "|"), normalSign:=normalSign ' Связь с исходным прямоугольником и нормалью в бетон.
            Set extraOptions = New CRebarRuleOptions ' Общие настройки только дополнительных рядов этой стороны.
            extraOptions.AnnotationGroup = .ID ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой стороны.
            extraOptions.Publication = PublishPerSourcePosition ' Нумеруем R1, R2, ...: за каждым основным стержнем идут связанные с ним стержни дополнительных рядов.
            extraOptions.Coincidence = KeepPublishedOwner ' В совпавшей позиции сохраняем ранее добавленный стержень.
            Set contact = New CRebarContactBuilder
            contact.Initialize .Diameter, normalSign, 1#, extraOptions ' Stacked направлен в бетон, SideBySide — вперёд по направлению ряда.
            For rowNumber = 2 To 3
                If .Diameters(rowNumber) > 0# Then
                    Set options = contact.AppendRow(.Diameters(rowNumber), .Locations(rowNumber), .Bindings(rowNumber), .ExtraSources(rowNumber), .BaseOffsets(rowNumber))
                    options.Comment = "Набор прямоугольников: " & .Caption & ", ряд " & CStr(rowNumber)
                    facade.RebarDependent .ID & ".Row" & CStr(rowNumber), .ID & ".Main", .Diameters(rowNumber), options ' Оба дополнительных ряда связаны с основным.
                End If
            Next rowNumber
        End With
NextFace:
    Next i
End Sub

' Выбирает участок общего смещённого контура, как у наружной грани HollowRectangle.
' Верх и низ включают угловые дуги и направлены вправо; левая и правая прямые направлены вниз.
Private Function AxisPath(ByVal facade As CSectionDefinitionBuilder, ByVal pathID As String, ByVal contour As CGeometryPath, ByVal side As Long) As CGeometryPath
    Dim first As String, last As String, firstParameter As Double, lastParameter As Double
    Select Case side
        Case 1: Set AxisPath = facade.Chain(pathID, contour.ID, "Left", 0#, "Left", 1#)
        Case 2: Set AxisPath = facade.Chain(pathID, contour.ID, "Right", 1#, "Right", 0#, False)
        Case 3, 4
            first = "Left": last = "Right"
            If side = 4 Then firstParameter = 1# Else lastParameter = 1#
            If Not contour.HasSegmentID(first) Then
                If side = 3 Then first = "Corner.TL": firstParameter = 1# Else first = "Corner.BL": firstParameter = 0#
            End If
            If Not contour.HasSegmentID(last) Then
                If side = 3 Then last = "Corner.TR": lastParameter = 0# Else last = "Corner.BR": lastParameter = 1#
            End If
            Set AxisPath = facade.Chain(pathID, contour.ID, first, firstParameter, last, lastParameter, side = 4)
    End Select
End Function

' Добавляет подписи арматуры, общие габариты, размеры отдельных частей и ненулевое смещение верхней части.
Private Sub AddAnnotations(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRectanglesDimensions, ByRef faces() As TRectangleFace, ByVal upper As CGeometryPath, ByVal lower As CGeometryPath, ByVal outer As CGeometryPath)
    Dim i As Long
    For i = 1 To UBound(faces)
        If faces(i).Count > 0 And faces(i).Diameter > 0# Then facade.RebarRowLabel faces(i).ID, faces(i).ID & ".Main", faces(i).LabelX, faces(i).LabelY ' Подпись учитывает все ряды стороны и крайние стержни основного.
    Next i
    facade.HorizontalExtent "DIM_B", outer.ID, "B", 0#, -1#, "Общая ширина Rectangles-сечения." ' Общая ширина сечения B.
    facade.VerticalExtent "DIM_H", outer.ID, "H", GeomMin(0#, sizes.Offset), -1#, "Общая высота Rectangles-сечения." ' Полная высота сечения H; размер вынесен влево.
    If sizes.HasUpper Then facade.HorizontalExtent "DIM_TOP_WIDTH", upper.ID, "B1", sizes.H1 + sizes.H2, 1#, "Ширина первого прямоугольника Rectangles-сечения."
    If sizes.HasLower Then
        facade.VerticalExtent "DIM_RIGHT_HEIGHT", lower.ID, "H2", sizes.B2, 1#, "Высота нижнего прямоугольника Rectangles-сечения." ' Высота нижнего прямоугольника; размер справа.
        If Not sizes.HasUpper Then facade.HorizontalExtent "DIM_BOTTOM_WIDTH", lower.ID, "B2", sizes.H2, 1#, "Ширина второго прямоугольника Rectangles-сечения."
    End If
    If sizes.HasUpper And sizes.HasLower Then
        If Abs(sizes.Offset) > GEOM_TOLERANCE Then facade.HorizontalDistance "DIM_UPPER_OFFSET_X", ChrW$(&H394), GeomMin(0#, sizes.Offset), GeomMax(0#, sizes.Offset), sizes.H2, 1#, "Горизонтальное смещение верхнего прямоугольника относительно базового положения." ' Размер идёт слева направо: его величина равна Abs(sizes.Offset), а выноски остаются у смещённых граней.
    End If
End Sub
