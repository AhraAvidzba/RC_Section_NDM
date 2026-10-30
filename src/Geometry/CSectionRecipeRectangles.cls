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
' Класс описывает один прямоугольник OneRect или два смежных прямоугольника TwoRect.
' В составном сечении верхняя часть может быть смещена относительно нижней.
' По каждой стороне свои отступы и армирование; дополнительные ряды арматуры связаны с основным.
'
' Параметры читаются из таблицы выбранного варианта. Граница бетона, расположение
' арматуры и аннотации схемы (размеры сечения, количество и диаметры стержней) передаются
' в CSectionDefinitionBuilder, а стержни и бетонную сетку создаёт общий CSectionBuilder.

Implements ISectionRecipe

Private Type TRectanglesDimensions
    Shape As String ' Название формы, составленной из прямоугольников.
    B1 As Double ' Ширина первого прямоугольника, мм.
    H1 As Double ' Высота первого прямоугольника, мм.
    B2 As Double ' Ширина второго прямоугольника, мм.
    H2 As Double ' Высота второго прямоугольника, мм.
    Offset As Double ' Смещение верхнего прямоугольника по X относительно нижнего, мм.
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
End Function

' Заполняет sizes размерами выбранного варианта в миллиметрах и проверяет их допустимость.
' Для одного прямоугольника не читает размеры нижней части; смещение нужно только двум прямоугольникам.
' В каждой паре сначала читаются H и B, затем проверяются их знаки. Так текст в B не
' теряется за ошибкой знака H, а сообщение содержит адрес исходной ячейки.
Private Sub ReadGeometryInputs(ByVal reader As CSectionInputReader, ByRef sizes As TRectanglesDimensions)
    sizes.Shape = reader.Choice("General", "Shape", Array("TwoRect", "OneRect"))
    sizes.H1 = reader.Length("General", "H1"): sizes.B1 = reader.Length("General", "B1")
    reader.Require "General", "H1", "val", sizes.H1 > 0#, "Высота верхнего прямоугольника должна быть положительной."
    reader.Require "General", "B1", "val", sizes.B1 > 0#, "Ширина верхнего прямоугольника должна быть положительной."
    If sizes.Shape <> "OneRect" Then
        sizes.H2 = reader.Length("General", "H2"): sizes.B2 = reader.Length("General", "B2")
        reader.Require "General", "H2", "val", sizes.H2 > 0#, "Высота нижнего прямоугольника должна быть положительной."
        reader.Require "General", "B2", "val", sizes.B2 > 0#, "Ширина нижнего прямоугольника должна быть положительной."
    End If
    If sizes.Shape = "TwoRect" Then
        sizes.Offset = reader.Length("General", "XOffset")
        reader.Require "General", "XOffset", "val", sizes.Offset < sizes.B2 - GEOM_TOLERANCE And sizes.Offset + sizes.B1 > GEOM_TOLERANCE, "Верхний и нижний прямоугольники должны иметь общий участок по X."
    End If
End Sub

' Читает параметры четырёх сторон каждого включённого прямоугольника.
' У одиночного прямоугольника строки нижней части не читаются; выключенные ряды не требуют отступов и привязок.
Private Sub ReadRebarInputs(ByVal reader As CSectionInputReader, ByRef sizes As TRectanglesDimensions, ByRef faces() As TRectangleFace)
    If sizes.Shape = "OneRect" Then ReDim faces(1 To 4) Else ReDim faces(1 To 8)
    ReadRectangleFaces reader, 0, faces
    If sizes.Shape <> "OneRect" Then ReadRectangleFaces reader, 4, faces
End Sub

' Сопоставляет ключи строк таблицы сечения со сторонами одного прямоугольника и готовит их параметры в мм.
' as всех сторон формирует общий контур осей; s0/s1 отсчитываются по его выбранному участку.
Private Sub ReadRectangleFaces(ByVal reader As CSectionInputReader, _
        ByVal base As Long, ByRef faces() As TRectangleFace)
    Dim side As Long, axisDistance As Double, rowNumber As Long, ids As Variant, keys As Variant
    Dim key As String, prefix As String, suffix As String, captions As Variant
    keys = Array("Left", "Right", "Top", "Bottom")
    captions = Array(" - левая грань", " - правая грань", " - верхняя грань", " - нижняя грань")
    If base = 0 Then ids = Array("H1.as_1", "H1.as_2", "B1.as_1", "B1.as_2"): prefix = "R1." Else ids = Array("H2.as_1", "H2.as_2", "B2.as_1", "B2.as_2"): prefix = "R2."
    For side = 1 To 4
        key = prefix & CStr(keys(side - 1)) ' Ключ R1./R2. и имя стороны; side с 1 сопоставляется с Array, начинающимся с 0.
        With faces(base + side) ' base равен 0 для R1 и 4 для R2; объединённый массив не задаёт адреса входных ячеек.
            .ID = CStr(ids(side - 1))
            .Count = reader.Count("Main", key, "n")
            .Diameter = reader.OptionalDiameter("Main", key, "d")
            .Source = reader.SourceReference("Main", key, "as")
            If base = 0 Then .ContourSource = "Upper" Else .ContourSource = "Lower"
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
                Case 1 ' Линия арматуры у левой грани: отступ as, снизу вверх.
                    .LabelX = -1#
                Case 2 ' Линия арматуры у правой грани: отступ as, сверху вниз.
                    .LabelX = 1#
                Case 3 ' Линия арматуры у верхней грани: отступ as, слева направо.
                    .LabelY = 1#
                Case 4 ' Линия арматуры у нижней грани: отступ as, справа налево.
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

' Добавляет контуры верхнего и, если он нужен, нижнего прямоугольника. Через upper,
' lower и outer отдаёт их линии и общую наружную границу; для одного прямоугольника lower пуст.
Private Sub AddGeometry(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRectanglesDimensions, ByRef upper As CGeometryPath, ByRef lower As CGeometryPath, ByRef outer As CGeometryPath)
    Set upper = facade.RectanglePath("Upper", sizes.Offset, sizes.H2, sizes.B1, sizes.H1, "Bottom", "Right", "Top", "Left") ' Верхняя часть: левый нижний угол, ширина и высота, мм.
    If sizes.Shape = "OneRect" Then
        Set outer = upper ' У одиночного прямоугольника его контур и есть наружная граница.
    Else
        Set lower = facade.RectanglePath("Lower", 0#, 0#, sizes.B2, sizes.H2, "Bottom", "Right", "Top", "Left") ' Нижняя часть начинается в (0, 0); B2 и H2 — её габариты, мм.
        Set outer = facade.AdjacentRectangles("Boundary", Array(Array(0#, 0#, sizes.B2, sizes.H2), Array(sizes.Offset, sizes.H2, sizes.B1, sizes.H1))) ' Получаем общую границу двух смежных прямоугольников.
    End If
    facade.Outer "OUTER", outer.ID, "Сегмент контура сечения.", namingPolicy:=SequentialByKind ' Назначаем эту замкнутую линию наружной границей бетона.

End Sub

' Создаёт линии центров арматуры и задаёт основной и дополнительные ряды по подготовленным сторонам.
Private Sub AddReinforcement(ByVal facade As CSectionDefinitionBuilder, ByRef faces() As TRectangleFace)
    Dim contact As CRebarContactBuilder, extraOptions As CRebarRuleOptions
    Dim tangentSign As Double, i As Long, rowNumber As Long, path As CGeometryPath, options As CRebarRuleOptions
    Dim contours(1 To 2) As CGeometryPath, groups As Variant, start As Long, side As Long, part As Long
    groups = Array("Left", "Right", "Top", "Bottom")
    For part = 1 To UBound(faces) \ 4
        start = (part - 1) * 4
        Set contours(part) = facade.OffsetContour("Rebar." & faces(start + 1).ContourSource, faces(start + 1).ContourSource, groups, _
            Array(faces(start + 1).AxisDistance, faces(start + 2).AxisDistance, faces(start + 3).AxisDistance, faces(start + 4).AxisDistance), _
            Array(faces(start + 1).Source, faces(start + 2).Source, faces(start + 3).Source, faces(start + 4).Source))
        For side = 1 To 4
            If faces(start + side).Count = 0 Or faces(start + side).Diameter = 0# Then facade.RebarContourBaseline "OUTER", faces(start + side).ContourSource, Array(CStr(groups(side - 1))), faces(start + side).AxisDistance, faces(start + side).Source
        Next side
    Next part
    For i = 1 To UBound(faces)
        With faces(i)
            If .Count = 0 Or .Diameter = 0# Then GoTo NextFace
            Set path = facade.Chain(.ID, contours((i - 1) \ 4 + 1).ID, CStr(groups((i - 1) Mod 4)), 1#, CStr(groups((i - 1) Mod 4)), 0#, False) ' Направления: Left снизу вверх, Right сверху вниз, Top слева направо, Bottom справа налево; концы учитывают соседние as.
            Set options = New CRebarRuleOptions ' Создаём настройки размещения этого ряда.
            options.StartSetback = .StartSetback: options.EndSetback = .EndSetback ' Отступы от начала и конца линии, мм.
            options.EndMode = .EndMode ' Применяем выбранное правило после концевых отступов.
            options.PositionOrder = FromPathMiddle ' Нумеруем стержни от середины линии.
            options.SourceReference = .Source ' Адреса полей для сообщения об ошибке размещения.
            options.AnnotationGroup = .ID ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой стороны.
            options.Publication = PublishPerSourcePosition ' Нумеруем R1, R2, ...: за каждым основным стержнем идут связанные с ним стержни дополнительных рядов.
            options.Coincidence = KeepPublishedOwner ' Первый добавленный физический стержень сохраняет свои диаметр и материал.
            options.WarnOnCoincidence = True ' Каждый пропуск совпавшего Main фиксируется предупреждением, включая одинаковые диаметры.
            options.Comment = "Набор прямоугольников: " & .Caption & ", ряд 1"
            facade.RebarByCount .ID & ".Main", path.ID, .Count, .Diameter, options ' Задаём основной ряд по числу стержней и диаметру.
            facade.RebarContour .ID & ".Main", "OUTER", .ContourSource, .AxisDistance, .ID & ".Row2", .ID & ".Row3", sourceSegments:=Array(CStr(groups((i - 1) Mod 4))), normalSign:=-1# ' Связь с исходным прямоугольником и известным направлением в бетон.
            Select Case (i - 1) Mod 4 + 1 ' Повторяем направления сторон Left/Right/Top/Bottom для обоих прямоугольников.
                Case 1, 4: tangentSign = -1#
                Case 2, 3: tangentSign = 1#
            End Select
            Set extraOptions = New CRebarRuleOptions ' Общие настройки только дополнительных рядов этой стороны.
            extraOptions.AnnotationGroup = .ID ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой стороны.
            extraOptions.Publication = PublishPerSourcePosition ' Нумеруем R1, R2, ...: за каждым основным стержнем идут связанные с ним стержни дополнительных рядов.
            extraOptions.Coincidence = KeepPublishedOwner ' В совпавшей позиции сохраняем ранее добавленный стержень.
            Set contact = New CRebarContactBuilder
            contact.Initialize .Diameter, -1#, tangentSign, extraOptions ' Stacked направлен в бетон справа от обхода; знак касательной сохраняет выбранную ориентацию SideBySide.
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

' Добавляет подписи арматуры, общие габариты, размеры отдельных частей и ненулевое смещение верхней части.
Private Sub AddAnnotations(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRectanglesDimensions, ByRef faces() As TRectangleFace, ByVal upper As CGeometryPath, ByVal lower As CGeometryPath, ByVal outer As CGeometryPath)
    Dim i As Long
    For i = 1 To UBound(faces)
        If faces(i).Count > 0 And faces(i).Diameter > 0# Then facade.RebarRowLabel faces(i).ID, faces(i).ID & ".Main", faces(i).LabelX, faces(i).LabelY ' Подпись учитывает все ряды стороны и крайние стержни основного.
    Next i
    facade.HorizontalExtent "DIM_B", outer.ID, "B", 0#, -1#, "Общая ширина Rectangles-сечения." ' Общая ширина сечения B.
    facade.VerticalExtent "DIM_H", outer.ID, "H", GeomMin(0#, sizes.Offset), -1#, "Общая высота Rectangles-сечения." ' Полная высота сечения H; размер вынесен влево.
    facade.HorizontalExtent "DIM_TOP_WIDTH", upper.ID, "B1", sizes.H1 + sizes.H2, 1#, "Ширина верхнего прямоугольника Rectangles-сечения." ' Ширина верхнего прямоугольника; размер над сечением.
    If sizes.Shape <> "OneRect" Then
        facade.VerticalExtent "DIM_RIGHT_HEIGHT", lower.ID, "H2", sizes.B2, 1#, "Высота нижнего прямоугольника Rectangles-сечения." ' Высота нижнего прямоугольника; размер справа.
        If Abs(sizes.Offset) > GEOM_TOLERANCE Then facade.HorizontalDistance "DIM_UPPER_OFFSET_X", ChrW$(&H394), GeomMin(0#, sizes.Offset), GeomMax(0#, sizes.Offset), sizes.H2, 1#, "Горизонтальное смещение верхнего прямоугольника относительно базового положения." ' Размер идёт слева направо: его величина равна Abs(sizes.Offset), а выноски остаются у смещённых граней.
    End If
End Sub
