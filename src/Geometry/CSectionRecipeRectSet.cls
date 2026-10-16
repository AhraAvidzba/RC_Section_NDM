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

' ==========================================================================
' Прямоугольное и составное бетонное сечение
' ==========================================================================
' Класс описывает один прямоугольник, Г-образное сечение или два смежных прямоугольника.
' В составном сечении верхняя часть может быть смещена относительно нижней.
' По каждой стороне свои отступы и армирование; дополнительные ряды арматуры связаны с основным.
'
' Параметры читаются из таблицы выбранного варианта. Граница бетона, расположение
' арматуры и аннотации схемы (размеры сечения, количество и диаметры стержней) передаются
' в CSectionDefinitionBuilder, а стержни и бетонную сетку создаёт общий CSectionBuilder.

Implements ISectionRecipe

Private Type TRectSetDimensions
    Shape As String ' Название формы, составленной из прямоугольников.
    B1 As Double ' Ширина первого прямоугольника, мм.
    H1 As Double ' Высота первого прямоугольника, мм.
    B2 As Double ' Ширина второго прямоугольника, мм.
    H2 As Double ' Высота второго прямоугольника, мм.
    Offset As Double ' Смещение присоединённого прямоугольника, мм.
End Type

Private Type TRectangleFace
    ID As String ' Имя стороны для привязки линий арматуры и подписей.
    Caption As String ' Название стороны из утверждённой таблицы Config.
    Count As Long ' Число стержней основного ряда.
    Diameter As Double ' Диаметр стержней основного ряда, мм.
    X1 As Double ' Начальная координата X линии центров арматуры, мм.
    Y1 As Double ' Начальная координата Y линии центров арматуры, мм.
    X2 As Double ' Конечная координата X линии центров арматуры, мм.
    Y2 As Double ' Конечная координата Y линии центров арматуры, мм.
    LabelX As Double ' Компонента X направления выноса подписи.
    LabelY As Double ' Компонента Y направления выноса подписи.
    TangentSign As Double ' Направление сдвига связанных стержней вдоль линии.
    StartSetback As Double ' Отступ первого стержня от начала стороны, мм.
    EndSetback As Double ' Отступ последнего стержня от конца стороны, мм.
    Diameters(2 To 3) As Double ' Диаметры второго и третьего рядов арматуры, мм.
    Locations(2 To 3) As String ' Расположение каждого дополнительного ряда относительно основного.
    Bindings(2 To 3) As String ' Выбор основных стержней для второго и третьего рядов.
    SkippedDiameters(2 To 3) As Double ' Полный диаметр промежуточного ряда в том же направлении, мм.
    Source As String ' Адреса отступов основного ряда для сообщения об ошибке размещения.
    ExtraSources(2 To 3) As String ' Адреса диаметров дополнительных рядов для сообщения об ошибке.
End Type

' Читает форму сечения и параметры армирования. Формирует описание наружной границы
' бетона, размещения стержней вдоль сторон и аннотаций схемы (размеры сечения,
' суммарное количество стержней каждого диаметра по каждой стороне) для общего сборщика.
Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
    Dim table As CSectionInputTable, units As CUnitSystem, facade As CSectionDefinitionBuilder
    Dim sizes As TRectSetDimensions, faces() As TRectangleFace, upper As CGeometryPath, lower As CGeometryPath, outer As CGeometryPath
    Set table = context.Table("rngRectSetGeometry") ' Берём копию входной таблицы и единицы измерения.
    Set units = context.Units
    ReadGeometryInputs table, units, sizes ' Читаем и проверяем размеры сечения.
    ReadRebarInputs table, units, sizes, faces ' Читаем параметры и готовим координаты линий армирования.
    Set facade = New CSectionDefinitionBuilder ' Создаём сборщик описания сечения.
    facade.Initialize "RectSet" ' Начинаем описание выбранного типа.
    AddGeometry facade, sizes, upper, lower, outer ' Добавляем подготовленные границы бетона.
    AddReinforcement facade, faces ' Создаём линии арматуры и задаём правила размещения стержней.
    AddAnnotations facade, sizes, faces, upper, lower, outer ' Добавляем подписи арматуры, габариты и размеры сечения.
    Set ISectionRecipe_CreateDefinition = facade.Finish ' Завершаем и возвращаем проверенное описание сечения.
End Function

' Заполняет sizes размерами выбранного варианта в миллиметрах и проверяет их допустимость.
' Для одного прямоугольника не читает размеры нижней части; смещение нужно только двум прямоугольникам.
' В каждой паре сначала читаются H и B, затем проверяются их знаки. Так текст в B не
' теряется за ошибкой знака H, а сообщение сохраняет прежний адрес исходной ячейки.
Private Sub ReadGeometryInputs(ByVal table As CSectionInputTable, ByVal units As CUnitSystem, ByRef sizes As TRectSetDimensions)
    Dim header As Long, row As Long
    sizes.Shape = table.Choice(table.RowByKey("RectSet.SectionType"), 2, Array("Rectangle", "LSection", "TwoRectangles", "RectSet", "TwoRectangle", "TwoRectanglesOffset", "OffsetTwoRectangles"))
    If sizes.Shape = "RectSet" Then sizes.Shape = "LSection"
    If sizes.Shape <> "Rectangle" And sizes.Shape <> "LSection" Then sizes.Shape = "TwoRectangles"
    header = table.RowByKey("H1")
    sizes.H1 = table.Length(header + 1, 1, units)
    sizes.B1 = table.Length(header + 1, 2, units)
    Require table, header + 1, 1, sizes.H1 > 0#, "Высота верхнего прямоугольника должна быть положительной."
    Require table, header + 1, 2, sizes.B1 > 0#, "Ширина верхнего прямоугольника должна быть положительной."
    If sizes.Shape <> "Rectangle" Then
        sizes.H2 = table.Length(header + 1, 3, units)
        sizes.B2 = table.Length(header + 1, 4, units)
        Require table, header + 1, 3, sizes.H2 > 0#, "Высота нижнего прямоугольника должна быть положительной."
        Require table, header + 1, 4, sizes.B2 > 0#, "Ширина нижнего прямоугольника должна быть положительной."
        If sizes.Shape = "LSection" Then Require table, header + 1, 2, sizes.B1 < sizes.B2, "Для LSection ширина верхнего прямоугольника должна быть меньше нижнего."
    End If
    If sizes.Shape = "TwoRectangles" Then
        row = table.RowByKey("RectSet.UpperOffsetX")
        sizes.Offset = table.Length(row, 2, units)
        Require table, row, 2, sizes.Offset < sizes.B2 - GEOM_TOLERANCE And sizes.Offset + sizes.B1 > GEOM_TOLERANCE, "Верхний и нижний прямоугольники должны иметь общий участок по X."
    End If
End Sub

' Читает параметры четырёх сторон каждого включённого прямоугольника.
' У одиночного прямоугольника строки нижней части не читаются; выключенные ряды не требуют отступов и привязок.
Private Sub ReadRebarInputs(ByVal table As CSectionInputTable, ByVal units As CUnitSystem, ByRef sizes As TRectSetDimensions, ByRef faces() As TRectangleFace)
    Dim header As Long
    header = table.RowByKey("H1")
    If sizes.Shape = "Rectangle" Then ReDim faces(1 To 4) Else ReDim faces(1 To 8)
    ReadRectangleFaces table, units, header, 0, sizes.Offset, sizes.H2, sizes.B1, sizes.H1, faces
    If sizes.Shape <> "Rectangle" Then ReadRectangleFaces table, units, header, 4, 0#, 0#, sizes.B2, sizes.H2, faces
End Sub

' Сопоставляет строки Config со сторонами одного прямоугольника и готовит их параметры в мм.
' Отступы считаются от полной грани; направления линий и связанных рядов сохраняются.
Private Sub ReadRectangleFaces(ByVal table As CSectionInputTable, ByVal units As CUnitSystem, ByVal header As Long, _
        ByVal base As Long, ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal height As Double, ByRef faces() As TRectangleFace)
    Dim side As Long, row As Long, extra As Long, axisDistance As Double, rowNumber As Long, column As Long, ids As Variant
    ' Строки каждого прямоугольника в Config: левая, правая, верхняя и нижняя стороны.
    If base = 0 Then ids = Array("H1.as_1", "H1.as_2", "B1.as_1", "B1.as_2") Else ids = Array("H2.as_1", "H2.as_2", "B2.as_1", "B2.as_2")
    For side = 1 To 4
        row = header + 3 + base + side
        extra = header + 13 + base + side
        With faces(base + side)
            .ID = CStr(ids(side - 1))
            .Count = table.Count(row, 4)
            .Diameter = table.OptionalDiameter(row, 3, units)
            If .Count = 0 Or .Diameter = 0# Then GoTo NextSide
            axisDistance = table.Length(row, 2, units)
            Require table, row, 2, axisDistance > .Diameter / 2#, "Расстояние до оси должно быть больше половины диаметра стержня."
            .StartSetback = table.Length(row, 5, units)
            .EndSetback = table.Length(row, 6, units)
            .Source = table.SourceReference(row, 2) & "; " & table.SourceReference(row, 5) & "; " & table.SourceReference(row, 6)
            .Caption = table.Text(row, 1) & " грань"
            Select Case side
                Case 1 ' Левая грань на расстоянии as от бетона, снизу вверх.
                    .X1 = x + axisDistance: .Y1 = y: .X2 = x + axisDistance: .Y2 = y + height
                    .LabelX = -1#: .TangentSign = -1#
                Case 2 ' Правая грань на расстоянии as от бетона, сверху вниз.
                    .X1 = x + width - axisDistance: .Y1 = y + height: .X2 = x + width - axisDistance: .Y2 = y
                    .LabelX = 1#: .TangentSign = 1#
                Case 3 ' Верхняя грань на расстоянии as от бетона, слева направо.
                    .X1 = x: .Y1 = y + height - axisDistance: .X2 = x + width: .Y2 = y + height - axisDistance
                    .LabelY = 1#: .TangentSign = 1#
                Case 4 ' Нижняя грань на расстоянии as от бетона, справа налево.
                    .X1 = x + width: .Y1 = y + axisDistance: .X2 = x: .Y2 = y + axisDistance
                    .LabelY = -1#: .TangentSign = -1#
            End Select
            .Diameters(2) = table.OptionalDiameter(extra, 2, units)
            .Diameters(3) = table.OptionalDiameter(extra, 5, units)
            For rowNumber = 2 To 3
                If .Diameters(rowNumber) > 0# Then
                    column = 2 + (rowNumber - 2) * 3 ' В таблице на каждый ряд отведены диаметр, расположение и привязка.
                    .Locations(rowNumber) = table.Choice(extra, column + 1, Array("Stacked", "SideBySide"))
                    .Bindings(rowNumber) = table.Choice(extra, column + 2, Array("EachBar", "EverySecondBar"))
                    .ExtraSources(rowNumber) = table.SourceReference(extra, column)
                End If
            Next rowNumber
            If .Locations(2) = .Locations(3) Then .SkippedDiameters(3) = .Diameters(2) ' Обходим номинальный второй ряд в том же направлении.
        End With
NextSide:
    Next side
End Sub

' Добавляет контуры верхнего и, если он нужен, нижнего прямоугольника. Через upper,
' lower и outer отдаёт их линии и общую наружную границу; для одного прямоугольника lower пуст.
Private Sub AddGeometry(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRectSetDimensions, ByRef upper As CGeometryPath, ByRef lower As CGeometryPath, ByRef outer As CGeometryPath)
    Set upper = facade.RectanglePath("Upper", sizes.Offset, sizes.H2, sizes.B1, sizes.H1) ' Верхняя часть: левый нижний угол, ширина и высота, мм.
    If sizes.Shape = "Rectangle" Then
        Set outer = upper ' У одиночного прямоугольника его контур и есть наружная граница.
    Else
        Set lower = facade.RectanglePath("Lower", 0#, 0#, sizes.B2, sizes.H2) ' Нижняя часть начинается в (0, 0); B2 и H2 — её габариты, мм.
        ' Получаем общую границу двух смежных прямоугольников.
        Set outer = facade.AdjacentRectangles("Boundary", Array(Array(0#, 0#, sizes.B2, sizes.H2), Array(sizes.Offset, sizes.H2, sizes.B1, sizes.H1)))
    End If
    facade.Outer "OUTER", outer.ID, "Сегмент контура сечения.", namingPolicy:=SequentialByKind ' Назначаем эту замкнутую линию наружной границей бетона.

End Sub

' Создаёт линии центров арматуры и задаёт основной и дополнительные ряды по подготовленным сторонам.
Private Sub AddReinforcement(ByVal facade As CSectionDefinitionBuilder, ByRef faces() As TRectangleFace)
    Dim i As Long, rowNumber As Long, path As CGeometryPath, options As CRebarRuleOptions
    For i = 1 To UBound(faces)
        With faces(i)
            If .Count = 0 Or .Diameter = 0# Then GoTo NextFace
            Set path = facade.LinePath(.ID, .X1, .Y1, .X2, .Y2) ' Линия центров стержней с подготовленными координатами, мм.
            Set options = New CRebarRuleOptions ' Создаём настройки размещения этого ряда.
            options.StartSetback = .StartSetback: options.EndSetback = .EndSetback ' Отступы от начала и конца линии, мм.
            options.PositionOrder = FromPathMiddle ' Нумеруем стержни от середины линии.
            options.SourceReference = .Source ' Адреса полей для сообщения об ошибке размещения.
            options.AnnotationGroup = .ID ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой стороны.
            options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
            options.Coincidence = KeepPublishedOwner ' При совпадении центров сохраняем стержень, который был добавлен раньше.
            options.Comment = "RectSet: " & .Caption & ", ряд 1"
            facade.RebarByCount .ID & ".Main", path.ID, .Count, .Diameter, options ' Задаём основной ряд по числу стержней и диаметру.
            For rowNumber = 2 To 3
                If .Diameters(rowNumber) > 0# Then
                    Set options = RecipeContactOptions(.ExtraSources(rowNumber), .Locations(rowNumber), .Bindings(rowNumber), _
                        .Diameter, .Diameters(rowNumber), -1#, .TangentSign, .SkippedDiameters(rowNumber))
                    options.AnnotationGroup = .ID ' Одна подпись для всех рядов этой стороны.
                    options.Publication = PublishPerSourcePosition ' Сначала основной стержень, затем связанные с ним.
                    options.Coincidence = KeepPublishedOwner ' В совпавшей позиции сохраняем ранее добавленный стержень.
                    options.Comment = "RectSet: " & .Caption & ", ряд " & CStr(rowNumber)
                    facade.RebarDependent .ID & ".Row" & CStr(rowNumber), .ID & ".Main", .Diameters(rowNumber), options ' Оба дополнительных ряда связаны с основным.
                End If
            Next rowNumber
        End With
NextFace:
    Next i
End Sub

' Добавляет подписи арматуры, общие габариты, размеры отдельных частей и ненулевое смещение верхней части.
Private Sub AddAnnotations(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRectSetDimensions, ByRef faces() As TRectangleFace, ByVal upper As CGeometryPath, ByVal lower As CGeometryPath, ByVal outer As CGeometryPath)
    Dim i As Long
    For i = 1 To UBound(faces)
        If faces(i).Count > 0 And faces(i).Diameter > 0# Then facade.RebarRowLabel faces(i).ID, faces(i).ID & ".Main", faces(i).LabelX, faces(i).LabelY ' Подпись учитывает все ряды стороны и крайние стержни основного.
    Next i
    facade.HorizontalExtent "DIM_B", outer.ID, "B", 0#, -1#, "Общая ширина RectSet-сечения." ' Общая ширина сечения B.
    facade.VerticalExtent "DIM_H", outer.ID, "H", GeomMin(0#, sizes.Offset), -1#, "Общая высота RectSet-сечения." ' Полная высота сечения H; размер вынесен влево.
    facade.HorizontalExtent "DIM_TOP_WIDTH", upper.ID, "B1", sizes.H1 + sizes.H2, 1#, "Ширина верхнего прямоугольника RectSet-сечения." ' Ширина верхнего прямоугольника; размер над сечением.
    If sizes.Shape <> "Rectangle" Then
        facade.VerticalExtent "DIM_RIGHT_HEIGHT", lower.ID, "H2", sizes.B2, 1#, "Высота нижнего прямоугольника RectSet-сечения." ' Высота нижнего прямоугольника; размер справа.
        ' Размер идёт слева направо: его величина равна Abs(sizes.Offset), а выноски остаются у смещённых граней.
        If Abs(sizes.Offset) > GEOM_TOLERANCE Then facade.HorizontalDistance "DIM_UPPER_OFFSET_X", ChrW$(&H394), GeomMin(0#, sizes.Offset), GeomMax(0#, sizes.Offset), sizes.H2, 1#, "Горизонтальное смещение верхнего прямоугольника относительно базового положения."
    End If
End Sub

' Проверяет условие для указанной ячейки. При нарушении сообщает причину и её фактический адрес.
Private Sub Require(ByVal table As CSectionInputTable, ByVal row As Long, ByVal column As Long, ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then SectionRecipeError "CSectionRecipeRectSet", reason, table.SourceReference(row, column)
End Sub
