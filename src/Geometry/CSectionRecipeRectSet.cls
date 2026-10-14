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
' У каждой грани свои отступы и армирование; дополнительные ряды связаны с основным.
'
' Размеры читаются из таблицы выбранного варианта. Контуры, ряды и размеры передаются
' в CSectionDefinitionBuilder, а стержни и бетонную сетку создаёт общий CSectionBuilder.

Implements ISectionRecipe

Private Type TRectSetDimensions
    Shape As String
    Header As Long
    B1 As Double
    H1 As Double
    B2 As Double
    H2 As Double
    Offset As Double
End Type

' Читает выбранный вариант сечения и его армирование. Возвращает описание общей
' наружной границы, рядов на гранях и размеров для общего сборщика.
Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
    Dim table As CSectionInputTable, units As CUnitSystem, facade As CSectionDefinitionBuilder
    Dim sizes As TRectSetDimensions, upper As CGeometryPath, lower As CGeometryPath, outer As CGeometryPath
    Set table = context.Table("rngRectSetGeometry") ' Берём копию входной таблицы и единицы измерения.
    Set units = context.Units
    ReadDimensions table, units, sizes ' Читаем и проверяем размеры сечения.
    Set facade = New CSectionDefinitionBuilder ' Создаём сборщик описания сечения.
    facade.Initialize "RectSet" ' Начинаем описание выбранного типа.
    AddGeometry facade, sizes, upper, lower, outer ' Добавляем подготовленные границы бетона.
    AddRectangleRows facade, table, units, sizes.Header, 0, sizes.Offset, sizes.H2, sizes.B1, sizes.H1 ' Армирование четырёх граней верхнего прямоугольника.
    If sizes.Shape <> "Rectangle" Then AddRectangleRows facade, table, units, sizes.Header, 4, 0#, 0#, sizes.B2, sizes.H2 ' Армирование нижней части, если она есть.
    AddDimensions facade, sizes, upper, lower, outer ' Добавляем габариты и размеры сечения.
    Set ISectionRecipe_CreateDefinition = facade.Finish ' Завершаем и возвращаем проверенное описание сечения.
End Function

' Заполняет sizes размерами выбранного варианта в миллиметрах и проверяет их допустимость.
' Для одного прямоугольника не читает размеры нижней части; смещение нужно только двум прямоугольникам.
' В каждой паре сначала читаются H и B, затем проверяются их знаки. Так текст в B не
' теряется за ошибкой знака H, а сообщение сохраняет прежний адрес исходной ячейки.
Private Sub ReadDimensions(ByVal table As CSectionInputTable, ByVal units As CUnitSystem, ByRef sizes As TRectSetDimensions)
    sizes.Shape = table.Choice(table.RowByKey("RectSet.SectionType"), 2, Array("Rectangle", "LSection", "TwoRectangles", "RectSet", "TwoRectangle", "TwoRectanglesOffset", "OffsetTwoRectangles"))
    If sizes.Shape = "RectSet" Then sizes.Shape = "LSection"
    If sizes.Shape <> "Rectangle" And sizes.Shape <> "LSection" Then sizes.Shape = "TwoRectangles"
    sizes.Header = table.RowByKey("H1")
    sizes.H1 = table.Length(sizes.Header + 1, 1, units)
    sizes.B1 = table.Length(sizes.Header + 1, 2, units)
    Require table, sizes.Header + 1, 1, sizes.H1 > 0#, "Высота верхнего прямоугольника должна быть положительной."
    Require table, sizes.Header + 1, 2, sizes.B1 > 0#, "Ширина верхнего прямоугольника должна быть положительной."
    If sizes.Shape <> "Rectangle" Then
        sizes.H2 = table.Length(sizes.Header + 1, 3, units)
        sizes.B2 = table.Length(sizes.Header + 1, 4, units)
        Require table, sizes.Header + 1, 3, sizes.H2 > 0#, "Высота нижнего прямоугольника должна быть положительной."
        Require table, sizes.Header + 1, 4, sizes.B2 > 0#, "Ширина нижнего прямоугольника должна быть положительной."
        If sizes.Shape = "LSection" Then Require table, sizes.Header + 1, 2, sizes.B1 < sizes.B2, "Для LSection ширина верхнего прямоугольника должна быть меньше нижнего."
    End If
    If sizes.Shape = "TwoRectangles" Then
        Dim row As Long
        row = table.RowByKey("RectSet.UpperOffsetX")
        sizes.Offset = table.Length(row, 2, units)
        Require table, row, 2, sizes.Offset < sizes.B2 - GEOM_TOLERANCE And sizes.Offset + sizes.B1 > GEOM_TOLERANCE, "Верхний и нижний прямоугольники должны иметь общий участок по X."
    End If
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

' Добавляет на схему общие габариты, размеры отдельных частей и ненулевое смещение верхней части.
Private Sub AddDimensions(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRectSetDimensions, ByVal upper As CGeometryPath, ByVal lower As CGeometryPath, ByVal outer As CGeometryPath)
    facade.HorizontalExtent "DIM_B", outer.ID, "B", 0#, -1#, "Общая ширина RectSet-сечения." ' Общая ширина сечения B.
    facade.VerticalExtent "DIM_H", outer.ID, "H", GeomMin(0#, sizes.Offset), -1#, "Общая высота RectSet-сечения." ' Полная высота сечения H; размер вынесен влево.
    facade.HorizontalExtent "DIM_TOP_WIDTH", upper.ID, "B1", sizes.H1 + sizes.H2, 1#, "Ширина верхнего прямоугольника RectSet-сечения." ' Ширина верхнего прямоугольника; размер над сечением.
    If sizes.Shape <> "Rectangle" Then
        facade.VerticalExtent "DIM_RIGHT_HEIGHT", lower.ID, "H2", sizes.B2, 1#, "Высота нижнего прямоугольника RectSet-сечения." ' Высота нижнего прямоугольника; размер справа.
        ' Ненулевое смещение верхней части относительно нижней.
        If Abs(sizes.Offset) > GEOM_TOLERANCE Then facade.HorizontalDistance "DIM_UPPER_OFFSET_X", ChrW$(&H394), 0#, sizes.Offset, sizes.H2, 1#, "Горизонтальное смещение верхнего прямоугольника относительно базового положения."
    End If
End Sub

' Задаёт основной и дополнительные ряды на четырёх гранях одного прямоугольника.
' Отступы считаются от полной грани; подпись каждой стороны объединяет все её ряды.
Private Sub AddRectangleRows(ByVal facade As CSectionDefinitionBuilder, ByVal table As CSectionInputTable, _
        ByVal units As CUnitSystem, ByVal header As Long, ByVal base As Long, _
        ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal height As Double)
    Dim side As Long, row As Long, extra As Long, axisDistance As Double, diameter As Double, count As Long
    Dim path As CGeometryPath, groupID As String, label As String, options As CRebarRuleOptions
    Dim nx As Double, ny As Double, sideSign As Double, d2 As Double, d3 As Double, loc2 As String, loc3 As String
    Dim prefix As String
    If base = 0 Then prefix = "1" Else prefix = "2"
    For side = 1 To 4
        row = header + 3 + base + side
        extra = header + 13 + base + side
        count = table.Count(row, 4)
        diameter = table.OptionalDiameter(row, 3, units)
        If count = 0 Or diameter = 0# Then GoTo NextSide
        axisDistance = table.Length(row, 2, units)
        Require table, row, 2, axisDistance > diameter / 2#, "Расстояние до оси должно быть больше половины диаметра стержня."
        Set options = New CRebarRuleOptions ' Создаём настройки размещения этого ряда.
        options.StartSetback = table.Length(row, 5, units) ' Отступ от начала линии, мм.
        options.EndSetback = table.Length(row, 6, units) ' Отступ от конца линии, мм.
        options.PositionOrder = FromPathMiddle ' Нумеруем стержни от середины линии.
        options.SourceReference = table.SourceReference(row, 2) & "; " & table.SourceReference(row, 5) & "; " & table.SourceReference(row, 6) ' Адреса полей для сообщения об ошибке размещения.
        label = table.Text(row, 1) & " грань"
        Select Case side
            Case 1
                groupID = "H" & prefix & ".as_1": nx = -1#: ny = 0#: sideSign = -1#
                Set path = facade.LinePath(groupID, x + axisDistance, y, x + axisDistance, y + height) ' Левая грань на расстоянии as от бетона, снизу вверх.
            Case 2
                groupID = "H" & prefix & ".as_2": nx = 1#: ny = 0#: sideSign = 1#
                Set path = facade.LinePath(groupID, x + width - axisDistance, y + height, x + width - axisDistance, y) ' Правая грань на расстоянии as от бетона, сверху вниз.
            Case 3
                groupID = "B" & prefix & ".as_1": nx = 0#: ny = 1#: sideSign = 1#
                Set path = facade.LinePath(groupID, x, y + height - axisDistance, x + width, y + height - axisDistance) ' Верхняя грань на расстоянии as от бетона, слева направо.
            Case 4
                groupID = "B" & prefix & ".as_2": nx = 0#: ny = -1#: sideSign = -1#
                Set path = facade.LinePath(groupID, x + width, y + axisDistance, x, y + axisDistance) ' Нижняя грань на расстоянии as от бетона, справа налево.
        End Select
        options.AnnotationGroup = groupID ' Общая подпись арматуры этой группы.
        options.Publication = PublishPerSourcePosition ' Связанные стержни выводим рядом с основным.
        options.Coincidence = KeepPublishedOwner ' При совпадении сохраняем первый опубликованный стержень.
        options.Comment = "RectSet: " & label & ", ряд 1"
        facade.RebarByCount groupID & ".Main", path.ID, count, diameter, options ' Задаём основной ряд по числу стержней и диаметру.
        d2 = table.OptionalDiameter(extra, 2, units)
        d3 = table.OptionalDiameter(extra, 5, units)
        If d2 > 0# Then
            loc2 = table.Choice(extra, 3, Array("Stacked", "SideBySide"))
            Set options = AdditionalOptions(table, extra, 2, groupID, label, 2, loc2, diameter, d2, sideSign) ' Настраиваем смещение и привязку добавляемого ряда.
            facade.RebarDependent groupID & ".Row2", groupID & ".Main", d2, options ' Второй ряд привязан к позициям основного; задаём его диаметр.
        End If
        If d3 > 0# Then
            loc3 = table.Choice(extra, 6, Array("Stacked", "SideBySide"))
            Dim distance As Double
            distance = 0#
            If d2 > 0# And loc2 = loc3 Then distance = d2 ' Третий ряд в том же направлении должен обойти стержень второго.
            Set options = AdditionalOptions(table, extra, 5, groupID, label, 3, loc3, diameter, d3, sideSign, distance) ' Настраиваем смещение и привязку добавляемого ряда.
            facade.RebarDependent groupID & ".Row3", groupID & ".Main", d3, options ' Третий ряд привязан к позициям основного; задаём его диаметр.
        End If
        facade.RebarRowLabel groupID, groupID & ".Main", nx, ny ' Подпись учитывает все ряды стороны и крайние стержни основного.
NextSide:
    Next side
End Sub

' Возвращает настройки дополнительного ряда: контактное расстояние, направление
' и размещение у каждого основного стержня или через один. Все ряды стороны имеют общую подпись.
Private Function AdditionalOptions(ByVal table As CSectionInputTable, ByVal row As Long, ByVal column As Long, _
        ByVal groupID As String, ByVal label As String, ByVal rowNumber As Long, ByVal location As String, _
        ByVal baseDiameter As Double, ByVal diameter As Double, ByVal sideSign As Double, _
        Optional ByVal interveningDiameterSum As Double = 0#) As CRebarRuleOptions
    Dim options As CRebarRuleOptions
    Set options = New CRebarRuleOptions ' Создаём настройки размещения этого ряда.
    Dim axis As ERebarOffsetAxis, directionSign As Double, everySecond As Boolean
    axis = RowNormal: directionSign = -1# ' По умолчанию смещаем ряд по нормали к исходной линии.
    If location = "SideBySide" Then axis = RowTangent: directionSign = sideSign ' Располагаем дополнительный стержень вдоль исходного ряда.
    everySecond = (table.Choice(row, column + 2, Array("EachBar", "EverySecondBar")) = "EverySecondBar") ' Выбираем все позиции или через одну, начиная с первой.
    options.SourceReference = table.SourceReference(row, column) ' Адреса полей для сообщения об ошибке размещения.
    options.ConfigureDependentContact baseDiameter, diameter, axis, directionSign, everySecond, interveningDiameterSum ' Учитываем диаметры обоих рядов и промежуточные стержни.
    options.AnnotationGroup = groupID ' Общая подпись арматуры этой группы.
    options.Publication = PublishPerSourcePosition ' Связанные стержни выводим рядом с основным.
    options.Coincidence = KeepPublishedOwner ' При совпадении сохраняем первый опубликованный стержень.
    options.Comment = "RectSet: " & label & ", ряд " & CStr(rowNumber)
    Set AdditionalOptions = options ' Возвращаем настройки дополнительного ряда.
End Function

' Проверяет условие для указанной ячейки. При нарушении сообщает причину и её фактический адрес.
Private Sub Require(ByVal table As CSectionInputTable, ByVal row As Long, ByVal column As Long, ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then SectionRecipeError "CSectionRecipeRectSet", reason, table.SourceReference(row, column)
End Sub
