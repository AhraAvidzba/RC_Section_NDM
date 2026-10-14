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
' Декларативный рецепт Rectangle, LSection и TwoRectangles
' ==========================================================================
Implements ISectionRecipe

' Один рецепт Rectangle/LSection/TwoRectangles: свой табличный ввод,
' смежные прямоугольники, восемь смысловых граней и зависимые ряды.
Private Type TRectSetDimensions
    Shape As String
    Header As Long
    B1 As Double
    H1 As Double
    B2 As Double
    H2 As Double
    Offset As Double
End Type

' Описывает контуры, арматурные правила и размеры выбранного рецепта. Сетку и физические стержни
' строит общий CSectionBuilder.
Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
    Dim table As CSectionInputTable, units As CUnitSystem, facade As CSectionDefinitionBuilder  ' Объявляем сборщик описания сечения.
    Dim sizes As TRectSetDimensions, upper As CGeometryPath, lower As CGeometryPath, outer As CGeometryPath  ' Объявляем линии и параметры геометрии.
    Set table = context.Table("rngRectSetGeometry")
    Set units = context.Units
    ReadDimensions table, units, sizes
    Set facade = New CSectionDefinitionBuilder  ' Создаём описание составного сечения.
    facade.Initialize "RectSet"  ' "RectSet" — имя типа сечения.
    ' facade — сборщик описания сечения; sizes — подготовленные размеры и параметры формы; upper —
    ' возвращаемый контур верхнего прямоугольника; lower — возвращаемый контур нижнего прямоугольника;
    ' outer — возвращаемая общая наружная граница.
    AddGeometry facade, sizes, upper, lower, outer
    ' facade — сборщик описания сечения; table — снимок таблицы армирования; units — перевод размеров
    ' в мм; sizes.Header — номер строки шапки таблицы; 0 — смещение группы строк: 0 верхний, 4 нижний
    ' прямоугольник; sizes.Offset — X нижнего левого угла, мм; sizes.H2 — Y нижнего левого угла, мм;
    ' sizes.B1 — ширина прямоугольника, мм; sizes.H1 — высота прямоугольника, мм.
    AddRectangleRows facade, table, units, sizes.Header, 0, sizes.Offset, sizes.H2, sizes.B1, sizes.H1
    ' facade — сборщик описания сечения; table — снимок таблицы армирования; units — перевод размеров
    ' в мм; sizes.Header — номер строки шапки таблицы; 4 — смещение группы строк: 0 верхний, 4 нижний
    ' прямоугольник; 0 — X нижнего левого угла, мм; 0 — Y нижнего левого угла, мм; sizes.B2 — ширина
    ' прямоугольника, мм; sizes.H2 — высота прямоугольника, мм.
    If sizes.Shape <> "Rectangle" Then AddRectangleRows facade, table, units, sizes.Header, 4, 0#, 0#, sizes.B2, sizes.H2
    ' facade — сборщик описания сечения; sizes — размеры и данные формы; upper — верхний
    ' прямоугольник; lower — нижний прямоугольник; outer — наружная граница для габаритов.
    AddDimensions facade, sizes, upper, lower, outer
    Set ISectionRecipe_CreateDefinition = facade.Finish  ' Проверяем и возвращаем готовое описание сечения.
End Function

' Каждый вариант читает только свои активные размеры; отсутствующая нижняя
' часть Rectangle не требует заполнения полей остальных двух вариантов.
' В каждой паре сначала читаются оба числа, затем проверяется положительность H и B.
' PositiveLength проверяет сразу: его прямой вызов изменил бы первую ошибку при H=0 и текстовом B.
' Поэтому здесь сохранены Length + Require, прежние адреса/пояснения и один перевод каждого размера.
Private Sub ReadDimensions(ByVal table As CSectionInputTable, ByVal units As CUnitSystem, ByRef sizes As TRectSetDimensions)
    sizes.Shape = table.Choice(table.RowByKey("RectSet.SectionType"), 2, _
        Array("Rectangle", "LSection", "TwoRectangles", "RectSet", "TwoRectangle", "TwoRectanglesOffset", "OffsetTwoRectangles"))
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
        If sizes.Shape = "LSection" Then Require table, sizes.Header + 1, 2, sizes.B1 < sizes.B2, _
            "Для LSection ширина верхнего прямоугольника должна быть меньше нижнего."
    End If
    If sizes.Shape = "TwoRectangles" Then
        Dim row As Long
        row = table.RowByKey("RectSet.UpperOffsetX")
        sizes.Offset = table.Length(row, 2, units)
        Require table, row, 2, sizes.Offset < sizes.B2 - GEOM_TOLERANCE And sizes.Offset + sizes.B1 > GEOM_TOLERANCE, _
            "Верхний и нижний прямоугольники должны иметь общий участок по X."
    End If
End Sub

' LSection и TwoRectangles имеют одно построение; различаются размеры и
' допустимое смещение, а не общий алгоритм объединения прямоугольников.
Private Sub AddGeometry(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRectSetDimensions, _
        ByRef upper As CGeometryPath, ByRef lower As CGeometryPath, ByRef outer As CGeometryPath)
    ' "Upper" — имя прямоугольника; sizes.Offset — X левого нижнего угла, мм; sizes.H2 — Y левого
    ' нижнего угла, мм; sizes.B1 — ширина, мм; sizes.H1 — высота, мм.
    Set upper = facade.RectanglePath("Upper", sizes.Offset, sizes.H2, sizes.B1, sizes.H1)
    If sizes.Shape = "Rectangle" Then
        Set outer = upper  ' Принимаем один прямоугольник за наружный контур.
    Else
        Set lower = facade.RectanglePath("Lower", 0#, 0#, sizes.B2, sizes.H2)  ' "Lower" — имя прямоугольника; 0 — X левого нижнего угла, мм; 0 — Y левого нижнего угла, мм; sizes.B2 — ширина, мм; sizes.H2 — высота, мм.
        ' "Boundary" — имя общего контура; rectangles — прямоугольники: X,Y нижнего угла, ширина,
        ' высота, мм.
        Set outer = facade.AdjacentRectangles("Boundary", Array(Array(0#, 0#, sizes.B2, sizes.H2), Array(sizes.Offset, sizes.H2, sizes.B1, sizes.H1)))
    End If
    ' "OUTER" — имя наружной границы бетона; outer.ID — замкнутый контур; comment — пояснение к
    ' границе; SequentialByKind — способ публикации имён участков.
    facade.Outer "OUTER", outer.ID, "Сегмент контура сечения.", namingPolicy:=SequentialByKind

End Sub

' Описывает габариты и размерные привязки рецепта отдельно от расчётных контуров.
Private Sub AddDimensions(ByVal facade As CSectionDefinitionBuilder, ByRef sizes As TRectSetDimensions, _
        ByVal upper As CGeometryPath, ByVal lower As CGeometryPath, ByVal outer As CGeometryPath)
    ' "DIM_B" — имя размера; outer.ID — контур, по которому измеряется ширина; "B" — обозначение на
    ' схеме; 0 — положение исходной размерной линии Y, мм; -1 — вынос размера вниз; comment —
    ' пояснение к размеру.
    facade.HorizontalExtent "DIM_B", outer.ID, "B", 0#, -1#, "Общая ширина RectSet-сечения."
    ' "DIM_H" — имя размера; outer.ID — контур, по которому измеряется высота; "H" — обозначение на
    ' схеме; GeomMin(0, sizes.Offset) — положение исходной размерной линии X, мм; -1 — вынос размера
    ' влево; comment — пояснение к размеру.
    facade.VerticalExtent "DIM_H", outer.ID, "H", GeomMin(0#, sizes.Offset), -1#, "Общая высота RectSet-сечения."
    ' "DIM_TOP_WIDTH" — имя размера; upper.ID — контур, по которому измеряется ширина; "B1" —
    ' обозначение на схеме; sizes.H1 + sizes.H2 — положение исходной размерной линии Y, мм; 1 — вынос
    ' размера вверх; comment — пояснение к размеру.
    facade.HorizontalExtent "DIM_TOP_WIDTH", upper.ID, "B1", sizes.H1 + sizes.H2, 1#, "Ширина верхнего прямоугольника RectSet-сечения."
    If sizes.Shape <> "Rectangle" Then
        ' "DIM_RIGHT_HEIGHT" — имя размера; lower.ID — контур, по которому измеряется высота; "H2" —
        ' обозначение на схеме; sizes.B2 — положение исходной размерной линии X, мм; 1 — вынос размера
        ' вправо; comment — пояснение к размеру.
        facade.VerticalExtent "DIM_RIGHT_HEIGHT", lower.ID, "H2", sizes.B2, 1#, "Высота нижнего прямоугольника RectSet-сечения."
        ' "DIM_UPPER_OFFSET_X" — имя размера; ChrW$(&H394) — обозначение на схеме; 0 — X первой точки,
        ' мм; sizes.Offset — X второй точки, мм; sizes.H2 — уровень исходной размерной линии Y, мм; 1
        ' — вынос размера вверх; comment — пояснение к размеру.
        If Abs(sizes.Offset) > GEOM_TOLERANCE Then facade.HorizontalDistance "DIM_UPPER_OFFSET_X", ChrW$(&H394), 0#, sizes.Offset, sizes.H2, 1#, "Горизонтальное смещение верхнего прямоугольника относительно базового положения."
    End If
End Sub

' База отступов - полная длина грани, не линия, укороченная соседним as.
' Направления обхода и касательного SideBySide заданы явно в рецепте.
Private Sub AddRectangleRows(ByVal facade As CSectionDefinitionBuilder, ByVal table As CSectionInputTable, _
        ByVal units As CUnitSystem, ByVal header As Long, ByVal base As Long, _
        ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal height As Double)
    Dim side As Long, row As Long, extra As Long, axisDistance As Double, diameter As Double, count As Long
    Dim path As CGeometryPath, groupID As String, label As String, options As CRebarRuleOptions  ' Объявляем линии геометрии и условия армирования.
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
        Set options = New CRebarRuleOptions  ' Создаём условия размещения арматуры.
        options.StartSetback = table.Length(row, 5, units)  ' StartSetback — отступ начала ряда от начала грани, мм.
        options.EndSetback = table.Length(row, 6, units)  ' EndSetback — отступ конца ряда от конца грани, мм.
        options.PositionOrder = FromPathMiddle  ' PositionOrder — FromPathMiddle нумерует позиции от середины линии.
        ' Публикация задаётся ниже после выбора смыслового groupID.  ' Publication — PublishPerSourcePosition сохраняет имена по исходным позициям.
        options.SourceReference = table.SourceReference(row, 2) & "; " & table.SourceReference(row, 5) & "; " & table.SourceReference(row, 6)  ' SourceReference — адреса ячеек, которые надо исправить при ошибке размещения.
        label = table.Text(row, 1) & " грань"
        Select Case side
            Case 1
                groupID = "H" & prefix & ".as_1": nx = -1#: ny = 0#: sideSign = -1#
                Set path = facade.LinePath(groupID, x + axisDistance, y, x + axisDistance, y + height)  ' groupID — имя линии; x + axisDistance — X начала, мм; y — Y начала, мм; x + axisDistance — X конца, мм; y + height — Y конца, мм.
            Case 2
                groupID = "H" & prefix & ".as_2": nx = 1#: ny = 0#: sideSign = 1#
                ' groupID — имя линии; x + width - axisDistance — X начала, мм; y + height — Y начала,
                ' мм; x + width - axisDistance — X конца, мм; y — Y конца, мм.
                Set path = facade.LinePath(groupID, x + width - axisDistance, y + height, x + width - axisDistance, y)
            Case 3
                groupID = "B" & prefix & ".as_1": nx = 0#: ny = 1#: sideSign = 1#
                ' groupID — имя линии; x — X начала, мм; y + height - axisDistance — Y начала, мм; x +
                ' width — X конца, мм; y + height - axisDistance — Y конца, мм.
                Set path = facade.LinePath(groupID, x, y + height - axisDistance, x + width, y + height - axisDistance)
            Case 4
                groupID = "B" & prefix & ".as_2": nx = 0#: ny = -1#: sideSign = -1#
                Set path = facade.LinePath(groupID, x + width, y + axisDistance, x, y + axisDistance)  ' groupID — имя линии; x + width — X начала, мм; y + axisDistance — Y начала, мм; x — X конца, мм; y + axisDistance — Y конца, мм.
        End Select
        options.SetPublication groupID, PublishPerSourcePosition, KeepPublishedOwner
        options.Comment = "RectSet: " & label & ", ряд 1"  ' Comment — пояснение к стержням данного ряда.
        ' groupID & ".Main" — имя нового ряда; path.ID — линия центров стержней; count — заданное
        ' число стержней; diameter — диаметр стержня, мм; options — отступы, нумерация и условия
        ' размещения.
        facade.RebarByCount groupID & ".Main", path.ID, count, diameter, options
        d2 = table.OptionalDiameter(extra, 2, units)
        d3 = table.OptionalDiameter(extra, 5, units)
        If d2 > 0# Then
            loc2 = table.Choice(extra, 3, Array("Stacked", "SideBySide"))
            ' table — снимок таблицы армирования; extra — строка дополнительных рядов; 2 — столбец их
            ' диаметра; groupID — общая группа подписи; label — название стороны; 2 — номер
            ' дополнительного ряда; loc2 — Stacked поперёк ряда, SideBySide вдоль; diameter — диаметр
            ' основного стержня, мм; d2 — диаметр добавляемого стержня, мм; sideSign — знак
            ' направления на данной стороне.
            Set options = AdditionalOptions(table, extra, 2, groupID, label, 2, loc2, diameter, d2, sideSign)
            ' groupID & ".Row2" — имя дополнительного ряда; groupID & ".Main" — исходный ряд, чьи
            ' позиции используются; d2 — диаметр дополнительного стержня, мм; options — смещение и
            ' выбор позиций исходного ряда.
            facade.RebarDependent groupID & ".Row2", groupID & ".Main", d2, options
        End If
        If d3 > 0# Then
            loc3 = table.Choice(extra, 6, Array("Stacked", "SideBySide"))
            Dim distance As Double
            distance = 0#
            If d2 > 0# And loc2 = loc3 Then distance = d2
            ' table — снимок таблицы армирования; extra — строка дополнительных рядов; 5 — столбец их
            ' диаметра; groupID — общая группа подписи; label — название стороны; 3 — номер
            ' дополнительного ряда; loc3 — Stacked поперёк ряда, SideBySide вдоль; diameter — диаметр
            ' основного стержня, мм; d3 — диаметр добавляемого стержня, мм; sideSign — знак
            ' направления на данной стороне; distance — сумма промежуточных номинальных диаметров, мм.
            Set options = AdditionalOptions(table, extra, 5, groupID, label, 3, loc3, diameter, d3, sideSign, distance)
            ' groupID & ".Row3" — имя дополнительного ряда; groupID & ".Main" — исходный ряд, чьи
            ' позиции используются; d3 — диаметр дополнительного стержня, мм; options — смещение и
            ' выбор позиций исходного ряда.
            facade.RebarDependent groupID & ".Row3", groupID & ".Main", d3, options
        End If
        ' groupID — группа общей подписи всех её рядов; groupID & ".Main" — ряд, к концам которого
        ' привязана подпись; nx — X-компонента направления подписи; ny — Y-компонента направления
        ' подписи.
        facade.RebarRowLabel groupID, groupID & ".Main", nx, ny
NextSide:
    Next side
End Sub

' Описывает смещение и привязку дополнительного ряда относительно основного, сохраняя смысловую
' группу.
Private Function AdditionalOptions(ByVal table As CSectionInputTable, ByVal row As Long, ByVal column As Long, _
        ByVal groupID As String, ByVal label As String, ByVal rowNumber As Long, ByVal location As String, _
        ByVal baseDiameter As Double, ByVal diameter As Double, ByVal sideSign As Double, _
        Optional ByVal interveningDiameterSum As Double = 0#) As CRebarRuleOptions
    Dim options As CRebarRuleOptions  ' Объявляем условия размещения арматуры.
    Set options = New CRebarRuleOptions  ' Создаём условия размещения арматуры.
    ' NormalOffset — знаковое смещение по нормали к исходному ряду, мм; TangentOffset — знаковое
    ' смещение вдоль исходного ряда, мм.
    Dim axis As ERebarOffsetAxis, directionSign As Double, everySecond As Boolean
    axis = RowNormal: directionSign = -1#
    If location = "SideBySide" Then axis = RowTangent: directionSign = sideSign
    everySecond = (table.Choice(row, column + 2, Array("EachBar", "EverySecondBar")) = "EverySecondBar")
    options.SourceReference = table.SourceReference(row, column)
    options.ConfigureDependentContact baseDiameter, diameter, axis, directionSign, everySecond, interveningDiameterSum
    options.SetPublication groupID, PublishPerSourcePosition, KeepPublishedOwner
    options.Comment = "RectSet: " & label & ", ряд " & CStr(rowNumber)  ' Comment — пояснение к стержням данного ряда.
    Set AdditionalOptions = options  ' Возвращаем условия размещения дополнительного ряда.
End Function

' Проверяет ограничение рецепта; ошибка содержит фактический адрес параметра во входной таблице.
Private Sub Require(ByVal table As CSectionInputTable, ByVal row As Long, ByVal column As Long, ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then SectionRecipeError "CSectionRecipeRectSet", reason, table.SourceReference(row, column)
End Sub
