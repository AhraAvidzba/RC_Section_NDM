VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionRecipeCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Декларативный рецепт круглого сечения и трёх арматурных рядов
' ==========================================================================
Implements ISectionRecipe

' Рецепт Circle: свой ввод, точный круг, основной и связанные ряды.
' Сетку, физическую расстановку и групповые подписи выполняет общий builder.
Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
    Dim table As CSectionInputTable, units As CUnitSystem, facade As CSectionDefinitionBuilder  ' Объявляем сборщик описания сечения.
    Dim diameter As Double, axisDistance As Double, count As Long, d As Double, d2 As Double, d3 As Double
    Dim loc2 As String, loc3 As String, options As CRebarRuleOptions, path As CGeometryPath, skipped As Double  ' Объявляем линии геометрии и условия армирования.

    ' Ввод: дополнительные поля читаются только для активных рядов.
    Set table = context.Table("rngCircleGeometry"): Set units = context.Units
    diameter = table.Length(table.RowByKey("Circle.Diameter"), 2, units)
    Require table, "Circle.Diameter", diameter > 0#, "Диаметр бетонного круга должен быть положительным."
    count = table.Count(table.RowByKey("Rebar.Count"), 2)
    d = table.OptionalDiameter(table.RowByKey("Rebar.Diameter"), 2, units)

    ' Геометрия: наружная окружность и точная траектория осей стержней.
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Circle"              ' Создаём описание круглого сечения.
    Set path = facade.CirclePath("CONTOUR_CIRCLE", 0#, 0#, diameter / 2#)               ' Создаём наружную окружность бетона.
    facade.Outer "OUTER", path.ID, "Круговой контур сечения."                           ' Назначаем наружную границу бетона.
    If count > 0 And d > 0# Then
        Require table, "Rebar.Count", count >= 2, "Для круговой раскладки требуется не меньше двух стержней первого ряда."
        axisDistance = table.Length(table.RowByKey("Rebar.AxisDistance"), 2, units)
        Require table, "Rebar.AxisDistance", axisDistance > d / 2#, "Расстояние до оси должно быть больше половины диаметра стержня."
        Require table, "Rebar.AxisDistance", axisDistance < diameter / 2#, "Расстояние до оси должно быть меньше радиуса бетонного круга."
        d2 = table.OptionalDiameter(table.RowByKey("Rebar.Diameter2"), 2, units)
        d3 = table.OptionalDiameter(table.RowByKey("Rebar.Diameter3"), 2, units)
        If d2 > 0# Then loc2 = table.Choice(table.RowByKey("Rebar.Loc2row"), 2, Array("Stacked", "SideBySide"))
        If d3 > 0# Then loc3 = table.Choice(table.RowByKey("Rebar.Loc3row"), 2, Array("Stacked", "SideBySide"))
        Set path = facade.CirclePath("Circle.Axis", 0#, 0#, diameter / 2# - axisDistance)  ' Создаём окружность центров стержней с отступом as.

        ' Арматура: основной ряд и зависимые дополнительные ряды.
        Set options = New CRebarRuleOptions                                             ' Создаём условия размещения арматуры.
        options.EndMode = ClosedUniform: options.SourcePrefix = "B": options.AnnotationGroup = "Circle"  ' Замкнутый ряд; имена и общая группа подписи.
        options.Publication = PublishPerSourcePosition: options.Comment = "Круг: основной ряд арматуры"  ' Задаём имена стержней и пояснение первого ряда.
        options.SourceReference = table.SourceReference(table.RowByKey("Rebar.Count"), 2) & "; " & table.SourceReference(table.RowByKey("Rebar.AxisDistance"), 2)  ' Привязываем ошибки размещения к ячейкам ввода.
        facade.RebarByCount "Circle.Main", path.ID, count, d, options                   ' Задаём основной ряд по числу и диаметру стержней.
        If d2 > 0# Then
            Set options = AdditionalOptions(table, "Rebar.Diameter2", "_R2", loc2, d / 2# + d2 / 2#)  ' Получаем смещение и привязку дополнительного ряда.
            facade.RebarDependent "Circle.Row2", "Circle.Main", d2, options             ' Задаём второй ряд относительно основного.
        End If
        If d3 > 0# Then
            If d2 > 0# And loc2 = loc3 Then skipped = d2
            Set options = AdditionalOptions(table, "Rebar.Diameter3", "_R3", loc3, d / 2# + skipped + d3 / 2#)  ' Получаем смещение и привязку дополнительного ряда.
            facade.RebarDependent "Circle.Row3", "Circle.Main", d3, options             ' Задаём третий ряд относительно основного.
        End If
        facade.RebarLabel "Circle", path.ID, 0#, 1#, axisDistance                       ' Объявляем общую подпись количества и диаметров группы.
    End If

    ' Аннотации: размер и подпись фактической арматуры.
    facade.VerticalExtent "DIM_CIRCLE_D", "CONTOUR_CIRCLE", "D", -diameter / 2#, -1#, "Диаметр параметрического круглого сечения."  ' Добавляем размер диаметра бетонного круга.
    Set ISectionRecipe_CreateDefinition = facade.Finish                                 ' Проверяем и возвращаем готовое описание сечения.
End Function

' Описывает смещение и привязку дополнительного ряда относительно основного, сохраняя смысловую
' группу.
Private Function AdditionalOptions(ByVal table As CSectionInputTable, ByVal key As String, ByVal suffix As String, _
        ByVal location As String, ByVal distance As Double) As CRebarRuleOptions
    Dim options As CRebarRuleOptions                                                    ' Объявляем условия размещения арматуры.
    Set options = New CRebarRuleOptions                                                 ' Создаём условия размещения арматуры.
    If location = "Stacked" Then options.NormalOffset = distance Else options.TangentOffset = -distance  ' Выбираем смещение поперёк ряда или вдоль него.
    options.SourceSuffix = suffix: options.AnnotationGroup = "Circle"                   ' Задаём суффикс ряда и общую группу подписи.
    options.Publication = PublishPerSourcePosition: options.PlacementCheck = CheckFullDisk  ' Сохраняем исходные имена; проверяем весь стержень.
    options.Comment = "Круг: дополнительный ряд арматуры"                               ' Задаём пояснение к стержням этого ряда.
    options.SourceReference = table.SourceReference(table.RowByKey(key), 2)             ' Привязываем ошибки размещения к ячейкам ввода.
    Set AdditionalOptions = options                                                     ' Возвращаем условия размещения дополнительного ряда.
End Function

' Проверяет ограничение рецепта; ошибка содержит фактический адрес параметра во входной таблице.
Private Sub Require(ByVal table As CSectionInputTable, ByVal key As String, ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then SectionRecipeError "CSectionRecipeCircle", reason & " Исправьте указанное значение.", table.SourceReference(table.RowByKey(key), 2)
End Sub
