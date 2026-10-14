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
' Круглое бетонное сечение с арматурой
' ==========================================================================
' Класс читает диаметр сечения и параметры армирования из таблицы круглого сечения.
' Основной ряд арматуры располагается по окружности на расстоянии as от края бетона. Второй и
' третий ряды привязаны к стержням первого ряда и могут находиться рядом с ними или глубже.
'
' Рецепт описывает границу бетона, расположение арматуры и аннотации схемы
' (размер диаметра сечения, количество и диаметры стержней). По этому описанию общий
' CSectionBuilder создаёт стержни и бетонную сетку; сам рецепт их не расставляет.

Implements ISectionRecipe

' Читает диаметр сечения и параметры включённых рядов арматуры. Формирует описание
' круга бетона, расположения стержней и аннотаций схемы (размер диаметра сечения,
' суммарное количество стержней каждого диаметра) для общего сборщика.
Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
    Dim table As CSectionInputTable, units As CUnitSystem, facade As CSectionDefinitionBuilder
    Dim diameter As Double, axisDistance As Double, count As Long, d As Double, d2 As Double, d3 As Double
    Dim loc2 As String, loc3 As String, options As CRebarRuleOptions, path As CGeometryPath, skipped As Double

    Set table = context.Table("rngCircleGeometry"): Set units = context.Units ' Берём копию входной таблицы и единицы измерения.
    diameter = table.PositiveLength(table.RowByKey("Circle.Diameter"), 2, units)
    count = table.Count(table.RowByKey("Rebar.Count"), 2)
    d = table.OptionalDiameter(table.RowByKey("Rebar.Diameter"), 2, units)

    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Circle" ' Создаём сборщик описания сечения.
    Set path = facade.CirclePath("CONTOUR_CIRCLE", 0#, 0#, diameter / 2#) ' Круг бетона с центром (0, 0) и радиусом D/2, мм.
    facade.Outer "OUTER", path.ID, "Круговой контур сечения." ' Назначаем эту замкнутую линию наружной границей бетона.
    If count > 0 And d > 0# Then
        Require table, "Rebar.Count", count >= 2, "Для круговой раскладки требуется не меньше двух стержней первого ряда."
        axisDistance = table.Length(table.RowByKey("Rebar.AxisDistance"), 2, units)
        Require table, "Rebar.AxisDistance", axisDistance > d / 2#, "Расстояние до оси должно быть больше половины диаметра стержня."
        Require table, "Rebar.AxisDistance", axisDistance < diameter / 2#, "Расстояние до оси должно быть меньше радиуса бетонного круга."
        d2 = table.OptionalDiameter(table.RowByKey("Rebar.Diameter2"), 2, units)
        d3 = table.OptionalDiameter(table.RowByKey("Rebar.Diameter3"), 2, units)
        If d2 > 0# Then loc2 = table.Choice(table.RowByKey("Rebar.Loc2row"), 2, Array("Stacked", "SideBySide"))
        If d3 > 0# Then loc3 = table.Choice(table.RowByKey("Rebar.Loc3row"), 2, Array("Stacked", "SideBySide"))
        Set path = facade.CirclePath("Circle.Axis", 0#, 0#, diameter / 2# - axisDistance) ' Линия центров основного ряда: радиус D/2 - as, мм.

        Set options = New CRebarRuleOptions ' Создаём настройки размещения этого ряда.
        options.EndMode = ClosedUniform: options.SourcePrefix = "B" ' Равномерно по кругу без повторного конца; имена начинаются с B.
        options.AnnotationGroup = "Circle" ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой окружности.
        options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
        options.Coincidence = RejectCoincidence ' Совпадение стержней считаем ошибкой.
        options.Comment = "Круг: основной ряд арматуры"
        ' Адреса полей для сообщения об ошибке размещения.
        options.SourceReference = table.SourceReference(table.RowByKey("Rebar.Count"), 2) & "; " & table.SourceReference(table.RowByKey("Rebar.AxisDistance"), 2)
        facade.RebarByCount "Circle.Main", path.ID, count, d, options ' Задаём основной ряд по числу стержней и диаметру.
        If d2 > 0# Then
            Set options = AdditionalOptions(table, "Rebar.Diameter2", "Rebar.Bind2row", "_R2", loc2, d, d2) ' Смещение и выбор основных стержней для второго ряда.
            facade.RebarDependent "Circle.Row2", "Circle.Main", d2, options ' Второй ряд привязан к позициям основного; задаём его диаметр.
        End If
        If d3 > 0# Then
            If d2 > 0# And loc2 = loc3 Then skipped = d2 ' Третий ряд в том же направлении должен обойти стержень второго.
            Set options = AdditionalOptions(table, "Rebar.Diameter3", "Rebar.Bind3row", "_R3", loc3, d, d3, skipped) ' Смещение и выбор основных стержней для третьего ряда.
            facade.RebarDependent "Circle.Row3", "Circle.Main", d3, options ' Третий ряд привязан к позициям основного; задаём его диаметр.
        End If
        facade.RebarLabel "Circle", path.ID, 0#, 1# ' Общая подпись всех круговых рядов над сечением.
    End If

    facade.VerticalExtent "DIM_CIRCLE_D", "CONTOUR_CIRCLE", "D", -diameter / 2#, -1#, "Диаметр параметрического круглого сечения." ' Диаметр бетонного круга; размер вынесен влево.
    Set ISectionRecipe_CreateDefinition = facade.Finish ' Завершаем и возвращаем проверенное описание сечения.
End Function

' Возвращает настройки дополнительного ряда арматуры: направление и расстояние от основного
' стержня, суффикс имени и общую группу подписи. При наличии второго ряда учитывает его диаметр.
Private Function AdditionalOptions(ByVal table As CSectionInputTable, ByVal key As String, ByVal bindingKey As String, ByVal suffix As String, _
        ByVal location As String, ByVal baseDiameter As Double, ByVal diameter As Double, _
        Optional ByVal interveningDiameterSum As Double = 0#) As CRebarRuleOptions
    Dim options As CRebarRuleOptions
    Set options = New CRebarRuleOptions ' Создаём настройки размещения этого ряда.
    Dim axis As ERebarOffsetAxis, directionSign As Double, everySecond As Boolean
    axis = RowNormal: directionSign = 1# ' По умолчанию смещаем ряд по нормали к исходной линии.
    If location = "SideBySide" Then axis = RowTangent: directionSign = -1# ' Располагаем дополнительный стержень вдоль исходного ряда.
    everySecond = (table.Choice(table.RowByKey(bindingKey), 2, Array("EachBar", "EverySecondBar")) = "EverySecondBar") ' Выбираем все позиции или через одну, начиная с первой.
    options.SourceReference = table.SourceReference(table.RowByKey(key), 2) & "; " & table.SourceReference(table.RowByKey(bindingKey), 2) ' Адреса диаметра и привязки ряда.
    options.ConfigureDependentContact baseDiameter, diameter, axis, directionSign, everySecond, interveningDiameterSum ' Учитываем диаметры рядов и выбор основных стержней.
    options.SourceSuffix = suffix ' Добавляем к имени основного стержня суффикс ряда.
    options.AnnotationGroup = "Circle" ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой окружности.
    options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
    options.Coincidence = RejectCoincidence ' Совпадение стержней считаем ошибкой.
    options.PlacementCheck = CheckFullDisk ' Весь круг стержня должен находиться в бетоне.
    options.Comment = "Круг: дополнительный ряд арматуры"
    Set AdditionalOptions = options ' Возвращаем настройки дополнительного ряда.
End Function

' Проверяет условие для указанного параметра. При нарушении сообщает причину и адрес его ячейки.
Private Sub Require(ByVal table As CSectionInputTable, ByVal key As String, ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then SectionRecipeError "CSectionRecipeCircle", reason & " Исправьте указанное значение.", table.SourceReference(table.RowByKey(key), 2)
End Sub
