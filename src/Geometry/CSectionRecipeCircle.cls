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
    Dim diameter As Double, axisDistance As Double, count As Long, d As Double, skipped As Double
    Dim diameters(2 To 3) As Double, locations(2 To 3) As String, diameterRows(2 To 3) As Long
    Dim countRow As Long, axisRow As Long, rowNumber As Long, bindingRow As Long, binding As String
    Dim options As CRebarRuleOptions, path As CGeometryPath

    Set table = context.Table("rngCircleGeometry"): Set units = context.Units ' Берём копию входной таблицы и единицы измерения.
    diameter = table.PositiveLength(table.RowByKey("Circle.Diameter"), 2, units)
    countRow = table.RowByKey("Rebar.Count")
    count = table.Count(countRow, 2)
    d = table.OptionalDiameter(table.RowByKey("Rebar.Diameter"), 2, units)

    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Circle" ' Создаём сборщик описания сечения.
    Set path = facade.CirclePath("CONTOUR_CIRCLE", 0#, 0#, diameter / 2#) ' Круг бетона с центром (0, 0) и радиусом D/2, мм.
    facade.Outer "OUTER", path.ID, "Круговой контур сечения." ' Назначаем эту замкнутую линию наружной границей бетона.
    If count > 0 And d > 0# Then
        Require table, countRow, count >= 2, "Для круговой раскладки требуется не меньше двух стержней первого ряда."
        axisRow = table.RowByKey("Rebar.AxisDistance")
        axisDistance = table.Length(axisRow, 2, units)
        Require table, axisRow, axisDistance > d / 2#, "Расстояние до оси должно быть больше половины диаметра стержня."
        Require table, axisRow, axisDistance < diameter / 2#, "Расстояние до оси должно быть меньше радиуса бетонного круга."
        For rowNumber = 2 To 3
            diameterRows(rowNumber) = table.RowByKey("Rebar.Diameter" & CStr(rowNumber))
            diameters(rowNumber) = table.OptionalDiameter(diameterRows(rowNumber), 2, units)
        Next rowNumber
        For rowNumber = 2 To 3
            If diameters(rowNumber) > 0# Then locations(rowNumber) = table.Choice(table.RowByKey("Rebar.Loc" & CStr(rowNumber) & "row"), 2, Array("Stacked", "SideBySide"))
        Next rowNumber
        Set path = facade.CirclePath("Circle.Axis", 0#, 0#, diameter / 2# - axisDistance) ' Линия центров основного ряда: радиус D/2 - as, мм.

        Set options = New CRebarRuleOptions ' Создаём настройки размещения этого ряда.
        options.EndMode = ClosedUniform: options.SourcePrefix = "B" ' Равномерно по кругу без повторного конца; имена начинаются с B.
        options.AnnotationGroup = "Circle" ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой окружности.
        options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
        options.Coincidence = RejectCoincidence ' Совпадение стержней считаем ошибкой.
        options.Comment = "Круг: основной ряд арматуры"
        ' Адреса полей для сообщения об ошибке размещения.
        options.SourceReference = table.SourceReference(countRow, 2) & "; " & table.SourceReference(axisRow, 2)
        facade.RebarByCount "Circle.Main", path.ID, count, d, options ' Задаём основной ряд по числу стержней и диаметру.
        For rowNumber = 2 To 3
            If diameters(rowNumber) > 0# Then
                skipped = 0#
                If rowNumber = 3 And locations(2) = locations(3) Then skipped = diameters(2) ' Обходим номинальный второй ряд в том же направлении.
                bindingRow = table.RowByKey("Rebar.Bind" & CStr(rowNumber) & "row")
                binding = table.Choice(bindingRow, 2, Array("EachBar", "EverySecondBar"))
                Set options = RecipeContactOptions(table.SourceReference(diameterRows(rowNumber), 2) & "; " & table.SourceReference(bindingRow, 2), _
                    locations(rowNumber), binding, d, diameters(rowNumber), 1#, -1#, skipped)
                options.SourceSuffix = "_R" & CStr(rowNumber) ' Добавляем к имени основного стержня суффикс ряда.
                options.AnnotationGroup = "Circle" ' Подпись объединяет стержни всех рядов этой окружности.
                options.Publication = PublishPerSourcePosition ' Сначала основной стержень, затем связанные с ним.
                options.Coincidence = RejectCoincidence ' Совпадение стержней считаем ошибкой.
                options.PlacementCheck = CheckFullDisk ' Весь круг стержня должен находиться в бетоне.
                options.Comment = "Круг: дополнительный ряд арматуры"
                facade.RebarDependent "Circle.Row" & CStr(rowNumber), "Circle.Main", diameters(rowNumber), options ' Оба дополнительных ряда связаны с основным.
            End If
        Next rowNumber
        facade.RebarLabel "Circle", path.ID, 0#, 1# ' Общая подпись всех круговых рядов над сечением.
    End If

    facade.VerticalExtent "DIM_CIRCLE_D", "CONTOUR_CIRCLE", "D", -diameter / 2#, -1#, "Диаметр параметрического круглого сечения." ' Диаметр бетонного круга; размер вынесен влево.
    Set ISectionRecipe_CreateDefinition = facade.Finish ' Завершаем и возвращаем проверенное описание сечения.
End Function

' Проверяет условие для указанного параметра. При нарушении сообщает причину и адрес его ячейки.
Private Sub Require(ByVal table As CSectionInputTable, ByVal row As Long, ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then SectionRecipeError "CSectionRecipeCircle", reason & " Исправьте указанное значение.", table.SourceReference(row, 2)
End Sub
