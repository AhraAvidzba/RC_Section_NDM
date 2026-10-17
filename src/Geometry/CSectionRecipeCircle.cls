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

Private Type TCircleRebar
    Count As Long ' Число стержней основного ряда.
    Diameter As Double ' Диаметр стержней основного ряда, мм.
    AxisDistance As Double ' Расстояние от края бетона до осей основного ряда, мм.
    BaseOffsets(2 To 3) As Double ' Базовые отступы второго и третьего рядов от основного, мм.
    Diameters(2 To 3) As Double ' Диаметры второго и третьего рядов арматуры, мм.
    Locations(2 To 3) As String ' Прочитанное расположение каждого дополнительного ряда относительно основного.
    Bindings(2 To 3) As String ' Прочитанный способ привязки дополнительных рядов к основному.
    ExtraSources(2 To 3) As String ' Адреса исходных полей дополнительных рядов для сообщения об ошибке.
    Source As String ' Адреса числа стержней и отступа основного ряда для сообщения об ошибке.
End Type

' Читает диаметр сечения и параметры включённых рядов арматуры. Формирует описание
' круга бетона, расположения стержней и аннотаций схемы (размер диаметра сечения,
' суммарное количество стержней каждого диаметра) для общего сборщика.
Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
    Dim table As CSectionInputTable, units As CUnitSystem, facade As CSectionDefinitionBuilder
    Dim diameter As Double, rebar As TCircleRebar
    Set table = context.Table("rngCircleGeometry"): Set units = context.Units ' Берём копию входной таблицы и единицы измерения.
    diameter = table.PositiveLength(table.RowByKey("Circle.Diameter"), 2, units)
    ReadRebarInputs table, units, diameter, rebar ' Читаем только параметры включённых рядов.
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Circle" ' Создаём сборщик описания сечения.
    AddGeometry facade, diameter ' Добавляем круговую границу бетона.
    AddReinforcement facade, diameter, rebar ' Создаём линии арматуры и задаём правила размещения стержней.
    AddAnnotations facade, diameter, rebar ' Добавляем размер диаметра и общую подпись арматуры.
    Set ISectionRecipe_CreateDefinition = facade.Finish ' Завершаем и возвращаем проверенное описание сечения.
End Function

' Читает число и диаметры стержней, а для включённых рядов — отступ и привязки.
' Если основного ряда нет, его отступ и параметры дополнительных рядов не требуются.
Private Sub ReadRebarInputs(ByVal table As CSectionInputTable, ByVal units As CUnitSystem, ByVal diameter As Double, ByRef rebar As TCircleRebar)
    Dim countRow As Long, axisRow As Long, rowNumber As Long, header As Long, extra As Long, column As Long
    header = table.AdditionalRowsHeader: extra = header + 1
    Require table, extra, StrComp(table.Text(extra, 1), "Круг", vbTextCompare) = 0, "В таблице дополнительных рядов ожидается строка Круг."
    countRow = table.RowByKey("Rebar.Count")
    rebar.Count = table.Count(countRow, 2)
    rebar.Diameter = table.OptionalDiameter(table.RowByKey("Rebar.Diameter"), 2, units)
    If rebar.Count = 0 Or rebar.Diameter = 0# Then Exit Sub
    Require table, countRow, rebar.Count >= 2, "Для круговой раскладки требуется не меньше двух стержней первого ряда."
    axisRow = table.RowByKey("Rebar.AxisDistance")
    rebar.AxisDistance = table.Length(axisRow, 2, units)
    Require table, axisRow, rebar.AxisDistance > rebar.Diameter / 2#, "Расстояние до оси должно быть больше половины диаметра стержня."
    Require table, axisRow, rebar.AxisDistance < diameter / 2#, "Расстояние до оси должно быть меньше радиуса бетонного круга."
    rebar.Source = table.SourceReference(countRow, 2) & "; " & table.SourceReference(axisRow, 2)
    For rowNumber = 2 To 3
        column = 2 + (rowNumber - 2) * 4
        rebar.Diameters(rowNumber) = table.OptionalDiameter(extra, column, units)
        If rebar.Diameters(rowNumber) > 0# Then
            rebar.BaseOffsets(rowNumber) = table.OptionalRowOffset(extra, column + 1, units)
            rebar.Bindings(rowNumber) = table.Choice(extra, column + 2, Array("EachBar", "EverySecondBar"))
            rebar.Locations(rowNumber) = table.Choice(extra, column + 3, Array("Stacked", "SideBySide"))
            rebar.ExtraSources(rowNumber) = table.SourceReference(extra, column) & "; " & table.SourceReference(extra, column + 1) & "; " & table.SourceReference(extra, column + 2) & "; " & table.SourceReference(extra, column + 3)
        End If
    Next rowNumber
End Sub

' Создаёт окружность бетона и назначает её наружной границей.
Private Sub AddGeometry(ByVal facade As CSectionDefinitionBuilder, ByVal diameter As Double)
    Dim path As CGeometryPath
    Set path = facade.CirclePath("CONTOUR_CIRCLE", 0#, 0#, diameter / 2#) ' Круг бетона с центром (0, 0) и радиусом D/2, мм.
    facade.Outer "OUTER", path.ID, "Круговой контур сечения." ' Назначаем эту замкнутую линию наружной границей бетона.
End Sub

' Создаёт окружность центров основного ряда и правила размещения всех включённых рядов.
Private Sub AddReinforcement(ByVal facade As CSectionDefinitionBuilder, ByVal diameter As Double, ByRef rebar As TCircleRebar)
    Dim contact As CRebarContactBuilder, defaults As CRebarRuleOptions
    Dim path As CGeometryPath, options As CRebarRuleOptions, rowNumber As Long
    If rebar.Count = 0 Or rebar.Diameter = 0# Then Exit Sub
    Set path = facade.CirclePath("Circle.Axis", 0#, 0#, diameter / 2# - rebar.AxisDistance) ' Линия центров основного ряда: радиус D/2 - as, мм.
    Set options = New CRebarRuleOptions ' Создаём настройки размещения этого ряда.
    options.EndMode = ClosedUniform: options.SourcePrefix = "B" ' Равномерно по кругу без повторного конца; имена начинаются с B.
    options.AnnotationGroup = "Circle" ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой окружности.
    options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
    options.Coincidence = RejectCoincidence ' Совпадение стержней считаем ошибкой.
    options.Comment = "Круг: основной ряд арматуры"
    options.SourceReference = rebar.Source ' Адреса полей для сообщения об ошибке размещения.
    facade.RebarByCount "Circle.Main", path.ID, rebar.Count, rebar.Diameter, options ' Задаём основной ряд по числу стержней и диаметру.
    Set defaults = New CRebarRuleOptions ' Общие настройки только дополнительных рядов.
    defaults.AnnotationGroup = "Circle" ' Подпись объединяет стержни всех рядов этой окружности.
    defaults.Publication = PublishPerSourcePosition ' Сначала основной стержень, затем связанные с ним.
    defaults.Coincidence = RejectCoincidence ' Совпадение стержней считаем ошибкой.
    defaults.PlacementCheck = CheckFullDisk ' Весь круг стержня должен находиться в бетоне.
    Set contact = New CRebarContactBuilder
    contact.Initialize rebar.Diameter, 1#, -1#, defaults ' Stacked направлен внутрь круга, SideBySide — против обхода линии.
    For rowNumber = 2 To 3
        If rebar.Diameters(rowNumber) > 0# Then
            Set options = contact.AppendRow(rebar.Diameters(rowNumber), rebar.Locations(rowNumber), rebar.Bindings(rowNumber), rebar.ExtraSources(rowNumber), rebar.BaseOffsets(rowNumber))
            options.SourceSuffix = "_R" & CStr(rowNumber) ' Добавляем к имени основного стержня суффикс ряда.
            options.Comment = "Круг: дополнительный ряд арматуры"
            facade.RebarDependent "Circle.Row" & CStr(rowNumber), "Circle.Main", rebar.Diameters(rowNumber), options ' Оба дополнительных ряда связаны с основным.
        End If
    Next rowNumber
End Sub

' Добавляет размер диаметра бетона и, при наличии арматуры, общую подпись всех круговых рядов.
Private Sub AddAnnotations(ByVal facade As CSectionDefinitionBuilder, ByVal diameter As Double, ByRef rebar As TCircleRebar)
    If rebar.Count > 0 And rebar.Diameter > 0# Then facade.RebarLabel "Circle", "Circle.Axis", 0#, 1# ' Общая подпись всех круговых рядов над сечением.
    facade.VerticalExtent "DIM_CIRCLE_D", "CONTOUR_CIRCLE", "D", -diameter / 2#, -1#, "Диаметр параметрического круглого сечения." ' Диаметр бетонного круга; размер вынесен влево.
End Sub

' Проверяет условие для указанного параметра. При нарушении сообщает причину и адрес его ячейки.
Private Sub Require(ByVal table As CSectionInputTable, ByVal row As Long, ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then SectionRecipeError "CSectionRecipeCircle", reason & " Исправьте указанное значение.", table.SourceReference(row, 2)
End Sub
