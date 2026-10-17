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
    StartSetback As Double ' Отступ от верхней точки окружности по часовой стрелке вдоль линии центров, мм.
    EndSetback As Double ' Отступ назад от конца полного оборота вдоль линии центров, мм.
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
    Dim reader As CSectionInputReader, facade As CSectionDefinitionBuilder
    Dim diameter As Double, rebar As TCircleRebar
    Set reader = context.Reader("rngCircleGeometry") ' Находим поля таблицы по ключам и используем единицы ввода.
    diameter = reader.PositiveLength("Geom", "D")
    ReadRebarInputs reader, diameter, rebar ' Читаем только параметры включённых рядов.
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "Circle" ' Создаём сборщик описания сечения.
    AddGeometry facade, diameter ' Добавляем круговую границу бетона.
    AddReinforcement facade, diameter, rebar ' Создаём линии арматуры и задаём правила размещения стержней.
    AddAnnotations facade, diameter, rebar ' Добавляем размер диаметра и общую подпись арматуры.
    Set ISectionRecipe_CreateDefinition = facade.Finish ' Завершаем и возвращаем проверенное описание сечения.
End Function

' Читает число и диаметры стержней, а для включённых рядов — отступ и привязки.
' Если основного ряда нет, его отступ и параметры дополнительных рядов не требуются.
Private Sub ReadRebarInputs(ByVal reader As CSectionInputReader, ByVal diameter As Double, ByRef rebar As TCircleRebar)
    Dim rowNumber As Long, suffix As String
    rebar.Count = reader.Count("Main", "Circle", "n")
    rebar.Diameter = reader.OptionalDiameter("Main", "Circle", "d")
    If rebar.Count = 0 Or rebar.Diameter = 0# Then Exit Sub
    reader.Require "Main", "Circle", "n", rebar.Count >= 2, "Для круговой раскладки требуется не меньше двух стержней первого ряда."
    rebar.AxisDistance = reader.Length("Main", "Circle", "as")
    reader.Require "Main", "Circle", "as", rebar.AxisDistance > rebar.Diameter / 2#, "Расстояние до оси должно быть больше половины диаметра стержня."
    reader.Require "Main", "Circle", "as", rebar.AxisDistance < diameter / 2#, "Расстояние до оси должно быть меньше радиуса бетонного круга."
    rebar.StartSetback = reader.NonNegativeLength("Main", "Circle", "s0")
    rebar.EndSetback = reader.NonNegativeLength("Main", "Circle", "s1")
    reader.Require "Main", "Circle", "s0", rebar.StartSetback + rebar.EndSetback < 2# * GEOM_PI * (diameter / 2# - rebar.AxisDistance) - GEOM_TOLERANCE, "Сумма концевых отступов должна быть меньше длины окружности центров стержней."
    rebar.Source = reader.SourceReferences("Main", "Circle", Array("as", "d", "n", "s0", "s1"))
    For rowNumber = 2 To 3
        suffix = CStr(rowNumber)
        rebar.Diameters(rowNumber) = reader.OptionalDiameter("Extra", "Circle", "d" & suffix)
        If rebar.Diameters(rowNumber) > 0# Then
            rebar.BaseOffsets(rowNumber) = reader.OptionalRowOffset("Extra", "Circle", "a" & suffix)
            rebar.Bindings(rowNumber) = reader.Choice("Extra", "Circle", Array("EachBar", "EverySecondBar"), "bind" & suffix)
            rebar.Locations(rowNumber) = reader.Choice("Extra", "Circle", Array("Stacked", "SideBySide"), "pos" & suffix)
            rebar.ExtraSources(rowNumber) = reader.SourceReferences("Extra", "Circle", Array("d" & suffix, "a" & suffix, "bind" & suffix, "pos" & suffix))
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
    Dim contact As CRebarContactBuilder, extraOptions As CRebarRuleOptions
    Dim path As CGeometryPath, options As CRebarRuleOptions, rowNumber As Long
    If rebar.Count = 0 Or rebar.Diameter = 0# Then Exit Sub
    If rebar.StartSetback = 0# And rebar.EndSetback = 0# Then
        Set path = facade.ArcPath("Circle.Axis", 0#, 0#, diameter / 2# - rebar.AxisDistance, GEOM_PI / 2#, -2# * GEOM_PI) ' Полный круг центров от верхней точки по часовой стрелке.
    Else
        Set path = facade.ArcPath("Circle.FullAxis", 0#, 0#, diameter / 2# - rebar.AxisDistance, GEOM_PI / 2#, -2# * GEOM_PI)
        Set path = facade.Slice("Circle.Axis", path.ID, rebar.StartSetback, path.Length - rebar.EndSetback) ' Открытая дуга после отсечения двух концевых отступов.
    End If
    Set options = New CRebarRuleOptions ' Создаём настройки размещения этого ряда.
    If path.IsClosed Then options.EndMode = ClosedUniform Else options.EndMode = IncludeBoth ' Полный круг без повторного конца либо дуга с обоими концами.
    options.SourcePrefix = "B" ' Имена стержней начинаются с B.
    options.AnnotationGroup = "Circle" ' В одной подписи указываем суммарное количество стержней каждого диаметра во всех рядах этой окружности.
    options.Publication = PublishPerSourcePosition ' Нумеруем R1, R2, ...: за каждым основным стержнем идут связанные с ним стержни дополнительных рядов.
    options.Coincidence = RejectCoincidence ' Совпадение стержней считаем ошибкой.
    options.Comment = "Круг: основной ряд арматуры"
    options.SourceReference = rebar.Source ' Адреса полей для сообщения об ошибке размещения.
    facade.RebarByCount "Circle.Main", path.ID, rebar.Count, rebar.Diameter, options ' Задаём основной ряд по числу стержней и диаметру.
    Set extraOptions = New CRebarRuleOptions ' Общие настройки только дополнительных рядов.
    extraOptions.AnnotationGroup = "Circle" ' Подпись объединяет стержни всех рядов этой окружности.
    extraOptions.Publication = PublishPerSourcePosition ' Нумеруем R1, R2, ...: за каждым основным стержнем идут связанные с ним стержни дополнительных рядов.
    extraOptions.Coincidence = RejectCoincidence ' Совпадение стержней считаем ошибкой.
    extraOptions.PlacementCheck = CheckFullDisk ' Весь круг стержня должен находиться в бетоне.
    Set contact = New CRebarContactBuilder
    contact.Initialize rebar.Diameter, -1#, 1#, extraOptions ' Stacked направлен внутрь круга, SideBySide — по часовой касательной.
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
