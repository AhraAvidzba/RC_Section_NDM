VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionRecipeRoundedTrapezoid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Учебный рецепт скруглённой трапеции через общий декларативный фасад
' ==========================================================================
Implements ISectionRecipe

' Рабочий пример нового типа: скругленная трапеция с круглым отверстием.
' Здесь только ввод и декларации. Сетка, стержни, измерения и НДС общие.
Private Type TInputs
    BottomWidth As Double
    TopWidth As Double
    Height As Double
    CornerRadius As Double
    OpeningRadius As Double
    OuterAxisDistance As Double
    InnerAxisDistance As Double
    OuterCount As Long
    InnerCount As Long
    OuterDiameter As Double
    InnerDiameter As Double
    Source As String
End Type

' Описывает контуры, арматурные правила и размеры выбранного рецепта. Сетку и физические стержни
' строит общий CSectionBuilder.
Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
    Dim table As CSectionInputTable, units As CUnitSystem, parameters As TInputs
    Set table = context.Table("rngRoundedTrapezoidGeometry"): Set units = context.Units
    parameters.BottomWidth = ReadLength(table, "BottomWidth", units)
    parameters.TopWidth = ReadLength(table, "TopWidth", units)
    parameters.Height = ReadLength(table, "Height", units)
    parameters.CornerRadius = ReadLength(table, "CornerRadius", units)
    parameters.OpeningRadius = ReadLength(table, "OpeningRadius", units)
    parameters.OuterAxisDistance = ReadLength(table, "OuterAxisDistance", units)
    parameters.InnerAxisDistance = ReadLength(table, "InnerAxisDistance", units)
    parameters.OuterDiameter = ReadLength(table, "OuterDiameter", units)
    parameters.InnerDiameter = ReadLength(table, "InnerDiameter", units)
    parameters.OuterCount = table.Count(table.RowByKey("OuterCount"), 2)
    parameters.InnerCount = table.Count(table.RowByKey("InnerCount"), 2)
    parameters.Source = table.SourceReference(table.RowByKey("BottomWidth"), 2) & "; " & _
        table.SourceReference(table.RowByKey("TopWidth"), 2) & "; " & table.SourceReference(table.RowByKey("Height"), 2) & "; " & _
        table.SourceReference(table.RowByKey("CornerRadius"), 2) & "; " & table.SourceReference(table.RowByKey("OpeningRadius"), 2) & "; " & _
        table.SourceReference(table.RowByKey("OuterAxisDistance"), 2) & "; " & table.SourceReference(table.RowByKey("InnerAxisDistance"), 2)
    Set ISectionRecipe_CreateDefinition = Describe(parameters, table)                   ' Возвращаем описание геометрии, арматуры и размеров.
End Function

' Программный вариант того же рецепта: размеры/диаметры уже во внутренних мм.
' Никакого скрытого чтения Config. Обязательное изменение создает новый snapshot.
Public Function ProgramDefinition(Optional ByVal bottomWidth As Double = 1200#, Optional ByVal topWidth As Double = 800#, _
        Optional ByVal height As Double = 800#, Optional ByVal cornerRadius As Double = 60#, Optional ByVal openingRadius As Double = 110#) As CSectionDefinition
    Dim parameters As TInputs
    parameters.BottomWidth = bottomWidth: parameters.TopWidth = topWidth: parameters.Height = height
    parameters.CornerRadius = cornerRadius: parameters.OpeningRadius = openingRadius
    parameters.OuterAxisDistance = 50#: parameters.InnerAxisDistance = 45#
    parameters.OuterCount = 24: parameters.InnerCount = 16: parameters.OuterDiameter = 20#: parameters.InnerDiameter = 16#
    parameters.Source = "ProgramInputs.BottomWidth/TopWidth/Height/CornerRadius/OpeningRadius"
    Set ProgramDefinition = Describe(parameters, Nothing)                               ' Возвращаем описание геометрии, арматуры и размеров.
End Function

' Читает размер из снимка таблицы и переводит во внутренние мм через CUnitSystem.
Private Function ReadLength(ByVal table As CSectionInputTable, ByVal key As String, ByVal units As CUnitSystem) As Double
    ReadLength = table.Length(table.RowByKey(key), 2, units)
    If ReadLength <= 0# Then SectionRecipeError "CSectionRecipeRoundedTrapezoid", "Размер должен быть положительным. Исправьте указанный параметр.", table.SourceReference(table.RowByKey(key), 2)
End Function

' Геометрическая сторона бетона задается явно: внутрь Outer, наружу Opening.
' Один замкнутый путь охватывает все линии и реальные дуги скруглений.
Private Function Describe(ByRef parameters As TInputs, ByVal table As CSectionInputTable) As CSectionDefinition
    On Error GoTo InvalidGeometry
    If parameters.BottomWidth <= 0# Or parameters.TopWidth <= 0# Or parameters.Height <= 0# Or parameters.CornerRadius <= 0# Or parameters.OpeningRadius <= 0# Then _
        SectionRecipeError "CSectionRecipeRoundedTrapezoid", "Размеры и радиусы должны быть положительными.", parameters.Source
    Dim facade As CSectionDefinitionBuilder, boundary As CGeometryPath, openingPath As CGeometryPath, axis As CGeometryPath  ' Объявляем сборщик описания сечения.
    Dim vertices As Variant, radii As Variant, corners As Variant, options As CRebarRuleOptions  ' Объявляем условия размещения арматуры.
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "RoundedTrapezoid"    ' Создаём описание трапециевидного сечения.
    vertices = Array(Array(-parameters.BottomWidth / 2#, -parameters.Height / 2#), Array(parameters.BottomWidth / 2#, -parameters.Height / 2#), Array(parameters.TopWidth / 2#, parameters.Height / 2#), Array(-parameters.TopWidth / 2#, parameters.Height / 2#))  ' Задаём координаты четырёх вершин трапеции.
    radii = Array(parameters.CornerRadius, parameters.CornerRadius, parameters.CornerRadius, parameters.CornerRadius)  ' Назначаем радиусы скругления вершин.
    corners = Array("BottomLeft", "BottomRight", "TopRight", "TopLeft")                 ' Назначаем имена четырём вершинам трапеции.
    facade.Source "Boundary", parameters.Source                                         ' Привязываем геометрию к ячейкам исходных данных.
    Set boundary = facade.RoundedPolygon("Boundary", vertices, radii, corners)          ' Создаём замкнутый контур с заданными скруглениями.
    facade.Outer "Concrete", boundary.ID, "Наружная граница скругленной трапеции."      ' Назначаем наружную границу бетона.
    Set openingPath = facade.CirclePath("Opening.Circle", 0#, 0#, parameters.OpeningRadius)  ' Создаём окружность отверстия.
    facade.Opening "Opening", openingPath.ID, "Concrete", "Круглое отверстие."          ' Объявляем отверстие внутри наружной границы.
    Set axis = facade.OffsetRoundedPolygon("Outer.Axis", vertices, radii, corners, parameters.OuterAxisDistance)  ' Смещаем контур к линии центров стержней.
    Set options = New CRebarRuleOptions                                                 ' Создаём условия размещения арматуры.
    options.EndMode = ClosedUniform: options.PlacementCheck = CheckFullDisk: options.OverlapPolicy = RejectOverlap  ' Замкнутый ряд; стержни целиком в бетоне, без пересечений.
    options.AnnotationGroup = "OuterBars": options.SourcePrefix = "Outer."              ' Задаём имена стержней и общую группу подписи.
    options.SourceReference = parameters.Source                                         ' Привязываем ошибки размещения к ячейкам ввода.
    If Not table Is Nothing Then options.SourceReference = table.SourceReference(table.RowByKey("OuterAxisDistance"), 2) & "; " & table.SourceReference(table.RowByKey("OuterCount"), 2)  ' Привязываем ошибки размещения к ячейкам ввода.
    facade.RebarByCount "OuterBars", axis.ID, parameters.OuterCount, parameters.OuterDiameter, options  ' Задаём основной ряд по числу и диаметру стержней.
    ' CirclePath имеет обход CCW: отрицательный left-offset выходит из пустоты.
    Set axis = facade.Offset("Inner.Axis", openingPath.ID, -parameters.InnerAxisDistance)  ' Смещаем окружность из отверстия в бетон для арматуры.
    Set options = New CRebarRuleOptions                                                 ' Создаём условия размещения арматуры.
    options.EndMode = ClosedUniform: options.PlacementCheck = CheckFullDisk: options.OverlapPolicy = RejectOverlap  ' Замкнутый ряд; стержни целиком в бетоне, без пересечений.
    options.AnnotationGroup = "InnerBars": options.SourcePrefix = "Inner."              ' Задаём имена стержней и общую группу подписи.
    options.SourceReference = parameters.Source                                         ' Привязываем ошибки размещения к ячейкам ввода.
    If Not table Is Nothing Then options.SourceReference = table.SourceReference(table.RowByKey("InnerAxisDistance"), 2) & "; " & table.SourceReference(table.RowByKey("InnerCount"), 2)  ' Привязываем ошибки размещения к ячейкам ввода.
    facade.RebarByCount "InnerBars", axis.ID, parameters.InnerCount, parameters.InnerDiameter, options  ' Задаём основной ряд по числу и диаметру стержней.
    facade.VerticalExtent "Height", boundary.ID, "H", -parameters.BottomWidth / 2#, -1#  ' Добавляем размер полной высоты сечения.
    facade.HorizontalExtent "Width", boundary.ID, "B", -parameters.Height / 2#, -1#     ' Добавляем размер полной ширины сечения.
    facade.Radius "OpeningRadius", openingPath.ID, openingPath.ID, 0.25, "Ro"           ' Добавляем размер радиуса круглого отверстия.
    facade.Anchor "Caption", -parameters.TopWidth / 4#, parameters.Height / 5#          ' Создаём точку привязки названия сечения.
    facade.Text "FormName", "Caption", "Скругленная трапеция"                           ' Добавляем название сечения на схему.
    facade.RebarLabel "OuterBars", "Outer.Axis", 0#, 1#, parameters.OuterAxisDistance   ' Объявляем общую подпись количества и диаметров группы.
    facade.RebarLabel "InnerBars", "Inner.Axis", 0#, -1#, parameters.InnerAxisDistance  ' Объявляем общую подпись количества и диаметров группы.
    Set Describe = facade.Finish                                                        ' Проверяем и возвращаем готовое описание сечения.
    Exit Function
InvalidGeometry:
    SectionRecipeGeometryError "CSectionRecipeRoundedTrapezoid", parameters.Source
End Function
