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
    BottomWidth As Double ' Расстояние между нижними вершинами до скругления, мм.
    TopWidth As Double ' Расстояние между верхними вершинами до скругления, мм.
    Height As Double ' Высота трапеции, мм.
    CornerRadius As Double ' Радиус скругления вершин трапеции, мм.
    OpeningRadius As Double ' Радиус круглого отверстия, мм.
    OuterAxisDistance As Double ' Расстояние от наружной границы бетона до центров арматуры, мм.
    InnerAxisDistance As Double ' Расстояние от границы отверстия до центров арматуры, мм.
    OuterCount As Long ' Число стержней по наружной границе.
    InnerCount As Long ' Число стержней вокруг отверстия.
    OuterDiameter As Double ' Диаметр наружной арматуры, мм.
    InnerDiameter As Double ' Диаметр арматуры вокруг отверстия, мм.
    Source As String ' Адрес исходных данных для пояснения ошибок.
End Type

' Описывает границы бетона и отверстия, армирование и аннотации (размеры, подписи арматуры и текст).
' Сетку и отдельные стержни создаёт общий CSectionBuilder.
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
    Set ISectionRecipe_CreateDefinition = Describe(parameters, table) ' Таблица уточняет адреса ошибок; геометрия и армирование описываются общим методом.
End Function

' Задаёт рецепт программно: размеры и радиусы передаются в мм, параметры арматуры заданы внутри метода.
' Config не читается. При изменении исходных параметров формируется новое описание сечения.
Public Function ProgramDefinition(Optional ByVal bottomWidth As Double = 1200#, Optional ByVal topWidth As Double = 800#, _
        Optional ByVal height As Double = 800#, Optional ByVal cornerRadius As Double = 60#, Optional ByVal openingRadius As Double = 110#) As CSectionDefinition
    Dim parameters As TInputs
    parameters.BottomWidth = bottomWidth: parameters.TopWidth = topWidth: parameters.Height = height
    parameters.CornerRadius = cornerRadius: parameters.OpeningRadius = openingRadius
    parameters.OuterAxisDistance = 50#: parameters.InnerAxisDistance = 45#
    parameters.OuterCount = 24: parameters.InnerCount = 16: parameters.OuterDiameter = 20#: parameters.InnerDiameter = 16#
    parameters.Source = "ProgramInputs.BottomWidth/TopWidth/Height/CornerRadius/OpeningRadius"
    Set ProgramDefinition = Describe(parameters, Nothing) ' Nothing означает программный ввод без адресов ячеек Excel.
End Function

' Читает положительный размер из снимка таблицы и переводит его в мм через CUnitSystem.
Private Function ReadLength(ByVal table As CSectionInputTable, ByVal key As String, ByVal units As CUnitSystem) As Double
    ReadLength = table.Length(table.RowByKey(key), 2, units)
    If ReadLength <= 0# Then SectionRecipeError "CSectionRecipeRoundedTrapezoid", "Размер должен быть положительным. Исправьте указанный параметр.", table.SourceReference(table.RowByKey(key), 2)
End Function

' Создаёт контур трапеции с прямыми участками и скруглениями, а отверстие задаёт отдельной окружностью.
' Наружная арматура расположена внутри трапеции, внутренняя — снаружи отверстия, в бетоне.
Private Function Describe(ByRef parameters As TInputs, ByVal table As CSectionInputTable) As CSectionDefinition
    On Error GoTo InvalidGeometry
    If parameters.BottomWidth <= 0# Or parameters.TopWidth <= 0# Or parameters.Height <= 0# Or parameters.CornerRadius <= 0# Or parameters.OpeningRadius <= 0# Then _
        SectionRecipeError "CSectionRecipeRoundedTrapezoid", "Размеры и радиусы должны быть положительными.", parameters.Source
    Dim facade As CSectionDefinitionBuilder, boundary As CGeometryPath, openingPath As CGeometryPath, axis As CGeometryPath ' boundary — наружный бетонный контур; openingPath — отверстие; axis — линия центров стержней.
    Dim polygon As CPolygonSpec, options As CRebarRuleOptions
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "RoundedTrapezoid"
    Set polygon = New CPolygonSpec ' Один описатель вершин служит для бетонной границы и её смещённой арматурной линии.
    polygon.AddVertex "BottomLeft", -parameters.BottomWidth / 2#, -parameters.Height / 2#, parameters.CornerRadius, parameters.Source
    polygon.AddVertex "BottomRight", parameters.BottomWidth / 2#, -parameters.Height / 2#, parameters.CornerRadius, parameters.Source
    polygon.AddVertex "TopRight", parameters.TopWidth / 2#, parameters.Height / 2#, parameters.CornerRadius, parameters.Source
    polygon.AddVertex "TopLeft", -parameters.TopWidth / 2#, parameters.Height / 2#, parameters.CornerRadius, parameters.Source
    polygon.Seal
    facade.Source "Boundary", parameters.Source ' Источник ошибки границы связывается с переданными параметрами.
    Set boundary = facade.RoundedPolygon("Boundary", polygon) ' Скругления строятся по завершённому описателю polygon; исходные размеры заданы в мм.
    facade.Outer "Concrete", boundary.ID, "Наружная граница скругленной трапеции." ' Concrete — имя бетонной области; boundary.ID — её замкнутая наружная граница.
    Set openingPath = facade.CirclePath("Opening.Circle", 0#, 0#, parameters.OpeningRadius) ' Окружность отверстия с центром в начале координат; радиус задан в мм.
    facade.Opening "Opening", openingPath.ID, "Concrete", "Круглое отверстие." ' Opening — имя отверстия, которое вычитается из области Concrete.
    Set axis = facade.OffsetRoundedPolygon("Outer.Axis", polygon, parameters.OuterAxisDistance) ' Положительное смещение направлено внутрь трапеции; стержни располагаются по центрам на этой линии.
    Set options = New CRebarRuleOptions
    options.EndMode = IncludeStart: options.PlacementCheck = CheckFullDisk: options.OverlapPolicy = RejectOverlap ' Замкнутая раскладка без повторения конца; весь диск каждого стержня должен лежать в бетоне и не пересекать другие.
    options.AnnotationGroup = "OuterBars": options.SourcePrefix = "Outer." ' Группа общей подписи и префикс исходных имён наружных стержней.
    options.SourceReference = parameters.Source ' Общий источник для программного ввода; при наличии таблицы ниже выбираются конкретные связанные поля.
    If Not table Is Nothing Then options.SourceReference = table.SourceReference(table.RowByKey("OuterAxisDistance"), 2) & "; " & table.SourceReference(table.RowByKey("OuterCount"), 2) ' Ошибка наружного ряда укажет его отступ и количество в реальной таблице.
    facade.RebarByCount "OuterBars", axis.ID, parameters.OuterCount, parameters.OuterDiameter, options ' Количество и диаметр относятся к наружному ряду; параметры размещения копируются из options.
    Set axis = facade.Offset("Inner.Axis", openingPath.ID, -parameters.InnerAxisDistance) ' Окружность идёт против часовой стрелки; отрицательное левое смещение направлено вправо от обхода, из отверстия в бетон.
    Set options = New CRebarRuleOptions
    options.EndMode = IncludeStart: options.PlacementCheck = CheckFullDisk: options.OverlapPolicy = RejectOverlap ' Замкнутая раскладка без повторения конца; весь диск каждого стержня должен лежать в бетоне и не пересекать другие.
    options.AnnotationGroup = "InnerBars": options.SourcePrefix = "Inner." ' Внутренние стержни получают отдельные группу подписи и префикс имён.
    options.SourceReference = parameters.Source ' Общий источник для программного ввода; при наличии таблицы ниже выбираются конкретные связанные поля.
    If Not table Is Nothing Then options.SourceReference = table.SourceReference(table.RowByKey("InnerAxisDistance"), 2) & "; " & table.SourceReference(table.RowByKey("InnerCount"), 2) ' Ошибка внутреннего ряда укажет его отступ и количество в реальной таблице.
    facade.RebarByCount "InnerBars", axis.ID, parameters.InnerCount, parameters.InnerDiameter, options ' Внутренний ряд имеет собственные количество, диаметр и условия размещения.
    facade.VerticalExtent "Height", boundary.ID, "H", -parameters.BottomWidth / 2#, -1# ' Высота бетонного контура; исходная размерная линия у левого края, знак -1 выносит её влево.
    facade.HorizontalExtent "Width", boundary.ID, "B", -parameters.Height / 2#, -1# ' Ширина бетонного контура; исходная размерная линия у нижнего края, знак -1 выносит её вниз.
    facade.Radius "OpeningRadius", openingPath.ID, openingPath.ID, 0.25, "Ro" ' Оба ID относятся к окружности и её сегменту; 0.25 задаёт точку на четверти обхода для размера Ro.
    facade.Anchor "Caption", -parameters.TopWidth / 4#, parameters.Height / 5# ' Программная точка размещения поясняющего текста, мм.
    facade.Text "FormName", "Caption", "Скругленная трапеция"
    facade.RebarLabel "OuterBars", "Outer.Axis", 0#, 1# ' Подпись наружной группы привязана к её осевому контуру и направлена вверх.
    facade.RebarLabel "InnerBars", "Inner.Axis", 0#, -1# ' Подпись внутренней группы привязана к её осевому контуру и направлена вниз.
    Set Describe = facade.Finish ' Общий фасад проверяет согласованность и возвращает неизменяемое описание.
    Exit Function
InvalidGeometry:
    SectionRecipeGeometryError "CSectionRecipeRoundedTrapezoid", parameters.Source
End Function
