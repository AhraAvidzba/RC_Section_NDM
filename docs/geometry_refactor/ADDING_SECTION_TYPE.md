# Добавление нового типа сечения

## Быстрый старт

Автор нового типа создаёт один `.cls`, реализующий `ISectionRecipe`.
Класс читает собственные исходные данные, объявляет форму/пути/правила через
`CSectionDefinitionBuilder` и возвращает `Finish`. Общая сборка создаёт
`CSectionModel`, сетку, физическую арматуру и аннотации. Решатель, Results и
схема работают с этой моделью независимо от имени типа.

1. Возьмите полный проверенный `examples/CSectionRecipeRoundedTrapezoid.cls`.
   Импортируйте файл через VBE → File → Import File. Имя VBComponent должно
   совпадать с `Attribute VB_Name`; имя класса ограничено правилами VBA.
2. Переименуйте класс и TypeID, задайте собственные параметры/геометрию.
   Не копируйте расчёт координат отдельных стержней или операции Chart.Shapes.
3. В `src/Excel/CSectionTypeRegistry.cls`, `Class_Initialize`, добавьте один
   `RegisterRecipe TypeID, New ClassName, Array(RootRangeName)`.
4. Создайте при необходимости один компактный именованный блок своего ввода.
   Затем явно выполните `modSectionRecipeSetup.UpdateSectionRecipeConfig`:
   штатный список Geometry.Type берётся из реестра. Расчёт не форматирует Config.
5. Скомпилируйте VBProject и выполните сборку/preview/реальное НДС/схему
   на отдельной fixture-книге. Измените параметр и повторите проверки.
   Сохраните/откройте книгу, проверьте выбранный тип и сохранённый Results.

Для примера строка регистрации ровно одна:

```vb
RegisterRecipe "RoundedTrapezoid", New CSectionRecipeRoundedTrapezoid, Array("rngRoundedTrapezoidGeometry")
```

Демонстрационная регистрация используется в отдельной книге проверки.
Она не добавляет пятое обязательное семейство в итоговый пользовательский Config.
Не нужен тип Custom, универсальный редактор примитивов или набор связанных
служебных таблиц. Математически новая общая операция добавляется в общую
библиотеку отдельно; новый тип на существующих операциях её не меняет.

## Архитектура одного рецепта

`ISectionRecipe.CreateDefinition(context)` — единственная точка входа.
`CSectionRecipeContext` содержит снимки таблиц и `CUnitSystem`, без чтения
живого листа в рецепте. `CSectionInputSnapshotReader.Capture` читает только
именованные блоки выбранного типа, вычисляя их формулы даже в manual mode.
`CSectionInputTable` проверяет числа/варианты/счётчики и знает реальные адреса.

Параметры задаются численно в их пользовательских единицах. Длины и диаметры
читайте `table.Length(row, column, context.Units)`; числа без размерности —
`Number`, счётчики — `Count`. `RowByKey` ищет смысловой ключ, поэтому строки
можно переставлять. `ColumnByCaption` позволяет рецепту найти своё поле.
Внутренние длины в мм, площадь в мм², углы в радианах. Переводится ввод один
раз; уже нормализованные мм нельзя переводить повторно.

`SourceReference` даёт адрес исходной ячейки. Передавайте его в
`CRebarRuleOptions.SourceReference` и/или `facade.Source(entityID, reference)`.
Ошибку своего условия выдавайте `SectionRecipeError`, указывая реальную
ссылку и способ исправления. Не разбирайте русский текст чужой ошибки:
общие операции предоставляют машинные Code/ID/Detail1/Detail2.

Фасад хранит декларации. `Finish` проверяет топологию, ссылки и зависимости,
готовит область бетона и публикует неизменяемое definition. `CSectionBuilder.Build`
принимает definition и `TSectionMeshOptions`; параметры сетки читаются общим
workflow. Готовое definition не хранит context/Excel. Новый ввод создаёт новый
snapshot; сохранённая схема или экспорт используют Results, без нового рецепта.

## Полный рабочий пример

Проверенный пример — скруглённая трапеция с круглым отверстием в (0,0).
Имеет четыре прямые грани и четыре настоящие дуги, Outer/Opening,
наружный и внутренний замкнутые пути, 24+16 стержней, размеры и подписи.
Стержни проверяются полными дисками; касание допускается, пересечения
дисков запрещены явно. Смещение направлено в бетон, независимо от центра.

Компактный блок `rngRoundedTrapezoidGeometry`: первая колонка ключ,
вторая значение, третья единица, четвёртая комментарий. Порядок ключей
не задаёт идентичность геометрии. Длины отображают текущую input-единицу
через формулу к rngUnitSettings; counts в штуках.

| Ключ | Проверенное значение в мм/шт | Смысл |
|---|---:|---|
| BottomWidth | 1200 | Нижняя исходная ширина до скругления |
| TopWidth | 800 | Верхняя исходная ширина до скругления |
| Height | 800 | Высота |
| CornerRadius | 60 | Радиусы четырёх углов |
| OpeningRadius | 110 | Радиус отверстия |
| OuterAxisDistance | 50 | Расстояние до осей от Outer внутрь бетона |
| InnerAxisDistance | 45 | Расстояние до осей от Opening из пустоты |
| OuterCount | 24 | Стержни по всему наружному пути |
| InnerCount | 16 | Стержни по всему внутреннему пути |
| OuterDiameter | 20 | Диаметр наружных стержней |
| InnerDiameter | 16 | Диаметр внутренних стержней |

Программный вариант тех же данных не читает Config:

```vb
Dim recipe As CSectionRecipeRoundedTrapezoid, definition As CSectionDefinition
Set recipe = New CSectionRecipeRoundedTrapezoid
Set definition = recipe.ProgramDefinition(1200#, 800#, 800#, 60#, 110#)
```

Все аргументы ProgramDefinition уже в мм. Сетка и физическая арматура
строятся общим CSectionBuilder; meshOptions задаётся вызывающим workflow.
Полный импортируемый код ниже синхронизируется с реально проверенным `.cls`.

<!-- complete-example:start -->
```vb
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
    Set ISectionRecipe_CreateDefinition = Describe(parameters, table)
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
    Set ProgramDefinition = Describe(parameters, Nothing)
End Function

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
    Dim facade As CSectionDefinitionBuilder, boundary As CGeometryPath, openingPath As CGeometryPath, axis As CGeometryPath
    Dim vertices As Variant, radii As Variant, corners As Variant, options As CRebarRuleOptions
    Set facade = New CSectionDefinitionBuilder: facade.Initialize "RoundedTrapezoid"
    vertices = Array(Array(-parameters.BottomWidth / 2#, -parameters.Height / 2#), Array(parameters.BottomWidth / 2#, -parameters.Height / 2#), _
        Array(parameters.TopWidth / 2#, parameters.Height / 2#), Array(-parameters.TopWidth / 2#, parameters.Height / 2#))
    radii = Array(parameters.CornerRadius, parameters.CornerRadius, parameters.CornerRadius, parameters.CornerRadius)
    corners = Array("BottomLeft", "BottomRight", "TopRight", "TopLeft")
    facade.Source "Boundary", parameters.Source
    Set boundary = facade.RoundedPolygon("Boundary", vertices, radii, corners)
    facade.Outer "Concrete", boundary.ID, "Наружная граница скругленной трапеции."
    Set openingPath = facade.CirclePath("Opening.Circle", 0#, 0#, parameters.OpeningRadius)
    facade.Opening "Opening", openingPath.ID, "Concrete", "Круглое отверстие."
    Set axis = facade.OffsetRoundedPolygon("Outer.Axis", vertices, radii, corners, parameters.OuterAxisDistance)
    Set options = New CRebarRuleOptions
    options.EndMode = ClosedUniform: options.PlacementCheck = CheckFullDisk: options.OverlapPolicy = RejectOverlap
    options.AnnotationGroup = "OuterBars": options.SourcePrefix = "Outer."
    options.SourceReference = parameters.Source
    If Not table Is Nothing Then options.SourceReference = table.SourceReference(table.RowByKey("OuterAxisDistance"), 2) & "; " & table.SourceReference(table.RowByKey("OuterCount"), 2)
    facade.RebarByCount "OuterBars", axis.ID, parameters.OuterCount, parameters.OuterDiameter, options
    ' CirclePath имеет обход CCW: отрицательный left-offset выходит из пустоты.
    Set axis = facade.Offset("Inner.Axis", openingPath.ID, -parameters.InnerAxisDistance)
    Set options = New CRebarRuleOptions
    options.EndMode = ClosedUniform: options.PlacementCheck = CheckFullDisk: options.OverlapPolicy = RejectOverlap
    options.AnnotationGroup = "InnerBars": options.SourcePrefix = "Inner."
    options.SourceReference = parameters.Source
    If Not table Is Nothing Then options.SourceReference = table.SourceReference(table.RowByKey("InnerAxisDistance"), 2) & "; " & table.SourceReference(table.RowByKey("InnerCount"), 2)
    facade.RebarByCount "InnerBars", axis.ID, parameters.InnerCount, parameters.InnerDiameter, options
    facade.VerticalExtent "Height", boundary.ID, "H", -parameters.BottomWidth / 2#, -1#
    facade.HorizontalExtent "Width", boundary.ID, "B", -parameters.Height / 2#, -1#
    facade.Radius "OpeningRadius", openingPath.ID, openingPath.ID, 0.25, "Ro"
    facade.Anchor "Caption", -parameters.TopWidth / 4#, parameters.Height / 5#
    facade.Text "FormName", "Caption", "Скругленная трапеция"
    facade.RebarLabel "OuterBars", "Outer.Axis", 0#, 1#, parameters.OuterAxisDistance
    facade.RebarLabel "InnerBars", "Inner.Axis", 0#, -1#, parameters.InnerAxisDistance
    Set Describe = facade.Finish
    Exit Function
InvalidGeometry:
    SectionRecipeGeometryError "CSectionRecipeRoundedTrapezoid", parameters.Source
End Function
```
<!-- complete-example:end -->

## Пути, границы и привязки

`CirclePath`, `LinePath`, `ArcPath` создают точные аналитические примитивы.
Дуга задаётся центром, радиусом, начальным углом и signed sweep: положительный
CCW, отрицательный CW; всё в радианах. Нулевая дуга или более одного оборота
отклоняются. Ломаная допускается только для графики, не для шага арматуры.

`RoundedPolygon` принимает вершины, радиусы и устойчивые vertexIDs.
На выходе именованные `Edge.<vertexID>` и `Corner.<vertexID>`; при R=0
Corner отсутствует. Получайте касательные точки `PathAnchor` по ID сегмента
и параметру 0/1. `Chain` выбирает участок по двум таким привязкам и явному
направлению; fullLoopIfSame нужен для полного обхода при совпавших концах.
По исчезнувшей дуге должна быть адресная ошибка, а не соседний индекс.

Только `Outer`/`Opening` включают путь в бетон. Opening принадлежит указанному
Outer, вложение и пересечения проверяются геометрически. Вспомогательная
линия ряда не становится границей бетона. Несколько Outer/Opening допустимы
в рамках общего topology-контракта. `AdjacentRectangles` объединяет смежные
неперекрывающиеся прямоугольники; это не произвольный CAD Boolean Union.

Path — одна непрерывная цепочка. PathGroup — независимые пути без фиктивных
соединений. `DistributionScope=PerPath` применяет количество/шаг и setbacks
отдельно к каждому пути. Режим TotalForGroup не реализован и не выбирается
молчаливо. Один боковой путь из пяти сегментов получает одно число стержней.

`Transform` переносит/поворачивает/отражает путь, сохраняя смысловые ID;
отражение относительно локальной X выполняется перед поворотом/переносом.
Для отдельной именованной точки используйте `TransformAnchor` с теми же
параметрами или PathAnchor на преобразованном пути. Рецепт не использует пиксели.

## Арматура и дополнительные ряды

AxisDistance (as) — положительное расстояние от границы до **оси**, а не
до поверхности стержня. `OffsetRoundedPolygon` положительным расстоянием
идёт внутрь площади, отрицательным из отверстия. `Offset` использует иной
явный контракт: signedLeftDistance положителен слева от обхода пути.
Не подменяйте эту сторону направлением к центру тяжести.

Сначала согласуйте осевой путь/его стыки и ограничения. `ConstrainLine`
задаёт его геометрические концы двумя поперечными плоскостями; затем
StartSetback/EndSetback отсекают рабочий участок **вдоль** пути. У полного
замкнутого ClosedUniform оба setbacks равны нулю. Для кольца с ненулевыми
setbacks сначала создайте явный открытый участок Chain/Slice.

ByCount сохраняет заданное n и распределяет станции по рабочей длине.
ByMaxStep берёт достаточное n, чтобы шаг не превышал заданный; фактический
шаг может быть меньше. ByFixedStep сохраняет точный шаг, размещая остаток
RemainderStart/End/Center; фиксированные оба конца/полный оборот требуют
совместимости длины с шагом. ExplicitPositions задаёт станции по длине пути.

EndMode выбирает включение концов: IncludeBoth, ExcludeBoth, IncludeStart,
IncludeEnd, ClosedUniform. Это отличается от обрезки рабочей длины setbacks.
Дуги используют свою длину, не хорду или проекцию. Длина дуги,
CenterStep и VerticalExtent измеряют разные величины.

ProjectFrom проецирует исходные позиции на конечный целевой путь. Направление
задаётся DirectionX/Y или локальной SourceNormalSign; target setbacks фильтруют
hit и не перераспределяют ряд. SkipMissingProjection явно разрешает пропуск
с диагностикой; иначе отсутствие hit — ошибка. После пропусков выбирайте,
нужно ли SuccessfulProjectionOrdinal/NewPublicationRoot для связанного ряда.

DependentOffset получает связанные позиции локальным TangentOffset и
NormalOffset. Stacked означает местную нормаль; SideBySide — касательную.
Названия Stacked/SideBySide читает рецепт, не общий solver. При разных
диаметрах расстояние центров включает радиусы обоих стержней; при пропущенном
втором ряде в третьем учитывается его диаметр согласно принятой раскладке.
EverySecond выбирает исходные ordinals, не номер позиции после сортировки.
IncomingTangentAtJoint явно выбирает входящую касательную в стыке.

Publication/PositionOrder управляют порядком вывода отдельно от геометрии.
Coincidence задаёт RejectCoincidence/MergeIdentical/KeepPublishedOwner.
Один физический стержень учитывается один раз; несовпавшие диаметры/материалы
не разрешаются случайным обходом. ReportOverlap сообщает геометрические
пересечения дисков; RejectOverlap запрещает их. Контакт сам по себе допустим.
Это не нормативный минимальный зазор. CheckCenter и CheckFullDisk — отдельные
проверки размещения относительно бетона/отверстий.

## Измерения, подписи и диагностика

Dimension измеряет Horizontal/Vertical/Aligned между именованными точками.
HorizontalExtent/VerticalExtent используют полные экстремумы выбранного
пути/группы, включая внутренние экстремумы дуги. Габарит не сокращается
от setbacks армирования. Radius/Diameter измеряют конкретный Arc/Circle,
PathLength — точную длину пути; Text имеет модельную привязку.

RebarLabel выбирает геометрический target и фактический состав группы;
RebarRowLabel/RebarRowExtentLabel/RebarRowExtremeLabel привязаны к созданному
правилу. Подписи строятся по физической арматуре. Выберите стороны/модельные
привязки так, чтобы строки не совпали; общий layout применит настройки
шрифта, стрелок и масштаба. Старые типы не получают новые радиусы автоматически.

Builder.Rules сообщает Requested/Created/Merged/Skipped/Conflicts/BarOverlaps,
RuleStepBounds и происхождение каждой позиции/её физического владельца.
В кольце учитывается замыкающий интервал. StationStep относится к пути
(у DependentOffset — исходному), CenterStep — хорда. Подробных образцов
до 128, DiagnosticRecordsOmitted сообщает остальные события; полные
счётчики сохранены. LastDiagnostic сохраняет последнее событие.

Неудача геометрии/численного решения не объявляется физическим разрушением.
Не менять единый mapper статусов. Results хранит геометрию/контуры/аннотации,
единицы и RecipeTrace; повторная схема/экспорт не импортируют и не решают НДС.

## Проверки и типичные ошибки

Полные рабочие команды и фактические результаты приведены в
ONE_CLASS_EXAMPLE.md, ONE_CLASS_EVIDENCE.json и DECLARATIVE_API_EVIDENCE.json.
Проверяйте новое определение, явные Outer/Opening, истинные длины дуг,
координаты/число/диаметры/принадлежность баров, A/центр области отдельно от
механической сетки, равновесие N+Mx+My, статусы/ResultComment и исходные единицы.
Затем preview, реальный расчёт, схему и save/reopen на отдельной книге;
после смены параметра повторите. Diff должен подтверждать одну регистрацию
и один алгоритмический класс типа. Generic tests/UI-метаданные допустимы.

E09 — программная рамка 1200×800 R80, отверстие 600×300 R40 в (+50,-20),
наружные 24 d20 as60, внутренние 16 d16 as45. Контроль A/центра/40 bars и
соответствующие методы находятся в tests/modTestDeclarativeExamples.bas.
Ro=0/ProjectFrom с пропусками проверены общими тестами Hollow и rule executor.

| Причина отказа | Что проверить и исправить |
|---|---|
| Не найден путь/Corner | ID и факт существования дуги при R=0; выбрать существующий Edge/явный участок |
| Разрыв цепочки | Реальные концы сегментов; несвязанные части назначить группе |
| Offset исчез/негладкий arc стык | Радиус/as/согласованное скругление; не выбирать длинную часть молча |
| Параллельные/развёрнутые ограничители | Нормали плоскостей, направление и порядок двух bounds |
| Setbacks больше длины | Рабочий участок и единицы; as не является продольным setback |
| FixedStep не совместим | Длину/шаг/EndMode либо явное размещение остатка |
| Projection отсутствует | Направление, конечный target, ограничения; выбрать Error/Skip явно |
| Диск вне бетона/пересечение баров | Диаметр, as, шаг и зависимые смещения; проверить фактического владельца |
| Пустой активный ввод/ошибка формулы | Адрес из SourceReference, именованный блок и обязательное поле |
| Новый тип отсутствует в выборе | Имя class/регистрацию и явное UpdateSectionRecipeConfig |

Расчётные проверки не заменяют нормативную верификацию. Полные длительные
on/off и Excel UI наборы требуют отдельного согласия; текущая задача использует
разрешённые направленные сценарии по требованиям.
