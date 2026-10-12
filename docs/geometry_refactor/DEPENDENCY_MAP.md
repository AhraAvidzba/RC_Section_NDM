# Исходная Карта Зависимостей

Актуализация 12.10.2026: заменяемые компоненты ниже удалены из source и
отдельного native staging; фактический состав — в LEGACY_RETIREMENT_MAP.md.
Generated: CSectionInputSnapshotReader → ISectionRecipe → CSectionDefinition
→ CSectionBuilder → CSectionModel. Импорт AutoCAD сохраняет свою сетку;
его аннотации готовит CSectionAnnotationBuilder. CSectionAnnotations хранит
данные. Основной output ещё предстоит обновить.

Состояние до переноса: HEAD a138c273, ТЗ 1.4. Карта будет уточняться по
потребителям и фактическому VBProject, не только по экспорту от 09.10.2026.

| Узел | Текущая роль | Дальнейшая ответственность |
|---|---|---|
| CSystemSettingsReader | Системные настройки и разбор типовых таблиц | Системные настройки; generic источник таблиц отдельно от рецептов |
| CSectionTypeRegistry | CreateShape/CreateGeometry/CreateRebars/BuildAnnotations/BuildGeneratedModel | Один реестр рецептов и метаданных выбора |
| modWorkbookCalculation.BuildWorkbookSectionModel | Generated и AutoCAD | Единый Generated recipe/build; независимый CAD source |
| 4 CGeometry* | Контуры и специфическая геометрия | 4 коротких рецепта плюс общие операции |
| 4 *RebarLayoutBuilder | Специфическая расстановка | Один общий исполнитель декларативных правил |
| 4 *AnnotationBuilder | Смысловые аннотации | Общие правила измерения и фактической арматуры |
| CSectionGeometryQuery | Точная область, clipping, пересечения, topology, mesh fallback | Сохранить и переиспользовать |
| CGeometryRegion/CFiberMeshBuilder | Общая область и сетка | Сохранить; не менять источник механических A/I |
| CSectionContours/CSectionModel | Контуры/Opening и расчетные элементы | Сохранить единый опубликованный снимок |
| CNDMResultsWriter/CSectionPlotDataReader | Сохранение и чтение Results | Сохранить атомарный контракт и единицы снимка |
| CPlotAnnotationLayout/CSectionPlotter | Excel схема | Сохранить общий layout, ownership и пассивное чтение |
| CAutoCADSectionModelImporter/modAutoCADStressExport | CAD source/export | Не пересоздавать импортированную сетку; экспортировать снимок |
| modTestShapeArchitecture | Неполный snapshot девяти fixtures | Полные независимые baseline/new данные и проверки |

Runtime-списки классов, имен, OnAction и References будут в Baseline inventory.
Ссылки на заменяемые символы фиксируются до начала удаления и повторно после него.

Первый фактический поиск: 729 строк src/tests/tools в
`docs/regression/GeometryDeclarative_2026-10-11/Baseline/DependencyReferences.txt`.
Расширенный snapshot тоже использует исходные registry CreateGeometry/CreateRebars;
после переноса он должен читать готовые mesh/layout из единого build, а не
оставаться адаптером старой архитектуры.

Первый новый общий слой:
- CGeometryPath -> CSectionGeometryQuery Friend-операции; никаких solver/Excel/СП.
- CSectionGeometryQuery.RoundedPolygon/RegionFromPath -> общий путь и прежняя topology.
- CSectionInputTable -> массивы скаляров/адресов и CUnitSystem для конвертации;
  типовых разборов и live Range нет. Workbook adapter и recipe context еще не подключены.
- modTestDeclarativeGeometry -> программные fixtures, без старых генераторов.

Registry/reader/layout потребители пока прежние; retirement-map остается ACTIVE.

Сквозной development-маршрут Circle:
- CSectionTypeRegistry.RegisterRecipe -> ISectionRecipe + имена input tables;
  одна регистрация Circle. Production selection пока прежний до полного переноса.
- CSectionInputSnapshotReader -> CSectionInputTable/CSectionRecipeContext:
  по одному Value2 на таблицу, адреса формируются без COM-чтения каждой ячейки.
- CSectionRecipeCircle -> CSectionDefinitionBuilder -> CSectionDefinition:
  никаких CGeometryCircle/CCircleRebarLayoutBuilder/CCircleAnnotationBuilder.
- CSectionBuilder -> CGeometryRegion/CFiberMeshBuilder/CRebarRuleExecutor/
  CSectionModelBuilder/CSectionAnnotationBuilder; нет ветвления по TypeID.
- CSectionAnnotationBuilder -> CRebarGroupAnnotationBuilder, один формат подписи.
- CRebarRuleExecutor -> CGeometryPath и query.BoundaryDistance для полного диска;
  NDM/СП35-группы не смешиваются с группами геометрических путей.
- modTestSectionRecipes -> реальная новая модель; Prepared*Snapshot читает
  mesh/layout этой же сборки, не реконструирует их через прежний registry.

Старый test SourceSnapshot удален из актуального harness; его прежний код
сохраняется только в неизменяемых baseline-артефактах.

Сквозной development-маршрут RectSet:
- Одна регистрация CSectionRecipeRectSet; одна таблица rngRectSetGeometry.
- Рецепт описывает Rectangle/LSection/TwoRectangles и восемь граней, без
  CGeometryRectSet/CRectSetRebarLayoutBuilder/CRectSetAnnotationBuilder.
- Общий query.AdjacentRectangles -> NormalizeBoundary/BoundaryRegion;
  только смежные неперекрывающиеся прямоугольники, без нового Boolean engine.
- CRebarRuleExecutor -> отдельные ordinal/publication rank, явные coincidence
  policies и preflight. CRebarLayout -> точная проверка кандидатов по индексу,
  геометрическая capacity, physical bar/memberships.
- CSectionAnnotationBuilder -> общий RuleEnds для подписи ряда; существующий
  форматировщик считает реальные стержни только по группе владельца;
  дополнительные memberships сохраняют трассировку, но не входят в подписи.
- Compare-GeometrySnapshots записывает bijection только source/layout имен
  сгенерированной арматуры. Контурные роли, ID модели и физические поля не
  переименовываются и не исключаются из сравнения.

Production-ветви registry/reader пока прежние; все заменяемые узлы остаются
ACTIVE в retirement-map до переноса остальных семейств и удаления потребителей.

Подготовка аналитического переноса RoundedRectangle/HollowRectangle:
- CSectionGeometryQuery.OffsetRoundedPolygon -> существующие RoundedPolygon/
  BoundaryRegion, согласованные опорные прямые и аналитические радиусы.
- CGeometryPath -> именованные привязки/цепочки и origin intervals; направление
  исходного пути не нормализуется вместе с бетонной областью.
- CSectionDefinitionBuilder.Chain/Transform/OffsetRoundedPolygon -> только
  общий query/path; в фасаде нет выбора конкретной формы.
- Сложный рецепт разделяется внутренними методами ввода/геометрии/армирования/
  аннотаций; это не новые обязательные классы и не дополнительный маршрут.

Сквозной development-маршрут RoundedRectangle:
- CSectionRecipeRoundedRectangle -> generic input snapshot; ReadDimensions,
  AddGeometry, AddReinforcement, AddDimensions показывают последовательность.
- Shape descriptors -> query RoundedPolygon/OffsetRoundedPolygon -> именованные
  Chain B.Top/B.Bottom/H.Left/H.Right. В рецепте нет расчета дуг/касательных
  или обхода сетки; старый CGeometryRoundedRectangle не вызывается.
- CRebarRuleOptions.IncomingTangentAtJoint -> общий CGeometryPath.AtDistance;
  параметр задает направление на острой вершине, не выбирает тип сечения.
- RuleBounds -> RebarRowExtentLabel; диапазон фактических основных позиций
  используется только для аннотации, не для повторного размещения стержней.
- DisplayIncrement округляет только справочное значение размера; реальные
  концы линии и расчетная геометрия не меняются.
- SectionRecipeGeometryError добавляет source reference к ошибке query/path,
  непредвиденные ошибки сохраняет с исходным кодом/источником.
- Verify-RoundedRecipeOracle использует только input и независимые формулы
  в Python, не фактические VBA пути/позиции как ожидаемый результат.

Старые три RoundedRectangle-компонента остаются ACTIVE в production до
общего retirement.

Сквозной development-маршрут HollowRectangle:
- Один рецепт: ReadDimensions/ReadFaces -> AddGeometry -> AddReinforcement ->
  AddDimensions. Наружный контур и Opening объявляются на уровне definition.
- Общие RoundedPolygon/OffsetRoundedPolygon/Chain задают точные дуги;
  при исчезновении прямой грани именованная цепочка начинается на дуге.
- RebarProjected получает локальную нормаль исходного пути. Номер после
  успешных проекций и самостоятельный публикационный корень заданы options;
  общий executor не выбирает тип сечения.
- RebarRowExtremeLabel использует крайние основные стержни на одной опорной
  линии, не повторный подбор или расчет положения арматуры.
- Python oracle независимо строит дуги, станции и лучевые пересечения по
  исходному INPUT. Изменения координат не скрываются изменением допусков.

Все четыре семейства подключены к новому source/native staging Generated route:
- BuildWorkbookSectionModel -> ReadWorkbookSectionDefinition -> registry metadata
  -> Capture выбранных таблиц -> CreateDefinition -> CSectionBuilder.
- modMeshConfiguration.ReadSectionMeshOptions общий для Generated и CAD mesh.
- CSystemSettingsReader больше не сканирует/разбирает geometry-таблицы.
- modTestShapeArchitecture читает prepared mesh/layout без повторной сборки.
- RegisteredSectionTypes дает установщику динамический список типов; подключение
  dropdown еще предстоит. Основная книга и ее общий dump пока прежние.

12 старых классов, ISectionShape и численные test constructors еще требуют
retirement. Они не вызываются новым production source-входом.

Перенос численных fixtures:
- modTestGeometryFixtures -> CSectionGeometryQuery/CSectionDefinitionBuilder/
  CGeometryRegion/CRebarRuleExecutor. Только явные test dimensions/paths,
  без Config и без регистрации производственных типов.
- 11 затронутых modules не вызывают прежние специализированные classes.
  Raw load-matrix snapshot берет model/mesh/layout одной prepared-сборки.
- Multi-opening ownership -> definition.Outer/Opening -> WriteContours ->
  Results/query; CFakeMultiOpeningGeometry больше не существует в source
  и native staging.
- Остались два test consumers старого слоя: modTestGeometry и
  modTestWorkbookInterface. Полный native retirement еще впереди.
