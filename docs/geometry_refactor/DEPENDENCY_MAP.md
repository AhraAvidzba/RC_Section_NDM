# Исходная Карта Зависимостей

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
