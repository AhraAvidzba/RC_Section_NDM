# SECTION_RECIPE_API: каталог текущих сигнатур

Сигнатуры, defaults и каждый контекстный вызов синхронизированы с source; native entrypoint/журнал указан у каждого члена. Workbook-расположение фиксирует NativeMap.json, итоговая приёмка книги — отдельный отчёт.

Все Public члены фасада, его options/результатов и общего ввода/единиц; общие query/region и build diagnostics для проверок. Private/Friend не API рецепта.

Внутри мм/мм²/Н/Н*мм/МПа; углы в радианах. ByRef/Optional/default приведены в точной сигнатуре. Краткий порядок подключения — в ADDING_SECTION_TYPE.md; методы автора сечения — в modSectionRecipeApiGuide, тела методов — в исходных классах VBE.

## ISectionRecipe

Вызывается реестром с готовым context; создаёт новое definition. Не выполняет mesh/layout/solver/Excel output.

<a id="api-isectionrecipe-createdefinition"></a>
### ISectionRecipe.CreateDefinition

Интерфейс единственного класса типа. Возвращает опубликованное definition через общий фасад, без сетки/баров/Shapes в рецепте.

```vb
Public Function CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| context | CSectionRecipeContext; ByVal; обязательный | Подготовленный снимок CSectionRecipeContext, не live Excel. |

Тип результата: CSectionDefinition. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Вызывается реестром с готовым context; создаёт новое definition. Не выполняет mesh/layout/solver/Excel output.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/ISectionRecipe.cls:14

Минимальный вызов; требуемый контекст: В теле modTestRoundedTrapezoidExample.RunAudit02MinimalRecipeTests с его подготовленными локальными объектами/входом.

```vb
    Set definition = recipe.CreateDefinition(context)
```

Полный контекст: docs/geometry_refactor/examples/modTestRoundedTrapezoidExample.bas:15.

Native entrypoint: modTestRoundedTrapezoidExample.RunAudit02MinimalRecipeTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/ExampleTests/modTestRoundedTrapezoidExample.RunAudit02MinimalRecipeTests.txt.

## CSectionDefinitionBuilder

New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

<a id="api-csectiondefinitionbuilder-initialize"></a>
### CSectionDefinitionBuilder.Initialize

Открывает один новый фасад с TypeID и версией схемы. Повторная инициализация запрещена.

```vb
Public Sub Initialize(ByVal typeID As String, Optional ByVal schemaVersion As Long = 1)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| typeID | String; ByVal; обязательный | Смысловое имя зарегистрированного типа; непустой текст. |
| schemaVersion | Long; ByVal; Optional; default=1 | Положительная целая версия схемы определения. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:21

Минимальный вызов; требуемый контекст: В теле modTestGeometryPublication.RunDefinitionLifecycleTests с его подготовленными локальными объектами/входом.

```vb
    facade.Initialize "Again"
```

Полный контекст: tests/modTestGeometryPublication.bas:217.

Native entrypoint: modTestGeometryPublication.RunDefinitionLifecycleTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/SolidFinal_2026-10-14/SolidTests/modTestGeometryPublication.RunDefinitionLifecycleTests.txt.

<a id="api-csectiondefinitionbuilder-circlepath"></a>
### CSectionDefinitionBuilder.CirclePath

Создаёт замкнутую окружность CCW от точки +X; один сегмент с тем же ID, что путь.

```vb
Public Function CirclePath(ByVal pathID As String, ByVal centerX As Double, ByVal centerY As Double, ByVal radius As Double) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| centerX | Double; ByVal; обязательный | Модельная координата centerX, мм; знаки соответствуют section XY/WCS. |
| centerY | Double; ByVal; обязательный | Модельная координата centerY, мм; знаки соответствуют section XY/WCS. |
| radius | Double; ByVal; обязательный | Положительный радиус, мм. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:28

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.AdvancedGuideDefinition с его подготовленными локальными объектами/входом.

```vb
    Set path = facade.CirclePath("Hole", 0#, -250#, 50#)
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:43.

Native entrypoint: modTestRecipeApiCatalogue.RunAdvancedGuideExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/AdvancedExamples.txt.

<a id="api-csectiondefinitionbuilder-linepath"></a>
### CSectionDefinitionBuilder.LinePath

Создаёт непрерывный открытый путь из одной прямой от первого конца ко второму.

```vb
Public Function LinePath(ByVal pathID As String, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| x1 | Double; ByVal; обязательный | Модельная координата x1, мм; знаки соответствуют section XY/WCS. |
| y1 | Double; ByVal; обязательный | Модельная координата y1, мм; знаки соответствуют section XY/WCS. |
| x2 | Double; ByVal; обязательный | Модельная координата x2, мм; знаки соответствуют section XY/WCS. |
| y2 | Double; ByVal; обязательный | Модельная координата y2, мм; знаки соответствуют section XY/WCS. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:38

Минимальный вызов; требуемый контекст: В теле modTestGeometryReview.TestPerPath с его подготовленными локальными объектами/входом.

```vb
    Set path = facade.LinePath("P1", 0#, 0#, 100#, 0#)
```

Полный контекст: tests/modTestGeometryReview.bas:365.

Native entrypoint: modTestGeometryReview.RunGeometryReviewTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/SolidFinal_2026-10-14/PublicationTests/modTestGeometryReview.RunGeometryReviewTests.txt.

<a id="api-csectiondefinitionbuilder-arcpath"></a>
### CSectionDefinitionBuilder.ArcPath

Создаёт точную дугу по центру/радиусу/startAngle/signed sweep; полный оборот замкнут.

```vb
Public Function ArcPath(ByVal pathID As String, ByVal centerX As Double, ByVal centerY As Double, ByVal radius As Double, ByVal startAngle As Double, ByVal sweep As Double) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| centerX | Double; ByVal; обязательный | Модельная координата centerX, мм; знаки соответствуют section XY/WCS. |
| centerY | Double; ByVal; обязательный | Модельная координата centerY, мм; знаки соответствуют section XY/WCS. |
| radius | Double; ByVal; обязательный | Положительный радиус, мм. |
| startAngle | Double; ByVal; обязательный | Начальный угол, радианы; от +X. |
| sweep | Double; ByVal; обязательный | Signed sweep, радианы: CCW положителен, CW отрицателен, 0 запрещён, \|sweep\|≤2*pi. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:48

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.TestAnnotationMeasurements с его подготовленными локальными объектами/входом.

```vb
    Set path = facade.ArcPath("Curve", 0#, 0#, 100#, 0#, GEOM_PI / 2#)
```

Полный контекст: tests/modTestSectionRecipes.bas:349.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-csectiondefinitionbuilder-roundedpolygon"></a>
### CSectionDefinitionBuilder.RoundedPolygon

Создаёт и регистрирует согласованную границу по одному завершённому CPolygonSpec. Edge.ID начинается у исходной вершины; Corner.ID существует при R>0. Исходный описатель не меняется.

```vb
Public Function RoundedPolygon(ByVal pathID As String, ByVal polygon As CPolygonSpec) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| polygon | CPolygonSpec; ByVal; обязательный | Завершённый CPolygonSpec исходных вершин до скругления. Nothing и не закрытый через Seal набор отклоняются; порядок, радиусы и источники сохраняются. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:60

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02PolygonApiTests с его подготовленными локальными объектами/входом.

```vb
    Set path = facade.RoundedPolygon("Boundary", crossed)
```

Полный контекст: tests/modTestAudit02Api.bas:39.

Native entrypoint: modTestAudit02Api.RunAudit02PolygonApiTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02PolygonApiTests.txt.

<a id="api-csectiondefinitionbuilder-offsetroundedpolygon"></a>
### CSectionDefinitionBuilder.OffsetRoundedPolygon

Создаёт и регистрирует смещённый скруглённый многоугольник по исходному CPolygonSpec, не меняя его. Положительное расстояние внутрь площади; отрицательное наружу, из Opening в бетон. Сохраняет привязки исчезнувших участков; отличается от signed-left Offset готового пути.

```vb
Public Function OffsetRoundedPolygon(ByVal pathID As String, ByVal polygon As CPolygonSpec, ByVal distance As Double) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| polygon | CPolygonSpec; ByVal; обязательный | Завершённый CPolygonSpec исходных вершин до скругления. Nothing и не закрытый через Seal набор отклоняются; порядок, радиусы и источники сохраняются. |
| distance | Double; ByVal; обязательный | Signed offset, мм: положительно внутрь площади, отрицательно наружу; это не left-offset обхода. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:68

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02GuideExamplesTests с его подготовленными локальными объектами/входом.

```vb
    Set path = facade.OffsetRoundedPolygon("OffsetPath", polygon, 50#)
```

Полный контекст: tests/modTestAudit02Api.bas:234.

Native entrypoint: modTestAudit02Api.RunAudit02GuideExamplesTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/FinalApiAcceptance/modTestAudit02Api.RunAudit02GuideExamplesTests.txt.

<a id="api-csectiondefinitionbuilder-addpath"></a>
### CSectionDefinitionBuilder.AddPath

Регистрирует непрерывную линию, созданную вручную или отдельным общим инструментом. Создающие команды фасада уже регистрируют свой результат; повторный AddPath для него является ошибкой. Назначение Outer/Opening выполняется отдельно.

```vb
Public Sub AddPath(ByVal path As CGeometryPath)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| path | CGeometryPath; ByVal; обязательный | Существующий непрерывный опубликованный CGeometryPath. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:77

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02PolygonApiTests с его подготовленными локальными объектами/входом.

```vb
    facade.AddPath path
```

Полный контекст: tests/modTestAudit02Api.bas:45.

Native entrypoint: modTestAudit02Api.RunAudit02PolygonApiTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02PolygonApiTests.txt.

<a id="api-csectiondefinitionbuilder-anchor"></a>
### CSectionDefinitionBuilder.Anchor

Объявляет именованную модельную точку. Данные копируются как два скаляра.

```vb
Public Sub Anchor(ByVal anchorID As String, ByVal x As Double, ByVal y As Double)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| anchorID | String; ByVal; обязательный | Непустой уникальный ID модельной привязки. |
| x | Double; ByVal; обязательный | Модельная координата x, мм; знаки соответствуют section XY/WCS. |
| y | Double; ByVal; обязательный | Модельная координата y, мм; знаки соответствуют section XY/WCS. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:83

Минимальный вызов; требуемый контекст: В теле modTestRecipeAnnotations.AnnotationFacade с его подготовленными локальными объектами/входом.

```vb
        facade.Anchor "TextNote", 200#, 0#
```

Полный контекст: tests/modTestRecipeAnnotations.bas:154.

Native entrypoint: modTestRecipeAnnotations.RunRecipeAnnotationTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/DeclarativeApi/FinalApiAnnotations.txt.

Native entrypoint: modTestRecipeAnnotations.RunRecipeAnnotationSnapshotTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/DeclarativeApi/Operations.txt.

<a id="api-csectiondefinitionbuilder-pathanchor"></a>
### CSectionDefinitionBuilder.PathAnchor

Получает точку конкретного сегмента по ID и параметру [0;1], объявляет именованную привязку.

```vb
Public Sub PathAnchor(ByVal anchorID As String, ByVal pathID As String, ByVal segmentID As String, Optional ByVal parameter As Double = 0#)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| anchorID | String; ByVal; обязательный | Непустой уникальный ID модельной привязки. |
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| segmentID | String; ByVal; обязательный | Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом. |
| parameter | Double; ByVal; Optional; default=0# | Нормированный параметр точки сегмента [0;1], не расстояние по всему пути. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:89

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02AnnotationTests с его подготовленными локальными объектами/входом.

```vb
    facade.PathAnchor "Touch", path.ID, "Edge.A", 0#
```

Полный контекст: tests/modTestAudit02Api.bas:165.

Native entrypoint: modTestAudit02Api.RunAudit02AnnotationTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02AnnotationTests.txt.

<a id="api-csectiondefinitionbuilder-source"></a>
### CSectionDefinitionBuilder.Source

Связывает смысловой ID со строкой источника/адреса без сохранения объекта Excel.

```vb
Public Sub Source(ByVal entityID As String, ByVal sourceReference As String)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| entityID | String; ByVal; обязательный | Смысловой ID сущности для SourceMap. |
| sourceReference | String; ByVal; обязательный | Строка исходного адреса/программного параметра; без объекта Excel. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:98

Минимальный вызов; требуемый контекст: В теле modTestRecipeAnnotations.RunRecipeAnnotationTests с его подготовленными локальными объектами/входом.

```vb
    facade.Source "Diagonal", "ProgramInputs.PointB"
```

Полный контекст: tests/modTestRecipeAnnotations.bas:13.

Native entrypoint: modTestRecipeAnnotations.RunRecipeAnnotationTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/DeclarativeApi/FinalApiAnnotations.txt.

<a id="api-csectiondefinitionbuilder-transformanchor"></a>
### CSectionDefinitionBuilder.TransformAnchor

Создаёт новую привязку из существующей: отражение по локальной X, поворот вокруг (0,0), перенос.

```vb
Public Sub TransformAnchor(ByVal anchorID As String, ByVal sourceAnchorID As String, ByVal shiftX As Double, ByVal shiftY As Double, Optional ByVal rotation As Double = 0#, Optional ByVal reflectX As Boolean = False)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| anchorID | String; ByVal; обязательный | Непустой уникальный ID модельной привязки. |
| sourceAnchorID | String; ByVal; обязательный | ID существующей исходной привязки. |
| shiftX | Double; ByVal; обязательный | Перенос по X, мм. |
| shiftY | Double; ByVal; обязательный | Перенос по Y, мм. |
| rotation | Double; ByVal; Optional; default=0# | Угол поворота вокруг (0,0), радианы, CCW положителен. |
| reflectX | Boolean; ByVal; Optional; default=False | Отражение относительно локальной оси X перед поворотом и переносом. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:105

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.BuildAudit02OperationsDefinition с его подготовленными локальными объектами/входом.

```vb
    facade.TransformAnchor "MovedA", "A", 0#, 700#
```

Полный контекст: tests/modTestAudit02Api.bas:316.

Native entrypoint: modTestAudit02Api.RunAudit02GuideExamplesTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/FinalApiAcceptance/modTestAudit02Api.RunAudit02GuideExamplesTests.txt.

<a id="api-csectiondefinitionbuilder-offset"></a>
### CSectionDefinitionBuilder.Offset

Смещает линию/аналитическую цепочку слева от обхода при положительном signedLeftDistance. Line/line стыки miter; arc стыки должны быть гладкими.

```vb
Public Function Offset(ByVal pathID As String, ByVal sourcePathID As String, ByVal signedLeftDistance As Double) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| sourcePathID | String; ByVal; обязательный | ID ранее объявленного пути, который читается без изменения. |
| signedLeftDistance | Double; ByVal; обязательный | Поперечное смещение, мм; положительно слева от направления пути, отрицательно справа. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:116

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    Set path = facade.Offset("Lost", "Small", 20#)
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:152.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-csectiondefinitionbuilder-constrainline"></a>
### CSectionDefinitionBuilder.ConstrainLine

Ограничивает или продлевает одну прямую двумя поперечными плоскостями nX*x+nY*y=bound; направление исходного пути сохраняется.

```vb
Public Function ConstrainLine(ByVal pathID As String, ByVal sourcePathID As String, ByVal startNX As Double, ByVal startNY As Double, ByVal startBound As Double, ByVal endNX As Double, ByVal endNY As Double, ByVal endBound As Double) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| sourcePathID | String; ByVal; обязательный | ID ранее объявленного пути, который читается без изменения. |
| startNX | Double; ByVal; обязательный | X-компонента начальной поперечной нормали; линия не должна быть параллельна плоскости. |
| startNY | Double; ByVal; обязательный | Y-компонента начальной поперечной нормали. |
| startBound | Double; ByVal; обязательный | Значение начальной плоскости startNX*x+startNY*y=startBound; мм при единичной нормали. |
| endNX | Double; ByVal; обязательный | X-компонента конечной поперечной нормали. |
| endNY | Double; ByVal; обязательный | Y-компонента конечной поперечной нормали. |
| endBound | Double; ByVal; обязательный | Значение конечной плоскости endNX*x+endNY*y=endBound; мм при единичной нормали. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:128

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeOperationTests с его подготовленными локальными объектами/входом.

```vb
    Set path = facade.ConstrainLine("Trimmed", "Axis", 1#, 0#, 10#, 1#, 0#, 80#)
```

Полный контекст: tests/modTestRecipeOperations.bas:24.

Native entrypoint: modTestRecipeOperations.RunRecipeOperationTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-csectiondefinitionbuilder-lastgeometrydiagnostic"></a>
### CSectionDefinitionBuilder.LastGeometryDiagnostic

Возвращает последнюю диагностику query фасада с SourceMap. Nothing означает отсутствие записанного отказа.

```vb
Public Property Get LastGeometryDiagnostic() As CGeometryDiagnostic
```

Тип результата: CGeometryDiagnostic. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:142

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    Set diagnostic = facade.LastGeometryDiagnostic
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:155.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-csectiondefinitionbuilder-chain"></a>
### CSectionDefinitionBuilder.Chain

Выбирает непрерывный участок между двумя именованными сегментными привязками. Направление обхода явно; fullLoopIfSame задаёт полный оборот при совпавших концах.

```vb
Public Function Chain(ByVal pathID As String, ByVal sourcePathID As String, ByVal startSegmentID As String, ByVal startParameter As Double, ByVal endSegmentID As String, ByVal endParameter As Double, Optional ByVal forward As Boolean = True, Optional ByVal fullLoopIfSame As Boolean = False) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| sourcePathID | String; ByVal; обязательный | ID ранее объявленного пути, который читается без изменения. |
| startSegmentID | String; ByVal; обязательный | Сегмент начальной привязки цепочки. |
| startParameter | Double; ByVal; обязательный | Параметр начальной точки выбранного сегмента [0;1]. |
| endSegmentID | String; ByVal; обязательный | Сегмент конечной привязки цепочки. |
| endParameter | Double; ByVal; обязательный | Параметр конечной точки выбранного сегмента [0;1]. |
| forward | Boolean; ByVal; Optional; default=True | Явное направление следования исходному обходу; False — противоположное. |
| fullLoopIfSame | Boolean; ByVal; Optional; default=False | Полный оборот при совпавших привязках замкнутого пути; иначе нулевая цепочка недопустима. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:154

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.RunPathAnchorTests с его подготовленными локальными объектами/входом.

```vb
    Set part = facade.Chain("Side", "Outer", "Edge.BR", 0#, "Edge.TR", 0#)
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:103.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-csectiondefinitionbuilder-transform"></a>
### CSectionDefinitionBuilder.Transform

Создаёт новый путь из существующего изометрией; сохраняет сегментные ID и происхождение, отражение меняет знак sweep.

```vb
Public Function Transform(ByVal pathID As String, ByVal sourcePathID As String, ByVal shiftX As Double, ByVal shiftY As Double, Optional ByVal rotation As Double = 0#, Optional ByVal reflectX As Boolean = False) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| sourcePathID | String; ByVal; обязательный | ID ранее объявленного пути, который читается без изменения. |
| shiftX | Double; ByVal; обязательный | Перенос по X, мм. |
| shiftY | Double; ByVal; обязательный | Перенос по Y, мм. |
| rotation | Double; ByVal; Optional; default=0# | Угол поворота вокруг (0,0), радианы, CCW положителен. |
| reflectX | Boolean; ByVal; Optional; default=False | Отражение относительно локальной оси X перед поворотом и переносом. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:166

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.RunPathAnchorTests с его подготовленными локальными объектами/входом.

```vb
    Set changed = facade.Transform("MovedSide", "Side", 1000000#, -1000000#)
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:104.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-csectiondefinitionbuilder-rectanglepath"></a>
### CSectionDefinitionBuilder.RectanglePath

Создаёт замкнутую наружную границу одного прямоугольника от (x,y) с положительными width/height.

```vb
Public Function RectanglePath(ByVal pathID As String, ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal height As Double) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| x | Double; ByVal; обязательный | Модельная координата x, мм; знаки соответствуют section XY/WCS. |
| y | Double; ByVal; обязательный | Модельная координата y, мм; знаки соответствуют section XY/WCS. |
| width | Double; ByVal; обязательный | Координата/размер width, мм; размеры width/height положительны. |
| height | Double; ByVal; обязательный | Координата/размер height, мм; размеры width/height положительны. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:176

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02PolygonApiTests с его подготовленными локальными объектами/входом.

```vb
    Set path = facade.RectanglePath("Other", 200#, 0#, 100#, 100#)
```

Полный контекст: tests/modTestAudit02Api.bas:50.

Native entrypoint: modTestAudit02Api.RunAudit02PolygonApiTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02PolygonApiTests.txt.

<a id="api-csectiondefinitionbuilder-adjacentrectangles"></a>
### CSectionDefinitionBuilder.AdjacentRectangles

Создаёт одну границу смежных неперекрывающихся прямоугольников. Произвольное пересекающееся CAD Union не выполняется.

```vb
Public Function AdjacentRectangles(ByVal pathID As String, ByVal rectangles As Variant) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| rectangles | Variant; ByVal; обязательный | Array(Array(x,y,width,height),...); мм, размеры положительные, прямоугольники смежные без перекрытий. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:184

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    Set path = facade.AdjacentRectangles("Concrete", Array(Array(-200#, -200#, 400#, 400#)))
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:99.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-csectiondefinitionbuilder-pathgroup"></a>
### CSectionDefinitionBuilder.PathGroup

Объявляет массив ID независимых путей; порядок значим, фиктивные соединения отсутствуют.

```vb
Public Sub PathGroup(ByVal groupID As String, ByVal pathIDs As Variant)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| groupID | String; ByVal; обязательный | Смысловой ID группы; путь и арматурная annotation-group — разные контракты. |
| pathIDs | Variant; ByVal; обязательный | Массив ID существующих независимых путей; без повторов. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:191

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    facade.PathGroup "Rows", Array("Row", "Row2")
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:103.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-csectiondefinitionbuilder-outer"></a>
### CSectionDefinitionBuilder.Outer

Назначает замкнутому пути роль наружной границы материала. Не требует наличия арматуры.

```vb
Public Sub Outer(ByVal loopID As String, ByVal pathID As String, Optional ByVal comment As String = "", Optional ByVal contourIDs As Variant, Optional ByVal arcComment As String = "", Optional ByVal namingPolicy As EContourNamingPolicy = PreserveSegmentIDs, Optional ByVal namingPrefix As String = "")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| loopID | String; ByVal; обязательный | Непустой устойчивый ID контура; у CConcreteRegion численный индекс loop 1..LoopCount. |
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| comment | String; ByVal; Optional; default="" | Пояснение происхождения/инженерного смысла; не управляющий статус. |
| contourIDs | Variant; ByVal; Optional; пропущенный Variant=Missing | Необязательный массив ID сегментов сохранённого контура, по одному на каждый сегмент пути. |
| arcComment | String; ByVal; Optional; default="" | Отдельный комментарий аналитических дуг; пустой использует обычный comment. |
| namingPolicy | EContourNamingPolicy; ByVal; Optional; default=PreserveSegmentIDs | PreserveSegmentIDs сохраняет имена сегментов/явные contourIDs; SequentialByKind создаёт префикс_LINE_i/ARC_i с общим номером сегмента. Явные ID несовместимы с SequentialByKind. |
| namingPrefix | String; ByVal; Optional; default="" | Префикс только SequentialByKind; пустой использует CONTOUR. Разным границам нужны разные префиксы; рецепты Hollow явно задают CONTOUR_OUTER/CONTOUR_OPENING. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:197

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.BuildAudit02OperationsDefinition с его подготовленными локальными объектами/входом.

```vb
    facade.Outer "Outer", path.ID
```

Полный контекст: tests/modTestAudit02Api.bas:289.

Native entrypoint: modTestAudit02Api.RunAudit02GuideExamplesTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/FinalApiAcceptance/modTestAudit02Api.RunAudit02GuideExamplesTests.txt.

<a id="api-csectiondefinitionbuilder-opening"></a>
### CSectionDefinitionBuilder.Opening

Назначает замкнутому пути роль отверстия с конкретным parentOuterID. Действительное вложение проверяется при Finish.

```vb
Public Sub Opening(ByVal loopID As String, ByVal pathID As String, ByVal parentOuterID As String, Optional ByVal comment As String = "", Optional ByVal contourIDs As Variant, Optional ByVal arcComment As String = "", Optional ByVal namingPolicy As EContourNamingPolicy = PreserveSegmentIDs, Optional ByVal namingPrefix As String = "")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| loopID | String; ByVal; обязательный | Непустой устойчивый ID контура; у CConcreteRegion численный индекс loop 1..LoopCount. |
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| parentOuterID | String; ByVal; обязательный | ID ранее объявленной наружной границы для Opening; проверяется реальное вложение. |
| comment | String; ByVal; Optional; default="" | Пояснение происхождения/инженерного смысла; не управляющий статус. |
| contourIDs | Variant; ByVal; Optional; пропущенный Variant=Missing | Необязательный массив ID сегментов сохранённого контура, по одному на каждый сегмент пути. |
| arcComment | String; ByVal; Optional; default="" | Отдельный комментарий аналитических дуг; пустой использует обычный comment. |
| namingPolicy | EContourNamingPolicy; ByVal; Optional; default=PreserveSegmentIDs | PreserveSegmentIDs сохраняет имена сегментов/явные contourIDs; SequentialByKind создаёт префикс_LINE_i/ARC_i с общим номером сегмента. Явные ID несовместимы с SequentialByKind. |
| namingPrefix | String; ByVal; Optional; default="" | Префикс только SequentialByKind; пустой использует CONTOUR. Разным границам нужны разные префиксы; рецепты Hollow явно задают CONTOUR_OUTER/CONTOUR_OPENING. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:210

Минимальный вызов; требуемый контекст: В теле modTestGeometry.TestHollowRectangleGeometry с его подготовленными локальными объектами/входом.

```vb
    facade.Opening "VOID", path.ID, "OUTER"
```

Полный контекст: tests/modTestGeometry.bas:1969.

Native entrypoint: modTestGeometry.RunGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-csectiondefinitionbuilder-rebarbycount"></a>
### CSectionDefinitionBuilder.RebarByCount

Задаёт места для указанного числа стержней на рабочей части линии размещения. EndMode определяет включение концов. В группе это число применяется отдельно к каждой линии; нулевое число стержней или нулевой диаметр отключают ряд.

```vb
Public Sub RebarByCount(ByVal ruleID As String, ByVal target As String, ByVal count As Long, ByVal diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal steelClass As String = "Rebar")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |
| target | String; ByVal; обязательный | ID пути/PathGroup, разрешаемый в этом определении. |
| count | Long; ByVal; обязательный | Число стержней на одной линии размещения; целое неотрицательное. В группе применяется отдельно к каждой линии. Ноль отключает ряд. |
| diameter | Double; ByVal; обязательный | Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд. |
| options | CRebarRuleOptions; ByVal; Optional; default=Nothing | CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт реальные default options. |
| steelClass | String; ByVal; Optional; default="Rebar" | Строка принадлежности материала арматуры; совпадения разных материалов требуют явной политики. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:224

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
    facade.RebarByCount "Base", "Row", 3, 10#
```

Полный контекст: tests/modTestRecipeOperations.bas:131.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-csectiondefinitionbuilder-rebarbymaxstep"></a>
### CSectionDefinitionBuilder.RebarByMaxStep

Задаёт места стержней так, чтобы расстояние между соседними стержнями вдоль линии не превышало step. Исполнитель подбирает число стержней; фактическое расстояние может быть меньше заданного.

```vb
Public Sub RebarByMaxStep(ByVal ruleID As String, ByVal target As String, ByVal step As Double, ByVal diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal steelClass As String = "Rebar")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |
| target | String; ByVal; обязательный | ID пути/PathGroup, разрешаемый в этом определении. |
| step | Double; ByVal; обязательный | Расстояние между соседними стержнями вдоль линии размещения, мм; строго положительное. |
| diameter | Double; ByVal; обязательный | Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд. |
| options | CRebarRuleOptions; ByVal; Optional; default=Nothing | CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт реальные default options. |
| steelClass | String; ByVal; Optional; default="Rebar" | Строка принадлежности материала арматуры; совпадения разных материалов требуют явной политики. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:235

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.RunRecipeRuleTests с его подготовленными локальными объектами/входом.

```vb
    facade.RebarByMaxStep "Step", "Row", 30#, 10#, options
```

Полный контекст: tests/modTestSectionRecipes.bas:198.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-csectiondefinitionbuilder-rebarbyfixedstep"></a>
### CSectionDefinitionBuilder.RebarByFixedStep

Задаёт точное расстояние step между соседними стержнями вдоль линии. RemainderPlacement определяет, где оставить неиспользованный остаток длины. При IncludeBoth или ClosedUniform рабочая длина должна соответствовать заданному расстоянию.

```vb
Public Sub RebarByFixedStep(ByVal ruleID As String, ByVal target As String, ByVal step As Double, ByVal diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal steelClass As String = "Rebar")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |
| target | String; ByVal; обязательный | ID пути/PathGroup, разрешаемый в этом определении. |
| step | Double; ByVal; обязательный | Расстояние между соседними стержнями вдоль линии размещения, мм; строго положительное. |
| diameter | Double; ByVal; обязательный | Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд. |
| options | CRebarRuleOptions; ByVal; Optional; default=Nothing | CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт реальные default options. |
| steelClass | String; ByVal; Optional; default="Rebar" | Строка принадлежности материала арматуры; совпадения разных материалов требуют явной политики. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:244

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.RunRecipeRuleTests с его подготовленными локальными объектами/входом.

```vb
    facade.RebarByFixedStep "Fixed", "Row", 30#, 10#, options
```

Полный контекст: tests/modTestSectionRecipes.bas:205.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-csectiondefinitionbuilder-rebaratstations"></a>
### CSectionDefinitionBuilder.RebarAtStations

Задаёт отдельные места стержней расстояниями от начала линии размещения, в мм. Порядок этих расстояний в массиве определяет исходную нумерацию стержней.

```vb
Public Sub RebarAtStations(ByVal ruleID As String, ByVal target As String, ByVal stations As Variant, ByVal diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal steelClass As String = "Rebar")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |
| target | String; ByVal; обязательный | ID пути/PathGroup, разрешаемый в этом определении. |
| stations | Variant; ByVal; обязательный | Массив расстояний от начала линии размещения до каждого стержня, мм; порядок элементов задаёт исходную нумерацию. |
| diameter | Double; ByVal; обязательный | Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд. |
| options | CRebarRuleOptions; ByVal; Optional; default=Nothing | CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт реальные default options. |
| steelClass | String; ByVal; Optional; default="Rebar" | Строка принадлежности материала арматуры; совпадения разных материалов требуют явной политики. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:252

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
    facade.RebarAtStations "Dense", "Row", dense, 20#
```

Полный контекст: tests/modTestRecipeOperations.bas:156.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-csectiondefinitionbuilder-rebardependent"></a>
### CSectionDefinitionBuilder.RebarDependent

Задаёт дополнительный ряд смещением относительно центров стержней исходного ряда. TangentOffset задаёт смещение вдоль касательной, NormalOffset — поперёк линии. Число и места стержней связаны с исходным рядом, новое равномерное распределение не выполняется.

```vb
Public Sub RebarDependent(ByVal ruleID As String, ByVal sourceRuleID As String, ByVal diameter As Double, ByVal options As CRebarRuleOptions, Optional ByVal steelClass As String = "Rebar")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |
| sourceRuleID | String; ByVal; обязательный | ID исходного правила зависимости; циклы запрещены. |
| diameter | Double; ByVal; обязательный | Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд. |
| options | CRebarRuleOptions; ByVal; обязательный | CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт реальные default options. |
| steelClass | String; ByVal; Optional; default="Rebar" | Строка принадлежности материала арматуры; совпадения разных материалов требуют явной политики. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:261

Минимальный вызов; требуемый контекст: В теле modTestGeometryReview.TestPerPath с его подготовленными локальными объектами/входом.

```vb
    facade.RebarDependent "Extra", "Main", 8#, options
```

Полный контекст: tests/modTestGeometryReview.bas:371.

Native entrypoint: modTestGeometryReview.RunGeometryReviewTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/SolidFinal_2026-10-14/PublicationTests/modTestGeometryReview.RunGeometryReviewTests.txt.

<a id="api-csectiondefinitionbuilder-rebarprojected"></a>
### CSectionDefinitionBuilder.RebarProjected

Задаёт новый ряд проекцией центров стержней исходного ряда на целевую линию target. Направление — нормаль исходной линии или общий вектор DirectionX/Y. StartSetback и EndSetback ограничивают рабочий участок целевой линии; отсутствие пересечения отклоняется либо пропускается по SkipMissingProjection.

```vb
Public Sub RebarProjected(ByVal ruleID As String, ByVal target As String, ByVal sourceRuleID As String, ByVal diameter As Double, ByVal options As CRebarRuleOptions, Optional ByVal steelClass As String = "Rebar")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |
| target | String; ByVal; обязательный | ID пути/PathGroup, разрешаемый в этом определении. |
| sourceRuleID | String; ByVal; обязательный | ID исходного правила зависимости; циклы запрещены. |
| diameter | Double; ByVal; обязательный | Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд. |
| options | CRebarRuleOptions; ByVal; обязательный | CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт реальные default options. |
| steelClass | String; ByVal; Optional; default="Rebar" | Строка принадлежности материала арматуры; совпадения разных материалов требуют явной политики. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:270

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
    facade.RebarProjected "Projected", "Target", "Base", 10#, options
```

Полный контекст: tests/modTestRecipeOperations.bas:135.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-csectiondefinitionbuilder-verticalextent"></a>
### CSectionDefinitionBuilder.VerticalExtent

Измеряет полный вертикальный габарит выбранного пути/группы с точными экстремумами дуг. placementX и normalX задают сторону отображения.

```vb
Public Sub VerticalExtent(ByVal annotationID As String, ByVal target As String, ByVal name As String, ByVal placementX As Double, Optional ByVal normalX As Double = -1#, Optional ByVal comment As String = "")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| annotationID | String; ByVal; обязательный | Непустой уникальный ID аннотации. |
| target | String; ByVal; обязательный | ID пути/PathGroup, разрешаемый в этом определении. |
| name | String; ByVal; обязательный | Обозначение измерения или имя контейнера. |
| placementX | Double; ByVal; обязательный | Модельная X размерной линии, мм. |
| normalX | Double; ByVal; Optional; default=-1# | X-компонента стороны размерной/подписной линии; layout нормирует вектор. |
| comment | String; ByVal; Optional; default="" | Пояснение происхождения/инженерного смысла; не управляющий статус. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:290

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.AdvancedGuideDefinition с его подготовленными локальными объектами/входом.

```vb
    facade.VerticalExtent "Height", "Concrete", "H", 740#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:84.

Native entrypoint: modTestRecipeApiCatalogue.RunAdvancedGuideExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/AdvancedExamples.txt.

<a id="api-csectiondefinitionbuilder-horizontalextent"></a>
### CSectionDefinitionBuilder.HorizontalExtent

Измеряет полный горизонтальный габарит выбранного пути/группы. Это не длина пути и не исходная ширина до скругления.

```vb
Public Sub HorizontalExtent(ByVal annotationID As String, ByVal target As String, ByVal name As String, ByVal placementY As Double, Optional ByVal normalY As Double = 1#, Optional ByVal comment As String = "")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| annotationID | String; ByVal; обязательный | Непустой уникальный ID аннотации. |
| target | String; ByVal; обязательный | ID пути/PathGroup, разрешаемый в этом определении. |
| name | String; ByVal; обязательный | Обозначение измерения или имя контейнера. |
| placementY | Double; ByVal; обязательный | Модельная Y размерной линии, мм. |
| normalY | Double; ByVal; Optional; default=1# | Y-компонента стороны размерной/подписной линии. |
| comment | String; ByVal; Optional; default="" | Пояснение происхождения/инженерного смысла; не управляющий статус. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:301

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestGeneratedContoursWithoutAnnotations с его подготовленными локальными объектами/входом.

```vb
        facade.HorizontalExtent "DIM_B", "Boundary", "B", 0#
```

Полный контекст: tests/modTestGeometryQuery.bas:797.

Native entrypoint: modTestGeometryQuery.RunGeneratedContourOwnershipTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-csectiondefinitionbuilder-horizontaldistance"></a>
### CSectionDefinitionBuilder.HorizontalDistance

Объявляет размер между двумя заданными X на уровне placementY. DisplayIncrement округляет только подпись, не геометрию.

```vb
Public Sub HorizontalDistance(ByVal annotationID As String, ByVal name As String, ByVal startX As Double, ByVal endX As Double, ByVal placementY As Double, Optional ByVal normalY As Double = 1#, Optional ByVal comment As String = "", Optional ByVal displayIncrement As Double = 0#)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| annotationID | String; ByVal; обязательный | Непустой уникальный ID аннотации. |
| name | String; ByVal; обязательный | Обозначение измерения или имя контейнера. |
| startX | Double; ByVal; обязательный | Модельная координата startX, мм; знаки соответствуют section XY/WCS. |
| endX | Double; ByVal; обязательный | Модельная координата endX, мм; знаки соответствуют section XY/WCS. |
| placementY | Double; ByVal; обязательный | Модельная Y размерной линии, мм. |
| normalY | Double; ByVal; Optional; default=1# | Y-компонента стороны размерной/подписной линии. |
| comment | String; ByVal; Optional; default="" | Пояснение происхождения/инженерного смысла; не управляющий статус. |
| displayIncrement | Double; ByVal; Optional; default=0# | Неотрицательный шаг округления только видимого справочного значения, мм; 0 означает отсутствие такого округления. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:312

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.TestAnnotationMeasurements с его подготовленными локальными объектами/входом.

```vb
    facade.HorizontalDistance "EXACT", "Exact", 0#, 123.6, 200#
```

Полный контекст: tests/modTestSectionRecipes.bas:355.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-csectiondefinitionbuilder-verticaldistance"></a>
### CSectionDefinitionBuilder.VerticalDistance

Объявляет вертикальный размер между двумя Y на линии placementX; значение — модуль разности.

```vb
Public Sub VerticalDistance(ByVal annotationID As String, ByVal name As String, ByVal startY As Double, ByVal endY As Double, ByVal placementX As Double, Optional ByVal normalX As Double = -1#, Optional ByVal comment As String = "")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| annotationID | String; ByVal; обязательный | Непустой уникальный ID аннотации. |
| name | String; ByVal; обязательный | Обозначение измерения или имя контейнера. |
| startY | Double; ByVal; обязательный | Модельная координата startY, мм; знаки соответствуют section XY/WCS. |
| endY | Double; ByVal; обязательный | Модельная координата endY, мм; знаки соответствуют section XY/WCS. |
| placementX | Double; ByVal; обязательный | Модельная X размерной линии, мм. |
| normalX | Double; ByVal; Optional; default=-1# | X-компонента стороны размерной/подписной линии; layout нормирует вектор. |
| comment | String; ByVal; Optional; default="" | Пояснение происхождения/инженерного смысла; не управляющий статус. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:325

Минимальный вызов; требуемый контекст: В теле modTestRecipeAnnotations.AnnotationFacade с его подготовленными локальными объектами/входом.

```vb
    If previewPlacement Then facade.VerticalDistance "VerticalExplicit", "v", -100#, 0#, 150#, 1# Else facade.VerticalDistance "VerticalExplicit", "v", -100#, 0#, -50#
```

Полный контекст: tests/modTestRecipeAnnotations.bas:160.

Native entrypoint: modTestRecipeAnnotations.RunRecipeAnnotationTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/DeclarativeApi/FinalApiAnnotations.txt.

Native entrypoint: modTestRecipeAnnotations.RunRecipeAnnotationSnapshotTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/DeclarativeApi/Operations.txt.

<a id="api-csectiondefinitionbuilder-dimension"></a>
### CSectionDefinitionBuilder.Dimension

Измеряет между двумя именованными точками: Horizontal/Vertical проекция или Aligned расстояние. Проекция размещается через первую точку.

```vb
Public Sub Dimension(ByVal annotationID As String, ByVal startAnchorID As String, ByVal endAnchorID As String, ByVal name As String, Optional ByVal measurement As String = "Aligned", Optional ByVal normalX As Double = 0#, Optional ByVal normalY As Double = 0#, Optional ByVal comment As String = "")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| annotationID | String; ByVal; обязательный | Непустой уникальный ID аннотации. |
| startAnchorID | String; ByVal; обязательный | ID начальной модельной точки измерения. |
| endAnchorID | String; ByVal; обязательный | ID конечной модельной точки измерения. |
| name | String; ByVal; обязательный | Обозначение измерения или имя контейнера. |
| measurement | String; ByVal; Optional; default="Aligned" | Ровно Horizontal, Vertical или Aligned; значение чувствительно к указанному контракту. |
| normalX | Double; ByVal; Optional; default=0# | X-компонента стороны размерной/подписной линии; layout нормирует вектор. |
| normalY | Double; ByVal; Optional; default=0# | Y-компонента стороны размерной/подписной линии. |
| comment | String; ByVal; Optional; default="" | Пояснение происхождения/инженерного смысла; не управляющий статус. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:337

Минимальный вызов; требуемый контекст: В теле modTestRecipeAnnotations.AnnotationFacade с его подготовленными локальными объектами/входом.

```vb
    facade.Dimension "Diagonal", "A", "B", "s"
```

Полный контекст: tests/modTestRecipeAnnotations.bas:159.

Native entrypoint: modTestRecipeAnnotations.RunRecipeAnnotationTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/DeclarativeApi/FinalApiAnnotations.txt.

Native entrypoint: modTestRecipeAnnotations.RunRecipeAnnotationSnapshotTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/DeclarativeApi/Operations.txt.

<a id="api-csectiondefinitionbuilder-pathlength"></a>
### CSectionDefinitionBuilder.PathLength

Объявляет точную суммарную длину target-путей и текстовую привязку; не создаёт фиктивную соединяющую линию группы.

```vb
Public Sub PathLength(ByVal annotationID As String, ByVal target As String, ByVal placementAnchorID As String, Optional ByVal name As String = "L", Optional ByVal comment As String = "")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| annotationID | String; ByVal; обязательный | Непустой уникальный ID аннотации. |
| target | String; ByVal; обязательный | ID пути/PathGroup, разрешаемый в этом определении. |
| placementAnchorID | String; ByVal; обязательный | ID точки размещения текстового измерения. |
| name | String; ByVal; Optional; default="L" | Обозначение измерения или имя контейнера. |
| comment | String; ByVal; Optional; default="" | Пояснение происхождения/инженерного смысла; не управляющий статус. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:352

Минимальный вызов; требуемый контекст: В теле modTestRecipeAnnotations.AnnotationFacade с его подготовленными локальными объектами/входом.

```vb
    facade.PathLength "Length", "Boundary", "LengthNote"
```

Полный контекст: tests/modTestRecipeAnnotations.bas:171.

Native entrypoint: modTestRecipeAnnotations.RunRecipeAnnotationTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/DeclarativeApi/FinalApiAnnotations.txt.

Native entrypoint: modTestRecipeAnnotations.RunRecipeAnnotationSnapshotTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/DeclarativeApi/Operations.txt.

<a id="api-csectiondefinitionbuilder-radius"></a>
### CSectionDefinitionBuilder.Radius

Измеряет радиус конкретной Arc/Circle; параметр выбирает конечную точку радиуса. Прямая — ошибка.

```vb
Public Sub Radius(ByVal annotationID As String, ByVal pathID As String, ByVal segmentID As String, Optional ByVal parameter As Double = 0.5, Optional ByVal name As String = "R", Optional ByVal comment As String = "")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| annotationID | String; ByVal; обязательный | Непустой уникальный ID аннотации. |
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| segmentID | String; ByVal; обязательный | Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом. |
| parameter | Double; ByVal; Optional; default=0.5 | Нормированный параметр точки сегмента [0;1], не расстояние по всему пути. |
| name | String; ByVal; Optional; default="R" | Обозначение измерения или имя контейнера. |
| comment | String; ByVal; Optional; default="" | Пояснение происхождения/инженерного смысла; не управляющий статус. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:363

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.BuildAudit02OperationsDefinition с его подготовленными локальными объектами/входом.

```vb
    facade.Radius "Radius", "Arc", "Arc", 0.5, "R"
```

Полный контекст: tests/modTestAudit02Api.bas:321.

Native entrypoint: modTestAudit02Api.RunAudit02GuideExamplesTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/FinalApiAcceptance/modTestAudit02Api.RunAudit02GuideExamplesTests.txt.

<a id="api-csectiondefinitionbuilder-diameter"></a>
### CSectionDefinitionBuilder.Diameter

Измеряет диаметр конкретной Arc/Circle с противоположным концом на той же окружности. Прямая — ошибка.

```vb
Public Sub Diameter(ByVal annotationID As String, ByVal pathID As String, ByVal segmentID As String, Optional ByVal parameter As Double = 0.5, Optional ByVal name As String = "D", Optional ByVal comment As String = "")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| annotationID | String; ByVal; обязательный | Непустой уникальный ID аннотации. |
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| segmentID | String; ByVal; обязательный | Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом. |
| parameter | Double; ByVal; Optional; default=0.5 | Нормированный параметр точки сегмента [0;1], не расстояние по всему пути. |
| name | String; ByVal; Optional; default="D" | Обозначение измерения или имя контейнера. |
| comment | String; ByVal; Optional; default="" | Пояснение происхождения/инженерного смысла; не управляющий статус. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:370

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.BuildAudit02OperationsDefinition с его подготовленными локальными объектами/входом.

```vb
    facade.Diameter "Diameter", "Hole", "Hole", 0.5, "D"
```

Полный контекст: tests/modTestAudit02Api.bas:322.

Native entrypoint: modTestAudit02Api.RunAudit02GuideExamplesTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/FinalApiAcceptance/modTestAudit02Api.RunAudit02GuideExamplesTests.txt.

<a id="api-csectiondefinitionbuilder-text"></a>
### CSectionDefinitionBuilder.Text

Объявляет простой текст на именованной модельной привязке; пиксели/шрифт задаёт общий layout.

```vb
Public Sub Text(ByVal annotationID As String, ByVal anchorID As String, ByVal textValue As String, Optional ByVal comment As String = "")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| annotationID | String; ByVal; обязательный | Непустой уникальный ID аннотации. |
| anchorID | String; ByVal; обязательный | Непустой уникальный ID модельной привязки. |
| textValue | String; ByVal; обязательный | Текст подписи; Unicode сохраняется. |
| comment | String; ByVal; Optional; default="" | Пояснение происхождения/инженерного смысла; не управляющий статус. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:388

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.BuildAudit02OperationsDefinition с его подготовленными локальными объектами/входом.

```vb
    facade.Text "Note", "MovedA", "Учебный пример"
```

Полный контекст: tests/modTestAudit02Api.bas:323.

Native entrypoint: modTestAudit02Api.RunAudit02GuideExamplesTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/FinalApiAcceptance/modTestAudit02Api.RunAudit02GuideExamplesTests.txt.

<a id="api-csectiondefinitionbuilder-rebarlabel"></a>
### CSectionDefinitionBuilder.RebarLabel

Привязывает фактическую подпись группы к габариту target-пути, со стороны normalX/Y.

```vb
Public Sub RebarLabel(ByVal groupID As String, ByVal target As String, ByVal normalX As Double, ByVal normalY As Double)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| groupID | String; ByVal; обязательный | Смысловой ID группы; путь и арматурная annotation-group — разные контракты. |
| target | String; ByVal; обязательный | ID пути/PathGroup, разрешаемый в этом определении. |
| normalX | Double; ByVal; обязательный | X-компонента стороны размерной/подписной линии; layout нормирует вектор. |
| normalY | Double; ByVal; обязательный | Y-компонента стороны размерной/подписной линии. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:398

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeExamples.RoundedFrameDefinition с его подготовленными локальными объектами/входом.

```vb
    facade.RebarLabel "OuterBars", "Outer.Axis", 0#, 1#
```

Полный контекст: tests/modTestDeclarativeExamples.bas:120.

Native entrypoint: modTestDeclarativeExamples.RunRoundedFrameExampleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/OneClassExample/InputAndBuild.txt.

<a id="api-csectiondefinitionbuilder-rebarrowlabel"></a>
### CSectionDefinitionBuilder.RebarRowLabel

Привязывает фактическую подпись группы к концам исполненного ruleID.

```vb
Public Sub RebarRowLabel(ByVal groupID As String, ByVal ruleID As String, ByVal normalX As Double, ByVal normalY As Double)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| groupID | String; ByVal; обязательный | Смысловой ID группы; путь и арматурная annotation-group — разные контракты. |
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |
| normalX | Double; ByVal; обязательный | X-компонента стороны размерной/подписной линии; layout нормирует вектор. |
| normalY | Double; ByVal; обязательный | Y-компонента стороны размерной/подписной линии. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:407

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.TestPublicationAndCoincidence с его подготовленными локальными объектами/входом.

```vb
    facade.RebarRowLabel "Base", "Base", 0#, -1#
```

Полный контекст: tests/modTestSectionRecipes.bas:383.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-csectiondefinitionbuilder-rebarrowextentlabel"></a>
### CSectionDefinitionBuilder.RebarRowExtentLabel

Привязывает подпись к фактическому габариту позиций правила; reverseDirection меняет ориентацию anchor-линии.

```vb
Public Sub RebarRowExtentLabel(ByVal groupID As String, ByVal ruleID As String, ByVal normalX As Double, ByVal normalY As Double, Optional ByVal reverseDirection As Boolean = False)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| groupID | String; ByVal; обязательный | Смысловой ID группы; путь и арматурная annotation-group — разные контракты. |
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |
| normalX | Double; ByVal; обязательный | X-компонента стороны размерной/подписной линии; layout нормирует вектор. |
| normalY | Double; ByVal; обязательный | Y-компонента стороны размерной/подписной линии. |
| reverseDirection | Boolean; ByVal; Optional; default=False | Обращает направление линии привязки подписи; физический ряд не меняется. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:416

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.BuildAudit02DemoDefinition с его подготовленными локальными объектами/входом.

```vb
    facade.RebarRowExtentLabel "H.Right", "H.Right.Main", 1#, 0#, True
```

Полный контекст: tests/modTestAudit02Api.bas:275.

Native entrypoint: modTestAudit02Api.RunAudit02GuideExamplesTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/FinalApiAcceptance/modTestAudit02Api.RunAudit02GuideExamplesTests.txt.

<a id="api-csectiondefinitionbuilder-finish"></a>
### CSectionDefinitionBuilder.Finish

Проверяет контуры, ссылки, зависимости и декларации аннотаций; публикует неизменяемое CSectionDefinition и закрывает фасад. Допустимость параметров и фактическое размещение арматуры проверяет CSectionBuilder до построения сетки.

```vb
Public Function Finish() As CSectionDefinition
```

Тип результата: CSectionDefinition. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:427

Минимальный вызов; требуемый контекст: В теле modTestGeometryPublication.RunDefinitionLifecycleTests с его подготовленными локальными объектами/входом.

```vb
    Set unused = facade.Finish
```

Полный контекст: tests/modTestGeometryPublication.bas:207.

Native entrypoint: modTestGeometryPublication.RunDefinitionLifecycleTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/SolidFinal_2026-10-14/SolidTests/modTestGeometryPublication.RunDefinitionLifecycleTests.txt.

<a id="api-csectiondefinitionbuilder-rebarrowextremelabel"></a>
### CSectionDefinitionBuilder.RebarRowExtremeLabel

Выбирает крайние созданные позиции в directionX/Y; сторона текста normalX/Y независима от выбора ряда.

```vb
Public Sub RebarRowExtremeLabel(ByVal groupID As String, ByVal ruleID As String, ByVal directionX As Double, ByVal directionY As Double, ByVal normalX As Double, ByVal normalY As Double)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| groupID | String; ByVal; обязательный | Смысловой ID группы; путь и арматурная annotation-group — разные контракты. |
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |
| directionX | Double; ByVal; обязательный | X-компонента направления проекции/выбора экстремума; вектор не нулевой, у RebarRowExtremeLabel единичный. |
| directionY | Double; ByVal; обязательный | Y-компонента направления проекции/выбора экстремума. |
| normalX | Double; ByVal; обязательный | X-компонента стороны размерной/подписной линии; layout нормирует вектор. |
| normalY | Double; ByVal; обязательный | Y-компонента стороны размерной/подписной линии. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Создающие методы уже регистрируют путь; AddPath нужен только для отдельного подготовленного пути. Ссылки требуют ранее объявленных путей/привязок; зависимости разрешаются при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:436

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.TestProjectionNormalsAndOrdinals с его подготовленными локальными объектами/входом.

```vb
    facade.RebarRowExtremeLabel "CurveBars", "CurveMain", 0#, 1#, 0#, 1#
```

Полный контекст: tests/modTestSectionRecipes.bas:286.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

## CGeometryPath

New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

<a id="api-cgeometrypath-lastdiagnostic"></a>
### CGeometryPath.LastDiagnostic

Последняя машинная диагностика; Nothing при отсутствии записанного события.

```vb
Public Property Get LastDiagnostic() As CGeometryDiagnostic
```

Тип результата: CGeometryDiagnostic. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:27

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    Set diagnostic = path.LastDiagnostic: Check "path.reason", Len(diagnostic.Reason) > 0
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:143.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-cgeometrypath-initialize"></a>
### CGeometryPath.Initialize

Создаёт mutable путь с непустым ID один раз.

```vb
Public Sub Initialize(ByVal pathID As String)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:32

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.AdvancedGuideDefinition с его подготовленными локальными объектами/входом.

```vb
    Set path = New CGeometryPath: path.Initialize "Nose"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:47.

Native entrypoint: modTestRecipeApiCatalogue.RunAdvancedGuideExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/AdvancedExamples.txt.

<a id="api-cgeometrypath-addline"></a>
### CGeometryPath.AddLine

Добавляет прямой сегмент с устойчивым ID; начало совпадает с предыдущим концом.

```vb
Public Sub AddLine(ByVal segmentID As String, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| segmentID | String; ByVal; обязательный | Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом. |
| x1 | Double; ByVal; обязательный | Модельная координата x1, мм; знаки соответствуют section XY/WCS. |
| y1 | Double; ByVal; обязательный | Модельная координата y1, мм; знаки соответствуют section XY/WCS. |
| x2 | Double; ByVal; обязательный | Модельная координата x2, мм; знаки соответствуют section XY/WCS. |
| y2 | Double; ByVal; обязательный | Модельная координата y2, мм; знаки соответствуют section XY/WCS. |

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:39

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.CompoundPath с его подготовленными локальными объектами/входом.

```vb
    path.AddLine "Bottom", 0#, 0#, 100#, 0#
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:789.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunPathTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/AnalyticPaths.txt.

<a id="api-cgeometrypath-addarc"></a>
### CGeometryPath.AddArc

Добавляет аналитическую дугу по центру/R/startAngle/sweep, без хордовой аппроксимации.

```vb
Public Sub AddArc(ByVal segmentID As String, ByVal centerX As Double, ByVal centerY As Double, ByVal radius As Double, ByVal startAngle As Double, ByVal sweep As Double)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| segmentID | String; ByVal; обязательный | Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом. |
| centerX | Double; ByVal; обязательный | Модельная координата centerX, мм; знаки соответствуют section XY/WCS. |
| centerY | Double; ByVal; обязательный | Модельная координата centerY, мм; знаки соответствуют section XY/WCS. |
| radius | Double; ByVal; обязательный | Положительный радиус, мм. |
| startAngle | Double; ByVal; обязательный | Начальный угол, радианы; от +X. |
| sweep | Double; ByVal; обязательный | Signed sweep, радианы: CCW положителен, CW отрицателен, 0 запрещён, \|sweep\|≤2*pi. |

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:48

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.InvalidPathRejected с его подготовленными локальными объектами/входом.

```vb
        Case 3: path.AddArc "ZeroSweep", 0#, 0#, 1#, 0#, 0#
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:805.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunPathTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/AnalyticPaths.txt.

<a id="api-cgeometrypath-seal"></a>
### CGeometryPath.Seal

Публикует путь. closed=True проверяет совпадение концов, не достраивает соединение.

```vb
Public Sub Seal(Optional ByVal closed As Boolean = False)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| closed | Boolean; ByVal; Optional; default=False | Явный Boolean режима замыкания/касательной; не исправляет геометрию автоматически. |

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:63

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.TestJointTangents с его подготовленными локальными объектами/входом.

```vb
    path.Seal
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:721.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunPathTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/AnalyticPaths.txt.

<a id="api-cgeometrypath-id"></a>
### CGeometryPath.ID

Смысловой ID этого пути.

```vb
Public Property Get ID() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:79

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.BuildAudit02OperationsDefinition с его подготовленными локальными объектами/входом.

```vb
    facade.Outer "Outer", path.ID
```

Полный контекст: tests/modTestAudit02Api.bas:289.

Native entrypoint: modTestAudit02Api.RunAudit02GuideExamplesTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/FinalApiAcceptance/modTestAudit02Api.RunAudit02GuideExamplesTests.txt.

<a id="api-cgeometrypath-segmentcount"></a>
### CGeometryPath.SegmentCount

Число аналитических line/arc сегментов.

```vb
Public Property Get SegmentCount() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:81

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.RunRoundedPolygonTests с его подготовленными локальными объектами/входом.

```vb
    For i = 1 To path.SegmentCount
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:193.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunRoundedPolygonTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/RoundedPolygons.txt.

<a id="api-cgeometrypath-issealed"></a>
### CGeometryPath.IsSealed

Опубликован ли объект; published геометрия не изменяется.

```vb
Public Property Get IsSealed() As Boolean
```

Тип результата: Boolean. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:83

Минимальный вызов; требуемый контекст: В теле modTestGeometryPublication.RunDefinitionLifecycleTests с его подготовленными локальными объектами/входом.

```vb
    Check "path.creator.sealed", path.IsSealed
```

Полный контекст: tests/modTestGeometryPublication.bas:191.

Native entrypoint: modTestGeometryPublication.RunDefinitionLifecycleTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/SolidFinal_2026-10-14/SolidTests/modTestGeometryPublication.RunDefinitionLifecycleTests.txt.

<a id="api-cgeometrypath-isclosed"></a>
### CGeometryPath.IsClosed

Замкнут ли фактический путь, без фиктивного соединения.

```vb
Public Property Get IsClosed() As Boolean
```

Тип результата: Boolean. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:85

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.TestArcs с его подготовленными локальными объектами/входом.

```vb
    Check "circle.closed", path.IsClosed
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:676.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunPathTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/AnalyticPaths.txt.

<a id="api-cgeometrypath-length"></a>
### CGeometryPath.Length

Точная суммарная длина пути, мм.

```vb
Public Property Get Length() As Double
```

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:87

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.InvalidPathRejected с его подготовленными локальными объектами/входом.

```vb
        Case 6: value = path.Length
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:808.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunPathTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/AnalyticPaths.txt.

<a id="api-cgeometrypath-segmentid"></a>
### CGeometryPath.SegmentID

Возвращает смысловой ID сегмента по индексу.

```vb
Public Property Get SegmentID(ByVal index As Long) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:89

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.TestCompoundPath с его подготовленными локальными объектами/входом.

```vb
    Check "compound.slice.id", part.SegmentID(2) = "Corner"
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:703.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunPathTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/AnalyticPaths.txt.

<a id="api-cgeometrypath-sourcepathid"></a>
### CGeometryPath.SourcePathID

ID исходного пути сегмента после slice/transform/offset.

```vb
Public Property Get SourcePathID(ByVal index As Long) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:94

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.RunPathAnchorTests с его подготовленными локальными объектами/входом.

```vb
    Check "anchor.facade.source", changed.SourcePathID(1) = "Outer" And changed.SourceSegmentID(1) = "Edge.BR"
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:107.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-cgeometrypath-sourcesegmentid"></a>
### CGeometryPath.SourceSegmentID

Смысловой исходный ID сегмента, сохраняемый при распаде/переименовании результата.

```vb
Public Property Get SourceSegmentID(ByVal index As Long) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:99

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunAdvancedGuideExamples с его подготовленными локальными объектами/входом.

```vb
        Check "advanced.chain.namedNose." & CStr(useMaxStep), path.SourceSegmentID(3) = "NoseArc"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:18.

Native entrypoint: modTestRecipeApiCatalogue.RunAdvancedGuideExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/AdvancedExamples.txt.

<a id="api-cgeometrypath-sourcestartparameter"></a>
### CGeometryPath.SourceStartParameter

Начальный параметр исходного сегмента; может выйти за [0;1] при явном продлении прямой.

```vb
Public Property Get SourceStartParameter(ByVal index As Long) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:104

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.RunPathAnchorTests с его подготовленными локальными объектами/входом.

```vb
    Near "anchor.reverse.first", part.SourceStartParameter(1), 0.75
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:65.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-cgeometrypath-sourceendparameter"></a>
### CGeometryPath.SourceEndParameter

Конечный параметр исходного сегмента, включая обращённый обход.

```vb
Public Property Get SourceEndParameter(ByVal index As Long) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:109

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.RunPathAnchorTests с его подготовленными локальными объектами/входом.

```vb
    Near "anchor.wrap.first.end", part.SourceEndParameter(1), 1#
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:57.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-cgeometrypath-segmentindex"></a>
### CGeometryPath.SegmentIndex

Разрешает существующий смысловой ID в текущий индекс; исчезнувшая дуга даёт MISSING_SEGMENT.

```vb
Public Function SegmentIndex(ByVal segmentID As String) As Long
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| segmentID | String; ByVal; обязательный | Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом. |

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:114

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.InvalidAnchorRejected с его подготовленными локальными объектами/входом.

```vb
            station = path.SegmentIndex("Corner.BL")
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:145.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-cgeometrypath-anchorstation"></a>
### CGeometryPath.AnchorStation

Находит расстояние от начала всей линии до точки заданного сегмента, мм. Точка задаётся именем сегмента и параметром [0;1]. Если при построении участок исчез, обе его границы сохраняют привязку к оставшемуся стыку. Неизвестное имя вызывает ошибку.

```vb
Public Function AnchorStation(ByVal segmentID As String, Optional ByVal parameter As Double = 0#) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| segmentID | String; ByVal; обязательный | Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом. |
| parameter | Double; ByVal; Optional; default=0# | Нормированный параметр точки сегмента [0;1], не расстояние по всему пути. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:124

Минимальный вызов; требуемый контекст: В теле modTestGeometryReview.TestCollapsedBoundaries с его подготовленными локальными объектами/входом.

```vb
    station = path.AnchorStation("Corner.D")
```

Полный контекст: tests/modTestGeometryReview.bas:205.

Native entrypoint: modTestGeometryReview.RunRoundedEndpointTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/UnifiedTestsCorrected/modTestGeometryReview.RunRoundedEndpointTests.txt.

<a id="api-cgeometrypath-withid"></a>
### CGeometryPath.WithID

Копирует опубликованный путь с новым ID, сохраняя сегментные ссылки и исходное происхождение.

```vb
Public Function WithID(ByVal pathID As String) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:136

Минимальный вызов; требуемый контекст: В теле modTestGeometryReview.TestCollapsedBoundaries с его подготовленными локальными объектами/входом.

```vb
    Set copy = path.WithID("Copy")
```

Полный контекст: tests/modTestGeometryReview.bas:214.

Native entrypoint: modTestGeometryReview.RunRoundedEndpointTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/UnifiedTestsCorrected/modTestGeometryReview.RunRoundedEndpointTests.txt.

<a id="api-cgeometrypath-betweenanchors"></a>
### CGeometryPath.BetweenAnchors

Возвращает точную цепочку между привязками с явным направлением и правилом полного обхода.

```vb
Public Function BetweenAnchors(ByVal pathID As String, ByVal startSegmentID As String, ByVal startParameter As Double, ByVal endSegmentID As String, ByVal endParameter As Double, Optional ByVal forward As Boolean = True, Optional ByVal fullLoopIfSame As Boolean = False) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| startSegmentID | String; ByVal; обязательный | Сегмент начальной привязки цепочки. |
| startParameter | Double; ByVal; обязательный | Параметр начальной точки выбранного сегмента [0;1]. |
| endSegmentID | String; ByVal; обязательный | Сегмент конечной привязки цепочки. |
| endParameter | Double; ByVal; обязательный | Параметр конечной точки выбранного сегмента [0;1]. |
| forward | Boolean; ByVal; Optional; default=True | Явное направление следования исходному обходу; False — противоположное. |
| fullLoopIfSame | Boolean; ByVal; Optional; default=False | Полный оборот при совпавших привязках замкнутого пути; иначе нулевая цепочка недопустима. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:149

Минимальный вызов; требуемый контекст: В теле modTestGeometryReview.TestCollapsedBoundaries с его подготовленными локальными объектами/входом.

```vb
    Set copy = path.BetweenAnchors("Empty", "Edge.A", 0#, "Edge.A", 1#)
```

Полный контекст: tests/modTestGeometryReview.bas:226.

Native entrypoint: modTestGeometryReview.RunRoundedEndpointTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/UnifiedTestsCorrected/modTestGeometryReview.RunRoundedEndpointTests.txt.

<a id="api-cgeometrypath-atdistance"></a>
### CGeometryPath.AtDistance

Возвращает координаты точки на заданном расстоянии от начала линии и единичный вектор касательной в этой точке. На стыке можно выбрать касательную предыдущего сегмента.

```vb
Public Sub AtDistance(ByVal distance As Double, ByRef x As Double, ByRef y As Double, ByRef tangentX As Double, ByRef tangentY As Double, Optional ByVal incomingTangentAtJoint As Boolean = False)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| distance | Double; ByVal; обязательный | Расстояние или граница интервала distance, внутренние мм; направление и допустимые пределы определяет вызываемый метод. |
| x | Double; ByRef; обязательный | Модельная координата x, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| y | Double; ByRef; обязательный | Модельная координата y, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| tangentX | Double; ByRef; обязательный | Компонента единичной касательной, без размерности. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| tangentY | Double; ByRef; обязательный | Компонента единичной касательной, без размерности. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| incomingTangentAtJoint | Boolean; ByVal; Optional; default=False | Явный Boolean режима замыкания/касательной; не исправляет геометрию автоматически. |

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:177

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.RunPathAnchorTests с его подготовленными локальными объектами/входом.

```vb
    part.AtDistance 0#, x, y, tx, ty
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:47.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-cgeometrypath-projectionbounds"></a>
### CGeometryPath.ProjectionBounds

Возвращает действительные min/max проекции по нормированному направлению, с экстремумами дуг.

```vb
Public Sub ProjectionBounds(ByVal directionX As Double, ByVal directionY As Double, ByRef minimum As Double, ByRef maximum As Double)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| directionX | Double; ByVal; обязательный | X-компонента направления проекции/выбора экстремума; вектор не нулевой, у RebarRowExtremeLabel единичный. |
| directionY | Double; ByVal; обязательный | Y-компонента направления проекции/выбора экстремума. |
| minimum | Double; ByRef; обязательный | Расстояние или граница интервала minimum, внутренние мм; направление и допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| maximum | Double; ByRef; обязательный | Расстояние или граница интервала maximum, внутренние мм; направление и допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:200

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.TestLine с его подготовленными локальными объектами/входом.

```vb
    path.ProjectionBounds 2#, 0#, minimum, maximum
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:622.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunPathTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/AnalyticPaths.txt.

<a id="api-cgeometrypath-slice"></a>
### CGeometryPath.Slice

Обрезает полный непрерывный путь по длине; точные дуги и source parameters сохраняются.

```vb
Public Function Slice(ByVal startDistance As Double, ByVal endDistance As Double) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| startDistance | Double; ByVal; обязательный | Расстояние от начала всей линии, мм, в [0;Length]; endDistance строго больше startDistance. |
| endDistance | Double; ByVal; обязательный | Расстояние от начала всей линии, мм, в [0;Length]; endDistance строго больше startDistance. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:215

Минимальный вызов; требуемый контекст: В теле modTestGeometryReview.TestCollapsedBoundaries с его подготовленными локальными объектами/входом.

```vb
    Set copy = path.Slice(0#, 300#)
```

Полный контекст: tests/modTestGeometryReview.bas:221.

Native entrypoint: modTestGeometryReview.RunRoundedEndpointTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/UnifiedTestsCorrected/modTestGeometryReview.RunRoundedEndpointTests.txt.

<a id="api-cgeometrypath-reversed"></a>
### CGeometryPath.Reversed

Копирует путь с обратным порядком сегментов/обходом; signed sweep и source parameters обращаются.

```vb
Public Function Reversed() As CGeometryPath
```

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:315

Минимальный вызов; требуемый контекст: В теле modTestGeometryReview.TestCollapsedBoundaries с его подготовленными локальными объектами/входом.

```vb
    Set copy = path.Reversed
```

Полный контекст: tests/modTestGeometryReview.bas:216.

Native entrypoint: modTestGeometryReview.RunRoundedEndpointTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/UnifiedTestsCorrected/modTestGeometryReview.RunRoundedEndpointTests.txt.

<a id="api-cgeometrypath-transformed"></a>
### CGeometryPath.Transformed

Копирует путь изометрией: отражение по локальной X, поворот, перенос; длина сохраняется.

```vb
Public Function Transformed(ByVal shiftX As Double, ByVal shiftY As Double, Optional ByVal rotation As Double = 0#, Optional ByVal reflectX As Boolean = False) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| shiftX | Double; ByVal; обязательный | Перенос по X, мм. |
| shiftY | Double; ByVal; обязательный | Перенос по Y, мм. |
| rotation | Double; ByVal; Optional; default=0# | Угол поворота вокруг (0,0), радианы, CCW положителен. |
| reflectX | Boolean; ByVal; Optional; default=False | Отражение относительно локальной оси X перед поворотом и переносом. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:330

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.TestLine с его подготовленными локальными объектами/входом.

```vb
    Set changed = path.Transformed(10#, 20#, GEOM_PI / 2#)
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:628.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunPathTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/AnalyticPaths.txt.

<a id="api-cgeometrypath-rayhit"></a>
### CGeometryPath.RayHit

Находит ближайшее положительное пересечение луча с конечным путём; False не экстраполирует линию.

```vb
Public Function RayHit(ByVal originX As Double, ByVal originY As Double, ByVal directionX As Double, ByVal directionY As Double, ByRef pathDistance As Double, ByRef rayDistance As Double) As Boolean
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| originX | Double; ByVal; обязательный | Модельная координата originX, мм; знаки соответствуют section XY/WCS. |
| originY | Double; ByVal; обязательный | Модельная координата originY, мм; знаки соответствуют section XY/WCS. |
| directionX | Double; ByVal; обязательный | X-компонента направления проекции/выбора экстремума; вектор не нулевой, у RebarRowExtremeLabel единичный. |
| directionY | Double; ByVal; обязательный | Y-компонента направления проекции/выбора экстремума. |
| pathDistance | Double; ByRef; обязательный | Расстояние или граница интервала pathDistance, внутренние мм; направление и допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| rayDistance | Double; ByRef; обязательный | Расстояние или граница интервала rayDistance, внутренние мм; направление и допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |

Тип результата: Boolean. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя сегмента или рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:353

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.TestRays с его подготовленными локальными объектами/входом.

```vb
    Check "ray.circle.hit", path.RayHit(-20#, 0#, 2#, 0#, station, distance)
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:758.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunPathTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/AnalyticPaths.txt.

## CPolygonSpec

New→AddVertex в нужном порядке→Seal→чтение и повторное использование для исходного/смещённых путей. Минимум три вершины. После Seal изменение запрещено, повторный Seal безопасен. Топологию и касания проверяет прежнее геометрическое ядро; автоматических anchors нет.

<a id="api-cpolygonspec-addvertex"></a>
### CPolygonSpec.AddVertex

Добавляет запись исходной вершины: X/Y, радиус, ID и текстовый источник. Проверяет структуру и числа без геометрических вычислений.

```vb
Public Sub AddVertex(ByVal vertexID As String, ByVal x As Double, ByVal y As Double, Optional ByVal filletRadius As Double = 0#, Optional ByVal sourceReference As String = "")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| vertexID | String; ByVal; обязательный | Непустое уникальное имя исходной вершины; регистр различается, повтор первой вершины для замыкания не нужен. |
| x | Double; ByVal; обязательный | Модельная координата x, мм; знаки соответствуют section XY/WCS. |
| y | Double; ByVal; обязательный | Модельная координата y, мм; знаки соответствуют section XY/WCS. |
| filletRadius | Double; ByVal; Optional; default=0# | Неотрицательный конечный радиус сопряжения, мм. Ноль задаёт острый угол; фактические точки касания вычисляет общее ядро. |
| sourceReference | String; ByVal; Optional; default="" | Строка исходного адреса/программного параметра; без объекта Excel. |

Жизненный цикл: New→AddVertex в нужном порядке→Seal→чтение и повторное использование для исходного/смещённых путей. Минимум три вершины. После Seal изменение запрещено, повторный Seal безопасен. Топологию и касания проверяет прежнее геометрическое ядро; автоматических anchors нет.

Ошибки/исправление: 5470: пустой/повторный ID, отрицательный/не конечный радиус или координата, меньше трёх вершин, добавление после Seal, чтение до Seal/неверный индекс. SourceReference сохраняет источник ошибки; при новых размерах создайте новый набор.

Source: src/Geometry/CPolygonSpec.cls:35

Минимальный вызов; требуемый контекст: В теле modTestPolygonSpec.RunPolygonSpecTests с его подготовленными локальными объектами/входом.

```vb
    polygon.AddVertex "A", 0#, 0#
```

Полный контекст: tests/modTestPolygonSpec.bas:44.

Native entrypoint: modTestPolygonSpec.RunPolygonSpecTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/PolygonSpecTests/modTestPolygonSpec.RunPolygonSpecTests.txt.

<a id="api-cpolygonspec-seal"></a>
### CPolygonSpec.Seal

Завершает набор из минимум трёх вершин в исходном порядке. Повторный Seal безопасен; неудачный оставляет описание открытым.

```vb
Public Sub Seal()
```

Жизненный цикл: New→AddVertex в нужном порядке→Seal→чтение и повторное использование для исходного/смещённых путей. Минимум три вершины. После Seal изменение запрещено, повторный Seal безопасен. Топологию и касания проверяет прежнее геометрическое ядро; автоматических anchors нет.

Ошибки/исправление: 5470: пустой/повторный ID, отрицательный/не конечный радиус или координата, меньше трёх вершин, добавление после Seal, чтение до Seal/неверный индекс. SourceReference сохраняет источник ошибки; при новых размерах создайте новый набор.

Source: src/Geometry/CPolygonSpec.cls:60

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.BuildAudit02DemoDefinition с его подготовленными локальными объектами/входом.

```vb
    polygon.Seal
```

Полный контекст: tests/modTestAudit02Api.bas:263.

Native entrypoint: modTestAudit02Api.RunAudit02GuideExamplesTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/FinalApiAcceptance/modTestAudit02Api.RunAudit02GuideExamplesTests.txt.

<a id="api-cpolygonspec-count"></a>
### CPolygonSpec.Count

Число исходных вершин, доступное при заполнении и после Seal; не ограничено шестью.

```vb
Public Property Get Count() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→AddVertex в нужном порядке→Seal→чтение и повторное использование для исходного/смещённых путей. Минимум три вершины. После Seal изменение запрещено, повторный Seal безопасен. Топологию и касания проверяет прежнее геометрическое ядро; автоматических anchors нет.

Ошибки/исправление: 5470: пустой/повторный ID, отрицательный/не конечный радиус или координата, меньше трёх вершин, добавление после Seal, чтение до Seal/неверный индекс. SourceReference сохраняет источник ошибки; при новых размерах создайте новый набор.

Source: src/Geometry/CPolygonSpec.cls:68

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.SpecSnapshot с его подготовленными локальными объектами/входом.

```vb
    For i = 1 To polygon.Count
```

Полный контекст: tests/modTestAudit02Api.bas:546.

Native entrypoint: modTestAudit02Api.RunAudit02PolygonApiTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02PolygonApiTests.txt.

<a id="api-cpolygonspec-issealed"></a>
### CPolygonSpec.IsSealed

True после успешного явного Seal; только такой набор принимает фасад.

```vb
Public Property Get IsSealed() As Boolean
```

Тип результата: Boolean. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→AddVertex в нужном порядке→Seal→чтение и повторное использование для исходного/смещённых путей. Минимум три вершины. После Seal изменение запрещено, повторный Seal безопасен. Топологию и касания проверяет прежнее геометрическое ядро; автоматических anchors нет.

Ошибки/исправление: 5470: пустой/повторный ID, отрицательный/не конечный радиус или координата, меньше трёх вершин, добавление после Seal, чтение до Seal/неверный индекс. SourceReference сохраняет источник ошибки; при новых размерах создайте новый набор.

Source: src/Geometry/CPolygonSpec.cls:70

Минимальный вызов; требуемый контекст: В теле modTestPolygonSpec.RunPolygonSpecTests с его подготовленными локальными объектами/входом.

```vb
        Check "sealed." & CStr(n), polygon.IsSealed
```

Полный контекст: tests/modTestPolygonSpec.bas:28.

Native entrypoint: modTestPolygonSpec.RunPolygonSpecTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/PolygonSpecTests/modTestPolygonSpec.RunPolygonSpecTests.txt.

<a id="api-cpolygonspec-vertexid"></a>
### CPolygonSpec.VertexID

Имя исходной вершины по номеру 1..Count, доступное после Seal.

```vb
Public Function VertexID(ByVal index As Long) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→AddVertex в нужном порядке→Seal→чтение и повторное использование для исходного/смещённых путей. Минимум три вершины. После Seal изменение запрещено, повторный Seal безопасен. Топологию и касания проверяет прежнее геометрическое ядро; автоматических anchors нет.

Ошибки/исправление: 5470: пустой/повторный ID, отрицательный/не конечный радиус или координата, меньше трёх вершин, добавление после Seal, чтение до Seal/неверный индекс. SourceReference сохраняет источник ошибки; при новых размерах создайте новый набор.

Source: src/Geometry/CPolygonSpec.cls:73

Минимальный вызов; требуемый контекст: В теле modTestPolygonSpec.RunPolygonSpecTests с его подготовленными локальными объектами/входом.

```vb
        Check "kernel.copy.independent." & CStr(n), polygon.X(1) = 10# And polygon.FilletRadius(1) = 0# And polygon.VertexID(1) = "V1"
```

Полный контекст: tests/modTestPolygonSpec.bas:36.

Native entrypoint: modTestPolygonSpec.RunPolygonSpecTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/PolygonSpecTests/modTestPolygonSpec.RunPolygonSpecTests.txt.

<a id="api-cpolygonspec-x"></a>
### CPolygonSpec.X

Исходная координата X, мм, по номеру 1..Count; не является автоматически точкой готового контура.

```vb
Public Function X(ByVal index As Long) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→AddVertex в нужном порядке→Seal→чтение и повторное использование для исходного/смещённых путей. Минимум три вершины. После Seal изменение запрещено, повторный Seal безопасен. Топологию и касания проверяет прежнее геометрическое ядро; автоматических anchors нет.

Ошибки/исправление: 5470: пустой/повторный ID, отрицательный/не конечный радиус или координата, меньше трёх вершин, добавление после Seal, чтение до Seal/неверный индекс. SourceReference сохраняет источник ошибки; при новых размерах создайте новый набор.

Source: src/Geometry/CPolygonSpec.cls:77

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02AnnotationTests с его подготовленными локальными объектами/входом.

```vb
    Check "A03.virtual.not.touch", x <> polygon.X(1) Or y <> polygon.Y(1)
```

Полный контекст: tests/modTestAudit02Api.bas:178.

Native entrypoint: modTestAudit02Api.RunAudit02AnnotationTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02AnnotationTests.txt.

<a id="api-cpolygonspec-y"></a>
### CPolygonSpec.Y

Исходная координата Y, мм, по номеру 1..Count.

```vb
Public Function Y(ByVal index As Long) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→AddVertex в нужном порядке→Seal→чтение и повторное использование для исходного/смещённых путей. Минимум три вершины. После Seal изменение запрещено, повторный Seal безопасен. Топологию и касания проверяет прежнее геометрическое ядро; автоматических anchors нет.

Ошибки/исправление: 5470: пустой/повторный ID, отрицательный/не конечный радиус или координата, меньше трёх вершин, добавление после Seal, чтение до Seal/неверный индекс. SourceReference сохраняет источник ошибки; при новых размерах создайте новый набор.

Source: src/Geometry/CPolygonSpec.cls:81

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02AnnotationTests с его подготовленными локальными объектами/входом.

```vb
    Check "A03.virtual.not.touch", x <> polygon.X(1) Or y <> polygon.Y(1)
```

Полный контекст: tests/modTestAudit02Api.bas:178.

Native entrypoint: modTestAudit02Api.RunAudit02AnnotationTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02AnnotationTests.txt.

<a id="api-cpolygonspec-filletradius"></a>
### CPolygonSpec.FilletRadius

Назначенный исходный радиус, мм, по номеру 1..Count; ноль — острый угол.

```vb
Public Function FilletRadius(ByVal index As Long) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→AddVertex в нужном порядке→Seal→чтение и повторное использование для исходного/смещённых путей. Минимум три вершины. После Seal изменение запрещено, повторный Seal безопасен. Топологию и касания проверяет прежнее геометрическое ядро; автоматических anchors нет.

Ошибки/исправление: 5470: пустой/повторный ID, отрицательный/не конечный радиус или координата, меньше трёх вершин, добавление после Seal, чтение до Seal/неверный индекс. SourceReference сохраняет источник ошибки; при новых размерах создайте новый набор.

Source: src/Geometry/CPolygonSpec.cls:85

Минимальный вызов; требуемый контекст: В теле modTestPolygonSpec.RunPolygonSpecTests с его подготовленными локальными объектами/входом.

```vb
        Check "kernel.copy.independent." & CStr(n), polygon.X(1) = 10# And polygon.FilletRadius(1) = 0# And polygon.VertexID(1) = "V1"
```

Полный контекст: tests/modTestPolygonSpec.bas:36.

Native entrypoint: modTestPolygonSpec.RunPolygonSpecTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/PolygonSpecTests/modTestPolygonSpec.RunPolygonSpecTests.txt.

<a id="api-cpolygonspec-sourcereference"></a>
### CPolygonSpec.SourceReference

Скалярный текст происхождения вершины по номеру 1..Count; ссылки на Excel-объекты не хранятся.

```vb
Public Function SourceReference(ByVal index As Long) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New→AddVertex в нужном порядке→Seal→чтение и повторное использование для исходного/смещённых путей. Минимум три вершины. После Seal изменение запрещено, повторный Seal безопасен. Топологию и касания проверяет прежнее геометрическое ядро; автоматических anchors нет.

Ошибки/исправление: 5470: пустой/повторный ID, отрицательный/не конечный радиус или координата, меньше трёх вершин, добавление после Seal, чтение до Seal/неверный индекс. SourceReference сохраняет источник ошибки; при новых размерах создайте новый набор.

Source: src/Geometry/CPolygonSpec.cls:89

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.SpecSnapshot с его подготовленными локальными объектами/входом.

```vb
        SpecSnapshot = SpecSnapshot & polygon.VertexID(i) & "|" & CStr(polygon.X(i)) & "|" & CStr(polygon.Y(i)) & _
            "|" & CStr(polygon.FilletRadius(i)) & "|" & polygon.SourceReference(i) & vbCrLf
```

Полный контекст: tests/modTestAudit02Api.bas:547.

Native entrypoint: modTestAudit02Api.RunAudit02PolygonApiTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02PolygonApiTests.txt.

## CSectionDefinition

Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

<a id="api-csectiondefinition-lastdiagnostic"></a>
### CSectionDefinition.LastDiagnostic

Последняя машинная диагностика; Nothing при отсутствии записанного события.

```vb
Public Property Get LastDiagnostic() As CGeometryDiagnostic
```

Тип результата: CGeometryDiagnostic. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:25

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    Set diagnostic = definition.LastDiagnostic: Check "definition.successDiagnostic", diagnostic Is Nothing
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:117.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-csectiondefinition-anchorcoordinates"></a>
### CSectionDefinition.AnchorCoordinates

Возвращает координаты именованной привязки в скалярные переменные вызывающего кода.

```vb
Public Sub AnchorCoordinates(ByVal anchorID As String, ByRef x As Double, ByRef y As Double)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| anchorID | String; ByVal; обязательный | Непустой уникальный ID модельной привязки. |
| x | Double; ByRef; обязательный | Модельная координата x, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| y | Double; ByRef; обязательный | Модельная координата y, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:49

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02PolygonApiTests с его подготовленными локальными объектами/входом.

```vb
    definition.AnchorCoordinates "A", x, y
```

Полный контекст: tests/modTestAudit02Api.bas:60.

Native entrypoint: modTestAudit02Api.RunAudit02PolygonApiTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02PolygonApiTests.txt.

<a id="api-csectiondefinition-sourcereference"></a>
### CSectionDefinition.SourceReference

Читает скалярный SourceMap по ID; отсутствующая связь возвращает пустую строку.

```vb
Public Function SourceReference(ByVal entityID As String) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| entityID | String; ByVal; обязательный | Смысловой ID сущности для SourceMap. |

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:62

Минимальный вызов; требуемый контекст: В теле modTestRecipeAnnotations.RunRecipeAnnotationTests с его подготовленными локальными объектами/входом.

```vb
    Check "source.scalar", definition.SourceReference("Diagonal") = "ProgramInputs.PointB"
```

Полный контекст: tests/modTestRecipeAnnotations.bas:15.

Native entrypoint: modTestRecipeAnnotations.RunRecipeAnnotationTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/DeclarativeApi/FinalApiAnnotations.txt.

<a id="api-csectiondefinition-typeid"></a>
### CSectionDefinition.TypeID

Смысловое имя типа опубликованного определения.

```vb
Public Property Get TypeID() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:67

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    Check "definition.type", definition.TypeID = "ApiRead"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:113.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-csectiondefinition-schemaversion"></a>
### CSectionDefinition.SchemaVersion

Положительная версия схемы определения.

```vb
Public Property Get SchemaVersion() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:69

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    Check "definition.schema", definition.SchemaVersion = 2
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:114.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-csectiondefinition-issealed"></a>
### CSectionDefinition.IsSealed

Опубликован ли объект; published геометрия не изменяется.

```vb
Public Property Get IsSealed() As Boolean
```

Тип результата: Boolean. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:71

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.RunRecipeDefinitionTests с его подготовленными локальными объектами/входом.

```vb
    Check "definition.sealed", definition.IsSealed
```

Полный контекст: tests/modTestSectionRecipes.bas:580.

Native entrypoint: modTestSectionRecipes.RunRecipeDefinitionTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-csectiondefinition-pathcount"></a>
### CSectionDefinition.PathCount

Число зарегистрированных путей.

```vb
Public Property Get PathCount() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:73

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02PolygonApiTests с его подготовленными локальными объектами/входом.

```vb
    Check "P06.no.partial.path", definition.PathCount = 3
```

Полный контекст: tests/modTestAudit02Api.bas:55.

Native entrypoint: modTestAudit02Api.RunAudit02PolygonApiTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02PolygonApiTests.txt.

<a id="api-csectiondefinition-loopcount"></a>
### CSectionDefinition.LoopCount

Число подготовленных материальных loops.

```vb
Public Property Get LoopCount() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:75

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.RunRecipeDefinitionTests с его подготовленными локальными объектами/входом.

```vb
    Check "definition.loops", definition.LoopCount = 3 And region.LoopCount = 3
```

Полный контекст: tests/modTestSectionRecipes.bas:581.

Native entrypoint: modTestSectionRecipes.RunRecipeDefinitionTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-csectiondefinition-rulecount"></a>
### CSectionDefinition.RuleCount

Число объявленных правил, включая выключенные.

```vb
Public Property Get RuleCount() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:77

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.CircleInputCase с его подготовленными локальными объектами/входом.

```vb
    If scenario = 3 Then CircleInputCase = (definition.RuleCount = 0): Exit Function
```

Полный контекст: tests/modTestSectionRecipes.bas:666.

Native entrypoint: modTestSectionRecipes.RunCircleRecipeInputTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CircleEvidencePack/Actual/InputTests.txt.

<a id="api-csectiondefinition-annotationcount"></a>
### CSectionDefinition.AnnotationCount

Число объявленных annotation rules.

```vb
Public Property Get AnnotationCount() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:79

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    Check "definition.annotationCount", definition.AnnotationCount = 1
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:115.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-csectiondefinition-path"></a>
### CSectionDefinition.Path

Читает опубликованный неизменяемый путь по ID.

```vb
Public Function Path(ByVal pathID As String) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:82

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    Set path = definition.Path("Row")
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:137.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-csectiondefinition-targetpaths"></a>
### CSectionDefinition.TargetPaths

Возвращает ID одного пути или массив независимых путей группы.

```vb
Public Function TargetPaths(ByVal targetID As String) As Variant
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| targetID | String; ByVal; обязательный | ID пути или группы путей. |

Тип результата: Variant. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:88

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    ids = definition.TargetPaths("Rows"): Check "group.independent", UBound(ids) - LBound(ids) + 1 = 2
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:116.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-csectiondefinition-concretedomain"></a>
### CSectionDefinition.ConcreteDomain

Возвращает clone подготовленной бетонной области, без Excel или повторной интерпретации рецепта.

```vb
Public Function ConcreteDomain() As CConcreteRegion
```

Тип результата: CConcreteRegion. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:100

Минимальный вызов; требуемый контекст: В теле modTestRecipeSetbacks.RunRecipeSetbackInputTests с его подготовленными локальными объектами/входом.

```vb
    geometryArea = def.ConcreteDomain.Area
```

Полный контекст: tests/modTestRecipeSetbacks.bas:202.

Native entrypoint: modTestRecipeSetbacks.RunRecipeSetbackInputTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/DeclarativeApi/FinalGeometry.txt.

<a id="api-csectiondefinition-materialcontours"></a>
### CSectionDefinition.MaterialContours

Возвращает отдельный контейнер аналитических материальных Outer/Opening.

```vb
Public Function MaterialContours() As CSectionContours
```

Тип результата: CSectionContours. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:320

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.RunRecipeDefinitionTests с его подготовленными локальными объектами/входом.

```vb
    Set contours = definition.MaterialContours
```

Полный контекст: tests/modTestSectionRecipes.bas:586.

Native entrypoint: modTestSectionRecipes.RunRecipeDefinitionTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-csectiondefinition-writecontours"></a>
### CSectionDefinition.WriteContours

Записывает единственный контурный snapshot в переданный контейнер; destination предварительно очищается.

```vb
Public Sub WriteContours(ByVal result As CSectionContours)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| result | CSectionContours; ByVal; обязательный | Подготовленный объект/контейнер выбранного запроса, без повторного чтения Excel или назначения типа. |

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:329

Минимальный вызов; требуемый контекст: В теле modTestGeometry.TestRoundedRectangleContourAnnotations с его подготовленными локальными объектами/входом.

```vb
    geom.WriteContours model.Contours
```

Полный контекст: tests/modTestGeometry.bas:1785.

Native entrypoint: modTestGeometry.RunGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

## CRebarRuleOptions

New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

<a id="api-crebarruleoptions-startsetback"></a>
### CRebarRuleOptions.StartSetback

Начальное отсечение вдоль каждого target-пути, мм≥0; у полного ClosedUniform 0.

```vb
Public StartSetback As Double
```

Default после New: 0#. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:14

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.RunRecipeRuleTests с его подготовленными локальными объектами/входом.

```vb
    options.StartSetback = 10#: options.EndSetback = 20#
```

Полный контекст: tests/modTestSectionRecipes.bas:184.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-crebarruleoptions-endsetback"></a>
### CRebarRuleOptions.EndSetback

Конечное отсечение вдоль каждого target-пути, мм≥0; ProjectFrom фильтрует hit без перераспределения.

```vb
Public EndSetback As Double
```

Default после New: 0#. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:15

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.RunRecipeRuleTests с его подготовленными локальными объектами/входом.

```vb
    options.StartSetback = 10#: options.EndSetback = 20#
```

Полный контекст: tests/modTestSectionRecipes.bas:184.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-crebarruleoptions-endmode"></a>
### CRebarRuleOptions.EndMode

IncludeBoth/ExcludeBoth/IncludeStart/IncludeEnd/ClosedUniform; последние требуют соответствующей замкнутости.

```vb
Public EndMode As ERebarEndMode
```

Default после New: IncludeBoth. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: ERebarEndMode. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:16

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.RunRecipeRuleTests с его подготовленными локальными объектами/входом.

```vb
    options.EndMode = ExcludeBoth
```

Полный контекст: tests/modTestSectionRecipes.bas:204.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-crebarruleoptions-distributionscope"></a>
### CRebarRuleOptions.DistributionScope

Допустимо только PerPath: число стержней, расстояния между ними и отступы от концов применяются отдельно к каждой независимой линии группы.

```vb
Public DistributionScope As ERebarDistributionScope
```

Default после New: PerPath. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: ERebarDistributionScope. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:17

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.AdvancedGuideDefinition с его подготовленными локальными объектами/входом.

```vb
    options.DistributionScope = PerPath: options.StartSetback = 20#: options.EndSetback = 30#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:61.

Native entrypoint: modTestRecipeApiCatalogue.RunAdvancedGuideExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/AdvancedExamples.txt.

<a id="api-crebarruleoptions-remainderplacement"></a>
### CRebarRuleOptions.RemainderPlacement

RemainderStart/End/Center: положение остатка для точного шага.

```vb
Public RemainderPlacement As ERemainderPlacement
```

Default после New: RemainderCenter. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: ERemainderPlacement. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:18

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    Check "defaults.remainder", options.RemainderPlacement = RemainderCenter
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:105.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-crebarruleoptions-tangentoffset"></a>
### CRebarRuleOptions.TangentOffset

Локальное связанное смещение вдоль касательной, мм; signed, не независимый шаг ряда.

```vb
Public TangentOffset As Double
```

Default после New: 0#. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:19

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    Check "defaults.tangent", options.TangentOffset = 0#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:106.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-crebarruleoptions-normaloffset"></a>
### CRebarRuleOptions.NormalOffset

Локальное связанное смещение по левой нормали (-ty,tx), мм; signed.

```vb
Public NormalOffset As Double
```

Default после New: 0#. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:20

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.BuildAudit02OperationsDefinition с его подготовленными локальными объектами/входом.

```vb
    options.NormalOffset = 40#: options.TangentOffset = 10#
```

Полный контекст: tests/modTestAudit02Api.bas:308.

Native entrypoint: modTestAudit02Api.RunAudit02GuideExamplesTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/FinalApiAcceptance/modTestAudit02Api.RunAudit02GuideExamplesTests.txt.

<a id="api-crebarruleoptions-everysecond"></a>
### CRebarRuleOptions.EverySecond

Брать локальные исходные позиции 1, 3, ... отдельно на каждом пути; SuccessfulProjectionOrdinal задаёт отдельный отсчёт успешных проекций.

```vb
Public EverySecond As Boolean
```

Default после New: False. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: Boolean. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:21

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02CloseoutContactTests с его подготовленными локальными объектами/входом.

```vb
    options.NormalOffset = 12#: options.TangentOffset = 34#: options.EverySecond = True
```

Полный контекст: tests/modTestAudit02Api.bas:349.

Native entrypoint: modTestAudit02Api.RunAudit02CloseoutContactTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02Closeout_2026-10-14/BaselineDirectedTests/modTestAudit02Api.RunAudit02CloseoutContactTests.txt.

<a id="api-crebarruleoptions-incomingtangentatjoint"></a>
### CRebarRuleOptions.IncomingTangentAtJoint

В точке соединения двух сегментов брать касательную предыдущего сегмента. По умолчанию берётся касательная следующего сегмента.

```vb
Public IncomingTangentAtJoint As Boolean
```

Default после New: False. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: Boolean. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:22

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.BuildAudit02DemoDefinition с его подготовленными локальными объектами/входом.

```vb
    options.IncomingTangentAtJoint = True
```

Полный контекст: tests/modTestAudit02Api.bas:273.

Native entrypoint: modTestAudit02Api.RunAudit02GuideExamplesTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/FinalApiAcceptance/modTestAudit02Api.RunAudit02GuideExamplesTests.txt.

<a id="api-crebarruleoptions-directionx"></a>
### CRebarRuleOptions.DirectionX

Глобальная X направления ProjectFrom, используется при SourceNormalSign=0.

```vb
Public DirectionX As Double
```

Default после New: 0#. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:23

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    Check "defaults.projectionDirection", options.DirectionX = 0# And options.DirectionY = 0#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:107.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-crebarruleoptions-directiony"></a>
### CRebarRuleOptions.DirectionY

Глобальная Y направления ProjectFrom; направление не нулевое.

```vb
Public DirectionY As Double
```

Default после New: 0#. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:24

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.InvalidRuleRejected с его подготовленными локальными объектами/входом.

```vb
            options.DirectionY = -1#
```

Полный контекст: tests/modTestSectionRecipes.bas:550.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-crebarruleoptions-sourcenormalsign"></a>
### CRebarRuleOptions.SourceNormalSign

-1/0/+1: локальная нормаль исходного ряда или глобальное DirectionX/Y при 0.

```vb
Public SourceNormalSign As Double
```

Default после New: 0#. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:25

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.TestProjectionNormalsAndOrdinals с его подготовленными локальными объектами/входом.

```vb
    options.SourceNormalSign = 1#
```

Полный контекст: tests/modTestSectionRecipes.bas:285.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-crebarruleoptions-successfulprojectionordinal"></a>
### CRebarRuleOptions.SuccessfulProjectionOrdinal

True: заново нумеровать только стержни, для которых найдено пересечение с целевой линией, подряд по всей группе. False: сохранять исходные номера стержней с возможными пропусками отдельно на каждой линии.

```vb
Public SuccessfulProjectionOrdinal As Boolean
```

Default после New: False. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: Boolean. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:26

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02OptionsTests с его подготовленными локальными объектами/входом.

```vb
    options.SourceNormalSign = -1#: options.SuccessfulProjectionOrdinal = True
```

Полный контекст: tests/modTestAudit02Api.bas:93.

Native entrypoint: modTestAudit02Api.RunAudit02OptionsTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02OptionsTests.txt.

<a id="api-crebarruleoptions-newpublicationroot"></a>
### CRebarRuleOptions.NewPublicationRoot

У ProjectFrom начинает новый корень публикации по успешным позициям.

```vb
Public NewPublicationRoot As Boolean
```

Default после New: False. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: Boolean. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:27

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02OptionsTests с его подготовленными локальными объектами/входом.

```vb
    options.SkipMissingProjection = True: options.NewPublicationRoot = True
```

Полный контекст: tests/modTestAudit02Api.bas:94.

Native entrypoint: modTestAudit02Api.RunAudit02OptionsTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02OptionsTests.txt.

<a id="api-crebarruleoptions-skipmissingprojection"></a>
### CRebarRuleOptions.SkipMissingProjection

Разрешить пропуск без hit с PROJECTION_MISSING; False выдаёт ошибку.

```vb
Public SkipMissingProjection As Boolean
```

Default после New: False. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: Boolean. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:28

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.AdvancedGuideDefinition с его подготовленными локальными объектами/входом.

```vb
    options.DirectionY = 1#: options.SkipMissingProjection = True
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:78.

Native entrypoint: modTestRecipeApiCatalogue.RunAdvancedGuideExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/AdvancedExamples.txt.

<a id="api-crebarruleoptions-publication"></a>
### CRebarRuleOptions.Publication

PublishByRule или PublishPerSourcePosition; влияет на вывод, не геометрический шаг.

```vb
Public Publication As ERebarPublication
```

Default после New: PublishByRule. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: ERebarPublication. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:29

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.TestProjectionNormalsAndOrdinals с его подготовленными локальными объектами/входом.

```vb
        options.Publication = PublishPerSourcePosition
```

Полный контекст: tests/modTestSectionRecipes.bas:302.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-crebarruleoptions-positionorder"></a>
### CRebarRuleOptions.PositionOrder

Порядок вывода стержней: AlongPath — от начала линии, FromPathMiddle — от её середины. Места стержней и их исходные номера сохраняются.

```vb
Public PositionOrder As ERebarPositionOrder
```

Default после New: AlongPath. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: ERebarPositionOrder. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:30

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02OptionsTests с его подготовленными локальными объектами/входом.

```vb
    options.PositionOrder = FromPathMiddle: options.IncomingTangentAtJoint = True
```

Полный контекст: tests/modTestAudit02Api.bas:92.

Native entrypoint: modTestAudit02Api.RunAudit02OptionsTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02OptionsTests.txt.

<a id="api-crebarruleoptions-coincidence"></a>
### CRebarRuleOptions.Coincidence

RejectCoincidence/MergeIdentical/KeepPublishedOwner; последнее сохраняет явно выбранного опубликованного владельца.

```vb
Public Coincidence As ERebarCoincidence
```

Default после New: RejectCoincidence. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: ERebarCoincidence. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:31

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.TestPublicationAndCoincidence с его подготовленными локальными объектами/входом.

```vb
    options.AnnotationGroup = "Second": options.Coincidence = MergeIdentical
```

Полный контекст: tests/modTestSectionRecipes.bas:394.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-crebarruleoptions-placementcheck"></a>
### CRebarRuleOptions.PlacementCheck

CheckCenter или CheckFullDisk относительно Outer/Opening; ужесточение задаётся явно.

```vb
Public PlacementCheck As ERebarPlacementCheck
```

Default после New: CheckCenter. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: ERebarPlacementCheck. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:32

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.AdvancedGuideDefinition с его подготовленными локальными объектами/входом.

```vb
    options.PlacementCheck = CheckFullDisk: options.OverlapPolicy = RejectOverlap
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:63.

Native entrypoint: modTestRecipeApiCatalogue.RunAdvancedGuideExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/AdvancedExamples.txt.

<a id="api-crebarruleoptions-overlappolicy"></a>
### CRebarRuleOptions.OverlapPolicy

ReportOverlap или RejectOverlap. Контакт допустим; нормативный минимальный зазор не вводится.

```vb
Public OverlapPolicy As ERebarOverlapPolicy
```

Default после New: ReportOverlap. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: ERebarOverlapPolicy. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:33

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
        If mode = 2 Then options.OverlapPolicy = RejectOverlap
```

Полный контекст: tests/modTestRecipeOperations.bas:102.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-crebarruleoptions-sourceprefix"></a>
### CRebarRuleOptions.SourcePrefix

Префикс SourcePositionId; пустой использует RuleId+'.' либо связь source+suffix.

```vb
Public SourcePrefix As String
```

Default после New: vbNullString. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:34

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.RunRecipeRuleTests с его подготовленными локальными объектами/входом.

```vb
    options.SourcePrefix = "P"
```

Полный контекст: tests/modTestSectionRecipes.bas:215.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-crebarruleoptions-sourcesuffix"></a>
### CRebarRuleOptions.SourceSuffix

Суффикс SourcePositionId, в том числе для связанных рядов.

```vb
Public SourceSuffix As String
```

Default после New: vbNullString. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:35

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.RunRecipeRuleTests с его подготовленными локальными объектами/входом.

```vb
    options.DirectionY = 1#: options.StartSetback = 110#: options.SourceSuffix = "_Hit"
```

Полный контекст: tests/modTestSectionRecipes.bas:245.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-crebarruleoptions-annotationgroup"></a>
### CRebarRuleOptions.AnnotationGroup

Смысловая принадлежность подписи физического стержня; не расчётная группа трещин СП35.

```vb
Public AnnotationGroup As String
```

Default после New: vbNullString. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:36

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.AdvancedGuideDefinition с его подготовленными локальными объектами/входом.

```vb
    options.AnnotationGroup = "RowsLabel"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:62.

Native entrypoint: modTestRecipeApiCatalogue.RunAdvancedGuideExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/AdvancedExamples.txt.

<a id="api-crebarruleoptions-comment"></a>
### CRebarRuleOptions.Comment

Инженерное пояснение правила; не selector или ResultCode.

```vb
Public Comment As String
```

Default после New: vbNullString. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:37

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    options.StartSetback = 20#: options.EndSetback = 30#: options.Comment = "Guide": options.SourceReference = "ProgramInputs.Row"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:109.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-crebarruleoptions-sourcereference"></a>
### CRebarRuleOptions.SourceReference

Реальные адреса исходных параметров правила или программная ссылка.

```vb
Public SourceReference As String
```

Default после New: vbNullString. Поле изменяется до объявления правила; facade копирует его значение.

Тип записываемого поля: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:38

Минимальный вызов; требуемый контекст: В теле modTestGeometryConfig.CheckRuleCounterBinding с его подготовленными локальными объектами/входом.

```vb
    Set options = New CRebarRuleOptions: options.SourceReference = "CounterBinding.Main"
```

Полный контекст: tests/modTestGeometryConfig.bas:72.

Native entrypoint: modTestGeometryConfig.RunAudit03RebarCounterTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/ProductionRecipeEvidencePack/data/ProductionRecipe/RebarCounters.txt.

<a id="api-crebarruleoptions-setpublication"></a>
### CRebarRuleOptions.SetPublication

Задаёт строго AnnotationGroup/Publication/Coincidence. Концы, offsets, проверка размещения, source и настройки проекции не меняются.

```vb
Public Sub SetPublication(ByVal annotationGroup As String, ByVal publication As ERebarPublication, ByVal coincidence As ERebarCoincidence)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| annotationGroup | String; ByVal; обязательный | Имя группы фактических стержней для общей подписи, независимо от расчётных групп трещин. |
| publication | ERebarPublication; ByVal; обязательный | PublishByRule или PublishPerSourcePosition; порядок и происхождение опубликованных стержней. |
| coincidence | ERebarCoincidence; ByVal; обязательный | RejectCoincidence, MergeIdentical или KeepPublishedOwner; ничего не меняет в концевых/проекционных настройках. |

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:55

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02OptionsTests с его подготовленными локальными объектами/входом.

```vb
    options.SetPublication "Group", PublishPerSourcePosition, KeepPublishedOwner
```

Полный контекст: tests/modTestAudit02Api.bas:95.

Native entrypoint: modTestAudit02Api.RunAudit02OptionsTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02OptionsTests.txt.

<a id="api-crebarruleoptions-configuredependentcontact"></a>
### CRebarRuleOptions.ConfigureDependentContact

Задаёт только NormalOffset/TangentOffset/EverySecond по сумме половин основного и добавляемого диаметров и промежуточных полных диаметров. Оба диаметра строго положительны: ноль вызывает ошибку. Отключённый ряд обрабатывает вызывающий рецепт до вызова. Вторая компонента обнуляется; остальные политики и source не наследуются.

```vb
Public Sub ConfigureDependentContact(ByVal baseDiameter As Double, ByVal diameter As Double, ByVal axis As ERebarOffsetAxis, ByVal directionSign As Double, ByVal everySecond As Boolean, Optional ByVal interveningDiameterSum As Double = 0#)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| baseDiameter | Double; ByVal; обязательный | Строго положительный диаметр основного стержня, мм (>0). Ноль и отрицательное значение вызывают ошибку; отключение ряда выполняет рецепт до вызова помощника. |
| diameter | Double; ByVal; обязательный | Строго положительный диаметр добавляемого стержня, мм (>0). Ноль и отрицательное значение вызывают ошибку; отключение ряда выполняет рецепт до вызова помощника. |
| axis | ERebarOffsetAxis; ByVal; обязательный | RowNormal — смещение по левой нормали исходного ряда; RowTangent — вдоль его касательной. Ориентацию задают фактический путь и directionSign. |
| directionSign | Double; ByVal; обязательный | Только -1 или +1: знак выбранной компоненты смещения; задаётся явно по направлению ряда. |
| everySecond | Boolean; ByVal; обязательный | True выбирает локальные исходные позиции 1,3,5…; отсчёт успешных проекций остаётся отдельной настройкой. |
| interveningDiameterSum | Double; ByVal; Optional; default=0# | Неотрицательная сумма диаметров номинально промежуточных стержней, мм. Для третьего ряда обычно d2 при том же направлении, иначе 0; пропуск EverySecond сам по себе её не уменьшает. |

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:66

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02OptionsTests с его подготовленными локальными объектами/входом.

```vb
    options.ConfigureDependentContact 20#, 12#, RowNormal, 0#, False
```

Полный контекст: tests/modTestAudit02Api.bas:114.

Native entrypoint: modTestAudit02Api.RunAudit02OptionsTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02OptionsTests.txt.

## CGeometryDiagnostic

Создаётся сервисом как immutable snapshot. Читайте через LastDiagnostic/DiagnosticAt; диагностический объект не изменяет геометрию или статусы.

<a id="api-cgeometrydiagnostic-code"></a>
### CGeometryDiagnostic.Code

Машинный код причины; назначается напрямую, не извлекается из русского текста.

```vb
Public Property Get Code() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Создаётся сервисом как immutable snapshot. Читайте через LastDiagnostic/DiagnosticAt; диагностический объект не изменяет геометрию или статусы.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Geometry/CGeometryDiagnostic.cls:26

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
            Check "overlap.reject", i = vbObjectError + 5530 And diagnostic.Code = "BAR_OVERLAP"
```

Полный контекст: tests/modTestRecipeOperations.bas:112.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-cgeometrydiagnostic-entityid"></a>
### CGeometryDiagnostic.EntityID

Смысловой ID сущности отказа, если он известен этому сервису.

```vb
Public Property Get EntityID() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Создаётся сервисом как immutable snapshot. Читайте через LastDiagnostic/DiagnosticAt; диагностический объект не изменяет геометрию или статусы.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Geometry/CGeometryDiagnostic.cls:28

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
            Check "overlap.source", diagnostic.EntityID = "Bars" And diagnostic.SourceReference = "ProgramInputs.Spacing"
```

Полный контекст: tests/modTestRecipeOperations.bas:113.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-cgeometrydiagnostic-relatedentityid"></a>
### CGeometryDiagnostic.RelatedEntityID

ID связанного сегмента/позиции/стержня для локализации причины.

```vb
Public Property Get RelatedEntityID() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Создаётся сервисом как immutable snapshot. Читайте через LastDiagnostic/DiagnosticAt; диагностический объект не изменяет геометрию или статусы.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Geometry/CGeometryDiagnostic.cls:30

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.InvalidOperation с его подготовленными локальными объектами/входом.

```vb
    If InvalidOperation Then InvalidOperation = diagnostic.Code = expected And diagnostic.EntityID = "Failed" And Len(diagnostic.RelatedEntityID) > 0
```

Полный контекст: tests/modTestRecipeOperations.bas:90.

Native entrypoint: modTestRecipeOperations.RunRecipeOperationTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-cgeometrydiagnostic-reason"></a>
### CGeometryDiagnostic.Reason

Человекочитаемая причина; не selector/status.

```vb
Public Property Get Reason() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Создаётся сервисом как immutable snapshot. Читайте через LastDiagnostic/DiagnosticAt; диагностический объект не изменяет геометрию или статусы.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Geometry/CGeometryDiagnostic.cls:32

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    Check "facade.offsetReason", Len(diagnostic.Reason) > 0
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:157.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-cgeometrydiagnostic-detail1"></a>
### CGeometryDiagnostic.Detail1

Первое численное поле конкретного Code; единицы определяет Code.

```vb
Public Property Get Detail1() As Double
```

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Создаётся сервисом как immutable snapshot. Читайте через LastDiagnostic/DiagnosticAt; диагностический объект не изменяет геометрию или статусы.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Geometry/CGeometryDiagnostic.cls:34

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
            Near "overlap.numeric.gap", diagnostic.Detail1, -10#
```

Полный контекст: tests/modTestRecipeOperations.bas:114.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-cgeometrydiagnostic-detail2"></a>
### CGeometryDiagnostic.Detail2

Второе численное поле конкретного Code.

```vb
Public Property Get Detail2() As Double
```

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Создаётся сервисом как immutable snapshot. Читайте через LastDiagnostic/DiagnosticAt; диагностический объект не изменяет геометрию или статусы.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Geometry/CGeometryDiagnostic.cls:36

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.InvalidOperation с его подготовленными локальными объектами/входом.

```vb
    If scenario = 2 And InvalidOperation Then InvalidOperation = diagnostic.Detail1 = 10# And diagnostic.Detail2 = 20#
```

Полный контекст: tests/modTestRecipeOperations.bas:91.

Native entrypoint: modTestRecipeOperations.RunRecipeOperationTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-cgeometrydiagnostic-sourcereference"></a>
### CGeometryDiagnostic.SourceReference

Реальный адрес/программный источник, добавленный SourceMap или правилом.

```vb
Public Property Get SourceReference() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Создаётся сервисом как immutable snapshot. Читайте через LastDiagnostic/DiagnosticAt; диагностический объект не изменяет геометрию или статусы.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Geometry/CGeometryDiagnostic.cls:38

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
            Check "overlap.source", diagnostic.EntityID = "Bars" And diagnostic.SourceReference = "ProgramInputs.Spacing"
```

Полный контекст: tests/modTestRecipeOperations.bas:113.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

## CSectionInputTable

Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

<a id="api-csectioninputtable-initialize"></a>
### CSectionInputTable.Initialize

Копирует двумерные scalar values и согласованные addresses один раз; Excel-объекты не сохраняет.

```vb
Public Sub Initialize(ByVal tableName As String, ByVal values As Variant, Optional ByVal addresses As Variant)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| tableName | String; ByVal; обязательный | Имя единственного корневого входного блока этого снимка. |
| values | Variant; ByVal; обязательный | Двумерный массив скалярных значений, копируется; вложенные массивы/объекты запрещены. |
| addresses | Variant; ByVal; Optional; пропущенный Variant=Missing | Необязательный согласованный 2D-массив адресов; без него применяется программный tableName[row,col]. |

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:20

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.InvalidInputRejected с его подготовленными локальными объектами/входом.

```vb
        Case 10: table.Initialize "Second", cells
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:531.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunInputTableTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/InputTables.txt.

<a id="api-csectioninputtable-name"></a>
### CSectionInputTable.Name

Имя корневого блока snapshot.

```vb
Public Property Get Name() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:57

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiInputAndDataExamples с его подготовленными локальными объектами/входом.

```vb
    Check "table.name", table.Name = "ProgramInputs"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:221.

Native entrypoint: modTestRecipeApiCatalogue.RunApiInputAndDataExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-csectioninputtable-rowcount"></a>
### CSectionInputTable.RowCount

Число строк неизменяемого input snapshot.

```vb
Public Property Get RowCount() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:59

Минимальный вызов; требуемый контекст: В теле modTestSectionSolver.TestRecipeInputCatalog с его подготовленными локальными объектами/входом.

```vb
            For row = 1 To table.RowCount
```

Полный контекст: tests/modTestSectionSolver.bas:504.

Native entrypoint: modTestSectionSolver.RunSectionSolverTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/NumericalRegression.txt.

<a id="api-csectioninputtable-columncount"></a>
### CSectionInputTable.ColumnCount

Число столбцов input snapshot.

```vb
Public Property Get ColumnCount() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:61

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.RunInputTableTests с его подготовленными локальными объектами/входом.

```vb
    Check "input.columns", table.ColumnCount = 3
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:469.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunInputTableTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/InputTables.txt.

<a id="api-csectioninputtable-sourcereference"></a>
### CSectionInputTable.SourceReference

Возвращает исходный адрес ячейки или программный tableName[row,col].

```vb
Public Function SourceReference(ByVal row As Long, ByVal column As Long) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| row | Long; ByVal; обязательный | Номер строки снимка 1..RowCount. |
| column | Long; ByVal; обязательный | Номер столбца снимка 1..ColumnCount. |

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:64

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.CloseoutRectSetInputMatrix с его подготовленными локальными объектами/входом.

```vb
                expectedAddress = table.SourceReference(row, column)
```

Полный контекст: tests/modTestAudit02Api.bas:430.

Native entrypoint: modTestAudit02Api.RunAudit02CloseoutRectSetInputTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02Closeout_2026-10-14/BaselineDirectedTests/modTestAudit02Api.RunAudit02CloseoutRectSetInputTests.txt.

<a id="api-csectioninputtable-cellvalue"></a>
### CSectionInputTable.CellValue

Читает scalar snapshot; ошибка Excel в ячейке даёт адресное исключение.

```vb
Public Function CellValue(ByVal row As Long, ByVal column As Long) As Variant
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| row | Long; ByVal; обязательный | Номер строки снимка 1..RowCount. |
| column | Long; ByVal; обязательный | Номер столбца снимка 1..ColumnCount. |

Тип результата: Variant. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:69

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.RunWorkbookRecipeInputTests с его подготовленными локальными объектами/входом.

```vb
        Set table = context.Table(rootName): initialCell = table.CellValue(1, 1)
```

Полный контекст: tests/modTestSectionRecipes.bas:38.

Native entrypoint: modTestSectionRecipes.RunWorkbookRecipeInputTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/ProductionRecipeEvidencePack/data/ProductionRecipe/WorkbookInput.txt.

<a id="api-csectioninputtable-isblank"></a>
### CSectionInputTable.IsBlank

Проверяет пустое значение снимка; ошибка Excel не подменяется пустотой.

```vb
Public Function IsBlank(ByVal row As Long, ByVal column As Long) As Boolean
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| row | Long; ByVal; обязательный | Номер строки снимка 1..RowCount. |
| column | Long; ByVal; обязательный | Номер столбца снимка 1..ColumnCount. |

Тип результата: Boolean. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:76

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.RunInputTableTests с его подготовленными локальными объектами/входом.

```vb
    Check "input.blank", table.IsBlank(4, 3)
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:479.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunInputTableTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/InputTables.txt.

<a id="api-csectioninputtable-text"></a>
### CSectionInputTable.Text

Читает непустой текст либо явно переданный fallback, обрезая крайние пробелы.

```vb
Public Function Text(ByVal row As Long, ByVal column As Long, Optional ByVal defaultValue As Variant) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| row | Long; ByVal; обязательный | Номер строки снимка 1..RowCount. |
| column | Long; ByVal; обязательный | Номер столбца снимка 1..ColumnCount. |
| defaultValue | Variant; ByVal; Optional; пропущенный Variant=Missing | Явный fallback конкретного вызова. Пропущенный Optional Variant без default означает обязательное поле. |

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:81

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiInputAndDataExamples с его подготовленными локальными объектами/входом.

```vb
    Check "table.text", table.Text(3, 2) = "Guide"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:222.

Native entrypoint: modTestRecipeApiCatalogue.RunApiInputAndDataExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-csectioninputtable-number"></a>
### CSectionInputTable.Number

Читает число, исключая Boolean/ошибки/объекты; обязательное пустое поле не заменяется нулём.

```vb
Public Function Number(ByVal row As Long, ByVal column As Long, Optional ByVal defaultValue As Variant) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| row | Long; ByVal; обязательный | Номер строки снимка 1..RowCount. |
| column | Long; ByVal; обязательный | Номер столбца снимка 1..ColumnCount. |
| defaultValue | Variant; ByVal; Optional; пропущенный Variant=Missing | Явный fallback конкретного вызова. Пропущенный Optional Variant без default означает обязательное поле. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:93

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.InvalidInputRejected с его подготовленными локальными объектами/входом.

```vb
        Case 11: value = table.Number(3, 2)
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:532.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunInputTableTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/InputTables.txt.

<a id="api-csectioninputtable-count"></a>
### CSectionInputTable.Count

Читает целое неотрицательное 0..2147483647; отдельный executor проверяет суммарное число позиций.

```vb
Public Function Count(ByVal row As Long, ByVal column As Long, Optional ByVal defaultValue As Long = 0) As Long
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| row | Long; ByVal; обязательный | Номер строки снимка 1..RowCount. |
| column | Long; ByVal; обязательный | Номер столбца снимка 1..ColumnCount. |
| defaultValue | Long; ByVal; Optional; default=0 | Явный fallback конкретного вызова. Пропущенный Optional Variant без default означает обязательное поле. |

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:111

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.RunInputTableTests с его подготовленными локальными объектами/входом.

```vb
    Check "input.count", table.Count(2, 2) = 4
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:475.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunInputTableTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/InputTables.txt.

<a id="api-csectioninputtable-choice"></a>
### CSectionInputTable.Choice

Проверяет вариант по allowed без учёта регистра и возвращает каноническое значение allowed.

```vb
Public Function Choice(ByVal row As Long, ByVal column As Long, ByVal allowed As Variant, Optional ByVal defaultValue As String = vbNullString) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| row | Long; ByVal; обязательный | Номер строки снимка 1..RowCount. |
| column | Long; ByVal; обязательный | Номер столбца снимка 1..ColumnCount. |
| allowed | Variant; ByVal; обязательный | Массив разрешённых строк; Choice возвращает каноническую строку из этого массива. |
| defaultValue | String; ByVal; Optional; default=vbNullString | Явный fallback конкретного вызова. Пропущенный Optional Variant без default означает обязательное поле. |

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:120

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.InvalidInputRejected с его подготовленными локальными объектами/входом.

```vb
        Case 9: value = table.Choice(2, 2, Array("Stacked", "SideBySide"))
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:530.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunInputTableTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/InputTables.txt.

<a id="api-csectioninputtable-length"></a>
### CSectionInputTable.Length

Читает обязательное число и один раз переводит input length в мм через CUnitSystem.

```vb
Public Function Length(ByVal row As Long, ByVal column As Long, ByVal units As CUnitSystem) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| row | Long; ByVal; обязательный | Номер строки снимка 1..RowCount. |
| column | Long; ByVal; обязательный | Номер столбца снимка 1..ColumnCount. |
| units | CUnitSystem; ByVal; обязательный | CUnitSystem, уже подготовленная вызывающим workflow; преобразование только через её методы. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:131

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.RunInputTableTests с его подготовленными локальными объектами/входом.

```vb
    Near "input.length", table.Length(4, 2, units), 16#
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:484.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunInputTableTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/InputTables.txt.

<a id="api-csectioninputtable-positivelength"></a>
### CSectionInputTable.PositiveLength

Читает обязательную длину существующим Length и один раз переводит её в мм; проверяет >0. Пустота/текст/ошибка не становятся нулём; отказ содержит фактический адрес. Не заменяет ограничения GEOM_TOLERANCE/as>d/2.

```vb
Public Function PositiveLength(ByVal row As Long, ByVal column As Long, ByVal units As CUnitSystem) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| row | Long; ByVal; обязательный | Номер строки снимка 1..RowCount. |
| column | Long; ByVal; обязательный | Номер столбца снимка 1..ColumnCount. |
| units | CUnitSystem; ByVal; обязательный | CUnitSystem, уже подготовленная вызывающим workflow; преобразование только через её методы. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:137

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02InputTests с его подготовленными локальными объектами/входом.

```vb
        number = table.PositiveLength(i, 1, units)
```

Полный контекст: tests/modTestAudit02Api.bas:139.

Native entrypoint: modTestAudit02Api.RunAudit02InputTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02InputTests.txt.

<a id="api-csectioninputtable-nonnegativelength"></a>
### CSectionInputTable.NonNegativeLength

Читает обязательную длину существующим Length и один раз переводит её в мм; проверяет >=0. Пустота/текст/ошибка не становятся нулём; отказ содержит фактический адрес. Не заменяет ограничения GEOM_TOLERANCE/as>d/2.

```vb
Public Function NonNegativeLength(ByVal row As Long, ByVal column As Long, ByVal units As CUnitSystem) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| row | Long; ByVal; обязательный | Номер строки снимка 1..RowCount. |
| column | Long; ByVal; обязательный | Номер столбца снимка 1..ColumnCount. |
| units | CUnitSystem; ByVal; обязательный | CUnitSystem, уже подготовленная вызывающим workflow; преобразование только через её методы. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:143

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02InputTests с его подготовленными локальными объектами/входом.

```vb
    Check "input.nonnegative.zero", table.NonNegativeLength(2, 1, units) = 0#
```

Полный контекст: tests/modTestAudit02Api.bas:136.

Native entrypoint: modTestAudit02Api.RunAudit02InputTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02InputTests.txt.

<a id="api-csectioninputtable-optionaldiameter"></a>
### CSectionInputTable.OptionalDiameter

Пустая ячейка или 0 отключают ряд; положительный диаметр переводится в мм. Текст, включая одиночный дефис и TODO, и отрицательный диаметр вызывают адресную ошибку ввода.

```vb
Public Function OptionalDiameter(ByVal row As Long, ByVal column As Long, ByVal units As CUnitSystem) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| row | Long; ByVal; обязательный | Номер строки снимка 1..RowCount. |
| column | Long; ByVal; обязательный | Номер столбца снимка 1..ColumnCount. |
| units | CUnitSystem; ByVal; обязательный | CUnitSystem, уже подготовленная вызывающим workflow; преобразование только через её методы. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:150

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.RunInputTableTests с его подготовленными локальными объектами/входом.

```vb
    Near "input.optional.zero", table.OptionalDiameter(3, 2, units), 0#
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:481.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunInputTableTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/InputTables.txt.

<a id="api-csectioninputtable-columnbycaption"></a>
### CSectionInputTable.ColumnByCaption

Находит уникальный столбец заголовка, без фиксированного адреса.

```vb
Public Function ColumnByCaption(ByVal headerRow As Long, ByVal caption As String) As Long
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| headerRow | Long; ByVal; обязательный | Строка заголовков, в пределах снимка. |
| caption | String; ByVal; обязательный | Смысловой заголовок столбца; повторные заголовки — ошибка. |

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:161

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.InvalidInputRejected с его подготовленными локальными объектами/входом.

```vb
        Case 6: value = table.ColumnByCaption(1, "Value")
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:528.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunInputTableTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/InputTables.txt.

<a id="api-csectioninputtable-rowbykey"></a>
### CSectionInputTable.RowByKey

Находит уникальную строку смыслового ключа в указанном подблоке.

```vb
Public Function RowByKey(ByVal key As String, Optional ByVal keyColumn As Long = 1, Optional ByVal firstRow As Long = 1, Optional ByVal lastRow As Long = 0) As Long
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| key | String; ByVal; обязательный | Смысловой ключ строки, не Excel-адрес. |
| keyColumn | Long; ByVal; Optional; default=1 | Столбец смысловых ключей 1..ColumnCount. |
| firstRow | Long; ByVal; Optional; default=1 | Начало диапазона поиска строки, включительно. |
| lastRow | Long; ByVal; Optional; default=0 | Конец поиска, включительно; 0 означает RowCount. |

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:177

Минимальный вызов; требуемый контекст: В теле modTestRecipeSetbacks.RunRecipeSetbackInputTests с его подготовленными локальными объектами/входом.

```vb
    main = table.RowByKey("as", 2)
```

Полный контекст: tests/modTestRecipeSetbacks.bas:191.

Native entrypoint: modTestRecipeSetbacks.RunRecipeSetbackInputTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/DeclarativeApi/FinalGeometry.txt.

<a id="api-csectioninputtable-rowbyanykey"></a>
### CSectionInputTable.RowByAnyKey

Находит уникальную строку по массиву допустимых альтернативных ключей; неоднозначность запрещена.

```vb
Public Function RowByAnyKey(ByVal keys As Variant, Optional ByVal keyColumn As Long = 1, Optional ByVal firstRow As Long = 1, Optional ByVal lastRow As Long = 0) As Long
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| keys | Variant; ByVal; обязательный | Массив допустимых альтернативных ключей; неоднозначное совпадение отклоняется. |
| keyColumn | Long; ByVal; Optional; default=1 | Столбец смысловых ключей 1..ColumnCount. |
| firstRow | Long; ByVal; Optional; default=1 | Начало диапазона поиска строки, включительно. |
| lastRow | Long; ByVal; Optional; default=0 | Конец поиска, включительно; 0 означает RowCount. |

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:184

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.InvalidInputRejected с его подготовленными локальными объектами/входом.

```vb
        Case 20: value = table.RowByAnyKey("Count")
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:539.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunInputTableTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/InputTables.txt.

## CSectionRecipeContext

Initialize копирует и закрывает Units→AddTable→Seal на внешней границе. Рецепт только читает независимый снимок Table/Units; не сохраняет context в готовом definition.

<a id="api-csectionrecipecontext-initialize"></a>
### CSectionRecipeContext.Initialize

Открывает контекст один раз и сохраняет независимую закрытую копию подготовленных единиц и знаков.

```vb
Public Sub Initialize(ByVal units As CUnitSystem)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| units | CUnitSystem; ByVal; обязательный | CUnitSystem, уже подготовленная вызывающим workflow; преобразование только через её методы. |

Жизненный цикл: Initialize копирует и закрывает Units→AddTable→Seal на внешней границе. Рецепт только читает независимый снимок Table/Units; не сохраняет context в готовом definition.

Ошибки/исправление: 5530: нет units/table, повтор имени, добавление после Seal. Готовьте новый context на внешней границе.

Source: src/Section/CSectionRecipeContext.cls:19

Минимальный вызов; требуемый контекст: В теле modTestGeometryPublication.RunRecipeContextFreezeTests с его подготовленными локальными объектами/входом.

```vb
    Set context = New CSectionRecipeContext: context.Initialize units
```

Полный контекст: tests/modTestGeometryPublication.bas:69.

Native entrypoint: modTestGeometryPublication.RunRecipeContextFreezeTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/SolidFinal_2026-10-14/PublicationTests/modTestGeometryPublication.RunRecipeContextFreezeTests.txt.

<a id="api-csectionrecipecontext-addtable"></a>
### CSectionRecipeContext.AddTable

Добавляет immutable input snapshot с уникальным именем до Seal.

```vb
Public Sub AddTable(ByVal table As CSectionInputTable)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| table | CSectionInputTable; ByVal; обязательный | Неизменяемый CSectionInputTable. |

Жизненный цикл: Initialize копирует и закрывает Units→AddTable→Seal на внешней границе. Рецепт только читает независимый снимок Table/Units; не сохраняет context в готовом definition.

Ошибки/исправление: 5530: нет units/table, повтор имени, добавление после Seal. Готовьте новый context на внешней границе.

Source: src/Section/CSectionRecipeContext.cls:29

Минимальный вызов; требуемый контекст: В теле modTestGeometryPublication.RunRecipeContextFreezeTests с его подготовленными локальными объектами/входом.

```vb
    context.AddTable table: context.Seal
```

Полный контекст: tests/modTestGeometryPublication.bas:70.

Native entrypoint: modTestGeometryPublication.RunRecipeContextFreezeTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/SolidFinal_2026-10-14/PublicationTests/modTestGeometryPublication.RunRecipeContextFreezeTests.txt.

<a id="api-csectionrecipecontext-seal"></a>
### CSectionRecipeContext.Seal

Закрывает состав таблиц; после этого добавление невозможно.

```vb
Public Sub Seal()
```

Жизненный цикл: Initialize копирует и закрывает Units→AddTable→Seal на внешней границе. Рецепт только читает независимый снимок Table/Units; не сохраняет context в готовом definition.

Ошибки/исправление: 5530: нет units/table, повтор имени, добавление после Seal. Готовьте новый context на внешней границе.

Source: src/Section/CSectionRecipeContext.cls:38

Минимальный вызов; требуемый контекст: В теле modTestGeometryPublication.RunRecipeContextFreezeTests с его подготовленными локальными объектами/входом.

```vb
    context.AddTable table: context.Seal
```

Полный контекст: tests/modTestGeometryPublication.bas:70.

Native entrypoint: modTestGeometryPublication.RunRecipeContextFreezeTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/SolidFinal_2026-10-14/PublicationTests/modTestGeometryPublication.RunRecipeContextFreezeTests.txt.

<a id="api-csectionrecipecontext-table"></a>
### CSectionRecipeContext.Table

Возвращает снимок зарегистрированного блока по имени; отсутствие — ошибка, не fallback.

```vb
Public Function Table(ByVal tableName As String) As CSectionInputTable
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| tableName | String; ByVal; обязательный | Имя единственного корневого входного блока этого снимка. |

Тип результата: CSectionInputTable. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Initialize копирует и закрывает Units→AddTable→Seal на внешней границе. Рецепт только читает независимый снимок Table/Units; не сохраняет context в готовом definition.

Ошибки/исправление: 5530: нет units/table, повтор имени, добавление после Seal. Готовьте новый context на внешней границе.

Source: src/Section/CSectionRecipeContext.cls:44

Минимальный вызов; требуемый контекст: В теле modTestGeometryConfig.CounterFixture с его подготовленными локальными объектами/входом.

```vb
    Set table = context.Table(CStr(names(0)))
```

Полный контекст: tests/modTestGeometryConfig.bas:116.

Native entrypoint: modTestGeometryConfig.RunAudit03RebarCounterTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/ProductionRecipeEvidencePack/data/ProductionRecipe/RebarCounters.txt.

<a id="api-csectionrecipecontext-units"></a>
### CSectionRecipeContext.Units

Закрытая копия системы единиц этого снимка ввода. Внешняя перенастройка units не меняет контекст; методы преобразования доступны, InitializeDefaults/LoadFromSettings запрещены.

```vb
Public Property Get Units() As CUnitSystem
```

Тип результата: CUnitSystem. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Initialize копирует и закрывает Units→AddTable→Seal на внешней границе. Рецепт только читает независимый снимок Table/Units; не сохраняет context в готовом definition.

Ошибки/исправление: 5530: нет units/table, повтор имени, добавление после Seal. Готовьте новый context на внешней границе.

Source: src/Section/CSectionRecipeContext.cls:51

Минимальный вызов; требуемый контекст: В теле modTestGeometryPublication.RunRecipeContextFreezeTests с его подготовленными локальными объектами/входом.

```vb
    Set captured = context.Units
```

Полный контекст: tests/modTestGeometryPublication.bas:71.

Native entrypoint: modTestGeometryPublication.RunRecipeContextFreezeTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/SolidFinal_2026-10-14/PublicationTests/modTestGeometryPublication.RunRecipeContextFreezeTests.txt.

## CSectionInputSnapshotReader

Внешний workbook+список имён+подготовленные units→новый sealed context. Невыбранные блоки не читаются; live Excel не передаётся геометрическому ядру.

<a id="api-csectioninputsnapshotreader-capture"></a>
### CSectionInputSnapshotReader.Capture

На внешней Excel-границе вычисляет и копирует только выбранные именованные блоки с адресами, затем Seal context.

```vb
Public Function Capture(ByVal workbook As Object, ByVal tableNames As Variant, ByVal units As CUnitSystem) As CSectionRecipeContext
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| workbook | Object; ByVal; обязательный | Excel workbook на внешней границе ввода; рецепт этот объект не хранит. |
| tableNames | Variant; ByVal; обязательный | Массив имён блоков выбранного типа; Empty-array допустим для программного рецепта. |
| units | CUnitSystem; ByVal; обязательный | CUnitSystem, уже подготовленная вызывающим workflow; преобразование только через её методы. |

Тип результата: CSectionRecipeContext. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Внешний workbook+список имён+подготовленные units→новый sealed context. Невыбранные блоки не читаются; live Excel не передаётся геометрическому ядру.

Ошибки/исправление: Адресные ошибки имени/блока/формулы. Повреждённый выбранный блок не равен отсутствию потребителя; восстановите имя/ячейку.

Source: src/Excel/CSectionInputSnapshotReader.cls:14

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.RunWorkbookRecipeInputTests с его подготовленными локальными объектами/входом.

```vb
        Set context = reader.Capture(ThisWorkbook, names, units)
```

Полный контекст: tests/modTestSectionRecipes.bas:56.

Native entrypoint: modTestSectionRecipes.RunWorkbookRecipeInputTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/ProductionRecipeEvidencePack/data/ProductionRecipe/WorkbookInput.txt.

## CSectionTypeRegistry

Новая instance регистрирует текущие типы. RegisterRecipe добавляет один экземпляр/metadata; consumer выбирает только зарегистрированный TypeID.

<a id="api-csectiontyperegistry-registerrecipe"></a>
### CSectionTypeRegistry.RegisterRecipe

Регистрирует один ISectionRecipe и массив имён его входных блоков; ID типа уникальный.

```vb
Public Sub RegisterRecipe(ByVal typeID As String, ByVal recipe As ISectionRecipe, ByVal tableNames As Variant)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| typeID | String; ByVal; обязательный | Смысловое имя зарегистрированного типа; непустой текст. |
| recipe | ISectionRecipe; ByVal; обязательный | Один объект с реализацией ISectionRecipe. |
| tableNames | Variant; ByVal; обязательный | Массив имён блоков выбранного типа; Empty-array допустим для программного рецепта. |

Жизненный цикл: Новая instance регистрирует текущие типы. RegisterRecipe добавляет один экземпляр/metadata; consumer выбирает только зарегистрированный TypeID.

Ошибки/исправление: 5530: неизвестный/повторный TypeID, пустой recipe/context, повреждённые metadata. Добавьте одну действительную регистрацию.

Source: src/Excel/CSectionTypeRegistry.cls:28

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.RunWorkbookRecipeInputTests с его подготовленными локальными объектами/входом.

```vb
    registry.RegisterRecipe "FixedMetadata", New CSectionRecipeCircle, Array()
```

Полный контекст: tests/modTestSectionRecipes.bas:80.

Native entrypoint: modTestSectionRecipes.RunWorkbookRecipeInputTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/ProductionRecipeEvidencePack/data/ProductionRecipe/WorkbookInput.txt.

<a id="api-csectiontyperegistry-createrecipe"></a>
### CSectionTypeRegistry.CreateRecipe

Возвращает зарегистрированный объект по TypeID; неизвестный тип — ошибка.

```vb
Public Function CreateRecipe(ByVal typeID As String) As ISectionRecipe
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| typeID | String; ByVal; обязательный | Смысловое имя зарегистрированного типа; непустой текст. |

Тип результата: ISectionRecipe. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Новая instance регистрирует текущие типы. RegisterRecipe добавляет один экземпляр/metadata; consumer выбирает только зарегистрированный TypeID.

Ошибки/исправление: 5530: неизвестный/повторный TypeID, пустой recipe/context, повреждённые metadata. Добавьте одну действительную регистрацию.

Source: src/Excel/CSectionTypeRegistry.cls:45

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiInputAndDataExamples с его подготовленными локальными объектами/входом.

```vb
    Set registry = New CSectionTypeRegistry: Set recipe = registry.CreateRecipe("Circle")
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:225.

Native entrypoint: modTestRecipeApiCatalogue.RunApiInputAndDataExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-csectiontyperegistry-recipetablenames"></a>
### CSectionTypeRegistry.RecipeTableNames

Возвращает независимый массив имён ввода, пустой для программного рецепта.

```vb
Public Function RecipeTableNames(ByVal typeID As String) As Variant
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| typeID | String; ByVal; обязательный | Смысловое имя зарегистрированного типа; непустой текст. |

Тип результата: Variant. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Новая instance регистрирует текущие типы. RegisterRecipe добавляет один экземпляр/metadata; consumer выбирает только зарегистрированный TypeID.

Ошибки/исправление: 5530: неизвестный/повторный TypeID, пустой recipe/context, повреждённые metadata. Добавьте одну действительную регистрацию.

Source: src/Excel/CSectionTypeRegistry.cls:51

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.RunWorkbookRecipeInputTests с его подготовленными локальными объектами/входом.

```vb
    names = registry.RecipeTableNames("Circle")
```

Полный контекст: tests/modTestSectionRecipes.bas:78.

Native entrypoint: modTestSectionRecipes.RunWorkbookRecipeInputTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/ProductionRecipeEvidencePack/data/ProductionRecipe/WorkbookInput.txt.

<a id="api-csectiontyperegistry-registeredrecipetypes"></a>
### CSectionTypeRegistry.RegisteredRecipeTypes

Возвращает фактически зарегистрированные TypeID — источник штатного списка Config.

```vb
Public Function RegisteredRecipeTypes() As Variant
```

Тип результата: Variant. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Новая instance регистрирует текущие типы. RegisterRecipe добавляет один экземпляр/metadata; consumer выбирает только зарегистрированный TypeID.

Ошибки/исправление: 5530: неизвестный/повторный TypeID, пустой recipe/context, повреждённые metadata. Добавьте одну действительную регистрацию.

Source: src/Excel/CSectionTypeRegistry.cls:57

Минимальный вызов; требуемый контекст: В теле modTestRoundedTrapezoidExample.RunExampleBuildTests с его подготовленными локальными объектами/входом.

```vb
    For Each id In registry.RegisteredRecipeTypes
```

Полный контекст: docs/geometry_refactor/examples/modTestRoundedTrapezoidExample.bas:40.

Native entrypoint: modTestRoundedTrapezoidExample.RunExampleBuildTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/OneClassExample/InputAndBuild.txt.

<a id="api-csectiontyperegistry-selectedrecipetype"></a>
### CSectionTypeRegistry.SelectedRecipeType

Читает обязательный Geometry.Type через settings и проверяет его по реестру.

```vb
Public Function SelectedRecipeType(ByVal settings As CSystemSettingsReader) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| settings | CSystemSettingsReader; ByVal; обязательный | CSystemSettingsReader, уже загрузивший Config. |

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Новая instance регистрирует текущие типы. RegisterRecipe добавляет один экземпляр/metadata; consumer выбирает только зарегистрированный TypeID.

Ошибки/исправление: 5530: неизвестный/повторный TypeID, пустой recipe/context, повреждённые metadata. Добавьте одну действительную регистрацию.

Source: src/Excel/CSectionTypeRegistry.cls:62

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.RunWorkbookRecipeInputTests с его подготовленными локальными объектами/входом.

```vb
    Check "workbook.registry.dynamicChoice", registry.SelectedRecipeType(settings) = "FixedMetadata"
```

Полный контекст: tests/modTestSectionRecipes.bas:86.

Native entrypoint: modTestSectionRecipes.RunWorkbookRecipeInputTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/ProductionRecipeEvidencePack/data/ProductionRecipe/WorkbookInput.txt.

<a id="api-csectiontyperegistry-createdefinition"></a>
### CSectionTypeRegistry.CreateDefinition

Вызывает ровно один выбранный ISectionRecipe с подготовленным context.

```vb
Public Function CreateDefinition(ByVal typeID As String, ByVal context As CSectionRecipeContext) As CSectionDefinition
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| typeID | String; ByVal; обязательный | Смысловое имя зарегистрированного типа; непустой текст. |
| context | CSectionRecipeContext; ByVal; обязательный | Подготовленный снимок CSectionRecipeContext, не live Excel. |

Тип результата: CSectionDefinition. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Новая instance регистрирует текущие типы. RegisterRecipe добавляет один экземпляр/metadata; consumer выбирает только зарегистрированный TypeID.

Ошибки/исправление: 5530: неизвестный/повторный TypeID, пустой recipe/context, повреждённые metadata. Добавьте одну действительную регистрацию.

Source: src/Excel/CSectionTypeRegistry.cls:68

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.BuildRecipe с его подготовленными локальными объектами/входом.

```vb
    Set definition = registry.CreateDefinition(typeID, context)
```

Полный контекст: tests/modTestSectionRecipes.bas:863.

Native entrypoint: modTestSectionRecipes.RunRecipeAnnotationCountTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/RoundedRecipeEvidencePack/RoundedRecipe/AnnotationCounts.txt.

Native entrypoint: modTestSectionRecipes.RunRoundedRecipeEquilibriumTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/RoundedRecipeEvidencePack/RoundedRecipe/EquilibriumTests.txt.

Native entrypoint: modTestSectionRecipes.RunHollowRecipeEquilibriumTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/HollowRecipeEvidencePack/data/HollowRecipe/EquilibriumTests.txt.

Native entrypoint: modTestSectionRecipes.RunHollowRecipeEdgeEquilibriumTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/HollowRecipeEvidencePack/data/HollowRecipe/EdgeEquilibriumYes.txt.

## CSectionBuilder

Build проверяет и размещает арматуру, строит сетку и модель, затем аннотации. После успеха Mesh/Rebars/Rules закрыты от изменений и выдаются без копирования массивов. Изменение геометрии модели блокирует getters; ранее полученная ссылка остаётся снимком прежней сборки. Новый Build создаёт новые данные.

<a id="api-csectionbuilder-build"></a>
### CSectionBuilder.Build

Общая сборка: область, проверка и исполнение правил арматуры, сетка, модель и аннотации. После успеха сетка, раскладка и исполнитель закрыты для изменений; большие массивы не копируются.

```vb
Public Function Build(ByVal definition As CSectionDefinition, ByRef meshOptions As TSectionMeshOptions) As CSectionModel
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| definition | CSectionDefinition; ByVal; обязательный | Успешно опубликованный неизменяемый CSectionDefinition. |
| meshOptions | TSectionMeshOptions; ByRef; обязательный | TSectionMeshOptions: StepX/StepY мм>0, BoundarySubdivisions целое≥1, MaterialID≥1. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |

Тип результата: CSectionModel. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Build проверяет и размещает арматуру, строит сетку и модель, затем аннотации. После успеха Mesh/Rebars/Rules закрыты от изменений и выдаются без копирования массивов. Изменение геометрии модели блокирует getters; ранее полученная ссылка остаётся снимком прежней сборки. Новый Build создаёт новые данные.

Ошибки/исправление: 5530: неопубликованное definition, неверные meshOptions, отсутствие баров/успешной сборки. Исправьте ввод; старый результат после отказа не выдаётся.

Source: src/Section/CSectionBuilder.cls:22

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02GuideExamplesTests с его подготовленными локальными объектами/входом.

```vb
    Set model = builder.Build(definition, mesh)
```

Полный контекст: tests/modTestAudit02Api.bas:213.

Native entrypoint: modTestAudit02Api.RunAudit02GuideExamplesTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/FinalApiAcceptance/modTestAudit02Api.RunAudit02GuideExamplesTests.txt.

<a id="api-csectionbuilder-mesh"></a>
### CSectionBuilder.Mesh

Та же закрытая сетка последней успешной сборки. BuildMesh запрещён. При изменении геометрии исходной модели getter отклоняет устаревшую сборку.

```vb
Public Property Get Mesh() As CFiberMeshBuilder
```

Тип результата: CFiberMeshBuilder. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Build проверяет и размещает арматуру, строит сетку и модель, затем аннотации. После успеха Mesh/Rebars/Rules закрыты от изменений и выдаются без копирования массивов. Изменение геометрии модели блокирует getters; ранее полученная ссылка остаётся снимком прежней сборки. Новый Build создаёт новые данные.

Ошибки/исправление: 5530: неопубликованное definition, неверные meshOptions, отсутствие баров/успешной сборки. Исправьте ввод; старый результат после отказа не выдаётся.

Source: src/Section/CSectionBuilder.cls:53

Минимальный вызов; требуемый контекст: В теле modTestGeometryPublication.RunDefinitionLifecycleTests с его подготовленными локальными объектами/входом.

```vb
    Set mesh = builder.Mesh
```

Полный контекст: tests/modTestGeometryPublication.bas:234.

Native entrypoint: modTestGeometryPublication.RunDefinitionLifecycleTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/SolidFinal_2026-10-14/SolidTests/modTestGeometryPublication.RunDefinitionLifecycleTests.txt.

<a id="api-csectionbuilder-rebars"></a>
### CSectionBuilder.Rebars

Та же закрытая физическая раскладка последней успешной сборки, без привязок аннотаций. Clear/AddBar/AddMembership запрещены; массивы не копируются.

```vb
Public Property Get Rebars() As CRebarLayout
```

Тип результата: CRebarLayout. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Build проверяет и размещает арматуру, строит сетку и модель, затем аннотации. После успеха Mesh/Rebars/Rules закрыты от изменений и выдаются без копирования массивов. Изменение геометрии модели блокирует getters; ранее полученная ссылка остаётся снимком прежней сборки. Новый Build создаёт новые данные.

Ошибки/исправление: 5530: неопубликованное definition, неверные meshOptions, отсутствие баров/успешной сборки. Исправьте ввод; старый результат после отказа не выдаётся.

Source: src/Section/CSectionBuilder.cls:60

Минимальный вызов; требуемый контекст: В теле modTestGeometryPublication.RunPublishedBuildTests с его подготовленными локальными объектами/входом.

```vb
    Set sameRebars = builder.Rebars
```

Полный контекст: tests/modTestGeometryPublication.bas:162.

Native entrypoint: modTestGeometryPublication.RunPublishedBuildTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/SolidFinal_2026-10-14/PublicationTests/modTestGeometryPublication.RunPublishedBuildTests.txt.

<a id="api-csectionbuilder-rules"></a>
### CSectionBuilder.Rules

Тот же закрытый исполнитель с фактическими счётчиками, шагами и происхождением. Повторный Build запрещён; диагностические методы доступны.

```vb
Public Property Get Rules() As CRebarRuleExecutor
```

Тип результата: CRebarRuleExecutor. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Build проверяет и размещает арматуру, строит сетку и модель, затем аннотации. После успеха Mesh/Rebars/Rules закрыты от изменений и выдаются без копирования массивов. Изменение геометрии модели блокирует getters; ранее полученная ссылка остаётся снимком прежней сборки. Новый Build создаёт новые данные.

Ошибки/исправление: 5530: неопубликованное definition, неверные meshOptions, отсутствие баров/успешной сборки. Исправьте ввод; старый результат после отказа не выдаётся.

Source: src/Section/CSectionBuilder.cls:67

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    Set rules = builder.Rules
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:121.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

## CRebarRuleExecutor

До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

<a id="api-crebarruleexecutor-requestedpositions"></a>
### CRebarRuleExecutor.RequestedPositions

Все запрошенные позиции, включая пропущенные проекции; это не число физических стержней.

```vb
Public Property Get RequestedPositions() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:36

Минимальный вызов; требуемый контекст: В теле modTestGeometry.TestRectSetAutoRebarLayout с его подготовленными локальными объектами/входом.

```vb
    AssertTrue stats, "rectset.rebar.requested", builder.RequestedPositions = 8
```

Полный контекст: tests/modTestGeometry.bas:1220.

Native entrypoint: modTestGeometry.RunGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-crebarruleexecutor-createdpositions"></a>
### CRebarRuleExecutor.CreatedPositions

Действительно добавленные физические стержни текущего Build, даже при отказе в середине.

```vb
Public Property Get CreatedPositions() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:38

Минимальный вызов; требуемый контекст: В теле modTestGeometry.TestAudit03CircleCountGuard с его подготовленными локальными объектами/входом.

```vb
        AssertClose stats, prefix & ".beforeCoordinates", CDbl(builder.CreatedPositions), 0#, 0#
```

Полный контекст: tests/modTestGeometry.bas:2654.

Native entrypoint: modTestGeometry.RunGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-crebarruleexecutor-skippedprojections"></a>
### CRebarRuleExecutor.SkippedProjections

Число явно разрешённых пропусков target hit.

```vb
Public Property Get SkippedProjections() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:40

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunAdvancedGuideExamples с его подготовленными локальными объектами/входом.

```vb
        Check "advanced.projection.skips." & CStr(useMaxStep), rules.SkippedProjections = 2
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:22.

Native entrypoint: modTestRecipeApiCatalogue.RunAdvancedGuideExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/AdvancedExamples.txt.

<a id="api-crebarruleexecutor-mergedpositions"></a>
### CRebarRuleExecutor.MergedPositions

Число заявленных позиций, объединённых с физическим владельцем.

```vb
Public Property Get MergedPositions() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:42

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
    Check "merge.counts", rules.RequestedPositions = 2 And rules.CreatedPositions = 1 And rules.MergedPositions = 1
```

Полный контекст: tests/modTestRecipeOperations.bas:165.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-crebarruleexecutor-conflicts"></a>
### CRebarRuleExecutor.Conflicts

Число несовместимых/отклонённых совпадений или запрещённых пересечений по явной политике.

```vb
Public Property Get Conflicts() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:44

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
            Check "overlap.error.actual.created", rules.RequestedPositions = 2 And rules.CreatedPositions = 1 And rules.Conflicts = 1
```

Полный контекст: tests/modTestRecipeOperations.bas:115.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-crebarruleexecutor-baroverlaps"></a>
### CRebarRuleExecutor.BarOverlaps

Полное число пар с отрицательным геометрическим зазором дисков; касание не включается.

```vb
Public Property Get BarOverlaps() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:46

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
            Check "overlap.report." & CStr(mode), rules.BarOverlaps = 1
```

Полный контекст: tests/modTestRecipeOperations.bas:119.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-crebarruleexecutor-diagnosticrecordsomitted"></a>
### CRebarRuleExecutor.DiagnosticRecordsOmitted

Число событий сверх 128 сохранённых подробных образцов; полные счётчики не сокращаются.

```vb
Public Property Get DiagnosticRecordsOmitted() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:48

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
    Check "diagnostics.bounded.records", rules.DiagnosticCount = 128 And rules.DiagnosticRecordsOmitted = 1582
```

Полный контекст: tests/modTestRecipeOperations.bas:159.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-crebarruleexecutor-lastdiagnostic"></a>
### CRebarRuleExecutor.LastDiagnostic

Последняя машинная диагностика; Nothing при отсутствии записанного события.

```vb
Public Property Get LastDiagnostic() As CGeometryDiagnostic
```

Тип результата: CGeometryDiagnostic. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:50

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
            Set diagnostic = rules.LastDiagnostic
```

Полный контекст: tests/modTestRecipeOperations.bas:111.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-crebarruleexecutor-diagnosticcount"></a>
### CRebarRuleExecutor.DiagnosticCount

Число сохранённых подробных immutable образцов, 0..128.

```vb
Public Property Get DiagnosticCount() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:52

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
    Check "diagnostics.bounded.records", rules.DiagnosticCount = 128 And rules.DiagnosticRecordsOmitted = 1582
```

Полный контекст: tests/modTestRecipeOperations.bas:159.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-crebarruleexecutor-diagnosticat"></a>
### CRebarRuleExecutor.DiagnosticAt

Читает immutable подробный образец 1..DiagnosticCount; максимум 128 образцов, счётчики полные.

```vb
Public Function DiagnosticAt(ByVal index As Long) As CGeometryDiagnostic
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Тип результата: CGeometryDiagnostic. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:57

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
            Check "overlap.machine." & CStr(mode), rules.DiagnosticAt(1).Code = "BAR_OVERLAP"
```

Полный контекст: tests/modTestRecipeOperations.bas:120.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-crebarruleexecutor-positioncount"></a>
### CRebarRuleExecutor.PositionCount

Число сгенерированных заявленных позиций до объединения физических владельцев.

```vb
Public Property Get PositionCount() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:63

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    Check "rules.positionCount", rules.PositionCount = 3
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:122.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-crebarruleexecutor-positionruleid"></a>
### CRebarRuleExecutor.PositionRuleID

ID правила заявленной позиции, включая объединённые позиции.

```vb
Public Function PositionRuleID(ByVal index As Long) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:65

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
            Check "trace.ids." & CStr(mode), rules.PositionRuleID(1) = "Bars" And rules.PositionPathID(1) = "Row" And rules.PositionSourceID(1) = "Bars.1"
```

Полный контекст: tests/modTestRecipeOperations.bas:123.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-crebarruleexecutor-positionpathid"></a>
### CRebarRuleExecutor.PositionPathID

ID геометрического пути заявленной позиции.

```vb
Public Function PositionPathID(ByVal index As Long) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:69

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
            Check "trace.ids." & CStr(mode), rules.PositionRuleID(1) = "Bars" And rules.PositionPathID(1) = "Row" And rules.PositionSourceID(1) = "Bars.1"
```

Полный контекст: tests/modTestRecipeOperations.bas:123.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-crebarruleexecutor-positionsourceid"></a>
### CRebarRuleExecutor.PositionSourceID

Смысловой SourcePositionId, из которого публиковался физический стержень.

```vb
Public Function PositionSourceID(ByVal index As Long) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:71

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
            Check "trace.ids." & CStr(mode), rules.PositionRuleID(1) = "Bars" And rules.PositionPathID(1) = "Row" And rules.PositionSourceID(1) = "Bars.1"
```

Полный контекст: tests/modTestRecipeOperations.bas:123.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-crebarruleexecutor-positionphysicalowner"></a>
### CRebarRuleExecutor.PositionPhysicalOwner

Индекс физического владельца в итоговом CRebarLayout; объединённые позиции имеют одного владельца.

```vb
Public Function PositionPhysicalOwner(ByVal index As Long) As Long
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:74

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
    Check "merge.physical.owner", rules.PositionPhysicalOwner(1) = 1 And rules.PositionPhysicalOwner(2) = 1
```

Полный контекст: tests/modTestRecipeOperations.bas:166.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-crebarruleexecutor-ruletrace"></a>
### CRebarRuleExecutor.RuleTrace

Текстовый отчёт конкретного правила с фактическими counts и шагами/хордами.

```vb
Public Function RuleTrace(ByVal ruleID As String) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:78

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
    Check "merge.trace", InStr(rules.RuleTrace("Second"), "Created=0; Merged=1") > 0
```

Полный контекст: tests/modTestRecipeOperations.bas:167.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-crebarruleexecutor-rulestepbounds"></a>
### CRebarRuleExecutor.RuleStepBounds

Возвращает число промежутков между стержнями, наименьшее и наибольшее расстояния между ними вдоль линии и по прямой, мм. Для замкнутого ряда учитывает промежуток от последнего стержня до первого.

```vb
Public Sub RuleStepBounds(ByVal ruleID As String, ByRef stepCount As Long, ByRef minStep As Double, ByRef maxStep As Double, ByRef minChord As Double, ByRef maxChord As Double)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |
| stepCount | Long; ByRef; обязательный | Целый неотрицательный счётчик. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| minStep | Double; ByRef; обязательный | Расстояние или граница интервала minStep, внутренние мм; направление и допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| maxStep | Double; ByRef; обязательный | Расстояние или граница интервала maxStep, внутренние мм; направление и допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| minChord | Double; ByRef; обязательный | Расстояние или граница интервала minChord, внутренние мм; направление и допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| maxChord | Double; ByRef; обязательный | Расстояние или граница интервала maxChord, внутренние мм; направление и допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:93

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
    rules.RuleStepBounds "RingBars", stepCount, minStep, maxStep, minChord, maxChord
```

Полный контекст: tests/modTestRecipeOperations.bas:147.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-crebarruleexecutor-diagnosticstext"></a>
### CRebarRuleExecutor.DiagnosticsText

Полный текстовый отчёт счётчиков и всех RuleTrace; численные длины в мм независимо от output units.

```vb
Public Function DiagnosticsText() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:133

Минимальный вызов; требуемый контекст: В теле modTestGeometryPublication.RunPublishedBuildTests с его подготовленными локальными объектами/входом.

```vb
    before = LayoutSnapshot(rebars): trace = rules.DiagnosticsText
```

Полный контекст: tests/modTestGeometryPublication.bas:128.

Native entrypoint: modTestGeometryPublication.RunPublishedBuildTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/SolidFinal_2026-10-14/PublicationTests/modTestGeometryPublication.RunPublishedBuildTests.txt.

<a id="api-crebarruleexecutor-build"></a>
### CRebarRuleExecutor.Build

Исполняет опубликованные правила в готовой бетонной области и возвращает физическую раскладку; mesh не строит.

```vb
Public Function Build(ByVal definition As CSectionDefinition, ByVal geometry As CGeometryRegion) As CRebarLayout
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| definition | CSectionDefinition; ByVal; обязательный | Успешно опубликованный неизменяемый CSectionDefinition. |
| geometry | CGeometryRegion; ByVal; обязательный | Подготовленный CGeometryRegion материала для исполнения правил. |

Тип результата: CRebarLayout. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:146

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
    Set layout = rules.Build(Nothing, Nothing)
```

Полный контекст: tests/modTestRecipeOperations.bas:176.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-crebarruleexecutor-isreadonly"></a>
### CRebarRuleExecutor.IsReadOnly

True после публикации успешной общей сборки. Позиции и диагностика сохраняются; новый Build требует нового исполнителя.

```vb
Public Property Get IsReadOnly() As Boolean
```

Тип результата: Boolean. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: До публикации Build начинает новый отчёт и позиции. После закрытия сборщиком повторный Build запрещён и сохраняет диагностику. Позиции и происхождение доступны для чтения.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:774

Минимальный вызов; требуемый контекст: В теле modTestGeometryPublication.RunPublishedBuildTests с его подготовленными локальными объектами/входом.

```vb
    Check "published.readOnly", mesh.IsReadOnly And rebars.IsReadOnly And rules.IsReadOnly
```

Полный контекст: tests/modTestGeometryPublication.bas:123.

Native entrypoint: modTestGeometryPublication.RunPublishedBuildTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/SolidFinal_2026-10-14/PublicationTests/modTestGeometryPublication.RunPublishedBuildTests.txt.

## CUnitSystem

InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

<a id="api-cunitsystem-initializedefaults"></a>
### CUnitSystem.InitializeDefaults

Задаёт программные defaults: input mm/mm²/N/N*mm/MPa/1/mm и Tension, output m/m²/N/N*mm/MPa/1/mm. Не заменяет обязательный Config ввод.

```vb
Public Sub InitializeDefaults()
```

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:45

Минимальный вызов; требуемый контекст: В теле modTestGeometryPublication.RunRecipeContextFreezeTests с его подготовленными локальными объектами/входом.

```vb
    copy.InitializeDefaults
```

Полный контекст: tests/modTestGeometryPublication.bas:93.

Native entrypoint: modTestGeometryPublication.RunRecipeContextFreezeTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/SolidFinal_2026-10-14/PublicationTests/modTestGeometryPublication.RunRecipeContextFreezeTests.txt.

<a id="api-cunitsystem-loadfromsettings"></a>
### CUnitSystem.LoadFromSettings

Загружает и проверяет все активные единицы и знаки через общий settings reader.

```vb
Public Sub LoadFromSettings(ByVal settings As CSystemSettingsReader)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| settings | CSystemSettingsReader; ByVal; обязательный | CSystemSettingsReader, уже загрузивший Config. |

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:61

Минимальный вызов; требуемый контекст: В теле modTestGeometryPublication.RunRecipeContextFreezeTests с его подготовленными локальными объектами/входом.

```vb
    captured.LoadFromSettings settings
```

Полный контекст: tests/modTestGeometryPublication.bas:83.

Native entrypoint: modTestGeometryPublication.RunRecipeContextFreezeTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/SolidFinal_2026-10-14/PublicationTests/modTestGeometryPublication.RunRecipeContextFreezeTests.txt.

<a id="api-cunitsystem-clone"></a>
### CUnitSystem.Clone

Возвращает независимую изменяемую копию всех единиц и знаков. Изменение копии не влияет на исходный объект или контекст.

```vb
Public Function Clone() As CUnitSystem
```

Тип результата: CUnitSystem. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:86

Минимальный вызов; требуемый контекст: В теле modTestGeometryPublication.RunRecipeContextFreezeTests с его подготовленными локальными объектами/входом.

```vb
    Set copy = captured.Clone
```

Полный контекст: tests/modTestGeometryPublication.bas:91.

Native entrypoint: modTestGeometryPublication.RunRecipeContextFreezeTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/SolidFinal_2026-10-14/PublicationTests/modTestGeometryPublication.RunRecipeContextFreezeTests.txt.

<a id="api-cunitsystem-isreadonly"></a>
### CUnitSystem.IsReadOnly

True для конфигурации единиц, принадлежащей снимку ввода; перенастройка запрещена, преобразование чисел доступно.

```vb
Public Property Get IsReadOnly() As Boolean
```

Тип результата: Boolean. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:99

Минимальный вызов; требуемый контекст: В теле modTestGeometryPublication.RunRecipeContextFreezeTests с его подготовленными локальными объектами/входом.

```vb
    Check "units.readOnly", captured.IsReadOnly
```

Полный контекст: tests/modTestGeometryPublication.bas:73.

Native entrypoint: modTestGeometryPublication.RunRecipeContextFreezeTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/SolidFinal_2026-10-14/PublicationTests/modTestGeometryPublication.RunRecipeContextFreezeTests.txt.

<a id="api-cunitsystem-inputlengthtointernal"></a>
### CUnitSystem.InputLengthToInternal

Переводит InputLengthToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function InputLengthToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:143

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.inputLength", units.InputLengthToInternal(20#), 20#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:170.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-internallengthtooutput"></a>
### CUnitSystem.InternalLengthToOutput

Переводит InternalLengthToOutput из внутренних единиц в подготовленные output units; сохранённая единица при наличии имеет явный смысл.

```vb
Public Function InternalLengthToOutput(ByVal value As Double, Optional ByVal savedUnit As String = vbNullString) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |
| savedUnit | String; ByVal; Optional; default=vbNullString | Единица сохранённого Results; пустая использует текущую output-единицу CUnitSystem. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:149

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.outputLength", units.InternalLengthToOutput(2000#), 2#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:171.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-outputlengthtointernal"></a>
### CUnitSystem.OutputLengthToInternal

Переводит OutputLengthToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function OutputLengthToInternal(ByVal value As Double, Optional ByVal savedUnit As String = vbNullString) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |
| savedUnit | String; ByVal; Optional; default=vbNullString | Единица сохранённого Results; пустая использует текущую output-единицу CUnitSystem. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:156

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.savedLength", units.OutputLengthToInternal(2#, "cm"), 20#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:172.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-inputareatointernal"></a>
### CUnitSystem.InputAreaToInternal

Переводит InputAreaToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function InputAreaToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:162

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.inputArea", units.InputAreaToInternal(100#), 100#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:173.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-internalareatooutput"></a>
### CUnitSystem.InternalAreaToOutput

Переводит InternalAreaToOutput из внутренних единиц в подготовленные output units; сохранённая единица при наличии имеет явный смысл.

```vb
Public Function InternalAreaToOutput(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:167

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.outputArea", units.InternalAreaToOutput(1000000#), 1#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:174.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-outputareatointernal"></a>
### CUnitSystem.OutputAreaToInternal

Переводит OutputAreaToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function OutputAreaToInternal(ByVal value As Double, Optional ByVal savedUnit As String = vbNullString) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |
| savedUnit | String; ByVal; Optional; default=vbNullString | Единица сохранённого Results; пустая использует текущую output-единицу CUnitSystem. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:173

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.savedArea", units.OutputAreaToInternal(2#, "cm2"), 200#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:175.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-inputfourthpowerlengthtointernal"></a>
### CUnitSystem.InputFourthPowerLengthToInternal

Переводит InputFourthPowerLengthToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function InputFourthPowerLengthToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:179

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.inputFourth", units.InputFourthPowerLengthToInternal(100#), 100#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:176.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-internalfourthpowerlengthtooutput"></a>
### CUnitSystem.InternalFourthPowerLengthToOutput

Переводит InternalFourthPowerLengthToOutput из внутренних единиц в подготовленные output units; сохранённая единица при наличии имеет явный смысл.

```vb
Public Function InternalFourthPowerLengthToOutput(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:184

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.outputFourth", units.InternalFourthPowerLengthToOutput(1000000000000#), 1#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:177.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-outputfourthpowerlengthtointernal"></a>
### CUnitSystem.OutputFourthPowerLengthToInternal

Переводит OutputFourthPowerLengthToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function OutputFourthPowerLengthToInternal(ByVal value As Double, Optional ByVal savedUnit As String = vbNullString) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |
| savedUnit | String; ByVal; Optional; default=vbNullString | Базовая единица длины сохранённого Results: mm/cm/m; степень 4 применяет сам метод. Не передавайте mm4/cm4/m4. Пустая строка использует текущую output-длину. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:190

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.savedFourth", units.OutputFourthPowerLengthToInternal(2#, "cm"), 20000#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:178.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-inputforcetointernal"></a>
### CUnitSystem.InputForceToInternal

Переводит InputForceToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function InputForceToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:196

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.force.input", units.InputForceToInternal(123#), 123#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:179.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-internalforcetooutput"></a>
### CUnitSystem.InternalForceToOutput

Переводит InternalForceToOutput из внутренних единиц в подготовленные output units; сохранённая единица при наличии имеет явный смысл.

```vb
Public Function InternalForceToOutput(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:201

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.force.output", units.InternalForceToOutput(123#), 123#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:180.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-internalforcelengthsquaredtooutput"></a>
### CUnitSystem.InternalForceLengthSquaredToOutput

Переводит InternalForceLengthSquaredToOutput из внутренних единиц в подготовленные output units; сохранённая единица при наличии имеет явный смысл.

```vb
Public Function InternalForceLengthSquaredToOutput(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:208

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.stiffness", units.InternalForceLengthSquaredToOutput(1000000#), 1#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:181.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-inputmomentmxtointernal"></a>
### CUnitSystem.InputMomentMxToInternal

Переводит InputMomentMxToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function InputMomentMxToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:213

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.moment.inputX", units.InputMomentMxToInternal(12#), 12#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:182.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-inputmomentmytointernal"></a>
### CUnitSystem.InputMomentMyToInternal

Переводит InputMomentMyToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function InputMomentMyToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:218

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.moment.inputY", units.InputMomentMyToInternal(-12#), -12#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:183.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-inputmomentperlengthtointernal"></a>
### CUnitSystem.InputMomentPerLengthToInternal

Переводит InputMomentPerLengthToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function InputMomentPerLengthToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:225

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.momentPerLength", units.InputMomentPerLengthToInternal(12#), 12#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:184.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-internalmomentmxtooutput"></a>
### CUnitSystem.InternalMomentMxToOutput

Переводит InternalMomentMxToOutput из внутренних единиц в подготовленные output units; сохранённая единица при наличии имеет явный смысл.

```vb
Public Function InternalMomentMxToOutput(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:230

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.moment.outputX", units.InternalMomentMxToOutput(12#), 12#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:185.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-internalmomentmytooutput"></a>
### CUnitSystem.InternalMomentMyToOutput

Переводит InternalMomentMyToOutput из внутренних единиц в подготовленные output units; сохранённая единица при наличии имеет явный смысл.

```vb
Public Function InternalMomentMyToOutput(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:235

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.moment.outputY", units.InternalMomentMyToOutput(-12#), -12#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:186.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-internalmomentmagnitudetooutput"></a>
### CUnitSystem.InternalMomentMagnitudeToOutput

Переводит InternalMomentMagnitudeToOutput из внутренних единиц в подготовленные output units; сохранённая единица при наличии имеет явный смысл.

```vb
Public Function InternalMomentMagnitudeToOutput(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:240

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.moment.magnitude", units.InternalMomentMagnitudeToOutput(12#), 12#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:187.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-inputstresstointernal"></a>
### CUnitSystem.InputStressToInternal

Переводит InputStressToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function InputStressToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:245

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.stress.input", units.InputStressToInternal(12#), 12#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:188.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-internalstresstooutput"></a>
### CUnitSystem.InternalStressToOutput

Переводит InternalStressToOutput из внутренних единиц в подготовленные output units; сохранённая единица при наличии имеет явный смысл.

```vb
Public Function InternalStressToOutput(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:250

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.stress.output", units.InternalStressToOutput(12#), 12#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:189.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-outputstresstointernal"></a>
### CUnitSystem.OutputStressToInternal

Переводит OutputStressToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function OutputStressToInternal(ByVal value As Double, Optional ByVal savedUnit As String = vbNullString) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |
| savedUnit | String; ByVal; Optional; default=vbNullString | Единица сохранённого Results; пустая использует текущую output-единицу CUnitSystem. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:256

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.stress.saved", units.OutputStressToInternal(12#, "MPa"), 12#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:190.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-inputcurvaturetointernal"></a>
### CUnitSystem.InputCurvatureToInternal

Переводит InputCurvatureToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function InputCurvatureToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:262

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.curvature.input", units.InputCurvatureToInternal(0.01), 0.01
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:191.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-internalcurvaturetooutput"></a>
### CUnitSystem.InternalCurvatureToOutput

Переводит InternalCurvatureToOutput из внутренних единиц в подготовленные output units; сохранённая единица при наличии имеет явный смысл.

```vb
Public Function InternalCurvatureToOutput(ByVal value As Double, Optional ByVal savedUnit As String = vbNullString) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |
| savedUnit | String; ByVal; Optional; default=vbNullString | Единица сохранённого Results; пустая использует текущую output-единицу CUnitSystem. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:268

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.curvature.output", units.InternalCurvatureToOutput(0.001, "1/m"), 1#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:192.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-outputcurvaturetointernal"></a>
### CUnitSystem.OutputCurvatureToInternal

Переводит OutputCurvatureToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function OutputCurvatureToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:274

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Near "units.curvature.inverse", units.OutputCurvatureToInternal(0.001), 0.001
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:193.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-outputlengthunit"></a>
### CUnitSystem.OutputLengthUnit

Возвращает текущий output selector LengthUnit, без преобразования или округления данных.

```vb
Public Property Get OutputLengthUnit() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:279

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Check "units.selector.length", units.OutputLengthUnit = "m"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:194.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-outputareaunit"></a>
### CUnitSystem.OutputAreaUnit

Возвращает текущий output selector AreaUnit, без преобразования или округления данных.

```vb
Public Property Get OutputAreaUnit() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:284

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Check "units.selector.area", units.OutputAreaUnit = "m2"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:195.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-outputmomentunit"></a>
### CUnitSystem.OutputMomentUnit

Возвращает текущий output selector MomentUnit, без преобразования или округления данных.

```vb
Public Property Get OutputMomentUnit() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:289

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Check "units.selector.moment", units.OutputMomentUnit = "N*mm"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:197.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-inputmomentperlengthunit"></a>
### CUnitSystem.InputMomentPerLengthUnit

Возвращает единицу пользовательского момента на длину по input force/moment/length settings.

```vb
Public Property Get InputMomentPerLengthUnit() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:294

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Check "units.selector.momentPerLength", units.InputMomentPerLengthUnit = "N*mm/mm"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:198.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-outputforceunit"></a>
### CUnitSystem.OutputForceUnit

Возвращает текущий output selector ForceUnit, без преобразования или округления данных.

```vb
Public Property Get OutputForceUnit() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:299

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Check "units.selector.force", units.OutputForceUnit = "N"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:196.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-outputforcelengthsquaredunit"></a>
### CUnitSystem.OutputForceLengthSquaredUnit

Возвращает текущий output selector ForceLengthSquaredUnit, без преобразования или округления данных.

```vb
Public Property Get OutputForceLengthSquaredUnit() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:304

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Check "units.selector.stiffness", units.OutputForceLengthSquaredUnit = "N*m2"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:199.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-outputstressunit"></a>
### CUnitSystem.OutputStressUnit

Возвращает текущий output selector StressUnit, без преобразования или округления данных.

```vb
Public Property Get OutputStressUnit() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:309

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Check "units.selector.stress", units.OutputStressUnit = "MPa"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:200.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-outputcurvatureunit"></a>
### CUnitSystem.OutputCurvatureUnit

Возвращает текущий output selector CurvatureUnit, без преобразования или округления данных.

```vb
Public Property Get OutputCurvatureUnit() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:314

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Check "units.selector.curvature", units.OutputCurvatureUnit = "1/mm"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:201.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-outputsignn"></a>
### CUnitSystem.OutputSignN

Возвращает текущий output selector SignN, без преобразования или округления данных.

```vb
Public Property Get OutputSignN() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:319

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Check "units.sign.N", units.OutputSignN = "Tension"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:202.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-outputsignmx"></a>
### CUnitSystem.OutputSignMx

Возвращает текущий output selector SignMx, без преобразования или округления данных.

```vb
Public Property Get OutputSignMx() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:324

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Check "units.sign.Mx", units.OutputSignMx = "+Y tension"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:203.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-cunitsystem-outputsignmy"></a>
### CUnitSystem.OutputSignMy

Возвращает текущий output selector SignMy, без преобразования или округления данных.

```vb
Public Property Get OutputSignMy() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. Контекст хранит отдельную закрытую копию; Clone возвращает новую изменяемую систему. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:329

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiUnitExamples с его подготовленными локальными объектами/входом.

```vb
    Check "units.sign.My", units.OutputSignMy = "+X tension"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:204.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

## CConcreteRegion

Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

<a id="api-cconcreteregion-setowner"></a>
### CConcreteRegion.SetOwner

Назначает скалярное происхождение инженерной области без пересчёта A/контуров; это внешний coordinator API.

```vb
Public Sub SetOwner(ByVal combinationID As String, ByVal standard As String, ByVal anchorID As String)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| combinationID | String; ByVal; обязательный | Скалярная metadata владельца/источника; геометрическая классификация от строки не зависит. |
| standard | String; ByVal; обязательный | Скалярная metadata владельца/источника; геометрическая классификация от строки не зависит. |
| anchorID | String; ByVal; обязательный | Непустой уникальный ID модельной привязки. |

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:76

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestAnalyticClipping с его подготовленными локальными объектами/входом.

```vb
    clipped.SetOwner "LC1", "TEST", "R1"
```

Полный контекст: tests/modTestGeometryQuery.bas:148.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-cconcreteregion-getsegment"></a>
### CConcreteRegion.GetSegment

Возвращает точную аналитическую запись одного ребра в скаляры: мм, sweep радианы, loop-индекс.

```vb
Public Sub GetSegment(ByVal index As Long, ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double, ByRef centerX As Double, ByRef centerY As Double, ByRef sweep As Double, ByRef loopID As Long)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |
| x1 | Double; ByRef; обязательный | Модельная координата x1, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| y1 | Double; ByRef; обязательный | Модельная координата y1, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| x2 | Double; ByRef; обязательный | Модельная координата x2, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| y2 | Double; ByRef; обязательный | Модельная координата y2, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| centerX | Double; ByRef; обязательный | Модельная координата centerX, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| centerY | Double; ByRef; обязательный | Модельная координата centerY, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| sweep | Double; ByRef; обязательный | Signed sweep, радианы: CCW положителен, CW отрицателен, 0 запрещён, \|sweep\|≤2*pi. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| loopID | Long; ByRef; обязательный | Непустой устойчивый ID контура; у CConcreteRegion численный индекс loop 1..LoopCount. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:82

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    region.GetSegment 1, x, y, x2, y2, cx, cy, sweep, loopID
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:133.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-cconcreteregion-segmentcount"></a>
### CConcreteRegion.SegmentCount

Число аналитических line/arc сегментов.

```vb
Public Property Get SegmentCount() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:93

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.CheckClosed с его подготовленными локальными объектами/входом.

```vb
    For i = 1 To region.SegmentCount
```

Полный контекст: tests/modTestGeometryQuery.bas:94.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

Native entrypoint: modTestGeometryQuery.RunGeneratedContourOwnershipTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-cconcreteregion-boundaryprobes"></a>
### CConcreteRegion.BoundaryProbes

Копия сохранённых пассивных пар P/Q; Empty допустим, контуры по ним не восстанавливаются.

```vb
Public Property Get BoundaryProbes() As Variant
```

Тип результата: Variant. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:99

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestPostAudit03RegionClone с его подготовленными локальными объектами/входом.

```vb
    readProbes = region.BoundaryProbes
```

Полный контекст: tests/modTestGeometryQuery.bas:381.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-cconcreteregion-boundaryprobecount"></a>
### CConcreteRegion.BoundaryProbeCount

Число сохранённых диагностических проб P/Q.

```vb
Public Property Get BoundaryProbeCount() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:103

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestPostAudit03RegionClone с его подготовленными локальными объектами/входом.

```vb
    Check "postAudit03.clone.probesIndependent", readProbes(1, 2) = 1# And region.BoundaryProbeCount = 2
```

Полный контекст: tests/modTestGeometryQuery.bas:382.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-cconcreteregion-loopcount"></a>
### CConcreteRegion.LoopCount

Число подготовленных материальных loops.

```vb
Public Property Get LoopCount() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:107

Минимальный вызов; требуемый контекст: В теле modTestAutoCADContours.CheckNestedRegion с его подготовленными локальными объектами/входом.

```vb
    For i = 1 To region.LoopCount
```

Полный контекст: tests/modTestAutoCADContours.bas:1024.

Native entrypoint: modTestAutoCADContours.RunAutoCADContourTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalAcceptance/FakeCAD.txt.

<a id="api-cconcreteregion-loopisopening"></a>
### CConcreteRegion.LoopIsOpening

Читает классификацию конкретной loop по подготовленной топологии.

```vb
Public Property Get LoopIsOpening(ByVal index As Long) As Boolean
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Тип результата: Boolean. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:113

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestPostAudit03RegionClone с его подготовленными локальными объектами/входом.

```vb
        Check "postAudit03.clone.role." & CStr(i), copy.LoopIsOpening(i) = region.LoopIsOpening(i) And nextCopy.LoopIsOpening(i) = region.LoopIsOpening(i)
```

Полный контекст: tests/modTestGeometryQuery.bas:376.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-cconcreteregion-area"></a>
### CConcreteRegion.Area

Аналитическая площадь материальной области, мм²; не механическая A сетки.

```vb
Public Property Get Area() As Double
```

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:118

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.TestAdjacentRectangles с его подготовленными локальными объектами/входом.

```vb
    Near "adjacent.area", region.Area, 240000#
```

Полный контекст: tests/modTestSectionRecipes.bas:1066.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-cconcreteregion-source"></a>
### CConcreteRegion.Source

Скалярное происхождение подготовленной области.

```vb
Public Property Get Source() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:122

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestKnownAndUnknownContours с его подготовленными локальными объектами/входом.

```vb
    Check "knownOuter.noOpening.authoritative", region.Source = "AuthoritativeContour"
```

Полный контекст: tests/modTestGeometryQuery.bas:604.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-cconcreteregion-sectionrevision"></a>
### CConcreteRegion.SectionRevision

Ревизия исходной модели, к которой относится область.

```vb
Public Property Get SectionRevision() As Long
```

Тип результата: Long. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:126

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestKnownAndUnknownContours с его подготовленными локальными объектами/входом.

```vb
    revision = region.SectionRevision
```

Полный контекст: tests/modTestGeometryQuery.bas:619.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-cconcreteregion-ownerid"></a>
### CConcreteRegion.OwnerID

Инженерный владелец области; назначается coordinator, не геометрической классификацией.

```vb
Public Property Get OwnerID() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:130

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestAnalyticClipping с его подготовленными локальными объектами/входом.

```vb
    Check "copy.owner", copy.OwnerID = "LC1" And copy.AnchorID = "R1"
```

Полный контекст: tests/modTestGeometryQuery.bas:150.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-cconcreteregion-standard"></a>
### CConcreteRegion.Standard

Скалярная нормативная принадлежность инженерного владельца; geometry не вычисляет нормативные коэффициенты.

```vb
Public Property Get Standard() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:135

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestPostAudit03RegionClone с его подготовленными локальными объектами/входом.

```vb
    Check "postAudit03.clone.owner", copy.OwnerID = "COPY_LC" And copy.Standard = "SP35" And copy.AnchorID = "G7"
```

Полный контекст: tests/modTestGeometryQuery.bas:365.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-cconcreteregion-anchorid"></a>
### CConcreteRegion.AnchorID

Скалярный ID сохранённой инженерной области.

```vb
Public Property Get AnchorID() As String
```

Тип результата: String. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:139

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestAnalyticClipping с его подготовленными локальными объектами/входом.

```vb
    Check "copy.owner", copy.OwnerID = "LC1" And copy.AnchorID = "R1"
```

Полный контекст: tests/modTestGeometryQuery.bas:150.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-cconcreteregion-clone"></a>
### CConcreteRegion.Clone

Копирует готовую область и metadata; временные массивы следующей операции не меняют исходный snapshot.

```vb
Public Function Clone() As CConcreteRegion
```

Тип результата: CConcreteRegion. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:145

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestAnalyticClipping с его подготовленными локальными объектами/входом.

```vb
    Set copy = clipped.Clone
```

Полный контекст: tests/modTestGeometryQuery.bas:145.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

## CSectionGeometryQuery

Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

<a id="api-csectiongeometryquery-lastdiagnostic"></a>
### CSectionGeometryQuery.LastDiagnostic

Последняя машинная диагностика; Nothing при отсутствии записанного события.

```vb
Public Property Get LastDiagnostic() As CGeometryDiagnostic
```

Тип результата: CGeometryDiagnostic. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:32

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.InvalidOperation с его подготовленными локальными объектами/входом.

```vb
    Set diagnostic = query.LastDiagnostic
```

Полный контекст: tests/modTestRecipeOperations.bas:88.

Native entrypoint: modTestRecipeOperations.RunRecipeOperationTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-csectiongeometryquery-transformcoordinates"></a>
### CSectionGeometryQuery.TransformCoordinates

Преобразует скалярную точку: reflectX, rotation, shift; общий порядок совпадает с путями.

```vb
Public Sub TransformCoordinates(ByRef x As Double, ByRef y As Double, ByVal shiftX As Double, ByVal shiftY As Double, Optional ByVal rotation As Double = 0#, Optional ByVal reflectX As Boolean = False)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| x | Double; ByRef; обязательный | Модельная координата x, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| y | Double; ByRef; обязательный | Модельная координата y, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| shiftX | Double; ByVal; обязательный | Перенос по X, мм. |
| shiftY | Double; ByVal; обязательный | Перенос по Y, мм. |
| rotation | Double; ByVal; Optional; default=0# | Угол поворота вокруг (0,0), радианы, CCW положителен. |
| reflectX | Boolean; ByVal; Optional; default=False | Отражение относительно локальной оси X перед поворотом и переносом. |

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:37

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    x = 1#: y = 2#: query.TransformCoordinates x, y, 10#, 20#, GEOM_PI / 2#, True
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:135.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-csectiongeometryquery-offsetpath"></a>
### CSectionGeometryQuery.OffsetPath

Смещает линию/аналитическую цепочку слева от обхода при положительном signedLeftDistance. Line/line стыки miter; arc стыки должны быть гладкими.

```vb
Public Function OffsetPath(ByVal pathID As String, ByVal source As CGeometryPath, ByVal distance As Double) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| source | CGeometryPath; ByVal; обязательный | Подготовленный объект/контейнер выбранного запроса, без повторного чтения Excel или назначения типа. |
| distance | Double; ByVal; обязательный | Расстояние или граница интервала distance, внутренние мм; направление и допустимые пределы определяет вызываемый метод. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:49

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.InvalidOperation с его подготовленными локальными объектами/входом.

```vb
            Set result = query.OffsetPath("Failed", path, 1#)
```

Полный контекст: tests/modTestRecipeOperations.bas:83.

Native entrypoint: modTestRecipeOperations.RunRecipeOperationTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-csectiongeometryquery-constrainline"></a>
### CSectionGeometryQuery.ConstrainLine

Ограничивает или продлевает одну прямую двумя поперечными плоскостями nX*x+nY*y=bound; направление исходного пути сохраняется.

```vb
Public Function ConstrainLine(ByVal pathID As String, ByVal source As CGeometryPath, ByVal startNX As Double, ByVal startNY As Double, ByVal startBound As Double, ByVal endNX As Double, ByVal endNY As Double, ByVal endBound As Double) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| source | CGeometryPath; ByVal; обязательный | Подготовленный объект/контейнер выбранного запроса, без повторного чтения Excel или назначения типа. |
| startNX | Double; ByVal; обязательный | X-компонента начальной поперечной нормали; линия не должна быть параллельна плоскости. |
| startNY | Double; ByVal; обязательный | Y-компонента начальной поперечной нормали. |
| startBound | Double; ByVal; обязательный | Значение начальной плоскости startNX*x+startNY*y=startBound; мм при единичной нормали. |
| endNX | Double; ByVal; обязательный | X-компонента конечной поперечной нормали. |
| endNY | Double; ByVal; обязательный | Y-компонента конечной поперечной нормали. |
| endBound | Double; ByVal; обязательный | Значение конечной плоскости endNX*x+endNY*y=endBound; мм при единичной нормали. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:119

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.InvalidOperation с его подготовленными локальными объектами/входом.

```vb
            Set result = query.ConstrainLine("Failed", path, 1#, 0#, 0#, 1#, 0#, 10#)
```

Полный контекст: tests/modTestRecipeOperations.bas:79.

Native entrypoint: modTestRecipeOperations.RunRecipeOperationTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-csectiongeometryquery-initialize"></a>
### CSectionGeometryQuery.Initialize

Привязывает query к CSectionModel/revision. Достоверные контуры приоритетны; иначе используются реальные границы КЭ.

```vb
Public Sub Initialize(ByVal section As CSectionModel)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| section | CSectionModel; ByVal; обязательный | CSectionModel на внешней границе query/build; ядро не читает Excel. |

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:153

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestMesh с его подготовленными локальными объектами/входом.

```vb
    query.Initialize model
```

Полный контекст: tests/modTestGeometryQuery.bas:257.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-csectiongeometryquery-concretedomain"></a>
### CSectionGeometryQuery.ConcreteDomain

Возвращает clone области текущей query-session; не подменяет A/I механической сетки.

```vb
Public Function ConcreteDomain() As CConcreteRegion
```

Тип результата: CConcreteRegion. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:180

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestMesh с его подготовленными локальными объектами/входом.

```vb
    Set region = query.ConcreteDomain
```

Полный контекст: tests/modTestGeometryQuery.bas:258.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-csectiongeometryquery-projectionbounds"></a>
### CSectionGeometryQuery.ProjectionBounds

Находит min/max аналитической области в заданном направлении, включая внутренние экстремумы дуг.

```vb
Public Sub ProjectionBounds(ByVal region As CConcreteRegion, ByVal nx As Double, ByVal ny As Double, ByRef minimum As Double, ByRef maximum As Double)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| region | CConcreteRegion; ByVal; обязательный | Готовый CConcreteRegion; Nothing не является пустой областью. |
| nx | Double; ByVal; обязательный | Компонента нормали/направления в section XY; вектор не нулевой. |
| ny | Double; ByVal; обязательный | Компонента нормали/направления в section XY; вектор не нулевой. |
| minimum | Double; ByRef; обязательный | Расстояние или граница интервала minimum, внутренние мм; направление и допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| maximum | Double; ByRef; обязательный | Расстояние или граница интервала maximum, внутренние мм; направление и допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:188

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestRotatedAndTouchingGeometry с его подготовленными локальными объектами/входом.

```vb
    query.ProjectionBounds region, Cos(angle), Sin(angle), minimum, maximum
```

Полный контекст: tests/modTestGeometryQuery.bas:726.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-csectiongeometryquery-openingbounds"></a>
### CSectionGeometryQuery.OpeningBounds

Возвращает габариты всех заданных отверстий из контуров, без ограничителя только первым отверстием.

```vb
Public Sub OpeningBounds(ByVal contours As CSectionContours, ByRef bounds() As Double, ByRef openingCount As Long)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| contours | CSectionContours; ByVal; обязательный | CSectionContours с достоверными аналитическими Outer/Opening. |
| bounds | Double(); ByRef; обязательный | Выходной массив габаритов, мм; формат определяется конкретным запросом. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| openingCount | Long; ByRef; обязательный | Целый неотрицательный счётчик. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:207

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestGeneratedContoursWithoutAnnotations с его подготовленными локальными объектами/входом.

```vb
        query.OpeningBounds model.Contours, bounds, openingCount
```

Полный контекст: tests/modTestGeometryQuery.bas:808.

Native entrypoint: modTestGeometryQuery.RunGeneratedContourOwnershipTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-csectiongeometryquery-polygonregion"></a>
### CSectionGeometryQuery.PolygonRegion

Готовит материальную область по замкнутому многоугольнику и проверяет топологию.

```vb
Public Function PolygonRegion(ByVal coordinates As Variant) As CConcreteRegion
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| coordinates | Variant; ByVal; обязательный | Массив координат вершин замкнутого многоугольника, мм. |

Тип результата: CConcreteRegion. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:233

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestRotatedAndTouchingGeometry с его подготовленными локальными объектами/входом.

```vb
    Set region = query.PolygonRegion(points)
```

Полный контекст: tests/modTestGeometryQuery.bas:719.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-csectiongeometryquery-circleregion"></a>
### CSectionGeometryQuery.CircleRegion

Создаёт точную круговую область по центру и положительному R.

```vb
Public Function CircleRegion(ByVal centerX As Double, ByVal centerY As Double, ByVal radius As Double) As CConcreteRegion
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| centerX | Double; ByVal; обязательный | Модельная координата centerX, мм; знаки соответствуют section XY/WCS. |
| centerY | Double; ByVal; обязательный | Модельная координата centerY, мм; знаки соответствуют section XY/WCS. |
| radius | Double; ByVal; обязательный | Положительный радиус, мм. |

Тип результата: CConcreteRegion. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:248

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestOpenings с его подготовленными локальными объектами/входом.

```vb
    Set hole = query.CircleRegion(40#, 50#, 10#)
```

Полный контекст: tests/modTestGeometryQuery.bas:211.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-csectiongeometryquery-boundaryregion"></a>
### CSectionGeometryQuery.BoundaryRegion

Готовит область по аналитическим рёбрам, сшивает/проверяет замкнутые loops и роли.

```vb
Public Function BoundaryRegion(ByRef edges() As TRegionEdge, ByVal count As Long, Optional ByVal source As String = "ValidatedContour") As CConcreteRegion
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| edges | TRegionEdge(); ByRef; обязательный | Массив TRegionEdge аналитических рёбер с loop-данными, мм/радианы. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| count | Long; ByVal; обязательный | Целый неотрицательный счётчик. |
| source | String; ByVal; Optional; default="ValidatedContour" | Подготовленный объект/контейнер выбранного запроса, без повторного чтения Excel или назначения типа. |

Тип результата: CConcreteRegion. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:259

Минимальный вызов; требуемый контекст: В теле modTestAutoCADContours.CheckImportedContourDimensions с его подготовленными локальными объектами/входом.

```vb
            Set region = query.BoundaryRegion(raw, count)
```

Полный контекст: tests/modTestAutoCADContours.bas:514.

Native entrypoint: modTestAutoCADContours.RunAutoCADContourTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalAcceptance/FakeCAD.txt.

<a id="api-csectiongeometryquery-loopisopening"></a>
### CSectionGeometryQuery.LoopIsOpening

Классифицирует подготовленную loop геометрически, не по строке имени.

```vb
Public Function LoopIsOpening(ByVal region As CConcreteRegion, ByVal loopID As Long) As Boolean
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| region | CConcreteRegion; ByVal; обязательный | Готовый CConcreteRegion; Nothing не является пустой областью. |
| loopID | Long; ByVal; обязательный | Непустой устойчивый ID контура; у CConcreteRegion численный индекс loop 1..LoopCount. |

Тип результата: Boolean. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:343

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    Check "query.loopRole.opening", query.LoopIsOpening(frame, 2)
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:127.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-csectiongeometryquery-connectedregionatpoint"></a>
### CSectionGeometryQuery.ConnectedRegionAtPoint

Возвращает связанную компоненту материальной области по точке.

```vb
Public Function ConnectedRegionAtPoint(ByVal region As CConcreteRegion, ByVal x As Double, ByVal y As Double) As CConcreteRegion
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| region | CConcreteRegion; ByVal; обязательный | Готовый CConcreteRegion; Nothing не является пустой областью. |
| x | Double; ByVal; обязательный | Модельная координата x, мм; знаки соответствуют section XY/WCS. |
| y | Double; ByVal; обязательный | Модельная координата y, мм; знаки соответствуют section XY/WCS. |

Тип результата: CConcreteRegion. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:355

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestConnectedRegionSelection с его подготовленными локальными объектами/входом.

```vb
    Set selected = query.ConnectedRegionAtPoint(region, 10#, 50#)
```

Полный контекст: tests/modTestGeometryQuery.bas:405.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-csectiongeometryquery-mergedisjointregions"></a>
### CSectionGeometryQuery.MergeDisjointRegions

Объединяет непересекающиеся области; не является универсальным Union перекрывающихся фигур.

```vb
Public Function MergeDisjointRegions(ByVal regions As Collection) As CConcreteRegion
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| regions | Collection; ByVal; обязательный | Collection готовых непересекающихся областей для MergeDisjointRegions. |

Тип результата: CConcreteRegion. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:411

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestPostAudit03RegionClone с его подготовленными локальными объектами/входом.

```vb
    Set region = query.MergeDisjointRegions(regions)
```

Полный контекст: tests/modTestGeometryQuery.bas:355.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-csectiongeometryquery-subtractcontainedregion"></a>
### CSectionGeometryQuery.SubtractContainedRegion

Вычитает строго вложенное отверстие, отклоняя нарушенные отношения границ.

```vb
Public Function SubtractContainedRegion(ByVal outer As CConcreteRegion, ByVal opening As CConcreteRegion) As CConcreteRegion
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| outer | CConcreteRegion; ByVal; обязательный | Исходная область материала/наружная область. |
| opening | CConcreteRegion; ByVal; обязательный | Вычитаемая область отверстия. |

Тип результата: CConcreteRegion. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:432

Минимальный вызов; требуемый контекст: В теле modTestContourMesh.CheckArbitraryDomains с его подготовленными локальными объектами/входом.

```vb
    Set region = query.SubtractContainedRegion(region, query.CircleRegion(70#, 70#, 23#))
```

Полный контекст: tests/modTestContourMesh.bas:523.

Native entrypoint: modTestContourMesh.RunContourMeshTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalAcceptance/FakeCAD.txt.

<a id="api-csectiongeometryquery-subtractregion"></a>
### CSectionGeometryQuery.SubtractRegion

Вычитает область по действующему ограниченному контурному контракту query; проверяет допустимость результата.

```vb
Public Function SubtractRegion(ByVal outer As CConcreteRegion, ByVal opening As CConcreteRegion) As CConcreteRegion
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| outer | CConcreteRegion; ByVal; обязательный | Исходная область материала/наружная область. |
| opening | CConcreteRegion; ByVal; обязательный | Вычитаемая область отверстия. |

Тип результата: CConcreteRegion. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:516

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestIntervals с его подготовленными локальными объектами/входом.

```vb
    Set region = query.SubtractRegion(region, hole)
```

Полный контекст: tests/modTestGeometryQuery.bas:230.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-csectiongeometryquery-regionscoincide"></a>
### CSectionGeometryQuery.RegionsCoincide

Проверяет совпадение подготовленных областей по геометрическому контракту.

```vb
Public Function RegionsCoincide(ByVal first As CConcreteRegion, ByVal second As CConcreteRegion) As Boolean
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| first | CConcreteRegion; ByVal; обязательный | Первая/вторая сравниваемая готовая область. |
| second | CConcreteRegion; ByVal; обязательный | Первая/вторая сравниваемая готовая область. |

Тип результата: Boolean. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:535

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestAnalyticClipping с его подготовленными локальными объектами/входом.

```vb
    Check "regionEquality.same", query.RegionsCoincide(region, copy)
```

Полный контекст: tests/modTestGeometryQuery.bas:120.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-csectiongeometryquery-cliphalfplane"></a>
### CSectionGeometryQuery.ClipHalfPlane

Обрезает материальную область полуплоскостью, сохраняя аналитические рёбра и топологию.

```vb
Public Function ClipHalfPlane(ByVal region As CConcreteRegion, ByVal nx As Double, ByVal ny As Double, ByVal bound As Double) As CConcreteRegion
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| region | CConcreteRegion; ByVal; обязательный | Готовый CConcreteRegion; Nothing не является пустой областью. |
| nx | Double; ByVal; обязательный | Компонента нормали/направления в section XY; вектор не нулевой. |
| ny | Double; ByVal; обязательный | Компонента нормали/направления в section XY; вектор не нулевой. |
| bound | Double; ByVal; обязательный | Расстояние или граница интервала bound, внутренние мм; направление и допустимые пределы определяет вызываемый метод. |

Тип результата: CConcreteRegion. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:553

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestPreparedGeometry с его подготовленными локальными объектами/входом.

```vb
    Set region = query.ClipHalfPlane(region, 1#, 0#, 31#)
```

Полный контекст: tests/modTestGeometryQuery.bas:182.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-csectiongeometryquery-clipwindow"></a>
### CSectionGeometryQuery.ClipWindow

Обрезает область конечным окном в направленных n/t координатах.

```vb
Public Function ClipWindow(ByVal region As CConcreteRegion, ByVal nx As Double, ByVal ny As Double, ByVal nMin As Double, ByVal nMax As Double, ByVal tMin As Double, ByVal tMax As Double) As CConcreteRegion
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| region | CConcreteRegion; ByVal; обязательный | Готовый CConcreteRegion; Nothing не является пустой областью. |
| nx | Double; ByVal; обязательный | Компонента нормали/направления в section XY; вектор не нулевой. |
| ny | Double; ByVal; обязательный | Компонента нормали/направления в section XY; вектор не нулевой. |
| nMin | Double; ByVal; обязательный | Расстояние или граница интервала nMin, внутренние мм; направление и допустимые пределы определяет вызываемый метод. |
| nMax | Double; ByVal; обязательный | Расстояние или граница интервала nMax, внутренние мм; направление и допустимые пределы определяет вызываемый метод. |
| tMin | Double; ByVal; обязательный | Расстояние или граница интервала tMin, внутренние мм; направление и допустимые пределы определяет вызываемый метод. |
| tMax | Double; ByVal; обязательный | Расстояние или граница интервала tMax, внутренние мм; направление и допустимые пределы определяет вызываемый метод. |

Тип результата: CConcreteRegion. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:625

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestFirstMaterialWall с его подготовленными локальными объектами/входом.

```vb
    Set wall = query.ClipWindow(wall, 0#, 1#, 0#, 100#, -50#, -30#)
```

Полный контекст: tests/modTestGeometryQuery.bas:571.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-csectiongeometryquery-materialregionthroughplane"></a>
### CSectionGeometryQuery.MaterialRegionThroughPlane

Получает материальную область по заданной плоскости и окну; не нормативный расчёт трещин.

```vb
Public Function MaterialRegionThroughPlane(ByVal region As CConcreteRegion, ByVal nx As Double, ByVal ny As Double, ByVal seedPlane As Double) As CConcreteRegion
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| region | CConcreteRegion; ByVal; обязательный | Готовый CConcreteRegion; Nothing не является пустой областью. |
| nx | Double; ByVal; обязательный | Компонента нормали/направления в section XY; вектор не нулевой. |
| ny | Double; ByVal; обязательный | Компонента нормали/направления в section XY; вектор не нулевой. |
| seedPlane | Double; ByVal; обязательный | Подготовленный объект/контейнер выбранного запроса, без повторного чтения Excel или назначения типа. |

Тип результата: CConcreteRegion. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:642

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestFirstMaterialWall с его подготовленными локальными объектами/входом.

```vb
    Set wall = query.MaterialRegionThroughPlane(region, 0#, 1#, 85#)
```

Полный контекст: tests/modTestGeometryQuery.bas:565.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-csectiongeometryquery-containspoint"></a>
### CSectionGeometryQuery.ContainsPoint

Проверяет действительную принадлежность материальной области с заданной политикой границы.

```vb
Public Function ContainsPoint(ByVal region As CConcreteRegion, ByVal x As Double, ByVal y As Double, Optional ByVal includeBoundary As Boolean = True) As Boolean
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| region | CConcreteRegion; ByVal; обязательный | Готовый CConcreteRegion; Nothing не является пустой областью. |
| x | Double; ByVal; обязательный | Модельная координата x, мм; знаки соответствуют section XY/WCS. |
| y | Double; ByVal; обязательный | Модельная координата y, мм; знаки соответствуют section XY/WCS. |
| includeBoundary | Boolean; ByVal; Optional; default=True | Считать точку на границе принадлежащей области. |

Тип результата: Boolean. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:801

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.TestAdjacentRectangles с его подготовленными локальными объектами/входом.

```vb
    Check "adjacent.lower", query.ContainsPoint(region, 500#, 100#)
```

Полный контекст: tests/modTestSectionRecipes.bas:1068.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-csectiongeometryquery-boundarydistance"></a>
### CSectionGeometryQuery.BoundaryDistance

Возвращает минимальное расстояние до любой внешней/внутренней границы, мм.

```vb
Public Function BoundaryDistance(ByVal region As CConcreteRegion, ByVal x As Double, ByVal y As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| region | CConcreteRegion; ByVal; обязательный | Готовый CConcreteRegion; Nothing не является пустой областью. |
| x | Double; ByVal; обязательный | Модельная координата x, мм; знаки соответствуют section XY/WCS. |
| y | Double; ByVal; обязательный | Модельная координата y, мм; знаки соответствуют section XY/WCS. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:836

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.RunRecipeRuleTests с его подготовленными локальными объектами/входом.

```vb
    Near "distance.inside", query.BoundaryDistance(region, 90#, -20#), 20#
```

Полный контекст: tests/modTestSectionRecipes.bas:259.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-csectiongeometryquery-roundedpolygon"></a>
### CSectionGeometryQuery.RoundedPolygon

Создаёт и регистрирует согласованную границу по одному завершённому CPolygonSpec. Edge.ID начинается у исходной вершины; Corner.ID существует при R>0. Исходный описатель не меняется.

```vb
Public Function RoundedPolygon(ByVal pathID As String, ByVal polygon As CPolygonSpec) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| polygon | CPolygonSpec; ByVal; обязательный | Завершённый CPolygonSpec исходных вершин до скругления. Nothing и не закрытый через Seal набор отклоняются; порядок, радиусы и источники сохраняются. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:868

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02PolygonApiTests с его подготовленными локальными объектами/входом.

```vb
    Set path = query.RoundedPolygon("Base", polygon)
```

Полный контекст: tests/modTestAudit02Api.bas:27.

Native entrypoint: modTestAudit02Api.RunAudit02PolygonApiTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02PolygonApiTests.txt.

<a id="api-csectiongeometryquery-offsetroundedpolygon"></a>
### CSectionGeometryQuery.OffsetRoundedPolygon

Создаёт и регистрирует смещённый скруглённый многоугольник по исходному CPolygonSpec, не меняя его. Положительное расстояние внутрь площади; отрицательное наружу, из Opening в бетон. Сохраняет привязки исчезнувших участков; отличается от signed-left Offset готового пути.

```vb
Public Function OffsetRoundedPolygon(ByVal pathID As String, ByVal polygon As CPolygonSpec, ByVal distance As Double) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| polygon | CPolygonSpec; ByVal; обязательный | Завершённый CPolygonSpec исходных вершин до скругления. Nothing и не закрытый через Seal набор отклоняются; порядок, радиусы и источники сохраняются. |
| distance | Double; ByVal; обязательный | Signed offset, мм: положительно внутрь площади, отрицательно наружу; это не left-offset обхода. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:952

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.RunRoundedOffsetTests с его подготовленными локальными объектами/входом.

```vb
    Set path = query.OffsetRoundedPolygon("Far", polygonSpec, 5#)
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:395.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-csectiongeometryquery-regionfrompath"></a>
### CSectionGeometryQuery.RegionFromPath

Готовит область по одному замкнутому аналитическому пути; открытый путь не становится бетоном.

```vb
Public Function RegionFromPath(ByVal path As CGeometryPath) As CConcreteRegion
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| path | CGeometryPath; ByVal; обязательный | Существующий непрерывный опубликованный CGeometryPath. |

Тип результата: CConcreteRegion. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:1026

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02PolygonApiTests с его подготовленными локальными объектами/входом.

```vb
    Set region = query.RegionFromPath(path)
```

Полный контекст: tests/modTestAudit02Api.bas:32.

Native entrypoint: modTestAudit02Api.RunAudit02PolygonApiTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02PolygonApiTests.txt.

<a id="api-csectiongeometryquery-adjacentrectangles"></a>
### CSectionGeometryQuery.AdjacentRectangles

Создаёт одну границу смежных неперекрывающихся прямоугольников. Произвольное пересекающееся CAD Union не выполняется.

```vb
Public Function AdjacentRectangles(ByVal pathID As String, ByVal rectangles As Variant) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| rectangles | Variant; ByVal; обязательный | Array(Array(x,y,width,height),...); мм, размеры положительные, прямоугольники смежные без перекрытий. |

Тип результата: CGeometryPath. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:1038

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.InvalidAdjacentRejected с его подготовленными локальными объектами/входом.

```vb
    Set path = query.AdjacentRectangles("Bad", Array(Array(0#, 0#, 100#, 100#), rect))
```

Полный контекст: tests/modTestSectionRecipes.bas:1089.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-csectiongeometryquery-lineintervals"></a>
### CSectionGeometryQuery.LineIntervals

Возвращает реальные материальные интервалы пересечения с линией, с отверстиями и всеми компонентами.

```vb
Public Function LineIntervals(ByVal region As CConcreteRegion, ByVal pointX As Double, ByVal pointY As Double, ByVal directionX As Double, ByVal directionY As Double, Optional ByVal parameterMin As Double = -1E+100, Optional ByVal parameterMax As Double = 1E+100) As Variant
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| region | CConcreteRegion; ByVal; обязательный | Готовый CConcreteRegion; Nothing не является пустой областью. |
| pointX | Double; ByVal; обязательный | Модельная координата pointX, мм; знаки соответствуют section XY/WCS. |
| pointY | Double; ByVal; обязательный | Модельная координата pointY, мм; знаки соответствуют section XY/WCS. |
| directionX | Double; ByVal; обязательный | X-компонента направления проекции/выбора экстремума; вектор не нулевой, у RebarRowExtremeLabel единичный. |
| directionY | Double; ByVal; обязательный | Y-компонента направления проекции/выбора экстремума. |
| parameterMin | Double; ByVal; Optional; default=-1E+100 | Граница нормированного параметра выбранной линии/интервала. |
| parameterMax | Double; ByVal; Optional; default=1E+100 | Граница нормированного параметра выбранной линии/интервала. |

Тип результата: Variant. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:1209

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestIntervals с его подготовленными локальными объектами/входом.

```vb
    intervals = query.LineIntervals(region, 10#, 50#, 1#, 0#)
```

Полный контекст: tests/modTestGeometryQuery.bas:231.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-csectiongeometryquery-samplematerialintervals"></a>
### CSectionGeometryQuery.SampleMaterialIntervals

Получает материальные интервалы для заданных направлений/уровней; не восстанавливает область по одной площади.

```vb
Public Function SampleMaterialIntervals(ByVal region As CConcreteRegion, ByVal nx As Double, ByVal ny As Double, ByVal seedPlane As Double, ByVal tMin As Double, ByVal tMax As Double, ByVal probeCount As Long) As Variant
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| region | CConcreteRegion; ByVal; обязательный | Готовый CConcreteRegion; Nothing не является пустой областью. |
| nx | Double; ByVal; обязательный | Компонента нормали/направления в section XY; вектор не нулевой. |
| ny | Double; ByVal; обязательный | Компонента нормали/направления в section XY; вектор не нулевой. |
| seedPlane | Double; ByVal; обязательный | Signed нормальная координата опорной плоскости, мм, относительно начала section XY; это скаляр Double, не объект плоскости. |
| tMin | Double; ByVal; обязательный | Расстояние или граница интервала tMin, внутренние мм; направление и допустимые пределы определяет вызываемый метод. |
| tMax | Double; ByVal; обязательный | Расстояние или граница интервала tMax, внутренние мм; направление и допустимые пределы определяет вызываемый метод. |
| probeCount | Long; ByVal; обязательный | Целый неотрицательный счётчик. |

Тип результата: Variant. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:1270

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    probes = query.SampleMaterialIntervals(region, 1#, 0#, 0#, -10#, 10#, 3)
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:129.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-csectiongeometryquery-concretecoverfrompointalongdirection"></a>
### CSectionGeometryQuery.ConcreteCoverFromPointAlongDirection

В текущей query-session находит геометрическое расстояние до бетонной границы по заданному направлению.

```vb
Public Function ConcreteCoverFromPointAlongDirection(ByVal pointX As Double, ByVal pointY As Double, ByVal directionX As Double, ByVal directionY As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pointX | Double; ByVal; обязательный | Модельная координата pointX, мм; знаки соответствуют section XY/WCS. |
| pointY | Double; ByVal; обязательный | Модельная координата pointY, мм; знаки соответствуют section XY/WCS. |
| directionX | Double; ByVal; обязательный | X-компонента направления проекции/выбора экстремума; вектор не нулевой, у RebarRowExtremeLabel единичный. |
| directionY | Double; ByVal; обязательный | Y-компонента направления проекции/выбора экстремума. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:1302

Минимальный вызов; требуемый контекст: В теле modTestGeometry.TestConcreteCoverUsesLocalContour с его подготовленными локальными объектами/входом.

```vb
    localCover = query.ConcreteCoverFromPointAlongDirection(0#, 60#, 0#, 1#)
```

Полный контекст: tests/modTestGeometry.bas:1198.

Native entrypoint: modTestGeometry.RunGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-csectiongeometryquery-regionfromcontours"></a>
### CSectionGeometryQuery.RegionFromContours

Готовит область из единственных достоверных CSectionContours Outer/Opening.

```vb
Public Function RegionFromContours(ByVal contours As CSectionContours) As CConcreteRegion
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| contours | CSectionContours; ByVal; обязательный | CSectionContours с достоверными аналитическими Outer/Opening. |

Тип результата: CConcreteRegion. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:1441

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestPreparedGeometry с его подготовленными локальными объектами/входом.

```vb
    Set region = query.RegionFromContours(contours)
```

Полный контекст: tests/modTestGeometryQuery.bas:163.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

<a id="api-csectiongeometryquery-regioncentroid"></a>
### CSectionGeometryQuery.RegionCentroid

Вычисляет аналитический центр материальной области, мм; не заменяет центр механической сетки НДС.

```vb
Public Sub RegionCentroid(ByVal region As CConcreteRegion, ByRef x As Double, ByRef y As Double)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| region | CConcreteRegion; ByVal; обязательный | Готовый CConcreteRegion; Nothing не является пустой областью. |
| x | Double; ByRef; обязательный | Модельная координата x, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| y | Double; ByRef; обязательный | Модельная координата y, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:1461

Минимальный вызов; требуемый контекст: В теле modTestGeometryQuery.TestPreparedGeometry с его подготовленными локальными объектами/входом.

```vb
    query.RegionCentroid region, x, y
```

Полный контекст: tests/modTestGeometryQuery.bas:179.

Native entrypoint: modTestGeometryQuery.RunGeometryQueryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/QueriesAndContours.txt.

## modSectionRecipeTypes

Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

<a id="api-modsectionrecipetypes-esectionlooprole"></a>
### modSectionRecipeTypes.ESectionLoopRole

Outer/Opening — геометрические роли, не выводятся из имени контура.

```vb
Public Enum ESectionLoopRole
```

Поля/селекторы по source:

```vb
LoopOuter = 1
LoopOpening = 2
```

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:6

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiInputAndDataExamples с его подготовленными локальными объектами/входом.

```vb
    loopValue.PathID = "Concrete": loopValue.Role = LoopOuter
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:231.

Native entrypoint: modTestRecipeApiCatalogue.RunApiInputAndDataExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-modsectionrecipetypes-econtournamingpolicy"></a>
### modSectionRecipeTypes.EContourNamingPolicy

PreserveSegmentIDs=0: исходные/явные ID; SequentialByKind=1: технический префикс_LINE_i/ARC_i с прежним общим номером сегмента.

```vb
Public Enum EContourNamingPolicy
```

Поля/селекторы по source:

```vb
PreserveSegmentIDs = 0
SequentialByKind = 1
```

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:12

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02PolygonApiTests с его подготовленными локальными объектами/входом.

```vb
    facade.Outer "Outer", path.ID, namingPolicy:=SequentialByKind
```

Полный контекст: tests/modTestAudit02Api.bas:71.

Native entrypoint: modTestAudit02Api.RunAudit02PolygonApiTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02PolygonApiTests.txt.

<a id="api-modsectionrecipetypes-erebaroffsetaxis"></a>
### modSectionRecipeTypes.ERebarOffsetAxis

RowNormal=1 или RowTangent=2: выбранная компонента контактного зависимого смещения; знак задаётся отдельно.

```vb
Public Enum ERebarOffsetAxis
```

Поля/селекторы по source:

```vb
RowNormal = 1
RowTangent = 2
```

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:18

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02OptionsTests с его подготовленными локальными объектами/входом.

```vb
    For Each axis In Array(RowNormal, RowTangent)
```

Полный контекст: tests/modTestAudit02Api.bas:97.

Native entrypoint: modTestAudit02Api.RunAudit02OptionsTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02OptionsTests.txt.

<a id="api-modsectionrecipetypes-erebarendmode"></a>
### modSectionRecipeTypes.ERebarEndMode

IncludeBoth/ExcludeBoth/IncludeStart/IncludeEnd/ClosedUniform; последние требуют соответствующей замкнутости.

```vb
Public Enum ERebarEndMode
```

Поля/селекторы по source:

```vb
IncludeBoth = 1
ExcludeBoth = 2
IncludeStart = 3
IncludeEnd = 4
ClosedUniform = 5
```

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:23

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.RunRecipeRuleTests с его подготовленными локальными объектами/входом.

```vb
    options.EndMode = ExcludeBoth
```

Полный контекст: tests/modTestSectionRecipes.bas:204.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-modsectionrecipetypes-erebardistribution"></a>
### modSectionRecipeTypes.ERebarDistribution

ByCount/ByMaxStep/ByFixedStep/ExplicitPositions/ProjectFrom/DependentOffset — режимы общего исполнителя.

```vb
Public Enum ERebarDistribution
```

Поля/селекторы по source:

```vb
ByCount = 1
ByMaxStep = 2
ByFixedStep = 3
ExplicitPositions = 4
ProjectFrom = 5
DependentOffset = 6
```

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:31

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiInputAndDataExamples с его подготовленными локальными объектами/входом.

```vb
    Check "raw.rule", ruleValue.RuleID = "Bars" And ruleValue.Mode = ByCount
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:236.

Native entrypoint: modTestRecipeApiCatalogue.RunApiInputAndDataExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-modsectionrecipetypes-erebardistributionscope"></a>
### modSectionRecipeTypes.ERebarDistributionScope

Допустимо только PerPath: число стержней, расстояния между ними и отступы от концов применяются отдельно к каждой независимой линии группы.

```vb
Public Enum ERebarDistributionScope
```

Поля/селекторы по source:

```vb
PerPath = 1
```

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:41

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.AdvancedGuideDefinition с его подготовленными локальными объектами/входом.

```vb
    options.DistributionScope = PerPath: options.StartSetback = 20#: options.EndSetback = 30#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:61.

Native entrypoint: modTestRecipeApiCatalogue.RunAdvancedGuideExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/AdvancedExamples.txt.

<a id="api-modsectionrecipetypes-eremainderplacement"></a>
### modSectionRecipeTypes.ERemainderPlacement

RemainderStart/End/Center: положение остатка для точного шага.

```vb
Public Enum ERemainderPlacement
```

Поля/селекторы по source:

```vb
RemainderStart = 1
RemainderEnd = 2
RemainderCenter = 3
```

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:45

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiGeometryReadExamples с его подготовленными локальными объектами/входом.

```vb
    Check "defaults.remainder", options.RemainderPlacement = RemainderCenter
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:105.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

<a id="api-modsectionrecipetypes-erebarpublication"></a>
### modSectionRecipeTypes.ERebarPublication

PublishByRule или PublishPerSourcePosition; влияет на вывод, не геометрический шаг.

```vb
Public Enum ERebarPublication
```

Поля/селекторы по source:

```vb
PublishByRule = 1
PublishPerSourcePosition = 2
```

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:51

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.TestProjectionNormalsAndOrdinals с его подготовленными локальными объектами/входом.

```vb
        options.Publication = PublishPerSourcePosition
```

Полный контекст: tests/modTestSectionRecipes.bas:302.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-modsectionrecipetypes-erebarpositionorder"></a>
### modSectionRecipeTypes.ERebarPositionOrder

Порядок вывода стержней: AlongPath — от начала линии, FromPathMiddle — от её середины. Места стержней и их исходные номера сохраняются.

```vb
Public Enum ERebarPositionOrder
```

Поля/селекторы по source:

```vb
AlongPath = 1
FromPathMiddle = 2
```

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:56

Минимальный вызов; требуемый контекст: В теле modTestAudit02Api.RunAudit02OptionsTests с его подготовленными локальными объектами/входом.

```vb
    options.PositionOrder = FromPathMiddle: options.IncomingTangentAtJoint = True
```

Полный контекст: tests/modTestAudit02Api.bas:92.

Native entrypoint: modTestAudit02Api.RunAudit02OptionsTests; журнал: docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14/Audit02ContractTests/modTestAudit02Api.RunAudit02OptionsTests.txt.

<a id="api-modsectionrecipetypes-erebarcoincidence"></a>
### modSectionRecipeTypes.ERebarCoincidence

RejectCoincidence/MergeIdentical/KeepPublishedOwner; последнее сохраняет явно выбранного опубликованного владельца.

```vb
Public Enum ERebarCoincidence
```

Поля/селекторы по source:

```vb
RejectCoincidence = 1
MergeIdentical = 2
KeepPublishedOwner = 3
```

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:61

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.TestPublicationAndCoincidence с его подготовленными локальными объектами/входом.

```vb
    options.AnnotationGroup = "Second": options.Coincidence = MergeIdentical
```

Полный контекст: tests/modTestSectionRecipes.bas:394.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-modsectionrecipetypes-erebarplacementcheck"></a>
### modSectionRecipeTypes.ERebarPlacementCheck

CheckCenter или CheckFullDisk относительно Outer/Opening; ужесточение задаётся явно.

```vb
Public Enum ERebarPlacementCheck
```

Поля/селекторы по source:

```vb
CheckCenter = 1
CheckFullDisk = 2
```

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:67

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.AdvancedGuideDefinition с его подготовленными локальными объектами/входом.

```vb
    options.PlacementCheck = CheckFullDisk: options.OverlapPolicy = RejectOverlap
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:63.

Native entrypoint: modTestRecipeApiCatalogue.RunAdvancedGuideExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/AdvancedExamples.txt.

<a id="api-modsectionrecipetypes-erebaroverlappolicy"></a>
### modSectionRecipeTypes.ERebarOverlapPolicy

ReportOverlap или RejectOverlap. Контакт допустим; нормативный минимальный зазор не вводится.

```vb
Public Enum ERebarOverlapPolicy
```

Поля/селекторы по source:

```vb
ReportOverlap = 1
RejectOverlap = 2
```

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:74

Минимальный вызов; требуемый контекст: В теле modTestRecipeOperations.RunRecipeDiagnosticTests с его подготовленными локальными объектами/входом.

```vb
        If mode = 2 Then options.OverlapPolicy = RejectOverlap
```

Полный контекст: tests/modTestRecipeOperations.bas:102.

Native entrypoint: modTestRecipeOperations.RunRecipeDiagnosticTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/FinalCandidate/CoreSmoke.txt.

<a id="api-modsectionrecipetypes-tsectionloop"></a>
### modSectionRecipeTypes.TSectionLoop

Raw скалярное описание роли/ссылок контура. Заполняет фасад Outer/Opening; автору не нужен собственный классификатор.

```vb
Public Type TSectionLoop
```

Поля/селекторы по source:

```vb
PathID As String
LoopID As String
Role As ESectionLoopRole
ParentID As String
Comment As String
ArcComment As String
ContourIDs As Variant
NamingPolicy As EContourNamingPolicy
NamingPrefix As String
```

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:79

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiInputAndDataExamples с его подготовленными локальными объектами/входом.

```vb
    Dim loopValue As TSectionLoop, ruleValue As TRebarRule, annotationValue As TSectionAnnotation, positionValue As TRebarPosition
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:230.

Native entrypoint: modTestRecipeApiCatalogue.RunApiInputAndDataExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-modsectionrecipetypes-trebarrule"></a>
### modSectionRecipeTypes.TRebarRule

Raw snapshot правила исполнителя. Фасад копирует options, source references и параметры, координаты баров вычисляет общий executor.

```vb
Public Type TRebarRule
```

Поля/селекторы по source:

```vb
RuleID As String
Target As String
Mode As ERebarDistribution
DistributionScope As ERebarDistributionScope
Enabled As Boolean
Diameter As Double
SteelClass As String
Count As Long
Step As Double
StartSetback As Double
EndSetback As Double
EndMode As ERebarEndMode
RemainderPlacement As ERemainderPlacement
SourceRuleID As String
TangentOffset As Double
NormalOffset As Double
EverySecond As Boolean
IncomingTangentAtJoint As Boolean
DirectionX As Double
DirectionY As Double
SourceNormalSign As Double
SuccessfulProjectionOrdinal As Boolean
NewPublicationRoot As Boolean
SkipMissingProjection As Boolean
Publication As ERebarPublication
PositionOrder As ERebarPositionOrder
Coincidence As ERebarCoincidence
PlacementCheck As ERebarPlacementCheck
OverlapPolicy As ERebarOverlapPolicy
SourcePrefix As String
SourceSuffix As String
AnnotationGroup As String
Comment As String
SourceReference As String
ExplicitStations As Variant
```

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:91

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiInputAndDataExamples с его подготовленными локальными объектами/входом.

```vb
    Dim loopValue As TSectionLoop, ruleValue As TRebarRule, annotationValue As TSectionAnnotation, positionValue As TRebarPosition
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:230.

Native entrypoint: modTestRecipeApiCatalogue.RunApiInputAndDataExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-modsectionrecipetypes-tsectionannotation"></a>
### modSectionRecipeTypes.TSectionAnnotation

Raw annotation rule с модельными привязками/измерением. Фасад объявляет, общий engine измеряет, layout оформляет.

```vb
Public Type TSectionAnnotation
```

Поля/селекторы по source:

```vb
AnnotationID As String
Kind As String
Target As String
SegmentID As String
StartAnchorID As String
EndAnchorID As String
Name As String
StartX As Double
StartY As Double
EndX As Double
EndY As Double
NormalX As Double
NormalY As Double
DirectionX As Double
DirectionY As Double
Value As Double
DisplayIncrement As Double
ReverseDirection As Boolean
Comment As String
```

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:129

Минимальный вызов; требуемый контекст: В теле modTestSectionRecipes.TestAnnotationOwnership с его подготовленными локальными объектами/входом.

```vb
    Dim anchors(1 To 4) As TSectionAnnotation
```

Полный контекст: tests/modTestSectionRecipes.bas:441.

Native entrypoint: modTestSectionRecipes.RunRecipeRuleTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-modsectionrecipetypes-tsectionmeshoptions"></a>
### modSectionRecipeTypes.TSectionMeshOptions

Настройки общей дискретизации: StepX/StepY в мм>0, subdivisions≥1, MaterialID≥1. Для книги читаются общим workflow.

```vb
Public Type TSectionMeshOptions
```

Поля/селекторы по source:

```vb
StepX As Double
StepY As Double
BoundarySubdivisions As Long
MaterialID As Long
```

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:151

Минимальный вызов; требуемый контекст: В теле modTestRecipeAnnotations.BuildDefinition с его подготовленными локальными объектами/входом.

```vb
    Dim mesh As TSectionMeshOptions
```

Полный контекст: tests/modTestRecipeAnnotations.bas:178.

Native entrypoint: modTestRecipeAnnotations.RunRecipeAnnotationTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/DeclarativeApi/FinalApiAnnotations.txt.

Native entrypoint: modTestRecipeAnnotations.RunRecipeAnnotationSnapshotTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/DeclarativeApi/Operations.txt.

<a id="api-modsectionrecipetypes-trebarposition"></a>
### modSectionRecipeTypes.TRebarPosition

Данные места стержня до объединения совпадающих стержней: координаты в мм, касательная, расстояние от начала линии, номера внутри правила и исходного ряда, уникальные имена.

```vb
Public Type TRebarPosition
```

Поля/селекторы по source:

```vb
X As Double
Y As Double
TangentX As Double
TangentY As Double
Station As Double
RuleIndex As Long
RootRuleIndex As Long
RootOrdinal As Long
RootPublicationRank As Long
SourceOrdinal As Long
SourcePositionID As String
PathID As String
```

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:158

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiInputAndDataExamples с его подготовленными локальными объектами/входом.

```vb
    Dim loopValue As TSectionLoop, ruleValue As TRebarRule, annotationValue As TSectionAnnotation, positionValue As TRebarPosition
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:230.

Native entrypoint: modTestRecipeApiCatalogue.RunApiInputAndDataExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-modsectionrecipetypes-sectionrecipeerror"></a>
### modSectionRecipeTypes.SectionRecipeError

Выдаёт общий код 5530 с owner/причиной/строкой источника; не назначает нормативный статус.

```vb
Public Sub SectionRecipeError(ByVal owner As String, ByVal reason As String, Optional ByVal sourceReference As String = "")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| owner | String; ByVal; обязательный | Имя компонента/владельца ошибки, а не тип сечения для ветвления. |
| reason | String; ByVal; обязательный | Понятная причина отказа и исправление; машинный код не извлекается из текста. |
| sourceReference | String; ByVal; Optional; default="" | Строка исходного адреса/программного параметра; без объекта Excel. |

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:174

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiInputAndDataExamples с его подготовленными локальными объектами/входом.

```vb
    SectionRecipeError "ApiGuide", "Invalid count", "ProgramInputs.Count"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:248.

Native entrypoint: modTestRecipeApiCatalogue.RunApiInputAndDataExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-modsectionrecipetypes-sectionrecipegeometryerror"></a>
### modSectionRecipeTypes.SectionRecipeGeometryError

В обработчике добавляет SourceReference к query/path отказу; неожиданные ошибки сохраняют исходный Err.Number/Source.

```vb
Public Sub SectionRecipeGeometryError(ByVal owner As String, ByVal sourceReference As String)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| owner | String; ByVal; обязательный | Имя компонента/владельца ошибки, а не тип сечения для ветвления. |
| sourceReference | String; ByVal; обязательный | Строка исходного адреса/программного параметра; без объекта Excel. |

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:187

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiInputAndDataExamples с его подготовленными локальными объектами/входом.

```vb
    SectionRecipeGeometryError "ApiGuide", "ProgramInputs.Arc"
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:251.

Native entrypoint: modTestRecipeApiCatalogue.RunApiInputAndDataExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

## modGeometryTypes

Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

<a id="api-modgeometrytypes-geom_pi"></a>
### modGeometryTypes.GEOM_PI

Число pi как сумма сохраняемых VBE литералов, без изменения Double при компиляции.

```vb
Public Const GEOM_PI As Double = 3.14159265358979 + 3.10862446895044E-15
```

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modGeometryTypes.bas:14

Минимальный вызов; требуемый контекст: В теле modTestGeometry.TestAutoCADImporterRotatedRectangleBounds с его подготовленными локальными объектами/входом.

```vb
    angle = GEOM_PI / 4#
```

Полный контекст: tests/modTestGeometry.bas:371.

Native entrypoint: modTestGeometry.RunGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-modgeometrytypes-geom_tolerance"></a>
### modGeometryTypes.GEOM_TOLERANCE

Геометрический допуск совпадений исходных координат, мм; не округление ключей результатов.

```vb
Public Const GEOM_TOLERANCE As Double = 0.000000001
```

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modGeometryTypes.bas:15

Минимальный вызов; требуемый контекст: В теле modTestGeometry.AssertRelative с его подготовленными локальными объектами/входом.

```vb
    If Abs(expected) <= GEOM_TOLERANCE Then
```

Полный контекст: tests/modTestGeometry.bas:2572.

Native entrypoint: modTestGeometry.RunGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

<a id="api-modgeometrytypes-geom_max_rebar_count"></a>
### modGeometryTypes.GEOM_MAX_REBAR_COUNT

Предел представимости счётчика позиций, не обещание доступности памяти для такого массива.

```vb
Public Const GEOM_MAX_REBAR_COUNT As Double = 2147483646#
```

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modGeometryTypes.bas:16

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiInputAndDataExamples с его подготовленными локальными объектами/входом.

```vb
    Check "limit.counter", GEOM_MAX_REBAR_COUNT > 2147483645# And GEOM_MAX_REBAR_COUNT < 2147483647#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:244.

Native entrypoint: modTestRecipeApiCatalogue.RunApiInputAndDataExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-modgeometrytypes-tpoint2d"></a>
### modGeometryTypes.TPoint2D

Два координатных скаляра X/Y, мм, без Excel или объекта на каждую точку.

```vb
Public Type TPoint2D
```

Поля/селекторы по source:

```vb
X As Double
Y As Double
```

Жизненный цикл: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modGeometryTypes.bas:18

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiInputAndDataExamples с его подготовленными локальными объектами/входом.

```vb
    Dim point As TPoint2D, fiber As TFiber, properties As TGeometryProperties
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:239.

Native entrypoint: modTestRecipeApiCatalogue.RunApiInputAndDataExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-modgeometrytypes-tregionedge"></a>
### modGeometryTypes.TRegionEdge

Ориентированное аналитическое ребро: endpoints/center мм, Sweep радианы, 0 означает прямую, LoopID — численный loop.

```vb
Public Type TRegionEdge
```

Поля/селекторы по source:

```vb
X1 As Double
Y1 As Double
X2 As Double
Y2 As Double
CenterX As Double
CenterY As Double
Sweep As Double
LoopID As Long
```

Жизненный цикл: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modGeometryTypes.bas:26

Минимальный вызов; требуемый контекст: В теле modTestDeclarativeGeometry.RunRoundedPolygonTests с его подготовленными локальными объектами/входом.

```vb
    Dim edge As TRegionEdge, i As Long, radius As Double, scenario As Long
```

Полный контекст: tests/modTestDeclarativeGeometry.bas:179.

Native entrypoint: modTestDeclarativeGeometry.RunCoreGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

Native entrypoint: modTestDeclarativeGeometry.RunRoundedPolygonTests; журнал: docs/regression/GeometryDeclarative_2026-10-11/CorePaths/RoundedPolygons.txt.

<a id="api-modgeometrytypes-tfiber"></a>
### modGeometryTypes.TFiber

Скалярные данные механического бетонного КЭ. Не заменяются аналитической площадью области.

```vb
Public Type TFiber
```

Поля/селекторы по source:

```vb
X As Double
Y As Double
Width As Double
Height As Double
Area As Double
FillFactor As Double
MaterialID As Long
```

Жизненный цикл: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modGeometryTypes.bas:37

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiInputAndDataExamples с его подготовленными локальными объектами/входом.

```vb
    Dim point As TPoint2D, fiber As TFiber, properties As TGeometryProperties
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:239.

Native entrypoint: modTestRecipeApiCatalogue.RunApiInputAndDataExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-modgeometrytypes-tgeometryproperties"></a>
### modGeometryTypes.TGeometryProperties

Скалярные геометрические свойства во внутренних единицах; источник аналитический/сеточный определяется вызывающим алгоритмом.

```vb
Public Type TGeometryProperties
```

Поля/селекторы по source:

```vb
Area As Double
StaticMomentX As Double
StaticMomentY As Double
CentroidX As Double
CentroidY As Double
Ix As Double
Iy As Double
Ixy As Double
Ixc As Double
Iyc As Double
Ixyc As Double
PrincipalI1 As Double
PrincipalI2 As Double
PrincipalAngleRad As Double
RadiusX As Double
RadiusY As Double
PrincipalRadius1 As Double
PrincipalRadius2 As Double
```

Жизненный цикл: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modGeometryTypes.bas:47

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiInputAndDataExamples с его подготовленными локальными объектами/входом.

```vb
    Dim point As TPoint2D, fiber As TFiber, properties As TGeometryProperties
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:239.

Native entrypoint: modTestRecipeApiCatalogue.RunApiInputAndDataExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-modgeometrytypes-geommax"></a>
### modGeometryTypes.GeomMax

Максимум двух чисел одной размерности, без преобразования единиц.

```vb
Public Function GeomMax(ByVal A As Double, ByVal B As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| A | Double; ByVal; обязательный | Число той же размерности, что второй аргумент; helper единицы не преобразует. |
| B | Double; ByVal; обязательный | Число той же размерности, что второй аргумент; helper единицы не преобразует. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modGeometryTypes.bas:70

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.Near с его подготовленными локальными объектами/входом.

```vb
    Check name, Abs(actual - expected) <= 0.000001 * GeomMax(1#, Abs(expected))
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:262.

Native entrypoint: modTestRecipeApiCatalogue.RunAdvancedGuideExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/AdvancedExamples.txt.

Native entrypoint: modTestRecipeApiCatalogue.RunApiGeometryReadExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.unit-label-probe.txt.

Native entrypoint: modTestRecipeApiCatalogue.RunApiUnitExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

Native entrypoint: modTestRecipeApiCatalogue.RunApiInputAndDataExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-modgeometrytypes-geommin"></a>
### modGeometryTypes.GeomMin

Минимум двух чисел одной размерности.

```vb
Public Function GeomMin(ByVal A As Double, ByVal B As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| A | Double; ByVal; обязательный | Число той же размерности, что второй аргумент; helper единицы не преобразует. |
| B | Double; ByVal; обязательный | Число той же размерности, что второй аргумент; helper единицы не преобразует. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modGeometryTypes.bas:80

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiInputAndDataExamples с его подготовленными локальными объектами/входом.

```vb
    Near "helper.min", GeomMin(1#, 2#), 1#
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:245.

Native entrypoint: modTestRecipeApiCatalogue.RunApiInputAndDataExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-modgeometrytypes-readcontourarcsweep"></a>
### modGeometryTypes.ReadContourArcSweep

Читает согласованный signed sweep точной контурной дуги, не выбирая произвольные несовместимые входы.

```vb
Public Function ReadContourArcSweep(ByVal value As Variant, ByVal inputLocation As String) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Variant; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |
| inputLocation | String; ByVal; обязательный | Строка реального исходного адреса для диагностического сообщения. |

Тип результата: Double. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modGeometryTypes.bas:93

Минимальный вызов; требуемый контекст: В теле modTestRecipeApiCatalogue.RunApiInputAndDataExamples с его подготовленными локальными объектами/входом.

```vb
    Near "saved.sweep.strict", ReadContourArcSweep("1.25", "ProgramInputs.Arc"), 1.25
```

Полный контекст: tests/modTestRecipeApiCatalogue.bas:246.

Native entrypoint: modTestRecipeApiCatalogue.RunApiInputAndDataExamples; журнал: docs/regression/GeometryDeclarative_2026-10-12/NativeGuide/ApiCallExamples.txt.

<a id="api-modgeometrytypes-buildgeneratedsectionmodel"></a>
### modGeometryTypes.BuildGeneratedSectionModel

Общий тонкий adapter mesh/layout→CSectionModel; автор рецепта использует CSectionBuilder, а не отдельный генератор типа.

```vb
Public Function BuildGeneratedSectionModel(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, Optional ByVal sourceType As String = "Generated") As CSectionModel
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| mesh | CFiberMeshBuilder; ByVal; обязательный | Подготовленный объект/контейнер выбранного запроса, без повторного чтения Excel или назначения типа. |
| rebars | CRebarLayout; ByVal; обязательный | Подготовленный объект/контейнер выбранного запроса, без повторного чтения Excel или назначения типа. |
| sourceType | String; ByVal; Optional; default="Generated" | Скалярная metadata владельца/источника; геометрическая классификация от строки не зависит. |

Тип результата: CSectionModel. Назначение, единицы и ограничения приведены выше. Для объектного результата используйте Set.

Жизненный цикл: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modGeometryTypes.bas:119

Минимальный вызов; требуемый контекст: В теле modTestGeometry.TestRectangularMeshSteps с его подготовленными локальными объектами/входом.

```vb
    Set model = BuildGeneratedSectionModel(mesh, Nothing)
```

Полный контекст: tests/modTestGeometry.bas:1390.

Native entrypoint: modTestGeometry.RunGeometryTests; журнал: docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/GeometryCore.txt.

