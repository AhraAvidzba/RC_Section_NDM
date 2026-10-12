# SECTION_RECIPE_API: каталог текущих сигнатур

Стадия: DRAFT. Сигнатуры, defaults, параметры и жизненный цикл синхронизированы с source; минимальные компилируемые примеры/ссылки и native Help ещё проверяются. Это не закрытый E04/E12.

Все Public члены фасада, его options/результатов и общего ввода/единиц; общие query/region и build diagnostics для проверок. Private/Friend не API рецепта.

Внутри мм/мм²/Н/Н*мм/МПа; углы в радианах. ByRef/Optional/default приведены в точной сигнатуре. Полный общий фасад и исходный класс представлены в ADDING_SECTION_TYPE.md.

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

Результат: CSectionDefinition. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Вызывается реестром с готовым context; создаёт новое definition. Не выполняет mesh/layout/solver/Excel output.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/ISectionRecipe.cls:14

## CSectionDefinitionBuilder

New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

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

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:18

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

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:24

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

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:33

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

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:42

<a id="api-csectiondefinitionbuilder-roundedpolygon"></a>
### CSectionDefinitionBuilder.RoundedPolygon

Создаёт согласованную границу по вершинам/радиусам/ID. Edge.ID начинается у указанной вершины, Corner.ID существует только при R>0.

```vb
Public Function RoundedPolygon(ByVal pathID As String, ByVal vertices As Variant, ByVal radii As Variant, ByVal vertexIDs As Variant) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| vertices | Variant; ByVal; обязательный | Двумерные вершины: Array(Array(x,y),...); координаты во внутренних мм. |
| radii | Variant; ByVal; обязательный | Массив неотрицательных радиусов вершин в мм, длина равна vertices; clamp бетонного радиуса отсутствует. |
| vertexIDs | Variant; ByVal; обязательный | Массив уникальных смысловых ID вершин; определяет Edge.ID/Corner.ID. |

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:52

<a id="api-csectiondefinitionbuilder-offsetroundedpolygon"></a>
### CSectionDefinitionBuilder.OffsetRoundedPolygon

Смещает весь скруглённый многоугольник с согласованными стыками. Положительное расстояние внутрь площади; отрицательное наружу, из Opening в бетон.

```vb
Public Function OffsetRoundedPolygon(ByVal pathID As String, ByVal vertices As Variant, ByVal radii As Variant, ByVal vertexIDs As Variant, ByVal distance As Double) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| vertices | Variant; ByVal; обязательный | Двумерные вершины: Array(Array(x,y),...); координаты во внутренних мм. |
| radii | Variant; ByVal; обязательный | Массив неотрицательных радиусов вершин в мм, длина равна vertices; clamp бетонного радиуса отсутствует. |
| vertexIDs | Variant; ByVal; обязательный | Массив уникальных смысловых ID вершин; определяет Edge.ID/Corner.ID. |
| distance | Double; ByVal; обязательный | Signed offset, мм: положительно внутрь площади, отрицательно наружу; это не left-offset обхода. |

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:59

<a id="api-csectiondefinitionbuilder-addpath"></a>
### CSectionDefinitionBuilder.AddPath

Добавляет уже опубликованный непрерывный CGeometryPath в определение; не включает его автоматически в бетон.

```vb
Public Sub AddPath(ByVal path As CGeometryPath)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| path | CGeometryPath; ByVal; обязательный | Существующий непрерывный опубликованный CGeometryPath. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:66

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

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:72

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

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:77

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

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:85

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

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:90

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

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:99

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

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:109

<a id="api-csectiondefinitionbuilder-lastgeometrydiagnostic"></a>
### CSectionDefinitionBuilder.LastGeometryDiagnostic

Возвращает последнюю диагностику query фасада с SourceMap. Nothing означает отсутствие записанного отказа.

```vb
Public Property Get LastGeometryDiagnostic() As CGeometryDiagnostic
```

Результат: CGeometryDiagnostic. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:121

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

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:131

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

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:141

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

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:149

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

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:155

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

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:161

<a id="api-csectiondefinitionbuilder-outer"></a>
### CSectionDefinitionBuilder.Outer

Назначает замкнутому пути роль наружной границы материала. Не требует наличия арматуры.

```vb
Public Sub Outer(ByVal loopID As String, ByVal pathID As String, Optional ByVal comment As String = "", Optional ByVal contourIDs As Variant, Optional ByVal arcComment As String = "")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| loopID | String; ByVal; обязательный | Непустой устойчивый ID контура; у CConcreteRegion численный индекс loop 1..LoopCount. |
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| comment | String; ByVal; Optional; default="" | Пояснение происхождения/инженерного смысла; не управляющий статус. |
| contourIDs | Variant; ByVal; Optional; пропущенный Variant=Missing | Необязательный массив ID сегментов сохранённого контура, по одному на каждый сегмент пути. |
| arcComment | String; ByVal; Optional; default="" | Отдельный комментарий аналитических дуг; пустой использует обычный comment. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:166

<a id="api-csectiondefinitionbuilder-opening"></a>
### CSectionDefinitionBuilder.Opening

Назначает замкнутому пути роль отверстия с конкретным parentOuterID. Действительное вложение проверяется при Finish.

```vb
Public Sub Opening(ByVal loopID As String, ByVal pathID As String, ByVal parentOuterID As String, Optional ByVal comment As String = "", Optional ByVal contourIDs As Variant, Optional ByVal arcComment As String = "")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| loopID | String; ByVal; обязательный | Непустой устойчивый ID контура; у CConcreteRegion численный индекс loop 1..LoopCount. |
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| parentOuterID | String; ByVal; обязательный | ID ранее объявленной наружной границы для Opening; проверяется реальное вложение. |
| comment | String; ByVal; Optional; default="" | Пояснение происхождения/инженерного смысла; не управляющий статус. |
| contourIDs | Variant; ByVal; Optional; пропущенный Variant=Missing | Необязательный массив ID сегментов сохранённого контура, по одному на каждый сегмент пути. |
| arcComment | String; ByVal; Optional; default="" | Отдельный комментарий аналитических дуг; пустой использует обычный comment. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:175

<a id="api-csectiondefinitionbuilder-rebarbycount"></a>
### CSectionDefinitionBuilder.RebarByCount

Объявляет n станций на всём рабочем пути с выбранным EndMode; для группы n применяется к каждому пути. n=0 или d=0 выключает правило.

```vb
Public Sub RebarByCount(ByVal ruleID As String, ByVal target As String, ByVal count As Long, ByVal diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal steelClass As String = "Rebar")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |
| target | String; ByVal; обязательный | ID пути/PathGroup, разрешаемый в этом определении. |
| count | Long; ByVal; обязательный | Целый неотрицательный счётчик; для ByCount применяется отдельно к каждому пути. |
| diameter | Double; ByVal; обязательный | Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд. |
| options | CRebarRuleOptions; ByVal; Optional; default=Nothing | CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт реальные default options. |
| steelClass | String; ByVal; Optional; default="Rebar" | Строка принадлежности материала арматуры; совпадения разных материалов требуют явной политики. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:185

<a id="api-csectiondefinitionbuilder-rebarbymaxstep"></a>
### CSectionDefinitionBuilder.RebarByMaxStep

Объявляет распределение с достаточным количеством станций для ограничения максимального шага. Фактический шаг может быть меньше заданного.

```vb
Public Sub RebarByMaxStep(ByVal ruleID As String, ByVal target As String, ByVal step As Double, ByVal diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal steelClass As String = "Rebar")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |
| target | String; ByVal; обязательный | ID пути/PathGroup, разрешаемый в этом определении. |
| step | Double; ByVal; обязательный | Положительный шаг вдоль аналитического пути, мм. |
| diameter | Double; ByVal; обязательный | Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд. |
| options | CRebarRuleOptions; ByVal; Optional; default=Nothing | CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт реальные default options. |
| steelClass | String; ByVal; Optional; default="Rebar" | Строка принадлежности материала арматуры; совпадения разных материалов требуют явной политики. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:194

<a id="api-csectiondefinitionbuilder-rebarbyfixedstep"></a>
### CSectionDefinitionBuilder.RebarByFixedStep

Объявляет точный шаг и явное размещение остатка. IncludeBoth/ClosedUniform требуют совместимости рабочей длины с шагом.

```vb
Public Sub RebarByFixedStep(ByVal ruleID As String, ByVal target As String, ByVal step As Double, ByVal diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal steelClass As String = "Rebar")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |
| target | String; ByVal; обязательный | ID пути/PathGroup, разрешаемый в этом определении. |
| step | Double; ByVal; обязательный | Положительный шаг вдоль аналитического пути, мм. |
| diameter | Double; ByVal; обязательный | Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд. |
| options | CRebarRuleOptions; ByVal; Optional; default=Nothing | CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт реальные default options. |
| steelClass | String; ByVal; Optional; default="Rebar" | Строка принадлежности материала арматуры; совпадения разных материалов требуют явной политики. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:201

<a id="api-csectiondefinitionbuilder-rebaratstations"></a>
### CSectionDefinitionBuilder.RebarAtStations

Объявляет явные станции по длине пути. Порядок массива задаёт исходные ordinals.

```vb
Public Sub RebarAtStations(ByVal ruleID As String, ByVal target As String, ByVal stations As Variant, ByVal diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal steelClass As String = "Rebar")
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |
| target | String; ByVal; обязательный | ID пути/PathGroup, разрешаемый в этом определении. |
| stations | Variant; ByVal; обязательный | Массив численных станций вдоль полного пути, мм; порядок задаёт ordinals. |
| diameter | Double; ByVal; обязательный | Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд. |
| options | CRebarRuleOptions; ByVal; Optional; default=Nothing | CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт реальные default options. |
| steelClass | String; ByVal; Optional; default="Rebar" | Строка принадлежности материала арматуры; совпадения разных материалов требуют явной политики. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:208

<a id="api-csectiondefinitionbuilder-rebardependent"></a>
### CSectionDefinitionBuilder.RebarDependent

Получает связанные позиции по TangentOffset/NormalOffset исходного ряда; не выполняет независимое перераспределение на новом пути.

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

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:215

<a id="api-csectiondefinitionbuilder-rebarprojected"></a>
### CSectionDefinitionBuilder.RebarProjected

Проецирует исходные позиции на конечный target. Setbacks фильтруют целевой участок; отсутствие hit обрабатывается явно.

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

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:222

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

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:240

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

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:249

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

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:258

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

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:270

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

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:282

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

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:295

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

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:304

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

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:309

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

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:325

<a id="api-csectiondefinitionbuilder-rebarlabel"></a>
### CSectionDefinitionBuilder.RebarLabel

Привязывает фактическую подпись группы к габариту target-пути, со стороны normalX/Y.

```vb
Public Sub RebarLabel(ByVal groupID As String, ByVal target As String, ByVal normalX As Double, ByVal normalY As Double, Optional ByVal axisDistance As Double = 0#)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| groupID | String; ByVal; обязательный | Смысловой ID группы; путь и арматурная annotation-group — разные контракты. |
| target | String; ByVal; обязательный | ID пути/PathGroup, разрешаемый в этом определении. |
| normalX | Double; ByVal; обязательный | X-компонента стороны размерной/подписной линии; layout нормирует вектор. |
| normalY | Double; ByVal; обязательный | Y-компонента стороны размерной/подписной линии. |
| axisDistance | Double; ByVal; Optional; default=0# | Положительное расстояние до оси, мм; не защитный слой до поверхности и не продольный setback. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:334

<a id="api-csectiondefinitionbuilder-rebarrowlabel"></a>
### CSectionDefinitionBuilder.RebarRowLabel

Привязывает фактическую подпись группы к концам исполненного ruleID.

```vb
Public Sub RebarRowLabel(ByVal groupID As String, ByVal ruleID As String, ByVal normalX As Double, ByVal normalY As Double, Optional ByVal axisDistance As Double = 0#)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| groupID | String; ByVal; обязательный | Смысловой ID группы; путь и арматурная annotation-group — разные контракты. |
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |
| normalX | Double; ByVal; обязательный | X-компонента стороны размерной/подписной линии; layout нормирует вектор. |
| normalY | Double; ByVal; обязательный | Y-компонента стороны размерной/подписной линии. |
| axisDistance | Double; ByVal; Optional; default=0# | Положительное расстояние до оси, мм; не защитный слой до поверхности и не продольный setback. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:343

<a id="api-csectiondefinitionbuilder-rebarrowextentlabel"></a>
### CSectionDefinitionBuilder.RebarRowExtentLabel

Привязывает подпись к фактическому габариту позиций правила; reverseDirection меняет ориентацию anchor-линии.

```vb
Public Sub RebarRowExtentLabel(ByVal groupID As String, ByVal ruleID As String, ByVal normalX As Double, ByVal normalY As Double, Optional ByVal axisDistance As Double = 0#, Optional ByVal reverseDirection As Boolean = False)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| groupID | String; ByVal; обязательный | Смысловой ID группы; путь и арматурная annotation-group — разные контракты. |
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |
| normalX | Double; ByVal; обязательный | X-компонента стороны размерной/подписной линии; layout нормирует вектор. |
| normalY | Double; ByVal; обязательный | Y-компонента стороны размерной/подписной линии. |
| axisDistance | Double; ByVal; Optional; default=0# | Положительное расстояние до оси, мм; не защитный слой до поверхности и не продольный setback. |
| reverseDirection | Boolean; ByVal; Optional; default=False | Обращает направление линии привязки подписи; физический ряд не меняется. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:353

<a id="api-csectiondefinitionbuilder-finish"></a>
### CSectionDefinitionBuilder.Finish

Проверяет контуры, ссылки, зависимости и аннотации; публикует неизменяемое CSectionDefinition. Фасад после этого закрыт.

```vb
Public Function Finish() As CSectionDefinition
```

Результат: CSectionDefinition. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:362

<a id="api-csectiondefinitionbuilder-rebarrowextremelabel"></a>
### CSectionDefinitionBuilder.RebarRowExtremeLabel

Выбирает крайние созданные позиции в directionX/Y; сторона текста normalX/Y независима от выбора ряда.

```vb
Public Sub RebarRowExtremeLabel(ByVal groupID As String, ByVal ruleID As String, ByVal directionX As Double, ByVal directionY As Double, ByVal normalX As Double, ByVal normalY As Double, Optional ByVal axisDistance As Double = 0#)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| groupID | String; ByVal; обязательный | Смысловой ID группы; путь и арматурная annotation-group — разные контракты. |
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |
| directionX | Double; ByVal; обязательный | X-компонента направления проекции/выбора экстремума; вектор не нулевой, у RebarRowExtremeLabel единичный. |
| directionY | Double; ByVal; обязательный | Y-компонента направления проекции/выбора экстремума. |
| normalX | Double; ByVal; обязательный | X-компонента стороны размерной/подписной линии; layout нормирует вектор. |
| normalY | Double; ByVal; обязательный | Y-компонента стороны размерной/подписной линии. |
| axisDistance | Double; ByVal; Optional; default=0# | Положительное расстояние до оси, мм; не защитный слой до поверхности и не продольный setback. |

Жизненный цикл: New→Initialize один раз→декларации→Finish. Ссылки требуют ранее объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации запрещены.

Ошибки/исправление: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks, незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по диагностике.

Source: src/Section/CSectionDefinitionBuilder.cls:371

## CGeometryPath

New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

<a id="api-cgeometrypath-lastdiagnostic"></a>
### CGeometryPath.LastDiagnostic

Последняя машинная диагностика; Nothing при отсутствии записанного события.

```vb
Public Property Get LastDiagnostic() As CGeometryDiagnostic
```

Результат: CGeometryDiagnostic. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:26

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

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:30

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

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:37

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

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:46

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

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:61

<a id="api-cgeometrypath-id"></a>
### CGeometryPath.ID

Смысловой ID этого пути.

```vb
Public Property Get ID() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:76

<a id="api-cgeometrypath-segmentcount"></a>
### CGeometryPath.SegmentCount

Число аналитических line/arc сегментов.

```vb
Public Property Get SegmentCount() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:77

<a id="api-cgeometrypath-issealed"></a>
### CGeometryPath.IsSealed

Опубликован ли объект; published геометрия не изменяется.

```vb
Public Property Get IsSealed() As Boolean
```

Результат: Boolean. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:78

<a id="api-cgeometrypath-isclosed"></a>
### CGeometryPath.IsClosed

Замкнут ли фактический путь, без фиктивного соединения.

```vb
Public Property Get IsClosed() As Boolean
```

Результат: Boolean. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:79

<a id="api-cgeometrypath-length"></a>
### CGeometryPath.Length

Точная суммарная длина пути, мм.

```vb
Public Property Get Length() As Double
```

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:80

<a id="api-cgeometrypath-segmentid"></a>
### CGeometryPath.SegmentID

Возвращает смысловой ID сегмента по индексу.

```vb
Public Property Get SegmentID(ByVal index As Long) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:81

<a id="api-cgeometrypath-sourcepathid"></a>
### CGeometryPath.SourcePathID

ID исходного пути сегмента после slice/transform/offset.

```vb
Public Property Get SourcePathID(ByVal index As Long) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:85

<a id="api-cgeometrypath-sourcesegmentid"></a>
### CGeometryPath.SourceSegmentID

Смысловой исходный ID сегмента, сохраняемый при распаде/переименовании результата.

```vb
Public Property Get SourceSegmentID(ByVal index As Long) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:89

<a id="api-cgeometrypath-sourcestartparameter"></a>
### CGeometryPath.SourceStartParameter

Начальный параметр исходного сегмента; может выйти за [0;1] при явном продлении прямой.

```vb
Public Property Get SourceStartParameter(ByVal index As Long) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:93

<a id="api-cgeometrypath-sourceendparameter"></a>
### CGeometryPath.SourceEndParameter

Конечный параметр исходного сегмента, включая обращённый обход.

```vb
Public Property Get SourceEndParameter(ByVal index As Long) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:97

<a id="api-cgeometrypath-segmentindex"></a>
### CGeometryPath.SegmentIndex

Разрешает существующий смысловой ID в текущий индекс; исчезнувшая дуга даёт MISSING_SEGMENT.

```vb
Public Function SegmentIndex(ByVal segmentID As String) As Long
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| segmentID | String; ByVal; обязательный | Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом. |

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:101

<a id="api-cgeometrypath-anchorstation"></a>
### CGeometryPath.AnchorStation

Переводит ID сегмента и параметр [0;1] в станцию вдоль полного пути, мм.

```vb
Public Function AnchorStation(ByVal segmentID As String, Optional ByVal parameter As Double = 0#) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| segmentID | String; ByVal; обязательный | Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом. |
| parameter | Double; ByVal; Optional; default=0# | Нормированный параметр точки сегмента [0;1], не расстояние по всему пути. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:110

<a id="api-cgeometrypath-withid"></a>
### CGeometryPath.WithID

Копирует опубликованный путь с новым ID, сохраняя сегментные ссылки и исходное происхождение.

```vb
Public Function WithID(ByVal pathID As String) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:118

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

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:130

<a id="api-cgeometrypath-atdistance"></a>
### CGeometryPath.AtDistance

Возвращает точку и единичную касательную на станции. На стыке можно явно выбрать входящую касательную.

```vb
Public Sub AtDistance(ByVal distance As Double, ByRef x As Double, ByRef y As Double, ByRef tangentX As Double, ByRef tangentY As Double, Optional ByVal incomingTangentAtJoint As Boolean = False)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| distance | Double; ByVal; обязательный | Длина/станция/граница distance, внутренние мм; направление и пределы задаёт выбранный запрос. |
| x | Double; ByRef; обязательный | Модельная координата x, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| y | Double; ByRef; обязательный | Модельная координата y, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| tangentX | Double; ByRef; обязательный | Компонента единичной касательной, без размерности. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| tangentY | Double; ByRef; обязательный | Компонента единичной касательной, без размерности. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| incomingTangentAtJoint | Boolean; ByVal; Optional; default=False | Явный Boolean режима замыкания/касательной; не исправляет геометрию автоматически. |

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:158

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
| minimum | Double; ByRef; обязательный | Длина/станция/граница minimum, внутренние мм; направление и пределы задаёт выбранный запрос. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| maximum | Double; ByRef; обязательный | Длина/станция/граница maximum, внутренние мм; направление и пределы задаёт выбранный запрос. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:180

<a id="api-cgeometrypath-slice"></a>
### CGeometryPath.Slice

Обрезает полный непрерывный путь по длине; точные дуги и source parameters сохраняются.

```vb
Public Function Slice(ByVal startDistance As Double, ByVal endDistance As Double) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| startDistance | Double; ByVal; обязательный | Станция вдоль полного пути, мм, в [0;Length]; endDistance строго больше startDistance. |
| endDistance | Double; ByVal; обязательный | Станция вдоль полного пути, мм, в [0;Length]; endDistance строго больше startDistance. |

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:195

<a id="api-cgeometrypath-reversed"></a>
### CGeometryPath.Reversed

Копирует путь с обратным порядком сегментов/обходом; signed sweep и source parameters обращаются.

```vb
Public Function Reversed() As CGeometryPath
```

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:238

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

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:252

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
| pathDistance | Double; ByRef; обязательный | Длина/станция/граница pathDistance, внутренние мм; направление и пределы задаёт выбранный запрос. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| rayDistance | Double; ByRef; обязательный | Длина/станция/граница rayDistance, внутренние мм; направление и пределы задаёт выбранный запрос. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |

Результат: Boolean. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New→Initialize→AddLine/AddArc→Seal. Методы чтения/slice/transform требуют Seal. Они возвращают отдельные опубликованные пути; исходный не меняется.

Ошибки/исправление: 5520: пустой ID/сегмент, разрыв, невалидный R/sweep, незамкнутые концы, отсутствие сегмента, станция вне пути. Исправьте геометрию/ID/рабочий интервал.

Source: src/Geometry/CGeometryPath.cls:274

## CSectionDefinition

Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

<a id="api-csectiondefinition-lastdiagnostic"></a>
### CSectionDefinition.LastDiagnostic

Последняя машинная диагностика; Nothing при отсутствии записанного события.

```vb
Public Property Get LastDiagnostic() As CGeometryDiagnostic
```

Результат: CGeometryDiagnostic. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:23

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

Source: src/Section/CSectionDefinition.cls:44

<a id="api-csectiondefinition-sourcereference"></a>
### CSectionDefinition.SourceReference

Читает скалярный SourceMap по ID; отсутствующая связь возвращает пустую строку.

```vb
Public Function SourceReference(ByVal entityID As String) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| entityID | String; ByVal; обязательный | Смысловой ID сущности для SourceMap. |

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:55

<a id="api-csectiondefinition-typeid"></a>
### CSectionDefinition.TypeID

Смысловое имя типа опубликованного определения.

```vb
Public Property Get TypeID() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:59

<a id="api-csectiondefinition-schemaversion"></a>
### CSectionDefinition.SchemaVersion

Положительная версия схемы определения.

```vb
Public Property Get SchemaVersion() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:60

<a id="api-csectiondefinition-issealed"></a>
### CSectionDefinition.IsSealed

Опубликован ли объект; published геометрия не изменяется.

```vb
Public Property Get IsSealed() As Boolean
```

Результат: Boolean. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:61

<a id="api-csectiondefinition-pathcount"></a>
### CSectionDefinition.PathCount

Число зарегистрированных путей.

```vb
Public Property Get PathCount() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:62

<a id="api-csectiondefinition-loopcount"></a>
### CSectionDefinition.LoopCount

Число подготовленных материальных loops.

```vb
Public Property Get LoopCount() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:63

<a id="api-csectiondefinition-rulecount"></a>
### CSectionDefinition.RuleCount

Число объявленных правил, включая выключенные.

```vb
Public Property Get RuleCount() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:64

<a id="api-csectiondefinition-annotationcount"></a>
### CSectionDefinition.AnnotationCount

Число объявленных annotation rules.

```vb
Public Property Get AnnotationCount() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:65

<a id="api-csectiondefinition-path"></a>
### CSectionDefinition.Path

Читает опубликованный неизменяемый путь по ID.

```vb
Public Function Path(ByVal pathID As String) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:67

<a id="api-csectiondefinition-targetpaths"></a>
### CSectionDefinition.TargetPaths

Возвращает ID одного пути или массив независимых путей группы.

```vb
Public Function TargetPaths(ByVal targetID As String) As Variant
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| targetID | String; ByVal; обязательный | ID пути или группы путей. |

Результат: Variant. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:73

<a id="api-csectiondefinition-concretedomain"></a>
### CSectionDefinition.ConcreteDomain

Возвращает clone подготовленной бетонной области, без Excel или повторной интерпретации рецепта.

```vb
Public Function ConcreteDomain() As CConcreteRegion
```

Результат: CConcreteRegion. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:83

<a id="api-csectiondefinition-materialcontours"></a>
### CSectionDefinition.MaterialContours

Возвращает отдельный контейнер аналитических материальных Outer/Opening.

```vb
Public Function MaterialContours() As CSectionContours
```

Результат: CSectionContours. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу возвращаются скаляры, копии массивов или sealed paths/clones.

Ошибки/исправление: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.

Source: src/Section/CSectionDefinition.cls:280

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

Source: src/Section/CSectionDefinition.cls:287

## CRebarRuleOptions

New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

<a id="api-crebarruleoptions-startsetback"></a>
### CRebarRuleOptions.StartSetback

Начальное отсечение вдоль каждого target-пути, мм≥0; у полного ClosedUniform 0.

```vb
Public StartSetback As Double
```

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:14

<a id="api-crebarruleoptions-endsetback"></a>
### CRebarRuleOptions.EndSetback

Конечное отсечение вдоль каждого target-пути, мм≥0; ProjectFrom фильтрует hit без перераспределения.

```vb
Public EndSetback As Double
```

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:15

<a id="api-crebarruleoptions-endmode"></a>
### CRebarRuleOptions.EndMode

IncludeBoth/ExcludeBoth/IncludeStart/IncludeEnd/ClosedUniform; последние требуют соответствующей замкнутости.

```vb
Public EndMode As ERebarEndMode
```

Результат: ERebarEndMode. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:16

<a id="api-crebarruleoptions-distributionscope"></a>
### CRebarRuleOptions.DistributionScope

Только PerPath: количество/шаг/setbacks отдельно для каждого пути группы.

```vb
Public DistributionScope As ERebarDistributionScope
```

Результат: ERebarDistributionScope. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:17

<a id="api-crebarruleoptions-remainderplacement"></a>
### CRebarRuleOptions.RemainderPlacement

RemainderStart/End/Center: положение остатка для точного шага.

```vb
Public RemainderPlacement As ERemainderPlacement
```

Результат: ERemainderPlacement. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:18

<a id="api-crebarruleoptions-tangentoffset"></a>
### CRebarRuleOptions.TangentOffset

Локальное связанное смещение вдоль касательной, мм; signed, не независимый шаг ряда.

```vb
Public TangentOffset As Double
```

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:19

<a id="api-crebarruleoptions-normaloffset"></a>
### CRebarRuleOptions.NormalOffset

Локальное связанное смещение по левой нормали (-ty,tx), мм; signed.

```vb
Public NormalOffset As Double
```

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:20

<a id="api-crebarruleoptions-everysecond"></a>
### CRebarRuleOptions.EverySecond

Брать каждую вторую исходную позицию по SourceOrdinal, начиная с первой.

```vb
Public EverySecond As Boolean
```

Результат: Boolean. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:21

<a id="api-crebarruleoptions-incomingtangentatjoint"></a>
### CRebarRuleOptions.IncomingTangentAtJoint

На станции стыка брать входящую касательную вместо default outgoing.

```vb
Public IncomingTangentAtJoint As Boolean
```

Результат: Boolean. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:22

<a id="api-crebarruleoptions-directionx"></a>
### CRebarRuleOptions.DirectionX

Глобальная X направления ProjectFrom, используется при SourceNormalSign=0.

```vb
Public DirectionX As Double
```

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:23

<a id="api-crebarruleoptions-directiony"></a>
### CRebarRuleOptions.DirectionY

Глобальная Y направления ProjectFrom; направление не нулевое.

```vb
Public DirectionY As Double
```

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:24

<a id="api-crebarruleoptions-sourcenormalsign"></a>
### CRebarRuleOptions.SourceNormalSign

-1/0/+1: локальная нормаль исходного ряда или глобальное DirectionX/Y при 0.

```vb
Public SourceNormalSign As Double
```

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:25

<a id="api-crebarruleoptions-successfulprojectionordinal"></a>
### CRebarRuleOptions.SuccessfulProjectionOrdinal

Нумеровать SourceOrdinal успешных hit подряд; False сохраняет исходные ordinals с пропусками.

```vb
Public SuccessfulProjectionOrdinal As Boolean
```

Результат: Boolean. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:26

<a id="api-crebarruleoptions-newpublicationroot"></a>
### CRebarRuleOptions.NewPublicationRoot

У ProjectFrom начинает новый корень публикации по успешным позициям.

```vb
Public NewPublicationRoot As Boolean
```

Результат: Boolean. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:27

<a id="api-crebarruleoptions-skipmissingprojection"></a>
### CRebarRuleOptions.SkipMissingProjection

Разрешить пропуск без hit с PROJECTION_MISSING; False выдаёт ошибку.

```vb
Public SkipMissingProjection As Boolean
```

Результат: Boolean. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:28

<a id="api-crebarruleoptions-publication"></a>
### CRebarRuleOptions.Publication

PublishByRule или PublishPerSourcePosition; влияет на вывод, не геометрический шаг.

```vb
Public Publication As ERebarPublication
```

Результат: ERebarPublication. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:29

<a id="api-crebarruleoptions-positionorder"></a>
### CRebarRuleOptions.PositionOrder

AlongPath или FromPathMiddle; геометрические станции/исходные ordinals сохраняются.

```vb
Public PositionOrder As ERebarPositionOrder
```

Результат: ERebarPositionOrder. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:30

<a id="api-crebarruleoptions-coincidence"></a>
### CRebarRuleOptions.Coincidence

RejectCoincidence/MergeIdentical/KeepPublishedOwner; последнее сохраняет явно выбранного опубликованного владельца.

```vb
Public Coincidence As ERebarCoincidence
```

Результат: ERebarCoincidence. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:31

<a id="api-crebarruleoptions-placementcheck"></a>
### CRebarRuleOptions.PlacementCheck

CheckCenter или CheckFullDisk относительно Outer/Opening; ужесточение задаётся явно.

```vb
Public PlacementCheck As ERebarPlacementCheck
```

Результат: ERebarPlacementCheck. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:32

<a id="api-crebarruleoptions-overlappolicy"></a>
### CRebarRuleOptions.OverlapPolicy

ReportOverlap или RejectOverlap. Контакт допустим; нормативный минимальный зазор не вводится.

```vb
Public OverlapPolicy As ERebarOverlapPolicy
```

Результат: ERebarOverlapPolicy. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:33

<a id="api-crebarruleoptions-sourceprefix"></a>
### CRebarRuleOptions.SourcePrefix

Префикс SourcePositionId; пустой использует RuleId+'.' либо связь source+suffix.

```vb
Public SourcePrefix As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:34

<a id="api-crebarruleoptions-sourcesuffix"></a>
### CRebarRuleOptions.SourceSuffix

Суффикс SourcePositionId, в том числе для связанных рядов.

```vb
Public SourceSuffix As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:35

<a id="api-crebarruleoptions-annotationgroup"></a>
### CRebarRuleOptions.AnnotationGroup

Смысловая принадлежность подписи физического стержня; не расчётная группа трещин СП35.

```vb
Public AnnotationGroup As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:36

<a id="api-crebarruleoptions-comment"></a>
### CRebarRuleOptions.Comment

Инженерное пояснение правила; не selector или ResultCode.

```vb
Public Comment As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:37

<a id="api-crebarruleoptions-sourcereference"></a>
### CRebarRuleOptions.SourceReference

Реальные адреса исходных параметров правила или программная ссылка.

```vb
Public SourceReference As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options в definition; последующее изменение options не меняет правило.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Section/CRebarRuleOptions.cls:38

## CGeometryDiagnostic

Создаётся сервисом как immutable snapshot. Читайте через LastDiagnostic/DiagnosticAt; диагностический объект не изменяет геометрию или статусы.

<a id="api-cgeometrydiagnostic-code"></a>
### CGeometryDiagnostic.Code

Машинный код причины; назначается напрямую, не извлекается из русского текста.

```vb
Public Property Get Code() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Создаётся сервисом как immutable snapshot. Читайте через LastDiagnostic/DiagnosticAt; диагностический объект не изменяет геометрию или статусы.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Geometry/CGeometryDiagnostic.cls:24

<a id="api-cgeometrydiagnostic-entityid"></a>
### CGeometryDiagnostic.EntityID

Смысловой ID сущности отказа, если он известен этому сервису.

```vb
Public Property Get EntityID() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Создаётся сервисом как immutable snapshot. Читайте через LastDiagnostic/DiagnosticAt; диагностический объект не изменяет геометрию или статусы.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Geometry/CGeometryDiagnostic.cls:25

<a id="api-cgeometrydiagnostic-relatedentityid"></a>
### CGeometryDiagnostic.RelatedEntityID

ID связанного сегмента/позиции/стержня для локализации причины.

```vb
Public Property Get RelatedEntityID() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Создаётся сервисом как immutable snapshot. Читайте через LastDiagnostic/DiagnosticAt; диагностический объект не изменяет геометрию или статусы.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Geometry/CGeometryDiagnostic.cls:26

<a id="api-cgeometrydiagnostic-reason"></a>
### CGeometryDiagnostic.Reason

Человекочитаемая причина; не selector/status.

```vb
Public Property Get Reason() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Создаётся сервисом как immutable snapshot. Читайте через LastDiagnostic/DiagnosticAt; диагностический объект не изменяет геометрию или статусы.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Geometry/CGeometryDiagnostic.cls:27

<a id="api-cgeometrydiagnostic-detail1"></a>
### CGeometryDiagnostic.Detail1

Первое численное поле конкретного Code; единицы определяет Code.

```vb
Public Property Get Detail1() As Double
```

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Создаётся сервисом как immutable snapshot. Читайте через LastDiagnostic/DiagnosticAt; диагностический объект не изменяет геометрию или статусы.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Geometry/CGeometryDiagnostic.cls:28

<a id="api-cgeometrydiagnostic-detail2"></a>
### CGeometryDiagnostic.Detail2

Второе численное поле конкретного Code.

```vb
Public Property Get Detail2() As Double
```

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Создаётся сервисом как immutable snapshot. Читайте через LastDiagnostic/DiagnosticAt; диагностический объект не изменяет геометрию или статусы.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Geometry/CGeometryDiagnostic.cls:29

<a id="api-cgeometrydiagnostic-sourcereference"></a>
### CGeometryDiagnostic.SourceReference

Реальный адрес/программный источник, добавленный SourceMap или правилом.

```vb
Public Property Get SourceReference() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Создаётся сервисом как immutable snapshot. Читайте через LastDiagnostic/DiagnosticAt; диагностический объект не изменяет геометрию или статусы.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Geometry/CGeometryDiagnostic.cls:30

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

Source: src/Section/CSectionInputTable.cls:18

<a id="api-csectioninputtable-name"></a>
### CSectionInputTable.Name

Имя корневого блока snapshot.

```vb
Public Property Get Name() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:54

<a id="api-csectioninputtable-rowcount"></a>
### CSectionInputTable.RowCount

Число строк неизменяемого input snapshot.

```vb
Public Property Get RowCount() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:55

<a id="api-csectioninputtable-columncount"></a>
### CSectionInputTable.ColumnCount

Число столбцов input snapshot.

```vb
Public Property Get ColumnCount() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:56

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

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:58

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

Результат: Variant. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:62

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

Результат: Boolean. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:68

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

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:72

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

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:83

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

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:101

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

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:109

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

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:119

<a id="api-csectioninputtable-optionaldiameter"></a>
### CSectionInputTable.OptionalDiameter

Пусто/TODO/0 выключает ряд; положительное переводится в мм, отрицательное — ошибка.

```vb
Public Function OptionalDiameter(ByVal row As Long, ByVal column As Long, ByVal units As CUnitSystem) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| row | Long; ByVal; обязательный | Номер строки снимка 1..RowCount. |
| column | Long; ByVal; обязательный | Номер столбца снимка 1..ColumnCount. |
| units | CUnitSystem; ByVal; обязательный | CUnitSystem, уже подготовленная вызывающим workflow; преобразование только через её методы. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:126

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

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:136

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

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:151

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

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного массива не меняет снимок; ошибка Excel не скрывается.

Ошибки/исправление: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.

Source: src/Section/CSectionInputTable.cls:158

## CSectionRecipeContext

Initialize→AddTable→Seal на внешней границе. Рецепт только читает Table/Units; не сохраняет context в готовом definition.

<a id="api-csectionrecipecontext-initialize"></a>
### CSectionRecipeContext.Initialize

Открывает контекст с подготовленной системой единиц один раз.

```vb
Public Sub Initialize(ByVal units As CUnitSystem)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| units | CUnitSystem; ByVal; обязательный | CUnitSystem, уже подготовленная вызывающим workflow; преобразование только через её методы. |

Жизненный цикл: Initialize→AddTable→Seal на внешней границе. Рецепт только читает Table/Units; не сохраняет context в готовом definition.

Ошибки/исправление: 5530: нет units/table, повтор имени, добавление после Seal. Готовьте новый context на внешней границе.

Source: src/Section/CSectionRecipeContext.cls:18

<a id="api-csectionrecipecontext-addtable"></a>
### CSectionRecipeContext.AddTable

Добавляет immutable input snapshot с уникальным именем до Seal.

```vb
Public Sub AddTable(ByVal table As CSectionInputTable)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| table | CSectionInputTable; ByVal; обязательный | Неизменяемый CSectionInputTable. |

Жизненный цикл: Initialize→AddTable→Seal на внешней границе. Рецепт только читает Table/Units; не сохраняет context в готовом definition.

Ошибки/исправление: 5530: нет units/table, повтор имени, добавление после Seal. Готовьте новый context на внешней границе.

Source: src/Section/CSectionRecipeContext.cls:26

<a id="api-csectionrecipecontext-seal"></a>
### CSectionRecipeContext.Seal

Закрывает состав таблиц; после этого добавление невозможно.

```vb
Public Sub Seal()
```

Жизненный цикл: Initialize→AddTable→Seal на внешней границе. Рецепт только читает Table/Units; не сохраняет context в готовом definition.

Ошибки/исправление: 5530: нет units/table, повтор имени, добавление после Seal. Готовьте новый context на внешней границе.

Source: src/Section/CSectionRecipeContext.cls:34

<a id="api-csectionrecipecontext-table"></a>
### CSectionRecipeContext.Table

Возвращает снимок зарегистрированного блока по имени; отсутствие — ошибка, не fallback.

```vb
Public Function Table(ByVal tableName As String) As CSectionInputTable
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| tableName | String; ByVal; обязательный | Имя единственного корневого входного блока этого снимка. |

Результат: CSectionInputTable. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Initialize→AddTable→Seal на внешней границе. Рецепт только читает Table/Units; не сохраняет context в готовом definition.

Ошибки/исправление: 5530: нет units/table, повтор имени, добавление после Seal. Готовьте новый context на внешней границе.

Source: src/Section/CSectionRecipeContext.cls:39

<a id="api-csectionrecipecontext-units"></a>
### CSectionRecipeContext.Units

Система единиц внешнего workflow; не перезагружайте её из Excel в рецепте.

```vb
Public Property Get Units() As CUnitSystem
```

Результат: CUnitSystem. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Initialize→AddTable→Seal на внешней границе. Рецепт только читает Table/Units; не сохраняет context в готовом definition.

Ошибки/исправление: 5530: нет units/table, повтор имени, добавление после Seal. Готовьте новый context на внешней границе.

Source: src/Section/CSectionRecipeContext.cls:45

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

Результат: CSectionRecipeContext. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Внешний workbook+список имён+подготовленные units→новый sealed context. Невыбранные блоки не читаются; live Excel не передаётся геометрическому ядру.

Ошибки/исправление: Адресные ошибки имени/блока/формулы. Повреждённый выбранный блок не равен отсутствию потребителя; восстановите имя/ячейку.

Source: src/Excel/CSectionInputSnapshotReader.cls:14

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

Source: src/Excel/CSectionTypeRegistry.cls:27

<a id="api-csectiontyperegistry-createrecipe"></a>
### CSectionTypeRegistry.CreateRecipe

Возвращает зарегистрированный объект по TypeID; неизвестный тип — ошибка.

```vb
Public Function CreateRecipe(ByVal typeID As String) As ISectionRecipe
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| typeID | String; ByVal; обязательный | Смысловое имя зарегистрированного типа; непустой текст. |

Результат: ISectionRecipe. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Новая instance регистрирует текущие типы. RegisterRecipe добавляет один экземпляр/metadata; consumer выбирает только зарегистрированный TypeID.

Ошибки/исправление: 5530: неизвестный/повторный TypeID, пустой recipe/context, повреждённые metadata. Добавьте одну действительную регистрацию.

Source: src/Excel/CSectionTypeRegistry.cls:43

<a id="api-csectiontyperegistry-recipetablenames"></a>
### CSectionTypeRegistry.RecipeTableNames

Возвращает независимый массив имён ввода, пустой для программного рецепта.

```vb
Public Function RecipeTableNames(ByVal typeID As String) As Variant
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| typeID | String; ByVal; обязательный | Смысловое имя зарегистрированного типа; непустой текст. |

Результат: Variant. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Новая instance регистрирует текущие типы. RegisterRecipe добавляет один экземпляр/metadata; consumer выбирает только зарегистрированный TypeID.

Ошибки/исправление: 5530: неизвестный/повторный TypeID, пустой recipe/context, повреждённые metadata. Добавьте одну действительную регистрацию.

Source: src/Excel/CSectionTypeRegistry.cls:48

<a id="api-csectiontyperegistry-registeredrecipetypes"></a>
### CSectionTypeRegistry.RegisteredRecipeTypes

Возвращает фактически зарегистрированные TypeID — источник штатного списка Config.

```vb
Public Function RegisteredRecipeTypes() As Variant
```

Результат: Variant. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Новая instance регистрирует текущие типы. RegisterRecipe добавляет один экземпляр/metadata; consumer выбирает только зарегистрированный TypeID.

Ошибки/исправление: 5530: неизвестный/повторный TypeID, пустой recipe/context, повреждённые metadata. Добавьте одну действительную регистрацию.

Source: src/Excel/CSectionTypeRegistry.cls:53

<a id="api-csectiontyperegistry-selectedrecipetype"></a>
### CSectionTypeRegistry.SelectedRecipeType

Читает обязательный Geometry.Type через settings и проверяет его по реестру.

```vb
Public Function SelectedRecipeType(ByVal settings As CSystemSettingsReader) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| settings | CSystemSettingsReader; ByVal; обязательный | CSystemSettingsReader, уже загрузивший Config. |

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Новая instance регистрирует текущие типы. RegisterRecipe добавляет один экземпляр/metadata; consumer выбирает только зарегистрированный TypeID.

Ошибки/исправление: 5530: неизвестный/повторный TypeID, пустой recipe/context, повреждённые metadata. Добавьте одну действительную регистрацию.

Source: src/Excel/CSectionTypeRegistry.cls:57

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

Результат: CSectionDefinition. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Новая instance регистрирует текущие типы. RegisterRecipe добавляет один экземпляр/metadata; consumer выбирает только зарегистрированный TypeID.

Ошибки/исправление: 5530: неизвестный/повторный TypeID, пустой recipe/context, повреждённые metadata. Добавьте одну действительную регистрацию.

Source: src/Excel/CSectionTypeRegistry.cls:63

## CSectionBuilder

Один Build выполняет все общие стадии для immutable definition. После успеха доступны Mesh/Rebars/Rules; после нового отказа прошлый успешный результат не выдаётся.

<a id="api-csectionbuilder-build"></a>
### CSectionBuilder.Build

Единая сборка definition→область→сетка/правила→CSectionModel→аннотации; не ветвится по типу.

```vb
Public Function Build(ByVal definition As CSectionDefinition, ByRef meshOptions As TSectionMeshOptions) As CSectionModel
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| definition | CSectionDefinition; ByVal; обязательный | Успешно опубликованный неизменяемый CSectionDefinition. |
| meshOptions | TSectionMeshOptions; ByRef; обязательный | TSectionMeshOptions: StepX/StepY мм>0, BoundarySubdivisions целое≥1, MaterialID≥1. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |

Результат: CSectionModel. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Один Build выполняет все общие стадии для immutable definition. После успеха доступны Mesh/Rebars/Rules; после нового отказа прошлый успешный результат не выдаётся.

Ошибки/исправление: 5530: неопубликованное definition, неверные meshOptions, отсутствие баров/успешной сборки. Исправьте ввод; старый результат после отказа не выдаётся.

Source: src/Section/CSectionBuilder.cls:19

<a id="api-csectionbuilder-mesh"></a>
### CSectionBuilder.Mesh

CFiberMeshBuilder только последней успешно собранной модели; готовую сетку повторно не создавайте.

```vb
Public Property Get Mesh() As CFiberMeshBuilder
```

Результат: CFiberMeshBuilder. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Один Build выполняет все общие стадии для immutable definition. После успеха доступны Mesh/Rebars/Rules; после нового отказа прошлый успешный результат не выдаётся.

Ошибки/исправление: 5530: неопубликованное definition, неверные meshOptions, отсутствие баров/успешной сборки. Исправьте ввод; старый результат после отказа не выдаётся.

Source: src/Section/CSectionBuilder.cls:45

<a id="api-csectionbuilder-rebars"></a>
### CSectionBuilder.Rebars

Физическая раскладка последней успешной сборки; совпадения уже обработаны явно.

```vb
Public Property Get Rebars() As CRebarLayout
```

Результат: CRebarLayout. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Один Build выполняет все общие стадии для immutable definition. После успеха доступны Mesh/Rebars/Rules; после нового отказа прошлый успешный результат не выдаётся.

Ошибки/исправление: 5530: неопубликованное definition, неверные meshOptions, отсутствие баров/успешной сборки. Исправьте ввод; старый результат после отказа не выдаётся.

Source: src/Section/CSectionBuilder.cls:50

<a id="api-csectionbuilder-rules"></a>
### CSectionBuilder.Rules

Исполнитель последней успешной сборки с фактическими счётчиками/шагами/происхождением.

```vb
Public Property Get Rules() As CRebarRuleExecutor
```

Результат: CRebarRuleExecutor. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Один Build выполняет все общие стадии для immutable definition. После успеха доступны Mesh/Rebars/Rules; после нового отказа прошлый успешный результат не выдаётся.

Ошибки/исправление: 5530: неопубликованное definition, неверные meshOptions, отсутствие баров/успешной сборки. Исправьте ввод; старый результат после отказа не выдаётся.

Source: src/Section/CSectionBuilder.cls:55

## CRebarRuleExecutor

Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

<a id="api-crebarruleexecutor-requestedpositions"></a>
### CRebarRuleExecutor.RequestedPositions

Все запрошенные позиции, включая пропущенные проекции; это не число физических стержней.

```vb
Public Property Get RequestedPositions() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:29

<a id="api-crebarruleexecutor-createdpositions"></a>
### CRebarRuleExecutor.CreatedPositions

Действительно добавленные физические стержни текущего Build, даже при отказе в середине.

```vb
Public Property Get CreatedPositions() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:30

<a id="api-crebarruleexecutor-skippedprojections"></a>
### CRebarRuleExecutor.SkippedProjections

Число явно разрешённых пропусков target hit.

```vb
Public Property Get SkippedProjections() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:31

<a id="api-crebarruleexecutor-mergedpositions"></a>
### CRebarRuleExecutor.MergedPositions

Число заявленных позиций, объединённых с физическим владельцем.

```vb
Public Property Get MergedPositions() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:32

<a id="api-crebarruleexecutor-conflicts"></a>
### CRebarRuleExecutor.Conflicts

Число несовместимых/отклонённых совпадений или запрещённых пересечений по явной политике.

```vb
Public Property Get Conflicts() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:33

<a id="api-crebarruleexecutor-baroverlaps"></a>
### CRebarRuleExecutor.BarOverlaps

Полное число пар с отрицательным геометрическим зазором дисков; касание не включается.

```vb
Public Property Get BarOverlaps() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:34

<a id="api-crebarruleexecutor-diagnosticrecordsomitted"></a>
### CRebarRuleExecutor.DiagnosticRecordsOmitted

Число событий сверх 128 сохранённых подробных образцов; полные счётчики не сокращаются.

```vb
Public Property Get DiagnosticRecordsOmitted() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:35

<a id="api-crebarruleexecutor-lastdiagnostic"></a>
### CRebarRuleExecutor.LastDiagnostic

Последняя машинная диагностика; Nothing при отсутствии записанного события.

```vb
Public Property Get LastDiagnostic() As CGeometryDiagnostic
```

Результат: CGeometryDiagnostic. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:36

<a id="api-crebarruleexecutor-diagnosticcount"></a>
### CRebarRuleExecutor.DiagnosticCount

Число сохранённых подробных immutable образцов, 0..128.

```vb
Public Property Get DiagnosticCount() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:37

<a id="api-crebarruleexecutor-diagnosticat"></a>
### CRebarRuleExecutor.DiagnosticAt

Читает immutable подробный образец 1..DiagnosticCount; максимум 128 образцов, счётчики полные.

```vb
Public Function DiagnosticAt(ByVal index As Long) As CGeometryDiagnostic
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Результат: CGeometryDiagnostic. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:41

<a id="api-crebarruleexecutor-positioncount"></a>
### CRebarRuleExecutor.PositionCount

Число сгенерированных заявленных позиций до объединения физических владельцев.

```vb
Public Property Get PositionCount() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:47

<a id="api-crebarruleexecutor-positionruleid"></a>
### CRebarRuleExecutor.PositionRuleID

ID правила заявленной позиции, включая объединённые позиции.

```vb
Public Function PositionRuleID(ByVal index As Long) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:48

<a id="api-crebarruleexecutor-positionpathid"></a>
### CRebarRuleExecutor.PositionPathID

ID геометрического пути заявленной позиции.

```vb
Public Function PositionPathID(ByVal index As Long) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:51

<a id="api-crebarruleexecutor-positionsourceid"></a>
### CRebarRuleExecutor.PositionSourceID

Смысловой SourcePositionId, из которого публиковался физический стержень.

```vb
Public Function PositionSourceID(ByVal index As Long) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:52

<a id="api-crebarruleexecutor-positionphysicalowner"></a>
### CRebarRuleExecutor.PositionPhysicalOwner

Индекс физического владельца в итоговом CRebarLayout; объединённые позиции имеют одного владельца.

```vb
Public Function PositionPhysicalOwner(ByVal index As Long) As Long
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:53

<a id="api-crebarruleexecutor-ruletrace"></a>
### CRebarRuleExecutor.RuleTrace

Текстовый отчёт конкретного правила с фактическими counts и шагами/хордами.

```vb
Public Function RuleTrace(ByVal ruleID As String) As String
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:57

<a id="api-crebarruleexecutor-rulestepbounds"></a>
### CRebarRuleExecutor.RuleStepBounds

Возвращает число интервалов, min/max станционных шагов и хорд, мм; замыкание кольца учитывается.

```vb
Public Sub RuleStepBounds(ByVal ruleID As String, ByRef stepCount As Long, ByRef minStep As Double, ByRef maxStep As Double, ByRef minChord As Double, ByRef maxChord As Double)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| ruleID | String; ByVal; обязательный | Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила. |
| stepCount | Long; ByRef; обязательный | Целый неотрицательный счётчик; для ByCount применяется отдельно к каждому пути. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| minStep | Double; ByRef; обязательный | Длина/станция/граница minStep, внутренние мм; направление и пределы задаёт выбранный запрос. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| maxStep | Double; ByRef; обязательный | Длина/станция/граница maxStep, внутренние мм; направление и пределы задаёт выбранный запрос. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| minChord | Double; ByRef; обязательный | Длина/станция/граница minChord, внутренние мм; направление и пределы задаёт выбранный запрос. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| maxChord | Double; ByRef; обязательный | Длина/станция/граница maxChord, внутренние мм; направление и пределы задаёт выбранный запрос. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |

Жизненный цикл: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:72

<a id="api-crebarruleexecutor-diagnosticstext"></a>
### CRebarRuleExecutor.DiagnosticsText

Полный текстовый отчёт счётчиков и всех RuleTrace; численные длины в мм независимо от output units.

```vb
Public Function DiagnosticsText() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:108

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

Результат: CRebarLayout. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение происхождения/шагов после исполнения; позиции не являются независимой копией mesh.

Ошибки/исправление: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.

Source: src/Section/CRebarRuleExecutor.cls:119

## CUnitSystem

InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

<a id="api-cunitsystem-initializedefaults"></a>
### CUnitSystem.InitializeDefaults

Задаёт программные defaults: input mm/mm²/N/N*mm/MPa/1/mm и Tension, output m/m²/N/N*mm/MPa/1/mm. Не заменяет обязательный Config ввод.

```vb
Public Sub InitializeDefaults()
```

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:44

<a id="api-cunitsystem-loadfromsettings"></a>
### CUnitSystem.LoadFromSettings

Загружает и проверяет все активные единицы и знаки через общий settings reader.

```vb
Public Sub LoadFromSettings(ByVal settings As CSystemSettingsReader)
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| settings | CSystemSettingsReader; ByVal; обязательный | CSystemSettingsReader, уже загрузивший Config. |

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:59

<a id="api-cunitsystem-inputlengthtointernal"></a>
### CUnitSystem.InputLengthToInternal

Переводит InputLengthToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function InputLengthToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:82

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

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:88

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

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:95

<a id="api-cunitsystem-inputareatointernal"></a>
### CUnitSystem.InputAreaToInternal

Переводит InputAreaToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function InputAreaToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:101

<a id="api-cunitsystem-internalareatooutput"></a>
### CUnitSystem.InternalAreaToOutput

Переводит InternalAreaToOutput из внутренних единиц в подготовленные output units; сохранённая единица при наличии имеет явный смысл.

```vb
Public Function InternalAreaToOutput(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:106

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

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:112

<a id="api-cunitsystem-inputfourthpowerlengthtointernal"></a>
### CUnitSystem.InputFourthPowerLengthToInternal

Переводит InputFourthPowerLengthToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function InputFourthPowerLengthToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:118

<a id="api-cunitsystem-internalfourthpowerlengthtooutput"></a>
### CUnitSystem.InternalFourthPowerLengthToOutput

Переводит InternalFourthPowerLengthToOutput из внутренних единиц в подготовленные output units; сохранённая единица при наличии имеет явный смысл.

```vb
Public Function InternalFourthPowerLengthToOutput(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:123

<a id="api-cunitsystem-outputfourthpowerlengthtointernal"></a>
### CUnitSystem.OutputFourthPowerLengthToInternal

Переводит OutputFourthPowerLengthToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function OutputFourthPowerLengthToInternal(ByVal value As Double, Optional ByVal savedUnit As String = vbNullString) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |
| savedUnit | String; ByVal; Optional; default=vbNullString | Единица сохранённого Results; пустая использует текущую output-единицу CUnitSystem. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:129

<a id="api-cunitsystem-inputforcetointernal"></a>
### CUnitSystem.InputForceToInternal

Переводит InputForceToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function InputForceToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:135

<a id="api-cunitsystem-internalforcetooutput"></a>
### CUnitSystem.InternalForceToOutput

Переводит InternalForceToOutput из внутренних единиц в подготовленные output units; сохранённая единица при наличии имеет явный смысл.

```vb
Public Function InternalForceToOutput(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:140

<a id="api-cunitsystem-internalforcelengthsquaredtooutput"></a>
### CUnitSystem.InternalForceLengthSquaredToOutput

Переводит InternalForceLengthSquaredToOutput из внутренних единиц в подготовленные output units; сохранённая единица при наличии имеет явный смысл.

```vb
Public Function InternalForceLengthSquaredToOutput(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:147

<a id="api-cunitsystem-inputmomentmxtointernal"></a>
### CUnitSystem.InputMomentMxToInternal

Переводит InputMomentMxToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function InputMomentMxToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:152

<a id="api-cunitsystem-inputmomentmytointernal"></a>
### CUnitSystem.InputMomentMyToInternal

Переводит InputMomentMyToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function InputMomentMyToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:157

<a id="api-cunitsystem-inputmomentperlengthtointernal"></a>
### CUnitSystem.InputMomentPerLengthToInternal

Переводит InputMomentPerLengthToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function InputMomentPerLengthToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:164

<a id="api-cunitsystem-internalmomentmxtooutput"></a>
### CUnitSystem.InternalMomentMxToOutput

Переводит InternalMomentMxToOutput из внутренних единиц в подготовленные output units; сохранённая единица при наличии имеет явный смысл.

```vb
Public Function InternalMomentMxToOutput(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:169

<a id="api-cunitsystem-internalmomentmytooutput"></a>
### CUnitSystem.InternalMomentMyToOutput

Переводит InternalMomentMyToOutput из внутренних единиц в подготовленные output units; сохранённая единица при наличии имеет явный смысл.

```vb
Public Function InternalMomentMyToOutput(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:174

<a id="api-cunitsystem-internalmomentmagnitudetooutput"></a>
### CUnitSystem.InternalMomentMagnitudeToOutput

Переводит InternalMomentMagnitudeToOutput из внутренних единиц в подготовленные output units; сохранённая единица при наличии имеет явный смысл.

```vb
Public Function InternalMomentMagnitudeToOutput(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:179

<a id="api-cunitsystem-inputstresstointernal"></a>
### CUnitSystem.InputStressToInternal

Переводит InputStressToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function InputStressToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:184

<a id="api-cunitsystem-internalstresstooutput"></a>
### CUnitSystem.InternalStressToOutput

Переводит InternalStressToOutput из внутренних единиц в подготовленные output units; сохранённая единица при наличии имеет явный смысл.

```vb
Public Function InternalStressToOutput(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:189

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

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:195

<a id="api-cunitsystem-inputcurvaturetointernal"></a>
### CUnitSystem.InputCurvatureToInternal

Переводит InputCurvatureToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function InputCurvatureToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:201

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

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:207

<a id="api-cunitsystem-outputcurvaturetointernal"></a>
### CUnitSystem.OutputCurvatureToInternal

Переводит OutputCurvatureToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой конвенцией знаков для N/Mx/My.

```vb
Public Function OutputCurvatureToInternal(ByVal value As Double) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| value | Double; ByVal; обязательный | Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный перевод. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:213

<a id="api-cunitsystem-outputlengthunit"></a>
### CUnitSystem.OutputLengthUnit

Возвращает текущий output selector LengthUnit, без преобразования или округления данных.

```vb
Public Property Get OutputLengthUnit() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:217

<a id="api-cunitsystem-outputareaunit"></a>
### CUnitSystem.OutputAreaUnit

Возвращает текущий output selector AreaUnit, без преобразования или округления данных.

```vb
Public Property Get OutputAreaUnit() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:221

<a id="api-cunitsystem-outputmomentunit"></a>
### CUnitSystem.OutputMomentUnit

Возвращает текущий output selector MomentUnit, без преобразования или округления данных.

```vb
Public Property Get OutputMomentUnit() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:225

<a id="api-cunitsystem-inputmomentperlengthunit"></a>
### CUnitSystem.InputMomentPerLengthUnit

Возвращает единицу пользовательского момента на длину по input force/moment/length settings.

```vb
Public Property Get InputMomentPerLengthUnit() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:230

<a id="api-cunitsystem-outputforceunit"></a>
### CUnitSystem.OutputForceUnit

Возвращает текущий output selector ForceUnit, без преобразования или округления данных.

```vb
Public Property Get OutputForceUnit() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:234

<a id="api-cunitsystem-outputforcelengthsquaredunit"></a>
### CUnitSystem.OutputForceLengthSquaredUnit

Возвращает текущий output selector ForceLengthSquaredUnit, без преобразования или округления данных.

```vb
Public Property Get OutputForceLengthSquaredUnit() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:239

<a id="api-cunitsystem-outputstressunit"></a>
### CUnitSystem.OutputStressUnit

Возвращает текущий output selector StressUnit, без преобразования или округления данных.

```vb
Public Property Get OutputStressUnit() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:243

<a id="api-cunitsystem-outputcurvatureunit"></a>
### CUnitSystem.OutputCurvatureUnit

Возвращает текущий output selector CurvatureUnit, без преобразования или округления данных.

```vb
Public Property Get OutputCurvatureUnit() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:247

<a id="api-cunitsystem-outputsignn"></a>
### CUnitSystem.OutputSignN

Возвращает текущий output selector SignN, без преобразования или округления данных.

```vb
Public Property Get OutputSignN() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:251

<a id="api-cunitsystem-outputsignmx"></a>
### CUnitSystem.OutputSignMx

Возвращает текущий output selector SignMx, без преобразования или округления данных.

```vb
Public Property Get OutputSignMx() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:255

<a id="api-cunitsystem-outputsignmy"></a>
### CUnitSystem.OutputSignMy

Возвращает текущий output selector SignMy, без преобразования или округления данных.

```vb
Public Property Get OutputSignMy() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.

Ошибки/исправление: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется default. Приводите единицы через CUnitSystem один раз.

Source: src/Excel/CUnitSystem.cls:259

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

<a id="api-cconcreteregion-segmentcount"></a>
### CConcreteRegion.SegmentCount

Число аналитических line/arc сегментов.

```vb
Public Property Get SegmentCount() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:93

<a id="api-cconcreteregion-boundaryprobes"></a>
### CConcreteRegion.BoundaryProbes

Копия сохранённых пассивных пар P/Q; Empty допустим, контуры по ним не восстанавливаются.

```vb
Public Property Get BoundaryProbes() As Variant
```

Результат: Variant. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:99

<a id="api-cconcreteregion-boundaryprobecount"></a>
### CConcreteRegion.BoundaryProbeCount

Число сохранённых диагностических проб P/Q.

```vb
Public Property Get BoundaryProbeCount() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:102

<a id="api-cconcreteregion-loopcount"></a>
### CConcreteRegion.LoopCount

Число подготовленных материальных loops.

```vb
Public Property Get LoopCount() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:105

<a id="api-cconcreteregion-loopisopening"></a>
### CConcreteRegion.LoopIsOpening

Читает классификацию конкретной loop по подготовленной топологии.

```vb
Public Property Get LoopIsOpening(ByVal index As Long) As Boolean
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| index | Long; ByVal; обязательный | 1-based индекс в пределах соответствующего готового контейнера. |

Результат: Boolean. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:111

<a id="api-cconcreteregion-area"></a>
### CConcreteRegion.Area

Аналитическая площадь материальной области, мм²; не механическая A сетки.

```vb
Public Property Get Area() As Double
```

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:115

<a id="api-cconcreteregion-source"></a>
### CConcreteRegion.Source

Скалярное происхождение подготовленной области.

```vb
Public Property Get Source() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:118

<a id="api-cconcreteregion-sectionrevision"></a>
### CConcreteRegion.SectionRevision

Ревизия исходной модели, к которой относится область.

```vb
Public Property Get SectionRevision() As Long
```

Результат: Long. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:121

<a id="api-cconcreteregion-ownerid"></a>
### CConcreteRegion.OwnerID

Инженерный владелец области; назначается coordinator, не геометрической классификацией.

```vb
Public Property Get OwnerID() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:124

<a id="api-cconcreteregion-standard"></a>
### CConcreteRegion.Standard

Скалярная нормативная принадлежность инженерного владельца; geometry не вычисляет нормативные коэффициенты.

```vb
Public Property Get Standard() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:127

<a id="api-cconcreteregion-anchorid"></a>
### CConcreteRegion.AnchorID

Скалярный ID сохранённой инженерной области.

```vb
Public Property Get AnchorID() As String
```

Результат: String. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:130

<a id="api-cconcreteregion-clone"></a>
### CConcreteRegion.Clone

Копирует готовую область и metadata; временные массивы следующей операции не меняют исходный snapshot.

```vb
Public Function Clone() As CConcreteRegion
```

Результат: CConcreteRegion. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.

Ошибки/исправление: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot; пустая область и Nothing различаются.

Source: src/Section/CConcreteRegion.cls:136

## CSectionGeometryQuery

Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

<a id="api-csectiongeometryquery-lastdiagnostic"></a>
### CSectionGeometryQuery.LastDiagnostic

Последняя машинная диагностика; Nothing при отсутствии записанного события.

```vb
Public Property Get LastDiagnostic() As CGeometryDiagnostic
```

Результат: CGeometryDiagnostic. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:31

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

Source: src/Section/CSectionGeometryQuery.cls:36

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
| distance | Double; ByVal; обязательный | Длина/станция/граница distance, внутренние мм; направление и пределы задаёт выбранный запрос. |

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:48

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

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:116

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

Source: src/Section/CSectionGeometryQuery.cls:150

<a id="api-csectiongeometryquery-concretedomain"></a>
### CSectionGeometryQuery.ConcreteDomain

Возвращает clone области текущей query-session; не подменяет A/I механической сетки.

```vb
Public Function ConcreteDomain() As CConcreteRegion
```

Результат: CConcreteRegion. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:177

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
| minimum | Double; ByRef; обязательный | Длина/станция/граница minimum, внутренние мм; направление и пределы задаёт выбранный запрос. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| maximum | Double; ByRef; обязательный | Длина/станция/граница maximum, внутренние мм; направление и пределы задаёт выбранный запрос. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:185

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
| openingCount | Long; ByRef; обязательный | Целый неотрицательный счётчик; для ByCount применяется отдельно к каждому пути. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:246

<a id="api-csectiongeometryquery-polygonregion"></a>
### CSectionGeometryQuery.PolygonRegion

Готовит материальную область по замкнутому многоугольнику и проверяет топологию.

```vb
Public Function PolygonRegion(ByVal coordinates As Variant) As CConcreteRegion
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| coordinates | Variant; ByVal; обязательный | Массив координат вершин замкнутого многоугольника, мм. |

Результат: CConcreteRegion. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:272

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

Результат: CConcreteRegion. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:287

<a id="api-csectiongeometryquery-boundaryregion"></a>
### CSectionGeometryQuery.BoundaryRegion

Готовит область по аналитическим рёбрам, сшивает/проверяет замкнутые loops и роли.

```vb
Public Function BoundaryRegion(ByRef edges() As TRegionEdge, ByVal count As Long, Optional ByVal source As String = "ValidatedContour") As CConcreteRegion
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| edges | TRegionEdge(); ByRef; обязательный | Массив TRegionEdge аналитических рёбер с loop-данными, мм/радианы. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется. |
| count | Long; ByVal; обязательный | Целый неотрицательный счётчик; для ByCount применяется отдельно к каждому пути. |
| source | String; ByVal; Optional; default="ValidatedContour" | Подготовленный объект/контейнер выбранного запроса, без повторного чтения Excel или назначения типа. |

Результат: CConcreteRegion. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:298

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

Результат: Boolean. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:382

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

Результат: CConcreteRegion. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:394

<a id="api-csectiongeometryquery-mergedisjointregions"></a>
### CSectionGeometryQuery.MergeDisjointRegions

Объединяет непересекающиеся области; не является универсальным Union перекрывающихся фигур.

```vb
Public Function MergeDisjointRegions(ByVal regions As Collection) As CConcreteRegion
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| regions | Collection; ByVal; обязательный | Collection готовых непересекающихся областей для MergeDisjointRegions. |

Результат: CConcreteRegion. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:450

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

Результат: CConcreteRegion. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:471

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

Результат: CConcreteRegion. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:535

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

Результат: Boolean. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:554

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
| bound | Double; ByVal; обязательный | Длина/станция/граница bound, внутренние мм; направление и пределы задаёт выбранный запрос. |

Результат: CConcreteRegion. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:572

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
| nMin | Double; ByVal; обязательный | Длина/станция/граница nMin, внутренние мм; направление и пределы задаёт выбранный запрос. |
| nMax | Double; ByVal; обязательный | Длина/станция/граница nMax, внутренние мм; направление и пределы задаёт выбранный запрос. |
| tMin | Double; ByVal; обязательный | Длина/станция/граница tMin, внутренние мм; направление и пределы задаёт выбранный запрос. |
| tMax | Double; ByVal; обязательный | Длина/станция/граница tMax, внутренние мм; направление и пределы задаёт выбранный запрос. |

Результат: CConcreteRegion. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:644

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

Результат: CConcreteRegion. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:661

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

Результат: Boolean. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:820

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

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:830

<a id="api-csectiongeometryquery-roundedpolygon"></a>
### CSectionGeometryQuery.RoundedPolygon

Создаёт согласованную границу по вершинам/радиусам/ID. Edge.ID начинается у указанной вершины, Corner.ID существует только при R>0.

```vb
Public Function RoundedPolygon(ByVal pathID As String, ByVal vertices As Variant, ByVal radii As Variant, ByVal vertexIDs As Variant) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| vertices | Variant; ByVal; обязательный | Двумерные вершины: Array(Array(x,y),...); координаты во внутренних мм. |
| radii | Variant; ByVal; обязательный | Массив неотрицательных радиусов вершин в мм, длина равна vertices; clamp бетонного радиуса отсутствует. |
| vertexIDs | Variant; ByVal; обязательный | Массив уникальных смысловых ID вершин; определяет Edge.ID/Corner.ID. |

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:862

<a id="api-csectiongeometryquery-offsetroundedpolygon"></a>
### CSectionGeometryQuery.OffsetRoundedPolygon

Смещает весь скруглённый многоугольник с согласованными стыками. Положительное расстояние внутрь площади; отрицательное наружу, из Opening в бетон.

```vb
Public Function OffsetRoundedPolygon(ByVal pathID As String, ByVal vertices As Variant, ByVal radii As Variant, ByVal vertexIDs As Variant, ByVal distance As Double) As CGeometryPath
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| pathID | String; ByVal; обязательный | Непустой уникальный ID пути в определении. |
| vertices | Variant; ByVal; обязательный | Двумерные вершины: Array(Array(x,y),...); координаты во внутренних мм. |
| radii | Variant; ByVal; обязательный | Массив неотрицательных радиусов вершин в мм, длина равна vertices; clamp бетонного радиуса отсутствует. |
| vertexIDs | Variant; ByVal; обязательный | Массив уникальных смысловых ID вершин; определяет Edge.ID/Corner.ID. |
| distance | Double; ByVal; обязательный | Signed offset, мм: положительно внутрь площади, отрицательно наружу; это не left-offset обхода. |

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:932

<a id="api-csectiongeometryquery-regionfrompath"></a>
### CSectionGeometryQuery.RegionFromPath

Готовит область по одному замкнутому аналитическому пути; открытый путь не становится бетоном.

```vb
Public Function RegionFromPath(ByVal path As CGeometryPath) As CConcreteRegion
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| path | CGeometryPath; ByVal; обязательный | Существующий непрерывный опубликованный CGeometryPath. |

Результат: CConcreteRegion. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:991

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

Результат: CGeometryPath. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:1003

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

Результат: Variant. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:1174

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
| seedPlane | Double; ByVal; обязательный | Подготовленный объект/контейнер выбранного запроса, без повторного чтения Excel или назначения типа. |
| tMin | Double; ByVal; обязательный | Длина/станция/граница tMin, внутренние мм; направление и пределы задаёт выбранный запрос. |
| tMax | Double; ByVal; обязательный | Длина/станция/граница tMax, внутренние мм; направление и пределы задаёт выбранный запрос. |
| probeCount | Long; ByVal; обязательный | Целый неотрицательный счётчик; для ByCount применяется отдельно к каждому пути. |

Результат: Variant. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:1235

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

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:1267

<a id="api-csectiongeometryquery-regionfromcontours"></a>
### CSectionGeometryQuery.RegionFromContours

Готовит область из единственных достоверных CSectionContours Outer/Opening.

```vb
Public Function RegionFromContours(ByVal contours As CSectionContours) As CConcreteRegion
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| contours | CSectionContours; ByVal; обязательный | CSectionContours с достоверными аналитическими Outer/Opening. |

Результат: CConcreteRegion. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие операции над готовыми областями/путями; Initialize нужен для query-session по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.

Ошибки/исправление: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.

Source: src/Section/CSectionGeometryQuery.cls:1406

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

Source: src/Section/CSectionGeometryQuery.cls:1426

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

Source: src/Common/modSectionRecipeTypes.bas:11

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

Source: src/Common/modSectionRecipeTypes.bas:19

<a id="api-modsectionrecipetypes-erebardistributionscope"></a>
### modSectionRecipeTypes.ERebarDistributionScope

Только PerPath: количество/шаг/setbacks отдельно для каждого пути группы.

```vb
Public Enum ERebarDistributionScope
```

Поля/селекторы по source:

```vb
PerPath = 1
```

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:29

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

Source: src/Common/modSectionRecipeTypes.bas:33

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

Source: src/Common/modSectionRecipeTypes.bas:39

<a id="api-modsectionrecipetypes-erebarpositionorder"></a>
### modSectionRecipeTypes.ERebarPositionOrder

AlongPath или FromPathMiddle; геометрические станции/исходные ordinals сохраняются.

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

Source: src/Common/modSectionRecipeTypes.bas:44

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

Source: src/Common/modSectionRecipeTypes.bas:49

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

Source: src/Common/modSectionRecipeTypes.bas:55

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

Source: src/Common/modSectionRecipeTypes.bas:62

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
```

Жизненный цикл: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации заполняет фасад; автор не реализует координатные циклы баров через UDT.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modSectionRecipeTypes.bas:67

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

Source: src/Common/modSectionRecipeTypes.bas:77

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

Source: src/Common/modSectionRecipeTypes.bas:115

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

Source: src/Common/modSectionRecipeTypes.bas:137

<a id="api-modsectionrecipetypes-trebarposition"></a>
### modSectionRecipeTypes.TRebarPosition

Скалярное происхождение заявленной позиции до объединения: координаты мм, касательная, станция, rule/root ordinals и ID.

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

Source: src/Common/modSectionRecipeTypes.bas:144

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

Source: src/Common/modSectionRecipeTypes.bas:159

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

Source: src/Common/modSectionRecipeTypes.bas:166

## modGeometryTypes

Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

<a id="api-modgeometrytypes-geom_pi"></a>
### modGeometryTypes.GEOM_PI

Число pi как сумма сохраняемых VBE литералов, без изменения Double при компиляции.

```vb
Public Const GEOM_PI As Double = 3.14159265358979 + 3.10862446895044E-15
```

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modGeometryTypes.bas:14

<a id="api-modgeometrytypes-geom_tolerance"></a>
### modGeometryTypes.GEOM_TOLERANCE

Геометрический допуск совпадений исходных координат, мм; не округление ключей результатов.

```vb
Public Const GEOM_TOLERANCE As Double = 0.000000001
```

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modGeometryTypes.bas:15

<a id="api-modgeometrytypes-geom_max_rebar_count"></a>
### modGeometryTypes.GEOM_MAX_REBAR_COUNT

Предел представимости счётчика позиций, не обещание доступности памяти для такого массива.

```vb
Public Const GEOM_MAX_REBAR_COUNT As Double = 2147483646# ' Long-счетчик должен сохранять представимость следующего шага обхода.
```

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modGeometryTypes.bas:16

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

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modGeometryTypes.bas:70

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

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modGeometryTypes.bas:80

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

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modGeometryTypes.bas:93

<a id="api-modgeometrytypes-rebarrequestedpositioncount"></a>
### modGeometryTypes.RebarRequestedPositionCount

Проверяет запрошенное число позиций основного/дополнительных рядов до выделения массива.

```vb
Public Function RebarRequestedPositionCount(ByVal baseCount As Long, ByVal baseDiameter As Double, ByVal row2Diameter As Double, ByVal row3Diameter As Double, ByVal row2Binding As String, ByVal row3Binding As String) As Double
```

| Параметр | Тип / передача / default | Смысл и единицы |
|---|---|---|
| baseCount | Long; ByVal; обязательный | Целый неотрицательный счётчик; для ByCount применяется отдельно к каждому пути. |
| baseDiameter | Double; ByVal; обязательный | Диаметр соответствующего ряда, мм≥0; 0 выключает дополнительный ряд. |
| row2Diameter | Double; ByVal; обязательный | Диаметр соответствующего ряда, мм≥0; 0 выключает дополнительный ряд. |
| row3Diameter | Double; ByVal; обязательный | Диаметр соответствующего ряда, мм≥0; 0 выключает дополнительный ряд. |
| row2Binding | String; ByVal; обязательный | Принятый selector связи соответствующего ряда; EachBar/EverySecondBar. |
| row3Binding | String; ByVal; обязательный | Принятый selector связи соответствующего ряда; EachBar/EverySecondBar. |

Результат: Double. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modGeometryTypes.bas:120

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

Результат: CSectionModel. Смысл результата и дальнейшее использование описаны выше; возвращаемый sealed snapshot не изменяется этим чтением.

Жизненный цикл: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип сечения и не обращаются к Excel.

Ошибки/исправление: Нарушенный контракт вызывающей операции даёт адресную ошибку; данные и selectors проверяются общим источником/фасадом.

Source: src/Common/modGeometryTypes.bas:142

