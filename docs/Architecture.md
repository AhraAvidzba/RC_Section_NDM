# Architecture

Дата актуализации: 2026-09-21.

## Общая Схема

```text
Config + rngLoadCombinations
  -> CSystemSettingsReader
  -> CUnitSystem
  -> CLoadCombinationReader
  -> BuildWorkbookSectionModel
       -> Generated: CSectionTypeRegistry + ISectionGeometry + builders
       -> AutoCAD: CAutoCADSectionModelImporter
  -> CSectionModel
  -> CBatchSectionCalculator
       -> CSectionLoadState / CMomentZeroFilter
       -> CStateSolutionRunner -> CSectionSolver
       -> CCapacitySolver
       -> CCrackWidthCalculator
       -> CStabilityCalculator
  -> Results writers
  -> Results sheet
  -> Excel plot / AutoCAD export
```

Расчетное ядро работает с `CSectionModel`, материалами и настройками. Источник геометрии для решателей не важен: встроенный генератор и AutoCAD-import должны приводить данные к одной модели.

`CUnitSystem` является границей между пользовательским интерфейсом и расчетным ядром. Он читает `rngUnitSettings` и `rngSignConventionSettings`, переводит входные значения в фиксированные внутренние единицы/знаки и переводит результаты обратно в выбранный пользователем формат вывода.

Внутренние единицы:

```text
length    = mm
area      = mm2
force     = N
moment    = N*mm
stress    = MPa
curvature = 1/mm
+N        = tension
+Mx       = +Y tension
+My       = +X tension
```

Классы расчетного ядра не выполняют пользовательские пересчеты единиц и знаков самостоятельно.

## Источники Геометрии

Настройка `Geometry.Source` выбирает источник расчетной модели.

| Значение | Поведение |
|---|---|
| `Generated` | Сетка и арматура строятся встроенными генераторами по `Geometry.Type`. |
| `AutoCAD` | Импортируются AutoCAD `Region` из активного чертежа на слоях `AutoCAD.Import.ConcreteLayer` и `AutoCAD.Import.RebarLayer`. |

При AutoCAD-импорте единицы чертежа считаются миллиметрами. Импортер передает фактические данные Region: площадь, центр, локальные центральные моменты инерции, произведение инерции и доступные подсказки ориентации. Он не выбирает материалы и не восстанавливает тип сечения.

## Расширение Типов Сечений

Встроенные геометрии собираются через одну точку выбора `CSectionTypeRegistry`:

```text
CSectionTypeRegistry
  -> CGeometry*
  -> C*RebarLayoutBuilder
  -> C*AnnotationBuilder
  -> CSectionModel
```

Поддерживаемые типы:

- `Circle`;
- `RectSet`;
- `RoundedRectangle`;
- `HollowRectangle`.

Каждый тип сечения отвечает за свою параметрическую геометрию, автоматическую арматуру и semantic-аннотации. Универсальные слои не должны содержать частных проверок вида “если Circle” или “если RectSet” после выбора типа в реестре.

## CSectionModel

`CSectionModel` является единым расчетным представлением сечения. Он хранит:

- бетонные элементы: `ID`, `SourceName`, `SourceHandle`, `X`, `Y`, `Area`, `MaterialID`, `GeometryInterpretationStatus`, `Width`, `Height`, `Rotation`, `LocalIx`, `LocalIy`, `LocalIxy`, `Comment`;
- арматурные элементы: `ID`, `SourceName`, `SourceHandle`, `X`, `Y`, `Area`, `Diameter`, `SteelClass`, `MaterialID`, `Comment`;
- источник модели и параметрический контур, если он известен.

Реальные `Area/LocalIx/LocalIy/LocalIxy` являются расчетными характеристиками. `Width/Height/Rotation` описывают только геометрическую оболочку элемента для границ, схемы, расстояний и AutoCAD.

Если реальные локальные инерции элемента отсутствуют, последний fallback выполняется в `CSectionModel`: элемент получает характеристики эквивалентного квадрата по площади. Если реальные инерции есть, оболочка никогда не подменяет их в механических расчетах.

Расчетные имена элементов создаются только внутри `CSectionModel`:

```text
бетон:    C1, C2, ...
арматура: R1, R2, ...
```

Имена генераторов и AutoCAD handles сохраняются только как трассировочные поля.

## AutoCAD Import

`CAutoCADSectionModelImporter` является адаптером внешнего источника данных.

Импортер:

- подключается к активному AutoCAD через COM;
- перебирает `ModelSpace`;
- принимает только объекты `AcDbRegion`;
- фильтрует бетон и арматуру по слоям из `Config`;
- игнорирует области меньше `AutoCAD.Import.MinArea`;
- для арматуры вычисляет эквивалентный диаметр из площади Region;
- для бетона сохраняет фактические `Area`, `CentroidX/Y`, `LocalIx`, `LocalIy`, `LocalIxy`;
- для изотропных элементов может получить подсказку угла по первой прямой грани, если направление из инерций неопределимо.

Интерпретация оболочки `Rectangle` / `Equivalent rectangle` / `Equivalent square` выполняется в `CSectionModel`.

## Встроенные Генераторы

Встроенный путь используется при `Geometry.Source = Generated`.

| Тип | Геометрия | Арматура | Аннотации |
|---|---|---|---|
| `Circle` | `CGeometryCircle` | `CCircleRebarLayoutBuilder` | `CCircleAnnotationBuilder` |
| `RectSet` | `CGeometryRectSet` | `CRectSetRebarLayoutBuilder` | `CRectSetAnnotationBuilder` |
| `RoundedRectangle` | `CGeometryRoundedRectangle` | `CRoundedRectRebarLayoutBuilder` | `CRoundedRectAnnotationBuilder` |
| `HollowRectangle` | `CGeometryHollowRectangle` | `CHollowRectRebarLayoutBuilder` | `CHollowRectAnnotationBuilder` |

`CFiberMeshBuilder` и `CRebarLayout` остаются внутренними объектами генераторов и не передаются в решатели.

## Нагрузки

Нагрузки читаются через `CLoadCombinationReader` и дальше интерпретируются централизованно. Перенос между точкой приложения нагрузки, центром тяжести бетонного сечения, центром приведенного сечения и главными осями выполняется в общем слое работы с нагрузками.

Фильтр практически нулевого момента использует настройку `Calculation.ZeroMomentPerDepth`:

```text
Mtol = ZeroMomentPerDepth * h
Abs(M) <= Mtol => M = 0
```

Для глобальных осей используется бетонный габарит по соответствующей оси, для главных осей - габарит по повернутой главной плоскости.

## Расчетное Ядро

`CSectionSolver`, `CCapacitySolver`, `CCrackWidthCalculator`, `CStabilityCalculator`, `CBatchSectionCalculator` и `CSectionPropertiesCalculator` работают с `CSectionModel`.

Расчетное ядро:

- не обращается к Excel-листам;
- не знает источник геометрии;
- решает только постановку `N + Mx + My`;
- одноосный изгиб считает как частный случай при `Mx = 0` или `My = 0`.

Поле деформаций:

```text
epsilon(x, y) = epsilon0 + kappaX * y + kappaY * x
```

Внутренние усилия:

```text
Nint  = sum(sigma_i * A_i)
Mxint = sum(sigma_i * A_i * y_i)
Myint = sum(sigma_i * A_i * x_i)
```

## Вывод

- `CBatchResultWriter` пишет компактную сводку в `rngBatchSummary`;
- `CStrengthSummaryWriter` пишет подробную таблицу НДС и несущей способности от `rngStrengthSummaryAnchor`;
- `CCrackSummaryWriter` пишет подробную таблицу трещин от `rngCrackSummaryAnchor`;
- `CStabilitySummaryWriter` пишет подробную таблицу устойчивости от `rngStabilitySummaryAnchor`;
- `CNDMResultsWriter` пишет расчетный snapshot: `rngNDMSectionGeometry`, `rngNDMElementResults`, `rngNDMSectionProperties`, `rngNDMSectionAnnotations`.

AutoCAD export и Excel-схема читают последний снимок `Results`. Они не запускают решатель, не читают исходную геометрию заново и не держат модель в памяти между макросами.
