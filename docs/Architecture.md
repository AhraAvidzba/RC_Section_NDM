# Architecture

Дата актуализации: 2026-08-14.

## Общая Схема

```text
Config + rngLoadCombinations
  -> CSystemSettingsReader
  -> CUnitSystem
  -> CLoadCombinationReader
  -> BuildWorkbookSectionModel
       -> Generated: ISectionGeometry + CFiberMeshBuilder + CRebarLayout
       -> AutoCAD: CAutoCADSectionModelImporter
  -> CSectionModel
  -> CBatchSectionCalculator
       -> CStateSolutionRunner -> CSectionSolver
       -> CCapacityLoadPath
       -> CCapacitySolver
       -> CCrackWidthCalculator
  -> CBatchResultWriter / CNDMResultsWriter
  -> Results sheet
  -> AutoCAD export
```

Расчетное ядро работает только с `CSectionModel`, материалами и настройками. Источник геометрии для решателей не важен.

`CUnitSystem` является границей между пользовательским интерфейсом и расчетным ядром. Он читает `rngUnitSettings` и `rngSignConventionSettings`, переводит входные значения в фиксированные внутренние единицы/знаки и переводит результаты обратно в выбранный пользователем формат вывода.

Внутри расчетного ядра:

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

Классы `CLoadCase`, `CSectionModel`, `CSectionSolver`, `CCapacitySolver`, `CCrackWidthCalculator` не должны выполнять пользовательские пересчеты единиц и знаков самостоятельно.

## Источники Геометрии

Настройка `Geometry.Source` выбирает источник расчетной модели:

| Значение | Поведение |
|---|---|
| `Generated` | Сетка и арматура строятся встроенными генераторами по `Geometry.Type`. |
| `AutoCAD` | Импортируются только AutoCAD `Region` из активного чертежа на слоях `AutoCAD.Import.ConcreteLayer` и `AutoCAD.Import.RebarLayer`. |

При AutoCAD-импорте единицы чертежа считаются миллиметрами. Импортер передает в расчет только геометрию и не выбирает материальные диаграммы. Готовые модели бетона и арматуры выбирает `CMaterialModelProvider` по расчетному контексту, а импортированные стержни получают только техническую метку `Rebar`. Коэффициент `phi2` для раскрытия трещин задается отдельно в `SLS.Crack.Phi2`.

## Section Type Extension Contract

Built-in generated geometry is assembled through one registry point:

```text
CSectionTypeRegistry
  -> CGeometry*              (mathematical shape)
  -> C*RebarLayoutBuilder    (bars and semantic rebar groups)
  -> C*AnnotationBuilder     (contours, dimensions, rebar labels)
  -> CSectionModel
  -> CNDMResultsWriter
  -> rngNDMSectionGeometry / rngNDMElementResults / rngNDMSectionProperties / rngNDMSectionAnnotations
```

`CSectionTypeRegistry` is the only place where a generated section type is selected by
`Geometry.Type`. To add a new generated section, add the shape-specific classes, for example
`CGeometryTShape`, `CTShapeRebarBuilder`, `CTShapeAnnotationBuilder`, and register that
set in `CSectionTypeRegistry`. There is no separate provider class per section type.

Universal layers must not contain shape-specific checks after that point:

- `CNDMResultsWriter` serializes `CSectionModel.Annotations`; it does not infer Circle/LShape/RoundedRectangle from elements.
- `CSectionPlotDataReader` reads saved Results tables and does not call geometry builders.
- `CPlotAnnotationLayout` only converts semantic annotations from model coordinates into chart layout.
- `CSectionPlotter` only draws prepared series and annotation layouts.

Semantic annotation coordinates are model/internal coordinates when they are created by a
shape-specific annotation builder. The output writer converts them to the selected snapshot
output units when writing `rngNDMSectionAnnotations`.

## CSectionModel

`CSectionModel` является единым расчетным представлением сечения. Он хранит:

- бетонные элементы: `ID`, `SourceName`, `SourceHandle`, `X`, `Y`, `Area`, `MaterialID`, `GeometryInterpretationStatus`, `Width`, `Height`, `Rotation`, `LocalIx`, `LocalIy`, `LocalIxy`, `Comment`;
- арматурные элементы: `ID`, `SourceName`, `SourceHandle`, `X`, `Y`, `Area`, `Diameter`, `SteelClass`, `MaterialID`, `Comment`;
- `SourceType` для трассировки источника модели.

Расчетные имена элементов создаются только внутри `CSectionModel`:

```text
бетон:    C1, C2, ...
арматура: R1, R2, ...
```

Имена генераторов и AutoCAD handles сохраняются только как трассировочные поля `SourceName`/`SourceHandle`.

## AutoCAD Import

`CAutoCADSectionModelImporter` находится вне `CSectionModel` и является адаптером внешнего источника данных.

Импортёр:

- подключается к активному AutoCAD через COM;
- перебирает `ModelSpace`;
- принимает только объекты `AcDbRegion`;
- фильтрует бетон и арматуру по слоям из `Config`;
- игнорирует области меньше `AutoCAD.Import.MinArea`;
- для арматуры вычисляет эквивалентный диаметр из площади Region;
- для бетона сохраняет площадь, центр и, если AutoCAD отдаёт данные, центральные моменты инерции;
- для последующего экспорта строит эквивалентный прямоугольник по площади и моментам инерции.

Если AutoCAD закрыт, активного чертежа нет или нужные Region не найдены, расчет прерывается понятным сообщением.

## Встроенные Генераторы

Встроенный путь используется при `Geometry.Source = Generated`:

- `ISectionGeometry`, `CGeometryCircle`, `CGeometryRoundedRectangle`, `CGeometryLShape` описывают принадлежность точек бетонному сечению;
- `CFiberMeshBuilder` строит бетонные элементы;
- `CCircleRebarLayoutBuilder` и `CLShapeRebarLayoutBuilder` строят автоматическую арматуру;
- `CSectionModelBuilder.BuildFromGenerated` собирает `CSectionModel`.

`CFiberMeshBuilder` и `CRebarLayout` остаются внутренними объектами генераторов и не передаются в решатели.

## Расчетное Ядро

`CSectionSolver`, `CCapacitySolver`, `CCrackWidthCalculator`, `CBatchSectionCalculator` и `CSectionPropertiesCalculator` работают с `CSectionModel`.

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

- `CBatchResultWriter` пишет компактную сводку в `rngBatchSummary`, а `CStabilitySummaryWriter` пишет отдельную подробную таблицу устойчивости от `rngStabilitySummaryAnchor`;
- `CNDMResultsWriter` пишет согласованный snapshot последнего расчета на лист `Results`: `rngNDMSectionGeometry` с постоянной геометрией, `rngNDMElementResults` с LC-зависимыми `Strain/Stress/PhysicalState`, `rngNDMSectionProperties` с общими свойствами сечения и состоянием выбранных LC, `rngNDMSectionAnnotations` с сохраненными semantic-аннотациями;
- writer-ы получают `CUnitSystem` и выводят числовые результаты в выбранных `OUTPUT`-единицах и пользовательских знаках;
- контрольная таблица арматуры и формульный блок трещин на `Config` больше не выводятся;
- AutoCAD export читает данные из листа `Results`, поэтому не хранит последнюю модель в памяти и не запускает повторный AutoCAD-import при выгрузке.
- `UpdateSectionPlot` строит схему на листе `Расчет` только по сохраненному snapshot `Results`; смена `Plot.LoadCase` или `Plot.ResultType` не запускает расчет и не меняет `Results`.
- Старый численный блок на листе `Расчет` удален: governing LC больше не пересчитывается отдельной веткой, а весь пользовательский вывод берется из batch/snapshot.
