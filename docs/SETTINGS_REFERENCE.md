# SETTINGS_REFERENCE

Дата актуализации: 2026-08-14.

## Источник Настроек

Все пользовательские настройки расчета находятся на листе `Config`.

`CSystemSettingsReader` читает:

1. `rngUnitSettings`;
2. `rngSignConventionSettings`;
3. общий диапазон `rngSystemSettings`;
4. активный геометрический диапазон по `Geometry.Type`;
5. параметры материалов `rngConcreteMaterialParameters` и `rngSteelMaterialParameters`;
6. настройки расчетных диаграмм `rngCalculationDiagramSettings`.

При `Geometry.Source = AutoCAD` геометрический диапазон по `Geometry.Type` может оставаться заполненным, но расчетная геометрия берется из AutoCAD `Region`, а не из встроенного генератора.

## Именованные Диапазоны

| Диапазон | Назначение |
|---|---|
| `rngUnitSettings` | Отдельная таблица единиц измерения `Quantity / INPUT / INTERNAL / OUTPUT`. `INPUT` и `OUTPUT` редактируются через выпадающие списки, `INTERNAL` является справочной фиксированной колонкой. |
| `rngSignConventionSettings` | Отдельная таблица пользовательских правил знаков для `+N`, `+Mx`, `+My`; применяется одновременно к вводу и выводу. |
| `rngSystemSettings` | Общие настройки: источник геометрии, сетка, материалы, решатели, несущая способность, трещины, AutoCAD. |
| `rngCircleGeometry` | Параметры круглого сечения и автоматической арматуры. |
| `rngRoundedRectangleGeometry` | Параметры прямоугольного сечения со скруглениями. |
| `rngLShapeGeometry` | Параметры Г-образного сечения и автоматической арматуры. |
| `rngConcreteMaterialParameters` | Параметры бетона для I/II ГПС: `Rb/Rbt`, `Rb,ser/Rbt,ser`, `Rb,mc2`, `Eb/Ebt` и расчетные деформационные точки. |
| `rngSteelMaterialParameters` | Параметры арматуры для I/II ГПС: `Rsc/Rs`, `Rsc,ser/Rs,ser`, `Esc/Es` и предельные деформации диаграмм. |
| `rngCalculationDiagramSettings` | Выбор расчетной диаграммы и режима растянутого бетона для `Strength`, `Mcrc` и `CrackedNDS`. |
| `rngLoadCombinations` | До 20 сочетаний нагрузок: `CalculationType = Group1` для расчета по первой группе предельных состояний, `CalculationType = Group2` для расчета трещин по второй группе. |
| `rngBatchSummary` | Сводка batch-расчета на `Results!A1:BC29`. |
| `rngNDMElementResults` | LC-зависимые результаты элементов на `Results!A32`: `RunID`, `LoadCase`, `ProfileId`, `StateType`, `ElementID`, `Strain`, `Stress`, `PhysicalState`. |
| `rngNDMSectionGeometry` | Неизменяемая расчетная геометрия snapshot на `Results!K32`: координаты, площадь, размеры/диаметр, материал и локальные характеристики. |
| `rngNDMSectionProperties` | Общие свойства всего сечения и LC-зависимые свойства уровня сечения на `Results!AB32`: Bounds, центр тяжести, главные оси, единицы output, named-state metadata, точка приложения нагрузки. |
| `rngNDMMaterialDiagrams` | Уникальный каталог фактических точек диаграмм материалов на `Results!AJ32`. Каждая диаграмма имеет `DiagramId`; named-state metadata в `rngNDMSectionProperties` ссылается на бетонную и арматурную диаграмму через `ConcreteDiagramId`/`RebarDiagramId`. |
| `rngNDMSectionAnnotations` | Сохраненные semantic-аннотации оформления на `Results!AW32`: размерные линии и групповые подписи арматуры. |
| `chtNDMSectionPlot` | ChartObject схемы сечения на листе `Расчет`; создается около столбца `AP` и дальше не пересоздается при обновлении. |

## Единицы И Знаки

Все пользовательские числовые значения перед расчетом переводятся в единую внутреннюю систему через `CUnitSystem`. Расчетное ядро получает только внутренние единицы и внутренние знаки.

| Quantity | INPUT по умолчанию | INTERNAL | OUTPUT по умолчанию |
|---|---|---|---|
| `Length` | `mm` | `mm` | `mm` |
| `Area` | `mm2` | `mm2` | `mm2` |
| `Force` | `tf` | `N` | `tf` |
| `Moment` | `tf*m` | `N*mm` | `tf*m` |
| `Stress` | `MPa` | `MPa` | `MPa` |
| `Curvature` | `1/mm` | `1/mm` | `1/mm` |

Поддерживаемые варианты списков:

| Quantity | INPUT / OUTPUT |
|---|---|
| `Length` | `mm`, `cm`, `m` |
| `Area` | `mm2`, `cm2`, `m2` |
| `Force` | `N`, `kN`, `tf` |
| `Moment` | `N*mm`, `kN*m`, `tf*m` |
| `Stress` | `Pa`, `kPa`, `MPa`, `kgf/cm2`, `tf/m2` |
| `Curvature` | `1/mm`, `1/m` |

Внутренние знаки:

| Величина | USER по умолчанию | INTERNAL |
|---|---|---|
| `+N` | `Compression` | `Tension` |
| `+Mx` | `+Y tension` | `+Y tension` |
| `+My` | `+X tension` | `+X tension` |

Колонка `Ед.` в таблицах входных настроек заполняется формулами Excel и автоматически подтягивает выбранную пользователем единицу `INPUT` из `rngUnitSettings`. Заголовки таблиц результатов и таблиц на листе `Results` записываются writer-ами уже в выбранных единицах `OUTPUT`.

Контрольные таблицы координат диаграмм на `Config` строятся формулами Excel только для проверки и визуального контроля. Расчетное ядро читает не эти таблицы, а параметры материалов и настройки расчетных режимов; готовые универсальные `CMaterialDiagram` создает `CMaterialModelProvider`.

## Геометрия

| Ключ | Значения / ед. | Назначение |
|---|---|---|
| `Geometry.Source` | `Generated`, `AutoCAD` | Источник расчетной модели. |
| `Geometry.Type` | `RoundedRectangle`, `Circle`, `LShape` | Тип встроенной геометрии при `Geometry.Source = Generated`. |
| `Mesh.Step` | мм | Базовый шаг квадратной бетонной сетки для встроенных генераторов. Один и тот же размер элемента используется по X и Y. |
| `Mesh.BoundarySubdivisions` | шт | Дробление граничных ячеек для встроенных генераторов; `1` означает быстрый режим без дополнительного дробления. |
| `Load.ReferenceOffsetX`, `Load.ReferenceOffsetY` | мм | Смещение точки приложения пользовательских нагрузок относительно центра тяжести бетонного сечения. |

При нулевых смещениях `Mx` и `My` из `rngLoadCombinations` считаются заданными относительно центра тяжести бетонного сечения без учета продольной арматуры.

## AutoCAD Import

| Ключ | По умолчанию | Назначение |
|---|---:|---|
| `AutoCAD.Import.ConcreteLayer` | `Concrete` | Слой бетонных `Region` для импорта. |
| `AutoCAD.Import.RebarLayer` | `Reinf` | Слой арматурных `Region` для импорта. |
| `AutoCAD.Import.MinArea` | `0.000001` мм2 | Минимальная площадь Region; меньшие области игнорируются. |

Правила импорта:

- импортируются только AutoCAD `Region`;
- единицы AutoCAD всегда считаются миллиметрами;
- материальная модель для расчета строится через `CMaterialModelProvider`; AutoCAD importer передает только геометрию и не выбирает диаграмму сам;
- коэффициент `phi2` при расчете трещин задается пользователем в `SLS.Crack.Phi2`;
- разные материалы арматуры в одном сечении не поддерживаются;
- при отсутствии AutoCAD, активного чертежа или нужных областей расчет останавливается с ошибкой ввода.

## AutoCAD Export

| Ключ | По умолчанию | Назначение |
|---|---:|---|
| `AutoCAD.Export.CombinationID` | `Worst` | Какое сочетание экспортировать из `Results`: `Worst` - определяющее сочетание из batch summary, либо конкретный `CombinationID` из `rngLoadCombinations`. Выпадающий список формируется динамически по таблице сочетаний. |
| `AutoCAD.Export.NeutralLineEnabled` | `Yes` | Выгружать нейтральную линию: `Yes` - выводить; `No` - не выводить. |
| `AutoCAD.Export.PrincipalAxesEnabled` | `Yes` | Выгружать главные центральные оси приведенного сечения: `Yes` - выводить; `No` - не выводить. |
| `AutoCAD.Export.LoadPointEnabled` | `Yes` | Выгружать точку приложения нагрузки: `Yes` - выводить; `No` - не выводить. |
| `AutoCAD.Export.ResultType` | `Stress` | Что экспортировать цветом и подписями: `Stress` или `Strain`. |
| `AutoCAD.Export.LabelMode` | `NamesAndValues` | `ValuesOnly` - только значение выбранного `ResultType`; `NamesAndValues` - имя элемента и значение. |
| `AutoCAD.Layer.Concrete` | `Concrete` | Слой областей бетона. |
| `AutoCAD.Layer.Rebar` | `Reinf` | Слой областей арматуры. |
| `AutoCAD.Layer.ConcreteTension` | `Anno_Concrete_Positive` | Слой подписей растянутого бетона. |
| `AutoCAD.Layer.ConcreteCompression` | `Anno_Concrete_Negative` | Слой подписей сжатого, нулевого и почти нулевого бетона. |
| `AutoCAD.Layer.RebarTension` | `Anno_Rebar_Positive` | Слой подписей растянутой арматуры. |
| `AutoCAD.Layer.RebarCompression` | `Anno_Rebar_Negative` | Слой подписей сжатой, нулевой и почти нулевой арматуры. |
| `AutoCAD.Color.ConcreteTension` | `9` | ColorIndex растянутых бетонных областей и подписей. |
| `AutoCAD.Color.ConcreteCompression` | `5` | ColorIndex сжатых бетонных областей и подписей. |
| `AutoCAD.Color.RebarTension` | `1` | ColorIndex растянутых стержней и подписей. |
| `AutoCAD.Color.RebarCompression` | `6` | ColorIndex сжатых стержней и подписей. |
| `AutoCAD.Color.Neutral` | `8` | ColorIndex нулевых/почти нулевых областей, подписей и нейтральной линии. |

Если слой экспорта отсутствует в активном чертеже, он создается автоматически.

## Выпадающие Списки

| Ключ | Значения |
|---|---|
| `Geometry.Source` | `Generated`, `AutoCAD` |
| `Geometry.Type` | `RoundedRectangle`, `Circle`, `LShape` |
| `Strength.ConcreteTensionMode` в `rngCalculationDiagramSettings` | `Ignore`, `UseDiagram` |
| `Calculation.Mode` | `DirectState`, `FullCapacity`, `CapacityOnly` |
| `Solver.Method` | `Newton`, `Secant` |
| `Capacity.SolutionStrategy` | `Auto`, `UltimateStrain`, `LoadMultiplier` |
| `Capacity.SearchMethod` | `Bisection`, `Brent`, `Secant` |
| `Solver.LineSearchEnabled` | `Yes`, `No` |
| `SLS.Crack.Phi3Mode` | `Auto`, `User` |
| `SLS.Crack.PsiMode` | `User`, `Auto` |
| `SLS.Crack.TensionZoneMode` | `Effective`, `FullTension` |
| `AutoCAD.Export.ResultType` | `Stress`, `Strain` |
| `AutoCAD.Export.LabelMode` | `ValuesOnly`, `NamesAndValues` |
| `AutoCAD.Export.NeutralLineEnabled` | `Yes`, `No` |
| `AutoCAD.Export.PrincipalAxesEnabled` | `Yes`, `No` |
| `AutoCAD.Export.LoadPointEnabled` | `Yes`, `No` |
| `AutoCAD.Export.CombinationID` | Динамический список: `Worst` + значения `CombinationID` из `rngLoadCombinations` |
| `Plot.Enabled` | `Yes`, `No` |
| `Plot.AutoUpdateAfterCalculation` | `Yes`, `No` |
| `Plot.LoadCase` | Динамический список: `Worst` + значения `CombinationID` из `rngLoadCombinations` |
| `Plot.ResultType` | `Stress`, `Strain` |
| `Plot.ResultGradient` | `Yes`, `No` |
| `Plot.ResultLabelsEnabled` | `Yes`, `No` |
| `Plot.ResultPrecision` | целое число знаков после запятой |
| `Plot.NeutralLineEnabled` | `Yes`, `No` |
| `Plot.PrincipalAxesEnabled` | `Yes`, `No` |
| `Plot.LoadApplicationPointEnabled` | `Yes`, `No` |
| `Plot.CentroidEnabled` | `Yes`, `No` |
| `Plot.Dimensions.Enabled` | `Yes`, `No` |
| `Plot.RebarLabels.Enabled` | `Yes`, `No` |
| `Plot.Dimensions.Placement` | `Outside`, `Inside` |
| `Plot.RebarLabels.Placement` | `Outside`, `Inside` |
| `Plot.Dimensions.Offset`, `Plot.RebarLabels.Offset` | мм сечения |
| `Plot.Dimensions.TextUnits`, `Plot.RebarLabels.TextUnits` | `mm`, `pt` |
| `Plot.Dimensions.TextHeight`, `Plot.RebarLabels.TextHeight` | по выбранному `TextUnits`: мм сечения или pt |
| `Plot.Dimensions.TextGap`, `Plot.RebarLabels.TextGap` | по выбранному `TextUnits`: мм сечения или pt |
| `Plot.RebarLabels.LineEnabled` | `Yes`, `No` |
| `Plot.Dimensions.ArrowType` | `Triangle`, `Stealth`, `Diamond`, `Oval`, `Open` |
| `Plot.Dimensions.ArrowSize` | `Small`, `Medium`, `Wide` |
| `Plot.LegendEnabled` | `Yes`, `No` |
| `rngLoadCombinations.CalculationType` | `Group1`, `Group2` |

Пустые или неподдерживаемые значения расчетных методов должны приводить к `InputError`, а не к скрытому переключению на значение по умолчанию.

