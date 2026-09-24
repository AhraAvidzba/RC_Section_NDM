# SETTINGS_REFERENCE

Дата актуализации: 2026-09-21.

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
| `rngCalculationProfiles` | Расчетные профили: какие ветки запускаются и какие модели материалов используются. |
| `rngCircleGeometry` | Параметры круглого сечения и автоматической арматуры. |
| `rngRoundedRectangleGeometry` | Параметры прямоугольного сечения со скруглениями. |
| `rngHollowRectangleGeometry` | Параметры пустотелого прямоугольника со скруглениями, внешней и внутренней арматурой. |
| `rngRectSetGeometry` | Параметры сечения из одного или двух прямоугольников и автоматической арматуры. |
| `rngConcreteMaterialParameters` | Параметры бетона для I/II ГПС: `Rb/Rbt`, `Rb,ser/Rbt,ser`, `Rb,mc2`, `Eb/Ebt` и расчетные деформационные точки. |
| `rngSteelMaterialParameters` | Параметры арматуры для I/II ГПС: `Rsc/Rs`, `Rsc,ser/Rs,ser`, `Esc/Es` и предельные деформации диаграмм. |
| `rngCalculationDiagramSettings` | Выбор расчетной диаграммы и режима растянутого бетона для `Strength`, `CrackInitiation` и `CrackedNDS`. |
| `rngLoadCombinations` | Сочетания нагрузок; фактическое число сочетаний равно числу строк внутри именованного диапазона. В шаблоне подготовлено 30 строк, минимально допустима 1 строка под заголовком. |
| `rngStabilityDurationLoads` | Постоянные или постоянные+длительные нагрузки для расчета устойчивости, сопоставленные с `CombinationID`. |
| `rngSP35Table721` | Пользовательская таблица коэффициентов СП 35 табл. 7.21; расчет интерполирует только по содержимому этой таблицы. |
| `rngBatchSummary` | Компактная сводка batch-расчета на `Results!A1`. |
| `rngStrengthSummaryAnchor` | Якорь первой строки данных подробной таблицы НДС и несущей способности по модели прочности на `Results!A48`; шапка таблицы формируется над якорем. |
| `rngCrackSummaryAnchor` | Якорь первой строки данных подробной таблицы нормальных и продольных трещин на `Results!A84`; шапка таблицы формируется над якорем. |
| `rngStabilitySummaryAnchor` | Якорь первой строки данных подробной таблицы устойчивости на `Results!A121`; шапка таблицы формируется над якорем. |
| `rngNDMElementResults` | LC-зависимые результаты элементов на `Results!A155`: `RunID`, `LoadCase`, `ProfileId`, `StateType`, `ElementID`, `Strain`, `Stress`, `PhysicalState`. |
| `rngNDMSectionGeometry` | Неизменяемая расчетная геометрия snapshot на `Results!L155`: координаты, площадь, размеры/диаметр, материал и локальные характеристики. |
| `rngNDMSectionProperties` | Общие свойства всего сечения и LC-зависимые свойства уровня сечения на `Results!AC155`: Bounds, центр тяжести, главные оси, единицы output, named-state metadata, точка приложения нагрузки. |
| `rngNDMMaterialDiagrams` | Уникальный каталог фактических точек диаграмм материалов на `Results!AK155`. Каждая диаграмма имеет `DiagramId`; named-state metadata в `rngNDMSectionProperties` ссылается на бетонную и арматурную диаграмму через `ConcreteDiagramId`/`RebarDiagramId`. |
| `rngNDMSectionAnnotations` | Сохраненные semantic-аннотации оформления на `Results!AX155`: размерные линии и групповые подписи арматуры. |
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
| `Geometry.Type` | `Circle`, `RectSet`, `RoundedRectangle`, `HollowRectangle` | Тип встроенной геометрии при `Geometry.Source = Generated`. |
| `Mesh.StepX` | мм | Шаг бетонной сетки по X для встроенных генераторов. |
| `Mesh.StepY` | мм | Шаг бетонной сетки по Y для встроенных генераторов. Если `Mesh.StepX = Mesh.StepY`, элементы квадратные; если значения отличаются, элементы прямоугольные. |
| `Mesh.BoundarySubdivisions` | шт | Дробление граничных ячеек для встроенных генераторов; `1` означает быстрый режим без дополнительного дробления. |
| `Load.ReferenceOffsetX`, `Load.ReferenceOffsetY` | мм | Смещение точки приложения пользовательских нагрузок относительно центра тяжести бетонного сечения. |

При нулевых смещениях `Mx` и `My` из `rngLoadCombinations` считаются заданными относительно центра тяжести бетонного сечения без учета продольной арматуры.

## Общие Расчетные Настройки

| Ключ | Значения / ед. | Назначение |
|---|---|---|
| `Calculation.ZeroMomentPerDepth` | Moment/Length в INPUT-единицах | Инженерный фильтр практически нулевого момента. Для каждой плоскости программа считает `Mtol = Calculation.ZeroMomentPerDepth · h`, где `h` - бетонный габарит сечения в этой плоскости. Если `Abs(M) <= Mtol`, момент принимается равным нулю до выбора расчетной ветки. Настройка убирает численный остаток после переносов и поворотов, но слишком большое значение может обнулить реальный малый момент или эксцентриситет. |

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
- importer читает фактические `Area`, центр и центральные `Ix/Iy/Ixy` бетонного `Region`;
- importer не использует `BoundingBox` для распознавания прямоугольника и не назначает `Width/Height`;
- `CSectionModel` строит прямоугольную оболочку элемента и выводит `GeometryInterpretationStatus`: `Rectangle`, `Equivalent rectangle` или `Equivalent square`;
- реальные `Area/Ix/Iy/Ixy` остаются расчетными характеристиками элемента, а `Width/Height/Rotation` используются только для границ, `h`, `a_s`, схемы и AutoCAD export;
- для арматурных `Region` диаметр и радиус стержня восстанавливаются по площади как эквивалентный круг;
- коэффициент `phi2` при расчете трещин задается пользователем в `SLS.Crack.Phi2`;
- разные материалы арматуры в одном сечении не поддерживаются;
- при отсутствии AutoCAD, активного чертежа или нужных областей расчет останавливается с ошибкой ввода.

## AutoCAD Export

| Ключ | По умолчанию | Назначение |
|---|---:|---|
| `AutoCAD.Export.CombinationID` | `Worst` | Какое сочетание экспортировать из `Results`: `Worst` - определяющее сочетание из batch summary, либо конкретный `CombinationID` из `rngLoadCombinations`. Выпадающий список формируется динамически по таблице сочетаний. |
| `AutoCAD.Export.NeutralLineEnabled` | `Yes` | Выгружать нейтральную линию: `Yes` - выводить; `No` - не выводить. |
| `AutoCAD.Export.PrincipalAxesMode` | `Transformed` | Какие главные центральные оси выгружать: `Transformed` - приведенного сечения, `Concrete` - бетонного сечения, `None` - не выводить. |
| `AutoCAD.Export.LoadPointEnabled` | `Yes` | Выгружать точку приложения нагрузки: `Yes` - выводить; `No` - не выводить. |
| `AutoCAD.Export.ContourEnabled` | `Yes` | Выгружать параметрический контур сечения, если он известен. Для импортированной сетки контур не восстанавливается. |
| `AutoCAD.Export.LabelMode` | `NamesAndValues` | `ValuesOnly` - только значение выбранной профильной величины `Visualization.Quantity`; `NamesAndValues` - имя элемента и значение. |
| `AutoCAD.Layer.Concrete` | `Concrete` | Слой областей бетона. |
| `AutoCAD.Layer.Rebar` | `Reinf` | Слой областей арматуры. |
| `AutoCAD.Layer.Contour` | `RC_NDM_Contour` | Слой параметрического контура сечения. Цвет назначается выводимой линии, а не свойствам слоя. |
| `AutoCAD.Layer.ConcreteTension` | `Anno_Concrete_Positive` | Слой подписей растянутого бетона. |
| `AutoCAD.Layer.ConcreteCompression` | `Anno_Concrete_Negative` | Слой подписей сжатого, нулевого и почти нулевого бетона. |
| `AutoCAD.Layer.RebarTension` | `Anno_Rebar_Positive` | Слой подписей растянутой арматуры. |
| `AutoCAD.Layer.RebarCompression` | `Anno_Rebar_Negative` | Слой подписей сжатой, нулевой и почти нулевой арматуры. |
| `AutoCAD.Color.ConcreteTension` | `9` | ColorIndex растянутых бетонных областей и подписей. |
| `AutoCAD.Color.ConcreteCompression` | `5` | ColorIndex сжатых бетонных областей и подписей. |
| `AutoCAD.Color.RebarTension` | `1` | ColorIndex растянутых стержней и подписей. |
| `AutoCAD.Color.RebarCompression` | `6` | ColorIndex сжатых стержней и подписей. |
| `AutoCAD.Color.Neutral` | `8` | ColorIndex нулевых/почти нулевых областей, подписей и нейтральной линии. |

Если слой экспорта отсутствует в активном чертеже, он создается автоматически. Если слой уже есть, программа использует его имя и не меняет свойства слоя; цвета назначаются создаваемым объектам.

## Выпадающие Списки

| Ключ | Значения |
|---|---|
| `Geometry.Source` | `Generated`, `AutoCAD` |
| `Geometry.Type` | `Circle`, `RectSet`, `RoundedRectangle`, `HollowRectangle` |
| `General.WorstCombinationCriterion` | `StrengthStrain`, `StrengthCapacity`, `Cracks`, `Stability` |
| `Strength.ConcreteTensionMode` в `rngCalculationDiagramSettings` | `Ignore`, `UseDiagram` |
| `Solver.Method` | `Newton`, `Secant` |
| `Capacity.SolutionStrategy` | `Auto`, `UltimateStrain`, `LoadMultiplier` |
| `Capacity.SearchMethod` | `Bisection`, `Brent`, `Secant` |
| `Solver.LineSearchEnabled` | `Yes`, `No` |
| `SLS.Crack.Phi3Mode` | `Auto`, `User` |
| `SLS.Crack.PsiMode` | `User`, `Auto` |
| `SLS.Crack.InitiationLoadPath` | `Auto`, `λ*Mxy`, `λ*N`, `λ*NMxy` |
| `SLS.Crack.InitiationSolutionStrategy` | `Auto`, `UltimateStrain`, `LoadMultiplier` |
| `SLS.Crack.TensionZoneMode` | `Effective`, `FullTension` |
| `SLS.Crack.CoverDistanceMode` | `NearestContour`, `GlobalExtreme` |
| `AutoCAD.Export.LabelMode` | `ValuesOnly`, `NamesAndValues` |
| `AutoCAD.Export.NeutralLineEnabled` | `Yes`, `No` |
| `AutoCAD.Export.PrincipalAxesMode` | `Transformed`, `Concrete`, `None` |
| `AutoCAD.Export.LoadPointEnabled` | `Yes`, `No` |
| `AutoCAD.Export.ContourEnabled` | `Yes`, `No` |
| `AutoCAD.Export.CombinationID` | Динамический список: `Worst` + значения `CombinationID` из `rngLoadCombinations` |
| `Plot.Enabled` | `Yes`, `No` |
| `Plot.AutoUpdateAfterCalculation` | `Yes`, `No` |
| `Plot.LoadCase` | Динамический список: `Worst` + значения `CombinationID` из `rngLoadCombinations` |
| `Plot.ResultGradient` | `Yes`, `No` |
| `Plot.ResultLabelsEnabled` | `Yes`, `No` |
| `Plot.NeutralLineEnabled` | `Yes`, `No` |
| `Plot.PrincipalAxesMode` | `Transformed`, `Concrete`, `None` |
| `Plot.LoadApplicationPointEnabled` | `Yes`, `No` |
| `Plot.ContourEnabled` | `Yes`, `No` |
| `Plot.AxisLabelsEnabled` | `Yes`, `No` |
| `Plot.AxisLabelsFontSize` | размер шрифта подписей `Ось X` и `Ось Y` в pt |
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
| `rngLoadCombinations.CapacityLoadPath` | `Auto`, `λ*Mx`, `λ*My`, `λ*Mxy`, `λ*N`, `λ*NMxy` |

Пустые или неподдерживаемые значения расчетных методов должны приводить к `InputError`, а не к скрытому переключению на значение по умолчанию.

