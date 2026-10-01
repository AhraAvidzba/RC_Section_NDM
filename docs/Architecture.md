# Architecture

Дата актуализации: 2026-10-01, Audit02 (приемка завершена в доступной среде).

Доказательства и ограничения: [итоговый отчет Audit02](NDM_Audit02_Final_Report.md).

Актуальные уточнения: [ТЗ Audit02](NDM_Audit02_Implementation_Spec_2026-10-01.md).
История решений хранится в v6/migration/progress; ограничения старых документов
«Extension только для DirectState» заменены единым `General.DiagramExtension`.

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
       -> CStateProvider -> CStateRepository / CStateSolutionRunner -> CSectionSolver
       -> CCapacityCalculator -> CLimitSearchCoordinator
       -> CCrackFormationCalculator -> CLimitSearchCoordinator
       -> CCrackWidthCalculator -> CCrackWidthFormulaCalculator
       -> CLongitudinalCrackCalculator
       -> CStabilityCalculator
  -> CCombinationResult (канонические typed results)
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

### State И Search

Обычное равновесие при конкретных `N/Mx/My` исполняется через общий
`CStateSolutionRunner` и `CSectionSolver`, включая LoadMultiplier-пробы.
`EvaluateStrainPlane` лишь оценивает заданную плоскость; она становится
подтвержденным НДС после проверки усилий `ConfirmEquilibrium` или успешного solve.
Runner сохраняет Object material API для линейных тестовых материалов.

`CStateProvider.GetOrSolve` получает конечное named-state. Reuse учитывает
нагрузки, material role/spec, разрешение продолжения в запросе и scoped контекст:
identity/revision геометрии и материалов, effective глобальный On/Off, допуски
приемки. Смена такого контекста инвалидирует repository. Warm-start, retry,
load steps и диагностика не создают новый физический ключ. Неуспешные попытки
сохраняют диагностику, но не являются reusable state и не блокируют новый solve.

В `CStateRepository` попадают только конечные named states: `StrengthState`,
`CapacityState`, `PreCrackState`, `PostCrackState`, `CrackedState`. Временные
lambda-точки остаются в локальном контексте search-session. Успешные повторно
используемые пробы и warm-start сохранены; failed probe не запрещает лучший retry.
Восстановление solver snapshot выполняется общим State-слоем без нового solve
и без требования живого `LastRunner`.

`CLoadMultiplierSearch` содержит единый bracket/refinement для обеих задач,
`CUltimateStrainSearch` общий Newton для физического предела и load-path residual,
`CLoadPathMath` математику `Offset + lambda * Base`. `ILimitSearchProblem` и
существующие Capacity/Crack adapters передают критерий и доменную оценку проб.
Generic Search не приводит объекты к инженерным калькуляторам и не назначает
пользовательский FAIL/BaseFail. Auto-смена физического пути принадлежит Formation.

`CLimitSearchResult` хранит независимые meta, диагностику, фактический метод,
признак найденной точки и ссылку на ее канонический проверенный State. Последняя
допустимая проба и достижение MaxLambda сами по себе не являются найденным пределом.
Найденная Capacity с lambda < 1 остается успешным поиском; FAIL текущего LC
назначается в Capacity-слое. На осевом плато предельная плоскость отдельно
подтверждается исходным критерием и всеми тремя усилиями, а не только скобкой силы.

### Инженерные Результаты

`CBatchSectionCalculator` управляет строками LC, последовательностью активных
расчетов, передачей настроек и выбором определяющего сочетания. Он выдает компактный
`ResultAt(index) As CCombinationResult`; индексированной скалярной витрины НДС нет.
Отказ устойчивости не отменяет другие запрошенные проверки.

`CCapacityCalculator` задает инженерную постановку и интерпретирует Search.
`CCapacitySolver` сохраняет цельную оценку capacity-пробы, физического критерия,
локальный cache/warm-start и специальные осевые операции, но не вторую реализацию
bracket/bisection/Newton. `CLoadPathVector` является общим путем; `CCapacityLoadPath`
отвечает за пользовательскую постановку Capacity и преобразование ее компонентов.

`CCrackFormationCalculator` реально владеет formation-критерием, Auto-путями,
поиском и Pre/Post states. `CCrackWidthCalculator` подготавливает данные по
физически допустимому текущему State и FormationResult: зону, стержни, `Abt/As`,
`ds/ls`, напряжения и коэффициенты. Чистая `CCrackWidthFormulaCalculator`
получает только готовые численные данные и вычисляет `a_crc`; State, выбор зоны,
статусы и новые solves ей не передаются. `CLongitudinalCrackCalculator`
независимо проверяет сжатый бетон, в том числе при отсутствии нормальной трещины.

Текущее CrackedState рассчитывается только при наличии активного потребителя:
ширины, продольных трещин или явно запрошенного снимка. Недопустимое/ненайденное
обязательное State сохраняет первичную причину; зависимая формула получает
`rsBlockedByDependency` и не вызывается по auxiliary-напряжениям.

Каноническое дерево одного LC:

```text
CCombinationResult
  StrengthResult -> DirectState -> CSectionStateResult
                 -> Capacity -> CapacityState / SearchResult
  CrackResult    -> Formation -> PreCrackState / PostCrackState / SearchResult
                 -> CurrentStateMeta, Width, Longitudinal
  StabilityResult
  StateRepository (те же конечные named-state, не копии плоскостей)
```

Числа НДС принадлежат `CSectionStateResult`. Direct-result хранит ссылку на State
и meta, а не параллельную плоскость/усилия. Запасы принадлежат результатам своих
проверок; Combination агрегирует критерий сводки, Batch выбирает LC. Плоские
таблицы Results являются сериализацией, а не дублированной доменной моделью.
Малые spec/meta/parameter snapshots изолированы; волокна на каждом getter не клонируются.

### Статусы И Комментарии

```text
InternalStatus + ResultCode + ResultKind
  -> CResultStatusPolicy
  -> ExternalStatus / Display
```

`rsSuccess`/`rsSuccessWithWarning` отображаются как `OK`, `rsCheckFailed` как
`FAIL`, реальная `rsNumericalFailure` как `NumFail`, input/config errors как
`InputErr`, внутренние нарушения как `CalcErr`, неприменимые/незапрошенные/
заблокированные результаты как `N/A`. Только Capacity с
`rsCheckFailed + rcInitialStateBeyondLimit` получает `BaseFail`: постоянная
часть выбранного пути уже не проходит физическую проверку. Несошедшаяся плоскость
при lambda=0 не доказывает BaseFail.

`rcCriterionNotReached` означает установленное недостижение; отсутствие трещины
трактуется Formation, Width неприменима. `rcSearchBoundReached` означает лишь
технический предел поиска и не превращается в доказанное отсутствие трещины.
Точка за текущим LC остается найденной точкой с `CrackFormed=False`.

Причины не восстанавливаются из StopReason/ResultComment. Ответственные
калькуляторы/results формируют русский текст, Combination собирает комментарии
нужного subtree без дублей в логическом порядке. Writers читают готовую meta:
подробный блок получает свои комментарии, batch summary все применимые разделы.
Простая presentation-арифметика и перевод единиц разрешены; назначать инженерный
статус, запускать solver или склеивать новые причины writer не должен.

### Численное Продолжение

Единственный runtime-переключатель `General.DiagramExtension = Yes/No` применяется
ко всем активным diagram-based маршрутам: direct/current states, пробам и
финализации Capacity/Formation, Pre/Post states. Старый ключ допускается только
в явной миграции Config. Явное No сохраняется, неверное значение вызывает InputErr.

Provider продолжает активные ветви за их исходным физическим пределом существующим
численным законом порядка `0.01 * E`; технический край обычно +/-10, при большем
физическом пределе безопасно выносится дальше. Физические точки, плато, сопротивления
и предельные деформации не меняются. При `Ignore` растянутый бетон по-прежнему
имеет sigma=0 и Et=0; разрешение Extension его не включает.

В State различаются `Converged`, `WithinPhysicalRange` и `ExtensionUsed` конечной
плоскости. On не означает ExtensionUsed=True, а промежуточный выход Newton
за предел не ухудшает физически допустимый финал. Auxiliary-равновесие вне
физического диапазона не является дополнительной несущей и не передается ширине
как допустимый вход. Stability, аналитические характеристики и нормативные модули
не получают техническую касательную вместо исходного модуля.

Режим No остается полноценным расчетом на обычных диаграммах; retry не включает
продолжение скрыто. Для одинаковых корректных физических решений On/Off сравниваются
в прежних численных допусках; различия итераций и времени допустимы.

### Сохраненный Вывод

- `CBatchResultWriter` пишет компактную сводку в `rngBatchSummary`;
- `CStrengthSummaryWriter` пишет подробную таблицу НДС и несущей способности от `rngStrengthSummaryAnchor`;
- `CCrackSummaryWriter` пишет подробную таблицу трещин от `rngCrackSummaryAnchor`;
- `CStabilitySummaryWriter` пишет подробную таблицу устойчивости от `rngStabilitySummaryAnchor`;
- `CNDMResultsWriter` пишет расчетный snapshot: `rngNDMSectionGeometry`, `rngNDMElementResults`, `rngNDMSectionProperties`, `rngNDMSectionAnnotations`.

AutoCAD export и Excel-схема читают последний снимок `Results`. Они не запускают решатель, не читают исходную геометрию заново и не держат модель в памяти между макросами.
