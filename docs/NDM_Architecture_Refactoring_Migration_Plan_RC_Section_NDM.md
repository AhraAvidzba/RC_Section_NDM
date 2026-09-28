# Подробный план миграции архитектуры RC Section NDM

Дата подготовки: 2026-09-28.

Исходный укрупненный документ: `docs/NDM_Architecture_Refactoring_Plan_Codex_v6.md`.

Назначение этого файла: развернуть укрупненный план применительно к текущему коду проекта `RC_Section_NDM`, существующим VBA-классам, листам Excel, writer-ам и уже согласованным расчетным правилам. Это план миграции, а не реализация.

---

## 1. Цель миграции

Нужно перестроить архитектуру расчетного ядра вокруг четырех независимых слоев:

```text
State
    Решение одного конечного НДС.

Search
    Поиск lambda/предельной точки по заданной траектории и критерию.

Engineering calculations
    Capacity, Crack, Stability и другие проверки, использующие готовые State/Search.

Result + Presentation
    Хранение результата, статусы, комментарии и пассивный вывод.
```

Главная инженерная цель: одно и то же расчетное состояние сечения не должно решаться повторно разными модулями, если оно уже было получено с теми же нагрузками, материалами и режимом решения.

Главная архитектурная цель: расчетные классы должны возвращать структурированный результат со статусом, кодом и понятным комментарием, а writer-ы должны только выводить готовые данные.

---

## 2. Жесткие ограничения

Рефакторинг не должен менять:

- расчетную постановку `N + Mx + My`;
- внутренние единицы: мм, мм2, Н, Н*мм, МПа, безразмерная деформация, 1/мм;
- пользовательскую систему единиц и знаков, проходящую через `CUnitSystem`;
- математику НДМ, материалов, трещин, несущей способности и устойчивости без отдельного решения;
- логику генерации геометрии и AutoCAD import/export, кроме адаптации к новым результатам;
- численные результаты существующих regression-тестов, за исключением случаев, где тест явно проверяет старое имя состояния или старую структуру результата;
- внешний смысл display-статусов `OK`, `FAIL`, `BaseFail`, `NumFail`, `InputErr`, `CalcErr`, `N/A`;
- существующий layout Results/Summary, кроме согласованного добавления `ResultComment` перед общим статусом каждого расчетного блока.

Рефакторинг может менять:

- внутренние классы и API;
- имена внутренних состояний;
- структуру `CCombinationResult`;
- место, где рассчитываются и агрегируются статусы;
- способ переиспользования уже найденных состояний;
- дублирующую search-логику Capacity/Crack;
- writer-ы, если они после изменения становятся более пассивными.

---

## 3. Текущая карта кода

### 3.1. Ввод, единицы и настройки

Текущие ответственные классы и модули:

- `CSystemSettingsReader` читает настройки листа `Config`;
- `CUnitSystem` переводит пользовательские единицы и знаки во внутренние и обратно;
- `CLoadCombinationReader` читает `rngLoadCombinations`;
- `CCalculationProfileCatalog` и `CCalculationProfile` задают, какие расчеты запускать по профилю;
- `CSectionLoadState` переносит нагрузки между пользовательской точкой, расчетными центрами и внутренним началом модели;
- `CMomentZeroFilter` применяет `Calculation.ZeroMomentPerDepth`;
- `CLoadPathVector` и `CLoadPathMath` описывают траекторию `Offset + lambda * Base`.

Эти элементы в целом соответствуют целевой архитектуре и должны переиспользоваться.

### 3.2. Сечение и материалы

Текущие ответственные классы:

- `CSectionModel` хранит бетонные конечные элементы, арматуру, реальные характеристики импортированных Region и геометрическое представление;
- geometry builders для `Circle`, `RoundedRectangle`, `HollowRectangle`, `RectSet`;
- `CSectionPropertiesCalculator` считает бетонные и приведенные характеристики;
- `CMaterialModelProvider`, `CMaterialModelSpec`, `CMaterialDiagram` создают диаграммы материалов для разных расчетных ролей.

Этот слой не является целью текущего архитектурного рефакторинга. Он должен остаться источником геометрии, материалов и характеристик для State/Search.

### 3.3. Прямое решение НДС

Текущие классы:

- `CSectionSolver` решает равновесие `N`, `Mx`, `My`;
- `CStateSolutionRunner` настраивает solver, выбирает стартовую плоскость, делает retry и extension warm start;
- `CSectionStateResult` сохраняет результат одного НДС;
- `modCalculationPurpose` задает `ESectionStateType`, `ECalculationPurpose` и текстовые преобразования.

Сильная сторона текущего кода: уже есть почти готовый общий runner и общий объект состояния.

Недостаток: state-result создается и складывается в разных местах batch/crack/capacity, нет общего `Get-or-Solve`, а ключ эквивалентности состояния не формализован.

### 3.4. Batch-оркестратор

`CBatchSectionCalculator` сейчас выполняет слишком много обязанностей:

- запускает stability filter;
- запускает capacity;
- запускает direct strength state;
- запускает crack calculation;
- решает отдельные named states через `RunProfileState`;
- хранит настройки solver-а;
- собирает статусы;
- пишет человекочитаемый execution report;
- выбирает worst combination;
- рассчитывает summary reserves.

Это рабочая, но перегруженная точка. В целевой архитектуре он должен остаться orchestration-слоем, но потерять подробную инженерную и status-логику.

### 3.5. Capacity

`CCapacitySolver` сейчас содержит:

- поиск предельной несущей способности;
- две стратегии `UltimateStrain` и `LoadMultiplier`;
- fallback из `Auto`;
- probe/retry/warm-start логику;
- локальный probe cache;
- классификацию численных и физических исходов;
- хранение итоговых `lambda`, `Nult`, `Mxult`, `Myult`, `CapacityState`;
- диагностический лог.

Именно из него нужно извлечь универсальный search-слой. При этом важно не потерять текущую оптимизацию пробных lambda-точек и текущие стартовые подсказки.

### 3.6. Crack

`CCrackWidthCalculator` сейчас объединяет несколько задач:

- проверку образования нормальной трещины;
- выбор пути `SLS.Crack.InitiationLoadPath`;
- стратегию `SLS.Crack.InitiationSolutionStrategy`;
- поиск `lambda_crc`, `Ncrc`, `Mxy,crc`;
- хранение `BeforeMcrcState` и `AfterMcrcState`;
- расчет текущего раскрытия;
- расчет `psi_s`, включая `User`, `Auto`, `AlwaysCalc`;
- расчет продольных трещин частично через batch;
- диагностический лог.

Это главный кандидат на разделение. Формула раскрытия трещины должна стать чистым расчетом по готовым численным данным, подготовленным из states и выбранных стержней вне формулы, без скрытого запуска solver-а.

### 3.7. Stability

`CStabilityCalculator` уже в основном похож на независимый инженерный расчет:

- получает сечение, материалы, усилия, расчетные длины и настройки;
- возвращает результаты по СП 35/СП 63;
- не запускает `CSectionSolver`;
- не должен блокировать другие запрошенные расчеты.

В рефакторинге Stability в первую очередь нужно подключить к общей системе result/status/comment и убрать локальное дублирование статусных решений, если оно появится.

### 3.8. Результаты и writer-ы

Текущие классы:

- `CCombinationResult` хранит много плоских полей, включая состояния, crack, capacity, stability;
- `CBatchStatusPolicy` переводит внутренние тексты в пять пользовательских статусов;
- `CBatchResultWriter` пишет `rngBatchSummary`;
- `CStrengthSummaryWriter` пишет подробный блок прочности;
- `CCrackSummaryWriter` пишет подробный блок трещин;
- `CStabilitySummaryWriter` пишет подробный блок устойчивости;
- `CNDMResultsWriter` пишет расчетный снимок `Results`: геометрию, элементы, свойства, annotations, диаграммы.

Сильная сторона: writer-ы уже в основном пишут блочно массивами.

Недостаток: часть writer-ов и batch-класс все еще знают слишком много о смысле статусов и отдельных fallback-ветках. В целевой архитектуре writer должен читать готовое дерево результата и не принимать инженерные решения.

---

## 4. Основные архитектурные разрывы

### 4.1. Нет общего StateRepository

Сейчас `CCombinationResult` умеет хранить список `CSectionStateResult`, но не является полноценным repository:

- нет ключа эквивалентности состояния;
- нет централизованного `Get-or-Solve`;
- расчетные модули могут запускать похожие состояния сами;
- reuse между Capacity/Crack/DirectState не гарантирован архитектурно.

Нужен отдельный `CStateRepository`, а `CCombinationResult` должен оставаться корнем готового результата LC.

### 4.2. Search-алгоритмы смешаны с физическим смыслом

Capacity уже содержит развитую инфраструктуру поиска, но она привязана к прочности. Crack имеет собственный поиск образования трещины. В целевой архитектуре общий search должен знать только:

```text
Target(lambda) = Offset + lambda * Base
criterion(lambda/state)
strategy = UltimateStrain / LoadMultiplier / Auto
```

А физический смысл предела должен жить в criterion-объекте.

### 4.3. Внутренний статус смешан с пользовательским статусом

Сейчас `CBatchStatusPolicy` уже сводит результаты к `OK/FAIL/NumFail/InputErr/N/A`, но расчетные классы часто передают наружу строковые `StopReason`, `LimitState`, `CrackStatus`, `CapacityStatus`. Это порождает неоднозначность: `NumFail` может всплыть в блоке, где solver уже не запускался.

Нужно отделить:

- внутренний `InternalStatus`;
- машинный `ResultCode`;
- внешний `ExternalStatus`/display-статус;
- человекочитаемый `ResultComment`;
- технический diagnostic log.

Цепочка должна быть однозначной: расчетный код формирует `InternalStatus` и `ResultCode`, инженерный слой при необходимости уточняет их смысл для своего блока, `CResultStatusPolicy` превращает эту пару во внешний `ExternalStatus`, а writer только выводит готовый display-статус.

### 4.4. Crack слишком монолитный

Текущий `CCrackWidthCalculator` знает и про поиск трещинообразования, и про расчет ширины, и про sigma_s_crc, и про часть state management. После рефакторинга должны быть отдельные шаги:

```text
CrackFormation
CurrentCrackedState
CrackWidthFormula
LongitudinalCrackCheck
CrackResult aggregation
```

### 4.5. CCombinationResult слишком плоский

Сейчас одно сочетание хранит много однотипных scalar-полей. Это удобно для writer-ов, но мешает архитектуре:

- сложно понять, к какому расчетному блоку относится поле;
- сложно добавить `InternalStatus`, `ResultCode` и `ResultComment`;
- сложно гарантировать, что writer не выводит случайный старый статус;
- сложно поддерживать несколько named states.

Нужно перейти к вложенной структуре результатов.

---

## 5. Целевая структура классов

Имена ниже являются рекомендуемыми для проекта. Во время реализации их можно уточнить, если фактический код покажет более естественный вариант. Не нужно создавать каждый предложенный класс как самоцель: если после анализа отдельная сущность оказывается пустой тонкой оболочкой, ее можно объединить с соседней ответственностью. При этом архитектурные границы `State / Search / Engineering / Result+Presentation` нарушать нельзя.

### 5.1. State-слой

#### `CSectionStateResult`

Развить существующий класс. Он должен хранить:

- `StateType`;
- `StateName`/текстовое имя;
- `MaterialModelRole`;
- `MaterialSpec`;
- целевые `TargetN`, `TargetMx`, `TargetMy`;
- фактические `Nint`, `Mxint`, `Myint`;
- `epsilon0`, `kappaX`, `kappaY`;
- невязки и относительные невязки;
- min/max strain/stress по бетону и арматуре;
- `ExtensionUsed`;
- `WithinPhysicalRange`;
- `Converged`;
- `SolverCallCount`;
- `IterationCount`;
- `ResultMeta`;
- diagnostic/stop reason.

Текущие свойства min/max и flags нужно сохранить, чтобы не ломать writer-ы одномоментно.

#### `CStateRequest`

`CStateRequest` должен разделять две группы данных: физический запрос состояния и solve-options.

Параметры физического запроса:

- `StateType`;
- `TargetN`, `TargetMx`, `TargetMy`;
- `MaterialRole`;
- `MaterialSpecKey`;
- `DiagramExtensionMode`, если он меняет фактическую диаграмму или допустимость результата;
- настройки solver-а только в той части, где они меняют допустимость физического результата, а не только шанс сходимости.

Solve-options:

- warm-start state;
- initial guess / initial plane policy;
- initial solver;
- retry policy;
- diagnostic/report flags;
- численные настройки, которые влияют только на путь поиска и вероятность сходимости, но не на физический результат и не на критерии его приемки.

Ключ эквивалентности state строится только из параметров физического запроса. Warm-start, initial guess, initial solver и retry policy не должны входить в ключ reuse, если они не меняют физический результат или его допустимость. Один и тот же state, найденный с разными начальными приближениями, должен переиспользоваться как эквивалентный.

Важно: ключ должен отличать `PostCrackState` от `CurrentCrackedState` даже при одинаковой диаграмме и нагрузке, потому что это разные инженерные named states.

#### `CStateRepository`

Новый repository на один LC:

- хранит как reusable только корректно полученные конечные named states;
- не хранит внутренние lambda-probes search-а;
- умеет `FindEquivalent(request)` через state equivalence key, а не через solve-options;
- умеет `Add(request, stateResult)` только для физически полученного reusable named state;
- возвращает состояние для writer-ов и downstream-расчетов;
- не запускает solver самостоятельно, если будет выделен отдельный provider.

Неуспешный solve не должен становиться reusable physical State. Результаты попыток с `rsNumericalFailure`, `rsInvalidInput`, `rsInvalidConfiguration`, `rsInternalError` или другим статусом, при котором физическое НДС не получено, не должны возвращаться через `FindEquivalent` как готовое состояние и блокировать новую попытку решения. Их диагностику можно сохранять в соответствующем result/diagnostic, но не использовать как cache успешного физического состояния.

Особенно важно: если тот же физический запрос позже выполняется с другим warm-start, initial guess, initial solver или retry policy, должна оставаться возможность новой попытки solve. Поэтому неуспешные attempts могут участвовать в диагностике и отчете, но не закрывают state equivalence key как уже решенный reusable state.

#### `CStateProvider`

Тонкий сервис над `CStateRepository` и `CStateSolutionRunner`:

```text
GetOrSolve(request, section, materials, optional initialState)
```

Ответственность:

- собрать материалы по `request.MaterialSpec`;
- запустить `CStateSolutionRunner`, если состояния нет;
- создать `CSectionStateResult`;
- присвоить `ResultMeta`;
- сохранить результат в `CStateRepository` только если получен корректный reusable named state;
- вернуть неуспешный result вызывающему расчету/diagnostic без записи как reusable state, если физическое НДС не получено;
- вернуть готовый state.

Если в VBA окажется удобнее объединить `CStateRepository` и provider в одном классе, это допустимо только при сохранении двух логических ответственностей.

### 5.2. Status/result-слой

#### `CResultMeta`

Новый маленький объект, полностью соответствующий v6:

- `InternalStatus` - внутренний enum, на котором работает расчетное ядро;
- `ResultCode` - машинный код точной причины;
- `ResultComment` - короткое русское пояснение для пользователя;
- `DiagnosticDetails` - технический лог, если нужен;
- optional `Applies`/`Calculated`.

Внутри расчетного ядра нельзя использовать строки `OK`, `FAIL`, `NumFail` как основной тип статуса. Эти строки являются только внешним display-представлением.

Обязательный внутренний enum должен иметь смысловой набор наподобие:

```text
EResultStatus

rsUnset
rsNotRequested
rsNotApplicable
rsBlockedByDependency

rsSuccess
rsSuccessWithWarning

rsCheckFailed

rsNumericalFailure
rsInvalidInput
rsInvalidConfiguration
rsInternalError
```

Смысл статусов:

| InternalStatus | Смысл |
|---|---|
| `rsUnset` | результат еще не сформирован; в завершенном расчете это ошибка архитектуры или непредусмотренная ветка |
| `rsNotRequested` | расчет отключен профилем |
| `rsNotApplicable` | проверка физически или методически неприменима |
| `rsBlockedByDependency` | этап не запускался, потому что обязательный предыдущий этап не дал данных |
| `rsSuccess` | расчет корректно завершен |
| `rsSuccessWithWarning` | корректный результат с fallback или особой веткой |
| `rsCheckFailed` | расчет корректен, но инженерная проверка не проходит |
| `rsNumericalFailure` | настоящая численная несходимость state/search |
| `rsInvalidInput` | ошибка пользовательских исходных данных |
| `rsInvalidConfiguration` | ошибка профиля или настроек |
| `rsInternalError` | непредусмотренная внутренняя ошибка программы |

#### `CResultStatusPolicy`

Преемник или развитие `CBatchStatusPolicy`:

- единственное место, где `InternalStatus + ResultCode + контекст результата` превращаются во внешний display-статус;
- единственное место, где задается приоритет агрегирования;
- единственное место для цветов статусов;
- не читает Excel;
- не запускает расчеты.

`CBatchStatusPolicy` можно переименовать или заменить, но финально не должно остаться двух конкурирующих политик.

Цепочка преобразования должна быть именно такой:

```text
InternalStatus
    + ResultCode
    + ResultKind
        ↓
ExternalStatus
        ↓
DisplayText / cell color
```

Внешние display-статусы:

```text
OK
FAIL
BaseFail
NumFail
InputErr
CalcErr
N/A
```

Базовое соответствие:

| InternalStatus | ExternalStatus |
|---|---|
| `rsUnset` | `CalcErr` |
| `rsNotRequested` | `N/A` |
| `rsNotApplicable` | `N/A` |
| `rsBlockedByDependency` | `N/A` |
| `rsSuccess` | `OK` |
| `rsSuccessWithWarning` | `OK` |
| `rsCheckFailed` | `FAIL` |
| `rsNumericalFailure` | `NumFail` |
| `rsInvalidInput` | `InputErr` |
| `rsInvalidConfiguration` | `InputErr` |
| `rsInternalError` | `CalcErr` |

Обязательный контекстный случай:

```text
ResultKind = CapacityResult
InternalStatus = rsCheckFailed
ResultCode = INITIAL_STATE_BEYOND_LIMIT
        ↓
ExternalStatus = BaseFail
```

`BaseFail` использовать только для Capacity / несущей способности по прочности. Он означает, что уже при `lambda = 0` базовая часть нагрузки не несется сечением. Этот исход не является `NumFail`.

#### ResultCode-группы

Минимальный набор групп:

```text
STATE_SOLVED
STATE_NUMERICAL_FAILURE
STATE_INPUT_ERROR
STATE_OUTSIDE_PHYSICAL_RANGE

LIMIT_FOUND
INITIAL_STATE_BEYOND_LIMIT
CRITERION_NOT_REACHED
SEARCH_BOUND_REACHED
LIMIT_NUMERICAL_FAILURE
LIMIT_INPUT_ERROR

CHECK_OK
CHECK_FAILED
CHECK_NOT_APPLICABLE
CHECK_INPUT_ERROR
```

Точные имена можно уточнить, но коды должны быть стабильными, без парсинга произвольного текста.

Особые правила для search-кодов:

- `INITIAL_STATE_BEYOND_LIMIT` - критерий уже достигнут при `lambda = 0`. Это физически определенный исход начальной точки траектории, а не numerical failure.
- Для Capacity этот код превращается инженерным слоем в `rsCheckFailed`, а внешне через policy в `BaseFail`.
- Для CrackFormation в режиме `Auto` этот код может быть промежуточным сигналом попробовать другой допустимый load path.
- Для CrackFormation с фиксированным path этот код должен остаться физическим результатом выбранной траектории и получить понятный `ResultComment`.
- `CRITERION_NOT_REACHED` применять только когда недостижение критерия действительно установлено физически или методически. Если search корректно доказал недостижение критерия, `CLimitSearchResult` может иметь `InternalStatus = rsSuccess` и `ResultCode = CRITERION_NOT_REACHED`.
- Для CrackFormation `rsSuccess + CRITERION_NOT_REACHED` означает корректный результат `CrackFormed = False`. Следующий расчет ширины нормальной трещины получает `rsNotApplicable`, а не `rsCheckFailed` и не `rsNumericalFailure`; внешний `N/A` появляется только через `CResultStatusPolicy`.
- `SEARCH_BOUND_REACHED` применять, когда достигнут технический предел поиска, например `MaxLambda`, но нельзя утверждать, что критерий физически недостижим.

Все enum/status/result-code declarations должны иметь русские комментарии прямо у объявления: что означает код, чем отличается от близких кодов и когда его допустимо формировать. Там, где generic search-code получает инженерный смысл в Capacity или CrackFormation, рядом должен быть отдельный русский комментарий о такой трактовке.

### 5.3. Search-слой

#### `CLimitSearchRequest`

Содержит:

- `LoadPathVector`;
- `SearchStrategy`: `Auto`, `UltimateStrain`, `LoadMultiplier`;
- `SearchMethod`: `Bisection`, `Brent`, `Secant` для LoadMultiplier;
- `MaxLambda`, `InitialLambdaStep`, tolerances;
- solver settings;
- material spec;
- имя будущего final state.

#### `CLimitSearchResult`

Содержит:

- `ResultMeta`;
- `Lambda`;
- `TargetN`, `TargetMx`, `TargetMy`;
- `LimitState`;
- `SolutionMethod`;
- `ResolvedLoadPath`;
- `FinalState` как `CSectionStateResult`, если найдено конечное состояние;
- `ProbeCount`;
- `SolverCallCount`;
- diagnostic log.

#### `ILimitCriterion` или VBA-эквивалент

VBA не поддерживает интерфейсы так же удобно, но можно использовать класс с единым набором методов или `Implements`.

Минимальная логика критерия:

```text
EvaluateState(state) -> criterion value / pass/fail / utilization
TargetReached(state)
InitialStateBeyondLimit(state)
BuildUltimateStrainTarget(...)
CommentForOutcome(...)
```

Конкретные критерии:

- `CCapacityLimitCriterion`;
- `CCrackFormationCriterion`.

#### `CLimitSearchCoordinator`

Оркестратор:

- получает request, criterion, state provider;
- выбирает порядок попыток;
- сначала пробует `UltimateStrain`, если стратегия разрешает;
- при численной неудаче пробует `LoadMultiplier`;
- возвращает `CLimitSearchResult`;
- пишет диагностические сообщения в структурированном виде.

#### `CUltimateStrainSearch`

Выделенная реализация поиска по предельной деформации. Она должна быть общей для:

- несущей способности;
- образования трещины по `eps_bt,ult`.

Различие должно быть только в criterion и material spec.

#### `CLoadMultiplierSearch`

Выделенная реализация одномерного поиска lambda через серию state-solve.

Обычные probe-точки `LoadMultiplier`, где для заданных `N/Mx/My` нужно найти равновесное НДС, должны использовать общий state-solve pipeline: тот же механизм настройки solver-а, построения request/options, решения равновесия, формирования `CSectionStateResult`, `ResultMeta` и диагностики. Это сохраняет единое solver behavior, единый учет solver calls, единые статусы и одинаковую диагностику.

При этом общий state-solve pipeline не означает запись probe-точек в shared `CStateRepository`. Промежуточные states одного search-session остаются только во внутреннем локальном probe-cache/search-session. Они могут переиспользоваться внутри текущего search, чтобы не терять существующую оптимизацию повторных lambda-точек, но не становятся reusable named states для других расчетных блоков.

Прямой низкоуровневый вызов `CSectionSolver` внутри Search допустим только там, где алгоритму действительно нужен не готовый reusable State, а специальная операция уровня solver-а, например оценка заданной плоскости деформаций через `EvaluateStrainPlane` или аналогичный внутренний шаг. Конкретный класс/API заранее не фиксируется: если фактический код покажет более естественную схему, ее можно выбрать при сохранении границы `State solve -> Search -> Engineering result`.

Нужно перенести из `CCapacitySolver`:

- bracket logic;
- bisection/brent/secant;
- обработку `MaxLambda`;
- уменьшение lambda после numerical failure;
- probe cache;
- warm start соседних точек.

Probe cache должен быть локальным для одного search-session и не попадать в общий `CStateRepository`. В shared repository попадает только итоговый физически значимый named state, например `CapacityState`, `PreCrackState`, `PostCrackState` или `CurrentCrackedState`, если он действительно получен и нужен downstream-расчетам.

### 5.4. Engineering-слой

#### Capacity

Предлагаемая структура:

- `CCapacityCalculator` или сохраненный `CCapacitySolver` как thin facade;
- внутри использует `CLimitSearchCoordinator`;
- больше не содержит собственную копию generic search;
- возвращает `CCapacityResult`.

`CCapacityResult`:

- `ResultMeta`;
- `LambdaUltimate`;
- `NUltimate`, `MxUltimate`, `MyUltimate`;
- `SafetyFactor`;
- `LimitState`;
- `LoadPath`;
- `SolutionMethod`;
- `CapacityState`;
- reserves и поля для writer-а.

#### Crack

Предлагаемая структура:

- `CCrackCalculator` как общий сценарий трещин;
- `CCrackFormationCalculator` для проверки образования нормальной трещины;
- `CCrackWidthFormulaCalculator` для чистой формулы ширины;
- `CLongitudinalCrackCalculator` для продольных трещин;
- `CCrackResult` как структурированный результат.

`CCrackFormationCalculator`:

- использует `CLimitSearchCoordinator`;
- сохраняет `PreCrackState` и `PostCrackState`;
- не считает ширину раскрытия;
- возвращает `ResultMeta`, `lambda_crc`, `Ncrc`, `Mxy_crc`, фактический метод/путь.

`CCrackWidthFormulaCalculator`:

- принимает не solver/state, а готовый набор чисел для формулы раскрытия;
- получает уже подготовленные данные расчетной зоны, арматуры, `sigma_s`, `psi_s` и расстояния между трещинами;
- считает только последнюю формульную часть раскрытия: применение готовых `phi/psi/ls`, `a_crc`, utilization;
- не запускает `CSectionSolver`.

Выбор tension zone, выбранных стержней, `Abt`, `As`, `ds`, `ls`, `sigma_s` и подготовка геометрических данных должны быть вне чистой формулы. В v6 для этого предложен `CCrackSectionDataCalculator`, но имя и количество классов можно уточнить, если фактический код подскажет более естественную структуру. Запрещено только снова смешивать solver/search/geometry/formula в одном монолите.

`CLongitudinalCrackCalculator`:

- принимает готовый `CurrentCrackedState`;
- считает проверку сжатого бетона;
- не зависит от результата нормальных трещин, кроме общего `CCrackResult`.

### 5.5. Result graph

`CCombinationResult` должен стать корнем дерева:

```text
CCombinationResult
    InputInfo
    StateRepository
    StrengthResult
        DirectStateResult
        CapacityResult
    CrackResult
        FormationResult
        CurrentCrackedState, если нужен хотя бы одному активному потребителю
        WidthResult
        LongitudinalResult
    StabilityResult
    OverallResultMeta
```

На переходном этапе можно оставить старые getters для writer-ов, но финально legacy-поля должны быть удалены, если они больше не нужны.

### 5.6. Writers

Writer-ы должны:

- читать только `CCombinationResult` и вложенные result-объекты;
- не запускать solver;
- не считать `lambda`, `psi_s`, `sigma_s`, reserves;
- не парсить `StopReason`;
- не решать, является ли numerical failure физическим FAIL;
- только форматировать значения, единицы, статусы, комментарии и цвета.

Каждый расчетный блок Results/Summary должен получить `ResultComment` перед своим общим статусом, как требует исходный план.

`ResultComment` должен собираться централизованно из соответствующего result-subtree. Для блока Crack комментарии берутся из ветвей formation, current state, width и longitudinal crack; для Capacity - из direct state, limit search и capacity state; для Stability - из активных проверок устойчивости. Сборщик комментариев должен:

- сохранять логический порядок этапов;
- удалять дубли одинаковых сообщений;
- не подтягивать сообщения из соседних независимых блоков;
- возвращать готовую строку или набор строк для writer-а.

Writer не должен сам склеивать `StopReason`, diagnostic text или display-статусы в пользовательский комментарий.

---

## 6. План миграции по этапам

## Этап 0. Подготовка и базовая фиксация

Цель: создать безопасную точку сравнения до архитектурных изменений и сделать длинный рефакторинг продолжимым после compaction / потери диалогового контекста.

Шаги:

1. Перед основными правками выполнить `git status` и понять текущее рабочее дерево.
2. Определить текущий рабочий `HEAD`, соответствующий последней заведомо рабочей версии старой архитектуры.
3. Сохранить hash этого commit как baseline рефакторинга.
4. Не переписывать и не разрушать baseline; при необходимости сравнивать с ним через Git, а не по памяти.
5. Создать короткий progress/status Markdown-файл в естественном месте проекта. Имя заранее не фиксировать, но файл должен быть легко найден из архитектурного плана.
   Для текущей миграции используется `docs/NDM_Architecture_Refactoring_Progress.md`.
6. В progress/status-файле поддерживать минимум:

```text
Baseline:
<commit hash>

Completed:
- ...

In progress:
- ...

Next:
- ...

Known risks / open questions:
- ...

Important decisions:
- ...

Last verified:
- tests/builds/manual checks
```

7. Собрать книгу через `tools/build_workbook/Build-Workbook.ps1` с timeout не меньше 300 секунд.
8. Запустить полный набор существующих тестов.
9. Сохранить baseline-поведение:
   - `rngBatchSummary`;
   - detailed strength/crack/stability blocks;
   - `rngNDMSectionProperties`;
   - execution report для нескольких характерных сценариев.
10. Зафиксировать список тестовых книг/сценариев:
   - стандартное Г-сечение;
   - круг;
   - RoundedRectangle;
   - HollowRectangle;
   - импортированная AutoCAD Region geometry;
   - случаи с direct state, capacity, crack, stability.

Выход этапа:

- baseline численных результатов;
- baseline commit hash;
- короткий progress/status MD как внешняя память рефакторинга;
- список тестов, которые должны остаться неизменными;
- перечень тестов, которые надо обновить из-за новых имен `PreCrackState/PostCrackState` или `ResultComment`.

Правила восстановления после compaction:

```text
git status
git diff
git log
baseline commit
docs/NDM_Architecture_Refactoring_Plan_Codex_v6.md
docs/NDM_Architecture_Refactoring_Migration_Plan_RC_Section_NDM.md
progress/status MD
```

Если после compaction непонятно, какой этап выполнен, сначала читать эти источники, а не начинать этап заново. При необходимости использовать `git show` и сравнение с baseline/milestone commits.

Правила безопасности Git:

- Git использовать как источник истории и сравнения, а не как повод автоматически откатывать пользовательские изменения.
- Не выполнять destructive reset / checkout / rewrite history без явной необходимости и явного понимания последствий.
- Не откатывать существующие пользовательские изменения, не относящиеся к текущему рефакторингу.
- Перед крупным этапом всегда смотреть `git status`.
- После этапа проверять diff и тесты, чтобы случайные изменения не протекли в соседние части проекта.

## Этап 1. Фундамент статусов и комментариев

Цель: отделить внутренние причины от пользовательских статусов до переноса solver/search-логики.

Шаги:

1. Ввести `CResultMeta`.
2. Расширить или заменить `CBatchStatusPolicy` на `CResultStatusPolicy`.
3. Зафиксировать единый словарь:
   - `OK`;
   - `FAIL`;
   - `BaseFail`;
   - `NumFail`;
   - `InputErr`;
   - `CalcErr`;
   - `N/A`.
4. Ввести enum/константы `ResultCode`, не завязанные на произвольный текст `StopReason`.
5. Перевести текущие места, где создаются статусы:
   - `CapacityStatusFromSolver`;
   - `CrackStatusFromCalculator`;
   - `LongitudinalCrackStatus`;
   - `StabilityStatusFromCalculator`;
   - direct state status.
6. Добавить в result-объекты поля:
   - `InternalStatus`;
   - `ResultCode`;
   - `ResultComment`.
7. Обновить aggregation: итоговый статус LC должен собираться только из `ResultMeta`, а не из диагностических строк.
8. Добавить тесты, что:
   - `NumFail` возникает только там, где запускался поиск равновесия или предельной точки;
   - расчет ширины раскрытия, продольных трещин и устойчивости без solver-а не возвращает `NumFail`;
   - `N/A` не ухудшает общий статус.

Что пока не менять:

- расчетную математику;
- структуру search;
- порядок вызовов solver-а.

## Этап 2. Развитие `CSectionStateResult`

Цель: сделать результат НДС самодостаточным.

Шаги:

1. Добавить в `CSectionStateResult` недостающие поля:
   - целевые усилия;
   - фактические внутренние усилия;
   - невязки;
   - `WithinPhysicalRange`;
   - `ResultMeta`;
   - counters/metadata solver-а.
2. Перенести статусную классификацию состояния из batch/runner в единый метод создания `CSectionStateResult`.
3. Проверить, что состояние содержит все данные, которые нужны:
   - `CStrengthSummaryWriter`;
   - `CCrackSummaryWriter`;
   - `CNDMResultsWriter`;
   - Excel plot;
   - AutoCAD export.
4. Переименовать состояния:
   - `BeforeMcrcState` -> `PreCrackState`;
   - `AfterMcrcState` -> `PostCrackState`.
5. Обновить `modCalculationPurpose`:
   - enum `ESectionStateType`;
   - `SectionStateTypeFromText`;
   - `SectionStateTypeToText`;
   - `MaterialRoleFromStateType`.
6. Обновить все ссылки в plot/export/tests/help.

Переходное правило:

- если для миграции нужен короткий переходный период, старые текстовые имена можно принять как входные aliases;
- в финальном коде старые имена не должны использоваться как основные.

Тесты:

- direct `StrengthState` сохраняет те же `epsilon0/kx/ky`, extrema и статусы;
- `CrackedState` сохраняет те же результаты;
- старые/новые имена в Results не расходятся;
- visualization state lookup работает после переименования.

## Этап 3. `CStateRepository` и `CStateProvider`

Цель: централизовать получение named state.

Шаги:

1. Создать `CStateRequest`.
2. Создать `CStateRepository`.
3. Создать `CStateProvider`.
4. Перевести `CBatchSectionCalculator.RunProfileState` на `CStateProvider.GetOrSolve`.
5. Перевести расчет `CurrentCrackedState` на `CStateProvider`.
6. Перевести `PreCrackState/PostCrackState` на `CStateProvider`.
7. Сохранять repository внутри `CCombinationResult`, но не давать ему запускать solver.
8. Убрать прямые добавления state в `CCombinationResult`, кроме одного контролируемого пути.
9. Обновить счетчик вызовов solver-а: он должен по-прежнему показывать фактическое число попыток равновесия.

Ключ состояния должен учитывать:

- state type;
- target N/Mx/My;
- material role/spec;
- extension mode;
- режим direct state solution, если он влияет на физический результат;
- версию/ключ диаграммы материала.

Ключ состояния не должен учитывать:

- человекочитаемый comment;
- diagnostic log;
- writer/output units;
- random order of requests;
- warm-start, если он не меняет физический результат;
- initial guess / initial plane;
- initial solver;
- retry policy, если она меняет только шанс сходимости, но не физический результат и не критерии его допустимости.

Тесты:

- повторный запрос того же state не увеличивает solver-call count;
- `PostCrackState` и `CurrentCrackedState` не склеиваются случайно;
- разные material spec дают разные states;
- разные target loads дают разные states.

## Этап 4. Общий `LimitSearch`

Цель: выделить search из `CCapacitySolver` и использовать его для Capacity и CrackFormation.

Шаги:

1. Создать `CLimitSearchRequest`.
2. Создать `CLimitSearchResult`.
3. Создать criterion для capacity.
4. Создать criterion для crack formation.
5. Создать `CLimitSearchCoordinator`.
6. Создать `CUltimateStrainSearch`.
7. Создать `CLoadMultiplierSearch`.
8. Перенести из `CCapacitySolver`:
   - lambda target logic;
   - bracket search;
   - bisection/brent/secant;
   - max lambda handling;
   - numerical failure classification;
   - probe retry;
   - probe cache;
   - diagnostic lines.
9. Сначала подключить новый search только к Capacity.
10. Добиться полного совпадения capacity baseline.
11. Только после этого подключить CrackFormation.

Стандартные outcome-коды:

```text
LIMIT_FOUND
INITIAL_STATE_BEYOND_LIMIT
CRITERION_NOT_REACHED
SEARCH_BOUND_REACHED
LIMIT_NUMERICAL_FAILURE
LIMIT_INPUT_ERROR
```

Особое правило:

- промежуточные lambda probes не сохраняются в `CStateRepository`;
- найденная итоговая точка `lambda_ult` или `lambda_crc` оформляется как named state и может сохраняться в `CStateRepository` только если соответствующее физическое НДС корректно получено;
- неуспешные probe/state attempts остаются в `CLimitSearchResult`/diagnostic и локальном search-session cache, но не становятся reusable state.

Тесты Capacity:

- `Capacity.SolutionStrategy = UltimateStrain`;
- `Capacity.SolutionStrategy = LoadMultiplier`;
- `Capacity.SolutionStrategy = Auto` с успешным UltimateStrain;
- `Auto` с fallback на LoadMultiplier;
- `λ*N`, `λ*Mxy`, `λ*NMxy`;
- чистая осевая нагрузка;
- внецентренная нагрузка;
- случаи с `MaxRetries` и `BaseLoadSteps`;
- неизменность `Nult/Mxult/Myult/lambda/status`.

Тесты CrackFormation:

- `SLS.Crack.InitiationLoadPath = λ*Mxy`;
- `λ*N`;
- `λ*NMxy`;
- `Auto` с переключением путей;
- начальная точка уже за пределом;
- критерий не достигается;
- search bound reached;
- numerical failure должен попадать именно в status поиска образования трещины.

## Этап 5. Перестройка Capacity

Цель: оставить Capacity инженерным расчетом, а не владельцем search-алгоритмов.

Шаги:

1. Оставить `CCapacitySolver` как временный facade или переименовать в `CCapacityCalculator`.
2. Вынести из него generic search-код в `CLimitSearch*`.
3. Внутри capacity оставить:
   - построение load path из текущего LC;
   - выбор criterion;
   - интерпретацию limit state;
   - расчет reserves;
   - формирование `CCapacityResult`.
4. `CBatchSectionCalculator.RunCapacity` должен:
   - собрать request;
   - вызвать capacity calculator;
   - сохранить `CCapacityResult`;
   - не разбирать детали search-а.
5. `CapacityState` должен попадать в `CStateRepository` только как корректно полученный итоговый named state.
6. Execution report должен брать diagnostic из `CLimitSearchResult`, а не из внутренних строк `CCapacitySolver`.

Тесты:

- полное совпадение подробной таблицы прочности;
- совпадение batch summary reserve по capacity;
- совпадение exported/visualized `CapacityState`;
- solver-call count не растет без причины.

## Этап 6. Перестройка Crack

Цель: разделить образование трещины, текущее трещиноватое НДС, формулу ширины и продольные трещины.

Шаги:

1. Создать `CCrackFormationResult`.
2. Создать `CCrackWidthResult`.
3. Создать `CLongitudinalCrackResult`.
4. Создать общий `CCrackResult`.
5. Перенести поиск образования трещины в `CCrackFormationCalculator`.
6. Использовать общий `CLimitSearchCoordinator`.
7. Переименовать:
   - `BeforeMcrcState` -> `PreCrackState`;
   - `AfterMcrcState` -> `PostCrackState`.
8. `CurrentCrackedState` считать только если он нужен хотя бы одному активному потребителю:
   - проверке продольных трещин;
   - расчету ширины нормальной трещины после подтвержденного образования нормальной трещины;
   - будущему выводу/визуализации, если она явно запрошена для этого named state.
9. Если полностью эквивалентный и корректно полученный `CurrentCrackedState` уже есть в `CStateRepository`, переиспользовать его, а не решать повторно.
10. Если нужный `CurrentCrackedState` не найден:
   - state сохраняет фактический `InternalStatus` и `ResultCode` причины: `rsNumericalFailure`, `rsInvalidInput`, `rsInvalidConfiguration`, `rsInternalError` или другой применимый внутренний статус;
   - `rsNumericalFailure` назначается только при реальной численной несходимости, а не при любом неуспехе получения state;
   - width и longitudinal results получают внутренний `rsBlockedByDependency` с кодом зависимости от `CurrentCrackedState`;
   - общий crack result агрегирует failure из state-result через `ResultMeta`, а внешний `NumFail` может появиться только через `CResultStatusPolicy`.
11. Если formation search не сошелся:
   - `FormationResult.ResultMeta.InternalStatus = rsNumericalFailure`;
   - width получает `rsBlockedByDependency`, если для него обязательна точка formation;
   - longitudinal crack может быть посчитан только при наличии `CurrentCrackedState`.
12. Если formation search корректно вернул `rsSuccess + CRITERION_NOT_REACHED`:
   - `FormationResult.CrackFormed = False`;
   - расчет ширины нормальной трещины получает `rsNotApplicable`;
   - это не `FAIL` и не `NumFail`;
   - продольная проверка остается отдельным потребителем `CurrentCrackedState`, если она активна.
13. Если `sigma_s,crc <= 0` при найденном `PostCrackState`:
   - это не numerical failure формулы;
   - `psi_s` принимается 1;
   - `ResultCode` должен объяснить, что `sigma_s,crc` неположительное и уточнение `psi_s` неприменимо.
14. Формула раскрытия должна работать только по готовым численным данным. Цепочка должна быть такой:
   - `CurrentCrackedState` / `PostCrackState`;
   - подготовка section/crack data вне формулы;
   - готовые `Abt`, `As`, `ds_eq`, `ls`, `sigma_s`, `sigma_s,crc`, `psi_s`, `phi1`, `phi2`, `phi3` и другие коэффициенты;
   - `CCrackWidthFormulaCalculator`;
   - результат `a_crc` и связанные формульные величины.
15. `CCrackWidthFormulaCalculator` не должен получать State-объекты, выбирать tension zone, выбирать стержни, считать `Abt/As/ds/ls/sigma_s` из геометрии или запускать solver.
16. Удалить скрытые solve-вызовы из формулы ширины.

Тесты:

- `CurrentCrackedState` создается только когда нужен хотя бы одному активному потребителю;
- нормальная трещина не образуется -> width result получает `rsNotApplicable`, а внешний `N/A` формируется только policy; longitudinal при активной проверке получает и использует `CurrentCrackedState`;
- normal crack formed -> width считается;
- `psi_s User`;
- `psi_s Auto`;
- `psi_s AlwaysCalc`;
- `SLS.Crack.SigmaSCrcAveragingMode`;
- круг, RectSet, RoundedRectangle, HollowRectangle;
- слабое армирование, где раньше мог возникать `NumFail` в неправильном столбце.

## Этап 7. Перестройка `CCombinationResult`

Цель: заменить плоский набор полей структурированным деревом.

Шаги:

1. Ввести вложенные result-объекты:
   - `CStrengthResult`;
   - `CCapacityResult`;
   - `CCrackResult`;
   - `CStabilityResult`;
   - `CDirectStateResult`, если нужен отдельный объект.
2. Перенести данные из плоских полей в соответствующие объекты.
3. Оставить временные read-only getters для writer-ов, если это снизит риск.
4. После миграции writer-ов удалить legacy getters, которые больше не нужны.
5. `CCombinationResult` должен:
   - хранить input identity LC;
   - хранить `StateRepository`;
   - хранить result tree;
   - хранить overall `ResultMeta`;
   - не решать инженерные проверки.
6. Worst-combination logic должна читать reserves/statuses из result tree.

Тесты:

- batch summary worst LC не меняется;
- пустые строки `rngLoadCombinations` остаются пустыми в outputs;
- skipped calculations получают `rsNotRequested`, а внешний `N/A` формируется через `CResultStatusPolicy`;
- все reserves совпадают с detailed output.

## Этап 8. Пассивные writer-ы

Цель: writer-ы не должны содержать инженерной логики.

Шаги:

1. Обновить `CBatchResultWriter`.
2. Обновить `CStrengthSummaryWriter`.
3. Обновить `CCrackSummaryWriter`.
4. Обновить `CStabilitySummaryWriter`.
5. Обновить `CNDMResultsWriter`.
6. Добавить `ResultComment` перед общим статусом каждого расчетного блока.
7. Все status cells должны краситься через `CResultStatusPolicy`.
8. Writer-ы должны брать:
   - internal status из `ResultMeta.InternalStatus`;
   - result code из `ResultMeta.ResultCode`;
   - comment из `ResultMeta.ResultComment`;
   - numbers из result-объектов.
9. Убрать из writer-ов:
   - парсинг stop reason;
   - расчет reserves;
   - выбор OK/FAIL;
   - fallback-решения.

Тесты:

- layout Results/Summary совпадает с baseline плюс `ResultComment`;
- цвета статусов совпадают;
- `N/A` выводится только в применимых местах;
- строки с пропущенными LC остаются пустыми;
- status legend работает.

## Этап 9. Execution report и diagnostics

Цель: отчет остается человекочитаемым, но берет данные из result graph.

Шаги:

1. Убрать из `CBatchSectionCalculator` низкоуровневые переводчики diagnostic, если они станут дублировать новую систему.
2. Каждый слой пишет diagnostic в свой result:
   - state runner;
   - limit search;
   - capacity;
   - crack formation;
   - crack width;
   - stability.
3. Batch report только объединяет блоки в правильном порядке.
4. Сохранить подробные итерации solver/search для анализа.
5. Перевести технические фразы на понятный русский там, где они видны пользователю.

Тесты:

- отчет содержит отдельные блоки Stability, Strength, Crack;
- в fallback Capacity/Crack видно, какой метод пробовался и почему произошел переход;
- `ResultComment` в Results согласован с report.

## Этап 10. Очистка legacy и финальная проверка

Шаги:

1. Поиск и удаление старых имен:
   - `BeforeMcrcState`;
   - `AfterMcrcState`;
   - старые status strings, если они заменены codes;
   - устаревшие helper-и, которые считают status локально.
2. Поиск прямых запусков `CSectionSolver` вне разрешенных мест:
   - `CStateSolutionRunner`;
   - специальные низкоуровневые search-операции, где нужен не reusable state, а прямой solver/evaluator вызов вроде `EvaluateStrainPlane`;
   - тестовые helpers.
3. Проверить, что crack width formula не запускает solver.
4. Проверить, что writer-ы не запускают solver и не считают инженерные проверки.
5. Обновить `AGENTS.md`, если новые архитектурные правила должны стать правилом проекта.
6. Обновить `docs/CoordinateSystem.md`, если меняются имена state или комментарии.
7. Обновить справку Excel.
8. Удалить временные compatibility aliases, если они не нужны для пользовательских входных данных.

Финальные проверки:

```powershell
rg "BeforeMcrcState|AfterMcrcState" src tests tools docs
rg "New CSectionSolver" src
rg "CrackStatusFromCalculator|CapacityStatusFromSolver" src
rg "StopReason" src\Excel src\Batch
rg "ResultComment" src tests tools
```

---

## 7. Миграция конкретных текущих классов

| Текущий класс | Целевая роль | Действие |
|---|---|---|
| `CSectionSolver` | Низкоуровневый solver равновесия | Сохранить. Добавлять только metadata/counters при необходимости. |
| `CStateSolutionRunner` | Общий runner одного state | Сохранить и усилить. Он становится единственным штатным входом для named state. |
| `CSectionStateResult` | Универсальный result одного НДС | Расширить. Добавить target/internal loads, residuals, ResultMeta. |
| `CCombinationResult` | Корень результата LC | Перестроить в result graph, добавить StateRepository. |
| `CBatchSectionCalculator` | Оркестратор LC | Существенно облегчить: он вызывает сервисы, но не владеет search/status/formula. |
| `CBatchStatusPolicy` | Status policy | Развить или заменить на `CResultStatusPolicy`. |
| `CCapacitySolver` | Сейчас search + capacity | Разделить. Generic search вынести; capacity оставить как инженерный facade/result. |
| `CCrackWidthCalculator` | Сейчас formation + width + psi + state handling | Разделить на formation, current state, width formula, longitudinal check. |
| `CStabilityCalculator` | Независимая инженерная проверка | Подключить к ResultMeta, без большой перестройки. |
| `CLoadPathVector` | Универсальный vector path | Сохранить. Использовать и для Capacity, и для CrackFormation. |
| `CLoadPathMath` | Остатки/нормы load path | Сохранить и использовать в общем LimitSearch. |
| `CSectionLoadState` | Централизованный перенос нагрузок | Сохранить. Не дублировать переносы в новых классах. |
| `CMomentZeroFilter` | Zero moment filter | Сохранить. Вызывать через load/state request construction. |
| `CStrengthSummaryWriter` | Пассивный writer | Перевести на result graph. |
| `CCrackSummaryWriter` | Пассивный writer | Перевести на `CCrackResult`. |
| `CStabilitySummaryWriter` | Пассивный writer | Перевести на `CStabilityResult`. |
| `CBatchResultWriter` | Пассивный summary writer | Перевести на result graph и `CResultStatusPolicy`. |
| `CNDMResultsWriter` | Пассивный snapshot writer | Проверить, что named states читаются из repository/result graph. |

---

## 8. Предлагаемый порядок commits

Чтобы проверка на стороне была проще, рефакторинг лучше делать серией небольших архитектурных commits:

1. `Introduce ResultMeta and status policy foundation`
2. `Extend SectionStateResult metadata`
3. `Add StateRequest and StateRepository`
4. `Route direct states through StateProvider`
5. `Extract generic LimitSearch from capacity`
6. `Move capacity to LimitSearch`
7. `Move crack formation to LimitSearch`
8. `Split crack width formula from state/search`
9. `Restructure CombinationResult`
10. `Make Results writers passive`
11. `Remove legacy state/status paths`
12. `Update help, tests and workbook`

Каждый commit должен проходить как минимум targeted-тесты затронутого слоя. Полный набор тестов обязателен после commits 4, 6, 8, 10 и 12.

---

## 9. Набор regression-тестов

### 9.1. State tests

Проверить:

- `StrengthState` для стандартного Г-сечения;
- `CurrentCrackedState` для Г-сечения и круглого сечения;
- `PreCrackState/PostCrackState`;
- state reuse без повторного solver call;
- разные material spec не смешиваются;
- разные load target не смешиваются;
- `ExtensionUsed` корректно сохраняется.

### 9.2. Capacity tests

Проверить:

- `UltimateStrain`;
- `LoadMultiplier`;
- `Auto` с успешным UltimateStrain;
- `Auto` с fallback на LoadMultiplier;
- чистое сжатие;
- чистое растяжение;
- чистый изгиб;
- `N + Mx`;
- `N + Mx + My`;
- `λ*N`;
- `λ*Mxy`;
- `λ*NMxy`;
- случаи, где moment zero filter влияет на путь.

Сравнивать:

- `lambda_ult`;
- `Nult`;
- `Mxult`;
- `Myult`;
- `CapacityState` plane;
- `InternalStatus`, `ResultCode`, `ResultComment`;
- solver-call count в пределах ожидаемой разницы.

### 9.3. Crack tests

Проверить:

- `SLS.Crack.InitiationLoadPath = Auto`;
- `λ*Mxy`;
- `λ*N`;
- `λ*NMxy`;
- `SLS.Crack.InitiationSolutionStrategy = UltimateStrain`;
- `LoadMultiplier`;
- `Auto`;
- `PsiMode = User`;
- `PsiMode = Auto`;
- `PsiMode = AlwaysCalc`;
- `SigmaSCrcAveragingMode`;
- центральное растяжение;
- внецентренное растяжение;
- сжатие без нормальной трещины, но с продольной проверкой;
- слабое армирование;
- круг, RectSet, RoundedRectangle, HollowRectangle.

Особые assertions:

- `NumFail` появляется только в `FormationResult`, `CurrentCrackedState` или другом state/search result;
- `CCrackWidthResult` / orchestration-слой при отсутствии обязательных подготовленных данных получает `rsBlockedByDependency` или `rsNotApplicable`, а `CCrackWidthFormulaCalculator` в такой ветке вообще не вызывается;
- `sigma_s,crc <= 0` не превращается в numerical failure;
- `CurrentCrackedState` не решается повторно, если корректно полученный эквивалентный state уже есть в repository.

### 9.4. Stability tests

Проверить:

- СП 35;
- СП 63;
- ветви `ec > r` и `ec <= r`;
- обе главные плоскости;
- случай, где stability FAIL не прерывает остальные расчеты;
- reserves в batch summary совпадают с detailed output.

### 9.5. Writer/layout tests

Проверить:

- `rngBatchSummary`;
- `rngStrengthSummaryAnchor`;
- `rngCrackSummaryAnchor`;
- `rngStabilitySummaryAnchor`;
- `rngNDMElementResults`;
- `rngNDMSectionGeometry`;
- `rngNDMSectionProperties`;
- `rngNDMSectionAnnotations`.

Особенно:

- ResultComment columns на месте;
- пустые строки LC сохраняют позицию;
- skipped LC пустые;
- `N/A` только там, где расчет не применялся;
- цвета статусов соответствуют словарю;
- writers не меняют расчетные результаты.

---

## 10. Риски и меры защиты

### Риск 1. Случайное изменение математики

Мера:

- сначала baseline;
- после каждого этапа сравнивать численные результаты;
- все изменения search-а делать через parity-тесты.

### Риск 2. Неправильный reuse state

Мера:

- строгий `CStateRequest`;
- ключ включает material spec, target loads, state type и extension policy;
- warm-start, initial guess, initial solver и retry policy остаются solve-options и не входят в ключ reuse, если не меняют физический результат или допустимость;
- тесты на близкие, но разные состояния.

### Риск 3. Потеря probe cache и падение производительности

Мера:

- до извлечения search-а описать текущий cache `CCapacitySolver`;
- перенести cache в search-session;
- сравнить solver-call count и время до/после.

### Риск 4. Неправильная классификация `NumFail`

Мера:

- запретить формульным проверкам создавать `NumFail`;
- `NumFail` только из state/search, где реально запускался solve;
- добавить тесты для crack width и longitudinal crack.

### Риск 5. Writer начнет принимать инженерные решения

Мера:

- code review по writer-ам;
- writer допускает только formatting, units, colors, array output;
- все status/reserve/comment должны быть готовы до writer-а.

### Риск 6. Переименование states сломает plot/export

Мера:

- мигрировать `modCalculationPurpose` централизованно;
- обновить profile visualization settings;
- обновить tests workbook interface;
- проверить Excel plot и AutoCAD export по saved Results.

---

## 11. Проверочные grep-аудиты после реализации

После завершения миграции выполнить:

```powershell
rg "BeforeMcrcState|AfterMcrcState" src tests tools docs
rg "New CSectionSolver" src
rg "\.Solve " src
rg "CrackStatusFromCalculator|CapacityStatusFromSolver" src
rg "StopReason" src\Excel src\Batch
rg "NumFail" src\Crack src\Stability src\Excel
rg "ResultComment" src tests tools docs
rg "CStateRepository|CStateRequest|CLimitSearch" src tests
```

Ожидаемый смысл:

- старые crack-state имена отсутствуют или только документированы как migration note;
- прямой solver создается только в `CStateSolutionRunner` и search internals;
- writer-ы не парсят `StopReason`;
- `NumFail` не рождается в формульных проверках;
- новые result/status/search классы покрыты тестами.

---

## 12. Критерии готовности

Рефакторинг можно считать завершенным, если выполнены все условия:

1. Все существующие regression/integration tests проходят или обновлены только из-за сознательно измененных внутренних имен.
2. Численные результаты baseline совпадают в пределах существующих tolerance.
3. `Capacity` и `CrackFormation` используют общий `LimitSearch`.
4. Named states получаются через общий `StateProvider`.
5. `CStateRepository` хранит только конечные reusable states.
6. Lambda probes не попадают в общий repository.
7. `CCombinationResult` хранит структурированное дерево результатов.
8. Каждый расчетный result имеет `InternalStatus`, `ResultCode`, `ResultComment`.
9. `NumFail` появляется только у state/search-результатов, где действительно решалось равновесие или предельная точка.
10. Writer-ы пассивны и не содержат инженерной логики.
11. Results/Summary визуально сохраняют прежний смысл и дополнены `ResultComment`.
12. Execution report остается подробным и человекочитаемым.
13. `BeforeMcrcState/AfterMcrcState` заменены на `PreCrackState/PostCrackState`.
14. Документация и справка Excel соответствуют новой архитектуре.

---

## 13. Рекомендуемая первая реализационная задача

Не начинать с полного переписывания `CCrackWidthCalculator`.

Правильный первый шаг:

```text
CResultMeta + CResultStatusPolicy + расширение CSectionStateResult
```

Причина: это дает общий язык результата и позволит переносить Capacity/Crack по частям, не создавая очередной слой временных строковых статусов.

Второй шаг:

```text
CStateRequest + CStateRepository + CStateProvider
```

Только после этого безопасно извлекать `LimitSearch` и перестраивать Crack.

---

## 14. Финальная сверка migration plan с v6 перед началом реализации

Перед тем как начинать кодовый рефакторинг, пройти этот migration plan против `docs/NDM_Architecture_Refactoring_Plan_Codex_v6.md` по разделам.

### State

Проверить:

- `CSectionStateResult` остается универсальной сущностью конечного НДС;
- `PreCrackState`, `PostCrackState`, `CurrentCrackedState` заменяют старые узкие имена;
- `CStateRepository` хранит только конечные named states;
- `CStateRequest` разделяет параметры физического state-запроса и solve-options;
- ключ эквивалентности state не включает warm-start, initial guess, initial solver и retry policy, если они не меняют физический результат или его допустимость;
- lambda probes search-а не попадают в `CStateRepository`;
- `CurrentCrackedState` создается только когда нужен активному потребителю и переиспользуется при эквивалентном запросе.

### Search

Проверить:

- Capacity и CrackFormation используют общий `LimitSearch`;
- search algorithms отделены от physical criterion;
- `INITIAL_STATE_BEYOND_LIMIT`, `CRITERION_NOT_REACHED`, `SEARCH_BOUND_REACHED` различаются строго по смыслу v6;
- `INITIAL_STATE_BEYOND_LIMIT` не является `NumFail`;
- `CRITERION_NOT_REACHED` может быть `rsSuccess` search-результатом, если недостижение критерия корректно установлено;
- для `CrackFormation Auto` `INITIAL_STATE_BEYOND_LIMIT` может быть сигналом попробовать другой load path;
- для фиксированного CrackFormation path это физический результат выбранной траектории;
- локальный probe-cache search-session сохранен и не смешан с state repository.

### Engineering

Проверить:

- Capacity превращает `rsCheckFailed + INITIAL_STATE_BEYOND_LIMIT` во внешний `BaseFail` через централизованную policy;
- Crack разделен на formation, подготовку section/tension-zone data, pure formula и longitudinal consumer;
- `CCrackWidthFormulaCalculator` не выбирает tension zone, стержни, `Abt`, `As`, `ds`, `ls`, `sigma_s` и не запускает solver;
- Stability подключена к ResultMeta, но ее расчетная методика не меняется;
- ни один инженерный расчет не парсит display-строки статусов.

### Result + Presentation

Проверить:

- у существенных result entities есть `InternalStatus`, `ResultCode`, `ResultComment`;
- внешний display-словарь содержит `OK`, `FAIL`, `BaseFail`, `NumFail`, `InputErr`, `CalcErr`, `N/A`;
- цепочка `InternalStatus -> ResultCode -> CResultStatusPolicy -> ExternalStatus` не обходится инженерными расчетами и writer-ами;
- writer-ы пассивны и не принимают инженерных решений;
- `ResultComment` централизованно собирается из соответствующего result-subtree без дублей и добавляется в согласованные блоки Results/Summary;
- enum/status/result-code declarations и инженерные трактовки generic кодов имеют русские комментарии в коде.

### Git long-horizon workflow

Проверить:

- baseline commit hash записан в progress/status MD;
- progress/status MD короткий и обновляется после значимых этапов;
- после compaction восстановление идет через `git status`, `git diff`, `git log`, baseline, architecture MD, migration plan и progress/status MD;
- destructive reset / checkout / rewrite history запрещены без явной необходимости;
- посторонние пользовательские изменения не откатываются.

Если по любому пункту обнаружено противоречие с v6, сначала исправить migration plan или реализационный plan-step, и только потом продолжать код.
