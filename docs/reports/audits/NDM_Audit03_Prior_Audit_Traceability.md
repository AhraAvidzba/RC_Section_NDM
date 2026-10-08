# Audit03: Повторная Трассировка Audit01 И Audit02

Исходный реестр: 2026-10-02. Финальная приемка: 2026-10-05, раздел
«Итоговая По-ID Сверка v328» ниже: все 70 требований двух прежних аудитов.
Начальные таблицы и их Pending сохраняются как история, не как текущие TODO.
Исторические отчеты не изменены. Новые контрпримеры F01-F07 и publication/P01/P02
проверяются отдельно, даже если прежний тест остается зеленым.

## Обозначения Логов

- F29: `docs/regression/Audit03/p01_full_off_v29_2026-10-02.txt`, 8390/0.
- N29: `docs/regression/Audit03/p01_full_on_v29_2026-10-02.txt`, 8391/0.
- SC29: `docs/regression/Audit03/source_contracts_v29_2026-10-02.txt`, 101/101, 0 ошибок.
- PUB: `a03_snapshot_negative_v19_2026-10-02.txt` 114/33 и positive_v19 147/0.
- RETRY: `p01_retry_session_v29_2026-10-02.txt`, 22/0.
- P4: `performance_p02_v4_2026-10-02.txt`, 150 измерений, пять повторов, failed=0.

Все runtime-логи получены в Excel COM на отдельных копиях; source unchanged=True.
Обычный Off не отменяет явно указанный On внутри исторического fixture.
`runtime PASS` означает конкретный перечисленный тест, а не автоматический
семантический PASS всего требования. Указанные здесь Pending-гейты были
обязательны на дату исходного реестра; итоговые доказательства приведены ниже.

## Первый Аудит

| ID | Актуальный Владелец / Метод | Test-ID | Лог / Результат / Граница |
| --- | --- | --- | --- |
| T01 | CLoadMultiplierSearch, финализация Secant | TestLimitSearchSecantFinalizesCheckedRoot / limitSearch.secant.root | F29/N29, runtime PASS; root=0.75. |
| T02 | CLoadMultiplierSearch, iteration budget | TestLimitSearchBisectionIterationLimitFails / limitSearch.bisection.noFakeLimit | F29/N29, runtime PASS. |
| T03 | Capacity ProbeOnce / typed initial result | TestInitialLambdaFailureStatusMapping; audit02.failedProbe.noBaseFail | F29/N29, runtime PASS; NumFail без выдуманного physical failure. |
| T04 | LoadMultiplier bounds; Formation result interpretation | TestCrackFormationSearchBoundKeepsTechnicalCode; audit02.genericMultiplier.*.boundCode | F29/N29, runtime PASS. |
| T05 | Formation finalization / named states | TestCrackFormationNoCrackDoesNotBuildPostState | F29/N29, runtime PASS; нет фиктивного Post. |
| T06 | CStateProvider.SolverSnapshot / Runner.RestoreSnapshot | TestCrackFormationCacheHitWithoutLastRunner; audit02.currentCache.noHeavySolve | F29/N29, runtime PASS. |
| T07 | CStateRepository.FindEquivalent / StoreState | stateRepository.failedNotReusable / retrySuccessReusable; TestAudit02RepositoryContextAndRetry | F29/N29, runtime PASS; PUB дополняет защитой публикации. |
| T08 | CCrackResult.ResultMeta / CrackSummary writer | TestCrackAggregateIncludesCurrentStateFailure; combinationTree.crack.aggregate.currentStateFailure | F29/N29, runtime PASS; все load-case comments дополнительно проверяются в Audit03 matrix. |
| A01 | CLimitSearchCoordinator, CLoadMultiplierSearch, CUltimateStrainSearch | TestAudit02GenericUltimateSearch; TestAudit02GenericLoadMultiplierMatrix | F29/N29/SC29, runtime/source PASS; один generic алгоритм обоих domains. |
| A02 | CCapacitySolver.ProbeOnce / CStateSolutionRunner | TestAudit02UnconvergedProbeIsNumerical; linear material tests | F29/N29, runtime PASS; RETRY сохраняет общий pipeline и локальный scope. |
| A03 | CCrackFormationCalculator.CheckFormation | TestAudit02IndependentFormation | F29/N29, runtime PASS; Formation не обертка Width. |
| A04 | CLimitSearchResult.Initialize / meta/state snapshots | TestLimitSearchResultKeepsCrackDiagnosticSnapshot; searchEngineering.* | F29/N29, runtime PASS; PUB защищает named state. |
| A05 | Существующие writers; готовые result-subtrees | TestBatchSummaryWriter; TestAudit03LoadPathComments | F29/N29, runtime PASS; report/save-reopen полной новой матрицы Pending. |
| A06 | CResultStatusPolicy.ExternalStatus; владельцы ResultComment | TestResultMetaStatusDictionary; TestCombinationResultTreeDrivesDisplayFields | F29/N29, runtime PASS; D01/D02 общий review Pending. |

## Audit02: Корректность И Архитектура

| ID | Актуальный Владелец / Метод | Test-ID | Лог / Результат / Граница |
| --- | --- | --- | --- |
| R01 | Provider.SolverSnapshot; Runner.RestoreSnapshot; Width.Calculate | TestAudit02CurrentCrackedStateCacheHitCalculatesWidth / audit02.currentCache.* | F29/N29, runtime PASS. |
| R02 | LoadMultiplier FindBracket / MaxLambda / last interval | TestAudit02CapacitySearchBoundary; audit02.genericMultiplier.* | F29/N29, runtime PASS; F06 arithmetic adds precision limits. |
| R03 | Capacity ProbeOnce / physical confirmation | audit02.failedProbe.* / failedPositive.* / offsetPair.* | F29/N29, runtime PASS; P01 не кеширует physical failure несошедшейся пробы. |
| R04 | ILimitSearchProblem / generic searches / independent result | TestAudit02GenericUltimateSearch; TestAudit02GenericLoadMultiplierMatrix | F29/N29/SC29, runtime/source PASS. |
| R05 | Formation.CheckFormation / собственный контекст | TestAudit02IndependentFormation; TestAudit02FormationOutcomeSemantics | F29/N29, runtime PASS. |
| R06 | Runner flags / State.InitializeFromSolver / ConfirmEquilibrium | audit02.plane.*; TestAudit02StateSnapshotIsolation | F29/N29, runtime PASS; F04 lifecycle directed adds attempted=false. |
| R07 | Combination/Strength/Crack canonical owners | TestAudit02CanonicalResultsAndReset / audit02.canonical.* | F29/N29/PUB, runtime PASS; public field/setter review SC29. |
| R08 | Capacity sign-specific SteelCompression/TensionLimit | TestAudit02AsymmetricSteelLimits / audit02.steelSign.* | F29/N29, runtime PASS, оба знака и оба search approaches. |
| R09 | Sequential guards / входные контракты | TestAudit03TypedStateFailures; TestAudit03CapacityTypedFailures | F29/N29, directed PASS; все 1205 guard-кандидатов F07 еще review Pending. |
| A01 | Существующие owners; без новых .cls | noNewClasses / twoSpecifiedMerges | SC29, PASS; Audit03 явно предписывает два объединения, 85 -> 83 production. |
| A02 | Batch orchestration / report / governing; typed results | TestAudit02CanonicalResultsAndReset; TestBatchSummaryWriter | F29/N29, runtime PASS; all-consumer ownership review продолжается. |
| A03 | Formation отдельно; Width preparation и чистые численные методы | TestAudit02IndependentFormation; TestCrackWidthFormulaCalculatorPure | F29/N29/SC29, PASS; отдельный Formula class удален по Audit03 A02. |
| A04 | Capacity domain callbacks и общий state runner | TestAudit02AsymmetricSteelLimits; TestCapacityLoadPathMethodMatrix | F29/N29, runtime PASS. |
| A05 | Shared numerical loops / generic Newton | audit02.genericUltimate.* / genericMultiplier.* | F29/N29/SC29, PASS; F06 добавляет finite-step guards. |
| A06 | CLimitSearchResult snapshot без live solver | TestLimitSearchResultKeepsCrackDiagnosticSnapshot; searchEngineering.* | F29/N29/PUB, runtime PASS. |
| A07 | Только canonical numerical fields / readonly publication | audit02.canonical.*; audit03.published.* | F29/N29/PUB, runtime PASS; новый Execute не меняет старые ссылки. |
| A08 | Пассивные writers; арифметика готовых величин | TestBatchSummaryReserveConsistencyTest; audit03.comments.* | F29/N29/SC29, runtime/source PASS; полная UI приемка Pending. |

## Audit02: Extension И Статусы

| ID | Актуальный Владелец / Метод | Test-ID | Лог / Результат / Граница |
| --- | --- | --- | --- |
| E01 | SettingsReader / Provider / SettingsCatalog, General.DiagramExtension | TestAudit02DiagramExtensionReaderMigration / audit02.migration.reader.* | F29/N29, runtime PASS; actual catalog есть в Config census. |
| E02 | Обновление сохраненной книги до defaults | migration tests предыдущего Audit02; reader priority tests | Reader F29/N29 PASS; повтор physical migration новой итоговой книги Pending. |
| E03 | Provider effective specs по ролям State/Search | TestAudit02OnOffPhysicalResults; TestCapacityLoadPathMethodMatrix; TestAudit03LoadPathComments | F29/N29, runtime PASS; широкая shape/material/setting matrix Pending. |
| E04 | MaterialDiagram Ignore / extension stress+tangent | TestProviderStateSolutionExtension; ignored tension assertions | F29/N29, runtime PASS. |
| E05 | Provider physical nodes / plateau / eps_ult | AssertPhysicalDiagramPair; material diagram suites | F29/N29, runtime PASS; expected/tolerance прежние. |
| E06 | MaterialDiagram safe technical outer limit | TestAudit02ExtensionBeyondTechnicalDefault; TestAudit02ExtensionOverflowIsExplicit | F29/N29, runtime PASS. |
| E07 | State Converged/WithinPhysicalRange/ExtensionUsed, permission | TestAudit02ExtendedInitialGuessPhysicalFinal; audit02.offsetPair.* | F29/N29, runtime PASS. |
| E08 | Capacity/Formation physical finalization; blocked current dependencies | audit02.failedProbe.*; TestCrackAggregateIncludesCurrentStateFailure | F29/N29, runtime PASS. |
| E09 | Один физический критерий обоих подходов | TestCapacitySolutionStrategyComparisons; TestCapacityLoadPathMethodMatrix | F29/N29, runtime PASS. |
| E10 | Реальное Off и явно отдельные On fixtures | TestAudit02OnOffPhysicalResults; mode-after-suite | F29/N29, runtime PASS; новые all-path кейсы логируют effective permission. |
| S01 | InternalStatus/ResultCode -> ResultStatusPolicy -> ExternalStatus | resultMeta.*; combinationTree.* | F29/N29/SC29, runtime/source PASS; no text-based mapping. |
| S02 | Capacity initial offset, только подтвержденный physical criterion | TestAudit02InitialOffsetBoundary; TestAudit02InitialOffsetOnOffStatuses | F29/N29, runtime PASS. |
| S03 | Formation initial crack, fixed/Auto, warning/psi1 | TestAudit02FormationOutcomeSemantics; TestAudit02PsiSignedInputsAndFallbackModes | F29/N29, runtime PASS. |
| S04 | CriterionNotReached / SearchBoundReached / later-than-current | TestCrackFormationNoCrackDoesNotBuildPostState; TestCrackFormationSearchBoundKeepsTechnicalCode | F29/N29, runtime PASS. |
| S05 | CurrentCrackedState точная primary причина / blocked dependents | TestCrackAggregateIncludesCurrentStateFailure; audit03.comments.*BlockActualReason | F29/N29, runtime PASS; F03 точные terminal codes сохранены. |
| S06 | Width Psi User/Auto/AlwaysCalc / sigma_s,crc | TestAudit02PsiSignedInputsAndFallbackModes; TestSigmaSCrcAveragingModeAllSelected | F29/N29, runtime PASS; L15 broader boundaries Pending. |
| S07 | Result-subtree comments во всех output blocks | TestCombinationResultTreeDrivesDisplayFields; TestAudit03LoadPathComments | F29/N29, runtime PASS; новой широкой matrix report/save-reopen Pending. |
| C01 | Единый StateSolutionRunner; общий material object API | Section linear tests; TestAudit02AsymmetricSteelLimits | F29/N29, runtime PASS. |
| C02 | State equivalence: context/revision/spec/permission/tolerance | TestAudit02RepositoryContextAndRetry / audit02.cache.* | F29/N29, runtime PASS. |
| C03 | Named Repository != local search probes; failed retry allowed | stateRepository.*; TestAudit02RepositoryContextAndRetry | F29/N29/RETRY, runtime PASS; probe reuse сохранен P4. |
| C04 | Converged не означает пригодный engineering State | TestAudit02EvaluatedPlaneRequiresEquilibrium; current/offsetPair tests | F29/N29, runtime PASS. |
| C05 | Results snapshot самодостаточен / Config не вызывает пересчет | TestAudit02SavedResultsIgnoreMaterialChanges; TestProfileDrivenPlotUsesSnapshotState | F29/N29, runtime PASS; save/reopen новой конечной output-книги Pending. |
| C06 | Clear/Initialize/meta/следующий solve изолированы | TestAudit02CanonicalResultsAndReset; TestAudit02StateSnapshotIsolation; audit03.published.* | F29/N29/PUB/RETRY, runtime PASS. |

## Audit02: Приемка, Документация И Workflow

| ID | Актуальное Доказательство | Результат / Оставшийся Gate |
| --- | --- | --- |
| Q01 | Progress baseline SHA; baseline build/Off/On 6745/0 | PASS исходного воспроизведения. |
| Q02 | F29, штатные numerical expected/tolerance | runtime PASS; отдельный final preservation gate еще нужен. |
| Q03 | N29/F29: historical explicit-On setup не переписан | runtime PASS; mode-after-suite восстановлен. |
| Q04 | TestAudit02OnOffPhysicalResults и component assertions | F29/N29 PASS текущих пар; широкие новые пары Pending. |
| Q05 | Старые 28 directed rows в Audit02 progress; F29/N29 | Повторенные тесты PASS; новая Audit03 K/T матрица Pending, старый счетчик не подменяет ее. |
| Q06 | Все T01-T08/A01-A06 строки выше | directed runtime PASS; ограничения отмечены по строкам. |
| Q07 | Все восемь suites F29/N29; negative reproducer logs | PASS текущего среза; финальная отдельная clean build Pending. |
| Q08 | P4; saved snapshot tests; A03 publication; native AutoCAD v284/v286 | runtime PASS перечисленного; настоящий AutoCAD shape roundtrip 1023/0, 22 случая. Пиксельная приемка Excel и итоговая performance/reopen еще Pending. |
| D01 | Code census / Config census / current architecture owners | Содержательный D01/D02 review и итоговая help-книга Pending. |
| D02 | General permission != use != physical limit, material tests | Runtime PASS; фактическая итоговая справка/validation Pending. |
| W01 | df10412f baseline, hashes original book/spec/export | PASS; пользовательская output-книга не менялась. |
| W02 | Audit03 Progress, ownership table, coverage и этот реестр | В работе, каждый незавершенный gate явно отмечен. |
| W03 | Checkpoints 1a65796, 3d0112e, 210a5a20 и directed/full logs | PASS существующих checkpoints; следующие срезы после gate. |
| W04 | Git/status/diff/log, Progress восстановление, запрет destructive Git | Правило действует; push/откаты чужих изменений не выполнялись. |
| W05 | Active goal, все шесть Audit03 gates | Pending; цель не объявлена завершенной на узком green suite. |
| W06 | Итоговые report/book/export/self-audit | Pending; Audit03 Final Report пока не выпускается. |

## Граница Среды

Excel runtime, COM-чтение, DisplayFormat и сохранение доступны. Native
accessibility вернул null, native screenshot завершился timeout; это не PASS
визуального оформления. Историческое ограничение по CAD снято отдельными
native gates: настоящий Autodesk AutoCAD 2023 подтвердил импорт/экспорт
Region и запись DWG. Shape roundtrip v284/v286 `1023/0`, 22 случая, включая
квадрат по грани и круг по среднему остальных ориентированных элементов.
Подробные ограничения, независимые oracle и файлы evidence сохранены в
`NDM_Audit03_AutoCAD_Config_Review.md`. Эти native gates не заменяют
финальные release/self-audit и пиксельную проверку оформления Excel.

## Итоговая По-ID Сверка v328, 2026-10-05

Финальный выпуск принят: full Off `108901/0`, full On `108912/0`, по восемь
suites, watchdog exit=0 и source unchanged. Исторические Pending выше
сохраняют значение на дату того среза. Все 70 строк ниже проверены текущими
runtime/source/directed/native/reopen evidence в явно указанной области.
Глобальный self-audit и контролируемая публикация завершены в Final Report.
Владельцы и Test-ID находятся в соответствующих исторических строках выше.

Обозначения актуальных доказательств:

- F328/N328: `full_off_release_v328b_2026-10-05.txt` /
  `full_on_release_v328_2026-10-05.txt`, восемь штатных suites.
- SC328: `source_contracts_release_clean_v328_2026-10-05.txt` и
  `source_contracts_saved_release_v328_2026-10-05.txt`, по 108/108, failed=0.
- MATRIX: `load_matrix_final_v320_acceptance_2026-10-05.json`, 52 принятых
  runs, 13 форм, 949 form/load fixtures, 22776 path cases, 957531/0.
- SELECTOR: `selector_summary_v306_full_2026-10-05.json`, 62 группы,
  4464 path cases, 302 пары, 48 рискованных четверок; пять negative mutations.
- CONFIG: `config_final_evidence_v328_2026-10-05.json` и
  `NDM_Audit03_Final_Config_Review.md`, точный минимум K02, не exhaustive Double.
- PUB: material publication v322 401/0, 40 API/role/permission вариантов;
  каноническая snapshot/publication матрица и F328/N328.
- NUM: `full_numbers_off_v322_v328_2026-10-05.json` и парный On;
  `NDM_Audit03_Final_Numerical_Comparison.md`, исходные допуски и 111 разобранных отличий.
- BOOK: clean/update Validate по 27/27, formatting по 1003/0, Help 2828 cells
  и одинаковые merges, palette/reopen 357/0, обычный расчет сохраненной копии.
- PERF: `performance_final_v328_2026-10-05.txt` и records/summary, 160 измерений,
  пять повторов, failed=0, корректный df10412f baseline.
- CAD: native shapes v286 1023/0, 22 случая; translation v286 66/0.
- REVIEW: Semantic/Code Comment/Guard и семейные Review MD, actual source manifest v328.

Имена raw-файлов относятся к `docs/regression/Audit03/`; Review MD к `docs/`.
Отсутствующая пиксельная QA не подменяется COM. Поздние v322/v325/v328
publication/адресные/технические guards не меняют принятую математику MATRIX;
их собственные negatives/positives и полные suites перечислены в Final Report.

### Audit01: Все 14 Пунктов

| ID | Проверенный Контракт | Актуальное Доказательство |
| --- | --- | --- |
| T01 | Secant финализирует проверенный root, не прежнюю нижнюю границу. | F328/N328, Search arithmetic directed, NUM. |
| T02 | Exhausted iteration budget не создает найденный предел. | F328/N328; соседние Double, budget=0 и Ultimate-stagnation directed. |
| T03 | Неподтвержденный initial solve остается numerical failure; BaseFail требует подтвержденного physical criterion. | F328/N328, On/Off offset-pair tests, MATRIX. |
| T04 | SEARCH_BOUND_REACHED сохраняет технический смысл, не превращается в найденный physical limit. | F328/N328; Crack bound positive 2157/0, русская причина без двойной пунктуации. |
| T05 | Доказанное отсутствие трещины не создает Post; найденный порог после LC дает NotCracked. | F328/N328, Formation gate/lifecycle tests, MATRIX. |
| T06 | Cache-hit восстанавливает solver snapshot без LastRunner и нового heavy solve. | F328/N328, currentCache/cacheHitWithoutLastRunner, PERF. |
| T07 | Failed solve не возвращается reusable; новый warm-start/retry допускается, context/revision проверяются. | F328/N328, retry-session 22/0, PUB, REVIEW. |
| T08 | Crack aggregate содержит primary current-state failure и правильное поддерево комментариев. | F328/N328, MATRIX, ResultComment Review, BOOK. |
| A01 | Общие LoadMultiplier/UltimateStrain loops действительно едины; domain adapters не содержат второго search. | SC328, F328/N328, REVIEW. |
| A02 | Probe равновесия проходит общий State pipeline; неуспех не становится physical failure. | F328/N328, typed-fault/recovery tests, PERF. |
| A03 | Formation самостоятельна, Width не исполняет Formation/Search. | SC328, independentFormation tests, REVIEW. |
| A04 | Search/state/meta/spec publication изолирована от дальнейшего live solver/context. | PUB, F328/N328, named-state/repository tests. |
| A05 | Writers выводят готовое свое поддерево; batch объединяет все разделы, без solver-text mapping. | MATRIX, BOOK, SC328, ResultComment Review. |
| A06 | Один display-словарь, typed codes и владелец комментариев; палитра отдельно. | SC328, F328/N328, palette BOOK, REVIEW. |

### Audit02: Все 56 Пунктов

| ID | Проверенный Контракт | Актуальное Доказательство |
| --- | --- | --- |
| R01 | Повтор CurrentCrackedState использует сохраненное НДС и дает ту же ширину без нового solve. | F328/N328, PERF crack-cache. |
| R02 | MaxLambda/последний интервал/representability проверяются без фиктивной конечной точки. | F328/N328, Search arithmetic и real Capacity precision. |
| R03 | Несошедшаяся probe не подтверждает physical failure; следующий корректный retry не блокируется. | F328/N328, typed faults, retry-session, MATRIX. |
| R04 | Generic Search работает через общий callback-контракт для обоих критериев. | SC328, fake linear problem и реальные Capacity/Formation tests. |
| R05 | Formation имеет собственный вход/контекст/result и не скрыта в Width. | SC328, F328/N328 independentFormation, REVIEW. |
| R06 | Evaluated plane не равна подтвержденному равновесию; lifecycle flags независимы. | F328/N328, F04 lifecycle matrices, typed State tests. |
| R07 | Числа принадлежат canonical results, aggregate/reset не оставляют второй изменяемый flat source. | SC328, PUB, F328/N328. |
| R08 | Пределы растяжения/сжатия стали выбираются по знаку для обоих search approaches. | F328/N328, asymmetricSteelLimits и limit-state tests, NUM. |
| R09 | Nothing/array/optional/Double guards имеют проверенный контракт и точную причину отказа. | REVIEW, directed negatives/positives F07, F328/N328. |
| A01 | Новых классов нет; выполнены два согласованных объединения. | SC328, manifest: 85 -> 83 production, прежние 3 test classes. |
| A02 | Batch остается orchestration/context/governing, не новым calculator/search монолитом. | REVIEW, SC328, F328/N328, source-line explanation в Final Report. |
| A03 | Formation/Width/Longitudinal разделены; pure numeric Width methods не получают State/status. | SC328, F328/N328 pure-formula/independentFormation, REVIEW. |
| A04 | Capacity adapter задает criterion/spec/path, общий State runner решает обычную probe. | F328/N328, SELECTOR, REVIEW. |
| A05 | Bisection/Brent/Secant/Ultimate общие для двух domains, с общими безопасными шагами. | SC328, F328/N328, directed stagnation/fault/recovery. |
| A06 | CLimitSearchResult не читает live Formation/solver после публикации. | PUB, F328/N328 diagnosticSnapshot. |
| A07 | Published API не меняет прежние numerical/meta/material snapshots при новом запуске. | PUB, SC328, F328/N328, material revision tests. |
| A08 | Writers пассивны; допустима арифметика представления, не инженерный verdict. | REVIEW, SC328, MATRIX, BOOK. |
| E01 | Канонический General.DiagramExtension един; alias имеет только узкую migration-роль. | F328/N328 readerMigration, CONFIG, SC328. |
| E02 | Обновление сохраняет пользовательские values/formulas/LC, priority и повторяемость migration. | saved Config migration/idempotence v323, BOOK, сохраненный обычный расчет v328. |
| E03 | Permission передается всем разрешенным State/Search roles, не меняя physical criterion. | F328/N328, MATRIX, SELECTOR, extreme On equilibrium 192 cases. |
| E04 | Ignore не оживляет tensile concrete; extension продолжает только активные ветви. | Material F328/N328, physicalDiagramPairs, CONFIG. |
| E05 | Физические узлы/плато/strain limits сохранены, контроля таблицы недостаточно без provider. | Material F328/N328, 23 input x 16 spec tests, NUM. |
| E06 | Техническая outer range и overflow определяются явно, не расширяют capacity произвольно. | Material/State F328/N328; +/-10 сверхтехнические cases и safe outer-limit tests. |
| E07 | Permission/ExtensionUsed/Converged/WithinPhysicalRange различаются; промежуточный extension не всегда означает конечный outside-state. | F328/N328 extendedInitialGuessPhysicalFinal, MATRIX, BOOK. |
| E08 | Auxiliary state не подается в неподходящую инженерную формулу; primary cause и blocked dependents согласованы. | F328/N328, F03/F04, MATRIX, ResultComment Review. |
| E09 | Обе стратегии используют один физический критерий и sign-specific limits. | Capacity/Crack F328/N328, SELECTOR, NUM. |
| E10 | Обычный Off сохраняет explicit-On исторические fixtures и восстанавливает effective mode. | F328/N328 GLOBAL_MODE/MODE_AFTER_SUITE, CONFIG isolation. |
| S01 | InternalStatus/ResultCode -> одна CResultStatusPolicy -> ExternalStatus, без назначения по comment text. | SC328, F328/N328 dictionary/aggregation, REVIEW. |
| S02 | Initial offset: BaseFail только при подтвержденном превышении, NumFail при численном неуспехе. | F328/N328 initialOffsetOnOff/offsetPair, MATRIX. |
| S03 | Fixed/Auto Formation различают постоянную часть, смену пути и консервативный fallback. | F328/N328, MATRIX все 12 путевых вариантов, typed terminal tests. |
| S04 | CRITERION_NOT_REACHED не смешан с SEARCH_BOUND_REACHED; NotCracked не вызывается из failed поиска. | F328/N328 FormationOutcome/FormationGate/bound tests. |
| S05 | Current state сохраняет фактическую typed причину; Width/Longitudinal получают blocked, не новый NumFail. | F328/N328, MATRIX, BOOK, F03 directed. |
| S06 | User/Auto/AlwaysCalc, signed sigma_s,crc и его усреднение имеют явные текущие правила. | Crack F328/N328, CONFIG, UserPsiFallback 440/0; отсутствующий Formation -> 1, confirmed NotCracked сохраняет User. |
| S07 | Подробные comments только своего subtree; summary все разделы, dedup и логический порядок. | MATRIX, ResultComment Review, F328/N328, BOOK. |
| C01 | State pipeline общий; material stress/tangent API не заменяется локальной доменной формулой. | SC328, F328/N328, REVIEW. |
| C02 | Reuse key учитывает context/revision/spec/permission/admissibility/tolerances, не warm-start policy. | F328/N328 repositoryContextAndRetry, REVIEW, PUB. |
| C03 | Named repository отдельно от локального search cache; failed attempts допускают новый solve. | F328/N328, retry-session 22/0, PERF, SC328. |
| C04 | Converged/EvaluatedPlane сами по себе не доказывают физическую пригодность State. | F328/N328 evaluatedPlane/current/offsetPair, MATRIX. |
| C05 | Saved Results имеет свои units/signs/role metadata; смена Config не делает скрытый import или solve. | F328/N328, imported snapshot unit-change 228/0, CAD, BOOK. |
| C06 | Initialize/Clear/clone/новый LC не переносят прежние flags/errors/result numbers. | F328/N328, read lifecycle, PUB, MATRIX. |
| Q01 | Корректный baseline df10412f воспроизведен до исправлений и сохранен. | baseline Off/On по 6745/0, baseline source equality 103/103, hashes в Progress. |
| Q02 | Прежние численные expected/tolerance не подогнаны; направления исправленного поведения раскрыты. | NUM; 536 неизменных из 542 AssertClose, два согласованных Psi и четыре ColumnWidth exceptions; все 277 Test-процедур сохранены. |
| Q03 | Explicit-On setup исторических tests сохранен, глобальный Off не маскируется. | F328/N328, mode-after-suite, CONFIG. |
| Q04 | On/Off внутри physical range сопоставлены численно; failures не сравниваются только итоговой строкой. | F328/N328, MATRIX, SELECTOR, NUM. |
| Q05 | Прежние directed gates повторены, новая Config/нагрузочная область не заменена старым счетчиком. | F328/N328, CONFIG, MATRIX, SELECTOR. |
| Q06 | Все 14 требований Audit01 имеют отдельную строку, владельца и evidence. | Итоговая таблица Audit01 выше, SC328, F328/N328. |
| Q07 | Финальная clean build содержит актуальный source, все восемь suites выполняются. | clean v328 build/SC328, F328/N328, BOOK. |
| Q08 | Performance, native CAD, снимки и save/reopen реально проверены в заявленных границах. | PERF, CAD, BOOK, MATRIX; пиксельная QA Excel не заявлена. |
| D01 | Действующие owners/comments описывают реальный код, не прежнюю/будущую архитектуру. | REVIEW, census 108/4567, два class merges, SC328; не автоматический PASS по счетчику. |
| D02 | Справка/Config/validation/links совпадают с текущими Extension/paths/units/physical contracts. | BOOK, CONFIG, actual Help comparison v328, REVIEW. |
| W01 | Baseline/spec/user hashes и чужие данные сохранены, исходный audit MD не изменен. | Progress/Final Report, saved Config guards, controlled publication. |
| W02 | Progress/ownership/evidence дают восстановление после compaction, решения сохранены. | Progress и семейные Review MD, final CONFIG/manifest. |
| W03 | Scoped checkpoints сделаны после конкретных gates, не вместо приемки. | git log до ca28919; negative/positive/raw reports не переписаны. |
| W04 | Нет push/destructive Git/отката чужих изменений; Excel cleanup ограничен собственным PID/startTicks. | Git diff/status/log, watchdog и исходные workbook hashes. |
| W05 | Узкий зеленый тест не объявляет весь goal завершенным; выполнены все шесть этапов DoD. | Final Report self-audit после завершения F328/N328 и publication. |
| W06 | Source/checked user book/VBE export/report согласованы и опубликованы. | Final Report identification, final publication receipt и scoped release commit. |
