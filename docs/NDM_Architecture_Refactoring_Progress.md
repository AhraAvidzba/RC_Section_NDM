# NDM Architecture Refactoring Progress

## Baseline

`6b9a6bea62e0436d206e698845332b6ad17fa231`

Baseline commit создан перед началом архитектурных этапов рефакторинга. Его нельзя
переписывать или разрушать; для сравнения использовать `git show`, `git diff` и
regression-отчеты.

## Completed

- Создан baseline commit перед началом архитектурных изменений.
- Рабочее дерево после baseline commit проверено через `git status`.
- Migration plan и укрупненный архитектурный план сохранены в `docs/`.
- Зафиксирован progress/status MD как внешняя память рефакторинга.
- Этап 0 завершен: создан baseline, проверен чистый старт, сохранены отчеты `Stage00`.
- Этап 1 завершен: введены `CResultMeta`, `CResultStatusPolicy`, единые enum-ы
  `InternalStatus`, `ResultCode`, `ResultKind` и внешний словарь статусов
  `OK`, `FAIL`, `BaseFail`, `NumFail`, `InputErr`, `CalcErr`, `N/A`.
- Текущий batch-слой начал агрегировать итоговый статус LC через `ResultMeta`,
  а не через разрозненные диагностические строки.
- Добавлены regression-тесты на словарь статусов, нейтральность `N/A` и запрет
  `NumFail` для формульных проверок без solver/search.
- Этап 2 завершен: `CSectionStateResult` хранит целевые и фактические усилия,
  невязки, признак физического диапазона, `ResultMeta` и диагностические данные
  solver-а; состояния трещинообразования переименованы в `PreCrackState` и
  `PostCrackState`.
- Этап 3 завершен: добавлен общий state-provider слой для решения именованных
  НДС без изменения расчетной постановки и существующих writer-контрактов.
- Этап 4 начат безопасным срезом для Capacity: введены `CLimitSearchRequest`,
  `CLimitSearchResult`, `CLimitSearchCoordinator`, `CUltimateStrainSearch` и
  `CLoadMultiplierSearch`; batch-слой больше не выбирает capacity-ветку сам,
  но низкоуровневая математика пока остается в проверенном `CCapacitySolver`.
- Этап 4 продолжен для CrackFormation: проверка образования нормальной трещины
  теперь идет через тот же `CLimitSearchCoordinator`, что и Capacity. Старый
  crack-диспетчер пути/стратегии убран, отдельная параллельная search-логика не
  вводилась; различия CrackFormation остались в критерии `eps_bt,ult`, пути
  нагрузки и material spec внутри `CCrackWidthCalculator`.
- Этап 4 продолжен на уровне `LoadMultiplier`: bisection-цикл Capacity и
  CrackFormation сведен в общий `CLoadMultiplierSearch.SearchByBisectionCore`.
  Capacity и CrackFormation теперь отличаются callback-ом оценки одной
  lambda-точки и финализацией найденной границы, а не отдельной копией
  bisection-алгоритма.
- Этап 4 завершен по низкоуровневому search-ядру: `LoadMultiplier` и
  `UltimateStrain` для Capacity и CrackFormation используют общие классы
  `CLoadMultiplierSearch` и `CUltimateStrainSearch`. Различия остались в
  критерии предела, load-path, material spec и callback-ах доменного слоя.
  Старые параллельные Newton/line-search/bisection реализации для
  CrackFormation и Capacity удалены из доменных классов.
- Этап 5 начат безопасным срезом для Capacity: добавлены `CCapacityCalculator`
  и `CCapacityResult`. `CBatchSectionCalculator` больше не строит
  `CLimitSearchRequest`, не вызывает `CLimitSearchCoordinator` напрямую и не
  раскладывает поля `CCapacitySolver` вручную. Batch получает готовый
  инженерный capacity-результат через `CCombinationResult.StoreCapacityResult`.
- Старые batch-методы ручного переноса capacity-данных (`StoreCapacity`,
  `StoreCapacityUltimateAtLoadPoint`) удалены. Перенос `Nult/Mxult/Myult` в
  пользовательскую точку приложения нагрузки и создание `CapacityState` теперь
  живут в `CCapacityResult`, без изменения расчетной математики.
- Execution report для Capacity больше не читает `CCapacitySolver.DiagnosticLog`
  напрямую из batch-слоя: диагностический текст проходит через
  `CLimitSearchResult` и `CCapacityResult`.
- Этап 5 продолжен следующим небольшим срезом: признак чисто осевой capacity-траектории
  после переноса нагрузок к расчетному центру больше не вычисляется в `CBatchSectionCalculator`.
  Его определяет `CCapacityCalculator`, а `CCapacityResult` хранит этот признак только для отчета.
- Этап 5 продолжен переносом ранних исходов Capacity: некорректный path,
  нулевая траектория и отсутствие ненулевой масштабируемой нагрузки теперь
  формируются в `CCapacityCalculator`/`CCapacityResult`. `RunCapacity` больше
  не записывает capacity-поля вручную для skipped/input-error веток, а только
  сохраняет готовый инженерный результат и пишет пояснение в отчет.
- Этап 5 продолжен переносом фабрики capacity load-path: `CBatchSectionCalculator`
  больше не содержит локальный `BuildLoadPathForCombination`. Траекторию из
  пользовательской строки и `CSectionLoadState` строит `CCapacityCalculator`,
  а batch-слой использует ее как готовую инженерную подготовку Capacity.
- Этап 5 продолжен переносом настройки capacity-solver-а: `CBatchSectionCalculator`
  больше не создает и не конфигурирует `CCapacitySolver`. Search-настройки,
  настройки внутреннего равновесия и strategy передаются в `CCapacityCalculator`,
  а материальные пределы берутся из построенных диаграмм конкретного
  material spec при создании solver-а.
- Этап 5 продолжен переносом capacity-report пояснений: план поиска,
  ранняя остановка, фактически выбранный метод и диагностический журнал
  последнего внутреннего solver-а выдаются `CCapacityCalculator`/`CCapacityResult`.
  Batch-слой только выводит готовые строки отчета.
- Этап 5 продолжен переносом typed capacity-meta: `CCapacityResult` теперь хранит
  `CResultMeta`, `CCombinationResult` использует его при `CapacityMeta`, а skipped
  capacity проходит через тот же `CCapacityResult` без ручного заполнения batch-полей.
- Этап 5 продолжен переносом интерпретации capacity-результата в LC-result:
  `CCombinationResult` теперь сам отдает display-name траектории и запас
  несущей способности через typed meta. Batch и detailed strength writer больше
  не вычисляют `запас(λ)` напрямую из `LambdaCapacity` для NumFail/InputErr/N/A.
- Этап 5 продолжен защитой downstream-вывода Capacity: `CNDMResultsWriter` выводит
  предельные компоненты `Nultimate/MxUltimate/MyUltimate` только при наличии
  физически пригодного capacity reserve, а regression-тесты фиксируют нулевой запас
  для NumFail/InputErr capacity-веток. Локальный вызов `MaxDouble` в настройке
  `CCapacityCalculator` заменен явным сравнением пределов арматуры, чтобы исключить
  compile-зависимость от helper-а в этом классе.
- Этап 5 продолжен отвязкой downstream-слоя Capacity от `CCapacitySolver`:
  `CLimitSearchResult` сохраняет численный snapshot найденной limit-точки
  (`lambda`, компоненты усилий, utilization, последний solver и диагностику),
  `CCapacityResult` читает эти данные из общего result, а внешний статус берет
  из `CResultMeta` через `CBatchStatusPolicy`. Старый batch-маппинг
  `CCapacitySolver -> display status` удален.
- Этап 5 продолжен сужением зависимости от capacity-фасада:
  `CLimitSearchResult` больше не хранит и не отдает `CCapacitySolver` как часть
  результата. Ссылки на `request.CapacitySolver` остаются только внутри
  search-слоя и `CLimitSearchRequest`, а execution report читает
  capacity-диагностику из сохраненного снимка result.
- Этап 5 завершен по основному batch-пути: `Capacity.SolutionStrategy = Auto`
  идет через `CLimitSearchCoordinator`, а не через shortcut внутри
  `CCapacitySolver`.
  Coordinator сначала запускает `CUltimateStrainSearch`, затем при необходимости
  повторяет ту же траекторию через `CLoadMultiplierSearch`; диагностика первой
  попытки объединяется в доменном `CCapacitySolver` без возврата search-логики
  в batch.
- `CCapacityLoadPath` оставлен как capacity-адаптер пользовательской настройки
  `CapacityLoadPath`. Для CrackFormation напрямую используется общий
  `CLoadPathVector` и `CLoadPathMath`; тащить `CCapacityLoadPath` в трещины
  не нужно, потому что он содержит capacity-специфичные правила Auto/ForceOnly.
- Этап 6 завершен основным срезом Crack: добавлены отдельные result-объекты
  `CCrackFormationResult`, `CCrackWidthResult`, `CLongitudinalCrackResult` и
  `CCrackResult`. Batch-слой хранит typed result-tree трещин, а внешний статус
  по-прежнему формируется через `CResultStatusPolicy`/`CBatchStatusPolicy`.
- Чистая формула раскрытия нормальной трещины вынесена в
  `CCrackWidthFormulaCalculator`. Она получает только готовые численные
  `phi/psi/sigma_s/E_s/l_s` и не получает State-объекты, не выбирает зону,
  не выбирает арматуру и не запускает solver.
- Для `CurrentCrackedState` больше не назначается `rsNumericalFailure`
  автоматически при любом отсутствии результата: сохраняется фактическая
  причина `InvalidInput`, `InvalidConfiguration`, `NumericalFailure` или
  `InternalError`; downstream width/longitudinal блокируются через dependency.
- Этап 7 завершен по result-tree: `CCombinationResult` хранит
  `StrengthResult`, `CrackResult` и `StabilityResult`, а display-поля
  синхронизируются из typed result-объектов для существующего формата Results.
  Direct state, capacity, crack aggregate и stability теперь проходят через
  вложенные result-объекты без изменения расчетной математики и структуры
  Results.
- Этап 8 завершен для текущего миграционного среза: batch summary и подробные
  writer-ы прочности, трещин, устойчивости и named-state properties выводят
  `ResultComment` из соответствующих `ResultMeta`/result-subtree. Статусы и
  окраска status-ячеек остаются централизованы через status policy, а layout
  Results сохранен с согласованным добавлением `ResultComment`.
- Этап 9 завершен: execution report берет пользовательские комментарии из
  result graph/`ResultMeta`, а не собирает инженерные причины в writer-е или
  batch-report слое. Подробные diagnostic-логи solver/search сохранены для
  анализа итераций и fallback-веток.
- Убрана зависимость назначения статусов от произвольного текста
  `StopReason`/`ResultComment`: внешний статус формируется через
  `CResultStatusPolicy` по `InternalStatus`, `ResultCode` и `ResultKind`.
- Этап 10 завершен по очистке status/state paths: удалены старые crack-state
  alias-ы, локальные string-status -> meta helper-ы и текстовый маппинг
  инженерных причин в пользовательские статусы. Допустимые внешние display
  statuses остаются только `OK`, `FAIL`, `BaseFail`, `NumFail`, `InputErr`,
  `CalcErr`, `N/A`.

## In progress

- Нет активного этапа после завершения этапа 10.

## Next

- Использовать baseline commit и отчеты `Stage00`/`Stage01`/`Stage02` как точки сравнения
  для следующих этапов.
- Перед следующим этапом снова проверить `git status` и убедиться, что
  нет посторонних пользовательских изменений.
- Для следующих этапов сохранять финальную цепочку
  `InternalStatus/ResultCode -> CResultStatusPolicy -> ExternalStatus`.

## Known risks / open questions

- Для AutoCAD Region geometry часть проверок может оставаться ручной, если
  тестовый AutoCAD недоступен в автоматическом прогоне.
- В дальнейшем не смешивать миграционные изменения разных этапов в один diff
  без отдельной проверки.
- Строковые display-поля остаются только форматом вывода Results; расчетная
  семантика статуса должна идти через `ResultMeta`.

## Important decisions

- `CStateRepository` должен хранить reusable только корректно полученные
  конечные named states.
- Search probe-cache остается локальным для search-session и не попадает в
  shared repository.
- Неуспешные solve attempts сохраняются только как диагностика/result, но не
  как reusable physical State.
- `CCrackWidthFormulaCalculator` остается чистой формульной частью и не получает
  State/status.
- `BaseFail` используется для Capacity, когда базовая точка траектории уже за
  предельным состоянием. Это не `NumFail`.
- `NumFail` должен соответствовать реальной численной несходимости solver/search,
  а не обычной инженерной проверке типа ширины трещины, продольных трещин или
  устойчивости.
- Канонические имена состояний трещинообразования в snapshot и настройках:
  `PreCrackState` и `PostCrackState`. Другие имена не принимаются.

## Baseline Test Scenarios

- Стандартное RectSet `LSection`.
- `Circle`.
- `RoundedRectangle`.
- `HollowRectangle`.
- Импортированная AutoCAD Region geometry, если доступна среда AutoCAD.
- Direct state по модели прочности.
- Capacity по выбранной lambda-траектории.
- Crack formation, crack width и longitudinal cracks.
- Stability по СП 35 и СП 63.
- Вывод `rngBatchSummary`.
- Detailed strength/crack/stability blocks.
- `rngNDMSectionProperties`.
- Execution report.

## Last verified

- 2026-09-30: этап 10 пересобрал `workbook/output/RC_Section_NDM.xlsm` через
  `tools/build_workbook/Build-Workbook.ps1` и прошел полный
  `tools/build_workbook/Run-AllTests.ps1` без failures. В отчете UI-тестов:
  `passed=308`, `failed=0`; batch: `passed=644`, `failed=0`;
  regression baseline: `passed=39`, `failed=0`.
- 2026-09-28: после этапа 2 `tools/build_workbook/Build-Workbook.ps1`
  завершился успешно.
- 2026-09-28: после этапа 2 `tools/build_workbook/Run-BatchTests.ps1`
  завершился успешно: `passed=625`, `failed=0`.
- 2026-09-28: `tools/build_workbook/Run-AllTests.ps1` завершился успешно;
  итоговый отчет этапа 2 сохранен в `docs/regression/Stage02_AllTests_Report.txt`,
  regression baseline raw report сохранен в
  `docs/regression/Stage02_RegressionBaseline_Raw.txt`.
- 2026-09-28: после первого среза этапа 4 `tools/build_workbook/Build-Workbook.ps1`
  завершился успешно.
- 2026-09-28: после первого среза этапа 4 `tools/build_workbook/Run-CapacityTests.ps1`
  завершился успешно: `passed=923`, `failed=0`.
- 2026-09-28: после первого среза этапа 4 `tools/build_workbook/Run-BatchTests.ps1`
  завершился успешно: `passed=631`, `failed=0`.
- 2026-09-28: после первого среза этапа 4 `tools/build_workbook/Run-AllTests.ps1`
  завершился успешно.
- 2026-09-29: после подключения CrackFormation к общему `CLimitSearchCoordinator`
  `tools/build_workbook/Build-Workbook.ps1` завершился успешно; отдельные
  `Run-CrackTests.ps1`, `Run-CapacityTests.ps1`, `Run-BatchTests.ps1` и полный
  `Run-AllTests.ps1` завершились успешно.
- 2026-09-29: после полного объединения низкоуровневой search-логики
  Capacity/CrackFormation основной `workbook/output/RC_Section_NDM.xlsm`
  пересобран успешно.
- 2026-09-29: финальная книга после этапа 4 прошла
  `tools/build_workbook/Run-CapacityTests.ps1`: `passed=923`, `failed=0`.
- 2026-09-29: финальная книга после этапа 4 прошла
  `tools/build_workbook/Run-CrackTests.ps1`: `passed=286`, `failed=0`.
- 2026-09-29: первый срез этапа 5 пересобрал
  `workbook/output/RC_Section_NDM.xlsm` через `tools/build_workbook/Build-Workbook.ps1`.
- 2026-09-29: первый срез этапа 5 прошел
  `tools/build_workbook/Run-CapacityTests.ps1`: `passed=923`, `failed=0`.
- 2026-09-29: первый срез этапа 5 прошел
  `tools/build_workbook/Run-BatchTests.ps1`: `passed=631`, `failed=0`.
- 2026-09-29: первый срез этапа 5 прошел
  `tools/build_workbook/Run-CrackTests.ps1`: `passed=286`, `failed=0`.
- 2026-09-29: следующий срез этапа 5 с переносом skipped/input-error
  capacity-path в `CCapacityCalculator` пересобрал
  `workbook/output/RC_Section_NDM.xlsm` через `tools/build_workbook/Build-Workbook.ps1`.
- 2026-09-29: этот же срез этапа 5 прошел
  `tools/build_workbook/Run-CapacityTests.ps1`: `passed=923`, `failed=0`.
- 2026-09-29: этот же срез этапа 5 прошел
  `tools/build_workbook/Run-BatchTests.ps1`: `passed=631`, `failed=0`.
- 2026-09-29: этап 6 пересобрал `workbook/output/RC_Section_NDM.xlsm`
  через `tools/build_workbook/Build-Workbook.ps1`.
- 2026-09-29: этап 6 прошел `tools/build_workbook/Run-CrackTests.ps1`:
  `passed=291`, `failed=0`.
- 2026-09-29: этап 6 прошел `tools/build_workbook/Run-BatchTests.ps1`:
  `passed=633`, `failed=0`.
- 2026-09-29: этап 6 прошел полный `tools/build_workbook/Run-AllTests.ps1`
  без failures.
- 2026-09-30: этап 7 пересобрал `workbook/output/RC_Section_NDM.xlsm`
  через `tools/build_workbook/Build-Workbook.ps1`.
- 2026-09-30: этап 7 прошел `tools/build_workbook/Run-BatchTests.ps1`:
  `passed=641`, `failed=0`.
- 2026-09-30: первый крупный срез этапа 8 пересобрал книгу и прошел полный
  `tools/build_workbook/Run-AllTests.ps1` без failures. Отдельно проверены
  batch-тесты `passed=641`, `failed=0` и UI-тесты `passed=308`, `failed=0`.
- 2026-09-29: срез этапа 5 с защитой capacity reserve в `CNDMResultsWriter`
  пересобрал `workbook/output/RC_Section_NDM.xlsm` через
  `tools/build_workbook/Build-Workbook.ps1`.
- 2026-09-29: этот же срез этапа 5 прошел
  `tools/build_workbook/Run-CapacityTests.ps1`: `passed=923`, `failed=0`.
- 2026-09-29: этот же срез этапа 5 прошел
  `tools/build_workbook/Run-BatchTests.ps1`: `passed=633`, `failed=0`.
- 2026-09-29: срез этапа 5 с переносом typed capacity-meta пересобрал
  `workbook/output/RC_Section_NDM.xlsm` через `tools/build_workbook/Build-Workbook.ps1`.
- 2026-09-29: этот же срез этапа 5 прошел
  `tools/build_workbook/Run-CapacityTests.ps1`: `passed=923`, `failed=0`.
- 2026-09-29: этот же срез этапа 5 прошел
  `tools/build_workbook/Run-BatchTests.ps1`: `passed=631`, `failed=0`.
- 2026-09-29: срез этапа 5 с переносом display path и capacity reserve в
  `CCombinationResult` пересобрал `workbook/output/RC_Section_NDM.xlsm` через
  `tools/build_workbook/Build-Workbook.ps1`.
- 2026-09-29: этот же срез этапа 5 прошел
  `tools/build_workbook/Run-CapacityTests.ps1`: `passed=923`, `failed=0`.
- 2026-09-29: этот же срез этапа 5 прошел
  `tools/build_workbook/Run-BatchTests.ps1`: `passed=631`, `failed=0`.
- 2026-09-29: срез этапа 5 с переносом capacity-report пояснений пересобрал
  `workbook/output/RC_Section_NDM.xlsm` через `tools/build_workbook/Build-Workbook.ps1`.
- 2026-09-29: этот же срез этапа 5 прошел
  `tools/build_workbook/Run-CapacityTests.ps1`: `passed=923`, `failed=0`.
- 2026-09-29: этот же срез этапа 5 прошел
  `tools/build_workbook/Run-BatchTests.ps1`: `passed=631`, `failed=0`.
- 2026-09-29: срез этапа 5 с переносом настройки capacity-solver-а пересобрал
  `workbook/output/RC_Section_NDM.xlsm` через `tools/build_workbook/Build-Workbook.ps1`.
- 2026-09-29: этот же срез этапа 5 прошел
  `tools/build_workbook/Run-CapacityTests.ps1`: `passed=923`, `failed=0`.
- 2026-09-29: этот же срез этапа 5 прошел
  `tools/build_workbook/Run-BatchTests.ps1`: `passed=631`, `failed=0`.
- 2026-09-29: этот же срез этапа 5 прошел
  `tools/build_workbook/Run-CrackTests.ps1`: `passed=286`, `failed=0`.
- 2026-09-29: срез этапа 5 с переносом фабрики capacity load-path пересобрал
  `workbook/output/RC_Section_NDM.xlsm` через `tools/build_workbook/Build-Workbook.ps1`.
- 2026-09-29: этот же срез этапа 5 прошел
  `tools/build_workbook/Run-CapacityTests.ps1`: `passed=923`, `failed=0`.
- 2026-09-29: этот же срез этапа 5 прошел
  `tools/build_workbook/Run-BatchTests.ps1`: `passed=631`, `failed=0`.

## Финальное закрытие post-refactoring audit (2026-10-01)

### Completed
- Закрыт audit по `docs/NDM_PostRefactoring_Audit_2026-09-30.md`: пункты T01-T08 и A01-A06 проверены по финальному коду.
- Исправлено regression-падение capacity-тестов после перехода на общий `CStateSolutionRunner`: runner снова принимает общий материал solver-а, а specialized active-set/extension подсказки применяет только для `CMaterialDiagram`.
- Добавлен итоговый self-audit: `docs/NDM_PostRefactoring_Audit_Closure_2026-10-01.md`.

### In progress
- Нет незавершенной работы в рамках текущего audit.

### Next
- Новые архитектурные задачи начинать только отдельным этапом после отдельного пользовательского задания.

### Known risks
- Не выявлено известных противоречий с v6/migration plan по цепочкам `State solve -> Search -> Engineering result`, `InternalStatus/ResultCode -> CResultStatusPolicy -> ExternalStatus`, `StateRepository -> reusable final states`.
- `git status` может показывать `src/Solver/CCapacitySolver.cls` как modified из-за невозможности обновить index stat внутри sandbox; `git hash-object` совпадает с `HEAD`, текстового diff нет.

### Important decisions
- Неуспешные solve не становятся reusable physical state.
- Writer-ы не формируют инженерские статусы и комментарии заново, а читают `ResultMeta`.
- `BaseFail` остается только внешним отображением для Capacity при `rsCheckFailed + INITIAL_STATE_BEYOND_LIMIT`.
- Для тестовых линейных материалов runner не использует `CMaterialDiagram`-specific проверки физического диапазона.

### Last verified
- `tools/build_workbook/Build-Workbook.ps1`: OK.
- `tools/build_workbook/Run-AllTests.ps1`: `TOTAL passed=496; failed=0`; `TOTAL_CAPACITY passed=934; failed=0`; `TOTAL_CRACK passed=328; failed=0`; `TOTAL_BATCH passed=659; failed=0`; `TOTAL_WORKBOOK_UI passed=308; failed=0`; `TOTAL_REGRESSION_BASELINE passed=39; failed=0`.
