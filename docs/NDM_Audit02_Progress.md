# NDM Audit02 Progress

Дата старта: 2026-10-01.

## Baseline

- Baseline commit: `855626e69b07f21dc636771416c501097d077ec9`
- Baseline label: `855626e6 Checkpoint post-refactoring audit fixes`
- Audit02 spec: `docs/NDM_Audit02_Implementation_Spec_2026-10-01.md`
- Audit02 spec SHA-256: `935CAB51FF62D6E0438A1C4D1082F21306CE181E515CAC7A91C8CA2F06E2B991`

## Initial Git State

Перед началом Audit02 рабочее дерево не было чистым. Изменения сохранены как входное состояние задачи, без отката:

- `workbook/output/RC_Section_NDM_execution_report.txt`
- `workbook/output/VBA_All_Code.txt`
- `docs/NDM_Audit02_Implementation_Spec_2026-10-01.md` (новый файл)

## Requirement Status

| ID | Статус | Комментарий |
| --- | --- | --- |
| R01-R09 | реализовано; финальная приемка в работе | Cache-hit, точная MaxLambda, подтверждение равновесия, общий Search, независимая Formation, snapshots, canonical results и sign-specific limits прошли принятые срезы. Последний R09 guard и новые граничные тесты проверяются чистой сборкой. |
| A01-A08 | реализовано; финальная приемка в работе | Нет новых production-классов; DomainContext и flat-дубли удалены. Запасы принадлежат typed results. Остаточное форматирование передано CExecutionReport; мертвый API пределов Batch удален. |
| E01-E10 | проверено в принятых срезах | Один General.DiagramExtension, active branches, неизменные узлы/пределы, effective context и Off. Добавленная малая окрестность eps_ult прошла material suite 1359/0. |
| S01-S07 | проверено в принятых срезах | Typed причины, warning/fallback, SearchBound, blocked formulas, psi. Парный overload-тест подтверждает Off NumFail / On BaseFail и одинаковый вывод в summary/detail. |
| C01-C06 | проверено; итоговая регрессия в работе | Scoped reuse, failed retry и mutation isolation приняты. Config change/read/export/plot: UI 574/0. Save/reopen: exact match, 0 solve. Добавлен явный PostState cache-hit. |
| Q01-Q08 | в работе | Baseline/Off сохранены, current On/Off 6222/0 и 1448 парных assertions приняты. Performance: 3 повтора, источники неизменны. Повторяются все suites на чистой итоговой сборке. |
| D01-D02 | реализовано; self-audit в работе | Config/help/runtime и Architecture согласованы с global Extension; исторические v6/migration/progress помечены. |
| W01-W06 | в работе | Baseline и checkpoints сохранены; метрики сняты. До завершения необходимы final accepted suites, итоговый отчет, self-audit и финальный checkpoint. |

## Responsibility Transfer Map

| Зона | Текущий целевой владелец | Что проверить/перенести по Audit02 |
| --- | --- | --- |
| Named State solve | `CStateSolutionRunner`, `CStateProvider`, `CStateRepository`, `CSectionStateResult` | Cache-hit без `LastRunner`, effective request key, физические признаки, snapshots без повторного solve. |
| Limit Search | `CLimitSearchCoordinator`, `CUltimateStrainSearch`, `CLoadMultiplierSearch`, `CLimitSearchResult` | Самодостаточный result, границы search, initial-state semantics, отсутствие скрытого знания Capacity/Crack в generic search. |
| Capacity | `CCapacityCalculator`, `CCapacityResult`, остаточный `CCapacitySolver` | BaseFail только из подтвержденного `INITIAL_STATE_BEYOND_LIMIT`, отсутствие физического отказа из несошедшейся плоскости, асимметричные пределы арматуры. |
| Crack Formation/Width | `CCrackFormationCalculator`, `CCrackWidthCalculator`, `CCrackWidthFormulaCalculator`, `CCrackResult` | Formation отдельно от Width, `CurrentCrackedState` только при необходимости, `sigma_s,crc <= 0` без ложного NumFail. |
| Combination result | `CCombinationResult` и typed result-tree | Удаление временных flat/proxy полей после миграции consumers, единый источник статусов и комментариев. |
| Writers | `CBatchResultWriter`, `CStrengthSummaryWriter`, `CCrackSummaryWriter`, `CStabilitySummaryWriter`, `CNDMResultsWriter` | Writer-ы остаются пассивными: formatting/units/presentation only, без инженерных статусов и проверок. |
| Diagram Extension | Material/model provider + state/search consumers + Config/help | Один `General.DiagramExtension`, active branches only, физические пределы не расширяются. |

## Stages

1. Контрольная точка: baseline build/tests, inventory R01-R09, search for legacy fields/consumers.
2. Канонические данные и дефекты: R01-R03/R06/R08/R09, directed tests.
3. Архитектурная очистка: R04/R05/R07/A01-A08, consumers -> typed results.
4. Global Extension: E01-E10, Config/help/runtime, On/Off comparisons.
5. Интерпретация и UI/output: S01-S07, D01-D02, comments/status/output checks.
6. Финальная приемка: full build/tests, regression logs, self-audit, final report.

## Completed

- Прочитан Audit02 implementation spec.
- Прочитаны `AGENTS.md`, architecture progress и closure report первого audit.
- Зафиксированы baseline SHA и исходное dirty-state.
- Baseline commit распакован отдельно без переключения текущих исходников. Сборка baseline и все 8 тестовых наборов прошли без ошибок.
- Реализован первый блок исправлений R02/R03/R08: проба точно на MaxLambda, запрет физической классификации несошедшейся плоскости, раздельные пределы сжатия и растяжения арматуры. Приемка этого блока пока не пройдена.
- В `CSectionSolver` оценка заданной плоскости больше не устанавливает Converged. Новый `ConfirmEquilibrium` проверяет фактические N/Mx/My против целевой нагрузки без нового solve; все 5 направленных проверок прошли.
- Восстановление сохраненного State перенесено в общий State-слой (`CStateProvider.SolverSnapshot` / `CStateSolutionRunner.RestoreSnapshot`). Дублирующий код Batch/Width удален. Интеграционный повтор расчета трещин прошел: ширина, sigma_s, sigma_s,crc и psi совпали, число новых solve равно нулю.
- Аналитическая жесткость новых capacity-тестов уточнена с учетом вытеснения бетона арматурой, как в существующем `CSectionSolver`. Допуски ожидаемых результатов не ослаблены.
- При подтверждении чисто осевого предела lambda восстанавливается из фактических усилий предельной плоскости и проверяется по выбранному load path. Исходные ошибки `capacity.nult.tension.*` больше не воспроизводятся в последнем suite.

## Current

- Последний принятый HEAD: `020ff8f9` (typed reserves + paired physical acceptance); baseline неизменен.
- `typed_reserves_pair_full_on/off_2026-10-01.txt`: по 6222 assertions, 8 suites, 0 ошибок. `on_off_physical_accepted_2026-10-01.txt`: 1448/0; MAXDIFF по каждой величине.
- Последний срез: ранняя подтвержденная pure-axial finalization, дополнительные Q05-06/08/09/17/19/24, фактический AutoCAD export-reader, перенос оставшегося форматирования к CExecutionReport, удаление шести неиспользуемых методов Batch, R09 guard в SectionTypeRegistry.
- `snapshot_export_ui_accepted_on_2026-10-01.txt`: 574/0. `snapshot_save_reopen_accepted_2026-10-01.txt`: смена семи материальных настроек, закрытие/открытие, exact match 38950 символов, read/export/plot 0 solve, source unchanged.
- `performance_confirmed_2026-10-01.txt`: baseline/current, 8 сценариев по 3 повтора, все assertions прошли, source unchanged. Pure axial: 30 -> 38 solve (против 501 до ранней finalization), probes 32 -> 12. Асимметрия: 20 -> 94 solve; baseline трижды принимал несошедшиеся пробы за физический предел, current делает подтвержденные retry (unconfirmedPhysical 3 -> 0). Это измеренное исправление R03, не потеря кэша; cache/reuse в methods и crack_cache сохранены.
- `final_clean_build_2026-10-01.log`: полная отдельная сборка. Первый полный Off НЕ принят: новый section-test не задавал rbMc2/Ebt; новый exact-offset тест игнорировал фактическую сторону границы из-за округления. Исправлен только setup/смысл новых тестов, production-критерий и исторические expected/tolerance не менялись. Негативный лог сохранен.
- Второй Off также выявил ошибку setup нового теста: `EvaluateStrainPlane` не означает подтвержденное равновесие. Тест явно вызвал `ConfirmEquilibrium`; production не менялся. Негативный лог `final_clean_accepted_off_2026-10-01.txt` сохранен.
- Итоговая независимая полная сборка: `final_build_accepted_2026-10-01.log`, exit 0. Полные `final_verified_off/on_2026-10-01.txt`: каждый 6745 assertions / 0 ошибок, 8 suites, source unchanged. Физические On/Off пары содержат полный MAXDIFF; максимальная разница Capacity Mx = 0.000650160015 Н*мм, намного меньше прежнего допуска 5000 Н*мм.
- Пользовательская output-книга обновлена только через `Refresh-VbaModules`: `final_output_refresh_2026-10-01.log`, exit 0. Штатный `Run-AllTests`: `final_output_all_tests_2026-10-01.txt`, exit 0, те же 6745/0; структура, включая Print area, прошла. Отдельная clean validation также exit 0.
- `saved_combinations_on_off_accepted_2026-10-01.txt`: в доступной output-книге одно сохраненное LC=1; Capacity OK в обоих режимах, summary/detail совпадают, исходник неизменен. Это не набор перегруженных LC со старых скриншотов. Для двух overload-сценариев OFFSET_2MN/OFFSET_3MN подтверждено Off=NumFail / On=BaseFail, typed codes и вывод совпадают.
- `expected_preservation_accepted_2026-10-01.txt`: сохранены 247 исходных test-процедур; 436 AssertClose expected/tolerance идентичны, 16 меняют только явно проверенный typed accessor, все tolerance сохранены. Финальные self-audit/report/checkpoint завершаются; до их сохранения цель не объявлять выполненной.
- `final_source_contracts_accepted_2026-10-01.txt`: 103 импортированных src/tests modules соответствуют исходникам с учетом штатного VBE представления ANSI/Double; нет новых/удаленных production-классов (85 -> 85), DomainContext/downcasts и State/status/solve в чистой формуле отсутствуют. SHA-256 артефактов записаны в лог.

### История Принятых Срезов И Предыдущих Точек Продолжения

- HEAD checkpoint R07: `8929d5d`, `Audit02: canonical typed results and block output formatting`.
- C02/C03 реализованы через identity/revision сечения и material provider, effective Extension и допуски в scoped repository; warm-start/retry/диагностика не входят в reuse key. Provider владеет небольшими копиями параметров материалов. Последний named-state обновляется и при solve, и при cache-hit; failed state не блокирует повтор. Directed `audit02.cache.*`: 43/0. Full On/Off `scoped_repository_full_*_2026-10-01.txt`: все 8 suites, по 4774 assertions, 0 ошибок, source unchanged; структурная проверка прошла.
- Следующий срез A02 реализуется отдельно: запасы отдельных проверок переходят к canonical typed results, Batch сохраняет только выбор определяющего LC. Эти исходники еще не включены в принятую C02/C03 книгу и требуют новой проверки.
- R07 acceptance: `canonical_accepted_full_on_2026-10-01.txt` и `canonical_accepted_full_off_2026-10-01.txt`: все 8 suites, по 4731 assertions, 0 ошибок, исходная книга неизменна. Структура книги проверена (`canonical_accepted_validation_2026-10-01.log`). Этот срез фиксируется отдельно от еще непроверенных C02/C03 изменений scoped repository/material snapshots.
- HEAD: `92d0a4ae Audit02: canonical search snapshots and unified load multiplier`; перед этим подтвержденные checkpoints `ff9e2d47`, `799bcd8e`. Baseline остается `855626e6`.
- В работе R07/A02/A07: удалены 140 flat-полей CCombinationResult и 151 indexed-forward getter Batch. Writers, workbook и тесты переведены на `ResultAt(index)` и typed subtrees; named-state читаются напрямую через StateRepository. Численные поля DirectState теперь принадлежат одному CSectionStateResult; CrackedState больше не перезаписывает strength direct-result. Форматирование отчетов перенесено в существующий CExecutionReport. Новых production-классов нет.
- Повторная инициализация Width/Stability полностью сбрасывает прежние числа/flags. `audit02.canonical.*` identity/reset/meta-isolation assertions прошли в batch 699/0. Cached display удален у Capacity/State, он вычисляется через policy. Последний indexed CapacityLoadPath proxy удален; UI-тест load-reference читает реально запрошенный CrackedState, не strength-ветвь другого профиля.
- Последний R07 full On `typed_canonical_full_on_2026-10-01.txt` НЕ принят: geometry 496/0, materials 1191/0, section 444/0, capacity 1194/0, crack 360/0, batch 672/2; UI остановился с Excel COM ошибкой ColumnWidth, regression не запущен. Две batch-регрессии исправлены в исходниках: invalid capacity path не выдает None как фактически выбранный путь; Stability.Branch сохраняет машинный ключ, пользовательская причина остается в meta. Новые правки требуют повторной проверки.
- Прерванные UI-прогоны `typed_ui_isolated_on*`, `typed_ui_traced_on*` показали рост памяти Excel свыше 3 ГБ даже при запуске отдельной UI-suite; тестовый Excel закрыт по прямому разрешению пользователя. Эти прогоны не считаются green/приемкой. В VBA UI-тестах добавлен только диагностический файл этапов `RC_NDM_ui_test_progress.txt` в папке независимой тестовой копии; численные ожидания/допуски не менялись.
- Причина роста памяти локализована по фазам: batch-only 0.23 s / 133 MB, тот же расчет с writer-ами около 18 s / 918 MB. Задание одинаковой ширины столбцов отдельными COM-вызовами заменено одной блочной операцией для каждого output-блока; default width листа задан через StandardWidth. Числа, шапки и merge-схема сохранены. Повтор no-plot: около 4 s / 195 MB (`block_width_ui_memory_2026-10-01.txt`). Полный UI больше не падает по памяти; Q08 измерения representative-сценариев остаются открытыми.
- В этом продолжении ошибочно использован исторический `Refresh-Workbook.ps1`: он меняет старую семиколоночную Config-схему. Только затронутые 15 строк B:G восстановлены из независимой копии, сохраненной до запуска; формулы оставлены локальными, прежняя область печати восстановлена. Лог `restore_refresh_config_cells_2026-10-01.log`. Использовать исключительно `Refresh-VbaModules.ps1`, он не меняет данные листов. Этот incident не относится к расчетной методике; пользовательские входы восстановлены, defaults не навязывались.
- `canonical_final_full_on_2026-10-01.txt`: все 8 suites, 4731 assertions, 0 ошибок (496/1191/444/1194/360/699/308/39), source unchanged. После этого удалены последние строковые fallback-ы writer-ов; batch normal-crack summary берет агрегированную CrackMeta, а не только WidthMeta. Итоговый Off `canonical_accepted_full_off_2026-10-01.txt` выполняется; затем нужны On и structural validation текущего среза.
- Тестовый PowerShell runner теперь сохраняет начало/окончание каждой suite до/после COM-вызова и не теряет отчет при cleanup-ошибке. Отключенный CExecutionReport не форматирует сырой diagnostic-блок; это вспомогательная оптимизация, не установленная причина UI-роста памяти.
- Актуальный блок R04/A05/A06: канонический Search snapshot и общий LoadMultiplier реализованы. Аналитическая матрица проверяет оба вида задачи, три метода, найденный предел/точную MaxLambda/неисследованный предел, recovery и повторную инициализацию. Отдельные численные циклы Capacity/Formation удалены; физические callbacks остаются у владельцев задач.
- Монотонный recovery после успешной промежуточной пробы изменил исторический порядок warm-start и вызвал 11 Off-регрессий осевого RectSet. Причина устранена в общем FindBracket: успешная проба снова расширяет нагрузку с улучшенным стартом, несошедшаяся проба уменьшает шаг и не объявляется физическим пределом. `shared_expansion_full_off_2026-10-01.txt` и `shared_expansion_full_on_2026-10-01.txt`: каждый прогон содержит все 8 suites, 4706 assertions, 0 ошибок, исходная книга неизменна. Structural validation прошла (`shared_search_validation_2026-10-01.log`). Ожидания/допуски не менялись. Capacity On 46.785 s / Off 45.180 s; Q08 performance остается открытым.
- Следующий блок после проверенного R07 checkpoint: C02/C03 scoped invalidation repository при смене effective Extension, геометрии, материалов и допусков; корректная замена конечных named states и failed retry. Новых production-классов не планируется. Далее Q04 численные On/Off сравнения, матрица Q05, representative performance, актуализация D01/D02 и полный self-audit.
- Работа возобновлена. Новое указание пользователя: больше не приостанавливать цель для перезагрузки без нового явного запроса.
- Проверенный checkpoint: `799bcd8`, `Audit02 checkpoint: state fixes and global extension migration`; push не выполнялся.
- Следующий проверенный checkpoint: `ff9e2d4`, `Audit02: independent crack formation and confirmed physical boundaries`. R05/Formation перенос прошел полные On/Off suites, направленные physical-block assertions и структурную проверку книги. Push не выполнялся.
- Baseline Off на независимой копии завершился с exit 0: `baseline_off_all_tests_2026-10-01.txt`, все 8 suites без ошибок. Каждый suite начинается со свежего открытия Off-копии; исторический UI explicit-On сохраняет свой setup, изменения отбрасываются при закрытии.
- Current Off выявил 7 Capacity-регрессий в batch для осевого RectSet. Ожидания и tolerance сохранены. Простое масштабирование retained-плоскости вывело все стержни на плато и не подтвердило путь. Старт общего Newton с нулевой деформацией на противоположном краю сохранил активную жесткость и подтвердил физический предел: `off_axial_zero_anchor_2026-10-01.txt`, lambda=3.932502306007, N/Mx/My подтверждены. Последующий полный Off прошел, все 7 регрессий устранены.
- Добавлена изоляция небольших spec/meta-снимков State и StateRequest; 11 новых проверок `audit02.stateSnapshot.*` прошли в последнем Off suite. Волокна не копируются при getters.
- R05/A03 реализован в исходниках: Formation владеет search/context/settings, Pre/Post states и Auto-путями; Width принимает FormationResult и CSectionStateResult, не решает НДС. Adapter/request/result переведены на Formation. Новых production-классов нет. Неиспользуемый StoreCrackResult удален. Width уменьшился примерно с 2865 до 1250 строк по ответственности, а не удалением комментариев. Профильный crack suite `formation_warm_start_tests_2026-10-01.txt`: 360/0, включая Formation-only, reuse и actual-method assertions.
- Первые два crack-прогона выявили ошибки переноса End Function/End Sub и отсутствующий IsZeroLoadProbe; исправлены, константа CRACK_DIRECTION_TOLERANCE сохранена равной исходному 0.001. Последующие прогоны компилируются, но численная приемка не проходит.
- Фактический дефект финализации: верхняя точка LoadMultiplier имеет epsBmax=0.000150123682 при физическом пределе 0.00015. Ранее такая точка могла использоваться шириной. После проверки диапазона она правильно отклоняется; нельзя убрать проверку ради старого green suite.
- Добавлено уточнение точки общим UltimateStrain Newton по тому же пути/критерию и более строгая внутренняя точность критерия 1e-9. Физические пределы и test expected/tolerance НЕ менялись. Последний прогон `formation_boundary_refinement_tests_2026-10-01.txt` завершился exit 1: Newton line search не улучшил невязку, Pre/Post не созданы, далее старый тест доступа к отсутствующему Post получает runtime 91. Это открытый численный дефект, не внешняя проблема Excel.
- Уточняющий Newton игнорировал переданный startSolver. Исправлен приоритет уже найденной плоскости; после этого crack suite прошел. Первый full On после переноса выявил batch 635/34 и UI 297/11 (`formation_full_on_2026-10-01.txt`), поэтому перенос еще не принят по полной интеграции.
- Подробный batch-журнал выявил насыщение нормированной невязки нулевого момента: M/(|M|*relativeTolerance) теряло производную. BuildResiduals теперь нормируется по target, не actual. Общий Newton масштабирует строки/столбцы той же линейной системы, уменьшает разностный шаг около предела и принимает достигнутые допуски перед сравнением round-off нормы. Финальные абсолютные допуски, физические пределы и старые test expected/tolerance не менялись. Добавлены `audit02.pathResidual.*`.
- Исправлена утечка sentinel +/-1e100 у отсутствующего знака напряжений из CSectionSolver в physical snapshot; добавлены `audit02.stressSign.*`. Это устраняет ложную огромную проверку продольных трещин при отсутствии сжатого бетона.
- `formation_accepted_diagnostic_2026-10-01.txt`: пакетный RectSet-порог найден, rsSuccess/CHECK_PASSED, текущий LC еще не достиг его, статус OK. Full On/Off `formation_finalization_full_*` выявили последний физический отказ Formation, который последующая численная проба неверно заменяла NumFail.
- Подтвержденный сошедшимся НДС предел другого материала до порога трещины теперь сохраняется как rsCheckFailed/rcPhysicalLimitExceeded. Неподтвержденная плоскость не используется для такого вывода; Auto может попробовать следующий путь. Все retry-ветви проверяют один и тот же физический критерий.
- Полные On и Off `formation_independent_full_*_2026-10-01.txt` завершились с exit 0: geometry 496/0, materials 1191/0, section 444/0, capacity 972/0, crack 360/0, batch 669/0, UI 308/0, regression 39/0. On: capacity 44.496 s, batch 17.949 s; Off: capacity 45.391 s, batch 18.723 s. Performance Capacity против baseline остается открытым.
- Перед checkpoint добавлены 5 assertions физического отказа Formation и отсутствия фиктивной точки/PreState; `formation_checkpoint_batch_2026-10-01.txt`: 674/0, exit 0. Structural validation также прошла (`formation_checkpoint_validation_2026-10-01.log`). Refresh автоматически формирует читаемый UTF-8 экспорт фактического VBA-проекта; исходный входной экспорт сохранен в HEAD 799bcd8.
- Последний refresh: `docs/regression/Audit02/refresh_search_semantics_2026-10-01.log`, exit 0. Полный suite: `docs/regression/Audit02/current_search_semantics_all_tests_2026-10-01.txt`, exit 0; structural validation также прошла, область печати присутствует.
- Totals: geometry 496/0, materials 1191/0, section 423/0, capacity 972/0, crack 344/0, batch 669/0, workbook UI 308/0, regression 39/0. Capacity 43.996 s, crack 4.227 s, batch 18.582 s, UI 24.695 s.
- Миграция E02 проверена 179 assertions: old/new/both/missing/invalid, сохранение No, warning, соседние настройки/таблицы, validation/help, вставка/перенос строки, идемпотентность и save/reopen. Лог: `docs/regression/Audit02/diagram_extension_migration_2026-10-01.txt`.
- Рабочая книга мигрирована точечно: General.DiagramExtension=Yes в общем блоке, legacy-строка удалена. Полная перестройка пользовательского Config не выполнялась.
- Исправлена потеря print names при Excel COM save в refresh/migration; исходная область печати восстановлена из baseline, остальные настройки не заменялись. Run-AllTests теперь учитывает exit codes вложенных скриптов: прежние логи с exit 0 и Print area=False не доказывают структурную приемку.
- Найденная Capacity-точка с lambda<1 дает Search rsSuccess; инженерный CCapacityResult дает rsCheckFailed/FAIL. Код поиска найденного crack-порога выше LC теперь rcCheckPassed, а не rcCriterionNotReached. Постоянная часть пути дает rsSuccessWithWarning+rcInitialStateBeyondLimit, warning не теряется при упаковке Search. Все новые направленные проверки прошли.
- Search.meta возвращается отдельной копией; последующее изменение этой копии не меняет сохраненный результат. Search numeric snapshot трещин содержит lambda/N/M и полный вектор при наличии PreCrackState.
- Полное самодостаточное состояние Search, реальные границы Formation/Width, упрощение ILimitSearchProblem, единый bracket и миграция всех flat consumers НЕ завершены. До их закрытия цель не выполнена.
- R04/A06 snapshot-перенос прошел `canonical_search_snapshot_full_on_2026-10-01.txt`: все 8 suites без ошибок, Capacity 976/0; 4 assertions подтверждают независимость State от рабочего solver-а и единый объект CapacityState у Search/CapacityResult. Последующий reset повторного CapacityResult добавлен и еще проверяется.
- R04/A05: в исходниках единые Execute/RunSearch/FindBracket/RecoverBracket/Bisection; физическая обработка начальной Formation-точки перенесена к владельцу задачи. Из ILimitSearchProblem удалены 8 crack-специфичных операций. Полный первый прогон завис на оставшемся старом `search.ExecuteCapacity request` в CCapacitySolver; тестовый Excel закрыт по разрешению пользователя, вызов исправлен. Этот прерванный прогон не считается приемкой. Добавлены аналитические matrix/recovery/bound/contract assertions; refresh и повторная приемка в работе.
- Следующий цельный блок: закончить канонический Search snapshot и перенос Formation к существующему CCrackFormationCalculator вместе с данными/настройками, затем перейти к удалению flat/proxy слоя. Не создавать второй search-алгоритм и helper-монолит.
- В дальнейшем обязательны explicit-Off baseline, вся On/Off матрица, output/performance/self-audit и final report; зеленый текущий suite не заменяет эти требования.

## Directed Coverage

| ID | Статус | Проверки / оставшийся объем |
| --- | --- | --- |
| Q05-01 | проверено | `TestAudit02PhysicalDiagramPairs`, `audit02.material.*`: TwoLine/ThreeLine, ULS/SLS, concrete/steel; Q04 solve-пары. |
| Q05-02 | проверено | `audit02.material.*.ignored*`: большие tensile strains, sigma=Et=0, нет tensile extension. |
| Q05-03 | проверено | `TestAudit02OnOffPhysicalResults`: один batch с PR1 Strength Ignore и PR2 Formation UseDiagram; `audit02.pair.*`, provider role tests. |
| Q05-04 | проверено | `audit02.material.*`: активное растяжение бетона и оба знака стали; физический диапазон и extension раздельно. |
| Q05-05 | проверено | `AssertPhysicalDiagramPair`: все узлы/середины, eps_ult, inner 1e-9, outer 0.5e-12/2e-12 с исходным range tolerance 1e-12. |
| Q05-06 | проверено | `TestAudit02ExtendedInitialGuessPhysicalFinal`: подтвержденный extended start -> physical final; final ExtensionUsed=False. |
| Q05-07 | проверено | `audit02.offsetPair.*.onAuxiliary/onExtended/onNotPhysical`, исторические explicit-On progression tests. |
| Q05-08 | проверено | `TestAudit02UnconvergedProbeIsNumerical` и `TestAudit02PositiveUnconvergedProbe`: lambda=0 и 1, несколько итераций, превышение без физического исхода. |
| Q05-09 | проверено | `TestInitialLambdaFailureStatusMapping`, `audit02.offset.*` (0.5/1/1.1) и `audit02.offsetPair.*` (2/3 МН). Точная граница классифицируется по реально подтвержденной плоскости. |
| Q05-10 | проверено | `TestAudit02FormationOutcomeSemantics`, `audit02.formation.constant.*`, fixed-N/fixed-M и Auto path tests; code/warning/no fictitious Post. |
| Q05-11 | проверено | `TestCrackFormationSearchBoundKeepsTechnicalCode`, no-crack/no-Post tests, `audit02.genericLoad.*` обоих доменов и `audit02.boundary.8.*`. |
| Q05-12 | проверено | `audit02.boundary.*`: физические пределы 3, 5.25, 6 при MaxLambda=6; отсутствие предела 8; generic matrix трех методов. |
| Q05-13 | проверено | `audit02.formation.aboveCurrent.*`: найденный порог выше текущего LC, Search success, no current crack/no PostState. |
| Q05-14 | проверено | `audit02.formation.physicalBlock.*`: другой физический предел до tensile-критерия, нет фиктивного PreState/точки. |
| Q05-15 | проверено | `TestCapacityLoadPathMethodMatrix`, zero-component matrix, crack strategy/path tests, generic Ultimate/LoadMultiplier, Q04 On/Off. |
| Q05-16 | проверено | `audit02.steelSign.*` сжатие/растяжение, LoadMultiplier/Ultimate; CreateConfiguredSolver передает оба предела spec, не общий максимум. |
| Q05-17 | проверено | `TestAudit02PsiSignedInputsAndFallbackModes`: отрицательный/нулевой crc, положительный случай, User/Auto/AlwaysCalc и psi1-fallback. |
| Q05-18 | проверено | `audit02.currentCache.*`: фактическая width/sigma/psi идентична без LastRunner, zero new solve. |
| Q05-19 | проверено | `TestCrackFormationCacheHitWithoutLastRunner`, `TestAudit02RepositoryContextAndRetry`: Pre/Post cache-hit, failed -> success, LastRunner отсутствует при reuse. |
| Q05-20 | проверено | `audit02.cache.*`: Extension, spec/role, geometry/material identity/revision, tolerance; warm-start/retry/diagnostics не дробят key. |
| Q05-21 | проверено | `TestAudit02CanonicalResultsAndReset`, generic result reinitialize, independent Formation repeat, `TestRepeatedRun`. |
| Q05-22 | проверено | `TestLinearMaterialEquilibrium`, replacement/runner tests, `audit02.steelSign.*`: Object API без CMaterialDiagram. |
| Q05-23 | проверено | `TestCrackAggregateIncludesCurrentStateFailure`, `TestFormulaChecksDoNotCreateNumFail` и explicit-On beyond-physical tests: Width/Long blocked, первичный Current status в aggregate. |
| Q05-24 | проверено | `TestAudit02SavedResultsIgnoreMaterialChanges` + `snapshot_save_reopen_accepted_2026-10-01.txt`: семь Config values, actual export-reader/plot, raw tables exact, 0 solve. |
| Q05-25 | проверено | `migration.*` 179 assertions и `audit02.migration.reader.*` 28 assertions; save/reopen, validation/help, No/default/conflict/invalid. |
| Q05-26 | проверено | `audit02.technical.*`: физические пределы 12/14, technical 24/28, сохраненная физика и явный overflow InputErr. |
| Q05-27 | проверено | generic invalid/Nothing contracts, State snapshot empty, Formation input/reset, Registry zero-rebar UI; остаточные And/Or содержат безопасные проверки без dereference. |
| Q05-28 | проверено | `audit02.searchEngineering.metaSnapshot`, canonical State identity, `audit02.stateSnapshot.*`, repository returned meta/spec/log isolation, independent Formation snapshot. |

Все строки Q05 подтверждены в итоговых `final_verified_off/on_2026-10-01.txt`,
кроме физической миграции и save/reopen, имеющих отдельные указанные логи.

## Historical Checkpoints

### Актуальная Точка Продолжения Перед Перезагрузкой

- Пользователь сообщил о новой перезагрузке. Работа приостановлена; следующие старые записи этого раздела являются историей предыдущей контрольной точки и не описывают текущие тестовые результаты.
- Последний подтвержденный полный suite: `docs/regression/Audit02/current_capacity_extension_starts_all_tests_2026-10-01.txt`, exit code 0. Geometry 496/0, materials 1191/0, section 395/0, capacity 965/0, crack 328/0, batch 669/0, workbook UI 308/0, regression 39/0.
- Последний проверенный refresh: `docs/regression/Audit02/refresh_capacity_extension_starts_2026-10-01.log`. Активных тестовых сессий нет. Нового коммита нет; HEAD остается `855626e69b07f21dc636771416c501097d077ec9`.
- Исправлены все 10 ранее обнаруженных Batch-регрессий: общий построитель стартовой плоскости учитывает отрицательные вклады вытесненного бетона; прямой runner допускает повторный старт для чистого N и использует extension warm-start только при наличии реально расширенных материалов.
- Проверены неизменность физических точек/касательных диаграмм On/Off, активная tensile-ветвь UseDiagram и сохранение Ignore, безопасные технические границы при физических деформациях больше 10, явная ошибка overflow. Materials suite увеличился с 54 до 1191 успешной проверки.
- ПОСЛЕ последнего зеленого suite изменены `CSystemSettingsReader`, `CExecutionReport`, чтение настройки в `CMaterialModelProvider`, `SettingsCatalog.ps1`; добавлены `tools/build_workbook/Migrate-DiagramExtension.ps1` и `TestAudit02DiagramExtensionReaderMigration` в section tests. Эти последние изменения еще НЕ прошли Excel/VBA проверку и НЕ применены физической миграцией к книге.
- Reader переносит legacy No до defaults, предпочитает явно заданный General при конфликте, удаляет старый ключ из внутреннего registry, сообщает конфликт и отвергает невалидное значение. Новый PowerShell-скрипт предназначен для физической миграции Config/help/validation без полной пересборки пользовательских настроек; требуется реальная проверка временными книгами и повторным запуском.
- Capacity suite занял 46.684 s против baseline 19.832 s. Проверка производительности и сохранение оптимизации probe-cache остаются открытыми.
- Архитектурные R04/R05/R07, значительная часть A/S/C, полная Off/On матрица и итоговый self-audit остаются открытыми. Зеленый suite не означает завершение цели.

### Порядок Продолжения После Этой Перезагрузки

1. Прочитать Git status/diff/log, Audit02 spec и эту актуальную точку; сохранить пользовательские изменения. Старые диагностические указания ниже сверять с последним зеленым логом.
2. Проверить BOM нового PowerShell-скрипта и создать проверку физической миграции на временных книгах: old/new/both/missing/invalid, конфликт, идемпотентность, сохранение соседних настроек, validation/help, save/reopen. Не запускать Excel задачи параллельно.
3. Выполнить миграцию книги и refresh последних VBA-правок; пройти весь suite. Сохранить логи, обновить карту требований и сделать разрешенный локальный checkpoint-коммит только проверенного блока.
4. Продолжить реальные архитектурные изменения по Audit02: вынести Formation в существующего владельца, объединить generic Search, перейти всеми consumers на canonical typed results, проверить статусы/snapshots/reuse/probe-cache. Не добавлять задачи вне ТЗ.
5. Выполнить полный baseline Off и historical explicit-On, направленную матрицу, количественные сравнения, output/smoke/performance и self-audit всех ID. Только затем создавать final report и завершать цель.

### История Предыдущей Контрольной Точки

- Контрольная точка перед перезагрузкой пользователя 2026-10-01 обновлена после последних правок. Работа приостановлена по сообщению пользователя. Все изменения сохранены на диск; нового коммита нет, HEAD остается baseline 855626e6.
- Excel COM работает вне sandbox. Последний полный suite завершился с exit code 1; на момент контрольной точки процесс EXCEL не найден, активных тестовых сессий нет.
- Baseline исходники: `C:\Users\avidzba\AppData\Local\Temp\RC_NDM_Audit02_Baseline_855626e6`. Их сборка и полный suite завершились успешно; отчет скопирован в `docs/regression/Audit02/baseline_all_tests_2026-10-01.txt`.
- Последний refresh VBA успешно завершен: `docs/regression/Audit02/refresh_typed_ultimate_2026-10-01.log`. Активных процессов Excel и тестовых сессий перед перезагрузкой не осталось.
- Последний full suite: `docs/regression/Audit02/current_retry_flags_all_tests_2026-10-01.txt`. Capacity: 962 passed / 0 failed; Batch: 659 passed / 10 failed. Последующий batch-прогон `batch_shared_flags_diagnostics_2026-10-01.log` подтвердил те же 10 ошибок. Приемка не пройдена.
- Исправлены три ошибки асимметричного осевого сжатия UltimateStrain: начальная деформация выбирается по минимальному физическому пределу. Capacity-проверки прошли до последнего изменения typed Ultimate-интерфейса; это последнее изменение еще не проверено тестами.
- Batch FAIL: пары `batch.rectset.n200.Auto/UltimateStrain/LoadMultiplier.notNumFail` и `.capacityStatus`, а также `batch.capacityPath.nmxyZeroM.notNumFail`, `batch.crack.pathAuto.summary.notNumFail`, `batch.crack.pathN.summary.notNumFail`, `batch.group1.extension.overall`.
- После явной проверки равновесия `UltimateStrain` продолжает Newton, если относительные критерии выполнены, но кандидатная точка не проходит абсолютную проверку N/Mx/My. Нужно выяснить оставшиеся причины отказа без ослабления старых тестов и без изменения расчетной методики.
- Направленная проверка несошедшейся probe-плоскости прошла, включая отсутствие ложного BaseFail. Проверка MaxLambda для ожидаемого корня 8 при границе 6 также прошла.
- R04/R05/R07 и значительная часть остальных требований Audit02 остаются открытыми. Финальная приемка и self-audit не выполнены.
- В `ILimitSearchProblem` и адаптерах удален `DomainContext`; `CUltimateStrainSearch` получает typed callbacks вместо downcasts к Capacity/Crack. Добавлен аналитический тест общего Newton. После этих правок выполнен только refresh VBA, тесты необходимо запустить первыми после продолжения.
- Capacity State и сохраненные crack states получают признаки extension/physical range из общего `CStateSolutionRunner`, а не из констант или дублирующих методов Width. Capacity-контракт проверен full suite; crack-контракт проверен batch-прогоном, но требует направленных flag-тестов.
- Для двух crack-path ошибок диагностика показывает действительную несходимость CurrentCrackedState (`Line search не смог уменьшить невязку`), хотя Formation найден выше текущей нагрузки. Нужно исследовать начальное приближение `CStateGuessBuilder`, а не скрывать этот отказ статусами.

## Important Decisions

- Не начинать новые архитектурные задачи за пределами `docs/NDM_Audit02_Implementation_Spec_2026-10-01.md`.
- Пользователь разрешил самостоятельно закрывать ошибки и принудительно завершать Excel при зависании тестов; на момент разрешения активных пользовательских книг нет. Перед завершением процесса проверить, что это восстановление тестового прогона, а не обычное закрытие с потерей новых пользовательских данных.
- Не откатывать входные изменения в `workbook/output/*` и новый spec-файл.
- Любое изменение статуса должно проходить через `InternalStatus/ResultCode -> CResultStatusPolicy -> ExternalStatus`, без маппинга по тексту комментария.
- `NumFail` допустим только для state/search routes, где реально решалось равновесие или предельная точка.
- Writer-ы не должны заново формировать инженерские причины отказа; они читают готовые result comments.
- Excel COM-сборку и VBA-тесты запускать вне sandbox: внутри sandbox Application не создается с CO_E_SERVER_EXEC_FAILURE; вне sandbox минимальная проверка проходит.
- Для R04/R05 использовать существующие CCrackFormationCalculator и search adapters; не создавать новый helper-монолит.
- R04/A06: CLimitSearchResult принимает только готовые meta/числа/State/diagnostics. Доменную интерпретацию Capacity выполняет существующий CCapacityLimitSearchProblem, Formation формирует свой snapshot у владельца задачи. Production Capacity передает фактический MaterialSpec для единственного CapacityState; низкоуровневый Object-material API без профиля сохраняет численный вектор без вымышленного spec. Живой LastSolver и инженерный utilization из Search-result удаляются.
- Несошедшаяся probe-плоскость не подтверждает физический предел. Failed probe-cache не должен препятствовать повтору с лучшим warm-start.
- `EvaluateStrainPlane` является только оценкой плоскости; физический State требует отдельного подтверждения равновесия с конкретной нагрузкой. Кандидатный UltimateStrain-предел нельзя принимать только по относительному направлению момента.
- Snapshot восстанавливается через общий State-слой без нового solve и без зависимости от LastRunner. Прошедший cache-hit тест подтверждает не только статус, но и фактический расчет ширины.
- Точная аналитическая Offset-нагрузка решается с исходными допусками усилий. Классифицируется фактическая подтвержденная плоскость: utilization >= 1 исчерпывает начальную Capacity-точку; небольшое округление ниже 1 не объявляется превышением по одному входному N. Физические eps_ult/range tolerance не менялись.

## Known Risks / Open Questions

- Входной `workbook/output/VBA_All_Code.txt` отличается от baseline export; это входное состояние задачи и должно быть сохранено.
- AutoCAD smoke-проверки могут остаться ручными, если AutoCAD недоступен.
- Full Excel tests могут требовать закрытого workbook; при зависшем Excel действовать по правилам `AGENTS.md`.

## Historical Last Verified Before Checkpoint 799bcd8

- `git rev-parse HEAD`: `855626e69b07f21dc636771416c501097d077ec9`.
- Dirty-tree включает исходники, тесты, SettingsCatalog, документацию и generated workbook/export/report; ничего не откатывалось.
- Baseline totals: geometry 496/0, material 54/0, section 390/0, capacity 934/0, crack 328/0, batch 659/0, workbook UI 308/0, regression 39/0.
- Последний current full-suite totals: geometry 496/0, material 54/0, section 395/0, capacity 962/0, crack 328/0, batch 659/10, workbook UI 308/0, regression 39/0.
- Время последнего current suite: capacity 45.875 s, batch около 19.45 s. Последующий batch-прогон занял около 17.81 s. Ухудшение capacity относительно baseline 19.832 s остается открытым и требует проверки локального probe-cache.
- Все 5 новых проверок `audit02.plane.*` и 10 проверок `audit02.currentCache.*` прошли.
- `Refresh-VbaModules.ps1` выполнен после typed Ultimate-правок. Книга и VBA export обновлены, но эти правки еще не прошли тестовый прогон; успешный refresh не означает успешную приемку.
- `git diff --check` завершился с exit code 0; только предупреждения о будущей нормализации LF/CRLF.

## How To Continue

1. Восстановить контекст по Current/spec/Architecture и git status/diff/log; HEAD=020ff8f9. Не останавливаться для перезагрузки без явного запроса. Excel COM только последовательно.
2. Итоговые clean On/Off, output refresh/Run-AllTests и structural validation прошли; не запускать заново без изменения соответствующего кода. Негативные логи сохранены отдельно.
3. Завершить статический self-audit/source-book consistency, записать final report с таблицей всех ID и Old-T/Old-A, метриками и количественными MAXDIFF/performance.
4. AutoCAD COM ProgID отсутствует: реальный DWG smoke не засчитывать. Actual export-reader/save-reopen/plot проверены в Excel, 0 дополнительного solve.
5. Зафиксировать принятый итоговый срез, указать SHA и только после сохранения всех доказательств завершить goal.

### Исторические Диагностические Чтения До Реализации

- `CBatchSectionCalculator`: ApplyLoadReference / NormalizeCombinationMoments и LoadStateForCombination; `CSectionSolver.ClearResult`; `CExecutionReport` API для attachment диагностики Capacity.
- `CStateGuessBuilder.AccumulateLinearElement` пропускает неположительные tangent contributions, тогда как solver учитывает их. Проверить причину неудачного active-set старта без изменения физической модели.
- В Material provider еще нужно включить extension активной tensile concrete-ветви UseDiagram, сохранить Ignore, сделать безопасную техническую границу при физических деформациях вне +/-10 и сохранить tangent в исходном физическом узле сжатия.
