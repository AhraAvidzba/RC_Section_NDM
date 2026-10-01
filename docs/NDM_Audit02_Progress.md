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
| R01-R09 | в работе | R01 подтвержден интеграционным cache-hit тестом без нового solve. R06: EvaluateStrainPlane отделен от проверки равновесия. Полная приемка R02/R03/R08 не пройдена; R04/R05/R07 остаются открытыми. |
| A01-A08 | в работе | DomainContext/downcasts удалены; успешный Search отделен от инженерного FAIL. Formation перенесен к существующему владельцу, отдельный crack suite 360/0; полная интеграционная приемка в работе. Контракты Search и flat-дубли открыты. |
| E01-E10 | в работе | General.DiagramExtension действует на расширенные материалы текущих маршрутов. E02 проверен на временных и рабочей книгах; полная матрица и API cleanup не завершены. |
| S01-S07 | в работе | S03/S04: warning-код постоянной части и успех порога выше текущего LC сохранены; направленные тесты прошли. Полная статусная приемка открыта. |
| C01-C06 | в работе | Cache-hit и shared runner проверены частично; snapshots, scoped invalidation и transient probe-cache требуют завершения. |
| Q01-Q08 | в работе | Baseline On и explicit Off прошли. После Formation переноса повторяется current full suite; полная численная On/Off матрица, output и performance acceptance остаются открытыми. |
| D01-D02 | в работе | Источник Config и справка глобального Extension обновлены; архитектурные документы и все текущие API-комментарии еще не завершены. |
| W01-W06 | в работе | Baseline и входные изменения сохранены; тестовые логи и точка продолжения записаны, готовится первый проверенный Audit02 checkpoint. |

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

- Актуальный блок R04/A05/A06: канонический Search snapshot и общий LoadMultiplier реализованы. Аналитическая матрица проверяет оба вида задачи, три метода, найденный предел/точную MaxLambda/неисследованный предел, recovery и повторную инициализацию. Отдельные численные циклы Capacity/Formation удалены; физические callbacks остаются у владельцев задач.
- Монотонный recovery после успешной промежуточной пробы изменил исторический порядок warm-start и вызвал 11 Off-регрессий осевого RectSet. Причина устранена в общем FindBracket: успешная проба снова расширяет нагрузку с улучшенным стартом, несошедшаяся проба уменьшает шаг и не объявляется физическим пределом. `shared_expansion_full_off_2026-10-01.txt` и `shared_expansion_full_on_2026-10-01.txt`: каждый прогон содержит все 8 suites, 4706 assertions, 0 ошибок, исходная книга неизменна. Structural validation прошла (`shared_search_validation_2026-10-01.log`). Ожидания/допуски не менялись. Capacity On 46.785 s / Off 45.180 s; Q08 performance остается открытым.
- Следующий цельный R07/A02/A07 блок: удалить скалярные копии CCombinationResult и indexed forwarding getters Batch; все production/test consumers перевести на ResultAt(index) и typed subtree. ProfileId и снимок глобальной настройки остаются собственными LC-данными, статусы/итоговое использование вычисляются из subtree. Нейтральные незапрошенные result-ветки имеют typed N/A, но не создают фиктивное физическое State. Новые production-классы не нужны.
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
| Q05-01 | проверено для material API | `audit02.material.*` TwoLine/ThreeLine, ULS/SLS, бетон/сталь; требуются парные solve-сценарии. |
| Q05-02 | проверено для material API | `audit02.material.*.ignored*` и suite материалов. |
| Q05-03 | в работе | Эффективные роли provider; нужен отдельный multi-role integration test. |
| Q05-04 | проверено для material API | `audit02.material.*` active compression/tension. |
| Q05-05 | в работе | Все узлы/середины сегментов проверены; малая окрестность физической границы еще не покрыта полностью. |
| Q05-06 | не начато | Нужна проверка промежуточного выхода и физического финала. |
| Q05-07 | в работе | Исторические explicit-On state-тесты сохранены; требуется полная directed приемка. |
| Q05-08 | проверено частично | `audit02.failedProbe.*`; добавить положительную lambda. |
| Q05-09 | в работе | Старый `TestInitialLambdaFailureStatusMapping`; точная начальная граница еще требует проверки. |
| Q05-10 | в работе | `audit02.formation.constant.*`, старые fixed-N/fixed-M fallback; Auto warning history еще требует проверки. |
| Q05-11 | в работе | Старые crack search-bound/no-crack и `audit02.boundary.8.*`; generic result contract еще не завершен. |
| Q05-12 | проверено | `audit02.boundary.*`: пределы 3, 5.25, 6 при MaxLambda=6; отсутствие предела 8. |
| Q05-13 | проверено | `audit02.formation.aboveCurrent.*`: успешный найденный порог, no current crack, no PostState. |
| Q05-14 | не начато | Ограничение другого материала раньше crack-критерия. |
| Q05-15 | в работе | Существующие strategy/path matrices прошли; нужны новые On/Off сравнения. |
| Q05-16 | проверено частично | `audit02.steelSign.*` обоих знаков и методов; сверить загрузку sign-specific limits из spec. |
| Q05-17 | в работе | Старые User/Auto/AlwaysCalc сохранены; directed zero/negative sigma_crc еще не завершены. |
| Q05-18 | проверено | `audit02.currentCache.*`, 10 assertions, фактическая ширина и zero new solve. |
| Q05-19 | в работе | Старый `TestCrackFormationCacheHitWithoutLastRunner`; retry/cache acceptance еще требует проверки. |
| Q05-20 | не начато | Scoped context/geometry/material/tolerance invalidation. |
| Q05-21 | не начато | Повторное использование всех объектов после удаления дублей. |
| Q05-22 | проверено частично | Существующие linear-material runner tests и `audit02.steelSign.*`; общий контракт сохранить. |
| Q05-23 | в работе | Current meta включена в aggregate; добавить directed state failures. |
| Q05-24 | не начато | Смена Config после сохранения snapshot. |
| Q05-25 | проверено | `migration.*` 179 assertions и `audit02.migration.reader.*` 28 assertions; actual output и validation подтверждены. |
| Q05-26 | проверено для material API | `audit02.technical.*`: физические пределы 12/14, technical 24/28, overflow InputErr. |
| Q05-27 | в работе | Несколько guards исправлены; оставшиеся Nothing contracts еще проверяются. |
| Q05-28 | в работе | `audit02.searchEngineering.metaSnapshot`, crack diagnostic snapshot; live solver snapshot еще необходимо устранить. |

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

1. Восстановить контекст по Current, spec и Git; HEAD=799bcd8. Не останавливаться для перезагрузки без явного нового запроса.
2. Дождаться `formation_finalization_full_on` и закрыть оставшиеся интеграционные ошибки, если появятся. Не снимать physical guards и не ослаблять expected/tolerance. Профильный Formation suite 360/0 уже подтвержден.
3. Выполнить полный current Off: осевой диагностический seed с нулем на противоположном краю подтвердил физический предел, но семь regression-ветвей еще должны пройти. Baseline Off подтвержден exit 0.
4. После согласованного зеленого блока - full suite, scoped checkpoint, затем единый generic bracket/контракт и переход всех flat/proxy consumers на typed tree.
5. Завершить оставшиеся C/S/E/Q/D требования, On/Off численные сравнения, snapshots, output/performance и полную финальную сборку.
6. Финальный self-audit каждого ID и final report только после приемки; непроверенную реализацию не объявлять завершенной.

### Следующие диагностические чтения

- `CBatchSectionCalculator`: ApplyLoadReference / NormalizeCombinationMoments и LoadStateForCombination; `CSectionSolver.ClearResult`; `CExecutionReport` API для attachment диагностики Capacity.
- `CStateGuessBuilder.AccumulateLinearElement` пропускает неположительные tangent contributions, тогда как solver учитывает их. Проверить причину неудачного active-set старта без изменения физической модели.
- В Material provider еще нужно включить extension активной tensile concrete-ветви UseDiagram, сохранить Ignore, сделать безопасную техническую границу при физических деформациях вне +/-10 и сохранить tangent в исходном физическом узле сжатия.
