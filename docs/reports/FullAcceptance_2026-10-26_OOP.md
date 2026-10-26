# Полная приёмка NDM 26.10.2026 — OOP/статусы/локали

Полный обязательный тестовый состав и нативная готовая книга приняты на SHA256 `AFF765CB9D7126158534C1752DDF0BEE475C3C3765111C66CFD84B4C4AC38698`. Это один запуск 03:08:53–06:28:40 МСК, SourceCommit `75de262f80e60c638b1c7cc8cd91e861495bc844`; никаких новых длинных повторов после него не было.

| Обязательная часть | Результат |
|---|---|
| Core ON/OFF, CoverageVersion5 | 12/12 групп, 168/168 макросов, ошибок 0 |
| Полный UI | No 70241/0; Yes 70241/0, оба хвоста завершены |
| Focused | 26/26 записей, 122/122 точных ID/порядка, 14000/0 |
| Batch, устойчивость, разная сталь | batch 15 No/Yes, Stability No/Yes, AsymmetricSteel No/Yes — PASS |
| Save/reopen | результаты и статусы/стили подтверждены в каждом требуемом отчёте |
| Настоящий AutoCAD | FarOrigin 447/0; FarCapacity 503/0, SourceUnchanged=true; пользовательские документы не использовались |
| Независимый контроль | 562 интегрирования PASS; 10 ожидаемых NumFail разобраны PASS; Unsupported=0 |
| Замороженные входы | 498/498 совпадали до конца всех процессов, HEAD75de неизменен |
| Нативный финальный output | 178 компонентов (172+6), источники и AllCode совпали; ReadOnly, без импорта/Save |
| Данные/оформление | 106 non-VBA ZIP частей побайтно сохранены относительно исходной книги1D7 |

Исходный outer Completion — REVIEW_REQUIRED/RunnerExit1. После полного исполнения и успешных verifiers произошёл PS5.1 GenericMeasurePropertyNotFound в записи ExpectedSuites финального футера. Ошибка независимо воспроизведена, исправлена одной строкой и затронутый шаг повторён на фактическом manifest; весь состав повторно прочитан и штатные Summarize/Finalize дали exit0. Этот исходный служебный сбой и его exit1 сохранены, не маскируются новым Completion. Книга и расчётные результаты при исправлении не менялись. Доказательство: `OOPLayers_2026-10-26_Run01/RunnerFooterFix/ExitReconciliation.json`.

Полный Source — frozen AFF, а Probe имеет другой SHA только из-за разрешённого тестового helper; preparation/input manifest/исходные компоненты и exact-case verifier связывают его с AFF. Это не другая производственная книга. Первичные 12 core +26 Focused и оба child NativeCAD отчёта прочитаны целиком, SHA сверены. Все 42 исходные собственные идентичности и финальный readonly Excel завершены; пользовательский Excel не трогался.

AllCode SHA256 `2688FF2C07BC4A903811F900C2E35E8655FB967FCB0B98677E90288ABDD79D67`. Финальная книга: `workbook/output/RC_Section_NDM.xlsm`. Отчёт по изменениям, поэтапным тестам и ограничениям локали: `NDM_OOP_Layers_Status_Refactoring_2026-10-26.md`. Новых VBA-классов/модулей 0; математика/единицы/форматы снимков не менялись.

Основные доказательства: CompleteCurrentAcceptance.json и Focused/Acceptance.json исходного полного каталога, ClosedPrimaryReview/FocusedPrimaryReview, RunnerFooterFix, FinalReleaseGate. Архив FullRequiredLogs.zip сохраняет оригинальные raw/metadata и ошибочный outer exit; manifest Processed.json перечисляет SHA и точный состав. Публикация выполняется в codex/linked-rebar-row-offsets после этой приёмки; успешный push подтверждается отдельной сверкой remote.

## Подтверждение публикации

Изменения опубликованы в codex/linked-rebar-row-offsets: a9372807, remote HEAD совпал. Книга AFF и AllCode2688 совпали с проверенными, исходный ТЗ025C сохранён точно; пользовательский execution_report/outputs не затронуты. Существующее ndm подтверждено PAUSED после полной приёмки/отчёта/push, расписание и prompt сохранены. Собственная одноразовая Windows-задача RC_NDM_OOP_20261026_Run01 удалена. PublicationVerification.json хранит подтверждение; следующий коммит содержит только эти итоговые метаданные.
