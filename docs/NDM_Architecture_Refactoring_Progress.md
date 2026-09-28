# NDM Architecture Refactoring Progress

## Baseline

`6b9a6bea62e0436d206e698845332b6ad17fa231`

Baseline commit создан перед началом архитектурных этапов рефакторинга. Его нельзя переписывать или разрушать; для сравнения использовать `git show`, `git diff` и regression-отчеты.

## Completed

- Создан baseline commit перед началом архитектурных изменений.
- Рабочее дерево после baseline commit проверено через `git status`.
- Migration plan и укрупненный архитектурный план сохранены в `docs/`.
- Зафиксирован progress/status MD как внешняя память рефакторинга.
- Книга собрана через `tools/build_workbook/Build-Workbook.ps1`.
- Полный набор существующих тестов запущен через `tools/build_workbook/Run-AllTests.ps1`.
- Baseline-отчет полного прогона сохранен в `docs/regression/Stage00_AllTests_Report.txt`.
- Regression baseline raw report обновлен существующим тестовым скриптом `Run-RegressionBaselineTests.ps1`.

## In progress

- Этап 0 завершен; следующие изменения должны начинаться с этапа 1.

## Next

- Использовать baseline commit и test report как точку сравнения для этапа 1.
- Перед началом этапа 1 снова проверить `git status` и убедиться, что нет посторонних пользовательских изменений.

## Known risks / open questions

- Для AutoCAD Region geometry часть проверок может оставаться ручной, если тестовый AutoCAD недоступен в автоматическом прогоне.
- В дальнейшем не смешивать миграционные изменения разных этапов в один diff без отдельной проверки.

## Important decisions

- `CStateRepository` должен хранить reusable только корректно полученные конечные named states.
- Search probe-cache остается локальным для search-session и не попадает в shared repository.
- Неуспешные solve attempts сохраняются только как диагностика/result, но не как reusable physical State.
- `CCrackWidthFormulaCalculator` остается чистой формульной частью и не получает State/status.

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

- 2026-09-28: `tools/build_workbook/Build-Workbook.ps1` завершился успешно.
- 2026-09-28: `tools/build_workbook/Run-AllTests.ps1 -ReportPath docs/regression/Stage00_AllTests_Report.txt` завершился успешно, итоговый отчет `failed=0`.
