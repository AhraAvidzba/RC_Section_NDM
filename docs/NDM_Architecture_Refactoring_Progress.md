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

## In progress

- Следующие изменения должны начинаться с этапа 2 migration plan.

## Next

- Использовать baseline commit и отчеты `Stage00`/`Stage01` как точки сравнения
  для следующих этапов.
- Перед началом этапа 2 снова проверить `git status` и убедиться, что нет
  посторонних пользовательских изменений.

## Known risks / open questions

- Для AutoCAD Region geometry часть проверок может оставаться ручной, если
  тестовый AutoCAD недоступен в автоматическом прогоне.
- В дальнейшем не смешивать миграционные изменения разных этапов в один diff
  без отдельной проверки.
- Этап 1 оставляет часть старых строковых status-полей как совместимый внешний
  контракт writer-ов. Полный перенос на result graph выполняется последующими
  этапами плана.

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
- 2026-09-28: `tools/build_workbook/Run-BatchTests.ps1` завершился успешно:
  `passed=606`, `failed=0`.
- 2026-09-28: `tools/build_workbook/Run-AllTests.ps1` завершился успешно;
  итоговый отчет сохранен в `docs/regression/Stage01_AllTests_Report.txt`,
  regression baseline raw report сохранен в
  `docs/regression/Stage01_RegressionBaseline_Raw.txt`.
