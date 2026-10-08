# Audit03: Auto Plot

Дата: 2026-10-04. Исходный checkpoint:
`1400b3b42eb76cc75353ac503c1a60bdb9dda255`.

## Дефект И Исправление

RunSectionCalculationForWorkbook читал Plot.AutoUpdateAfterCalculation
после расчета и записи Results, с default=True для отсутствующего ключа.
Plot.Enabled не участвовал в ветке очистки схемы при отсутствии named-state.
В результате отключенная схема очищалась, а неактивный CVErr AutoUpdate
останавливал уже выполненный запуск.

Подтвержденный negative v217d: **45/36**. Неверные активные переключатели
не проверялись до построения геометрии; missing AutoUpdate не имел ранней
диагностики; Plot.Enabled=No не сохранял Chart без named-state и потреблял
неактивный CVErr. Журнал и прежняя книга сохранены, не перезаписаны.

Исправление в существующем workbook-entrypoint:

- Plot.Enabled обязателен и проверяется до модели и очистки Results.
- AutoUpdate обязателен только при Plot.Enabled=Yes. При No не читается.
- Результаты early validation кешируются двумя локальными Boolean, без
  новых классов/инженерного policy и повторного чтения этих flags в конце.
- No любого активного переключателя сохраняет существующую схему.
- Только Enabled=Yes + AutoUpdate=Yes разрешает обновление либо очистку
  прежней схемы, если расчетных состояний нет. Очистка имеет пояснение.
- Расчетные статусы, критерии, solver/search и expected/tolerance не менялись.

Комментарии Config и полная справка объясняют обязательность, inactive No,
ранний отказ и поведение при отсутствии данных. Новых настроек не добавлено.

## Отделение Ошибок Fixture

v217 (**42/36**) использовал профиль без активных проверок: четыре запуска
правильно возвращали InputErr, поэтому не считались доказательством валидного
stability-only сценария. v217b (**42/39**) ошибочно искал PR1 в первой строке
шапки и читал ResultComment вместо общего статуса; noSolve assertion выявил
оставшиеся включенными solver-потребители. v217c (**35/35**) требовал ключ
ProfileId во второй колонке, которого реальная схема не гарантирует.

Fixture исправлен по фактическому catalog contract: PR1 ищется в двухстрочной
шапке динамически, изменяется только его колонка; включена только устойчивость,
для которой named НДС не нужен. Общий статус читается из правильной колонки.
Эти ошибки нового теста не считаются production defects. v217d полностью
завершил все четыре запуска, подтвердил валидный профиль и нулевой solve count.

## Приемка

- Directed v218: **82/0**, 20 invalid/switch вариантов, одна собственная
  полная копия книги, equilibriumCases=0. Проверены перенос rngSystemSettings,
  фактические адреса, Results/Chart до early failure и inactive AutoUpdate.
- Frozen v219, General Plot + metadata + AutoUpdate: **1962/0**.
- Imported snapshot после смены units: **228/0**; workbook-сценарий и
  корректный возврат после invalid input/output units сохранены.
- Read-only structure: **27/27**. Config formatting: **1003/0**, 1002 адреса,
  deviations=0, без исправления форматирования проверяющим helper.
- Help: **2004** строки, **140** прямых links, **118** shapes, failed=0;
  пользовательские values/formulas/validation и PrintArea сохранены.
- Palette/save-reopen: **355/0**. Results before/after SHA
  `E05C0BF950A1F5DA36E0BF42A1B9A7C7C68FBA4EA8006FC9B4F2470AC42B6C96`,
  styles SHA `1EABA40A72D5A28392F11FC5D38F0D2B1E9B264FDFAD9F5514ACB9EE1A1E1132`.
- Frozen book `RC_Section_NDM_auto_plot_positive_v219.xlsm`, SHA
  `714394331AC3930418E8F21C785E7CDBCF2A9972D9513329F7886C2E5CBEF95F`.
- Actual read-only VBE export: **111** компонентов. Export
  `VBA_All_Code_auto_plot_v219_2026-10-04.txt`, SHA
  `A3FC7805EBB2378C770388C12DA014E366942C492A8FF2F079A78DF6336211B3`;
  canonical VBA_All_Code.txt содержит эти фактические байты.
- Source/export equality: **106/106**, failed=0; deleted-class consumers
  отсутствуют, 83 production + 3 test classes сохранены.
- Census: **106** modules/**4415** methods/**1705** guard candidates,
  semantic acceptance Pending; новых классов нет.

Все watchdog source unchanged=True. Main output/ТЗ и пользовательский Excel
PID 23476 не менялись. Полный Off/On не повторялся по разрешению пользователя:
последний полный numerical gate v206 **82221/0 Off**, **82232/0 On**, финальный
release gate остается обязательным.

## Границы

Общий Audit03 не завершен. Directed-набор проверяет управление схемой без
named-state, не новую численную нагрузочную матрицу и не пиксели Excel.
Обновление схемы с реально рассчитанным НДС остается частью финального
workbook/UI release gate. Полный Config range/pairwise, metadata resolvers/
CONTOUR_ARC parsing, Excel guard, extreme arithmetic и semantic review
не получают blanket PASS от данного среза.
