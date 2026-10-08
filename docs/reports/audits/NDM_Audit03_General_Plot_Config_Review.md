# Audit03: Общие Настройки Excel-Схемы Из Настоящего Config

## Scope

Срез после `2d01dbe`. Меняются только существующий `modTestPlotConfig`,
per-key evidence и progress. Production, расчетная методика, численные
expected/tolerance прежних тестов и классы не меняются.

Прежний general-plot fixture сворачивал System/Annotation в трехколоночный
Range. Его направленные проверки сохраняются, но вход переводится на
настоящие именованные System/Unit/Sign/Annotation таблицы через
`CSystemSettingsReader.LoadFromWorkbook`. System переносится из A5 в CH800;
динамические адреса ошибок проверяются в обоих положениях. Компактная
таблица аннотаций сохраняет свой формат и отдельное имя.

## Поведенческие Проверки

- Активные значения 15 параметров оформления: blank/TODO/INVALID/CVErr,
  missing key, адрес, действие, сохранность прежнего Chart и recovery.
- Contour/NeutralLine/LoadPoint/PrincipalAxes/ResultLabels/Legend: включение
  и выключение меняют реальные фигуры или подписи, не только reader fields.
- AxisLabelsEnabled/FontSize: настоящая AxisLabelX получает 18 pt; No
  убирает X/Y и не читает ошибочную неактивную высоту текста.
- Все четыре цвета: RGB(17,34,51) проверяется в Fill.ForeColor.RGB
  соответствующего реального бетонного или арматурного элемента.
- Separate/Common: цвет бетона меняется с RGB(10,20,30) на общую
  арматурную шкалу RGB(100,110,120). При градиенте и отношении 1.25/2.75
  независимый oracle равен RGB(164,169,173).
- Второе сочетание ALT имеет отдельные State/element rows, без дублирования
  ALL properties. Выбор ALT вместо PLOT возвращает 9.5 MPa и подпись 9.50.
- Небольшой положительный шаг подписей 1e-8, отклонение 0/-1,
  geometry-preview inactive contract, 15 сочетаний стрелок, стабильная
  рамка Chart и все 17 gradient buckets на сетке более 5000 элементов.
- Соседние включенные проверки используют настоящие workbook entrypoints
  для Plot.Enabled и AutoUpdateAfterCalculation, готовые snapshots,
  read lifecycle, semantic arcs и компактные annotation cells.
- После собственных контролируемых замен element rows восстанавливается
  исходный synthetic Results. Итоговое сравнение снимка и счетчик solve
  подтверждают отсутствие непредусмотренной записи Results и решения НДС.

## Gates И Исправление Тестового Ожидания

Первый v234: 3218/2. Обе ошибки относятся исключительно к новому ожиданию
текста ALT: тест ожидал `9.5`, хотя fixture явно задает StressPrecision=2,
а действующий formatter сохраняет два знака `9.50`. Число reader-а,
цвета и остальные проверки прошли. Это не найденный дефект production;
журнал с exit=1 сохранен, не переименован в PASS.

В v235 исправлено только это новое ожидание. General presentation:
3220/0, watchdog exit=0, source unchanged=True. В журнале счетчик
GENERAL_PLOT_CASES=211; он не является числом всех assertions набора.
Read-only структура: 27/27, исходная книга не изменена.
Реальный VBE export содержит 111 components. Source/export equality:
106/106, failed=0; подтверждено отдельным source-contracts журналом,
не выведено из одного успешного VBA run.

## Артефакты

Final fixture `RC_Section_NDM_general_plot_named_v235.xlsm`:
SHA-256 `A356D31B2EC11991E6ABF0D65424A341EAB356D860846E75D41EBF7952605A83`.
Actual export `VBA_All_Code_general_plot_named_v235_2026-10-04.txt`:
SHA-256 `821BB0C981645A9960229E64833B066102675BE41FFD6C8CEA85291704105B86`.
Directed logs: `general_plot_named_v234.log`, `general_plot_named_v235.log`.
Structure: `general_plot_named_structure_v235.json/.log`.
Per-key evidence: `NDM_Audit03_General_Plot_Config_Evidence.json`.
Registry: `config_behavior_registry_general_plot_v235_2026-10-04.json/.csv`.

## Ограничения

Адресно принято 18 дополнительных active-behavior полей; итог 705/1064,
из 760 editable полей 55 pending. `fullAcceptance=False` сохранен. Это
не blanket PASS всех диапазонов, aliases/Worst/Transformed, RGB/Double
краев, отсутствующего Plot.Enabled, mutation sensitivity и pairwise.
Реальные графические свойства проверены COM; новый визуальный screenshot
gate не заявлен, так как production/layout не менялись.

Полный восьми-suite On/Off не повторяется для test-only presentation среза
по разрешению пользователя. Он остается обязательным для финального выпуска.
Main output, пользовательский Excel PID 23476 и audit specification не
изменялись. Следующий объем: оставшиеся Config, range/pairwise, metadata/
geometry guard paths, semantic review и общий release gate.
