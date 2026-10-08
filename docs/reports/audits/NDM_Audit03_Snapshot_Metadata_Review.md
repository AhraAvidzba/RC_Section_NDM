# Audit03: Snapshot Metadata

Дата: 2026-10-04. Исходный checkpoint: `ce232662e26477aeb481675706d0659aa5ef4784`.
Срез относится к чтению сохраненного Results для схемы, не к численному ядру.

## Подтвержденные Дефекты

- При пустой/служебной Unit-ячейке пересчет общих координат и кривизн зависел
  от порядка строк: Output.*Unit, расположенные ниже численных свойств,
  прочитывались слишком поздно. Например, Bounds.MinX в cm восстанавливался
  как -10 mm вместо -100 mm; для m получалось -0.1 mm.
- Неизвестный/пустой/CVErr ExtensionUsed молча превращался в False; статус
  допускал неизвестный либо пустой текст. Это могло скрыть предупреждение
  о физической недопустимости сохраненного состояния.
- CVErr в Unit-ячейке свойства или аннотации подменялся fallback-единицей.
  Ошибочные Output.*Unit давали пустое значение, Type mismatch или проходили
  без содержательной диагностики текущей ячейки.

Новый directed-набор на прежнем production: v212, **183/266**. Журнал не
перезаписан. Это assertions presentation/metadata, не задачи равновесия.

## Исправления И Границы

- CSectionPlotDataReader сначала читает активные единицы snapshot, затем
  численные свойства. Коэффициенты пересчета остаются в CUnitSystem.
- Явная единица численного поля имеет приоритет над metadata. Пустое/"-"
  поле Unit использует единицу snapshot; оно не означает текущие units Config.
  Существующий контракт неполного снимка без Output.*Unit не ужесточается.
  Если строка Output.*Unit присутствует, ее значение должно быть корректным.
- Флаги поддерживают Boolean и явные пары True/False, Yes/No, да/нет, 1/0.
  Остальные значения отвергаются без молчаливого False.
- Сохраненный display-статус проверяется через CResultStatusPolicy. Reader
  не создает инженерный статус из поврежденного текста и не анализирует
  ResultComment для выбора статуса.
- Диагностика указывает поле, фактический лист/адрес и действие восстановления.
  Value и Unit определяются по шапке; перестановка колонок и перенос якоря
  не оставляют жестко заданного адреса.
- Неуспешное чтение сохраняет atomic lifecycle: Count/annotations/selection/
  plane сбрасываются, тот же объект успешно читает восстановленный snapshot.
- Обычный import-preview не потребляет plane/status/extension/stress units.
  Missing-state reader продолжает проверять уже записанные данные выбранного
  LC, как до данного среза. Другие LC и другие StateType не потребляются.
- CUnitSystem.OutputStressToInternal получил optional savedUnit по аналогии
  с существующими snapshot conversions. Без аргумента расчет и порядок
  арифметических операций остаются прежними; новых таблиц коэффициентов нет.

Промежуточный v213 directed: **449/0**. Расширенный v214: **471/0**, но
соседний lifecycle-набор выявил **3749/180**: новая правка слишком широко
отключила проверку сохраненной плоскости в missing-state режиме. Это регрессия
новой реализации, не исходный дефект. Она устранена; прежние numerical
expected/tolerance и lifecycle assertions не ослаблялись.

## Проверки v215

- General Plot + новые metadata tests: **1880/0**; metadata subset **471/0**.
  Основной счетчик metadata содержит 66 перечисленных unit/error вариантов;
  дополнительно проверены 10 вариантов Boolean, семь display-статусов,
  пять stress units, column reorder, явная Unit, другой LC и inactive preview.
  Один synthetic geometry fixture, equilibriumCases=0, noSolve assertion.
- Reader lifecycle: **3929/0**; geometry snapshot: **1158/0**;
  profile presentation: **681/0**. Все source unchanged=True.
- Read-only structure: **27/27**; Config formatting: **1003/0**,
  1002 адреса, deviations=0, без ApplyAlignments.
- Palette/save-reopen: **355/0**. Results before/after SHA
  `E05C0BF950A1F5DA36E0BF42A1B9A7C7C68FBA4EA8006FC9B4F2470AC42B6C96`,
  styles SHA `1EABA40A72D5A28392F11FC5D38F0D2B1E9B264FDFAD9F5514ACB9EE1A1E1132`.
  Это равенство до/после reopen, не сравнение с другим запуском.
- Frozen book: `RC_Section_NDM_snapshot_metadata_positive_v215.xlsm`, SHA
  `0606E721B2931D4323C83A05EEA12B00048451A1E81CFBCD48DED9B2D8125E4B`.
- Actual read-only VBE export: 111 компонентов. Versioned export
  `VBA_All_Code_snapshot_metadata_v215_2026-10-04.txt`, SHA
  `E14DA340A6602A0B64E0776A8FF74E2C8D72F9E87A6A9B698ADED48D23DA4755`.
  Для этого промежуточного экспорта source equality: **106/106**, failed=0.
- Census: 106 modules/4412 methods/1700 guard candidates; semantic acceptance
  остается Pending. Новых классов нет, новых нормативных величин нет.

Финальный v216 отличается от v215 только уточнением комментария UnitOrDefault
о текущем контракте metadata, без исторической ссылки. General+metadata
повторно **1880/0**, source unchanged=True. Book SHA
`FD5B54048C18E5E13601F4E364C7DA1C8D8F60F9ED1E356C414E9A1B9835E91C`.
Его фактический export `VBA_All_Code_snapshot_metadata_v216_2026-10-04.txt`
содержит 111 компонентов, SHA
`0156B1644A3ED6BAA2B5A20165F9FCCE5187854D5766D73317AC55AC7645CFE5`;
именно эти байты записаны в canonical VBA_All_Code.txt.
Source/export **106/106**, failed=0. Structure **27/27**, palette/reopen
**355/0**, unit/sign consumers **44/0**, imported snapshot после смены units
**228/0**. Results v216 before/after SHA
`4229D60E3CDE8B06B0568364C1C74FB748E53F1378ACCE794DD62BF6EF73AA98`,
status styles SHA прежний; все source unchanged=True.

Один запуск Validate завершился ошибкой форматирующего PowerShell pipeline
после записи успешных 27 checks. Повтор с корректным выводом завершен exit=0;
это не дефект Excel/VBA и не ослабление проверок.

Полный Off/On не повторялся по разрешению пользователя: solver/search и
обычные численные conversions не менялись. Последний полный numerical gate
остается v206, **82221/0 Off**, **82232/0 On**; финальный release gate обязателен.
ТЗ и main output сохранили исходные SHA, пользовательский Excel PID 23476
не использовался и не закрывался.

## Открытый Объем

Общий Audit03 не завершен. Этот срез не заявляет приемку всех metadata
resolver-ов, ошибок заголовков, CONTOUR_ARC text parsing или экстремальной
арифметики Double. AutoUpdate, полный Config range/pairwise, semantic guard
review и финальная выпускная матрица остаются открытыми. Реальная запись DWG
и пиксельная приемка Excel данным набором не выполнялись.
