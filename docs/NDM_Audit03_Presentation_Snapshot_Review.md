# Audit03: Оформление И Сохраненный Snapshot

## Граница Среза v193-v199

Предыдущий принятый checkpoint: `4535c99cf4bb2e12056c8cb4479f294b76e7969f`.
Меняются существующие Excel-владельцы и существующие тестовые модули.
Новых классов, физических критериев, Search или численных допусков нет.
Основная пользовательская книга не публикуется до выпускной приемки Audit03.

## Подтвержденные Исправления

- `CSystemSettingsReader` сохраняет применимые пустые поля и их фактический
  адрес; ошибки формул оформления откладываются до активного обращения.
  Неприменимые колонки аннотаций определяются по схеме, а не по значению `-`.
  Тип `rngPlotAnnotationSettings` сохраняется по имени диапазона: удаление
  строки Placement не превращает всю таблицу в таблицу другого формата.
- RGB проверяется в общем reader-е: три целых компонента 0..255. Черный 0
  не является признаком отсутствия цвета, дробные каналы не округляются.
- Layout требует активные параметры, не читает выключенную группу и
  сбрасывает элементы при повторном Initialize. Экранные ограничения
  шрифта/отступов сохранены; неиспользуемый collision-код удален после call-map.
- Plotter проверяет активное оформление до очистки прежней схемы. Цвета,
  размер стрелок и высота подписей осей не заменяются скрытым default.
  В Common цвета бетона не читаются: применяются цвета общей шкалы арматуры.
- `CUnitSystem` принимает явную единицу сохраненного snapshot через optional
  аргумент существующих методов. Reader/plotter/layout больше не хранят свои
  факторные таблицы длин/площадей/кривизн. Порядок деления кривизны сохранен.
  Текущий Config не применяется повторно к уже сохраненным числам Results.
- Повреждение существующей таблицы аннотаций больше не скрывается как
  отсутствие аннотаций или нулевые координаты. Нормали, координаты и текст
  проверяются с адресом Results. Пустой необязательный блок допустим.
  Текст подписи сохраняется, включая окружающие пробелы.
- GeometryStatusColumn требует GeometryInterpretationStatus либо ShapeType;
  отсутствие обоих заголовков дает управляемую причину, а не индекс 0/raw VBA 9.

## Runtime Evidence

| Артефакт В `docs/regression/Audit03` | Результат И Смысл |
| --- | --- |
| `presentation_negative_v193_2026-10-04.txt` | 280/120, 98 counted cases; black/default/active-inactive/reset defects воспроизведены на прежнем production. |
| `presentation_v194_2026-10-04.txt` | Незавершенный compile-run из-за двух оставшихся имен переименованного helper-а; не NumFail и не PASS. Закрыт только созданный test Excel. |
| `presentation_snapshot_negative_v195_2026-10-04.txt` | 284/122; snapshot-часть имела ошибочный test scope CurrentRegion. Не является доказательством snapshot-дефекта. |
| `presentation_snapshot_negative_v195b_2026-10-04.txt` | 293/130, 105 counted cases; scope исправлен по именованному якорю/схеме. Поврежденные координаты/текст и пропавший заголовок воспроизведены. |
| `presentation_v196_2026-10-04.txt` | 677/2; оба failures относятся к утрате Placement и ошибочному распознаванию таблицы. Исправлена маршрутизация named-table, oracle не ослаблен. |
| `presentation_v197_2026-10-04.txt` | 679/0, 105 counted cases; active input, inactive CVErr, recovery, moved Name, Chart black, snapshot, no solve/unchanged Results. |
| `presentation_profiles_v198_2026-10-04.txt` | 8740/0, 276 counted cases; PR1-PR4, пять states, Stress/Strain, значения всех элементов по плоскости/материалу, фактический текст легенды при precision 0..10. |
| `help_presentation_v199_2026-10-04.txt` | failed=0; 1988 строк, 140 ссылок, 118 shapes. Все 760 input-ячейки, формулы, validation и PrintArea сохранены после save/reopen. |
| `code_census_presentation_v199_2026-10-04_*` | 105 модулей, 4387 методов, 1618 guard-кандидатов; это индекс, semantic acceptance остается Pending. |

Все завершенные watchdog-directed прогоны сохранили SHA исходной книги.
`presentation_black_v197_2026-10-04.png` прочитан через view_image: график
не пустой, черные размерные/выносные линии и подпись арматуры видны,
геометрия и легенды сохраняются. Это не визуальная приемка всех форм/настроек.

## Замороженный Gate v199

Тестовая книга: `RC_Section_NDM_presentation_all_v199.xlsm`.
SHA-256: `C91FB920944A5977474C9947D431E2A4270CAA5656A41D5486117A85E54934B5`.
Она содержит также тест сохранения пробелов подписи и optional/invalid/recovery
описания профиля. Для физической геометрии восстановлена точная GEOM_PI.
Во время full Off/On этот source/test-срез не изменяется.

Full Off v199 завершен: восемь suites, 77134/0; UI 55231/0, profile
8784/0 и presentation 681/0. 13641 общий numeric actual-value с Off v192
совпал точно, missing=0/differences=0. Первый On v199 не начал ни одной suite:
COM factory Excel.Application вернула CO_E_SERVER_EXEC_FAILURE (80080005).
Исходная книга не изменилась, own Excel не остался; это failure запуска среды,
не numerical outcome и не PASS. Первый structure attempt получил тот же отказ
COM factory. Оба отрицательных артефакта сохранены. Protected PID 23476 не
затрагивается; повтор выполнялся в отдельном собственном Excel.

Повтор On v199b завершил восемь suites: 77145/0. Все 13646 общих numeric
actual-values с On v192 совпали точно, missing=0/differences=0. Source unchanged
подтвержден обоими watchdog-прогонами. Дополнительные завершенные gates:

- `validate_presentation_v199b_2026-10-04.json`: 27/27, книга не изменена.
- `formatting_presentation_v199_2026-10-04.txt`: 1003/0, 1002 адреса,
  deviations=0, без исправляющего ApplyAlignments.
- `palette_reopen_presentation_v199_2026-10-04.txt`: 355/0; Results и фактические
  статусные стили одинаковы до/после сохранения и повторного открытия.
  Results SHA `9F6AB303F8493E8761F2694159CB3CC2EB9B0248E466B3EC9221A14018704414`;
  style SHA `1EABA40A72D5A28392F11FC5D38F0D2B1E9B264FDFAD9F5514ACB9EE1A1E1132`.
- `source_contracts_presentation_v199_2026-10-04.txt`: 105/105, failed=0.
  Read-only VBE export содержит 110 фактических components; 105 source/test
  modules совпадают с исходниками после нормализации только представления export.
  Export SHA `DFF5001EDDE14B983E70F491E91C1BE9189EC3F8DD172502F2D3627FCFFE4257`.

После этих gates присоединяется evidence ровно 20 profile presentation/description
полей. Реестр v199: 666/1064 active-reviewed, 94 editable pending из 760.
Это не полная приемка диапазонов/взаимодействий и не завершение Config или Audit03.
21 annotation-поле этим merge не принимается: соответствующее активное поведение
проверено направленно, но полная поадресная трассировка еще не завершена.

## Оставшийся Объем

- Все допустимые ArrowType/ArrowSize, независимые изменения всех толщин,
  отступов/зазоров и полный visual-range/pairwise остаются отдельной приемкой.
- General Plot active-inactive/error-поведение, geometry-only unused style,
  поздний LC comment lookup, reset флага CrackWidth в переиспользованном reader-е.
- Содержательный разбор остальных dead private-групп Plotter и численного
  парсинга contour arc, настоящий bucket-render/rounding reproducer.
- Общие оставшиеся Config поля, guard/lifecycle/arithmetic-кандидаты,
  выпускные матрицы/benchmarks и self-audit остаются в текущей цели.

Защищенный пользовательский Excel PID 23476 не закрывается и не используется.
Main output SHA сохраняется `AAF5D4FAF06197F01966463DE85BA813EB4216B455AA8504B2DE4F0BA717C012`.
