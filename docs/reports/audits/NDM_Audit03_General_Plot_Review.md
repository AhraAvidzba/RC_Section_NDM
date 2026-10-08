# Audit03: Общая Excel-Схема И Аннотации

## Scope

Presentation-срез после `40626a67bb3e162db8dcb2efa901426e90dc841d`.
Решатели, Search, физические критерии, material spec, численные expected и
tolerance не изменяются. Новые классы не добавляются. Тесты используют
отдельные временные книги и готовые синтетические Results, без поиска НДС.

## Подтвержденные Дефекты

- Активные переключатели Contour/NeutralLine/LoadApplicationPoint/Labels/
  Legend и режим главных осей читались после удаления прежней схемы.
  Отсутствующий ключ мог выбирать default; ошибочная формула могла очистить
  правильный старый рисунок до возврата понятного сообщения.
- Пустой PrincipalAxesMode выбирал Transformed, а неизвестный режим не
  называл фактическую ячейку. Неиспользуемый альтернативный ключ
  PrincipalAxesEnabled создавал дополнительную fallback-ветвь.
- Пустой/TODO ResultLabelSpacing выбирал default, ноль/отрицательное число
  назначали автоматический шаг. Маленький положительный шаг переполнял Long.
- `CLng(value + 0.5)` применял банковское округление к уже смещенному целому
  номеру bucket. На 5001-точечном fixture отсутствовали восемь нечетных
  ступеней из 17, хотя исходные Results содержали каждую величину точно.
- Preview-entrypoint не проверял Plot.Enabled и читал оформление даже при No.
- Пустой Plot.LoadCase выбирал Worst вместо обязательного явного выбора.

## Изменения Владельцев

`CSectionPlotter` проверяет активное оформление до первого изменения Chart.
Один проход Draw сохраняет собственные flags; выключенные result-потребители
в geometry-only режиме не читаются. PrincipalAxesMode сохраняет поддержанные
текстовые обозначения, но не пустой default/альтернативный ключ.

Шаг численных подписей положительный и задан в мм. Целые индексы пространственных
ячеек хранятся в Double, а не искусственно ограничиваются Long. Подпись выбирается
по ячейке сетки, что не является гарантией минимального расстояния между любой
парой подписей. Округление неотрицательного bucket выполняет Int перед CLng.

`modWorkbookCalculation` одинаково проверяет Plot.Enabled перед обычным
обновлением и import-preview. `CSectionPlotDataReader` требует Plot.LoadCase
только у потребителей выбранного LC; обычный import-preview его не читает.
SettingsCatalog/справка описывают эти текущие контракты.

## Negative Evidence

- v207: 0/2, ошибка нового fixture: пятиколоночная таблица аннотаций ошибочно
  читалась как scalar. Исправлен только test setup.
- v207b: 775/308. Усиленные проверки выявили также неверные test oracle:
  диагностическое сообщение использует `лист ..., ячейка ...`, не `Sheet!A1`;
  TextGap задает смещение центра текста, не края. Исправлены только oracle.
- v207c: 1000/123 на прежнем production. Подтверждены очистка Chart, отсутствие
  обязательных ключей, spacing и восемь потерянных цветовых уровней.
- v208: 1121/2 после первой production-правки. Остались два test-only падения:
  fixture не задавал StressPrecision для проверяемой точной подписи.
- v209: 1057/174. Новый fixture записывал RGB `140,140,140` в General-ячейку,
  и Excel интерпретировал строку как число с разделителями тысяч. Контрольные
  таблицы fixture получили текстовый формат, исходные Config не менялись.
- v209b: 1147/168 на прежнем production; дополнительно подтверждены Enabled
  и обязательный LoadCase. Журналы отрицательных исходов сохраняются полностью.

Первый запуск watchdog из PowerShell 7 не нашел ожидаемый `powershell.exe`
в PSHOME. Повтор через Windows PowerShell выполнен; это ошибка запуска helper,
не compile/runtime defect книги.

## Gates

- Positive v210 и окончательный v211: 1409/0. 123 presentation variants,
  два geometry fixtures, 206 annotation variants и один независимый layout.
  Поиск равновесия не выполняется; проверены фактические Chart.Shapes/Series,
  15 комбинаций стрелок, численные mutation-oracle, recovery и текущие адреса.
- Directed regressions v211: read lifecycle 3929/0, geometry snapshot 1158/0,
  profile presentation 681/0. Все watchdog exit=0, source unchanged=True.
- Новые проверки находятся в стандартном `modTestPlotConfig`; общий UI suite
  агрегирует его отчет и счетчики. Новых `.cls` нет. Snapshot invariant
  проверяет точное Value2 и тип, а не допуск приблизительной эквивалентности.
- Help: 2002 строки, 140 links, 118 shapes, input/formula/validation/PrintArea
  сохранены. При v210 -> v211 совпали 2650 непустых ячеек справки и merges.
  Это проверка двух обновленных книг, не отдельная clean-сборка.
- Read-only structure: 27/27. Formatting: 1003/0, 1002 адреса, deviations=0,
  без ApplyAlignments. Source hash до/после проверки не изменился.
- Palette/save-reopen: 355/0. Results SHA
  `294BBE8FFB8EB1DEA5E09377A662DDB02388F07C8E8114F02353E000B5454F40`,
  status styles SHA
  `1EABA40A72D5A28392F11FC5D38F0D2B1E9B264FDFAD9F5514ACB9EE1A1E1132`;
  оба совпали после повторного открытия.
- Actual read-only VBE export: 111 компонентов, source/export 106/106,
  failed=0. Export
  `VBA_All_Code_general_plot_v211_2026-10-04.txt` SHA
  `E56FF811B396F8FB37C5920B81444D2E79E8EC98303515CE30955B330E8DFC67`.
  Canonical `VBA_All_Code.txt` содержит те же фактические байты.
- Frozen book `RC_Section_NDM_general_plot_positive_v211.xlsm` SHA
  `6947B2AB2061B9EACA9C84F69DCA64ACB3ADD649EB908CE694F4EF8EDF876CAE`.
- Census: 106 modules, 4407 methods, 1675 guard candidates; semantic review
  Pending. 83 production + 3 test classes сохранены, template suspects=0.

Полный On/Off для этого presentation-среза не повторялся по разрешению
пользователя. Последний полный numerical gate относится к v206:
82221/0 Off и 82232/0 On. Он не переименовывается в прогон v211;
окончательный полный release gate остается обязательным.

Основная output-книга и пользовательское ТЗ сохранили исходные SHA из
Progress. Пользовательский Excel PID 23476 не использовался и не закрывался.
Локальный checkpoint включает только собственный source/tests/help/evidence,
без push, destructive Git или публикации незавершенного Audit03 в main output.

## Ограничения

Это не пиксельная приемка Excel или реальная запись DWG. Численные значения
схемы проверяются по настоящим COM Chart.Shapes/Series и готовому layout.
Полный диапазон Config/pairwise и крайние representability-сценарии Double
для пространственной сетки еще не получают PASS. Общий Audit03 не завершен;
настройка AutoUpdateAfterCalculation, snapshot Unit/логические metadata,
Excel guard и прочие F07-кандидаты остаются отдельным незавершенным объемом.
