# Audit03: Контракт Дуг Results

Дата: 2026-10-04. Исходный checkpoint: `4e25d63`.

## Подтвержденный Дефект

Два независимых ParseInvariantDouble в Excel-схеме и AutoCAD export
переходили к Val после ошибки CDbl. Поэтому поврежденное Text вроде
`1.25garbage` принималось по числовому префиксу, пустота превращалась
в нулевую дугу, а большой угол доходил до CLng после очистки схемы.
Reader сам не проверял численный смысл Text у CONTOUR_ARC.

Negative v220: **71/158**, 44 варианта, отдельная синтетическая книга,
equilibriumCases=0. Корректные signed quarter-circle/decimal/E-notation
прошли прежнюю отрисовку; ошибочные строки и ранняя сохранность Chart
не прошли. Это направленный контрпример, не нагрузочный расчет.

## Исправление

- Общий ReadContourArcSweep в существующем modGeometryTypes читает все
  значение, проверяет десятичную запись и преобразует число в локали VBA.
  Нет Val, обращения к Excel, новой таблицы единиц или нового класса.
- Text допускает точку/запятую, знак и E-нотацию. Геометрическая дуга
  меньше полного оборота; полная окружность представляется CONTOUR_CIRCLE.
  Ноль сохраняет существующее поведение вырожденного отрезка. Это контракт
  представления контура, не новый инженерный критерий или требование СП.
- CSectionPlotDataReader проверяет угол до выдачи снимка, хранит готовый
  Double и сбрасывает его вместе с массивами при неуспешном чтении.
- CSectionPlotter получает готовый AnnotationSweepAngle, не разбирает Text.
  Некорректный угол не доходит до очистки Chart/счетчика сегментов.
- AutoCAD preflight готовит углы до подключения и создания объектов DWG.
  DrawParametricSectionContour использует подготовленные значения; два
  прежних parser-а удалены. Необязательное отсутствие contour Name допустимо.
- Ошибка на русском содержит Text, радианы, актуальный Sheet!Address и
  действие для восстановления snapshot; перенос именованного якоря учитывается.
- Справка AutoCAD.Export.ContourEnabled объясняет контракт Text/радианов
  и ранний отказ. Solver/Search/material criteria/expected/tolerance не менялись.

## Ошибки Собственного Теста

v221 не выполнил assertions: после переименования тестовой переменной
оставлен Next со старым именем. Watchdog остановил созданный им Excel
через 240 секунд; это compile error нового теста, не production finding.
Исправлено без изменения production или исторических expectations.

Первый общий вызов v222 был задан без имени модуля при двух существующих
одноименных entrypoint-ах. Excel.Run не выполнил macro; это NotRun, не
провал assertions и не PASS. Frozen v223 использует проверенную ранее
модульно квалифицированную запись modTestPlotConfig.RunAudit03GeneralPlotTests.
Оба исходных журнала сохранены отдельно, не перезаписаны зелеными результатами.

## Приемка Frozen v223

- Directed v222: **393/0**, 56 основных вариантов и отдельные locale checks.
  Две позиции anchor, ранний отказ настоящего preview-entrypoint,
  reader reset/recovery, подготовка AutoCAD sweep, точные Results Value2/type,
  quarter-circle 18 хорд, +/- угол, нулевой отрезок, noSolve.
- Общий presentation/metadata/AutoUpdate/arc набор: **2355/0**.
- Три режима reader lifecycle: **3929/0**. Geometry snapshot: **1158/0**.
- Unit/sign consumers и сохраненный export snapshot: **44/0**.
- Structure read-only: **27/27**. Config formatting: **1003/0**, 1002 адреса,
  deviations=0, проверка не исправляла оформление.
- Help: **2005** строк, **140** ссылок, **118** shapes, failed=0;
  input values/formulas/validation и PrintArea сохранены.
- Palette/save-reopen: **355/0**, Results before/after SHA
  `294BBE8FFB8EB1DEA5E09377A662DDB02388F07C8E8114F02353E000B5454F40`,
  styles SHA `1EABA40A72D5A28392F11FC5D38F0D2B1E9B264FDFAD9F5514ACB9EE1A1E1132`.
- Frozen book RC_Section_NDM_contour_arc_positive_v223.xlsm, SHA
  `2A211839469722EE697E743941E7DF3EAFD938041E7D65954DB776101ED17A31`.
- Actual read-only VBE export: **111** компонентов, файл
  VBA_All_Code_contour_arc_v223_2026-10-04.txt, SHA
  `8F270447E794FDFA5FBCE4C528A746335708509272F613A6792349A649CC501D`.
  Canonical VBA_All_Code.txt содержит те же фактические байты.
- Source/export equality: **106/106**, failed=0; удаленные классы и их
  consumers не возвращены, состав 83 production + 3 test classes сохранен.
- Census: **106** модулей/**4419** методов/**1713** guard candidates;
  это индекс проверки, semantic acceptance Pending.

Все завершенные watchdog source unchanged=True. Main output и исходное ТЗ
сохранили baseline SHA; пользовательский Excel PID 23476 не изменялся.
Полный numerical Off/On не повторялся: последнее доказательство v206
82221/0 Off, 82232/0 On сохраняется, выпускной полный gate обязателен.

## Оставшиеся Границы

Этот срез не подтверждает фактическое создание DWG, пиксельную приемку,
все metadata resolver-ы или крайние координаты Plot. Полный Config/pairwise,
численная release matrix, Excel guard, extreme arithmetic и общий semantic
review остаются в текущей цели. Общий Audit03 не завершен.
