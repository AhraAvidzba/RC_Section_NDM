# Audit03: Общие Настройки И Геометрия

## Объем

Девять actual Config полей: General.DiagramExtension,
Calculation.ZeroMomentPerDepth, Geometry.Source/Type, Mesh.StepX/StepY,
Mesh.BoundarySubdivisions, Load.ReferenceOffsetX/Y. Новых классов нет.
Отдельный standard test module не дублирует production builders/solver.

## Подтвержденные Дефекты И Исправления

- F02/K01: полная книга без обоих Extension keys получала скрытое Yes.
  Теперь это ошибка ввода. Узкая миграция прежнего ключа и приоритет
  канонического сохранены; неполный автономный LoadFromRange не приравнивается
  к полной книге. Material provider использует обязательный boolean reader.
- F02/K01: blank/TODO/missing у фильтра малых моментов и смещений нагрузки
  давали default. Поля теперь читаются обязательно; отрицательный порог
  сохраняет знак при пересчете и отклоняется владельцем CMomentZeroFilter.
  Batch добавляет текущий адрес по машинному номеру ошибки владельца,
  не назначает статус по тексту. Сам фильтр не знает Excel.
- F05/K01: нулевой/отрицательный шаг сетки не давал текущего адреса и действия.
  Excel-registry выполняет адресный preflight и нормализацию INPUT длины;
  автономный mesh builder сохраняет независимую внутреннюю защиту.
- F07/T03: при неполном последнем шаге mesh builder отбрасывал базовую ячейку
  по ее центру до рассмотрения внутренних подъячеек. Удален этот предварительный
  отбор; ContainsPoint проверяет центры фактических элементов. Для круга D=300,
  шага 55 и subdivisions=4 прежняя площадь 65415.625 мм2, независимый эталон
  квадратуры 71465.625 мм2: 378 точек шага 13.75 мм внутри x²+y²≤150².
  Прежний coarse режим subdivisions=1: 22 точки, 66550 мм2, без изменения.

Нормативные формулы, пределы материалов, solver tolerances и исторические
expected values не изменены. Изменение состава сетки в неполных крайних
ячейках требует регрессии геометрии, batch и итоговой нагрузочной матрицы.

## Runtime Evidence

- v258: новый test setup не скомпилировался из-за имени circle в объявлении;
  actual VBE выделил строку 23. Assertions не начались: NotRun. Закрыт только
  подтвержденный owned Excel 8624, пользовательский 23476 не затронут.
- v259: 753/89; первый setup дополнительно не передал profile catalog
  отрицательному batch-consumer. Этот отказ не объявляется дефектом ядра.
- v260: 754/90 после исправления profile fixture и добавления независимого
  численного эталона сетки. Production без исправлений; helper только открывает
  существующий private Excel-adapter точки нагрузки без изменения поведения.
- v261: тот же направленный набор на исправленном production: 844/0.
- v262: усиленный frozen набор, 1698/90; 164 consumer-сценария.
- v263: окончательный усиленный набор, 1788/0; 164 consumer-сценария,
  source unchanged=True. Реальные единицы, модели, НДС, comments и writers.

Оба положения rngSystemSettings проверены: обычное и CH800. 88 ошибочных
значений, 18 missing keys и 18 recoveries проходят реальные consumers.
Актуальные адреса смещений после переноса: CI814/CI815, а не B17/B18.
Положительные/отрицательные журналы сохраняются без редактирования.

## Поведение

- Все четыре Geometry.Type строят настоящие бетон, арматуру и annotations.
- Оба шага меняют фактические элементы; subdivisions меняет граничную сетку.
  Эквивалентные INPUT мм/м сохраняют каждый элемент и стержень модели.
- AutoCAD source использует настоящий импорт Region fixtures, Results writer
  и readback. Generated поля не читаются и не влияют на snapshot; смена INPUT
  Length не масштабирует уже импортированное сечение. Реальный DWG: NotRun,
  отдельная диагностика среды в NDM_Audit03_AutoCAD_Environment_2026-10-04.md.
- Фильтр малых моментов проверен ниже/на/выше порога, при 0 и в kN*m/m.
- X/Y offsets проверены обоих знаков по равновесным моментам N*offset,
  эквивалентной плоскости и неактивности при N=0 с тем же solve count.
- Extension Yes/No проверен в текущем Strength и Cracked State. При перегрузке
  -10000000 Н Yes дает найденное вспомогательное НДС и физический FAIL,
  No дает действительный numerical failure. Физическая Capacity и нормальная
  плоскость не увеличиваются от Yes; writer/subtree comments проверены.

## Справка И Открытые Gates

Saved Help v263: failed=0, 2042 строки, 140 links, 118 Shapes. Все 760 inputs,
их Formula/Value и validation сохранены после update/reopen. Это проверка
содержимого, не pixel/normative acceptance.

Дополнительное чтение сохраненного XML справки: семь контрактов присутствуют,
failed=0 (`general_geometry_saved_help_v263b.log`). Первый lookup v263 дал
6/1: probe искал несуществующую фразу «Смещение обязательно задается числом»;
реальный каталог и книга содержат «Оба смещения обязательны: допускаются ...».
Исправлен только lookup следующей проверки, книга и текст не переписывались;
первый журнал сохранен. Новый текст не объявляется нормативной трассировкой.

Полный Batch из eight-suite Off завершен: 25607/0, elapsed=572.265625 s.
Его исходная секция извлечена без редактирования assertions в
`general_geometry_batch_from_full_v263.log`. Все 2264 общих numerical
actual-values с completed Stability v257 совпали точно; missing=0,
differences=0 (`general_geometry_batch_numbers_v263.json`). Диагностическое
время не объявляется benchmark-ускорением по одному запуску.

Geometry regression v263: 536/0; actual source/export equality 107/107,
112 native components; structure 27/27, source unchanged=True. Полный
промежуточный eight-suite Off v263 завершен с exit=1: Geometry 536/0,
Material 2856/0, Section 1031/0, Capacity 3074/0, Crack 1532/0,
Batch 25607/0, WorkbookUI 64151/12, RegressionBaseline 39/0.
Все восемь suites завершились; timeout не было, source unchanged=True.
Отрицательный журнал не редактируется и не считается зеленым gate.
Поадресная приемка этих девяти
полей пока не присоединена к registry. Full-range/extreme/pairwise, final eight-suite On/Off, широкая
матрица и benchmark, clean build/update/UI и итоговый self-audit обязательны.
Основная пользовательская output-книга и ТЗ не заменяются этим срезом.

Native book SHA-256:
`3E5E46D429B78FDABDA8C00F02FC1CEE2764A56BFDE4B92C5B0B75BC93789939`.
Actual CodeModule export SHA-256:
`B4F35A8DC7130E23E19E170BFD6E8E7C0097C9E8DB9C12F296C796021D1B01CF`.
Экспорт: `docs/regression/Audit03/VBA_All_Code_general_geometry_v263_2026-10-04.txt`.
Это код реальной изолированной книги после всех source/test изменений,
не сборка текста из файлов src/tests. Поадресная трассировка для следующего
registry gate: `NDM_Audit03_General_Geometry_Config_Evidence.json`.

## Двенадцать Отказов Полного Прогона

- Шесть unit-caption assertions: Capacity InitialLambda/MaxLambda/
  ToleranceLambda/MaxRetries/BaseLoadSteps/SolverMaxIterations. В сохраненной
  v263-книге колонка «Ед.» действительно содержала числа. Устаревший
  Refresh-Workbook.ps1 писал семиколоночную схему поверх текущей таблицы
  и повторно назначал пользовательские значения. Entry point теперь
  использует существующий source-only Refresh-VbaModules.ps1, без второго
  importer-а и без записи defaults. Явный update справки восстанавливает
  только unit captions из общего каталога; сохранность 760 inputs обязательна.
- Четыре validProfile assertions AutoPlot: stability-only fixture наследовал
  пользовательский знак +N, хотя тесту требовалось сжатие. Fixture теперь
  явно задает Compression только в собственной копии книги.
- Один AutoPlot noSolve: batch сбрасывает общий счетчик в начале запуска,
  поэтому прежнее сравнение с накопленным счетчиком других tests некорректно.
  Проверяется ноль у каждого действительного stability-only запуска.
- Один settingsTable runtime: файловая копия промежуточного test-журнала
  получила Permission denied и прервала assertions. Отказ optional I/O теперь
  сохраняется как warning в возвращаемом отчете; assertions продолжаются.
  Добавлена отдельная проверка с блокировкой файла и восстановлением записи.

Изолированная v264-книга обновлена и прошла направленные проверки:
General/Geometry 1788/0 (164 consumer cases), settings-table guards 583/0,
AutoPlot 86/0, structure 27/27 и actual source/export 107/107.
Все runtime-журналы завершены exit=0, source unchanged=True.
Saved help failed=0; сохранены 760 inputs после save/reopen.
Сравнение фактического Config XML v263 -> v264 подтвердило ноль изменений
всех 760 inputs, их формул, alignment, merge и validation. В Config поменялись
только 15 unit captions, включая девять трещинных, не охваченных прежним
UI-test. Все 101 системная подпись единиц в сохраненной книге соответствует
каталогу, включая динамические формулы размерных величин.

Доказательства: general_geometry_refresh_input_preservation_v264.json,
general_geometry_saved_units_v264.json, help/table_guard/auto_plot/positive/
source_contracts/structure v264. Эти directed gates подтверждают конкретные
исправления, но не подменяют новый полный eight-suite gate.
До нового зеленого полного gate registry остается v257: 29 editable полей
не приняты. Численные expected/tolerance и итог отрицательного прогона
не меняются. Пользовательская книга и ТЗ остаются нетронутыми.

В двух обзорных абзацах guide уточнено включение границы M_tol:
«модуль не больше M_tol», в согласии с |M| <= M_tol в коде и подробном
разделе настройки. Это не изменение численного условия.

## Текущий Native Snapshot v264

- Книга: RC_Section_NDM_general_geometry_final_v264.xlsm.
  SHA-256: 60569C86E31F230B0C21B1DC78165F07EC817AC029C77F9E4E38B685EF71B191.
- Native CodeModule export: VBA_All_Code_general_geometry_v264_2026-10-04.txt.
  SHA-256: 995A6DD3416333E2DD3B0E954FD47689C03F6A6F7B3DDB82720D072FF308CF54.
  Текущий VBA_All_Code.txt содержит тот же экспорт реальной книги.
- 112 native components, 83 production classes и прежние три test classes.
  Новые классы не добавлялись. Основная output-книга и ТЗ сохраняют baseline
  hashes, пользовательский Excel не использовался.
