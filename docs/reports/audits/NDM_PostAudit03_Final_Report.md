# Post-Audit03: Итог Исправлений И Снимка Контуров

Дата: 07.10.2026. Основание: `NDM_PostAudit03_Fixes_And_Contours.md`.
Статус выпуска: **шесть пунктов реализованы и проверены; книга и фактический VBA-экспорт опубликованы**.

Отчет относится к шести локальным изменениям после Audit03/СП35. Он не является
независимой нормативной приемкой программы и не закрывает вопросы, оставленные
в `OpenNormativeQuestions.md`.

## Baseline И Границы

- Исходный Git: `872d481117fca33bd0fe60b953ffb9587266b5df`, ветка
  `codex/material-diagram-architecture`.
- До начала изменений сохранены книга, фактический VBA TXT, пользовательский
  execution report, формулы/значения Config/Results и ширины столбцов.
- 120 исходных VBA-компонентов, из них 93 класса: количество не увеличено.
- Новые VBA-классы, формулы СП35/СП63, диаграммы, численные допуски и отдельные
  новые расчетные ветки в рамках этого задания не вводились.
- Старые сторонние/untracked материалы и пользовательский dirty-report не
  очищались, не откатывались и не включаются в коммит наших изменений.

## Пункт, Владелец, Доказательство

Все пути логов ниже относительно `docs/regression/PostAudit03/`.

| Пункт | Владелец И Изменение | Тест/Лог | Итог |
|---|---|---|---|
| 1. Исходные LC неизменяемы | `CBatchSectionCalculator`: raw N/Mx/My/путь отдельно от `CSectionLoadState` рабочего запуска; Stability изменяет только рабочие усилия; кэш центра пересоздается | `IsolatedFinalDirected/modTestBatchCalculation.RunPostAudit03LifecycleTests.txt`: 301/0 | PASS |
| 2. Достоверный Formation | `CCrackFormationCalculator`/`CCrackFormationResult` подтверждают точку и факт трещины; отказ Post не уничтожает порог. `CCrackSummaryWriter` использует Formation и общую policy | `IsolatedFinalDirected/modTestBatchCalculation.RunPostAudit03FormationTests.txt`: 49/0; `StateDisplay/StateDisplay.txt`: 57/0 | PASS |
| 3. Однонаправленная подготовка СП35 | `CSP35CrackData` владеет beta и множителем радиуса; `CCrackWidthCalculator` сохраняет Rr/psi/раскрытие | Все 9 логов `FinalSP35/`: 2583/0 | PASS |
| 4. Пассивная готовая область | `CSectionGeometryQuery` один раз определяет роли; `CConcreteRegion` копирует границы/роли/диагностику без Query | `CloneIsolated/modTestGeometryQuery.RunGeometryQueryTests.txt`: 1084/0 | PASS |
| 5. Пассивная сериализация | `CSectionPropertiesCalculator` готовит A/I/габариты; `CSectionAnnotations` готовит импортные размеры; `CSectionStateResult` хранит фактические frozen-диаграммы и готовит временный массив элементов | `ContourSnapshotOwner/modTestBatchCalculation.RunPostAudit03SnapshotOwnerTests.txt`: 38/0; `FinalRepeatContours/modTestBatchCalculation.RunPostAudit03SnapshotRepeatTests.txt`: 5979/0; `SnapshotComparison/Comparison.json` | PASS |
| 6. Отдельный снимок контуров | `CNDMResultsWriter` сериализует; `ReadSavedSectionContours` читает только v1; узкий legacy-reader доступен только миграции. Build/layout/preview/clear/plot/export/help обновлены | `FinalRemainingDirected/modTestAutoCADContours.RunPostAudit03ContourSnapshotTests.txt`: 309/0; native CAD ниже | PASS |

Lifecycle проверен против свежего объекта: повторный Execute, Stability On/Off/On
по СП63/СП35, высокий/низкий порог моментов, перенос точки приложения туда-обратно,
смена модели/provider и неизменность ранее опубликованных result-ссылок.
Контрпример baseline сохранен: `LifecycleBaseline/`, 146 PASS / 108 FAIL.

Formation проверен реальными ячейками: аналитическое Ncrc до отказа Pre;
принятая точка с отказом Post; обычные Cracked/NotCracked; чистое сжатие;
граница поиска не становится NotCracked; физический FAIL Pre при нетипичном
порядке пределов материалов. Typed display дополнительно проверяет фактическую
meta/комментарии без локального whitelist и без повторного solve.

## Формат Контуров

Нижний ряд: ElementResults A (9), Geometry L (15), Contours AC (17),
Properties AV (6), MaterialDiagrams BD (11), Annotations BQ (22).
Якоря в одной строке, между блоками минимум два пустых столбца.
Все потребители используют фактические именованные якоря, не жесткие адреса.

`rngNDMSectionContours`: RunID v1, LoopID, SegmentID, Sequence, LoopRole,
SegmentType, StartX/StartY, EndX/EndY, CenterX/CenterY, Radius,
SweepAngle в радианах, SourceID, Comment, LengthUnit (SectionXY).

- Линия хранит концы; окружность - центр и радиус; дуга - начало, центр и
  знаковый угол. Конец/радиус дуги не дублируются.
- Роли Outer/Opening, порядок, ID и связь с RunID сохраняются.
- Несколько outer/opening, только opening и неизвестный outer допустимы.
- Пустая шапка v1 не доказывает отсутствие отверстий. Пустая/поврежденная
  шапка либо broken-name не вызывает скрытое чтение старого Geometry.
- Geometry содержит только элементы; расчетные области CRACK_REGION и
  CRACK_PROBE остаются LC-аннотациями, исходные контуры там не дублируются.
- Единицы принадлежат сохраненному снимку, а не текущему Config.

Миграция читает старый снимок до очистки, проверяет v1 на временном листе,
сдвигает только нижнюю полосу ячеек, сохраняет первые 15 механических полей
и удаляет прежние contour-строки/поля. Повтор безопасен. Проверены защитные
отказы для защищенного листа/структуры с понятной инструкцией и без изменения
старого снимка. Новый импорт и solve для миграции не вызываются.

309 проверок включают mm/cm/m, большие и отрицательные дуги, окружность,
несколько компонентов и отверстий, перенос якоря, актуальные адреса ошибок,
большой/малый/пустой набор, сохранение и повторное открытие, broken-name,
форматирование, пользовательские ячейки/ширины и идемпотентность.

Пользовательский кандидат: 16 сегментов перенесены без solve/import;
`ReleaseCandidate/Migration.json`. Чистая сборка:
`CleanFinal/Validation.json`, структурные проверки PASS.

## Числа И Производительность

`SnapshotComparison/Comparison.json`: **26569 фактических Excel Double**,
0 пропусков, **0 отличий** baseline/current; сравнение точное, без изменения
допусков. Исключен только служебный RunID. Включены физический успех, Off NumFail,
On сходящийся физический FAIL с расширением, характеристики и элементы.

| Сценарий | Baseline, с | Current С Подготовкой, с | Отдельная Подготовка, с |
|---|---:|---:|---:|
| 0 | 0.4296875 | 0.3828125 | 0.015625 |
| 1 | 0.5859375 | 0.20703125 | 0.0078125 |
| 2 | 0.17578125 | 0.1875 | 0.0078125 |
| 3 | 0.171875 | 0.16796875 | 0.0078125 |
| 4 | 0.15234375 | 0.1796875 | 0.00390625 |
| 5 | 0.1484375 | 0.16015625 | 0.01171875 |

1000 clone: frozen baseline 0.046875 с, current 0.0546875 с, оба 1084/0.
Это грубый VBA Timer и параллельно нагруженная машина; ускорение или замедление
статистически не заявляется. Статически повторный Query/classification из clone
удален. Все snapshot-сценарии подтверждают отсутствие дополнительного solve.

Постоянные массивы всех волокон всех НДС не добавлены. У состояния сохраняются
две небольшие фактические диаграммы: массивы точек требуют
`16 * (ConcretePointCount + SteelPointCount)` байт чисел на state, плюс обычные
объекты/метаданные. Модель общая с контролем ревизии; элементный Variant-массив
существует только на время упаковки одного состояния. Временное выделение
памяти не используется как заявление о гарантированном ускорении.

## Полные Наборы

Off: первый полный штатный прогон сохранен в `Full_Off.txt` и **не объявляется
успешным**. Geometry 668/0, Materials 4353/0, Solver 1492/0,
Capacity 3562/0, Crack 2158/0, RegressionBaseline 39/0.
Batch 29027/598, WorkbookInterface 65479/73 выявили исторический implicit-SP63
setup и адреса прежней 49-колоночной таблицы при действующей 72-колоночной.
Исправлены явный test setup и accessors, инженерные expected/tolerance сохранены.

Повторный полный Batch Off: `Full_Off_Batch_Current.txt`, 30893/4.
Четыре оставшихся отказа закрыты `FinalRemainingDirected/` (writer/reserve 190/0):
два старых адреса запасов и две проверки реально исправленной шапки Es/acrc.
Все расчетные поднаборы этого повторного Batch прошли без отказов.
Финальный source guard runner обнаружил собственное обновление кандидата:
runner остался exit 1, проверенная исходная копия с тем же SHA сохранена
в `FinalRepeatContours`. Это не выдавалось за успешный source guard.

Адресный Off UI: reader 67/0, INPUT units 1294/0, plot/contours 3347/0,
saved snapshot 267/0. Его source guard также не объявляется PASS после
технической блокировки кандидата updater-ом; provenance recovery в
`Off_UI_Provenance_Recovery.json` подтверждает идентичность замороженного исходника.
Полный On: первые пять suites 668/0, 4353/0, 1492/0, 3562/0, 2158/0;
Batch **30909/0**, 1764.62890625 с. Полный WorkbookInterface Off:
`Full_Off_WorkbookInterface_Final.txt`, **67869/0**, 5547.91796875 с,
exit 0 и SOURCE_UNCHANGED True. Явный setup одного из тестов оставил Yes
в временной fixture; override записан и изменения отброшены при закрытии.
Полный WorkbookInterface On: **67869/0**, 4580.31640625 с;
RegressionBaseline **39/0**. Все восемь On suites прошли, runner exit 0,
SOURCE_UNCHANGED True. Начальные FAIL-логи Off сохранены; исправленные
accessors/шапка закрыты адресно и проверены полным актуальным On, а не
подменены новым «успешным» названием исторического Off-лога.

Все 9 актуальных СП35 (`FinalSP35/`): Pipeline 34/0, Workbook 133/0,
Formula 27/0, LocalOpening 36/0, Preparation 405/0, NeighborSettings 18/0,
CurrentProjection 97/0, SavedHollow 1563/0, EndToEnd 270/0; итого **2583/0**.
Неделимость групп, t/L, RowTolerance, независимые opening и исходные численные
правила сохранены. После этих наборов менялось только оформление/защитный
preflight миграции и Unicode-представление двух тестовых строк, не математика СП35.

## Настоящий AutoCAD И Оформление

Native AutoCAD 2023: **82/0**, `NativeCADRawDispatch/NativeCAD.txt`,
`NativeCADRawDispatch/NativeCADIdentity.json` и DWG.
Проверены путь настоящего `AutoCAD 2023/acad.exe`, COM server и собственный PID;
SOFiPLUS и документы пользователя не использовались. Собственный экземпляр
закрыт runner-ом. Ранее неудачные activation/getter попытки не считаются PASS.

COM-fixture не подменяет native: отдельно 283/0 AutoCAD-contour fixtures и
1158/0 geometry snapshot в `ContourSnapshotOwner/`.

CleanFinal: **1583/0** Config presentation, `CleanFinal/Config_Presentation.txt`.
Проверяются фактические варианты источников, отдельные строки вариантов,
InCellDropdown, смысл списков, независимые грани RectSet и правые рамки
столбца «Справка», а не только строка Validation.Formula1.
Финальный кандидат: **1583/0**, `ReleaseCandidate/Config_Presentation_Final.txt`.
После сохранения и повторного открытия точно совпали данные Results и
оформление статусных ячеек; источник не изменен. Структурная проверка
`ReleaseCandidate/Validation_Final.json`: **27/0**.
Пиксельный QA всех 15 страниц из восьми PDF завершен: `FinalVisual/Visual_QA.md`.
Правая рамка «Справка», независимые грани RectSet, шапки и формулы просмотрены.
Полный текст длинных ячеек Results при сохраненных пользовательских ширинах
доступен в строке формул; перенос строк данных не включался.

## Сохранность И Код

`ReleaseCandidate/Preservation_Final.json`: **137179 ячеек, 332 ширины,
0 отличий**, пользовательский execution report сохранен. Исключены только
согласованно перестроенная нижняя полоса Results и пять исправленных ячеек шапки.
Первый save/reopen runner остановлен техническим COM-null; повтор после
прямого IDispatch getter дошел до чтения, но получил RPC_E_SYS_CALL_FAILED
(80010100). Еще один Unicode-runner получил тот же RPC до импорта/тестов.
Специальный reader/export/plot reopen после последовательного повтора PASS:
`ReleaseCandidate/Snapshot_Reopen_Sequential_Final.txt`. Все 39103 символа
снимка совпали точно после изменения семи Config-параметров и save/reopen;
чтение/export/plot вызвали **0 новых solve**, исходная книга не изменена.
Предыдущие технические отказы сохранены и не считаются PASS.
Отдельный адресный Unicode/NotCracked
gate завершен: `ReleaseCandidate/Unicode_NotCracked_Final.txt`, **240/0**.
Данные Results и цвета статусов также точно совпали после save/reopen этого теста.

Фактический VBE проверен по всем **120 исходным компонентам и типам**:
`ReleaseCandidate/VBA_Equality_Final.json`, **0 missing/differences/extra**;
всего экспортированы 125 компонентов с документными модулями.
Документные модули листов/книги отдельно генерируются сборщиком. Нормализация
строго ограничена печатным форматом VBE: регистр кода, Double G15, ANSI CP1251
комментарии и окончания строк. Строковые литералы, пробелы и типы целых
констант не скрываются; положительные/отрицательные самопроверки verifier
сохраняются в JSON. Raw-comparison также сохранен.
Два тестовых литерала psi, которые VBE заменял вопросительным знаком,
исправлены через ChrW; затронутый тест прошел 240/0. Десять self-checks verifier
прошли в PowerShell 7 и Windows PowerShell. Windows PowerShell требовал UTF-8 BOM
для русских/Unicode self-checks самого скрипта; исправлена кодировка verifier,
а не ослаблены сравнения. Фактический финальный TXT
`ReleaseCandidate/VBA_Final_Verification.txt` экспортирован из сохраненной книги
без импорта исходников и без ее изменения.

## Выпуск И Self-Audit

Все проверочные gates завершены. Опубликованы следующие артефакты:

| Файл | Размер, Байт | SHA-256 |
|---|---:|---|
| `workbook/output/RC_Section_NDM.xlsm` | 5537034 | `D48A5FFCC17A3FC4F1F2A6FC7080952BD01BA08FCB7151CFBF15EA6E887EDB25` |
| `workbook/output/VBA_All_Code.txt` | 5772286 | `E20F213475B2DF0592E4CEECABA4E706C7B1AADBB822D23FD6A3E0D197B5EEFE` |

Копии побайтно совпадают с проверенным кандидатом и его фактическим VBE-экспортом.
Непосредственно output-файл также проверен: `ReleaseCandidate/Published_Output_Validation.json`,
27/0, исходный hash не изменен. Проверка пробелов исходников не выявила ошибок;
фактический VBA TXT оставлен сырым экспортом VBE, включая его пустые строки
и пробелы, чтобы не подменять выгрузку ручным редактированием.
Последний пользовательский `RC_Section_NDM_execution_report.txt` не подменен:
33154 байта, SHA-256 `05303BCAFDF04378813F3EE6F30109A0258D608DF92C0E280E77E593F304AC79`.
Тестовые логи хранятся отдельно в `docs/regression/PostAudit03/`.
Локальная фиксация выпуска выполняется с темой
`Complete post-Audit03 fixes and separate contour snapshot`, без push.
Пользовательский dirty-report, исходное untracked ТЗ и старые QA-файлы
в эту фиксацию не включаются.

Финальный self-audit от начала до конца ТЗ:

- Шесть пунктов имеют отдельных владельцев и доказательства в таблице выше;
  запрещенные обратные зависимости и инженерная подготовка writer удалены.
- Исходные N/Mx/My/LoadPath записываются только при добавлении LC. Рабочий
  запуск создается заново; прежние result-ссылки не мутируют.
- Formation сохраняет подтвержденную точку при отказе Post и не выводит
  пробные числа как NotCracked. Named-state идет через общую policy без whitelist.
- beta/radius только в CSP35CrackData; Rr/psi/ширина остаются у Width.
  Неделимость групп, t/L и прежняя физика не менялись.
- Готовая Region/clone не создают Query. Напряжения элементов готовятся
  существующим state по фактическим frozen-диаграммам, не по позднему Config.
- Контуры один раз в отдельном 17-колоночном v1-блоке, Geometry только 15
  механических полей; LC-области трещин остаются в аннотациях. Legacy-reader
  ограничен явно вызываемой миграцией, не обычным чтением/export/plot.
- Сборка/миграция/readers/preview/clear/layout/help обновлены; protected и
  broken-name случаи проверены, новый solve/import при переносе запрещен.
- 93 класса до/после, добавленных/удаленных `.cls` нет. Математика и допуски
  не подгонялись; численная сверка точная, память и Timer оговорены выше.
- Списки/рамки проверены в чистой сборке и сохраненном кандидате; 15 страниц
  просмотрены визуально. Реальный AutoCAD отделен от COM-fixture.
- Сохранность Config/верхних Results/332 ширин/report и фактический VBA
  подтверждены. Технические и исторические FAIL-логи не скрыты.
- Независимая нормативная приемка не входит в этот локальный архитектурный
  выпуск; открытые нормативные вопросы остаются явно открытыми.

По последнему указанию пользователя будущие самые долгие полные прогоны
требуют отдельного явного согласия в чате. Ранее начатые UI завершены;
новые полные тяжелые наборы без согласия не запускаются. Правило закреплено
в AGENTS.md и не заменяется общей просьбой завершить цель.

## Пакет Для Внешней Проверки

Нельзя оценивать актуальный код только по отчету Audit03 от 05.10.2026:
после него реализованы СП35 и изменения этого задания. Передавать единый пакет:

- `workbook/output/RC_Section_NDM.xlsm` и `workbook/output/VBA_All_Code.txt`
  именно финального выпуска, с hashes из этого отчета.
- `NDM_Audit03_Implementation_Spec_2026-10-01.md` и `NDM_Audit03_Final_Report.md`:
  исходные требования и исторический итог Audit03.
- `трещины_сп35/NDM_SP35_CrackWidth_Architecture_TZ_v6_CodexGoal.md`,
  `трещины_сп35/NDM_SP35_Implementation_Final_Report_2026-10-06.md`,
  `трещины_сп35/1.png`, `2.png`, `3.png`: ТЗ, отчет и эталон оформления СП35.
- `NDM_SP35_Neighbor_Ratio_2026-10-06.md` и
  `NDM_SP35_Indivisible_Group_Restore_2026-10-06.md`: окончательные правила
  боковых границ и неделимости групп; более ранний частичный отбор отменен.
- `NDM_PostAudit03_Fixes_And_Contours.md`, этот отчет и
  `NDM_PostAudit03_Progress.md`: текущий объем, результаты и история проверок.
- `AGENTS.md`, `CoordinateSystem.md`, `OpenNormativeQuestions.md` и использованные
  локальные нормативные источники из `norms/`: контракт проекта и явно открытые
  вопросы нормативной приемки.
- Указанные в отчете финальные логи/JSON и native DWG из
  `regression/PostAudit03/`; исторические FAIL-логи не подменять успешными.

Тестовые копии книг не следует выдавать за выпуск. Проверяющий должен сверить
hash книги/TXT, актуальные уточнения ТЗ, исходники и конечные результаты тестов.
