# NDM Audit03 Progress

Дата начала: 2026-10-01. Цель: полное выполнение Audit03, не только план.

## Baseline И Входные Данные

- Baseline SHA: `df10412f0e0baf918f5e97cbc87b6bf16c3d4cae`.
- Ветка: `codex/material-diagram-architecture`.
- Входной dirty-tree: только новый пользовательский `docs/NDM_Audit03_Implementation_Spec_2026-10-01.md`; сохранить без изменения.
- SHA-256 ТЗ: `F3B32621FD3D6BDFF0313B76599EDD5D93F637971A5F91818902FF4418118E98`.
- SHA-256 output-книги: `AAF5D4FAF06197F01966463DE85BA813EB4216B455AA8504B2DE4F0BA717C012`.
- SHA-256 экспорта: `51C20ABF981EB9AD569F23648E8ABE846E7ECE05637D8043098C155BC8F6F185`; совпадает с Audit03 и отчетом Audit02.
- Классы: 85 production + 3 test. Новые `.cls` запрещены; запланированы два конкретных объединения до 83 + 3.
- Безопасная baseline-копия: `%TEMP%/RC_NDM_Audit03_Baseline_df10412f`; не переключать текущий checkout.

## Шесть Этапов

| Этап | Статус | Выход И Gate |
| --- | --- | --- |
| 1. Контрольная точка и карта проверок | завершен | Baseline build и Off/On 6745/0; 103 исходных модуля сверены с VBE. Фактические Config/help/class inventories, карта владельцев и F-reproducers сохранены; user Config неизменен. Это подготовительный этап, не итоговая приемка остальных этапов. |
| 2. Корректность входа, поиска и метаданных | в работе | F01/F02/F03 и основные F06 контрпримеры имеют runtime evidence; F04/F05/F07 и окончательная приемка еще не завершены. |
| 3. Упрощение архитектуры | в работе | A01/A02 и перенос агрегации A03 имеют runtime evidence. Один итог crack workflow, изоляция Search и комментарии всех путей проверены; окончательная проверка всех классов/consumers A03-A05 продолжается. |
| 4. Измеряемая оптимизация | в работе | P01/P03 benchmark v6: 160 измерений, 0 ошибок, exact duplicates 10 -> 0; P02 сохраняет 3600 волокон и точную pi. Финальная повторная приемка на выпускном исходнике еще нужна. |
| 5. Config/краевые нагрузки/документация/UI | в работе | Все 13 форм и 52 all-path runs v68 приняты для того численного среза. v95 имеет 79 адресно принятых активных полей из 1065; остальные поля, полный диапазон/pairwise, комментарии всех методов, выпускной help/UI и актуальный повтор матрицы еще открыты. Search-селекторы v96-v102 проходят отдельную приемку, не blanket PASS. |
| 6. Независимая приемка и выпуск | не начато | Полная отдельная сборка, все suites Off/On, config/edge/benchmarks/snapshots, все три audits, final report, source/book/export equality. |

## Карта Обязанностей

| Владелец | Сохраняемая Или Уточняемая Ответственность |
| --- | --- |
| CLoadCombinationReader | Структура 5/6/7+ колонок, отдельные пустые/нулевые/ошибочные/малые ненулевые строки, source slot, row-level errors. |
| CUnitSystem | Единственный пересчет input/output units/signs; overflow conversion не скрывается. |
| CSectionSolver / CStateSolutionRunner / CStateProvider | Typed failure, terminal retry policy, единый solve и scope/reuse; без лишнего solve из потребителя. |
| CSectionStateResult / CResultMeta | Независимый snapshot, точная причина, Applies/Calculated и физические flags; общая blocked meta без инженерной формулы. |
| CResultStatusPolicy / modStatusFormatting | Единственный внешний словарь/приоритет; отдельная существующая палитра. Не получают инженерные формулы удаляемой Batch policy. |
| CLoadMultiplierSearch / CUltimateStrainSearch | Общий numerical search, progress/representability/overflow guards; не трактуют Capacity BaseFail по тексту. |
| Capacity / Formation calculators и adapters | Физический критерий/материалы/load path/проверенная финализация; локальный probe cache и воспроизводимые retries. |
| CCrackWidthCalculator | Подготовка crack data и отдельные чистые численные методы после объединения Formula; без State solve. |
| CLongitudinalCrackCalculator / CStabilityCalculator | Собственные инженерные meta и понятные причины отказа. |
| CCrackResult / CStrengthResult / CCombinationResult | Итоги соответствующих subtrees, общая policy для приоритетов; один численный источник, изолированные snapshots. |
| CBatchSectionCalculator | Порядок LC/этапов/context/governing; без второй формулы/search/flat facade. |
| Geometry classes | Валидация/подготовленные характеристики при мутации, ContainsPoint без повторной дорогой подготовки. |
| Writers / CExecutionReport | Размещение, units, presentation, готовые comments/status; не исполняют инженерную проверку. |

## Реестр Требований

| ID | Статус | Доказательство / Следующий Объем |
| --- | --- | --- |
| F01 | directed runtime PASS | Реальные Range 5/6/7/9 и 4-column rejection; negative/positive_v3 logs. |
| F02 | directed runtime PASS, расширение покрытия впереди | abc N/M, формулы/CVErr, numeric string, Empty/zero/space, tiny в двух системах единиц, overflow и последующий LC; требуется полный output/comment matrix и отдельный Null путь. |
| F03 | directed runtime PASS, расширение проверки продолжается | Provider -> runner -> state: 73/0; real Capacity typed failures/precision: 54/0; Formation все четыре пути и terminal config: 64/0. Два старых expected general numerical code уточнены до rcSingularTangent, внешний NumFail и физика не менялись. |
| F04 | directed runtime PASS, workflow приемка продолжается | SetResult обеспечивает lifecycle для not-requested/not-applicable/blocked/validation; ранние missing/internal factories передают False. Matrix 172/0: десять исходов x attempted, clone/reset и result factories; повторные реальные LC еще входят в расширенную приемку. |
| F05 | directed runtime PASS, расширенная приемка впереди | Физический отказ, BaseFail, numerical failure, blocked Width/Longitudinal, отсутствие трещины: actual writers всех путей. Исправлена передача успешного StopReason в blocked reason. Save/reopen/report/остальные формы еще впереди. |
| F06 | основные directed runtime PASS, полный аудит продолжается | Baseline adjacent-Double Bisection budget=0 завис до watchdog; исправленный generic matrix 114/0, real Capacity tol=1e-18/budget=0 завершен за 58 probes без ложной точки. Остальные Ultimate/recovery call paths еще проверяются. |
| F07 | в работе | Исправлены Nothing guard в stability, Split bounds в plotter и два unsafe test guards. Статическая проверка IIf/array call paths выполнена частично; окончательная приемка не заявлена. |
| F08 | трассировка подготовлена, приемка продолжается | Per-ID реестр NDM_Audit03_Prior_Audit_Traceability.md; незавершенные K/T/D/UI/save-reopen пункты не получают PASS. |
| A01-A05 | в работе | 83 + 3 classes; A01/A02 и текущий A03/A04 срез с actual writers прошли full Off v13. Crack/Strength сами собирают свои итоги; общий приоритет без дублей; shared named-state сохранен. Полная acceptance всех остальных классов и A05 еще впереди. |
| P01-P04 | в работе | 15 сценариев x 5 повторов x 2 версии; asymmetric: 94 solves/3874 iterations/20 retries/10 эквивалентных попыток. P02 численные и point-grid инварианты подтверждены; окончательный benchmark/разбор повторов впереди. |
| K01-K04 | структурный реестр в работе; поведенческая приемка впереди | Read-only census фактической книги: 35617 ячеек, 14 диапазонов, 72 validation. Поадресный реестр 1065 полей, 241 default из каталога; pending metadata не получает PASS. Полный behavior/active-inactive/mutation sensitivity/isolation еще не завершен. |
| T01-T05 | в работе | v68: все 13 форм, Light/Stress, Off/On, 52/52 chunks и 18 980 all-path cases, failed=0. Это приемка численного среза v68, не поздних Config/UI правок; selector variants/pairwise/high-risk tuples, независимые near-limit gates и финальный повтор на выпускном source впереди. |
| D01-D02 | в работе | Comment-only ревизия export/writers/enum/workbook entrypoints выполнена частично; все остальные methods/tests и фактический help/Config/validation/links в clean/update впереди. |
| UI01 | directed COM PASS, выпускная приемка впереди | v39 Off/On: 351/0; семь статусов, DisplayFormat, чувствительность к чужому CF, очистка старых строк и сохранность оформления после save/reopen. Проверка clean/update итоговой книги еще предстоит. |
| W01 | в работе | Git/base/spec/hash/progress сохранены; checkpoints без push/destructive Git. |

Историческая поадресная приемка v95 сохраняет denominator 1065; адресно
принято 79 активных полей (13 Solver, 23 Material, 15 Unit/Sign, 1 Worst,
16 общих RectSet selectors, 1 AutoCAD MinArea, 10 Crack). Остальные 986 адресов не получают blanket PASS.
Срезы metadata/full-range/downstream и K03 остаются отдельными задачами.

## Актуальная Точка Продолжения

- Geometry Config v134 зафиксирован коммитом `78fd486`. Следующий F07 срез
  v135 проверяет представимость счетчиков трех некруглых builders до циклов
  и выделения массивов. Negative `7/18`: реальные Overflow=6 воспроизведены
  для RoundedRectangle H и HollowRectangle B при n=2147483647; остальные
  случаи намеренно имеют неверный отступ, чтобы старый код завершался быстро,
  не выделяя огромные массивы. Новая общая численная оценка запрошенных
  позиций учитывает дополнительные ряды/EverySecondBar и выключенные линии;
  Rectangle не включает нижние грани. Это техническая граница Long, не новый
  физический/нормативный лимит армирования. Opening сохраняет фактическое
  количество проекций, общая раскладка проверяет фактический счетчик.
  Directed positive `29/0`, source/export `103/103`, failed=0; full Off
  `23324/0`, общие actual-числа с v134 `5577/5577`, missing=0, differences=0.
  Help `2536/2536`, mergesEqual=True. Полный On завершен `23335/0`, общие
  числа с v134 `5582/5582`, missing=0, differences=0. Universal `827/0`,
  palette `351/0`: Results и status-style save/reopen=True. Read-only
  formatting `1003/0`, deviations=0; Validate `25/25`. Все процессы завершены.
  Книга SHA `C5812332A3B8BC705EC0C6D02E3FA1F16ADEC1E3C502DD38252A3A52FFB5BFA3`,
  export `367F712EB9BEB75550C32AF3C9C5ACC4A1D1D133F4FFC8BF2051E45BE57ABA3A`.
  ТЗ и основная output-книга сохраняют исходные хеши. Registry `118/1064`,
  fullAcceptance=False: counter gate не добавляет поадресный Config PASS.
  Census 103 модуля, 4240 методов,
  1490 guard-кандидатов, 83 production + 3 test classes. Новых `.cls` нет;
  Formation absence -> psi=1/warning и current-state blocking не менялись.
  Следующий цельный geometry input gate: обязательные активные размеры/
  selectors и raw-input/адреса в таблицах; при missing/blank/TODO/невалидном
  вводе не подставлять программные defaults. Разрешенные blank/0 выключения
  рядов и неактивные геометрические поля рассматривать отдельно. Пока это
  обнаруженный call-site риск, не завершенный runtime gate.
- Срез v133 зафиксирован коммитом `27d55e4`. Текущий следующий geometry Config
  срез v134: 24 параметра форм + четыре наружных счетчика HollowRectangle.
  Negative `161/40` подтвердил молчаливое отключение наружной грани при abc
  или ошибке формулы; raw значение и адрес теперь сохраняются, отрицательный
  счетчик объясняет builder. Blank/0 и автоматические количества отверстия
  сохраняют утвержденное поведение. Directed final `213/0`, source/export
  `103/103`, failed=0; full Off `23295/0`, On `23306/0`, все восемь suites.
  Общие actual-числа с v133: `5577/5577` Off, `5582/5582` On, missing=0,
  differences=0. Universal `827/0`, palette `351/0`: Results и status-style
  save/reopen=True; formatting `1003/0`, deviations=0; Validate `25/25`.
  Help clean/update `2536/2536`, mergesEqual=True. Новый адресный evidence
  принят только после обоих full gates: registry `118/1064`, fullAcceptance=False.
  Подробности в NDM_Audit03_Geometry_Config_Review.md. Книга SHA
  `75EBCDE707473B914DAE64233B6428319F3247960AC4E8E2C3C9F7EBA9291498`,
  export `0736EB9FB69BC5C235CE3FE1776AF8110CB252D7CF3122B664AD257688179B74`.
  Все процессы завершены; срез готов к scoped checkpoint. Новых `.cls` нет.
  ТЗ и основная output-книга сохраняют исходные хеши. Следующий цельный срез
  геометрии: представимость счетчиков некруглых builders и оставшиеся поля
  арматуры, затем полные диапазоны/pairwise; не начинать новую архитектуру.
- F05 срез v131 зафиксирован коммитом `25a5cab`; все его directed/full/
  numeric/save-reopen/formatting/структурные gates завершены. Следующий рабочий
  срез F07/D01/K02: ранняя сумма рядов Circle, содержательные комментарии и
  активная проверка фактических Config-полей круга. Negative count guard v132
  `4/8` воспроизводит неверную причину координат до проверки количества;
  positive `12/0`. Это не наблюдение миллиардного цикла/фактического Overflow:
  малый radius fixture специально дает быстрый отказ старого маршрута.
  Full Off v132 завершен `23029/0`; общие числа с v131 совпали `5574/5574`,
  missing=0, differences=0; source/export этого снимка `102/102`, failed=0.
  Последующие comment-only изменения Width и новый standard test module
  modTestGeometryConfig входят в следующий v133 snapshot: directed Config
  `53/0`, source/export `103/103`, clean/update help `2536/2536`.
  Полные v133 Off/On завершены `23082/0` / `23093/0`, все восемь suites,
  watchdog exit=0 и source unchanged=True. Общие числа с v131 совпали:
  `5574/5574` Off и `5579/5579` On, missing=0, differences=0. Universal
  `827/0`, Results save/reopen=True; palette `351/0`, values/style reopen=True;
  read-only formatting `1003/0`, deviations=0; Validate-Workbook `25/25`.
  UI-suite сохраняет исторические явные On setup; его копия закрывается без
  сохранения, следующий suite открывает исходный Off fixture. Этот override
  явно отмечен runner-ом, не скрытое расширение обычного Off расчета.
  Реестр текущего Config содержит 1064 адреса; 82 исторически принятых поля
  v103 перенесены по Id без удаленного InitiationLoadPath, восемь новых полей
  Circle имеют actual active/invalid/recovery/inactive evidence. Итого 90,
  fullAcceptance=False; исходные диапазоны/validation/help взяты из v133.
  Новых `.cls` нет. SHA книги v133
  `6E9A5EDF4EA5A769C130DBEEF01E7B98A94151EF18B6EA5BCFEB0430CE1894A6`,
  export `6CDC1CF6D85DF6308E5A362A1FE65905C40E94C4531CA34D5B5E081BB898B111`.
  Все процессы gates завершены. Следующая K02/F05 работа: полный input/range
  contract геометрии и адресная приемка оставшихся семей Config, затем
  актуальные broad/pairwise/near-limit и benchmark перед финальным DoD.
  Отсутствие Formation/Post по-прежнему допускает Width с psi=1/warning;
  только обязательный текущий State блокирует зависимые crack-проверки.
- Census v132: 102 модуля, 4210 методов, 1481 кандидатов guards, 83 production
  classes. Template candidates=0; среди методов без собственного комментария
  нет body длиннее двух строк. Это индекс и проверка конкретных комментариев,
  а не автоматическая содержательная приемка всех методов/guards.
- Проверенный Search lifecycle/K04 срез v130 зафиксирован локальным коммитом
  `2e17181`. Текущий следующий срез F05 - итоговое сообщение InputErr.
  Negative v131b `9/3`: теряется ResultComment DirectState, Capacity и
  CurrentCrackedState. V131c добавляет требования к имени Config/ключу Capacity;
  прежние журналы сохранены. Подготовлена локальная правка: popup читает готовый
  OverallMeta.ResultComment; Capacity называет ошибочную строку настройки.
  Positive v131: popup `14/0`, Capacity lifecycle/config messages `72/0`;
  full Off завершен `23017/0`, восемь suites; source/export `102/102`,
  failed=0. Числа общих Off assertions с v130 совпали: `5574/5574`, missing=0,
  differences=0. Clean/update help `2536/2536`, mergesEqual=True.
  Полный On завершен `23028/0`; общие численные assertions с v130
  совпали `5579/5579`, missing=0, differences=0. Universal `827/0`,
  Results save/reopen=True; палитра `351/0`, values/style save/reopen=True.
  Read-only Config formatting `1003/0`, deviations=0; Validate-Workbook
  выполнил 25 структурных проверок, все True, exit=0. Срез готов к локальному
  checkpoint, но не закрывает весь Audit03. SHA книги v131
  `67952A24F354E94819CE78B66D63BA19E633AA79F637F3F22748E9603DCEF2E5`,
  export `650349E6851692471CD1F280267986ACEA0BA3C0B36FFF210A622818719BBCE6`.
  Auto-path гипотеза потери флага
  не подтверждена: v131b выполняет все три моментные траектории и сохраняет
  SearchExecuted=True, итог rsNumericalFailure. Два неудачных assertions были
  неверным ожиданием теста, production Auto-поиск по ним не изменен.
- Текущий проверяемый source-срез v130: full Off `22969/0`, On `22980/0`,
  по восемь suites, watchdog exit=0, source unchanged=True. Книга SHA256
  `961FF06B3A1272B3FBF0405E6C0D86AD059BBAF1C8A3D1105E3EB41ECE4A257B`,
  export SHA256 `2E247A7EC8F7BBE67F2158D8BA5F79A9C09034D35BF89422B609CCA8A0232A50`.
  Source/export `102/102`, failed=0; clean/update help `2536/2536`,
  mergesEqual=True, failed=0. Census v130: 102 модуля, 4205 методов,
  1485 кандидатов guards; это индекс дальнейшей ревизии, не blanket PASS.
  Проверка последовательности прежних 17 tests + pair/Search `3595/0`.
  Точные actual-числа общих assertions относительно принятой v112 совпали:
  Off `5574/5574`, On `5579/5579`, missing=0, differences=0. Скрипт сравнения
  проверен на намеренном изменении числа и пропуске ID: оба выявлены.
  Это сравнение чисел журнала, не blanket приемка всех сохраненных Results.
  Directed v130: Capacity `38/0`, Formation `147/0`, временный диапазон `10/0`,
  universal On `827/0`, Results save/reopen=True; палитра `351/0`, значения
  и status-style save/reopen=True. Config `1003/0`, 1002 адреса,
  deviations=0; validation `25/25`, exit=0. Все COM-проверки завершены.
  Проверенный срез готов к локальному checkpoint; полный Audit03 не завершен.
  Следующий отдельный F04-контрпример: факт выполненной попытки при Auto-переходе
  между разными LoadPath, когда следующий путь завершается до probe.
  Primary output и неизменяемое ТЗ сохраняют исходные хеши.
- Принятый universal LoadPath/Width-fallback срез зафиксирован коммитом
  `7839c65`. Текущий следующий срез -
  [Search lifecycle F04](NDM_Audit03_Search_Lifecycle_Review.md).
  Negative Formation v113: 120/27, Capacity: 15/7 на прежнем production.
  Positive v113: Formation 147/0, Capacity 26/0; окончательный v114 directed
  Capacity/generic 31/0. Реальные callbacks хранят факт численной попытки,
  аналитический Ncrc не считается Search. EarlyStopReport Capacity использует
  русский ResultComment вместо внутреннего LimitState.
  v114 source/export 102/102, failed=0, но full Off остановлен ошибкой
  компиляции новой строки EarlyStopReport (mMeta вместо mResultMeta).
  Этот прогон не принят; его тестовый Excel закрыт, RPC-error сохранена.
  v115 directed Capacity/generic/report 34/0, source/export 102/102, failed=0;
  full Off v115 не завершен из-за runtime-ошибки 14 `Out of string space`
  в batch-suite. Изолированный Search Config прошел `1813/0`, но отдельная
  batch-suite аварийно завершила Excel с `RPC_E_SERVERFAULT`; оба отрицательных
  журнала сохранены, причина пока не установлена. Пустой Excel с панелью
  восстановления закрыт отдельно. Чистая сборка v116 завершена, source/export
  `102/102`, help `2536/2536`, mergesEqual=True, failed=0; ее full Off тоже
  остановлен ошибкой 14, память Excel достигла 3 572 256 768 байт.
  Версия о проблеме только обновленного VBA не подтверждена. Запущен видимый
  batch v116 для локализации строки; он тоже аварийно завершил Excel.
  Повтор batch на принятой v112 также аварийно завершен с ростом памяти
  до 3 582 541 824 байт; привязка отказа только к F04 не подтверждена.
  EventsOff v117 не устраняет аварию. Тестовый диагностический entrypoint
  изолирует настройки Application и не заменяет обычный full gate.
  Найден отдельный direct Auto lifecycle пропуск: negative v118 `37/1`,
  единственная ошибка `directAuto.firstAttemptKept`. Исправлены только три
  строки сохранения флага; положительная v118 Capacity `38/0`, source/export
  `102/102`, failed=0. Export SHA256
  `E14D01D53F8D8829EC6FD940886406628A1BADD7F1F2B7F3BE8B19CD5AE2F52B`.
  ManualCalculation v118 тоже аварийно завершен; режим пересчета не устранил
  этот воспроизведенный отказ. v119: 75 повторных записей одного готового batch
  завершены, privateBytes после набора 97 632 256; размеры Results, 49 styles
  и отсутствие CF стабильны. Полный batch Observe v119 вновь аварийно завершен;
  контекст диапазонов нормальный, память росла уже на On/Off pair и затем
  Search Config. Эти факты не устанавливают причину утечки/роста памяти.
  Добавлены только тестовые диагностические входы и отдельный журнал контекста.
  v120 сравнивает одинаковый Search Config при вложенном выражении склейки
  и при отдельном вызове: оба `1813/0`, численные expected/tolerance не менялись.
  v121 изолированный On/Off pair + Search Config: `3261/0`, память после
  138 973 184 байт. В полном batch перед pair уже 517 622 символа отчета;
  этап занимает около 55 секунд вместо одной в изолированном вызове.
  Сравнение 892 значений Config показывает только ожидаемый temporary
  Stability PR1. v122 report-only 3000 строк поверх 500 000 символов завершен
  без сбоя, privateBytes 90 656 768; pair + Search с таким же длинным префиксом
  `3261/0`. Размер отчета сам по себе причину не воспроизводит.
  v123 заключительная группа перед pair воспроизвела ошибку 14 уже в Search
  Config; Excel закрыт после чтения диалога, точная строка Debug недоступна
  (coordinate input geometry is unavailable). v124 контроль None без
  предварительных тестов `3261/0`; далее изолируются Writers/Meta/Repository.
  Repository v124 `3364/0`, но Writers снова воспроизвел runtime 14 при
  3 567 603 712 байт privateBytes. После чтения диалога закрыт только тестовый
  Excel, отрицательный исход сохранен. V125 отдельные действия: Summary
  `3425/0`, Rows `3263/0`, Invalid `3264/0`; production остается v118,
  assertions/tolerance не менялись. Settings/Gaps/Meta и взаимодействие
  действий еще проверяются, root cause не установлен.
  Продолжение диагностики: Settings `3263/0`, Gaps `3268/0`, Summaries
  `3432/0`, OtherWriters `3268/0`. Summary + Rows воспроизводит ошибку 14;
  Rows не восстанавливал активный лист и оставлял Results. ScreenOff дает
  `3427/0`, но 1,16 ГБ памяти, поэтому не принят как исправление. K04 negative
  v129 `5/1` подтверждает именно activeSheetRestored. Cleanup возвращает
  исходный лист до удаления временного: isolation `6/0`; тот же исходный
  ресурсный порядок `3427/0`, privateBytes 140 165 120, Config остается активным.
  В штатный Rows-тест v130 добавлены четыре restoration assertions.
  Внутренний механизм аллокации Excel не заявляется установленным; полные
  обычные Off/On и saved Results остаются обязательными после этой правки.
  Численная причина аварии этим не установлена и NumFail не назначается.
  Далее диагностика вызовов/объектов, затем обязательны
  full Off/On и saved Results gate; общий DoD остается открытым.
  Основная книга и ТЗ не изменены; общий DoD остается открытым.
- Checkpoint Search-среза: `ff5caa46737f64be4d8e5d11e9e99eed7cf5c019`.
  Принятый следующий срез v112: единый LoadPath согласно
  [дополнению текущей цели](NDM_Audit03_Universal_LoadPath_Scope.md).
  Оно включено в цель после завершения предыдущих gates и коммита;
  направленные gates проходят отдельную приемку. Исторические доказательства
  v103 не подменяют тестирование общего перебора и нового fallback Width.
- v110 directed Universal: 782/0, включая actual writer psi_s/a_crc/Es/status,
  save/reopen=True; source/export 102/102, failed=0. Полный Off v109 прошел
  восемь suites без failures, но предшествует позднему исправлению writer и
  не заменяет full Off/On v110. Финальные full v110: Off 22552/0, On 22563/0,
  восемь suites; directed On также 782/0 с Results save/reopen=True.
- Negative writer gate: 779/3; прежний Formation guard скрывает actual
  psi_s/a_crc/Es и обнаруживается независимыми cell assertions. Только этот
  guard изменен в отдельной книге, статус Width-meta не подменяется. Начальная
  попытка создания fixture была отвергнута до мутации из-за регистра VBE;
  точное единственное case-insensitive сопоставление устранило ошибку runner-а.
- v110 formatting: 1003/0, 1002 адреса, deviations=0, source unchanged=True;
  Validate 25/25. Clean/update help OpenXML: 2536 непустых ячеек, объединения
  равны, failed=0. Это не пиксельная приемка справки. Первый новый broad-run
  CircleUneven/Light/Off v110: 588 случаев, 23882/191. Отрицательный журнал
  сохранен; следующие семь chunks не запускались после этого отказа.
- Подтвержденная причина v110: новый fallback запускал Width для нулевого
  текущего НДС, а общая ветка делила на нулевую кривизну. Исключение Overflow
  уходило в общий batch handler и искажало статусы/комментарии других блоков.
  v111 обрабатывает равномерное поле до построения нейтральной линии: текущее
  растяжение использует существующую центральную ветку; нулевое/сжатое поле
  дает rsNotApplicable. Дополнительный скрытый порог 1e-9 удален из описателя
  LoadPath: ненулевой нормализованный ввод остается масштабируемым. Добавлены
  направленные тесты. v111: Crack 880/0, Universal Off 822/0, Results
  save/reopen=True; source/export 102/102, failed=0. Повтор восьми scoped
  matrix chunks остановился на Stress/Off: 13243/8, 288 случаев. Оба Light
  Off/On завершены: 24117/0 и 23108/0, по 588 случаев с save/reopen=True.
- Второй подтвержденный дефект: CCapacityResult при принятом пределе lambda<1
  сохранял FAIL, но заменял ResultComment Search и переносил историю Auto в
  DiagnosticDetails. v112 сохраняет прежние причины в пользовательском
  комментарии, добавляет инженерное объяснение FAIL и не подменяет diagnostics.
  Добавлен независимый result-level тест; отрицательный broad v111 сохранен.
- v112 полный Off: 22780/0, все восемь suites; embedded Universal 827/0 и
  Search 1813/0 не добавляются второй раз к общему числу. Source/export
  102/102, failed=0; help compare 2536 ячеек, mergesEqual=True, failed=0.
  Полный On также завершен: 22791/0, восемь suites. Все восемь scoped
  нагрузочных chunks завершены: 3504 случая, 147163/0 assertions, Results
  save/reopen=True в каждом, source unchanged=True. CSV v112 содержит только
  эти восемь логов: 200 Width fallback без точки Formation, все с psi_s=1
  и предупреждением; блокировок при успешном CurrentState и формульных
  rsNumericalFailure нет. Общий Audit03 не считается закрытым.
  Проверяемая книга SHA256:
  `32EF5B9594561C8E25596B4FCCA533829E7EFF3FB8484E3C7A4F78678D174A28`;
  export SHA256:
  `A7EBB0C0B72DD5EE95A5624BD1D6426E6B202EED2A90BA7DAEFE70DC46D7A83C`.
- Directed On v112: 827/0, отдельный Results save/reopen=True. Formatting:
  1003/0, 1002 адреса, deviations=0; Validate 25/25. По принятому срезу v112
  checkpoint определяется в git log по сообщению
  `Audit03 universal load paths and conservative crack fallback`.
- Актуальная отрицательная writer-мутация v112: 824/3, ожидаемые failures
  только actual psi_s/a_crc/Es. Изменен один Formation guard в отдельной книге;
  статус Width не подменяется. Save/reopen=True, source unchanged=True.
  Для продолжения Audit03 следующий подтвержденный кандидат - F04
  SearchExecuted ниже; общий Config/F07/D/P/high-risk/final объем остается открытым.
- В Width удалена зависимость допуска формулы от успеха Formation. Отсутствие
  пригодной точки/AfterMcrcState дает psi_s=1 и предупреждение, при пригодном
  текущем НДС; Longitudinal независима. Formation сохраняет свой typed исход
  и общий статус. Actual writer теперь использует собственный Calculated
  Width, а не Formation.CrackFormed; negative mutation gate принят выше.
- Новый census v108: 1064 адреса, 84 validation. Из знаменателя удалена только
  SLS.Crack.InitiationLoadPath; LC LoadPath остается в реестре. Исторические
  per-address evidence не превращаются автоматически в приемку новой книги.
  Новая нагрузочная матрица имеет 12 вариантов: шесть для каждого критерия,
  Light 588 и Stress 288 независимых случаев на форму и Extension-режим.
- F04-кандидат после v112 воспроизведен и исправлен в текущем v114 source:
  BuildSearchSnapshot Formation и Capacity BuildSnapshot больше не используют
  default SearchExecuted=True. Negative/positive lifecycle доказательства
  перечислены выше; окончательные full/save-reopen gates этого среза еще идут.
- Итог текущего Search-среза v103: directed 1812/0; полные восемь suites
  Off 21534/0 и On 21543/0; unit/sign equivalence On 1584/0;
  formatting 1004/0, Validate 25/25, source/export 102/102.
- Все четыре изолированные selector mutations обнаружены соответствующими
  active assertions (33/26/30/36 ожидаемых failures). Stress-повтор
  CircleUneven/HollowThin, Off/On: 4/4 chunks, 960 independent cases,
  failed=0, Results save/reopen=True, source unchanged=True.
- Проверенная книга v103 SHA256:
  `E42DD9EDEFBEA6BC6A35D1E5E240CDA20DB61A3C063CCC880EC4DE4009A361C5`;
  общий export SHA256:
  `4FA295D689013E602E0D5A1656A17EB83C9919454F9A4065862271C02EB51529`.
  Основная output-книга и исходное ТЗ Audit03 не изменены.
- Checkpoint нового среза выполнять после final gates, не объявляя весь
  Audit03 закрытым. Приведенные ниже записи v101/v102 являются историей
  диагностики, а не текущей незавершенной очередью проверок.

### История Предыдущего Search-Среза

- Предыдущий промежуточный коммит: `59ef7cf92cb25847bd7773996164950944b1b904`.
- Текущий dirty-срез: четыре Search-селектора, направленные Config-тесты и
  подтвержденные дефекты Brent/Secant/Formation line-search. Чужие untracked
  файлы и основная output-книга не меняются, новые классы не добавлены.
- Directed v100 и v101: 1792/0, результаты сохранены и после повторного
  открытия совпадают. Help v101: 1933 строки, 141 ссылка, 118 shapes,
  input records/Print_Area сохранены. Source/export v101: 102/102, failed=0.
- Full Off v101 выявил шесть регрессий точного корня Brent в старых generic
  tests; этот прогон не является приемкой. Исправленный точный корень сначала
  проходит доменный FinalizeAt, узкая скобка отдельно FinalizeBracket.
- Проверка фактических INPUT_MESSAGE v101 обнаружила неполную навигацию для
  двух неизвестных Formation-селекторов. В v102 тест требует реальную ячейку
  либо раздел для утраченной строки. Их нормализация остается у Formation;
  Batch проверяет ее до запуска НДС, с адресной диагностикой reader-а.
- Следующий порядок: завершить текущий Off v101 без второго COM процесса,
  выполнить test-only address negative на его копии, затем Refresh v102,
  directed/full Off/On, formatting/Validate, selector mutations и scoped
  all-path нагрузочную матрицу. Реестр обновлять только после всех gates.
  Общий DoD Audit03 по-прежнему не выполнен.

Уточнение после завершения этих шагов: v102 directed 1812/0 и full Off v102b
21534/0 уже завершены. Full On выполняется. v102 source/export: 102/102,
failed=0; export SHA `B5310EAA47C4B18D952B3CB4DB50C74BCCB858FD893CB46FACDF99F49D55D346`.
Full Off v102 без суффикса b был ошибочным вызовом runner-а с общим
VerifyResultsReopen: этот флаг допускается только для одного выбранного macro.
Он остановился до suites и сохранен как script-error, не numerical failure.
Правильный полный вызов v102b не содержит этого флага; directed save/reopen
выполнен отдельно, SHA `5BE1E24E028DEB9881077C850691975F0AD5F70F4481A235D9D12AAFFAEDBDD1`.

Full On v102 завершен с 128 failures в старом unit-equivalence test, остальные
suites зеленые. Тест записывал ToleranceN/Mx/My в выбранных INPUT-единицах,
но вызывал Batch.ApplySettings без units: внутренние допуски различались
между сравниваемыми вариантами. В v103 исправлен только этот один вызов,
production, expected values и прежний строгий comparison tolerance не изменены.
Directed 72-case On equivalence выполняется; после него повторить окончательные
source/export, Off/On и остальные gates на v103. v102 On не объявлять PASS.

Directed unit/sign equivalence v103 On: 1584/0, все 72 варианта, Results
save/reopen SHA `259C5DA1A69923917EA72EEB24184ADF68B0865F76117277404499E301880952`.
Production не изменен относительно v102; исправлен только отсутствующий units
аргумент реального Batch.ApplySettings в тесте. Окончательные directed Search,
full Off/On, source/export, formatting/Validate, четыре mutations и Stress
CircleUneven/HollowThin Off/On выполняются последовательно на v103.

## Important Decisions

- Единственный LoadPath строки LC применяется к обоим критериям. Общий Search
  выбирает Auto по активным N/Mx/My и перебирает M -> N -> NMxy только до
  принятой точки; предел ниже текущей нагрузки не запускает поиск удобного OK.
  Предыдущие причины Auto сохраняются в инженерном ResultComment даже при FAIL.
- Formation/Post failure не является обязательной зависимостью формул:
  пригодный CurrentState разрешает Width с psi_s=1 и warning в любом PsiMode;
  Longitudinal независима. Фактическая Formation meta не скрывается в общем
  результате. Нулевое/сжатое текущее поле без растяжения дает N/A Width,
  а отсутствие или физический отказ CurrentState блокирует обе формулы.
- Checkpoint `a5ed06a`: F07/K02 input-contract guards, all-path физическая приемка Capacity и full v37 Off/On; последующая нагрузочная матрица продолжает приемку, checkpoint не закрывает Audit03.
- Audit02 E08/E09 разрешает техническое расширение промежуточных Formation/Capacity проб. Запрет относится к выдаче extended-кандидата как физической предельной точки, а не к самому общему state-solve. Failed named-state может оставаться диагностикой и сохранять собственный typed status, но не становится reusable физическим state и не включает HasLimitPoint.
- Дополнение пользователя от 2026-10-03: нагрузочная матрица покрывает Auto и пять фиксированных путей (`lambda*Mx`, `lambda*My`, `lambda*Mxy`, `lambda*N`, `lambda*NMxy`) для обоих критериев. Проверять base/offset на lambda=0, фактически масштабируемые компоненты и момент от эксцентриситета N, физический критерий конечной точки, статусы и ResultComment. Не подменять полный перебор тестом только Auto или общим coordinator. Для Auto отдельно проверять последовательность путей и причины перехода; для фиксированного пути - собственный физический результат без незаявленной смены траектории.
- Checkpoint `170ff62d`: проверенный A03/P02/P01/P03 срез и доказательства. Новые input-contract tests были намеренно оставлены отдельным dirty-срезом до направленной и полной приемки.
- Broad load matrix: отдельный parametrized macro в существующем test-модуле, без новых классов. Каждый chunk проверяет все comments в четырех writers и настоящем execution report; watchdog умеет передать shape/family и сравнить все значения Results после save/reopen. Это не заменяет самостоятельные UI/Config и независимые near-limit gates.

- Audit02 сохраняется как история, не переписывается задним числом. Его green logs не являются доказательством новых контрпримеров.
- Не менять физические expected/tolerance; исправленные ошибочные ветви обосновывать F-ID и negative/positive runtime logs.
- Off-suite сохраняет исторические explicit-On setup; effective setting каждого теста фиксируется, изменения копии отбрасываются.
- Только существующие классы. Новые test procedures/standard modules допустимы, новые test classes нет.
- Изначальное замедление asymmetric от исправленной физической классификации не откатывать. Оптимизации сравнивать с df10412f.
- Для output использовать Refresh-VbaModules, не исторический Refresh-Workbook; help обновлять отдельно с доказанной сохранностью Config.
- Excel COM только последовательно, build ждать не менее 300 секунд. Прежнее разрешение пользователя закрывать тестовый Excel сохраняется; не выдавать недоступные DWG/screenshots за PASS.
- Excel COM в текущем окружении запускается с escalation: default sandbox дал 0x80080005, escalated запуск работает. Это ограничение среды, не дефект НДМ. Watchdog закрывает только новые test Excel процессы и считает timeout ошибкой; незавершенные exploratory логи не являются PASS.
- При крайне малом lambda tolerance локальный probe-cache использует точное равенство lambda и вектора нагрузки. Близкие, но различные representable points не склеиваются; повтор идентичной точки остается cache hit. Нет скрытого увеличения пользовательского допуска для объявления сходимости.
- CResultMeta.SetSolverFailure централизует только typed failure mapping и lifecycle; инженерные критерии и внешний display остаются у своих владельцев.
- Checkpoint перед A01/A02: `6373d57` (finite search/lifecycle; Off/On 7289/0).
- A01: CLongitudinalCrackCalculator.CalculateFromStress выполняет содержательную проверку подготовленного напряжения и готовит причину; Width/Stability публикуют свою ResultMeta. Capacity больше не принимает неиспользуемую status-policy. Невостребованные alias/DisplayStatus удаляемого Batch policy не перенесены в common policy; физические LimitState поля writers по-прежнему получают напрямую.
- A02: численная формула перенесена без изменения порядка операций/единиц/zero guards в Width.CrackWidthFromData/UtilizationFromData. Эти методы не принимают State и не меняют статус; отдельный Formula class удален. Новых классов нет.
- A03: strength/crack-specific итоги и комментарии находятся в CStrengthResult/CCrackResult; CCombinationResult объединяет готовые разделы. WorstResultMeta использует один общий приоритет, при равном внешнем OK сохраняет rsSuccessWithWarning. Логические status/code/lifecycle setter-ы CResultMeta удалены; SetResult остается атомарным входом.
- A03 snapshots: опубликованный Search-контейнер не разделяет повторно заполняемый исходный CLimitSearchResult. Getter возвращает независимый небольшой контейнер; конечный CSectionStateResult сохраняет идентичность и не копируется по волокнам. Проверены мутация исходного Search, выданной копии и рабочего solver-а.
- A03 publication: ResultAt завершает заполнение LC и его result-subtrees/repository. Публичные fill/reset методы уже выданного результата отклоняют мутацию; Stability и profile/extension поля доступны только для чтения. Повторный Execute создает новый LC, не меняя старые ссылки. Named-state сохраняет идентичность без клонирования волокон. Рабочий repository внутри еще не опубликованного workflow остается доступным для reuse; повторный crack-cache тест выполняется до ResultAt.
- A04: StoreCrackAggregateSnapshot вызывается один раз в окончательной точке workflow. Промежуточные named-states и диагностика сохраняются до упаковки, ранние выходы получают свои настоящие current/blocked meta.
- Нагрузочный matrix fixture явно ставит ZeroMomentPerDepth=0 и stability=No, затем восстанавливает настройки. Это проверка самой траектории, а не отмена общего фильтра: v10 показал, что малый внутренний My после переноса обнуляется существующим фильтром. Expected численных компонент не ослаблены; фильтр отдельно покрывается прежними boundary-тестами. Реальный effective Extension берется из настроек provider-а, а не из подписи mode-runner.
- Дополнение пользователя: при всех нагрузочных тестах проверять заполнение, инженерный смысл и читаемость каждого ResultComment. Проверять все применимые leaf/meta и subtree-итоги, подробные strength/crack/stability blocks, batch summary, txt-report и повторно прочитанный Results. Обязательная причина каждого отказа/blocked/warning, правильные units/символы, отсутствие противоречия статусу, дублей/двойных точек/чужого раздела; логический порядок сборки. Автоматические проверки дополнять содержательным чтением уникальных шаблонов и фактических контрпримеров, не сводить качество к наличию строки.

## Текущие Проверки

- Прочитаны Audit03 целиком, AGENTS, README/сборочные entrypoints, фактическая Architecture, Audit01/closure и Audit02 report/spec.
- Git/hash/class census: подтверждено, код после Audit02 неизменен.
- Отдельная baseline-сборка выполнена; `baseline_build_2026-10-01.log`. Оба исходных full Off/On: 6745/0; source/export contract: 103 совпадения, 0 ошибок. Логи в `docs/regression/Audit03`.
- F01/F02 подтверждены исходным reader-ом: `f01_f02_reader_confirmed_negative_2026-10-01.txt`, 29/5. Первый exploratory negative содержал ошибку нового settings fixture (2 вместо 3 колонок); он не используется как доказательство production-дефекта.
- Positive reader: `f01_f02_reader_positive_v3_2026-10-01.txt`, 67/0. В v1/v2 новый fixture не передавал обязательный путь и затем SP35 table; это исправлено без ослабления expected.
- F03 negative: `f03_state_confirmed_negative_2026-10-01.txt`, 32/11; positive: `f03_state_positive_2026-10-01.txt`, 73/0.
- Первый полный Off после reader/state/lifecycle изменений: `f01_f05_full_off_2026-10-01.txt`, 6885/0. Это промежуточный gate, не завершение F01-F07 либо всего Audit03.
- F06 baseline: `f06_stagnation_negative_2026-10-01.txt`, timeout=60 секунд на соседних Double при budget=0; подтвержденное зависание, test Excel завершен watchdog.
- Generic positive: `f06_arithmetic_positive_v4_2026-10-02.txt`, 114/0: Capacity/Formation adapters, Bisection/Brent/Secant, соседние Double, large finite bounds, unchanged recovery, final MaxLambda. Ранние v1-v3 timeout сохранены как диагностика, не приняты как доказательства прохождения.
- Real Capacity/F03 positive: `f03_f06_capacity_positive_v6_2026-10-02.txt`, 54/0. Для precision case выполнено 58 реальных probes; invalid method/budget, missing section/material, empty section и singular tangent сохраняют свои typed причины, terminal ошибки не повторяются и не превращаются в BaseFail. v5 выявил необработанную missing-section ошибку request factory; исправление подтверждено v6.
- Formation/F03 positive: `f03_formation_positive_v7_2026-10-02.txt`, 64/0: Auto/LambdaMxy/LambdaN/LambdaNMxy x invalid method/iteration budget; один terminal solve, rsInvalidConfiguration/rcInvalidConfiguration, no point/Pre/Post и Calculated=False.
- Полный Off v6: `f01_f07_full_off_v6_2026-10-02.txt` завершен, но не green: только два failure в Audit02 offset-pair ожидали общий rcNumericalFailure вместо теперь сохраненного rcSingularTangent. Все остальные suites без ошибок. Expected исправлен именно по F03, нагрузки/числа/tolerances не изменены; повторные Off/On после lifecycle изменений еще предстоят.
- Lifecycle positive: `f04_lifecycle_positive_v8_2026-10-02.txt`, 172/0: все десять internal outcomes x attempted True/False, clone/reset и missing/early factories. Search, который действительно запускался, сохраняет Calculated=True при internal failure; отсутствующий/ранний результат не выдумывает попытку.
- Lifecycle negative: `f04_lifecycle_negative_2026-10-02.txt`, 145/27 на unchanged production baseline. Подтверждены неверные flags и missing-result factories; текущий positive применяет тот же тест без изменения expected. В baseline temp source дополнительно заменен только test Batch module, production файлы прежние.
- Полный Off v8: `f01_f07_full_off_v8_2026-10-02.txt`, 7289/0; watchdog exit=0, source unchanged=True. Это промежуточный общий regression gate, не закрытие всех требований Audit03.
- Полный On v8: `f01_f07_full_on_v8_2026-10-02.txt`, 7289/0; watchdog exit=0, source unchanged=True. Исторические explicit-On setups внутри suites сохранены.
- A01/A02 full Off v9: `a01_a02_full_off_v9_2026-10-02.txt`, 7289/0; две удаленные class modules, численные expected/tolerances прежние.
- Paths/comments v10: `a03_paths_comments_off_v10_2026-10-02.txt`, 492/4. Четыре сравнения компонент не учитывали разрешенное обнуление внутреннего My фильтром; этот лог не считается PASS. Кроме того, чтение реальных blocked comments выявило неверную передачу успешного solver.StopReason вместо typed физической причины; исправлено по F05.
- Directed Off v11: `a03_paths_comments_off_v11_2026-10-02.txt`, 495/0, effective Extension=False. Пятнадцать Capacity и двенадцать Formation LC, ненулевой offset N, все leaf/meta и четыре actual output-блока. Честные NumFail проб записаны отдельно от физических отказов.
- Directed On v13: `a03_paths_comments_on_v13_2026-10-02.txt`, 504/0, effective Extension=True; дополнительно проверено, что blocked Width/Longitudinal содержат именно текущую typed-причину, а не сообщение об успешном Newton.
- Full Off v13: `a01_a04_full_off_v13_2026-10-02.txt`, 7798/0: Geometry 496, Material 1359, Section 523, Capacity 1383, Crack 447, Batch 2910, Workbook UI 641, Regression 39. Source unchanged=True. Позднейшая защита typed warnings требует повторного прогона v14.
- Full Off v14: `a01_a04_full_off_v14_2026-10-02.txt`, 7800/0; Full On v14: `a01_a04_full_on_v14_2026-10-02.txt`, 7801/0. Оба watchdog exit=0 и source unchanged=True. Проверено сохранение rsSuccessWithWarning независимо от порядка агрегации и при успешной формуле ширины. Разница числа проверок обусловлена активным Extension, не пропуском suite.
- Source contracts v14: `source_contracts_v14_2026-10-02.txt`, 101/101 модулей совпали с VBE, 0 ошибок; production 83 + test 3, удаленные классы отсутствуют в consumers/export.
- Все тесты выполнялись на отдельных книгах; пользовательская output-книга и Config не изменены. В temporary baseline-source copy добавлен только новый test module для F03 reproducer; production baseline остался прежним.
- Checkpoint `1a65796`: согласованный A01-A04 срез, полный Off/On v14 и all-path comments. Это не завершение Audit03.
- F06 Ultimate negative: `f06_ultimate_stagnation_negative_2026-10-02.txt`, watchdog timeout=60 секунд. В изолированной книге заменен только CUltimateStrainSearch на неизмененный модуль commit `1a65796`; остальные зависимости и новый reproducer текущие. Нулевая граница alpha и неизменная норма действительно зависают; тестовый Excel завершен watchdog.
- F03/F06 Ultimate positive v15: `f03_f06_ultimate_positive_v15_2026-10-02.txt`, 316/0: конфигурация, непредставимое приращение, точные причины отказа Якобиана, контролируемый Nothing и реальные Capacity Ultimate-маршруты. В v16 дополнительно фиксируется ResultKind каждого fake-domain; численные expected не меняются.
- Full Off/On v15: `f03_f06_full_off_v15_2026-10-02.txt` 8261/0 и `f03_f06_full_on_v15_2026-10-02.txt` 8262/0, source unchanged=True. Provider/state matrix расширен с 6 до 11 typed сценариев, Formation четыре пути проверены также с min alpha=0. Для v16 с уточнением test-kind идут свежие прогоны.
- Уточненная приемка v16: directed Ultimate 340/0, полный Off 8285/0 и On 8286/0; все три watchdog exit=0, source unchanged=True. В обоих generic доменах проверен точный ResultKind. `source_contracts_v16_2026-10-02.txt`: 101/101 модулей, 0 ошибок. Физические expected и допуски не изменены.
- Checkpoint Config census/registry: `210a5a20`; это структурная инвентаризация, не завершение поведенческой приемки настроек.
- Config census и registry: `docs/NDM_Audit03_Config_Coverage.md`, JSON/CSV в `docs/regression/Audit03`; пользовательская книга прочитана без Excel, исходный SHA сохранен. Структурное наличие поля не объявлено поведенческим покрытием. Первый медленный census остановлен; индексированный v2 успешно завершен.
- A03 mutation reproducer: `a03_snapshot_negative_v19_2026-10-02.txt`, 114/33 на прежнем production-коде; `a03_snapshot_positive_v19_2026-10-02.txt`, 147/0 после защиты. Семнадцать отдельных попыток изменения nested results и повторный реальный Execute. Early v1/v2 timeout был ошибкой нового fixture (пропущенный обязательный аргумент), v17/v18 имели неверное ожидание COM-обертки ошибки; эти логи не являются PASS.
- Full Off v19 прерван старым test-only повтором уже опубликованного LC; full Off v20 завершен с четырьмя failure нового stability fixture (значения Auto/Both вместо допустимых AutoWithL/BothPlanes). Исправлены только сценарии тестов и сохранение Err до cleanup; защита результата и физические expected не ослаблены. Reset теперь проверяется на реально заполненных, еще не опубликованных Width/Stability, а не записью в публичные поля.
- Full Off/On v21: `a03_full_off_v21_2026-10-02.txt` 8340/0 и `a03_full_on_v21_2026-10-02.txt` 8341/0 по всем восьми suites; source unchanged=True. Это приемка A03 publication среза, не завершение Audit03.
- P01/P04: `performance_pre_p02_v2_2026-10-02.txt` и `performance_p02_v3_2026-10-02.txt`: каждый содержит 150 измерений (15 сценариев x 5 повторов x 2 версии), все baseline assertions green, источники неизменны. В v2 время write ошибочно суммировало root и дочерние writer-ы; только этот показатель v2 не используется. В v3 измеряется root writer, сохранены размеры модели и точный scalar-key повторных solve. Asymmetric: 94 solves, 3874 iterations, 20 retries, 25 несошедшихся probes, 0 неподтвержденных physical outcomes, 10 полностью эквивалентных solve-attempts; причины повторов еще разбираются.
- P02 lifecycle: до и после valid-invalid-valid Rounded/Hollow проходят 27/0. `p02_point_grid_comparison_v22_2026-10-02.json` и v24 сохранили расхождение одного ряда на нижней границе; они не объявлены PASS. Контурный диагностический вывод v23 случайно менял историю компиляции, поэтому совпадение v23 само по себе не считалось приемкой.
- Причина граничного различия установлена отдельным опытом: VBE отображает длинный GEOM_PI как 3.14159265358979 и при перекомпиляции использует уже округленный текст. В v25 обе книги получили deltaFrom3*1e15=141592653589790 и один результат границы; после импорта неизмененного полного исходного литерала v26 обе получили 141592653589793 и другой одинаковый результат. `p02_point_grid_comparison_v26_2026-10-02.json`: 3626 точек, 0 различий. Это дефект воспроизводимости литерала, не изменение ContainsPoint математикой кеша.
- GEOM_PI теперь записан суммой двух коротких литералов, точно равной прежнему полному Double; число сегментов, шаги, subdivisions и point-in-contour не изменены. После повторного импорта теста и save/reopen v27: 28/0, exact GEOM_PI=4*Atn(1), все 3626 точек совпали с исходной полной константой v26. Контрольная pre-P02 книга получила ту же точную константу только для сопоставимого benchmark, сами geometry algorithms там прежние.
- Full P02 Off v22: 8367/0. Full Off/On v27: 8368/0 и 8369/0, восемь suites, source unchanged=True; source contracts 101/101, 0 ошибок. Численные expected и tolerance прежние. Один дополнительный assert v27 проверяет точность pi после VBE canonicalization. Финальная приемка Audit03 не заявлена.
- P02 v4: 150 измерений, все assertions green и sources unchanged=True. Стабильная точная pi одинакова; Hollow mesh имеет 3600 волокон в обеих версиях и прежние A/центры/I. Медиана полного времени 9.8740 -> 0.1582 с, geometry 9.7764 -> 0.0615 с. Asymmetric до P01 по-прежнему 94 solve/3874 iterations; вариация времени не выдается за численную оптимизацию.
- P01: десять Asymmetric дубликатов являются одинаковыми direct active-set retries после возврата к высокой lambda. Основной solve с новым соседним warm start сохраняется. CStateSolutionRunner теперь держит ограниченные 64 scalar-записи только numerical failures одной capacity-search session; другой start/options/context/revision и вытесненная попытка допускают новый solve. Успешные/terminal results не кешируются, shared repository не участвует, при Clear capacity session уничтожается. P03 не форматирует residual retries при отключенной диагностике.
- `p01_retry_session_v29_2026-10-02.txt`: 22/0, recovery/новый старт/options/context/terminal/bounded memory/eviction. `p01_full_off_v29_2026-10-02.txt`: 8390/0; On v29: 8391/0; все восемь suites, source unchanged=True. `source_contracts_v29_2026-10-02.txt`: 101/101, 0 ошибок; export SHA `5F9FE986813AD49E83ABD22FE6BD6419712F186642D49D0C5DDBD4DCD6BFA920`. Производительный gate после P01/P03 еще выполняется, общий Audit03 не завершен.
- Code census: 101 модуль, 4024 метода, 1205 guard-кандидатов. Это индекс для review, не семантический PASS; часть отсутствующих непосредственных комментариев относится к разрешенным тривиальным свойствам/групповым wrappers. D01/F07 еще требуют содержательного просмотра.
- UI environment: native accessibility вернул null, native screenshot завершился timeout. VBE selection прочитан через специальный COM API для диагностики compile error. Нативная визуальная приемка не объявляется выполненной; это не мешает Excel COM/DisplayFormat/save-reopen проверкам.
- P01/P03 `performance_p01_p03_v6_2026-10-02.txt`: 160 измерений, 16 cases x 5 x 2, failed=0, sources unchanged=True. Asymmetric solve 94 -> 84, iterations 3874 -> 3864, exact duplicates 10 -> 0, unconfirmedPhysical=0; медиана 15.2002 -> 15.1133 с. Существенное численное ускорение не заявлено. Diagnostics Off: formatting calls 150 -> 0. P02 Hollow mesh: 9.7852 -> 0.1367 с с одинаковыми 3600 волокнами. v5 sharing conflict test report сохранен как неуспешный запуск, не PASS; logger исправлен и весь набор повторен.
- F08: `docs/NDM_Audit03_Prior_Audit_Traceability.md` содержит per-ID Audit01/Audit02 владельцев, test-ID и свежие логи. Runtime evidence не подменяет Pending новые K/T/D/UI/physical migration/save-reopen gates.
- F07/K02 новый negative `f07_k02_input_negative_v30_2026-10-02.txt`: 22/14 на unchanged v29 production, только test-module import. Подтверждены неверный Boolean/молчаливый default, округление дробного Long, raw overflow, scalar/narrow/Nothing Settings Range, scalar Profiles Range, Nothing/tiny/product mesh и молчаливая нормализация subdivisions=0. Исправления еще не реализованы; подробный контракт в `docs/NDM_Audit03_Input_Contracts.md`. Новый test module source пока отдельный незавершенный срез, не входит в green v29 acceptance.
- F07/K02 v31 negative расширен до 22/17 на трех настоящих consumers. После минимальных guards v33 directed 54/0; full Off 8444/0, On 8445/0, source unchanged=True. Пустой/TODO Mesh.BoundarySubdivisions не заменяется 1, большой счетчик подъячеек отклоняется до арифметики массива. Все F07/K02 требования целиком еще не закрыты.
- D02: генератор справки больше не утверждает остановку всех расчетов после stability FAIL. Утвержденное продолжение уже покрыто TestStabilityFailContinuesDownstream; actual rebuilt/update help еще предстоит проверить. Добавлен контракт допустимых subdivisions, без изменения геометрической методики.
- T04 all-path Light v35: 441 случай, 12687/135; отрицательный лог сохранен. Подтвержден физически недопустимый UltimateStrain final при OK в двух My-путях; финализация теперь проверяет StateWithinPhysicalRange после ConfirmEquilibrium, иначе общий Newton уточняет ту же точку. PhysicalBoundary v36: 18 случаев, 538/0, save/reopen True, source unchanged=True. Новый тест отдельно трактует явный путь с нулевой масштабируемой компонентой и блокировку Width из-за Formation; production-контракты этих ветвей не изменены. Полная повторная нагрузочная приемка продолжается.
- D01: уточнены комментарии к существующим AutoCAD export методам и устаревшие архитектурные подписи status formatting/batch. Расчет и построение не изменены; текущая изолированная книга еще не включает последний comment-only срез. Census v36: 101 модуль, 4043 метода, 1227 guard-кандидатов; это индекс, не семантический PASS.
- T04 v36 RoundedSimple Off: Light 441 независимый путевой случай, 12954/0; Stress 216 случаев, 8491/0. Во всех прогонах проверены comments/actual writers/report и save/reopen, source unchanged=True. CSV-реестр `load_matrix_summary_2026-10-02_cases.csv` содержит внутренние статусы/коды/flags/leaf и output comments; отдельный runs manifest сохраняет отрицательные/незавершенные версии как NotPassed. Повторные progress-копии логов не считаются отдельными кейсами.
- Full v37 Off 8444/0, On 8445/0, все восемь suites, source unchanged=True. `source_contracts_v37_2026-10-02.txt`: 101/101, 0 ошибок, export SHA `750EA6FC744A97415DBFB6CC47379EB61ABD532B90D85A97E0915F5CEAD2E086`. В книге текущий production/test source, включая comment-only срез; пользовательская output-книга имеет прежний SHA `AAF5D4FAF06197F01966463DE85BA813EB4216B455AA8504B2DE4F0BA717C012`. Эти gates подтверждают текущий checkpoint, а не финальное завершение Audit03.
- FormationBoundary v38 после исправления HasLimitPoint: Off 9 случаев, 354/0; On 9 случаев, 370/0. CircleSym Stress On: 216 случаев, 8289/0. Full v38 Off 8444/0, On 8445/0. Несимметричный круг также завершил Light/Stress Off/On; сохранение Results и неизменность источников подтверждены. Неуспешный extended PreCrackState остается диагностикой со своей причиной, а не физической точкой.
- RectRectangle Light v38: 441 случай, 13039/88. Все ошибки относились к неверному fixture: ApplyLoadReference получил абсолютные (10; -7), тогда как независимое сравнение ожидало бетонный центр + (10; -7). Исправлена только постановка теста, добавлены проверки LoadReferenceOffsetX/Y. Повтор v39: 441 случай, 12996/0; save/reopen/source unchanged=True. Production-координаты и допуски не менялись.
- UI01 v39: `palette_off_v39_2026-10-02.txt` и `palette_on_v39_2026-10-02.txt`, по 351/0. Проверены Interior и DisplayFormat, реальная легенда и четыре writer-а, переход FAIL -> OK -> N/A, очистка прежних строк при уменьшении пакета. Направленная CF rule доказала чувствительность теста к переопределению цвета. Values и status-style SHA до/после повторного открытия равны; пиксельная визуальная приемка этим не заявляется.
- K02 profiles v40 negative: unchanged production + только test module, 32/40. Все четыре Calculation.* переключателя молча принимали Maybe/TODO/пустую/ошибочную ячейку и отсутствующую строку как No. Catalog теперь сохраняет восемь Boolean aliases, но возвращает адресную ошибку 3988 с профилем и key для неверного обязательного значения. Directed positive: 72/0; full Off/On v40 еще выполняются. Это input-contract исправление, без новой методики или классов.
- Full v40 Off 8516/0, On 8517/0; все восемь suites, source unchanged=True. Source contracts v40: 101/101, 0 ошибок. Эти gates подтверждают profile input и предыдущий срез, не весь Audit03.
- D02 update v41: 777 полей ввода сохранили значения/формулы/validation/форматирование, но ссылка Config P36 осталась на прежней строке A693 вместо нового заголовка A697. Повторное обновление не распознавало название с уже добавленным `(Подробнее)`. Исправлен только генератор ссылок. `help_updated_v42_2026-10-02.txt`: 141 ссылка корректна, seven-status dictionary и измененные контракты совпали; input SHA до/после/save-reopen одинаков, failed=0. Это целевой gate обновленной справки, не полная нормативная или пиксельная приемка.
- Пользовательское требование всех путей уточнено в широком runner-е: кроме пяти физических Capacity paths добавлен отдельный выбор Auto; Formation по-прежнему имеет четыре варианта. Независимый test oracle проверяет выбор по пользовательским N/M и отсутствие поиска без нагрузки. Новый десятивариантный runner еще требует runtime gate; прежние девятипутевые PASS-логи сохраняют свое исходное покрытие.
- CircleSym Light v42 Off: 490 независимых путевых случаев, 14608/0. Отдельно проверен Auto Capacity, включая нулевые/tiny нагрузки; все comments/4 writers/actual execution report и Results save-reopen совпали, source unchanged=True.
- Чистая сборка v42: `RC_Section_NDM_clean_v42.xlsm`, build exit=0. Full clean Off 8516/0, On 8517/0; source contracts 101/101, 0 ошибок, export SHA `616B047AB3C6962967E3773F26EF8C1CF0B01D71BEAC59B7F30F5DD9A3480212`. Пользовательская output-книга и Audit03 ТЗ сохраняют исходные SHA.
- Clean palette v42 Off/On: по 351/0; Values/Interior/DisplayFormat и статусное оформление после save-reopen сохранены. Style SHA `1EABA40A72D5A28392F11FC5D38F0D2B1E9B264FDFAD9F5514ACB9EE1A1E1132`. Пиксельная приемка не заявлена.
- Полное OpenXML-сравнение clean/update справки v42: 2473 непустые ячейки совпали точно, объединения совпали, обе книги неизменны, failed=0. Negative сравнение с прежней справкой обнаруживает смещение/изменение 2574 адресов и объединений; это не 2574 независимых дефекта и не PASS. Нормативный смысл текста отдельно не подтверждается этим сравнением.
- Load CSV обновлен: 21 исторический прогон, 5368 записей, 16 PASS runs; отрицательные и незавершенные версии остаются NotPassed. Содержательная частичная F07-ревизия и оставшиеся numeric/array/L15 вопросы сохранены в `NDM_Audit03_Guard_Review.md`.
- Checkpoint `705e2a8`: physical Formation point, profile/help и clean v42 Off/On/source/palette gates, десять вариантов путей. Это проверенный промежуточный срез, не завершение Audit03.
- RectL Stress v42 On: 240 независимых случаев всех десяти путевых вариантов, 9269/0; actual comments/writers/report/save-reopen проверены, source unchanged=True.
- K02 numeric exploratory v43 содержал неверные ожидания Excel-error кода и поведения Batch.ApplySettings; его 49/63 не является числом подтвержденных дефектов. Исправленный negative v44: 82/60; positive на том же тесте: 142/0. Пустые/TODO Solver/Capacity значения теперь отклоняются адресно, Batch не запускает solve и возвращает InputErr; optional диаметры и отсутствующие optional API keys сохранены.
- Расширенный numeric negative v45: 394/296 по всем 56 маршрутам численных Solver/Capacity ключей четырех consumers. Clean v45 с тем же тестом и обновленной справкой: positive 690/0; full Off 9206/0, On 9207/0. Source contracts 101/101, 0 ошибок, export SHA `EC327B2C65CA0996078CAF1CC3DF4818B5DBC65E8340ED959B488914A7842BCD`. Sources unchanged=True; per-key активные эффекты/вся приемка K02 не завершены.
- Validate clean v45 exit=0. Help update v45: 777 input fields/формулы/validation/формат сохранены, 141 ссылка корректна; save/reopen signature неизменен, failed=0. OpenXML compare clean/update: 2476 непустых ячеек и объединения совпали, обе книги неизменны, failed=0. Helper Compare исполняется в PowerShell 7; ошибочный запуск в 5.1 не является дефектом расчетной книги.
- Последовательный broad runner подготовлен для 13 форм x Light/Stress x Off/On. Resume проверен на завершенном 490-case отчете и отклоняет отчет другой source-книги; отрицательные/неполные логи не перезаписываются. Все 52 новых прогона еще не выполнены.
- Checkpoint этого проверенного среза: `731a0108` (`Audit03 strict numeric input and complete load-path runner`). Последовательная 52-run матрица v45 запущена; источник не изменяется. Завершены первые 28 прогонов, текущие результаты находятся в `broad_matrix_runner_v45_2026-10-02.log`; это промежуточный счет, не финальная приемка.
- После checkpoint изменены комментарии production/test методов, добавлены направленные F07 тесты массивов диаграмм и экстремальных Double-входов State. В source добавлены guards против повторной ошибочной инициализации материала и различение численного переполнения/деления на ноль от внутренних runtime-ошибок. Эти правки еще не имеют runtime acceptance и не входят в неизменную v45-книгу.
- Содержательное чтение реальных ResultComment выявило сырой `ConcreteStrainLimit` в success leaf, `NumericalFailure:` в отказе и тройной повтор общей причины CurrentCrackedState в aggregate. Новый тест читает все непустые комментарии, а отдельный composition-test сохраняет typed исходы/leaf и требует единственную общую причину. Source-исправления подготовлены; negative/positive и full gates еще предстоят. Подробности: `NDM_Audit03_ResultComment_Review.md`.

## Как Продолжить

0. Текущий исходник в `RC_Section_NDM_source_v48.xlsm`, SHA `650C8C68236B12A850320A560120932F90425D7B948A7E7D2A49ECCDA4713DBC`. Полные восемь suites v48: Off 9656/0, On 9657/0, оба watchdog завершены и source unchanged=True. Source contracts v48b: 101/101, failed=0; актуальный export SHA `40FB2684D8827AFE6FFF7CD546C1B2FE2CD7A95047E091E2A38291DB05DB109E`. Первый v48 contract обнаружил отставший export test-модуля; read-only экспорт из книги восстановил совпадение без изменения source SHA. Пользовательская output-книга не изменена.
   v45 матрица завершена: 52/52 runs, 18980 параметризованных путевых случаев. v46 повтор остановлен на четвертом run из-за четырех ошибок comment-oracle, не production; исправленный CircleSym/Stress/On v47: 10414/0, 240 cases, save/reopen/source gates успешны. Отрицательные логи сохранены. Полная матрица текущего v48 еще нужна.
   Новый directed extension gate: 4 формы x 48 известные равновесия = 192/192, 32 Off отказа и 8 нагрузок за технической возможностью; 2008/0 assertions. Реальный v47 Secant NumFail исправлен сохранением улучшенного failed retry-start в CStateSolutionRunner, без новой физики/допусков/repository-кэша. Подробности в `NDM_Audit03_Extended_Equilibrium_Review.md`.
   Solver setting effects v47: 84/0; mutation Method v48b: 78/6 ожидаемых отказов на отдельной книге. Все 13 Solver keys имеют active evidence; четырнадцатая строка fixture была General.DiagramExtension, не Solver key. Per-key JSON/CSV v48 сохраняет 1065 адресов и 13 active-reviewed полей без blanket PASS; остальные контракты и negative unit-sign еще не завершены.
1. Полный 52-run повтор v48 завершен: 18980 независимых путевых случаев, все failed=0; source unchanged=True. CSV `load_matrix_summary_v48_2026-10-02_*` содержит также исторические версии, поэтому для текущего v48 выбирать только его 52 runs. Checkpoint проверенного среза: `2597d293`. Согласованность вывода не доказывает, что каждый NumFail неизбежен.
   Negative v49: unit-sign Solver effects 88/12 (четыре отрицательных допуска скрывались Abs после CUnitSystem); actual L15 73/12 (psi=1 верно, но ResultComment не объяснял неположительное среднее); RoundedTapered Newton 171/3, Newton8 171/3, Secant 173/2. Диагностика подтвердила rank-deficient tangent: активные стержни на одной линии, остальные на плато. Исправлены знак четырех параметров у всех четырех consumers, сброс кривизн самостоятельного осевого fallback и поздний старт через оценку усилий арматуры/обратную фактическую диаграмму. Физика, методы и допуски не изменены. v50 unit effects 100/0, L15 85/0; v51 тот же RoundedTapered Newton 177/0 (48 случаев, 16 независимо невозможных). SLS axialT0.95/1 физически допустимы при своих сопротивлениях; Strength выходит за физические пределы и остается FAIL.
   Техническая точка +/-10 не является hard cap solver-а: EvaluateAtStrain за ней сохраняет последнее напряжение и нулевую касательную. Поэтому равновесие с отдельными крайними волокнами за +/-10 законно как вспомогательное FAIL, не физический OK. Новый ошибочный blanket endpoint assert v52 выявил именно это (CircleSym axialC100, два отказа oracle), не новый production-дефект; отрицательный лог сохранен. v53 проверяет физический FAIL таких состояний, а известные 192 плоскости по-прежнему должны находиться внутри технических точек.
   Directed Stress v53 завершен: 26 runs (13 форм x Newton/Secant), семь отрицательных runs; кроме modest Secant/Imported отказов оставалась непроверенная невозможность RectL/Mixed100. Source v53 SHA `01F68B4CA3F1631FD1D5A09E2B2D0FDB5E2A75CFC78ABD452DB0F0EBD675E7B0`. Known-state v53: четыре runs по 502/0, все 192 известные равновесия найдены; full Off/On также завершены без failures. Это не закрывает новые направленные отрицательные примеры.
   v54: стабилизированный active-set уточняет только новый inverse-steel-force старт, а не меняет старые попытки. ImportedFixture/Cracked/axialT1.05 стал сходиться; семь из десяти problem runs остаются отрицательными. v55 кандидат с весами Et*As ухудшил RoundedTapered/Newton/Cracked/axialT0.95, поэтому отменен вручную; v55 остается отрицательным доказательством и не принят. В v56 возвращены веса As, расширен независимый dual-certificate направлениями фактической failed-плоскости и добавлена раздельная диагностика холодных стартов Newton/Secant. Production статусы не назначаются по сертификату теста.
   v56 завершен: семь отрицательных problem runs. Кандидат v57 размерно согласует Broyden по геометрическим плечам, без изменения strain-field/материалов/tolerances и без перехода Secant в Newton. Три прежних modest Secant случая проходят; остаются RectTwoLeft/Secant/Strength/T1.05, ImportedFixture/Newton/Cracked/T1 и RectL/Mixed100. Это пока directed evidence кандидата, не финальная numerical acceptance.
   v58 невалиден: новый тест использовал отсутствующий MinDouble; compile error показан пользователем и подтвержден исходником. Серия остановлена, только ее parent/child runner и automation Excel закрыты; фактический пользовательский Excel с Давление.xlsx (PID 13684 на момент проверки) не тронут. Ни v58 directed, ни его full suites не засчитываются. В v59 MinDouble заменен простым If, короткий runtime gate успешно компилируется, но Imported/Newton остается отрицательным (218/1).
   v59 точные границы замещения steel-concrete по общим деформациям дают независимый сертификат невозможности RectL/Mixed100 для обеих ролей: оба Newton/Secant directed runs 221/0. Прежняя независимая разность крайних напряжений была слишком консервативна. Production статусы по сертификату не назначаются. RectTwoLeft/Secant с 20 явно заданными рестартами проходит 225/0; default=2 не отменяется и не скрывается. Imported/Newton500 все еще 218/1, то есть простое увеличение итераций не устраняет rank-deficient старт.
   v59 завершен: full Off 9672/0 и On 9678/0; source unchanged=True, SHA `3E0951B865373777537DA7CF4A6151DA4EF697256F39EDCA70D2F82C6AA3DAF7`. Эти gates принимают срез с размерно согласованным Secant, но не решают Imported/Newton. Поздний neutral-line старт v60 не устранил этот отказ: Imported/Newton 218/1. Unit-sign consumer тест v60/v60b остановлен watchdog на 30/120 секундах, причина пока не установлена, PASS не заявлен. В v61 добавлены test-only отметки Capacity/Formation/Batch и независимая диагностика масштабов neutral-line старта. Все прежние exec sessions завершены; новые COM проверки выполнять строго последовательно. Пользовательская output-книга не изменена. Все K/T/D/final pending остаются открытыми.
2. Завершить F03/F04/F05/F06/F07 call-site acceptance (включая конечность Ultimate line-search), A03-A05 all-class/consumer/snapshot аудит и фактический Config census. Нагрузочный 27-LC matrix не заменяет полную L01-L17/shape/setting приемку, report и save/reopen.
3. До завершения всех шести gates цель остается активной; новые ограничения/изменения после green suite записывать сюда.

## Последний Проверенный Срез v66

- Сессия 80149 полностью завершена, SERIES_FAILURES=0. Все 26 directed Stress
  runs (13 форм x Newton/Secant, 1248 задач) прошли: 404 задачи с невозможным
  равновесием имеют независимый stress certificate, 844 состояния найдены. Production
  не назначает статус по сертификату теста. Четыре known-state runs по 502/0
  повторили 192/192 известных состояний, 32 парных Off отказа и 8 технически
  невозможных задач.
- Все восемь suites v66: Off 9716/0, On 9722/0, source unchanged=True.
  SHA книги `4B09D90316A1EF3443E098C7FEF49428C780384E5BE3158E7C090CDA0F60BA90`.
  Imported/Newton найден обычным CStateProvider с late compression start;
  RectTwoLeft/Secant проходит с default=2. Физика и допуски не менялись.
- Unit-sign fixture v61 имела пропущенный обязательный аргумент AddCombination,
  v62 не инициализировала CUnitSystem, v63 считала внешние search probes вместо
  внутренних итераций solver-а. Эти ошибки теста устранены. Directed v64 44/0
  включен в full v66. Negative v67 на unchanged production v48: 20/24,
  подтверждены все четыре неверных параметра у Capacity, Formation и Batch.
- Новый material Config test в существующем bas: 23 входа, 16 спецификаций,
  active/inactive, неверный ввод и recovery. v67 имеет ошибку компиляции
  из-за имени scale; этот run не принят. После переименования в metricScale
  v68 directed 1442/0, все восемь suites Off 11158/0, On 11164/0.
- Source v68 SHA `6A17F72B1EB21173CB9A4BC635C38DDA1AA0E63B457CE40891D142DD6A49EB9F`.
  Source contracts 101/101, failed=0; export SHA
  `FAAA000B9C9E74A3B980022A1FC0D4BA1D732574B80674BCC30AF841C818A8CB`.
  Реестр v68 сохраняет 1065 адресов, 36 имеют active acceptance с незакрытым
  полным диапазоном. Остальные 1029 не объявлены принятыми.
- Пользовательская output-книга и Audit03 ТЗ не изменены. K/T/D/final остаются
  открытыми. Новый код после v66 не считать принятым по историческому full gate.
- Командные Excel-runner-ы запускать через Windows PowerShell 5.1, read-only
  census и evidence merger через PowerShell 7. Пробный запуск watchdog из PS7
  не прошел из-за оболочки/PSModulePath; расширение harness не принято и удалено.
  Это не ошибка НДМ и не PASS. Все перечисленные VBA gates завершены в штатной
  оболочке. Negative evidence merger отклоняет направленный green лог вместо
  полного набора восьми suites; число 36 не подменяет все 1065 полей.

## Продолжение После Checkpoint 4ddda47

- Проверенный срез v68 зафиксирован локально в `4ddda47`
  (`Audit03 robust equilibrium starts and material Config evidence`).
  Baseline не меняется. Push, destructive Git и откаты пользовательских
  изменений не выполнялись.
- Последовательная broad matrix v68 завершена и сессия 22975 закрыта:
  52/52 runs, 18980 независимых путевых случаев, все watchdog exit=0,
  sourceUnchanged=True. Источник и пользовательская output-книга разделены.
  Эти результаты не являются приемкой последующих Config/UI правок.
- D01: уточняются комментарии о действующих material-role API, сохраненных
  точках Search и чтении snapshot. Удалены исторические обещания и пояснения
  отсутствующего будущего слоя. Расчетный код этими правками не меняется;
  экспорт v68 остается доказательством checkpoint, не новых комментариев.
- K02: подтвержден source-дефект `General.WorstCombinationCriterion`: любое
  неизвестное значение молча выбирало StrengthCapacity. Подготовлены strict
  проверка в Batch и тест четырех настоящих критериев с independent typed-leaf
  oracle, фактическим writer, invalid input и recovery. Runtime еще не выполнен.
- K03: исходная книга содержит 19 объединений, но среди 1065 адресов реестра
  нет merged followers. Подозрение на исходный двойной учет этим опровергнуто.
  Подтвержден другой дефект: 16 вторых dropdown-ячеек RectSet игнорируются,
  поскольку положение/привязка дополнительного ряда общие для пары сторон.
  Решение сохраняет общую расчетную семантику: объединение двух строк,
  единственный selector, строгий конфликт у reader-а для необъединенного ввода.
  Migration helper сначала проверяет все пары, не меняет конфликтующую книгу,
  пишет прежние 16 значений в отчет; остальные 761 input-поле сохраняются
  строго, включая формулы, validation и форматирование. Пользовательская
  output-книга не обновлялась. Новые source/UI/help тесты пока не приняты runtime.
- Следующий COM gate выполнять только после завершения сессии 22975:
  negative на unchanged v68 с импортом только существующих test bas; затем
  isolated v69 с текущими source, обновленной справкой и общими RectSet
  селекторами, directed positive, все восемь suites Off/On и source contracts.
  Broad v68 не является приемкой последующих input/UI/comment правок.
- D01: найдены и содержательно заменены еще 11 однотипных комментариев
  `Validate*`/`Clear` у геометрии, импорта, Capacity и workbook-сценария.
  Они описывают именно проверяемые размеры/активные ряды, typed причину,
  очищаемые данные и границы ответственности. Census дополнен этими шаблонами
  и отделяет простые accessor-ы от review-кандидатов; это не автоматический
  семантический PASS всех 101 модулей. Новые комментарии войдут в v69.

## Единицы, Знаки И Фактическое Оформление

- По отдельному запросу пользователя проверка Config включает все двенадцать
  INPUT/OUTPUT селекторов единиц и три правила знаков N/Mx/My. В существующем
  test bas добавлены independent conversion oracle, все допустимые варианты,
  явный неверный ввод, отсутствие обязательной строки и recovery. Подготовлен
  сквозной тест 72 физических эквивалентов (8 знаков x 3 Force x 3 Moment),
  с обычным reader-ом, solver-ами, всеми named-state и фактическим Results.
  Перевод OUTPUT отдельно проверяет отсутствие дополнительного solve.
- Negative `unit_sign_choices_negative_v69_2026-10-02.txt`: 273/90. Пустой
  выбор, пробел и TODO скрывались default во всех 15 селекторах. Расширенный
  `unit_sign_choices_missing_negative_v69b_2026-10-02.txt`: 282/126, дополнительно
  подтвержден default при повреждении/отсутствии девяти строк таблиц.
  Исправлены GetRawString в CUnitSystem и обязательность пользовательских
  unit/sign keys в LoadFromWorkbook; автономный LoadFromRange не превращается
  в полную книгу. Коэффициенты, физика, допустимые значения и знаки не менялись.
  Positive v71: choices 408/0; equivalence 72 физических варианта, 1440/0.
  Values Results совпали после save/reopen. Полные Off/On v71: 13300/0 и
  13306/0, восемь suites, sources unchanged=True. Input Length/Stress/Curvature
  варианты проверены в адаптере, но их сквозные geometry/material сценарии
  этим тестом не объявляются полностью закрытыми; Area.Input требует отдельной
  K03 трассировки, так как пользовательского ввода площади в текущем flow нет.
- Worst criterion negative v69: 25/45; RectSet common selector negative v69:
  160/64. Неуспешные логи сохранены, а не объявлены PASS.
- Первая проверка сохранности обновления справки v69 и диагностический v70
  выявили только четыре изменения horizontal alignment у Config!K118,
  K120, K122, K124: слева -> центр. Значения, формулы, validation и number
  format этих полей не менялись. Исходные negative logs сохранены.
  Уточнена явная область ожидаемого изменения: центрирование только 16 общих
  selector anchors, без исключения их данных из строгой проверки. 16 followers
  выводятся из editable роли с отдельным отчетом прежних значений; остальные
  761 input fields остаются в сравнении до/после/save-reopen.
- Isolated source v71 содержит все новые VBA правки и принят текущим full
  Off/On gate. Help/common-selector update: failed=0, 141 ссылка, 761 input
  fields сохранены до/после/save-reopen. Worst criterion positive 70/0;
  RectSet shared selectors positive 224/0. Source contracts: 101/101, failed=0;
  версионный export SHA `92317AA79A1ABE8CBFC54446ED287B81770B81F93987438B4565CCC4266E13FB`.
  Source v71 SHA `7A3F2198E271DCCECE63B34A2E2BA67DAA72FCB57B39F98DDA11F3EF06EE3E20`.
  Пользовательская output-книга не изменялась, Audit03 цель остается активной.

### Выявленные Отклонения Оформления

- Config!K118, K120, K122, K124: пользовательские селекторы привязки
  дополнительных рядов RectSet были выровнены слева, вопреки центрированию
  INPUT по AGENTS.md. В migration helper исправлено на центр по обеим осям;
  runtime/save-reopen приемка выполняется отдельно.
- Фактический OpenXML census v71 обнаружил еще 27 input-полей с выравниванием
  слева/general вместо центра (vertical уже center): RoundedRectangle L146:L149
  (диаметры третьего ряда); HollowRectangle H159/K159 (размеры), K163:K166
  (число стержней), K173:K180 (привязка дополнительных рядов); RectSet K105
  (B2), K108:K115 (число стержней). Они включены в список изменений немедленно
  по запросу пользователя. Исправление и COM/save-reopen доказательство еще
  предстоят; это не скрывается общим PASS справки v71.
- По запросу пользователя каждое следующее подтвержденное отклонение
  оформления сразу включать в этот список. Проверить фактические editable
  ячейки, заголовки Значение, короткие Ед./Справка/INTERNAL и выравнивание
  комментариев с учетом наличия столбцов справа; не выдавать census за PASS.
- Дополнительно сразу включены в UI-правки: I99 (шапка Значение), J99
  (шапка Ед.) и L139 (Ед. основного армирования RoundedRectangle) не
  центрированы; шапки комментариев O107/P117/M139/P145/M162/P172 прижаты
  вправо, тогда как другие шапки комментариев читаются слева. Общие
  комментарии составной геометрии должны выравниваться по границе своей
  подтаблицы, а не всего именованного диапазона. Точный перечень тела
  комментариев будет записан независимым COM-тестом до/после/save-reopen.
- Причина 27 input-отклонений: правило комментария из верхней подтаблицы
  проходило до конца составного диапазона, перекрывая выравнивание полей
  ниже. Также обнаружены смещенные offsets HollowRectangle: common-input
  строки и строка шести размеров. В SettingsCatalog исправлены границы
  подтаблиц и адреса; числа, формулы, validation и расчетная методика не меняются.
  Этот новый formatting-срез пока не имеет runtime приемки.
- Независимый COM negative v71 проверил 999 адресов и подтвердил 52
  отклонения: 27 input-полей, девять ячеек Ед. (J99:J101, L105, L139:L143),
  шапка I99, девять комментариев K100:K101/K155:K156/L130/L134:L137
  и шесть шапок комментариев. Полный поадресный список
  `config_formatting_negative_v71_2026-10-02.txt` сохраняется как negative,
  source SHA не изменился. Форматный positive будет выполняться на копии v72.
- Первый positive v72 не принят: все 999 проверяемых адресов выровнены верно,
  но строгий snapshot указал еще четыре изменения Config!K167:K170.
  Диагностический v72b подтвердил, что менялось только выравнивание прочерков
  автоматического n у Opening. Это не пользовательские вводы, но служебные
  маркеры количества тоже должны быть центрированы. Они сразу включены в
  список правок и независимый ожидаемый scope; данные, validation, number
  format, merge, заливка и шрифт остались прежними. Оба неуспешных лога сохранены.
- Formatting positive v73: 3015/0, 1003 адреса, ноль отклонений после первого
  и повторного применения и save/reopen. Всего исправлено 56 H/V отклонений;
  данные всего Config и все 1065 зарегистрированных адресов сохранены,
  включая формулы, validation, number format, merge, заливки и шрифты.
  Source v73 SHA `3207167AF3C216CEDF73E2A3C7A5CE1E2DF6A30D14E77C07B8FD59DAD8ED4B30`.
  Чистая сборка v73 запущена отдельно для проверки тех же правил с нуля.
  Это COM/save-reopen приемка выравнивания, не пиксельная приемка всех листов.
- Чистая сборка v73 завершена exit=0. Независимый Verify без мутаций:
  `config_formatting_clean_v73_2026-10-02.txt`, 1004/0, 1003 адреса,
  исходник неизменен. SHA clean book
  `B09CF8BCFC65EB6897DE1D0A88A0EAD8A68BDD4531F2775AAB443B8246CECFCE`.
  Full clean Off/On завершены: 13300/0 и 13306/0, восемь suites,
  sources unchanged=True. Штатный Validate-Workbook завершился exit=0.
  Source contracts clean v73: 101/101, failed=0, тот же export SHA v71.
  Полные clean/update листы справки совпали: 2492 непустых ячейки,
  одинаковые объединения, sources unchanged=True, failed=0.
  Версионный VBA export v71 уже сохранен отдельно и не перезаписывается.
- Важно для hash-трассировки: штатный Validate-Workbook открывает ZIP в Update
  и при наличии удаляет неканоническое дублирующее имя Print_Area из
  workbook.xml перед read-only Excel-проверкой. Это не read-only процедура
  относительно контейнера книги. После нее SHA clean v73 стал
  `2D3B067BF4A7FA1CE2A9B9DE69D5361BCBDC17F14EE29466EE5166F231B7D22D`.
  Off/On выше относятся к предшествующему SHA B09C..., а не выдаются за
  побайтовый gate этого нового файла. Расчетный source/VBA и Config не
  изменяются этой функцией; post-validation format gate: 1004/0,
  1003 адреса, ноль отклонений, source unchanged=True. В текущем workbook.xml
  неканонических Print_Area имен нет. Сам факт смены SHA не доказывает, что
  такое имя действительно присутствовало до открытия ZIP в Update.
  Полная финальная приемка должна идти после всех таких cleanup-операций.
- Нагрузочный v68 manifest повторно сверен по всем 52 исходным txt-отчетам:
  каждый watchdog завершен, source hash совпал, save/reopen=True, cases=18980.
  Сохранены v68-only runs/cases и acceptance JSON. Отдельный общий runner-log
  v68 не найден и не указывается как существующий артефакт; исторический v48
  runner-log относится к другому source и не используется как v68 evidence.
- Перед checkpoint нет незавершенных COM/test sessions. Output SHA повторно
  прочитан с ReadWrite sharing (книга сейчас занята другим процессом): тот же
  исходный `AAF5D4FAF06197F01966463DE85BA813EB4216B455AA8504B2DE4F0BA717C012`.
  Ее не закрывать и не обновлять без безопасной финальной процедуры.
  Следующий объем: INPUT Length/Stress/Curvature через реальные geometry/
  material consumers, K03 Area.Input; остальные per-key и pairwise Config,
  D01/F07 semantic review и финальные benchmarks/output/self-audit.
  Audit03 остается активным, финальный отчет и общий DoD не заявлены.

## Продолжение После Checkpoint cadcb986

- Текущий принятый units/RectSet/formatting срез зафиксирован в `cadcb986`.
  Tracked Git tree и index после проверки чистые; исторические registry
  файлы сохранены побайтово. Baseline и пользовательская output-книга не менялись.
- По запросу пользователя все подтвержденные отклонения оформления сразу
  заносить в список выше. Новые единичные успешные тесты не подменяют полный
  Audit03 DoD и не означают обновление занятой пользовательской книги.
- Подготовлен сквозной `TestAudit03InputUnitConsumers`: 4 формы x 3 INPUT
  длины x 5 INPUT напряжений x 2 INPUT кривизны = 120 вариантов, по 4 LC.
  Запускается настоящий workbook-сценарий со strength direct/capacity,
  formation/width/longitudinal и stability. Независимые коэффициенты
  пересчитывают исходные размерные ячейки; сравниваются геометрия, named-state,
  диаграммы и все значения/ResultComment трех подробных блоков.
  Runtime пока не выполнен; новый тест не объявляется принятым.
- K03 `Units.Area.Input`: подтверждены только adapter API и фиксированный
  mm2 маршрут AutoCAD. Сквозной пользовательский потребитель площади не найден.
  Не придумывать новую физику для активации этой настройки; решение об актуальном
  интерфейсе, справке и migration предстоит в отдельной K03 проверке.
- Первый input consumer прогон v74 завершен: 120 случаев, 1269/24,
  source unchanged=True, Results save/reopen=True. Все 24 отказа относятся
  только к Transformed.Ixy около нуля у Rounded/Hollow в kgf/cm2 и tf/m2.
  Разность 1.5e-8..7.9e-8 мм4 возникает при round-trip модулей и составляет
  машинное округление относительно Ix/Iy; НДС, диаграммы и подробные
  strength/crack/stability числа, статусы и ResultComment совпали во всех случаях.
  Лог сохранен как неуспешный, не переименовывается в PASS.
  В новом тесте добавлен ограниченный roundoff oracle 8*2^-52*sqrt(Ix*Iy)
  только для Ixy/Ixyc. Остальные сравнения и все исторические expected/tolerance,
  solver settings и production source не меняются. При расхождении логируются
  оба фактических значения и bound. Повторный runtime gate еще предстоит.
- Уточнение пользователя: AutoCAD export всегда в мм независимо от OUTPUT
  Excel. Реальный export reader уже восстанавливает мм/мм2/мм4 из сохраненных
  заголовков Results. Добавлена проверка всех координат/габаритов/диаметров/
  площадей/локальных инерций и отсутствия solve в 72 OUTPUT/sign вариантах;
  runtime еще не объявляется выполненным. Фактическая запись DWG этим не заменяется.
- INPUT consumer v75 завершен: 120 вариантов, 1293/0, source unchanged=True,
  Results save/reopen=True. Это эквивалентность фактически активных стадий:
  некоторые normal-width/PostCrackState ветви неприменимы (в частности последние
  Hollow LC имеют NotCracked), а не 120 независимых активных Width-проверок.
  Source SHA `8C6FC7F79386A3B6BB2DA04DC77F8B6768A8ED8B1CE5AD1EBE8EBD16D5D96741`.
  Полный Off gate v75 еще выполняется; AutoCAD assertions пока не приняты.
- В сохраненном Results v75 обнаружен сырой `SP35-eta` в комментарии успешной
  устойчивости. `CStabilityCalculator.ResultMeta` возвращает пустой комментарий,
  а `CStabilityResult.InitializeFromCalculator` подставляет машинный Branch.
  Сразу включено в список исправлений пользовательского текста: calculator
  должен формировать русскую причину собственного OK/FAIL, result сохранять ее
  без подстановки Branch, writers не формировать объяснение. Требуется отдельный
  negative/positive gate для СП 63 и eta/table/mixed СП 35, включая output-блоки.
- Первый stability-comments запуск v76 отклонен до runtime: в новом test
  пропущен обязательный comment аргумент трех AddCombination. Это ошибка
  подготовки теста, не доказательство дефекта production; лог не перезаписывать.
  Вызовы исправлены, воспроизводящий negative gate запускается заново как v76b.
- Negative v76b: 90/10, 12 LC, source unchanged/Results reopen=True. После
  правки calculator-а positive v77: 100/0, те же четыре ветви/LC, без State solve,
  комментарии leaf/подробного блока/batch согласованы, Results reopen=True.
- Actual help v77: правило экспорта геометрии в мм добавлено в Units и
  AutoCAD.Export.CombinationID; два новых assertions успешны до/после reopen,
  141 прямая ссылка корректна, все 761 input-поле сохранено, failed=0.
- Общий Validate v77 выявил потерю области печати после help-update. Это
  подтвержденное отклонение оформления сразу включено в список изменений.
  Help updater должен сохранять/восстанавливать существующие Print_Area через
  уже общий SettingsCatalog API, как Refresh-VbaModules, и проверять snapshot.
  Первый Validate остается отрицательным, полный gate еще не запущен.
- Help updater после исправления сохраняет область печати (count=1), все вводы
  и прямые ссылки; повторный Validate успешен. Source contracts v77: 101/101,
  failed=0, export SHA `066479801FCBE981520BA450BB3F77E090BD5E302D4935A43D3BB52B813C242E`.
  Новый full v77 запуск не начал suites: runner отклонил VerifyResultsReopen
  без явно выбранной одной macro. Это неверные аргументы запуска, не runtime
  дефект; negative log сохраняется, полный повтор идет как v77b без этого флага.
- OpenXML-чтение actual v77 справки подтвердило места текста: B655 под
  заголовком "Единицы измерения" и B1380 в "AutoCAD.Export.CombinationID".
  Geometry export возвращает мм, а OUTPUT-единицы расчетных подписей остаются
  отдельным контрактом. Старый output и Audit03 ТЗ сохраняют исходные SHA.
  Полная v75 Off серия: 14737/0, включая 72 geometry-mm и 72 no-solve assertions.
  v77b full Off/On еще не объявлены завершенными.
- Full Off v77b завершен: 14837/0, восемь suites, source unchanged=True.
  On gate еще выполняется. Подготовленный INPUT evidence ограничен двумя
  активными полями Length/Stress: их обрыв меняет геометрию/диаграммы.
  Curvature INPUT входит во все 120 equivalence-вариантов и validation,
  но заданный clamp 0.00005 не доказан как binding. Один успешный одинаковый
  ответ не закрывает K02 этого поля; нужен отдельный binding-clamp маршрут.
  Существующий active MaxDeltaKappa gate не подменяет такую combined-проверку.
- Full On v77b завершен: 14843/0, восемь suites, source unchanged=True.
  Оба full gate относятся к SHA v77
  `01FFA526A60791E70EB1B5C5B252A14D83D7A548D9E079E3A9A16D1D6A9B2333`,
  после help/Print_Area/Validate операций, а не к более раннему контейнеру.
  INPUT evidence merge: 1065 адресов, 66 active-reviewed, fullAcceptance=False.
  Дополнительные 85 numeric Results ячеек последнего mixed snapshot совпали;
  A4 elapsed исключен явно, исходный mismatch JSON сохранен.
- Post-update format v77: 1004/0, 1003 адреса, source unchanged=True, тот же
  SHA 01FFA... . Все нужные COM/test sessions завершены; основной output
  по-прежнему AAF5... и не заменен промежуточной книгой. Версионный экспорт
  `VBA_All_Code_v77_2026-10-02.txt` сохранен отдельно, SHA 066479... .
  Этот срез готов к scoped checkpoint; Audit03/Final Report/общий DoD не завершены.
  Следующий объем: binding INPUT Curvature gate, K03 INPUT Area, остальные
  per-key/pairwise и semantic D01/F07; затем финальные clean build, benchmarks,
  сохранение пользовательских данных при output update и независимый self-audit.
  После compaction читать этот хвост вместе с Git status/diff/log, baseline,
  ТЗ и архитектурными MD; старый раздел "Как продолжить" не подменяет новые gates.
- Checkpoint `2d656722`: INPUT Length/Stress, AutoCAD geometry-mm contract,
  русские комментарии устойчивости, actual help/Print_Area и v77b full Off/On.
  Коммит подтвержден через git log; tracked dirty после него отсутствовал.
- INPUT Curvature v78: отдельный binding gate 132/0, source unchanged=True,
  Results save/reopen=True. По двум осям Newton/Secant ограничитель 1e-7 1/мм
  и 1e-4 1/м дает одинаковые 22/7 итераций вместо свободных 2; отключение
  адаптера в отдельном test-call наблюдаемо снимает ограничение. Это не новая
  production методика и не подгонка solver tolerance. Full Off v78: 14969/0;
  On еще выполняется. Source SHA 67399F22383CC706EF95F92C96E671E7B62030584005E1AA4134484C85511A98.
- Уточнение K03 INPUT Area: прежняя формулировка об отсутствии workbook
  consumer была неполной. Реальный путь ImportGeometryFromAutoCADForWorkbook
  передает Config units в ImportFromActiveDocument, где INPUT Area применяется
  к AutoCAD.Import.MinArea. Геометрия самих Region остается в мм/мм2/мм4.
  Поэтому поле не удаляется; справку фиксированного мм2 порога нужно исправить.
  Подготовка импорта выделена в метод того же importer-а, который используется
  live wrapper-ом и тестом существующего CFakeAcadRegion. Новых классов нет.
  Добавлен обязательный/неотрицательный MinArea, адресная ошибка overflow,
  nine-case equivalence/boundary и invalid/missing tests; runtime еще впереди.
- Full On v78: все семь первых suites, включая UI 5010/0, завершили свои
  assertions, но runner после UI упал с Null method invocation до восьмой
  suite. Это не full PASS. Отрицательный лог сохраняется. Runner получает
  точный ScriptStackTrace/PositionMessage и guards недоступного Mode Range;
  повтор v78b нужен до принятия Curvature evidence и дальнейших COM запусков.
- Повтор Full On v78b завершен: 14975/0, восемь suites, source unchanged=True.
  Null post-UI сбой не воспроизвелся; его причина не выдается за исправленный
  production дефект. Добавленная диагностика остается в runner-е для точной
  локализации при повторении. Curvature evidence теперь имеет Off/On gates.
- Новое требование пользователя: проверить import -> смена INPUT/OUTPUT ->
  расчет и единицы исходного import snapshot. Добавлен nine-transition тест
  mm/cm/m с INPUT Area/Force/Moment/Stress/Curvature и пользовательскими знаками,
  независимым переводом чисел и сравнением НДС/геометрии/свойств Results.
  Fake Region идут через общий импорт, реальный preview writer и кнопку расчета;
  это не реальный DWG smoke. Несуществующие import layers и огромный MinArea
  после snapshot должны доказать отсутствие повторного импорта/фильтрации.
  Отдельно проверяются непереведенное число нагрузки и invalid OUTPUT с
  сохранностью старой геометрии/recovery. Runtime еще предстоит.
- v79 actual help: 761 input-значение/формула/validation сохранены, 141 ссылка,
  Print_Area count=1, failed=0; новые import/MinArea тексты проверены на листе.
  Validate v79 успешен. INPUT Area directed: 89/0, source unchanged=True,
  Results save/reopen=True. Порог включительно сохраняет Region ровно на границе.
- Import-unit workflow v79 отрицательный: 2/1, ошибка только тестового поиска
  заголовка X вместо X, mm (индекс 0). В новой v80 используется существующий
  ResultHeaderColumnByBaseName. v80: 221/7, геометрия всех девяти переходов,
  отсутствие reimport, invalid OUTPUT/recovery и save/reopen уже проходят.
  Три пары numerical snapshot не совпали: helper пересчета INPUT повторно
  устанавливал ZeroMomentPerDepth, в tf*m зануляя тестовый момент. Новый вызов
  явно отключает этот фильтр после helper. Отдельный axial strainRatio пока
  не принят; добавлена диагностика epsilon0/kappa/Nint. Отрицательные книги
  и журналы не перезаписываются; требуется новый v81 runtime и full gates.
- v81 directed import-unit workflow: 228/0, source unchanged=True,
  Results save/reopen SHA `41F8614F9929EC76D4E5944B93FA62F3353DF17A8C1CA82F071E5C31B752DA58`.
  После исключения фильтра оба numerical mismatches устранены без изменений
  production математики: чисто осевые epsilon0 -6.62038e-8/-6.62038e-5,
  отношение 1000, почти нулевые kappaX/Y и правильный Nint. Предыдущие
  negative tests выявляли плохую настройку самого сценария, не unit defect.
  Source contracts v81: 101/101, failed=0; raw export SHA
  `5C71C72EF47490F65872F221ABE8C5DE93382E31F23490E73772655550DED41C`.
  Full Off выполняется; per-key evidence для Curvature и Area/MinArea создан,
  но merge не запускается до обоих full gates.
- Full Off v81: 15286/0, восемь suites, source unchanged=True; SHA
  `051FA4BE33BC899504D5A88CBD315F6DBE66C9807A40E5987293A6FD8CECE263`.
  Full On v81 не принят: UI assertions 5327/0 завершены, затем runner упал
  на строке 44 `$Book.Names.Item(...).RefersToRange` до восьмой suite.
  Новый stack точно локализует COM-сбой, не numerical failure расчета.
  Get-ModeSettingCell теперь читает тот же именованный диапазон через
  Config.Range с проверками отсутствующего листа/range; не подставляет mode
  и не пропускает post-suite контроль. Полный повтор еще нужен.
- Последнее дополнение пользователя: контракт исходных AutoCAD Region
  добавлен в справку AutoCAD.Import.ConcreteLayer: X/Y/размеры мм, Area мм2,
  Ix/Iy/Ixy мм4, независимо от INPUT/OUTPUT, без определения масштаба INSUNITS.
  Отдельно объяснены MinArea INPUT и геометрический snapshot OUTPUT.
  Неточное описание "площади в миллиметрах" в Geometry.Source исправлено.
  v82 содержит предыдущую справку, v83 получает окончательный текст; его
  actual help/format/full gates еще предстоят. Основной output не изменен.
- v83 actual help завершен: 761 input-records/validation/formulas сохранены,
  141 ссылка, Print_Area count=1, failed=0. OpenXML подтвердил окончательный
  контракт Region в Справка!B1488, отсутствие INSUNITS scaling в B1489.
  Validate успешен; source contracts 101/101, failed=0; export SHA
  `49FDEF50E4A1DEAE662869B4A4F24904F61FC2A1885C7554807B85CD0D64E3C8`.
  Книга SHA `DC646D3C622F95E1B696390C4EDFEBFD4E30E59E62DFC0B47BD75DC0EC6EC4DE`.
  Directed import-unit On повторно 228/0, save/reopen=True. Полный On теперь
  выполняется с Range-based mode-reader, после него требуется full Off и
  actual formatting/evidence merge. Output и ТЗ сохраняют исходные SHA.
- Full On v83 снова отрицательный: все UI 5327/0 закончены, но теперь Null
  на `$Book.Worksheets.Item('Config')`, так что замена Names на Range не
  устранила сбой. Предыдущая гипотеза об одной коллекции Names не подтверждена.
  Runner после VBA Run получает ту же открытую книгу из Excel.Workbooks,
  строго сверяет FullName и ReadOnly и логирует sameProxy. Это не reopen,
  не потеря dirty данных и не mode fallback. Полный v83b On запущен; до
  его завершения и полного Off merge/commit этого среза не делать.
- Full v83b On/Off завершены: 15292/0 и 15286/0, все восемь suites,
  source unchanged=True. COM_WORKBOOK_AFTER_SUITE: sameProxy=True и ReadOnly=True
  во всех восьми шагах; это успешный повтор, а не доказательство внутренней
  причины прежнего transient COM Null. Negative v78/v81/v83 журналы сохранены.
  Evidence merge v83: 1065 адресов, active-reviewed=69, fullAcceptance=False.
  Добавлены только Units.Curvature.Input, Units.Area.Input и AutoCAD.Import.MinArea.
  Основной output остается AAF5... и не заменен промежуточным срезом.
- Actual formatting v83: 1004/0, 1003 адреса, deviations=0, тот же SHA DC646...;
  все нужные test/COM sessions завершены, процессов EXCEL не осталось.
  Версионный raw export `VBA_All_Code_v83_2026-10-02.txt` сохранен отдельно.
  Срез готов к scoped checkpoint. Audit03 и Final Report не завершены;
  следующая работа - остальные per-key/active-inactive/pairwise, semantic D01,
  F07 реальные array callers и L15/L16 independent proof, затем final clean
  build/benchmarks, output update с сохранением пользовательских данных и
  итоговый self-audit. Не подменять это зеленым import-unit срезом.
- Crack Config v84-v87: подтверждено, что все 12 SLS.Crack полей ранее
  могли подменять missing/blank/TODO значениями классов. Корректный negative
  v84b: 746/60; v84 дополнительно содержал 25 ошибок самого unit-fixture
  (721/85). Batch, Formation и Width теперь используют обязательные getters;
  начальные значения программно создаваемых классов не изменены. Формулы,
  физические критерии, expected и tolerance не менялись.
- Новый directed Crack Config тест: 31 валидный параметризованный запуск и
  75 invalid/missing запусков на трех нагрузочных постановках, включая
  ступенчатый контур. Все 12 ключей проходят invalid/missing/recovery,
  десять имеют отдельный active proof. Нельзя выдавать это за проверку всех
  форм/профилей или активность двух Formation-селекторов. v85: 994/2 из-за
  слишком большого тестового выступа, исключавшего арматуру из зоны;
  отрицательный лог сохранен. После исправления только fixture v86 и v87:
  2726/0, независимая численная формула, реальные writers и ResultComment.
  v87 Results save/reopen SHA `3E9F5EFC6437BD20AA01048EC2CDA84B95A94B1FF3101356018C5559E1097A34`.
- Full Off v86/v86b останавливались после Batch 6211/0 на null Workbooks.
  Диагностический v86c доказал: тот же Excel был Ready=True, raw Workbooks
  имел Count=1, книга оставалась с правильным FullName и ReadOnly=True,
  raw Worksheets имел Count=4. Прямое чтение COM-свойств в runner-е заменяет
  неверно возвращавшую null PowerShell-обертку; нет reopen/retry/default.
  Это доказательство текущего сбоя, а не утверждение о причинах всех прежних
  transient failures. Отрицательные журналы v86/v86b/v86c сохраняются.
- Full Off v86d завершил все восемь suites без COM-сбоя, но UI дал 5239/88:
  минимальный Formation Config старого numeric-fixture не содержал двух
  обязательных селекторов. В v87 они явно заданы; проверяемый Solver-key
  по-прежнему удаляется в optionalAbsent. Направленный повтор numeric input:
  690/0. Ни один expected error code или допуск не ослаблен.
- v87 actual help: 1925 строк, 141 ссылка, 118 shapes, сохранены 761 input
  records/формулы/validation и Print_Area count=1, failed=0. Контракт Region
  мм/мм2/мм4 без INSUNITS scaling присутствует и проходит actual-sheet gate.
  Validate успешен. Source contracts: 101/101, failed=0, 83 + 3 classes.
  Raw export SHA `2681F25F692E3A257782A1AB5AE049B5AE600C922D825515F8F21281006A1A22`;
  книга SHA `11152F2A2C7984697DF4F061C45CB13999E85AA57C4DCEA597FCC3111BB6CED4`.
  Полный Off v87 выполняется; полный On и formatting/merge еще нужны.
- Новый read-only census v87: 101 module, 4142 methods, 1378 guards;
  semantic acceptance=Pending. 971 метода без собственной подписи требуют
  контекстной проверки групп простых свойств, а не 971 автоматической правки.
  D01 worklist: шаблонная первая строка ClearGeneratedShapes в CSectionPlotter
  (2166); историческое объяснение ClearLegacyWorksheetGeneratedShapes (2178)
  проверить и описать текущую очистку принадлежащих схеме объектов. Код этих
  методов в этом срезе не менялся. Census запускать через PowerShell 7:
  WindowsPowerShell 5.1 не прочитал UTF-8 без BOM и дал parser error; это не
  ошибка VBA. Основной output и пользовательское ТЗ сохраняют исходные SHA.
- Новый запрос пользователя: обязательный ввод должен давать понятные причины,
  место и действие для исправления. Reader сохраняет фактические адреса значений
  обычных таблиц, единиц/знаков и material aliases; для удаленной строки адрес
  не придумывает, называет раздел и восстановление строки. Материалы и адаптер
  единиц добавляют это место к собственной валидации, не меняя машинные коды.
  Ошибки SLS-селекторов/диапазонов указывают Config, имя строки и действие.
  Ошибки отсутствующих обязательных named tables отдельно объясняют восстановление
  имени в Excel. Нет нового класса, назначения статусов по тексту или default.
- Full Off v87 завершил все восемь suites, failed=0: Geometry 524, Material 2856,
  Solver 895, Capacity 1676, Crack 484, Batch 6211, UI 5327, Baseline 39.
  Runner/исходная книга сохранены, watchdog exit=0. On v87 не запускался:
  новый запрос требует финального повтора на более свежем v89.
- Input messages v88 directed: 1897/4. Все четыре failures - ошибочный адрес
  batch-комментария в новой проверке (Offset 1 вместо принятого Offset 12),
  а итоговое workbook-сообщение уже содержало SLS.Crack.Allowable, Config B75
  и действие. Исправлен только test offset; отрицательный лог/книга сохранены.
  Первый source-contract gate был запущен до окончания обновления export и
  получил 93/101, 8 mismatches именно в измененных модулях; не считается PASS.
  v89 собран из v88 с исправленным test; актуализация help и повтор source
  contracts выполняются. Directed/full Off+On, formatting и evidence merge
  остаются обязательными до приемки/коммита этого среза.
- Directed messages v89 и v90: 1901/0, watchdog exit=0, source unchanged=True.
  Workbook message и batch Results содержат SLS.Crack.Allowable, Config B75,
  причину и действие. Проверены 32 системных ключа, 23 material aliases и
  выбор единиц/знаков; количество assertions не равно количеству настроек.
  Source contracts v89 и v90: 101/101, failed=0. Финальный VBE export v90 SHA
  `8338B482D03496F6189796C52C3C391CF9A391808C48FB66CE451D5E80FF54CF`.
- Отрицательный help v89: failed=2, новый gate читал только B и пропускал
  общие абзацы, объединенные от A. Проверка теперь читает A и B; содержимое
  справки не удалено ради PASS. v89b и v90 failed=0. В v90 1930 строк, 141
  ссылка, 118 shapes, 761 input records и Print_Area count=1 сохранены.
  Добавлено различие между оформленным InputErr в Results и ранней остановкой
  до записи нового snapshot. Поздние изменения диапазонов Stability/MinArea
  не меняют критерии и error codes; их дополнительные assertions входят в
  штатные suites. Полные Off/On v90 и formatting еще выполняются/ожидаются.
  Изолированная v90 SHA `4EF49CC1A19EBC3D3E5E289C562ABE8E75803BB84115F8D2EEC24EBEBE48CFF8`.
- Crack directed v90/v91 завершились отрицательным watchdog timeout до
  решения НДС. В v91 последний этап `baseline.read`: новая сигнатура
  Validate(settings) материалов не учитывала программные Initialize/Clone,
  вызывавшие Validate без аргумента. Это внесенный дефект компиляции,
  а не numerical failure. Диагностический reader теперь необязателен;
  физическая проверка и машинные коды сохраняются для обоих способов создания.
  Directed messages 1901/0 не проверяли эти программные маршруты и не
  являются доказательством полной компилируемости. Отрицательные артефакты
  сохранены, окончательные directed/full gates повторяются на новом v92.
- v92 Material suite 2856/0 подтвердила исправление Initialize/Clone.
  В v93 добавлены четыре направленных случая поврежденных имен обязательных
  таблиц: rngUnitSettings, rngSignConventionSettings, rngSystemSettings,
  rngPlotAnnotationSettings. Каждое сообщение называет Config, таблицу/имя,
  восстановление ссылки в диспетчере имен; после восстановления reader работает.
  Directed messages v93: 1921/0, watchdog exit=0, source unchanged=True.
  Directed Crack Config v93: 2837/0, Results save/reopen SHA
  `6971B8B12E62C112D76A80F5D28FC75CE8BAF7085F8889B322B21ED9C5FF67C1`.
  Source contracts v93: 101/101, failed=0; export SHA
  `A83FCDC2BB759DFA56420A680D1E6499618D1088E313DCD4B45248D006838BF1`.
  Full Off/On и formatting еще не приняты; evidence merge не запускать раньше.
  D01 worklist дополнен: у NormalizePlaneMoments в CStabilityCalculator
  первая строка комментария относится к другому методу; RequirePositive
  арматуры содержит историческое слово "прежний" вместо текущего контракта.
  Эти comment-only правки относятся к очередной семантической D01-проверке,
  не к выпуску до приемки всех требований Audit03.
- Полный Off v93 завершен: 19650/0, восемь suites, source unchanged=True.
  Этот прогон не заменяет повтор позднего исправления пересчета единиц на v95.
  Направленный negative v94: 1950/36. Семь настроек возвращали только Overflow
  без имени/ячейки/действия; восьмой случай Calculation.ZeroMomentPerDepth
  прерывался при добавлении LC из-за фильтра с уже невалидными настройками.
  Исправление: ApplySettings сохраняет ключ текущего пересчета и по машинному
  Err.Number=6 формирует понятную причину; NormalizeCombinationMoments не
  обрабатывает LC при mSettingsValid=False. Арифметика CUnitSystem, критерии
  и допуски не менялись, новая методика/классы не вводились.
- Directed messages v95: 1993/0. Все восемь конечных Double, переполняющих
  пересчет INPUT, дают InputErr с собственным адресом и действием, без solve;
  после восстановления тот же batch исполняется. Кнопка расчета показывает
  ошибки независимо от NonCriticalMessages (проверено по entrypoint и
  фактически возвращенному workbook-сообщению, не по screenshot MsgBox).
  Crack directed v95: 2837/0, Results save/reopen SHA
  `3E9F5EFC6437BD20AA01048EC2CDA84B95A94B1FF3101356018C5559E1097A34`.
  Help v95: 1932 строки, 141 ссылка, 118 shapes, сохранены 761 input records
  и Print_Area count=1, failed=0. Source contracts 101/101, failed=0.
  Export SHA `B74AEF7FCE42D45F413983229A34F051C1E4572DFD0555A496E89D3E85B3E5B3`;
  книга SHA `585719A6972C8E1D49AED0463562A98B862A88151F7D85ECE33FC4CA890B75F4`.
  Полные Off/On, formatting и merge v95 на этом шаге еще выполнялись/ожидались;
  итоговая приемка этого среза приведена ниже.
- Окончательные full v95 Off/On завершены: 19722/0 и 19728/0. Все восемь
  suites прошли: Geometry 524, Material 2856, Solver 895, Capacity 1676,
  Crack 484, Batch 6324/6330, UI 6924, Baseline 39. Watchdog exit=0,
  source unchanged=True; исторические explicit-On setup в Off сохранены.
  Formatting: 1004/0, 1003 адреса, deviations=0, без исправлений оформления.
  Validate: все 25 структурных проверок True, фактическая книга не изменена.
  В справке фактической v95 A22/A23 присутствует полный текст о переполнении
  INPUT и исправлении указанной ячейки; это OpenXML-проверка содержимого,
  не пиксельная приемка интерфейса или MsgBox.
- Per-key evidence десяти Crack-настроек присоединен только после направленного
  gate, двух полных suites и formatting. Реестр v95: 1065 адресов,
  activeReviewed=79, fullAcceptance=False. Два Formation-селектора не получают
  новую активную приемку; full-range, interactions и оставшиеся K/T/D/F07/P
  требования по-прежнему открыты. Merge запускать PowerShell 7: проба через
  Windows PowerShell 5.1 не прочитала UTF-8 без BOM и дала parser error;
  повтор в PowerShell 7 успешен, VBA/расчет к этому сбою не относится.
  Проверенный checkpoint содержит source, изолированную v95-книгу, export,
  положительные и отрицательные доказательства. Основной output не опубликован
  до полного DoD Audit03 и сохраняет SHA
  `AAF5D4FAF06197F01966463DE85BA813EB4216B455AA8504B2DE4F0BA717C012`.
  Пользовательское ТЗ также не изменено: SHA
  `F3B32621FD3D6BDFF0313B76599EDD5D93F637971A5F91818902FF4418118E98`.
  `git diff --cached --check` для source/tests/tools и вручную правленных MD/JSON
  успешен. Общий check также видит конечные пробелы и пустые строки в неизмененных
  raw VBE exports и runtime logs; эти доказательства не форматировались задним
  числом ради зеленой проверки. Их байтовые SHA и отрицательные исходы сохранены.
- Search Config, рабочий срез v96-v98: отдельный standard test module, без
  новых classes, подключен к полному Batch suite. Проверяет девять пар
  Capacity strategy/method, двенадцать Formation path/strategy, настоящий
  reader/Config, физические точки, три компоненты равновесия, четыре writer
  comments, 20 ошибочных вводов и recovery. Пока acceptance не присвоен.
  Первый setup v96 не нашел profile key: таблица профилей использует key во
  втором столбце. v96b/v96c остановились на ссылках теста на чужие private
  helpers; через read-only Excel VBE API установлена точная строка компиляции,
  собственные helpers добавлены. Это дефекты нового test setup, не NumFail ядра.
  Полные отрицательные журналы сохранены.
- v96d: 762/38 до ошибочного вывода неинициализированного тестового unit
  adapter. Подтверждены blank/TODO fallback Capacity.SolutionStrategy и
  недостаточная навигация Unknown; последующие invalid cases не приняты.
  Brent для N=100000 Н, Mx=4000000 Н*мм, My=2000000 Н*мм при lambda tolerance
  0.0001 выдавал OK/HasLimitPoint с physical=False, ratio=1.00046210519761.
  Исправление сохраняет обе solver-точки скобки, передает их доменному
  FinalizeBracket и использует именно переданное допустимое НДС, без re-solve.
- v97: 1442/8, все 20 invalid/recovery cases выполнены и дают InputErr с
  понятным местом и действием, без solve. Три новых ratio assertions были
  строже точности собственного lambda-fixture; для девяти точных сравнений
  этот fixture теперь задает 0.000001. Отдельный Brent-контрпример сохраняет
  прежние нагрузки и 0.0001: запрет extended/нефизической точки не ослаблен.
  Малая нагрузка v96 не требовала PostCrackState; невырожденный активный
  Formation fixture v97 использует N=130000 Н, Mx=5200000, My=2600000 Н*мм.
- v97 также воспроизвел NumFail фиксированного Mxy/UltimateStrain: line-search
  вес критерия 10 не учитывал его действующий допуск 1e-9 и отвергал полный
  Newton-шаг к трещине из-за промежуточной невязки N. Merit function v98
  использует отношение уже принятых допусков равновесия/критерия; сами
  conditions/tolerances и материальные пределы не изменены. Отрицательные
  пробы/диагностика сохранены; positive, full suites и повтор матрицы еще
  необходимы. v97 source/export equality 102/102, failed=0, не закрывает runtime.
- v98: 1479/1; все пути/стратегии Formation получили физическую точку.
  Новый точный ratio oracle был строже force/moment tolerance собственного
  fixture. Для девяти точных Capacity сравнений fixture v99 использует
  N tolerance=0.01 Н и M tolerance=1 Н*мм; исторические expected/tolerance
  и отдельный Brent-контрпример не изменены.
- v99: 1479/1. Secant выдавал физически недопустимую конечную точку при
  ratio=1.00000003003019. Capacity FinalizeAt теперь проверяет physical range
  и может отклонить кандидата; общий Secant учитывает Boolean callback-а,
  продолжает уточнение и финализирует retained bracket без чужого LastSolver.
  v100: 1792/0, включая 13 cross-section isolation вариантов и два baseline.
  Results save/reopen SHA `434F226C4944C475A09274185BF6554A4A45020910F0E7BD31D2982A056752F9`.
- Directed v101: 1792/0, Results save/reopen SHA
  `66EE7587ADE8C90DDF1653D49F4B8E923180418A043AAD2C257D308167069484`.
  Source/export equality: 102/102, failed=0, export SHA
  `4A54B5A63AE69D30817CAFA3BC061EA044C14E3459AE77A66DB7B2E2EDEF7C04`.
  Help: 1933 строки, 141 ссылка, 118 shapes, input/Print_Area сохранены.
  Full Off v101 завершен с шестью regression failures в прежних Brent tests:
  точный корень финализировался как сторона скобки. Они не скрыты и ожидаемые
  lambda/tolerance не изменены; общий Brent снова сначала вызывает FinalizeAt
  для точного корня, затем отдельно FinalizeBracket для узкой скобки.
- Address negative v102: 1810/2 на v101 production + только усиленный тест.
  Обе ошибки относятся к неизвестному Formation path/strategy без адреса.
  Новые нормализаторы Formation сохраняют aliases и автономный контракт,
  optional reader только добавляет адресную диагностику. Batch использует их
  до любого solve; нет нового класса, словаря статусов или формулы.
  Watchdog negative exit=1, source unchanged=True. Положительные/final gates
  на v102 еще выполняются, per-key registry пока остается 79/1065.
