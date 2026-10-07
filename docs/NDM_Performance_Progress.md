# NDM Performance Progress

Дата начала: 07.10.2026. Цель: выполнение `NDM_Performance_Optimization_Spec.md`.
Статус: цель завершена; проверенные book/TXT опубликованы, итог в Final Report.

## Baseline И Сохранность

- Git baseline: `df3d5eb787e95246b2b01dfb19bdc76ed82d98f8`.
- До правок production/tests/tools чистые. Единственное tracked-изменение пользователя:
  `workbook/output/RC_Section_NDM_execution_report.txt`; оно не включается в наши коммиты.
- Книга: `D48A5FFCC17A3FC4F1F2A6FC7080952BD01BA08FCB7151CFBF15EA6E887EDB25`.
- VBA TXT: `E20F213475B2DF0592E4CEECABA4E706C7B1AADBB822D23FD6A3E0D197B5EEFE`.
- Актуальный пользовательский report: `D65B41D99B4E8969F5ADBB7196C9FB41989DEBB1287827F4F18D5D17B98ECCBC`.
- Побайтовые копии книги/TXT/report и исходников находятся в
  `regression/Performance/Baseline/`; полный SHA manifest в `Manifest.json`.
- Baseline содержит 97 production-компонентов, 88 production-классов;
  вместе с tests 120 исходных компонентов и 93 класса. Новые классы запрещены.
- Текущие книге/экспорту ничего не записывается до публикации проверенного кандидата.
  Проверки выполняются последовательно в собственных копиях/процессах Excel.

Прочитаны AGENTS, Performance Spec, Audit03 Performance/Final Report,
PostAudit03 Final Report, уточнения NeighborRatio и IndivisibleGroups СП35.
Ранее выполненные material/retry/state/mesh-capacity и passive-snapshot
оптимизации не считаются новыми изменениями этой цели.

## Порядок Работ

1. P04/P06: ограниченное блочное чтение, operation-local reuse, емкость массивов;
   компиляция, адресные edge/lifecycle tests и A/B.
2. P01/P02: точные повторы Secant/Ultimate/Newton и разделение K/физических данных;
   узлы диаграмм, финальная приемка, отрицательные cache-контексты и exact A/B.
3. P03/B01-B04: ленивая общая подготовка Query и mechanical properties у прежних
   владельцев, изоляция LC и освобождение временных объектов.
4. P05: измерить отдельно форматирование; принимать только доказанно полезные
   изменения с восстановлением поврежденного оформления.
5. B05: пакеты 1/10/30, направленная регрессия, save/reopen, память, выпуск и self-audit.

## Владельцы И Границы Подготовки

| ID | Владелец / Содержимое | Точный Контекст | Сброс И Память | Проверка |
|---|---|---|---|---|
| P04 | Reader схемы / уже прочитанные Results-таблицы | Одна операция, фактическая книга/именованный якорь | Очистка в начале, при успехе/ошибке Load; не между кнопками. Только таблицы одного snapshot | blank/формула/пробел/error, 1 ячейка, край листа, перенос якоря, новое чтение после изменения |
| P04 | Export / bounded scan и таблицы текущего ReadResultsExportState | identity книги, фактический именованный якорь одной операции | Блок 1024 ячейки; Value2 таблицы только внутри подготовки одного состояния. Успех/ошибка очищают dictionary и ссылку книги до рисования; между кнопками данные не живут | Config после solve, LC/State/RunID, повтор export, broken snapshot |
| P06 | Model / резерв параллельных массивов | Тот же model, фактический Count отдельно от capacity | Clear стирает массивы/capacity; конец жизни модели освобождает память. Рост менее двух емкостей либо точный Reserve | границы роста, invalid add, Clear/rebuild, IDs и точные поля |
| P06 | Plot reader / резерв элементов | Только один Load, порядок snapshot | Reset уничтожает массивы/capacity; хвост не публикуется | большой/малый/пустой/ошибка, реальные Count и прежние getters |
| P01/P02 | Solver / пригодная численная оценка | Точная плоскость, identity/revision модели, фактические frozen materials, stress API/mode, маска forces/extrema/K | Только одна задача/попытка; принятую точку не заменяет Jacobian probe. Небольшие scalar/K данные; без массива всех LC | knots/окрестности, новая цель и ConfirmEquilibrium, changed context, Secant/Broyden |
| P03/B01 | Query / общая бетонная область | identity/revision Model, включая Contours.Revision; материал не нужен | Один расчетный запуск, ленивая подготовка; clone для LC. Не создавать цикл Model->Query->Model | contours-only mutation, independent opening, автономный вызов, новый похожий model |
| P03/B01 | Properties / mechanical A/I/оси/проекции | identity/revision Model, фактические Eb/Es и направление/точка/представление | Один запуск; отдельные LC-dependent Ncr/eta/design. Без контурного пересчета mechanical A/I | provider/moduli смена, направления, свежий объект, Stability On/Off/On |

Перед реализацией конкретного кэша его точные поля и reset дополняются здесь.
Защитные copies meta/material/region остаются независимыми. Физические решения
между LC не кэшируются; repository каждого LC сохраняет собственную область жизни.

Фактические ключи реализации: solver хранит две отдельные точные оценки
forces-only/physical (physical имеет маску K), identity frozen диаграмм и
identity/Revision модели; принятую Ultimate-оценку хранит local RunNewton,
с шестью точными Offset/Base и Revision. Query проверяет Model identity/Revision.
Properties хранит представление, фактические Eb/Es и до 16 точных запросов
оси/точки/IncludeRebar; сверх емкости выполняется обычный расчет без расширения.
Batch владеет только ссылками текущего Execute; `mExecuting` запрещает
восстанавливать объектный кэш при последующем справочном чтении writer-ом.
Счетчики обнуляются на каждом Execute; scalar-центры и опубликованные results
сохраняют прежний срок жизни. Ни Model, ни result не ссылаются обратно на Batch.

## Фиксированный Набор Пакетов

Индексы `j=1..30`, меньшие пакеты всегда первые 1/10 строк. Нагрузки ниже во
внутренних Н и Н*мм; fixtures используют прежние builders/профили, одинаковые A/B.

- DIRECT: N=(-1)^j*(100000+11000*j), Mx=(j-15)*1000000,
  My=((7*j Mod 31)-15)*700000; контрольные строки осевые/одноосные явно фиксируются в fixture.
- CAPACITY: разные N/M и шесть LoadPath, Ultimate/Multiplier и Newton/Secant
  в отдельных конфигурациях; начало с короткого pilot для оценки длительности.
- CRACK: отдельные SP35/SP63, реальные NotCracked/Cracked и разные профили;
  нормальная нагрузка и прежний сохраненный Hollow пример.
- STABILITY: отдельные SP35/SP63, разные полные/длительные нагрузки и отключенные профили.
- MIXED: разные допустимые профили, tiny/near-limit/overload, ошибочная строка
  не последняя; проверяются следующие LC и governing tie.
- REPEAT: отдельная проверка повторов/порядка/Execute, не основной benchmark.

Фактические 30 строк, геометрии и параметры сохраняются машинно перед A/B.
Обязательны normal/reverse/fixed-shuffle, смены context/filter/reference,
Off/On/Ignore, узлы диаграмм, результаты 30->3->empty и save/reopen.

## Замеры И Тесты

Первичные ориентиры, не новое A/B: пользовательский report 07.10 17:41,
108 concrete / 56 rebar / 1 LC, batch=0.09375 с; summary=0.457 с;
snapshot=0.137 с; plot=0.281 с; full=1.0546875 с. Исторический v328
не является baseline производительности текущего выпуска.

Направленные проверки на частных копиях: Storage 532/0; Reader 67/0;
Numeric 184/0; SectionSolver 1492/0; Preparation 698/0; UltimateContracts 756/0.
Первый UltimateContracts выявил нашу регрессию для callback с Nothing;
журнал 481/275 сохранен, guard исправлен без изменения expected, повтор 756/0.

SolverMeasurements: пять чередующихся A/B, точные fingerprints совпали.
На 1600 бетонных элементах и 4 стержнях медиана ядра для 30 LC:
Newton 1.36347 -> 1.07952 с (-20.8% времени),
Secant 2.92858 -> 2.26528 с (-22.7%). Это не сквозное ускорение книги.
Начальный короткий StorageMeasurements с Timer слишком грубый;
StorageQPCMeasurements завершен; окончательная серия с теми же
50000 элементов/100000 строк принята в StorageQPCMeasurementsV2.
BatchPilot DIRECT/1: Staged и настоящий Full имеют точные совпадения снимков.
Первый pilot выявил восстановление preparation через справочный getter;
исправление подтверждено DirectMeasurements: после записи released=True.
Полная серия DIRECT 1/10/30 с пятью A/B завершена (60 наблюдений);
все снимки и meta совпали точно. На маленьком пользовательском Hollow
Staged/30: 2.33 -> 2.08 с, настоящий Full/30: 2.122 -> 2.110 с.
Первоначальный Full/1 показал 1.428 -> 1.560 с; этот результат сохранен.
Повторная контролируемая серия на явно скомпилированных и сохраненных A/B,
по пять открытий с двумя расчетами на открытие, дала Full/1:
первый 0.8913 -> 0.8928 с, повторный 0.8782 -> 0.8610 с.
Прежнее замедление не воспроизвелось; нельзя приписать его одной причине
без отдельного доказательства. RunWithinOpen не смешивается в статистике.
В Full измеряется весь настоящий макрос; нулевые phase-поля означают
NotMeasured, а не отсутствие затрат. Фазы измеряются в отдельном Staged.
Перед пакетным CRACK найден сброс injected Query внутри InitializeForFormation.
Передача заменена optional-аргументом Calculate после штатной инициализации;
это сохраняет прежний новый Query для автономных вызовов. Счетчик batch-подготовки
должен подтвердить реальный reuse, а не только совпадение результата.
CompiledCandidate прошел явную компиляцию проекта до Save, Preparation 698/0
и направленный SP35Pipeline 34/0. Пользовательская книга/TXT/report не изменены.
Пятиразовая серия СП35 1/10/30 закончена: 60 наблюдений, точные снимки/meta
совпали. На 30 LC общий domain строится один раз; released=True.
DirectedFinal: все 17 адресных gates завершены успешно, включая полный
SP35 EndToEnd, capacity contracts/lifecycle, published/retained snapshots,
unit-change import, dropdown/borders/status palette и восстановление Excel guard.
Первый GateSeries после Directed содержал ошибку оболочки аргументов PowerShell:
семейства и Count были переданы неверно. Его measurement-папки НЕ используются
как приемка заявленных сценариев; исправленный runner проверяет фактический
Case/число наблюдений/Count и сохраняет Arguments.json. Повтор выполнен в *V2.
Локальный промежуточный коммит: 32ec475; на том этапе выпуск не публиковался.
Исправленные V2 завершили все 12 measurement-jobs: CRACK/SP63,
STABILITY/SP35/SP63, CAPACITY/Auto, CAPACITY_MULTIPLIER/Secant/RectSet,
MIXED, Reverse/Shuffle, Circle/Rounded, Off и Storage QPC.
Основные семейства: по 60 наблюдений, пять чередующихся A/B для 1/10/30;
направленные геометрии/порядки: по 10. Все снимки/meta точные.
Capacity Auto Full/30: 2.970 -> 2.590 с; LoadMultiplier/Secant Full/30:
8.452 -> 7.115 с; MIXED Full/30: 5.930 -> 4.807 с.
Storage QPC: 50000 add 0.533666 -> 0.267015 с, 100000 extent 0.237864 -> 0.047207 с
(точные медианы/диапазоны в AcceptedMeasurements).
IsolationFinal завершил 9 семейств без ошибок: каждый LC в пакете и отдельно,
повтор Execute/retained result/Reverse/Shuffle, tiny/repeated physical LC,
Results-pipeline AutoCADImport с точными дугами/двумя opening.
Все инженерные getters, named-state, stress snapshots и данные кандидатов
сравнены побитно для чисел, точно для кодов/комментариев/ID.
FinalCandidateV2 явно скомпилирован, export lifecycle 6/0; FinalBaselineV2
также скомпилирован, export lifecycle 4/0 (нет новых счетчиков baseline).
Первая попытка FinalCandidate в sandbox не получила Excel COM (80080005),
до импорта; продолжено в разрешенной внешней среде. Первый FinalBaseline
сохранился после compile, но smoke был вызван с несуществующим именем macro;
эта папка не используется. Исправленный fresh V2 прошел настоящий smoke.
Принят локальный export-cache только внутри ReadResultsExportState:
книга по identity, ключ имени таблицы, сброс до рисования и на ошибке.
Этот дополнительный вариант прошел compile/lifecycle/A/B: 30 наблюдений,
по пять A/B для снимков с 1/10/30 LC, десять повторных чтений в каждом замере.
Export/30: 0.931040 -> 0.689044 с, 25.99% сокращения времени читателя.
Это не замер построения графики в AutoCAD; CAD не запускался.
Окончательная повторная регрессия соседних readers принята: 3929/0 и 228/0.
AcceptedMeasurements.json/.md пересобираются только из проверенных серий,
пилоты и папки ошибочного runner не участвуют.
Требуются >=5 чередующихся A/B, raw/median/range/max,
отдельные prep/core/output/full и память, первый и повторные запуски.
Удаленные evaluation/K-build счетчики меняются честно, инженерные поля сравниваются точно.
Пользователь 07.10 явно разрешил необходимые длительные проверки этой стадии.
Старые полные On/Off suites не запускать. Непроверенные gates не объявляются PASS.

## Освобождение Памяти

Все новые кэши ограничиваются операцией/попыткой/запуском и очищаются на
нормальном и ошибочном выходе. Опубликованные результаты нельзя стирать
внутри Batch.Execute: автономный caller имеет право их прочитать после Execute.
Пользовательский макрос после сериализации/схемы освобождает batch/model/provider
и расчетные временные объекты; Results остается только сохраненным Excel-снимком.
Проверяются ссылочные циклы и повторные расчеты при открытой книге. Память
процесса Excel не обязана вернуться побайтно к исходной: его внутренний allocator
может удерживать освобожденные страницы, но VBA-ссылки и расчетные кэши не остаются активными.

## Решения P И B

P01/P02: реализовано точное reuse в одном Solve и принятой Ultimate-точки;
физическая оценка отделена от K без смены stress API, финальная приемка сохранена.
P03/B01: общая Query и mechanical preparation ленивые, результаты LC независимы;
SP35 Freeze отпускает preparation-ссылки, Batch освобождает подготовку при выходе.
P04/P06: bounded scan, локальные таблицы Plot и capacity массивов реализованы.
Пакетная приемка завершена: 74 точных A/B сравнения, минимум пять повторов каждой пары.
P05 допускает предметный отказ, если надежное восстановление делает fast path
дороже прежнего оформления. B02: меж-LC решения/warm start запрещены.
P05: никаких FormattingReady и пропуска очистки не внедрено. Измерены отдельно
clear/summary/snapshot/plot и повторное snapshot-format. Неизменные размеры
не доказывают сохранность merge/font/borders/format/заливки. B04: profile validation
и meta clone измерены серией; новая индексация длительных LC не принята,
чтобы не менять first-match/duplicate/error и порядок ленивой валидации.

## Итог И Восстановление Контекста

GateSeriesV2, DirectedFinal (17) и IsolationFinal (9) завершены успешно.
Imported fixture: 1232 бетонных элемента, 8 стержней, точный круговой outer
и два opening по AutoCADImport snapshot-контракту. A/B 1/10/30 приняты;
Full/30 25.873746 -> 22.248864 с (14.01%). Исходный сохраненный LC без
ConfigureFixture: первый 1.111171 -> 0.955424, повтор 1.085982 -> 1.056626 с.

ReleaseV2: десять расчетов exact, Config/формулы/имена/260 ширин сохранены,
save/reopen exact; plot/export после reopen без solve и изменения Results.
Память и освобождение временных ссылок описаны в Final Report без обещания
побайтного возврата allocator Excel. UI release и fresh build 1583/0,
структурная валидация каждой 27/0. Активных runner-сессий этапа больше нет.

Первичная raw source/book сверка не принята из-за 14 нормализаций VBE.
Не применялось округление чисел для скрытия различий: независимый fresh build
скомпилирован и экспортирован read-only. Все 121 source + 5 document components
совпали с native reference; raw-отчет сохранен, literals не менялись.
В performance prepare/release применен существующий print-area preserve helper,
чтобы COM-save не оставлял неканонический alias Print_Area. Validator не ремонтирует.

Проверенные book/TXT опубликованы в workbook/output. SHA книги:
DE653B7ACFA9EE4A97D1C7114A1067AADDC0229F118E50EE6849A016473DE282;
SHA TXT: 5113A6C88851D65EEDF80F09F4DDAA3A15F6AF18CE40769CFAE8AB976BE504B0.
Publication.json содержит повторные guards и все source hashes.
Пользовательский execution report сохранил исходный SHA и не коммитится.
Исторические private fixtures/непринятые pilots сохранены, не очищаются.

Финальный self-audit в NDM_Performance_Final_Report.md закрывает все P/B пункты
текущего ТЗ. Старые полные On/Off suites не запускались; scope-разрешенные
длительные проверки выполнены. Дальнейшие задачи производительности/архитектуры
не добавляются в этот этап. Коммиты только локальные, без push.
