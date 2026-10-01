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
| 1. Контрольная точка и карта проверок | в работе | Baseline build/Off/On, фактические Config/help/class inventories, F-reproducers; источники и user Config неизменны. |
| 2. Корректность входа, поиска и метаданных | в работе | F01/F02/F03 и основные F06 контрпримеры имеют runtime evidence; F04/F05/F07 и окончательная приемка еще не завершены. |
| 3. Упрощение архитектуры | не начато | Удалить CBatchStatusPolicy и CCrackWidthFormulaCalculator после переноса обязанностей; canonical aggregates, consumers/build/tests; без новых классов. |
| 4. Измеряемая оптимизация | не начато | Profile корректного baseline, минимум пять повторов, asymmetric/retries/geometry/allocation; каждая оптимизация подтверждена числами/status/cache. |
| 5. Config/краевые нагрузки/документация/UI | не начато | Все editable-поля, каждый enum и активный эффект; широкая L01-L17/shape/setting матрица; обязательный ResultComment-контроль каждого нагрузочного кейса, русские комментарии, actual help/validation/colors/CF/save-reopen. |
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
| F05 | в работе | PhysicalRangeMeta явно объясняет выход за физические деформации/extension; продольные причины, subtree и фактический output еще не приняты. |
| F06 | основные directed runtime PASS, полный аудит продолжается | Baseline adjacent-Double Bisection budget=0 завис до watchdog; исправленный generic matrix 114/0, real Capacity tol=1e-18/budget=0 завершен за 58 probes без ложной точки. Остальные Ultimate/recovery call paths еще проверяются. |
| F07 | в работе | Исправлены Nothing guard в stability, Split bounds в plotter и два unsafe test guards. Статическая проверка IIf/array call paths выполнена частично; окончательная приемка не заявлена. |
| F08 | не начато | Каждый Audit01/Audit02 ID -> owner/method/test/log/result/environment, не только общая фраза. |
| A01-A05 | не начато | Два слияния, владельцы meta/aggregation, canonical data, output и все consumers. |
| P01-P04 | не начато | Правильный baseline, профиль и >=5 повторов; geometry/retries/cache не подменяют физику. |
| K01-K04 | не начато | Полный фактический Config census/behavior/active-inactive/mutation sensitivity/isolation. |
| T01-T05 | не начато | L01-L17, все формы/selector variants/pairwise/high-risk tuples, независимые инварианты. |
| D01-D02 | не начато | Все classes/modules/nontrivial methods; фактический help/Config/validation/links в clean и update. |
| UI01 | не начато | Все семь внешних статусов, actual DisplayFormat/CF/legend/reset/save-reopen. |
| W01 | в работе | Git/base/spec/hash/progress сохранены; checkpoints без push/destructive Git. |

## Important Decisions

- Дополнение пользователя: нагрузочная матрица обязательно покрывает все пути Capacity (`lambda*Mx`, `lambda*My`, `lambda*Mxy`, `lambda*N`, `lambda*NMxy`) и CrackFormation (`Auto`, `lambda*Mxy`, `lambda*N`, `lambda*NMxy`). Проверять base/offset на lambda=0, фактически масштабируемые компоненты и момент от эксцентриситета N, физический критерий конечной точки, статусы и ResultComment. Не подменять полный перебор тестом только Auto или общим coordinator. Для Auto отдельно проверять последовательность путей и причины перехода; для фиксированного пути - собственный физический результат без незаявленной смены траектории.

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
- Все тесты выполнялись на отдельных книгах; пользовательская output-книга и Config не изменены. В temporary baseline-source copy добавлен только новый test module для F03 reproducer; production baseline остался прежним.

## Как Продолжить

1. Обновить isolated test book текущими исходниками; выполнить lifecycle directed и полные Off/On через watchdog. При новом failure сохранять лог, исправлять причину, не ослаблять физические expected.
2. Завершить F03/F04/F05/F06/F07 call-site acceptance, затем A01/A02 целым согласованным объемом и фактический Config census.
3. До завершения всех шести gates цель остается активной; новые ограничения/изменения после green suite записывать сюда.
