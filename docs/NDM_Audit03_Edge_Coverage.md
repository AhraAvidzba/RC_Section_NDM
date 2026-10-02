# NDM Audit03: Матрица Краевых Случаев

Статус: 13-формная Light/Stress матрица v48 завершена: 52 новых runs,
18980 независимых путевых случаев, без failures. Приемка всех семейств,
селекторов и их взаимодействий еще не завершена. Новые после v48 правки
имеют отдельные directed/full gates и не приписываются прежней книге.
Baseline: `df10412f0e0baf918f5e97cbc87b6bf16c3d4cae`.

## Уровни

1. Быстрые F-reproducers и smoke: реальный Excel Range, typed provider/state, generic search и progress guard.
2. Все штатные suites Off/On, с сохраненными историческими explicit-On setup.
3. Полная нагрузочная матрица: реальные формы/нагрузочные семейства/селекторы, pairwise и явные high-risk комбинации, независимые проверки.

## Обязательный Перебор Путей Нагрузки

Дополнение пользователя от 2026-10-01, подтверждено 2026-10-02: во время нагрузочных тестов проверять все пути поиска несущей способности и точки трещинообразования.

| Поиск | Обязательные Варианты | Что Проверяется Для Каждого |
| --- | --- | --- |
| Capacity | `lambda*Mx`, `lambda*My`, `lambda*Mxy`, `lambda*N`, `lambda*NMxy` | Постоянная и масштабируемая части N/Mx/My, lambda=0, обычный/предельный/перегруженный LC, конечный физический критерий, фактический метод, typed status/code и ResultComment. |
| Auto Capacity | `Auto` отдельно от пяти физических траекторий | Выбор `lambda*Mxy` при пользовательском моменте, `lambda*N` при только N и неприменимость поиска без масштабируемой нагрузки. Момент от смещения N не подменяет пользовательский момент при выборе пути. |
| CrackFormation | `Auto`, `lambda*Mxy`, `lambda*N`, `lambda*NMxy` | Before/After состояния, фактические Ncrc/Mcrc/lambda_crc, постоянная часть и отсутствие достижимой трещины, численная неудача отдельно от физического результата, typed status/code и ResultComment. |
| Auto CrackFormation | Последовательность всех трех фиксированных путей | Причина перехода или остановки, выбранный путь, отсутствие незаконной смены фиксированной траектории; успешная проверка Auto не заменяет проверки отдельных путей. |

Перебор выполнять на осевом сжатии/растяжении, изгибе обоих знаков по обеим осям, четырех квадрантах Mx/My, смешанных N+M, постоянной части около/за пределом и больших нагрузках. Нулевая масштабируемая компонента - отдельная проверка применимости, а не фиктивный найденный предел. Для каждого пути сохранять результаты и проверять состав комментариев по всем применимым output-блокам. Реестр должен явно связывать путь с параметрами, assert-ами и логом; до выполнения строка не считается покрытой.

Runner имеет десять вариантов на нагрузку: шесть Capacity (включая выбор
Auto) и четыре Formation. Полный v48 проверяет все 13 форм, Light/Stress
и Off/On, comments/writers/txt report и save-reopen; source неизменен.
CSV v48 включает исторические логи, поэтому для этого gate выбирать только
его 52 runs. Прежние v35-v39 девятипутевые логи сохраняют исходное покрытие,
но не выдаются за новый Auto Capacity gate. Внутренняя согласованность таблиц
не доказывает, что каждый NumFail неизбежен; дополнительные независимые
extended-equilibrium и stress-feasibility тесты выявили новые modest failures.

## Проверки Текста И Сборки

Для каждого нагрузочного варианта сохранять исходные параметры, effective settings,
leaf/subtree meta и фактические ResultComment. Проверять численные результаты и
комментарии совместно; зеленое равновесие не означает автоматически правильный текст.

| Слой | Обязательная Проверка |
| --- | --- |
| State / Search | Точная typed-причина, сообщение соответствует converged/physical/extension и конкретному пределу/траектории; NumericalFailure не определяется по тексту. |
| Инженерные leaf results | Отказ, blocked и warning имеют информативную русскую причину; результат по формуле не приписывает себе solve; longitudinal указывает превышение сжатия над Rb,mc2. |
| Strength / Crack / Stability subtree | Только относящиеся дочерние причины, логический порядок, сохраненные warnings, без дублей и противоречий. |
| Подробные блоки Results | Текст равен подготовленному комментарию соответствующего result-subtree; нет чужих разделов и потери relevant-причины. |
| Batch summary | Объединены все применимые проблемные разделы LC с понятными метками; нет повторного склеивания writer-ом. |
| Report / save-reopen | Причина не теряется при упаковке/сохранении/чтении; units/знаки/обозначения согласованы. |
| Читаемость | Не только непустая строка: нет служебного имени вместо инженерного смысла, сообщения об успехе при FAIL, двойных точек, сломанной кириллицы/символов, полного iteration-log вместо причины. |

Автоматические assertions проверяют присутствие/состав/границы/согласованность.
Содержательное чтение всех уникальных шаблонов и фактически полученных редких
сообщений дополняет тесты. До выполнения нельзя обозначать эту проверку PASS.

## Первоначальный Реестр Семейств

Статусы таблицы ниже отражают первоначальную постановку. Свежие выполненные
gates перечислены в следующих разделах; таблица не используется как итоговый
PASS всей L01-L17 приемки.

| ID | Планируемая Граница / Инвариант | Статус |
| --- | --- | --- |
| L01 | Empty/zero/invalid row; 5/6/7+ columns, source gaps | не начато |
| L02 | Ненулевые tiny N/M обоих знаков и эквивалентные units, реальные residual tolerance | не начато |
| L03 | Чистое растяжение/сжатие, physical/crack applicability | не начато |
| L04 | +/-Mx, +/-My, symmetry only when valid | не начато |
| L05 | Все квадранты biaxial, normalized ratios 1/10/1000/1000000 и обратные | не начато |
| L06 | Оба знака N, малый/большой eccentricity, почти осевое состояние | не начато |
| L07 | ZeroMomentPerDepth ниже/на/выше, 0/positive, размеры/units | не начато |
| L08 | Capacity below/on/above/near-limit, подтвержденный критерий | не начато |
| L09 | Formation до/на/после/above-current, без фиктивного Post | не начато |
| L10 | Overload 2/10/100/1000/до 1000000, конечное выполнение/правильная причина | не начато |
| L11 | Offset lambda=0 acceptable/boundary/beyond, fixed-N и другие paths | не начато |
| L12 | Formation fixed/Auto, crack in constant part, warnings/fallback | не начато |
| L13 | CriterionNotReached vs SearchBoundReached, последний интервал/MaxLambda | не начато |
| L14 | Adjacent Double/huge bounds/tiny tolerance/overflow/stagnation/watchdog | generic 114/0 и real Capacity 58 probes; весь high-risk набор еще не завершен |
| L15 | Sigma_s,crc sign/zero/averaging, Psi modes/upper bound | не начато |
| L16 | Reorder/repeat/overload-normal-switch, lifecycle/meta/comment isolation | не начато |
| L17 | Invalid input/config/missing dependencies, no accidental OK/NumFail | provider/state 73/0, Capacity contracts 54/0, Formation four-path config 64/0; полная output/shape матрица впереди |

## Полученные Directed Evidence

- `f06_arithmetic_positive_v4_2026-10-02.txt`: 2 generic result kinds x 3 methods x 4 arithmetic scenarios, 114 assertions; baseline stagnation подтверждена watchdog timeout.
- `f03_f06_capacity_positive_v6_2026-10-02.txt`: real precision case и 6 typed failure scenarios, 54 assertions; precision case заканчивается честным NumFail без fictitious limit point.
- `f03_formation_positive_v7_2026-10-02.txt`: все четыре пути x 2 terminal configuration errors, 64 assertions. Это проверка ошибок настройки, не замена полного нагрузочного перебора этих путей.
- `f04_lifecycle_positive_v8_2026-10-02.txt`: 172 assertions флагов/снимков/фабрик; не заменяет reorder/repeat расчетов с реальными нагрузками.
- `f01_f07_full_off_v8_2026-10-02.txt`: 7289/0 по всем восьми suites. Счетчик assertions не равен числу независимых инженерных случаев.
- `a03_paths_comments_off_v11_2026-10-02.txt`: 27 независимых LC (Capacity 5 paths x 3 load cases; Formation 4 paths x 3), 495/0; actual provider Extension=False. RoundedRectangle 300 x 200 мм, четыре стержня d20, offset X=10/Y=-7 мм. N/Mx/My во внутренних Н/Н*мм: (-150000/-6000000/-3000000), (-1500000/-60000000/20000000), (200000/9000000/-3000000). ZeroMomentPerDepth=0, stability=No; временные настройки восстановлены.
- `a03_paths_comments_on_v13_2026-10-02.txt`: те же 27 LC, effective Extension=True, 504/0. Все leaf comments, subtree assembly и фактические подробные/batch ячейки совпали. Для каждой найденной Capacity-точки независимо проверены масштабируемые компоненты и момент N*offset; для formation сохранены выбранный путь и Pre/Post physical flags.
- `a03_paths_comments_off_v10_2026-10-02.txt`: exploratory 492/4, не PASS. Expected не учитывал моментный zero-filter после переноса N; fixture теперь явно изолирует проверку траектории от фильтра. Сам фильтр и его границы не менялись.
- `a01_a04_full_off_v13_2026-10-02.txt`: 7798/0 по восьми suites; это промежуточная приемка текущего среза, не завершение широкой матрицы.
- `f06_ultimate_stagnation_negative_2026-10-02.txt`: реальный VBA UltimateStrain с прежним модулем `1a65796` завис при min alpha=0 и постоянной норме; watchdog остановил test Excel через 60 секунд. Это узкий контрпример, остальные зависимости и тесты текущие.
- `f03_f06_ultimate_positive_v16_2026-10-02.txt`: 340/0. Два result-kind, 13 guard-сценариев, точные InvalidConfiguration/InternalError/SingularTangent, непредставимый шаг, переполнение и реальные Capacity Ultimate-маршруты. Ни одна ошибочная проба не принята как найденный предел.
- `f03_f06_full_off_v16_2026-10-02.txt` и `f03_f06_full_on_v16_2026-10-02.txt`: 8285/0 и 8286/0; все восемь suites, source unchanged=True. `source_contracts_v16_2026-10-02.txt`: 101/101, 0 ошибок. Это приемка текущего среза, не закрытие всех L01-L17.

В 27-LC matrix присутствуют физические FAIL/BaseFail, numerical failure и
blocked формулы. Честный NumFail найденной пробы не объявляется ошибкой самого
регрессионного теста, если точка не выдумана и сохранена точная typed-причина.
Качество каждого уникального текста проверено по логам; обнаруженная подмена
blocked reason успешным solver.StopReason исправлена. Повторное чтение после
save/reopen, txt-report, другие формы и все L01-L17 остаются отдельными gates.

Конкретные параметры, test-ID, число независимых кейсов/вариантов/assertions,
логи и обоснованные неприменимые комбинации будут добавляться по мере реализации.

## Расширенный All-Path Runner

`modTestBatchCalculation.RunAudit03BroadLoadMatrixTests(shape, family)` выполняет
каждую независимую нагрузку по шести Capacity (пять fixed + Auto) и четырем Formation путям. Light:
49 нагрузок, включая ноль, оба знака tiny/осевой/одноосной, четыре квадранта
косого изгиба с нормированными отношениями 1/10/1000/1000000 и обратными,
смешанные N+Mx+My. Stress: 24 нагрузки с масштабами 0.95/1/1.05/2/10/100/1000/1000000
от независимой силовой оценки сопротивления. Эта оценка не объявляется точной
capacity при ненулевом эксцентриситете и не заменяет отдельную приемку L08.

Проверяются named states, три компонентные невязки, физический диапазон,
extension и Psi; все leaf/subtree comments, четыре реальных блока writer,
настоящий txt-report и равенство всех значений Results после save/reopen.
Исходная тестовая книга не меняется: результаты остаются в сохраненной копии.

- Smoke v35 RoundedSimple Off: 9 независимых путевых случаев, 273/0; save/reopen
  совпал. Это smoke, а не приемка всех форм/настроек.
- Light v35 RoundedSimple Off: 441 случай, 12687/135. Лог сохранен как отрицательный.
  Два Capacity UltimateStrain результата имели найденную точку и OK, но
  `WithinPhysicalRange=False`. Финализация проверяла равновесие, но не
  физический диапазон. Введена обязательная проверка физики перед публикацией;
  непринятая плоскость уточняется следующим общим Newton-шагом. Допуски,
  диаграммы и исходные численные эталоны не ослаблены.
- Остальные failures v35 относятся к двум неверным предположениям нового теста.
  Явный путь без ненулевой масштабируемой компоненты имеет существующий
  `rsInvalidInput/INVALID_INPUT`, без Search/CapacityState; тест теперь проверяет
  именно этот контракт. При успешном CurrentState ширина может быть заблокирована
  отказом Formation; причина сохраняется в Formation leaf и crack subtree,
  Width указывает отсутствующую зависимость, без копии успешного комментария НДС.
  Это не изменение production-методики и не ослабление физической приемки.
- PhysicalBoundary v36 RoundedSimple Off: две контрпримерные нагрузки по всем
  девяти путям, 18 независимых случаев, 538/0. CapacityState в обоих проблемных
  My-путях имеет min eps бетон = -0.0035, confirmed equilibrium, physical=True,
  extension=False. Save/reopen равенство True; источники неизменны. Лог:
  `broad_matrix_RoundedSimple_PhysicalBoundary_off_v36_2026-10-02.txt`.

Полный Light/Stress перебор 13 форм и двух режимов v45 завершен, см. ниже.
Он не заменяет per-selector/pairwise, near-limit, reorder и actual UI gates.

Повторный v36 Light RoundedSimple Off: 441 случай, 12954/0; Stress: 216 случаев,
8491/0. Оба complete/save-reopen/source-unchanged. Full v37 Off/On: 8444/0 и
8445/0 по восьми suites; source/export 101/101, 0 ошибок. Для просмотра результатов
создан CSV cases и runs manifest через `Export-Audit03LoadMatrixSummary.ps1`.
В таблице неуспешные/незавершенные прогоны не получают PASS, фактические typed
статусы и комментарии сохраняются без нового display-маппинга в скрипте.

## Завершенная Матрица v45

`broad_matrix_runner_v45_2026-10-02.log`: 52/52 прогонов, 18 980
параметризованных путевых случаев = 13 форм x (49 Light + 24 Stress нагрузки)
x 10 путевых вариантов x 2 effective Extension режима. Это не 18 980 разных
нагрузочных векторов и не количество assertions.

Каждый отчет подтвердил complete, failed=0, Results save/reopen equal=True
и неизменность source. SHA-256 source:
`592C1268CA38964CBBB432AEEDB3961CEAC95B2EC83B50D25A7DB81DF918674C`.
Все 52 строки v45 в `load_matrix_summary_v45_2026-10-02_runs.csv` имеют PASS.
CSV cases сохраняет фактические нагрузки, пути, meta и output comments;
старые отрицательные/неполные версии остаются NotPassed в общем manifest.

Содержательная дополнительная ревизия комментариев обнаружила нарушения,
которые прежний автоматический тест не проверял для success leaf. Поэтому
PASS v45 не объявляется приемкой исправлений v46 или полной читаемости текста.

## Directed Gates v46

| Контракт | Negative v45 Production + Новые Тесты | Positive v46 | Граница Доказательства |
| --- | --- | --- | --- |
| API массивов диаграммы, valid-invalid-valid | `f07_diagram_array_negative_v46_2026-10-02.txt`: 30/25 | `f07_diagram_array_positive_v46_2026-10-02.txt`: 55/0 | Девять invalid входов и допустимый дополнительный индекс 0; не подмена Config-ошибки низкоуровневым массивом. |
| Экстремальные представимые нагрузки | `f07_extreme_state_negative_v46_2026-10-02.txt`: 28/28 | `f07_extreme_state_positive_v46_2026-10-02.txt`: 56/0 | Восемь Newton/Secant x center/shift x Off/On, фактический State/status и восстановление обычного solve. |
| Общая причина CurrentState в Crack aggregate | `comment_composition_negative_v46_2026-10-02.txt`: 28/4 | `comment_composition_positive_v46_2026-10-02.txt`: 32/0 | Четыре typed причины, leaf unchanged и отдельная другая dependency reason. |
| Читаемость и состав actual comments всех путей | `path_comments_negative_v46_2026-10-02.txt`: 703/17 | Off 720/0, On 721/0 | Реальные четыре writer-а, execution report, save/reopen; full v46 Off 9572/0, On 9573/0. |

Новые `MATRIX_CRACK_DATA` строки записывают подготовленные sigma_s/sigma_s,crc,
psi_s, Abt/As/ds/ls/a_crc и выбранные стержни во внутренних единицах. Их нет
в исторических v45 логах; CSV явно отмечает CrackDataLogged=False, а не
подставляет нули вместо непроверенных величин. Live L15 coverage еще требуется.

## Расширенное Равновесие v48

Дополнение пользователя проверено независимыми заданными плоскостями, из
которых интегрированием получены нагрузки. Все 192 равновесия On на четырех
формах найдены общим provider-ом с холодного старта, Newton/Secant и двумя
material roles. Проверены деформации вплоть до 9,5, 32 парных Off отказа и
8 нагрузок выше осевой технической возможности при +/-10. По 502/0 assertions
на форму. Реальный дефект потери лучшего failed retry-start исправлен;
подробности и границы доказательства в `NDM_Audit03_Extended_Equilibrium_Review.md`.

v46 широкая матрица остановлена после трех успешных runs и четвертого с четырьмя
ошибками comment-oracle. Две независимые стадии имели одинаковый текст причины,
а oracle ошибочно требовал одно вхождение вообще. После исправления только
тестового подсчета CircleSym/Stress/On v47: 240 случаев, 10414/0,
Results save/reopen equal=True и source unchanged=True. Полный v48 повтор
нужен отдельно; отрицательный v46 отчет сохранен и не считается PASS.
