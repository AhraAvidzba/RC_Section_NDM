# NDM Audit03: Матрица Краевых Случаев

Статус: directed runtime evidence частично получен; полная нагрузочная матрица еще не выполнена.
Baseline: `df10412f0e0baf918f5e97cbc87b6bf16c3d4cae`.

## Уровни

1. Быстрые F-reproducers и smoke: реальный Excel Range, typed provider/state, generic search и progress guard.
2. Все штатные suites Off/On, с сохраненными историческими explicit-On setup.
3. Полная нагрузочная матрица: реальные формы/нагрузочные семейства/селекторы, pairwise и явные high-risk комбинации, независимые проверки.

## Обязательный Перебор Путей Нагрузки

Дополнение пользователя от 2026-10-01: во время нагрузочных тестов проверять все пути поиска несущей способности и точки трещинообразования.

| Поиск | Обязательные Варианты | Что Проверяется Для Каждого |
| --- | --- | --- |
| Capacity | `lambda*Mx`, `lambda*My`, `lambda*Mxy`, `lambda*N`, `lambda*NMxy` | Постоянная и масштабируемая части N/Mx/My, lambda=0, обычный/предельный/перегруженный LC, конечный физический критерий, фактический метод, typed status/code и ResultComment. |
| CrackFormation | `Auto`, `lambda*Mxy`, `lambda*N`, `lambda*NMxy` | Before/After состояния, фактические Ncrc/Mcrc/lambda_crc, постоянная часть и отсутствие достижимой трещины, численная неудача отдельно от физического результата, typed status/code и ResultComment. |
| Auto CrackFormation | Последовательность всех трех фиксированных путей | Причина перехода или остановки, выбранный путь, отсутствие незаконной смены фиксированной траектории; успешная проверка Auto не заменяет проверки отдельных путей. |

Перебор выполнять на осевом сжатии/растяжении, изгибе обоих знаков по обеим осям, четырех квадрантах Mx/My, смешанных N+M, постоянной части около/за пределом и больших нагрузках. Нулевая масштабируемая компонента - отдельная проверка применимости, а не фиктивный найденный предел. Для каждого пути сохранять результаты и проверять состав комментариев по всем применимым output-блокам. Реестр должен явно связывать путь с параметрами, assert-ами и логом; до выполнения строка не считается покрытой.

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

## Семейства

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

Конкретные параметры, test-ID, число независимых кейсов/вариантов/assertions,
логи и обоснованные неприменимые комбинации будут добавляться по мере реализации.
