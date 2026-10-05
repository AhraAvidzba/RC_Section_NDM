# Audit03: Итоговая Область Проверки Config

Дата: 2026-10-05. Численный выпускной срез: `ca28919` / clean v328.
Это актуальная трактовка минимального покрытия K01-K04. Старые реестры и
обзорные MD сохраняют статусы на дату своего создания и не переписываются
задним числом. Полная выпускная приемка final suites, source/book equality,
benchmark и self-audit завершена и изложена отдельно в Final Report.

## Знаменатель

Фактический census clean v328: 35617 сохраненных Config cells, 14 основных
именованных диапазонов, 84 validation. Поадресный реестр содержит 1064
уникальных поля: 760 обычных полей роли `UserInput`, 104 контрольных точки диаграмм,
108 активных и 36 справочных клеток СП 35, 30 производных ID, 16 followers
объединенных селекторов и десять неприменимых клеток. Ничего из этих групп
не исключается из знаменателя. Удаление ровно одной obsolete Formation-path
настройки объясняет переход исходных 1065 к 1064; пользовательский LoadPath
в каждой строке сочетаний остается единственным путем для обоих расчетов.

Текущие значения, формулы, адреса, validation и ссылки берутся из
`regression/Audit03/config_field_registry_release_clean_v328_2026-10-05.json`,
а не копируются из старой книги. Исторический directed union v306 связывает
все те же 1064 ID со своими активными assertions. Для всех 760 inputs в нем
есть тип, внутренние единицы, default/активность, потребитель, ожидаемый эффект,
test-ID, допустимый диапазон и runtime evidence. Наличие этих полей само по
себе не является PASS; выполненные проверки перечислены ниже.

760 не является числом всех изменяемых ячеек книги: 144 клетки таблицы СП 35
учтены отдельной ролью `NormativeTableValue`. Из них 108 имеют расчетного
потребителя и отдельные поведенческие/invalid проверки, 36 являются справочными.
Итоговый поадресный artifact различает verified table behavior, derived
dependence, control value и inactive/shared-anchor contract, не объединяет
все поля вне `UserInput` в неприменимые.

## Минимальное Поведенческое Покрытие

| Семейство | Фактическая Проверка И Граница |
| --- | --- |
| Solver и Capacity options | Настоящий reader, допустимые/недопустимые числа, эффект iteration/load-step/damping/tolerance/restart, автономный и batch consumers, active/inactive. `modTestSectionSolver`, `modTestConfiguration`; прежние component tolerances сохранены. |
| Search strategy/method и LoadPath | Все стратегии, Bisection/Brent/Secant, Auto и пять фиксированных путей для Capacity и Formation. Selector matrix v306: 62 группы, 4464 path cases, 302 пары, 48 рискованных четверок. |
| Технические численные границы | Соседние Double и representability Search, iteration budget=0, extreme formation/state, Long counters, сумма рядов и произведение retries/load steps. Поведение вне арифметической области не выдается за физический предел. |
| Материалы | 23 вводимых значения, оба знака, 16 spec-комбинаций, активные и неприменимые ветви, ошибки/aliases/единицы, 104 производных controls. Solver не читает контрольную таблицу диаграмм. |
| Единицы и знаки | Все 12 input/output selectors и три знака; эквивалентные внутренние нагрузки/геометрия, преобразования Results/plot, смена единиц после импортного snapshot. AutoCAD имеет фиксированный mm-контракт. |
| Геометрия и арматура | Circle, RoundedRectangle, HollowRectangle, три RectSet формы; размеры/радиусы/смещения, count/diameter/position, включенные/выключенные ряды и формы, required-input errors/recovery. RequiredGeometry и RebarField suites дополняют ранний effect-only registry. |
| Circle дополнительный ряд | Две позиции именованной таблицы, второй/третий ряд, Stacked/SideBySide: frozen negative 24/40, positive v325 64/0; Geometry Off/On 668/0. Ошибка размещения содержит реальные адреса трех участвующих параметров. |
| Crack options | Все Psi/zone/cover/averaging modes и активные numerical effects; независимая формула, signed sigma_s,crc, current-state availability, NotCracked, предупреждение только реально использованного fallback. Width не решает State. |
| Stability | 17 полей, две главные плоскости, СП 63 и ветви СП 35, Auto/User, units/signs, active/inactive, ошибки и перенос Name. Полный v328 10016/0; отдельно Off/On по 2072/0 для 72 крайних L/mu случаев и recovery. |
| Профили | Все PR1-PR4, calculation flags, material spec, quantity/state selectors, aliases и descriptions; каждое enum-значение в активном сценарии и invalid choice. Wrong duration-reference не блокирует независимые расчеты. |
| Таблицы сочетаний и duration | Все editable строки, индексы/порядок/дубликаты/ссылки, blank/zero/ошибки/overflow и продолжение следующих строк. Пустой LoadPath допустим и означает Auto; пустота нагрузки не равна ошибке формулы. |
| General/presentation/annotations | Настоящий эффект report/plot/worst/source/LC/labels/styles; errors и recovery, inactive consumer isolation, чтение только frozen Results, без нового solve. Размеры столбцов задаются сборкой, writer их не переустанавливает. |
| AutoCAD | Все 20 editable CAD fields, layers/ACI/scales/cleanup/MinArea, actual Region import/export/DWG, фиксированные mm и frozen Results. Native shape acceptance 1023/0, 22 формы/ориентации; translation 66/0. |
| СП 35 table | 144 адресных случая 829/0; 540 invalid variants 1621/0; order/missing-name/relocation/comments/reopen 105/0. 108 активных клеток сверены независимой интерполяцией, 36 справочных не влияют на вычисление. |
| Derived/inactive | 30 ID следуют текущим LC; 16 merged followers имеют единственный anchor; десять неприменимых клеток не читаются активным consumer. Errors в неиспользуемых клетках не меняют normalized payload. |

Полные runtime evidence находятся в directed union v306, соответствующих
семейных Review/Evidence MD/JSON и финальных eight-suite журналах. Широкая
матрица v320 независимо проверяет 13 форм и 949 form/load fixtures: 22776
путевых вариантов, 957531/0 assertions, save/reopen и source unchanged.
Она не называется 22776 независимыми векторами нагрузки.

## Чувствительность И Допустимая Область

Пять локальных отрицательных мутаций проверяют передачу Solver.Method и
четырех Search selectors. Они дают ожидаемые failures, не входят в release.
Mass mutation-framework не добавлен, новых `.cls` нет. Лимиты, допуски,
профили и inactive ветви дополнительно проверяются наблюдаемым effect oracle,
а не только одинаковым успешным конечным числом.

K02 не требует перебрать все значения Double или все декартовы произведения.
Приемка включает допустимые границы и классы невалидного ввода, перечисленные
в тестах; она не утверждает, что всякое положительное инженерное число
нормативно допустимо. Для Count проверяется representability до выделения
массива; миллиардные сетки намеренно не выделяются. Для L/mu проверены
переполнение произведения/квадрата и округление квадрата до нуля, с ранним
InputErr и действующими адресами. Для Search проверены крайние шаги и
скобки без ослабления физического или equilibrium критерия.

Полные blank/CVErr/missing комбинации для каждого отдельного profile enum
и каждое возможное сочетание aliases не являются заявленным exhaustive
покрытием. Контракт чтения определяется общим reader/catalog и проверен
на направленных маршрутах; каждый активный enum имеет свой invalid-value
тест. `Null` не имитируется путем записи в Excel cell: реальный Range не
хранит такой payload. Автономные Variant/API проверки рассматриваются
отдельно от реального ввода таблицы.

## Поздние Уточнения Реестра

- `General.WorstCombinationCriterion`: в полной книге ключ обязателен.
  Default автономного Batch API не является fallback прочитанного Config;
  актуальный контракт и negative/positive 71/20 -> 91/0 находятся в
  `NDM_Audit03_Worst_Config_Evidence.json` и v322 runtime reports.
- `Solver.MaxDeltaKappa`: ноль отключает ограничение приращения; отрицательное
  значение невалидно. Ранняя фраза RangeReview о запрете нуля была неточной,
  не изменяет фактический код и явно не переносится в текущий контракт.
- Для восьми Circle fields и 24 material cells (23 активных ввода и одна
  неприменимая растянутая клетка Rb,mc2) старый `PendingBoundaryTests`
  является метаданными первого census, а не отсутствием поздних Required/
  Rebar/Material input tests. Конкретные тесты привязаны к каждому ID в union.
- Положение всех input errors определяется текущим named range, не строковым
  адресом из каталога. Проверены переносы реальных Name и динамический текст.

## Сохранность Пользовательской Книги

Обновление выполняется только на копии исходной AAF5D4FA-книги. Сохраняются
все значения/формулы LC и все исходные inputs; меняются согласованные
служебные подписи, прямые контрольные формулы, validation, справка и
объединение общих selectors. Migration повторяем; default новой clean-книги
не записывается поверх пользовательского значения. Help v328 сохраняет все
760 signatures, 140 links, 118 shapes и Print_Area. Clean/update help имеет
2828 одинаковых cells и одинаковые merges, failed=0.

Итоговые SHA, завершенные final suites, source equality и публикация
фиксируются в `NDM_Audit03_Final_Report.md`. Соответствующие runtime-журналы,
publication receipt и read-only Validate опубликованной книги сохранены;
они не подменяются одной статической проверкой реестра.
