# Audit03: Покрытие Config

История приемки начата 2026-10-02. Итог K01-K04 на 2026-10-05 завершен
в минимально достаточной области K02: 1064 фактических поля, 760 UserInput
и 144 клетки нормативной таблицы учитываются отдельно. Поадресные текущие
контракты не содержат Pending/NotRun; это не исчерпывающий перебор Double
или декартова произведения настроек.

Итоговые роли, эффект и ограничения: `NDM_Audit03_Final_Config_Review.md`,
`regression/Audit03/config_final_evidence_v328_2026-10-05.json/.csv`.
Full Off/On v328: 108901/0 и 108912/0; публикация и self-audit в
`NDM_Audit03_Final_Report.md`. Следующие записи сохраняют состояние своих
датированных срезов; старые Pending и denominator 1065 не являются итогом.

## Фактическая Книга

Исходная пользовательская книга `workbook/output/RC_Section_NDM.xlsm` прочитана
как OpenXML, без открытия в Excel и изменения ячеек. SHA-256:
`AAF5D4FAF06197F01966463DE85BA813EB4216B455AA8504B2DE4F0BA717C012`.

- Полный census: `docs/regression/Audit03/config_census_2026-10-02.json`.
- Поадресный реестр: `docs/regression/Audit03/config_field_registry_2026-10-02.json` и `.csv`.
- Census содержит 35617 сохраненных ячеек, 14 именованных диапазонов и 72 validation.
- Реестр содержит 1065 отдельных адресов, включая пустые поля, все четыре профиля,
  30 строк сочетаний и 30 строк длительных нагрузок. Повторных адресов нет.
- Это не 1065 независимых переключателей: повторные строки LC и нормативная
  таблица развернуты по ячейкам для контроля сохранения и граничных ошибок.
- 241 default сопоставлен с действующим каталогом новой книги. Сохраненное
  пользовательское значение не считается default и не перезаписывается им.

## Состав Реестра

| Блок | Полей | Примечание |
| --- | ---: | --- |
| rngSystemSettings | 102 | Групповые заголовки не считаются настройками. |
| rngUnitSettings | 12 | Ввод и вывод для каждой физической величины. |
| rngSignConventionSettings | 3 | Внутренняя система знаков не меняется. |
| rngSteelMaterialParameters | 10 | Сжатие и растяжение отдельно. |
| rngConcreteMaterialParameters | 14 | Включая неприменимую растянутую ячейку Rb,mc2. |
| rngCalculationProfiles | 92 | 23 поля на каждый PR1-PR4. |
| rngPlotAnnotationSettings | 26 | Две группы, включая явно неприменимые поля `-`. |
| rngCircleGeometry | 8 | Пустые диаметры дополнительных рядов включены. |
| rngRoundedRectangleGeometry | 46 | Также поля вне узкого именованного прямоугольника. |
| rngHollowRectangleGeometry | 80 | Четыре внутренние ячейки количества `-` не являются вводом. |
| rngRectSetGeometry | 94 | Геометрия, грани и дополнительные ряды. |
| rngLoadCombinations | 210 | Семь колонок, включая пользовательский комментарий. |
| rngStabilityDurationLoads | 120 | Формульные ID отмечены как производные значения. |
| rngSP35Table721 | 144 | Численные узлы и заголовки нормативной таблицы. |
| Контрольные точки диаграмм | 104 | Неназванный контроль; solver не читает эту таблицу. |

## Границы Приемки

Реестр различает пользовательский ввод, формулы, нормативные данные и
неприменимые ячейки. Скрытые списки validation и служебные ссылки присутствуют
в полном census, но не выдаются за самостоятельные расчетные настройки.

Для каждого поля сохранены адрес, значение/формула, validation, область
активности, первый маршрут потребителя и имеющаяся ссылка справки. Незавершенные
пункты обозначены `Pending...`/`NotRun`: runtime-default, точный контракт пустого
ввода и ошибки формулы, согласованные единицы, per-key активность, наблюдаемый
эффект и test-ID. До их проверки строка не получает PASS.

Следующий gate: сверка всех полей с генераторами и reader-ами; каждый enum,
Boolean и активный численный параметр проверяется направленным поведением.
Для расчетных селекторов дополнительно нужны active/inactive и чувствительность
к локальному отключению передачи настройки. Служебные столбцы и формулы
проверяются через зависимости и UI, без искусственного ручного ввода.

Для текущего universal LoadPath запланированы пять фиксированных путей и
`Auto` отдельно для Capacity и Formation: двенадцать путевых вариантов.
Исторические прогоны с четырьмя вариантами Formation не являются приемкой
добавленных Mx/My. Нагрузочные проверки контролируют все
ResultComment, их инженерный смысл, отсутствие
дублей, порядок и принадлежность конкретному output-поддереву. Полный перебор
геометрий/режимов, txt-report и save/reopen не подменяется уже выполненным
27-LC направленным тестом.

### Universal LoadPath, 2026-10-03

Фактический census изолированной v108-книги содержит 1064 адреса и 84 validation:
`docs/regression/Audit03/config_field_registry_universal_v108_2026-10-03.json`
и `.csv`. Из исходного знаменателя 1065 удалена одна obsolete-настройка
`SLS.Crack.InitiationLoadPath`; в `rngSystemSettings` теперь 101 поле.
Общая колонка `LoadPath` остается в каждой строке LC, с default/blank `Auto`.
Остальные блоки и исходный пользовательский census выше не переписаны.

Directed Universal v110: 782/0, включая настоящий reader с пустым значением,
formula-empty, Unknown, CVErr и recovery; actual output psi_s/a_crc/Es/status
при Formation BaseFail; Results save/reopen=True. Source/export 102/102.
Этот gate не присваивает blanket PASS всем 1064 адресам. Полные Off/On v110:
22552/0 и 22563/0; negative writer 779/3; formatting 1003/0, deviations=0,
Validate 25/25. Broad v110 CircleUneven/Light/Off выявил 191 failures,
поэтому новая матрица не принята. Исправленный v111 имеет directed Universal
822/0 (включая малые ненулевые компоненты), Crack 880/0, source/export 102/102;
полный повтор и нагрузочная матрица v111 выявили отдельную потерю Auto-history
при инженерном FAIL. Принятый v112: full Off 22780/0 и On 22791/0, directed
On 827/0 с save/reopen; source/export 102/102, formatting 1003/0 и Validate
25/25. Scoped matrix: 8/8 PASS, 3504 случая, 147163/0 assertions. CSV содержит
только восемь current-source логов. VersionFilter проверен на пустом scope
(ошибка, не PASS) и на неизменности нефильтрованного списка исторических логов.
Ни это evidence, ни удаление одного obsolete-поля не закрывают весь K01-K04.
Историческое адресное evidence требует отдельной
проверки актуальности, особенно для смещенных после удаления строки адресов.

Numeric input gate v45 охватывает все 56 consumer/key маршрутов численных
Solver/Capacity настроек в CSectionSolver, CCapacitySolver, Formation и Batch.
Проверены пустое/TODO, нечисловой текст, ошибка Excel, читаемое число и
отсутствующий optional key. Negative: 394/296; positive: 690/0; full Off/On:
9206/0 и 9207/0. Batch проверяется через Execute на полном Config, с InputErr,
отсутствием solve и восстановлением следующего запуска. Это не подменяет
проверку активного численного эффекта/диапазона каждого параметра.

## Направленные Эффекты Solver v47

`solver_effects_v47c_2026-10-02.txt`: 84/0 через реальный трехколоночный
Range -> CSystemSettingsReader -> ApplySettings -> Solve. Тринадцать ключей
Solver проверены отдельно: Method, MaxIterations, LoadSteps, три Tolerance,
LineSearchEnabled, DampingInitial, MinLineSearchAlpha, два MaxDelta,
SecantMaxRestarts и SecantMinStepNorm. Изменяются настоящие ступени, итерации,
принятые шаги, рестарты и компонентные невязки; getter-ов настроек нет.
Двенадцать неверных значений дают typed InvalidConfiguration без итераций.
Secant options отдельно не влияют на Newton. Временный лист удаляется,
исходная книга не меняется. Отрицательная подмена передачи Method выполнена
только в отдельной mutation-книге: `solver_method_mutation_v48b_2026-10-02.txt`,
78/6. Тест обнаруживает отключенную передачу Secant, неверные рестарты и
потерянную валидацию Method. Это ожидаемый отрицательный результат, не PASS
production. Mutation не входит в исходники или выпускную книгу.

Два подготовительных failure-лога не являются дефектами production: первый
fixture ошибочно имел два столбца, второй ожидал роста общего числа evaluations
при line search. Фактические Off/On имели по шесть evaluations, но разные
iterations и реальное дробление шага; oracle теперь проверяет нужную ветвь.

Эта направленная приемка еще не переносится автоматически на все 1065 адресов:
остальные consumers, material/geometry/presentation поля и interactions
по-прежнему требуют своей трассировки и доказательств.

Per-key evidence присоединено к полному поадресному census отдельным
`Merge-Audit03ConfigEvidence.ps1`. Выходной JSON/CSV
`config_behavior_registry_v48_2026-10-02` содержит те же 1065 адресов, из них
13 имеют `ActiveBehaviorAccepted:FullRangeReviewNotComplete`. Для остальных
1052 не добавляется искусственный PASS. `General.DiagramExtension` был
четырнадцатой строкой численного fixture, но не четырнадцатым Solver key;
его доказательства относятся к отдельной migration/On-Off группе.

Фактические внутренние defaults и диапазоны этих 13 полей записаны
по call-site CSectionSolver. Они отличаются от defaults новой книги, например
LoadSteps=5 в чистом объекте и 1 в каталоге. Дополнительный negative v49
`solver_effects_units_negative_v49_2026-10-02.txt` (88/12) подтвердил, что
Abs после CUnitSystem скрывал отрицательные ToleranceN/Mx/My и MaxDeltaKappa.
Исходный знак теперь сохраняется во всех четырех consumers до их валидации.
Тот же направленный тест v50 (100/0) подтверждает реальный typed отказ без
итераций через CSectionSolver.ApplySettings/Solve. Полные Off/On v53 также
прошли. Runtime unit-sign маршруты остальных трех consumers проверены отдельно:
`unit_sign_consumers_negative_v67_2026-10-02.txt` на неизменном production v48
дает 20/24, а `unit_sign_consumers_v64_2026-10-02.txt` на исправленном коде 44/0.
Capacity, Formation и Batch получают отрицательные ToleranceN/Mx/My и
MaxDeltaKappa через фактический Config и CUnitSystem. Проверяются typed
InvalidConfiguration, отсутствие внутренней итерации/retry/formation point,
InputErr в Batch и содержательная причина. Полные Off/On v66 включают этот
тест и дают 9716/0 и 9722/0. Тест восстанавливает все затронутые таблицы.

Выполнен адресный тест 23 редактируемых параметров материалов во всех
16 спецификациях ULS/SLS, TwoLine/ThreeLine бетона/арматуры и Ignore/UseDiagram.
Он читает настоящие таблицы, проверяет ожидаемый активный эффект, неизменность
неактивной диаграммы и другого материала, шесть неверных значений и recovery.
Rb,mc2 проверяется через продольную проверку, а не через диаграмму. Directed
`material_config_v68_2026-10-02.txt`: 1442/0. Full v68 Off 11158/0, On 11164/0.
Положительные значения fixture различают сжатые/растянутые ветви и ULS/SLS;
ошибочный v67 с зарезервированным именем scale не засчитывается. В актуальный
JSON/CSV `config_behavior_registry_v68_2026-10-02` присоединены 36 полей из
1065: 13 Solver и 23 material input. Остальные 1029 не получили PASS.
Приемка активности не равна полному диапазонному покрытию: связанные границы,
все downstream consumers и полная K01-K04 приемка по-прежнему открыты.

## Worst, Общие Селекторы И Единицы v71

`General.WorstCombinationCriterion` имеет четыре согласованных значения.
Непустой неизвестный текст, явный пустой ввод и ошибка Excel не должны
молча превращаться в StrengthCapacity. Подготовленный тест проходит настоящий
Config -> reader -> Batch -> typed reserves -> writer, отдельно проверяет
каждый критерий, ошибочный ввод без solve и восстановление следующего запуска.
Negative v69: 25/45; positive v71: 70/0. В полной v71 Off/On серии тест
сохранен, все восемь suites green (13300/0 и 13306/0), источники неизменны.

Положение и привязка дополнительных рядов RectSet общие для двух сторон каждой
грани H1/B1/H2/B2; диаметры рядов остаются раздельными. В исходной книге нижние
16 dropdown-ячеек дублируют общий выбор, но reader игнорирует их значения.
Подготовлено объединение этих повторных ячеек с верхним общим selector-ом.
Reader необъединенной таблицы разрешает одинаковые значения, пустой follower
или `-`, но отклоняет конфликт с содержательной причиной. Новая независимая
расчетная настройка для каждой стороны не вводится.

Исходный denominator остается 1065 адресов. После фактического обновления
16 адресов должны стать `NotEditable:MergedFollower`, а не исчезнуть из реестра.
При миграции прежние значения записываются отдельно; конфликт проверяется до
изменения первого объединения. Для остальных 761 editable-поля сохраняется
полная проверка значений/формул/validation/формата до, после и save-reopen.
Negative common-selector v69: 160/64; positive v71: 224/0. Фактическая
изолированная книга имеет 16 общих объединений, один dropdown на выбор,
idempotence и save/reopen. Пользовательская output-книга не обновлялась.

Все 15 селекторов единиц/знаков имеют directed invalid/recovery/choice test:
408/0. Подтвержденные defaults явного пустого/TODO выбора устранены без
изменения коэффициентов единиц. Обязательные unit/sign keys полной книги
проверяются reader-ом; отсутствующий key автономного частичного API сохраняет
программный default. Negative логи 273/90 и 282/126 сохранены.

Сквозная серия содержит 72 физически эквивалентных задания (8 знаков x
3 Force INPUT x 3 Moment INPUT), четыре LC каждого варианта, реальные reader,
Batch/State и writer-ы. Все шесть OUTPUT selectors меняются после solve,
не запускают повторный solver; normalized named-state/geometry/properties
совпадают с baseline. Directed 1440/0, save/reopen snapshot SHA совпадает.
Stress/Strain и кривизны сохраняют внутренние оси/знаки; пользовательские
знаки преобразуют усилия, а не материальную диаграмму.

Этот тест не выдается за сквозное покрытие INPUT Length/Stress/Curvature:
их варианты пока проверены в адаптере, downstream geometry/material proof
требует отдельной серии. На этом срезе INPUT Area еще не прослежен до своего
workbook consumer; адаптерный round-trip сам по себе не доказывает активность.
Последующая проверка выявила потребителя AutoCAD.Import.MinArea (см. ниже),
поэтому первоначальное предположение о мертвом INPUT Area не подтверждается.

По запросу пользователя отклонения оформления сразу фиксируются в Progress.
COM negative v71 нашел 52 отклонения input/service/comment ячеек; строгий
snapshot дополнительно выявил четыре служебных прочерка Opening. Positive
v73: 3015/0, 1003 адреса проверены до/после повторного применения и save/reopen.
Данные всего Config и все 1065 field-records, validation, number formats,
merges, заливки и шрифты сохранены. Это не blanket PASS K01-K04 или пиксельная
приемка; чистая сборка с этими исправлениями проверяется отдельно.

`config_behavior_registry_v73_2026-10-02.json/.csv` теперь содержит 64 поля
с адресными доказательствами активности из исходных 1065. Добавлены только
11 Unit/Sign, 1 Worst и 16 общих RectSet selectors; INPUT Length/Area/Stress/
Curvature не получают PASS сверх доказанного адаптерного контракта. Для
сквозной unit-серии merge-tool отдельно проверяет directed supplementary gate
и его наличие в обеих полных suites, а не только наличие unit-choice assertion.
Merged followers не удалены из census. Чистая v73 форматная проверка: 1004/0,
1003 ожидаемых адреса и source unchanged=True. Остальная K01-K04 приемка открыта.

## INPUT Consumers И AutoCAD v77

Новый сквозной INPUT тест выполняет обычный workbook-сценарий: четыре формы,
три длины, пять напряжений, две кривизны, всего 120 вариантов по четыре LC.
Независимое масштабирование исходных ячеек сохраняет физическую постановку;
сравниваются геометрия, элементы, диаграммы, свойства и подробные численные
значения/статусы/ResultComment. Directed v75: 1293/0, Results save/reopen=True.
В части LC Width/PostCrackState неприменимы; это не 120 активных Width-задач.

Первый v74 прогон дал 24 различия только почти нулевого Transformed.Ixy.
Они находятся в пределах машинного округления относительно Ix/Iy. Только
новый test-oracle для Ixy/Ixyc использует bound 8*2^-52*sqrt(Ix*Iy); исторические
expected/tolerance и solver не меняются. Отрицательный лог сохранен.

Full v77b Off/On: 14837/0 и 14843/0, source unchanged=True. Это позволяет
присоединить active evidence для Units.Length.Input и Units.Stress.Input.
Реестр v77 содержит 1065 прежних адресов, active-reviewed=66, fullAcceptance=False.
Units.Curvature.Input проверен по valid/invalid выбору и equivalence, но clamp
в этой серии не доказан как binding. Его active K02 gate остается открытым;
отдельный тест эффекта Solver.MaxDeltaKappa не заменяет combined-unit маршрут.
Units.Area.Input на срезе v77 еще не имеет runtime evidence своего workbook
consumer. Последующая трассировка подтвердила его использование порогом
AutoCAD.Import.MinArea; это не основание удалять настройку как мертвую.

В 72 OUTPUT/sign вариантах реальный AutoCAD export reader возвращает
геометрию в мм, площади в мм2 и локальные инерции в мм4; еще 72 assertions
доказывают отсутствие повторного solve. Это не приемка фактической записи DWG.
Справка в test/update книге объясняет фиксированный масштаб в Units и
AutoCAD.Export.CombinationID. Input data/формулы/validation/format 761 полей
сохранены, 141 ссылка корректна, область печати сохранена после help-update.

## Import Units И Binding Curvature v83

INPUT Curvature теперь имеет отдельный active proof: 132/0. Два solver method,
две оси и обе единицы проходят полный Config adapter; физически одинаковый
binding clamp дает 22/7 итераций вместо свободных 2. Аналитика дискретной
сетки и равновесие проверены независимо, solver tolerance не меняется.

INPUT Area не является мертвой настройкой: consumer - AutoCAD.Import.MinArea.
Общий Config preparation live-import и directed теста один, новых классов нет.
Порог переводится во внутренние мм2, а сами Region остаются мм/мм2/мм4.
Directed 89/0: три единицы, три порога, включительная граница, отсутствие
скрытого default, missing/blank/TODO/CVErr/negative/overflow, полное удаление
бетона/арматуры фильтром и восстановление Config. Исключение адаптера в
отдельном test-call меняет отбор, подтверждая чувствительность теста.

Import -> смена INPUT/OUTPUT -> Calculate: девять переходов mm/cm/m, 228/0.
Сохраненный snapshot читается по своим старым заголовкам и не масштабируется
повторно. Новый snapshot получает новые OUTPUT-единицы. Порог/слои после
импорта не влияют на Calculate, импорт заново не выполняется. Некорректный
OUTPUT отклоняется до solve и сохраняет прежнюю геометрию; recovery успешен.
Непереведенное число при смене INPUT Force N -> kN меняет нагрузку/деформацию
в 1000 раз при неизменной геометрии. Save/reopen сохраняет Results точно.

Фактическая справка v83: явный Region contract в B1488/B1489 - X/Y/размеры мм,
Area мм2, Ix/Iy/Ixy мм4, без определения масштаба INSUNITS. Пользовательские
INPUT/OUTPUT не конвертируют исходный DWG. Сохранены 761 input-records,
141 ссылка и Print_Area. Это не реальный DWG smoke и не пиксельная приемка.

Full v83b Off/On: 15286/0 и 15292/0, source unchanged=True. Реестр сохраняет
1065 адресов, active-reviewed=69, fullAcceptance=False. Остальные per-key,
inactive, full-range и pairwise требования остаются открытыми. Прежние
отрицательные workflow/COM-runner журналы не переписаны и не названы PASS.

## Crack Config v87

Все 12 полей SLS.Crack проверены через фактические ячейки Config, reader,
ApplySettings/Execute, typed results и реальные writers. Пустой ввод, TODO,
ошибка формулы и missing key больше не подменяются начальными значениями
класса. Для пяти чисел отдельно отклоняются 0/-1/overflow. Валидное значение
восстанавливается после каждого отрицательного сценария. Начальные значения
программного конструктора не менялись; переданный Config обязан быть валиден.

Directed v87: 2726/0, 31 валидный вариант и 75 invalid/missing запусков,
три нагрузочные постановки, а не 106 независимых физических сечений.
Phi1/Phi2/Phi3/PsiS проверены отношениями a_crc и независимым произведением
готовых численных данных. Allowable меняет utilization/status без изменения
ширины. User/Auto/AlwaysCalc и активность/неактивность User-параметров имеют
собственные assertions. AllSelected/TensionOnly изменяют sigma_s,crc, но не
текущий sigma_s и выбор стержней; неположительное среднее не создает NumFail
формулы. Effective/FullTension изменяют глубину/Abt и ширину. Для двух
CoverDistanceMode доказано различие подготовленного a_s и выполнение формулы
в обоих режимах; совпадение a_crc не выдается за доказанный эффект этой
настройки на итоговую ширину или самостоятельную проверку writer-а для a_s.

ResultComment каждого варианта проверен по поддереву и всем output-блокам.
Results сохраняется и повторно открывается без изменения значений и comments.
Справка v87 содержит обязательный input-contract и явный Region contract
мм/мм2/мм4; actual help gate failed=0, 761 input records сохранены.

Подготовлен per-key evidence десяти настроек. Два Formation-селектора имеют
invalid/missing/recovery, но отдельная активная per-key приемка не объявлена.
Реестр пока сохраняет 1065 адресов и 69 принятых активных полей; merge новых
десяти допускается только после обоих полных зеленых Off/On и formatting.
Full v86d имел 88 ошибок неполной numeric Formation fixture, исправленной
только в тесте v87; этот отрицательный журнал не является full PASS.

## Понятные Ошибки Обязательного Ввода

В текущем срезе reader сохраняет фактическое место значения в обычной таблице,
таблице материалов и селекторах единиц/знаков. Отсутствующая строка сообщает
раздел Config и действие по восстановлению, без выдуманного адреса. Ошибка
числа/формулы и пустой обязательный ввод сообщают имя настройки, лист/ячейку,
причину и действие. Material validation сохраняет собственные коды и правильно
указывает сжатую либо растянутую ячейку. CUnitSystem перечисляет допустимые
единицы/знаки; инженерные проверки Crack/Stability указывают строку Config и
требуемое исправление, не получают Excel-объектов и не назначают статус по тексту.

Directed v89: 1901/0. Проверены 32 обязательных системных ключа, 23 материальных
alias-ключа, пользовательские единицы/знаки, missing/blank/TODO/abc/formula error
в применимых случаях. Это число assertions, а не число независимых параметров.
Полный workbook-сценарий передал понятную ошибку SLS.Crack.Allowable в сообщение
запуска и настоящий batch-комментарий. Negative v88: 1897/4 из-за неверного
test offset комментария; исправлен тест, не writer. Source v89: 101/101, failed=0.

После этого добавлены сообщения диапазонов Stability и MinArea с неизменными
машинными кодами. Финальные gates выполняются на v90; результаты v89 не
выдаются за приемку этих поздних правок. Help v90: 1930 строк, 141 ссылка,
118 shapes, 761 input records сохранены, Print_Area count=1, failed=0.
Справка различает ошибку, уже оформленную в Results, и остановку до записи
нового расчетного снимка. Все незавершенные K/T/D и выпускные gates сохраняются.

Повтор v93: directed messages 1921/0, включая четыре поврежденных обязательных
именованных таблицы и восстановление их ссылок. Directed Crack Config 2837/0,
Results save/reopen совпадает. Проверка programmatic Initialize/Clone материалов
на v92: 2856/0. v90/v91 Crack timeout был вызван внесенной ошибкой сигнатуры
Validate, исправленной до v92; эти отрицательные журналы не являются NumFail
расчета и не получают PASS. Source v93: 101/101. Полные suites и formatting
на v93 остаются обязательными перед merge десяти активных Crack-настроек.

Крайнее переполнение пересчета добавлено отдельно. Negative v94: 1950/36,
исходные сообщения содержали только Overflow, последний случай прерывался
до оформления результата. После адресной диагностики пересчета и запрета
нормализации LC при ошибочном Config directed v95: 1993/0. Восемь ключей:
Allowable, ElementLength, два User eccentricity, ToleranceN/Mx/My и
ZeroMomentPerDepth. Каждый дает InputErr без solve и проходит восстановление.
Это проверка INPUT-диагностики, а не новая физическая верхняя граница чисел.
Help v95 включает этот контракт; 761 input-record сохранен. Directed Crack
v95: 2837/0, save/reopen=True. Окончательные full Off/On v95: 19722/0 и
19728/0, все восемь suites, watchdog exit=0, source unchanged=True.
Formatting v95: 1004/0, 1003 адреса, deviations=0; все 25 checks Validate True.
Source/export contract: 101/101. Фактический текст о переполнении и исправлении
INPUT прочитан в A22/A23 листа Справка, без заявления о пиксельной приемке.

После этих gates десять Crack-настроек присоединены к реестру v95:
1065 адресов, active-reviewed=79, fullAcceptance=False. Это приемка активного
поведения и обязательного input-contract указанного среза, а не всех диапазонов,
профилей, форм, взаимодействий или двух Formation-селекторов. Отрицательные
журналы не переписаны; оставшиеся K/T/D и выпуск основной книги не закрыты.

## Search-Селекторы И Адреса Ошибок

Четыре ключа Capacity.SolutionStrategy, Capacity.SearchMethod,
SLS.Crack.InitiationLoadPath и SLS.Crack.InitiationSolutionStrategy проверяются
через реальные Config-ячейки, профили, материалы, batch и общий LimitSearch.
Пустое/TODO, неизвестный вариант, ошибка формулы и утраченная строка дают
InputErr до решения НДС; после исправления повторный запуск работает.
Для найденной строки сообщение называет фактическую ячейку, для утраченной
строки - раздел и действие по восстановлению. Writer только выводит готовую
причину, машинный статус не определяется по тексту сообщения.

Directed v102: 1812/0, source unchanged=True, Results save/reopen=True.
Address negative v102: 1810/2 на прежнем production с усиленным тестом;
обе ошибки относились к неизвестным Formation-селекторам без адреса.
Нормализация остается у Formation и переиспользуется Batch до любого solve.
Full Off v102b: 21534/0, все восемь suites. Full On, formatting, четыре
selector mutations и scoped all-path повтор еще выполняются/ожидаются.
До завершения этих gates реестр остается 79/1065, приемка четырех новых
полей и полный Audit03 не объявляются завершенными.

Окончательные gates среза v103 завершены: directed 1812/0; full Off
21534/0 и On 21543/0; formatting 1004/0; Validate 25/25; source/export
102/102. Исправление unit-equivalence fixture передает выбранные единицы
в Batch.ApplySettings; production и исторические comparison tolerances
не менялись. Отдельный On-повтор 72 вариантов: 1584/0.

Четыре selector mutations обнаружены нужными active assertions, соответственно
33/26/30/36 failures, без изменения source. Матрица CircleUneven/HollowThin
Stress Off/On: 4/4 chunks, 960 независимых случаев, failed=0; Results
save/reopen и неизменность source подтверждены. После merge реестр v103
содержит 83 active-reviewed из 1065, fullAcceptance=False. Это приемка данного
среза, а не полного Audit03 или будущего общего LoadPath.

Первый запуск mutation helper остановился из-за UTF-8 без BOM в Windows
PowerShell 5.1 до выполнения VBA. Совместимость кодировки исправлена;
отрицательные книги создавались только после этого, исходный журнал сохранен
в диагностике запуска. Такая ошибка не является численной несходимостью.

## Актуальная Геометрия Circle v133

Фактический census v133: 1064 уникальных адреса, 35617 ячеек Config,
14 именованных блоков, 72 validation. После универсального LoadPath отдельного
SLS.Crack.InitiationLoadPath в текущей книге нет. При переносе по Id из v103
сохранены 82 исторически принятых активных поля, а не прежние 83 с удаленным ключом.
Адреса, значения, формулы, validation и help взяты из актуального census,
не скопированы из прежней раскладки.

Новые восемь полей Circle прошли штатный маршрут Config -> reader -> units ->
registry -> geometry/rebars: directed `53/0`. Проверены активные численные
изменения, оба положения дополнительных рядов, ошибочные значения, recovery,
пустые/нулевые диаметры, count=0/2 и неактивность при RectSet. Формулы трех
входных таблиц восстановлены. Конкретный смысл/границы/evidence записаны в
`NDM_Audit03_Circle_Config_Evidence.json`; непройденные границы и missing/error
контракты не помечены проверенными.

Full v133 Off `23082/0`, On `23093/0`; все восемь suites. Source/export
`103/103`, failed=0; production/test `.cls` остаются `83/3`. Общие численные
assertions с v131 совпадают: 5574 Off и 5579 On, differences=0. Universal
`827/0` и palette `351/0` сохраняют Results после reopen; палитра сохраняет
также оформление. Config formatting `1003/0`, deviations=0; структурные
checks `25/25`; clean/update help `2536/2536`, mergesEqual=True.

После подтверждения обоих full gates registry v133 содержит 90 адресно
принятых активных полей, fullAcceptance=False. Это не полная приемка Config:
range/default/pairwise и остальные семьи сохраняются в текущей цели. Основная
пользовательская книга не заменена промежуточным снимком.

## Некруглые Формы v134

Адресно проверены 6 параметров RectSet, 10 RoundedRectangle, 8 HollowRectangle
и четыре наружных счетчика n. Directed final `213/0` использует реальные
Config/reader/units/registry, активные мутации, недопустимые значения, recovery,
неактивные таблицы/режимы, mm/cm-equivalence и восстановление Formula.
Дополнительные oracle проверяют площадь/центр; дискретизация дуг не меняется.

Negative `161/40` подтвердил молчаливое выключение наружной грани при abc или
CVErr в Hollow n. Ошибка теперь сохраняет key/ячейку и инженерное пояснение;
blank/0 и автоматические счетчики отверстия остаются допустимыми.

Все восемь full suites: Off `23295/0`, On `23306/0`. Общие actual-числа
с v133 совпали `5577`/`5582`, missing=0, differences=0. Source/export `103/103`;
help `2536/2536`, mergesEqual=True; Universal `827/0` и palette `351/0`
сохранены после reopen; formatting `1003/0`, deviations=0; Validate `25/25`.

`NDM_Audit03_Shape_Config_Evidence.json` содержит 28 новых адресных записей,
а `config_behavior_registry_v134_2026-10-03` содержит 118 active-reviewed из
1064. Full-range/default/missing/error-cell для остальных размеров, остальные
поля армирования и pairwise не закрыты этой приемкой. Перенос прежних 90
семантических записей выполнен по точному Id с сохранением текущего census;
он не является повторной blanket приемкой всего Config.

## Ранний Ввод Width v139

Дополнительный active-contract gate проверяет пять режимов и пять числовых
параметров Width на реальном сжатом сочетании с подтвержденным NotCracked.
55 повреждений включают пустоту/TODO/текст/CVErr, неизвестный selector и
нулевые/отрицательные коэффициенты. Прежний production ошибочно принял
20 вариантов: final negative `1515/120`. Positive `1635/0` требует InputErr,
ключ/Config/адрес, содержательное сообщение до solve и последующее восстановление
NotCracked без формулы, с независимой продольной проверкой. Full Off/On
завершены `26264/0` и `26275/0`; общие actual-числа с v137 совпали
в 5601/5606 assertions. Окончательная D/UI приемка среза еще идет;
registry `118/1064` и fullAcceptance=False не изменены.

Это уточнение обязательного ввода, не новый нормативный диапазон и не новая
ветка СП35. Положительные допустимые коэффициенты и aliases не изменены;
инженерную допустимость проверяет Width, не writer и не reader.

## Таблица LC И Запросы Профилей v173

В текущий поадресный реестр присоединены 210 полей таблицы LC и 16 расчетных
переключателей PR1-PR4. Адресный LC-блок: `11021/0`, 910 параметризованных
случаев; profile-scope: `3572/0`, все 64 маски четырех запросов. Это разные
числа: количество assertions не объявляется количеством инженерных задач.
Реальные recovery-ветви используют контролируемые единицы/знаки и проверяют
запуск собственных calculators, а не только чтение флага reader-ом.

Обязательные полные gates окончательной v173 завершены: Off `44045/0`, On
`44056/0`, восемь suites в каждом. Все 9225/9230 общих actual-чисел с v165
совпали, missing/differences=0; допуски и физические expected не менялись.
Ширина 84 столбцов сохраняется при записи/очистке шести consumers:
`1116/0` в каждом режиме. Source/export `104/104`, Validate `27/27`,
formatting `1003/0`, help `2564/2564`, mergesEqual=True. Directed profile
Results и status-style одинаковы после save/reopen; source unchanged=True.

Реестр `config_behavior_registry_profile_scope_v173_2026-10-03` содержит
344 active-reviewed поля из 1064. Все 760 editable адресов присутствуют
в census, но это не означает их полной поведенческой приемки. Остальные
76 material/presentation полей профилей, другие семейства, полный диапазон
и взаимодействия остаются в работе; fullAcceptance=False. Runtime injection
технической остановки orchestration не выполнялся. Журналы отрицательных
v166/v167/v169/v170 и структурного отклонения v171 сохранены, а не переписаны.

## Профильные Модели И Динамические Адреса v176

Новый стандартный test module проверяет все 48 material-spec ячеек PR1-PR4,
четыре stability value-set и четыре изменяемых имени профиля на настоящем
LC/Batch/State/Search маршруте. Directed `3260/0`, 276 случаев, включает
независимые N/Mx/My, actual spec и comments соответствующих writers.
Visualization.State/Quantity и precision 0..10 проверены в snapshot-reader;
все форматированные надписи схемы не приняты этой серией. Description остается
поясняющим metadata, а не придуманным физическим коэффициентом.

Динамические адреса проверены отдельно: 12 named input ranges, другой лист,
две позиции E10/BH800, actual formula errors, конфликты двух общих селекторов
RectSet, ResultComment и writers. Окончательный gate `674/0`, 98 случаев,
сохранение Results/status-style после reopen=True. В VBA сообщения используют
фактическое начало Range и положение поля либо Range.Cells.Address; writer
не восстанавливает адрес по hardcoded Config-строке. После следующего чтения
перемещенная таблица дает новый адрес, старые расчетные comments сохраняют
трассировку того запуска. UI Cut/drag не имитировался, проверена переадресация
имени Excel; все ссылки/формулы восстановлены после серии.

Negative v174b `2924/336`, v174c `318/132` и v175c `514/160` сохранены.
В v176 исправлены потеря geometry/annotation адресов, профильная диагностика,
строгий integer precision и обе ячейки конфликта RectSet. Это не изменения
физических критериев/допусков. Source/export `105/105`, Validate `27/27`,
formatting `1003/0`, deviations=0. Полные восемь suites завершены: Off v176b
`47979/0`, On v176 `47990/0`; общие actual-числа с v173 совпали в 9285/9290
assertions, missing=0, differences=0. Отдельная clean-сборка прошла Validate
`27/27`; вся справка clean/update совпала `2569/2569`, mergesEqual=True.
После обоих full gates присоединены 56 per-key записей. Актуальный реестр
`config_behavior_registry_profile_config_v176_2026-10-03` содержит 400/1064
active-reviewed поля, fullAcceptance=False. Из 760 editable адресов 360 еще
не имеют активной адресной приемки; full-range/pairwise для остальных принятых
полей также не закрыт этим числом.

## Дополнительные Нагрузки Устойчивости v179

Присоединены все 90 editable N/Mx/My ячеек 30 строк duration-таблицы.
Тридцать производных Combination ID не объявляются физическими настройками.
Каждая ячейка проверена через настоящий workbook-reader, CUnitSystem,
Batch/Stability и собственные comments/числа writers. Пустота/ноль допускаются,
непрочитанная таблица, текст, CVErr и переполнение дают адресный InputErr
только устойчивости соответствующего LC; остальные проверки и следующий
корректный LC продолжаются. Две позиции named range на другом листе проверены
по фактическому адресу, не по строке шаблона.

Negative v177b `7926/576` сохранен. Directed v178 `9964/0`, 613 случаев,
включает 390 числовых вариантов, 24 сочетания INPUT units/signs, 180 ошибок
активных ячеек, recovery и структурные края. Full v179 в каждом режиме
дополнительно проверяет comments recovery и содержит `10504/0` этого блока.
Все восемь suites прошли: Off `58483/0`, On `58494/0`; source unchanged=True.
10305/10310 общих actual-чисел с v176 совпали точно, missing/differences=0.
Source/export `105/105`; Validate clean/update `27/27`, formatting `1003/0`,
deviations=0; справка совпала по 2579 ячейкам и всем объединениям.

Актуальный реестр `config_behavior_registry_duration_v179_2026-10-03`
содержит 490/1064 active-reviewed, fullAcceptance=False. Остаются 270 editable
полей без активной адресной приемки. У уже принятых полей full-range/pairwise
и остальные условия K01-K04 не объявляются автоматически закрытыми.
Подробная трассировка: `NDM_Audit03_Duration_Config_Review.md` и per-key
`NDM_Audit03_Duration_Config_Evidence.json`. Завершение clean gates: 2026-10-04.

## Арматурные Поля v192

После negative/positive и полных clean Off/update On присоединены 176
поадресных entries non-circle, из них 156 новых. Проверяются реальные ячейки
Config, независимые нормали/касательные прямых граней, signed смещение as/t,
количество/диаметр и направление дополнительных рядов, selectors/bindings,
активный обязательный ввод, blank/0 выключение, inactive CVErr и recovery.
Opening as остается геометрическим ограничением соседних проекций и при
отсутствии собственного ряда; наружный as выключенной грани не читается.

Negative v186: 7546/1376; v190: 10098/24. Positive v192: 10146/0,
2744 сценария. Circle negative v188: 366/106; positive v189: 524/0,
160 сценариев. Новые fixtures не ослабляют исторические correct expected
или tolerance. Полные восемь suites: Off 70929/0, On 70940/0,
UI 49026/0 в каждом; source unchanged=True. Общие actual-values v185:
Off 13641, On 13646, missing=0/differences=0, без новых допусков.

Source/VBE equality clean/update 105/105 каждая; Validate 27/27 каждая;
formatting 1003/0 каждая, 1002 адреса/deviations=0, без ApplyAlignments.
Palette 355/0 с одинаковыми Results/status-style после save/reopen.
Справка clean/update равна по 2604 непустым ячейкам и merges; все 2681
Config-ячейки/формулы/merges/PrintArea v185 сохранены в updated v192.

Актуальный registry `config_behavior_registry_rebar_v192_2026-10-04`
содержит 646/1064 active-reviewed, fullAcceptance=False. Остаются 114
editable полей: 73 system, 20 profile presentation/description, 21 annotation.
Расширенные вещественные границы, curved fixtures и рискованные пары
армирования остаются обязательными, а не считаются принятыми этим числом.
Трассировка: `NDM_Audit03_Rebar_Config_Review.md` и
`NDM_Audit03_Rebar_Config_Evidence.json`. Основная output-книга, итоговый
Audit03 отчет и полная выпускная приемка пока не обновлены/не завершены.

В актуальном registry также исправлены runtime-contracts восьми Circle-полей:
активный обязательный ввод не подменяется описанием прежних defaults.
Evidence `NDM_Audit03_Circle_Rebar_Input_Evidence.json` использует реальные
53 effect и 524 input assertions и оба full gates. Эти восемь адресов уже
были приняты ранее, поэтому счетчик 646 не увеличивается; исторические
registry сохраняют исходные данные своего среза.

## Профильное Представление v199

После двух полных gates присоединены ровно 20 editable полей PR1-PR4:
Description, Visualization.State, Quantity, StressPrecision и StrainPrecision.
Описание проверено как metadata с blank/text/CVErr/recovery. Все пять states,
Stress/Strain и precision 0..10 проверены через сохраненную плоскость/материалы,
значения всех элементов и фактическую легенду реального Chart, без нового solve
и без округления чисел Results. Полный missing/blank/error контракт selectors,
pixel-приемка всех вариантов и high-risk pairwise еще не объявляются закрытыми.

Directed v198: 8740/0, 276 counted cases. Frozen v199 включает description
и сохранение пробелов аннотации: profile block 8784/0, 292 counted cases;
presentation block 681/0, 105 counted cases. Full Off 77134/0, On v199b 77145/0,
восемь suites каждый; 13641/13646 общих actual-values с v192 совпали точно.
Source/VBE equality 105/105; Validate повтор 27/27; formatting 1003/0,
deviations=0; palette/save-reopen 355/0. First On/structure COM-factory failures
сохранены как ошибки запуска среды и не присваиваются численному result.

Актуальный registry `config_behavior_registry_presentation_v199_2026-10-04`:
666/1064 active-reviewed, fullAcceptance=False. Из 760 editable полей 94 еще
не имеют активной поадресной приемки: 73 system и 21 annotation. Directed
presentation-тест сам по себе не принимает эти 21 annotation-поле и не закрывает
весь Config. Трассировка: `NDM_Audit03_Profile_Presentation_Evidence.json` и
`NDM_Audit03_Presentation_Snapshot_Review.md`. Основная output-книга и final
report остаются невыпущенными до полного DoD Audit03.

## Численные Настройки Capacity v251

К текущему реестру добавлены ровно семь editable полей: SolverMaxIterations,
ToleranceStrain, MaxLambda, InitialLambda, ToleranceLambda, MaxRetries и
BaseLoadSteps. Это реальные named Config -> reader -> batch -> Capacity/Search
проверки, а не только чтение свойства. Проверены два положения таблицы, активные
эффекты, inactive/isolation, 88 invalid cases, recovery и комментарии writers.
Некорректные диапазоны валидирует CCapacityCalculator до Search; сообщение
использует текущий адрес CSystemSettingsReader. Отсутствующие optional numeric
keys сохраняют принятый default-контракт; полная приемка missing keys не заявлена.

Frozen v249: 2318/34, все 34 ошибки относятся к отсутствующему динамическому
адресу. Positive v250: 2804/0; усиленный v251: 2904/0, 154 реальных batch-сценария.
Capacity suite v250: 3074/0; Batch v250: 15775/0. Прежние 601 и 1864 numerical
actual-values совпали с completed v206 точно, missing/differences=0. Последний
срез v251 усиливает только новые тесты; полного On/Off gate он не заменяет.
Help/update/reopen сохраняет 760 inputs; Saved Help failed=0, structure 27/27,
actual source/export 106/106, 111 components.

Registry `config_behavior_registry_capacity_numeric_v251_2026-10-04` содержит
714/1064 active-reviewed. Из 760 editable полей 46 остаются без активной
поадресной приемки, около 6.05%. Full-range/pairwise, экстремальные Double/Long,
переполнение произведения retries/steps, missing-key contracts и metadata
остаются отдельными открытыми проверками: fullAcceptance=False.
Трассировка: `NDM_Audit03_Capacity_Numeric_Config_Review.md` и
`NDM_Audit03_Capacity_Numeric_Config_Evidence.json`. Итоговый выпуск Audit03
еще не завершен; основная output-книга не заменяется промежуточной копией.

## Настройки Устойчивости v257

Добавлены ровно 17 editable полей Stability через настоящий Config -> reader ->
units -> batch -> Calculator -> writers. Два положения именованной таблицы,
обе главные плоскости, SP63 и SP35 table/eta, User/Auto/planes/signs,
активные эффекты и isolation при отключенной устойчивости проверены отдельно.
Equivalent INPUT Length мм -> м сохраняет внутренние длины, Ncr и статус.
Есть 188 invalid values, 34 missing keys и 34 recovery; ошибки содержат ключ,
текущий адрес и действие. Каждый output-блок получает комментарий своего subtree.

Frozen v254: 7680/138; positive v255: 7944/0, 390 реальных batch-сценариев.
Два автономных Calculator проверяются отдельно, не входят в число 390.
Batch v256: 23818/1, единственный отказ прежнего текстового контракта Mu.
Исправлены только две фразы production; тест и numerical expected/tolerance
сохранены. Повторный Batch v257: 23819/0, source unchanged=True; все 1958 общих
numerical actual-values с completed v250 совпали точно, missing/differences=0.
Actual source/VBE 106/106, 111 components; structure 27/27. Help/update/reopen
сохранил 760 inputs; 11 saved-text проверок failed=0. Справка v255/v257 совпала
по 2716 непустым ячейкам и объединениям, source unchanged=True.

Registry `config_behavior_registry_stability_v257_2026-10-04`: 731/1064
active-reviewed. Из 760 editable полей 29 еще не имеют активной поадресной
приемки (3.82%): девять общих/geometry/mesh/load-reference и двадцать AutoCAD.
Extreme Double, переполнение mu*L/L^2, полный диапазон и рискованные пары
устойчивости остаются открытыми; нормативная допустимость произвольного
положительного коэффициента не утверждается программной проверкой диапазона.
FullAcceptance=False. Полный final eight-suite On/Off, нагрузочная матрица,
benchmarks, help/UI/clean-update и self-audit обязательны до выпуска.
Трассировка: `NDM_Audit03_Stability_Config_Review.md` и
`NDM_Audit03_Stability_Config_Evidence.json`. Основная пользовательская книга
и входное ТЗ не заменялись; Audit03 не объявляется завершенным этим срезом.

## Служебные И Табличные Поля v300-v302

Свежий census clean v302 сохраняет 1064 уникальных адреса и 84 validation:
760 UserInput, 144 SP35 table values, 104 diagram controls, 30 derived IDs,
16 merged followers и десять NotApplicable cells. Ни одна из этих групп не
исключена из знаменателя. Поадресные evidence views пока отдельные; наличие
runtime evidence не означает full-range/pairwise/final-release acceptance.

- Actual metadata v300b `209/0`: 30 ID-formulas следуют LC, включая blank и
  перенос named range; 16 followers имеют один ввод/validation у merged anchor.
- Diagram controls v299 `1096/0` и полный v300 Off/On подтверждают пять Stress
  INPUT units, 52 physical-provider points, 88 direct formulas и 16 zeros.
  Это 104 производные контрольные клетки, не независимые настройки solver.
- SP35 v302: 144 адресных behavior cases `829/0`, 540 invalid variants
  `1621/0`, node order/missing-name/relocation/comment/save-reopen `105/0`.
  36 reference cells не меняют расчета; 108 активных table cells проверены
  по независимой интерполяции. 576 чисел точно совпали с frozen baseline.
  Normative trace отдельна: positive validation не утверждает произвольный
  положительный phi как нормативно допустимый.
- Ten inactive cells v302 `102/0`: число/text/CVErr/Empty/NA formula не входят
  в normalized payload. Исправлено прежнее чтение ошибки Rb,mc2.Tension.

Views: `config_metadata_evidence_v300_2026-10-05`,
`config_controls_evidence_v300_2026-10-05`,
`config_cad_controls_evidence_v300_2026-10-05`,
`config_table_evidence_v302_2026-10-05` (JSON/CSV в regression/Audit03).
Каждый helper отвергает incomplete/failed runtime. Актуальные final
matrices/full suites остаются отдельными незавершенными gates.

## Единый Directed Evidence Registry v306

`config_directed_evidence_combined_v306_2026-10-05.json/.csv` объединяет
ровно 1064 уникальных поля актуального registry без изменения его SavedValue,
формул или структурных признаков. Сверены идентификаторы, адреса, роли,
единый знаменатель и существование каждого файла runtime evidence.

| Группа | Число | Граница Подтвержденного |
| --- | ---: | --- |
| Active inputs | 760 | Активный directed effect и соответствующие invalid/isolation gates; полный диапазон и все взаимодействия еще не приняты. |
| Diagram controls | 104 | Прямые формулы, provider points, пять INPUT Stress units и подписи. |
| Active SP35 table cells | 108 | Адресные эффекты, independent interpolation, invalid variants и динамические адреса. |
| Reference SP35 cells | 36 | Не меняют расчет; не трактуются как активные phi. |
| Derived IDs | 30 | Следуют LC, включая blank и перенос named range. |
| Merged followers | 16 | Единственный ввод и validation у anchor. |
| Inactive consumers | 10 | Неприменимые значения не вмешиваются в normalized payload. |

Итого directed address coverage: 1064/1064. `FullAcceptance=False`
сохранен явно: selector runtime, широкая актуальная матрица, full-range,
mutation sensitivity, release/update/self-audit не подменяются этой суммой.

## Матрица Селекторов v306

Полный runner завершил 62/62 группы: 4464 независимых путевых случая,
302 покрытые пары и 48 рискованных четверок. Машинная сводка:
`docs/regression/Audit03/selector_summary_v306_full_2026-10-05.json`;
`FullSelectorAcceptance=True`, source unchanged=True, exit=0.
Приемка относится к перечисленным селекторам/нагрузкам и численному срезу
v306. Она не означает исчерпывающую проверку всех Double, замену mutation
gates или полную приемку Audit03; общий `FullAcceptance=False` сохраняется.
