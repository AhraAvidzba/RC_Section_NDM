# NDM Audit02: Итоговый Отчет

Дата: 2026-10-01. ТЗ: `NDM_Audit02_Implementation_Spec_2026-10-01.md`, версия 1.0.

## Итог

Все применимые требования Audit02 реализованы и подтверждены. Полная независимая
сборка, восемь штатных VBA suites при Off и On, штатный Run-AllTests для output-книги,
направленная матрица, численные пары, сохранение/повторное открытие Results,
проверка вывода и измерение производительности выполнены. В каждом итоговом полном
прогоне **6745 успешных assertions, 0 ошибок**. Это число assertions, не число
независимых инженерных задач.

Реальный DWG-export не выполнен: AutoCAD COM ProgID отсутствует. Это ограничение
доступной среды, не засчитанное как успешный smoke. Фактический reader экспорта,
сериализованные данные и Excel plot проверены без нового solve. Отдельная визуальная
проверка пикселей экрана также не засчитана: native capture дважды завершился timeout.

Закрытие этого ТЗ не означает окончательной нормативной верификации по СП или
гарантии отсутствия любых будущих дефектов. Ограничения `AGENTS.md` остаются в силе.
Дополнительные архитектурные задачи за пределами Audit02 не начинались.

## Git И Артефакты

- Baseline: `855626e69b07f21dc636771416c501097d077ec9`.
- Итоговый проверенный implementation SHA: `a0e44989d59c7e68511e4078e60b70bba3a20162`.
- Ветка: `codex/material-diagram-architecture`. Push, reset, stash, переписывание
  истории и откат пользовательских изменений не выполнялись.
- Последующий закрывающий commit содержит этот отчет, обновленный progress и
  финальные сохраненные артефакты. Он не меняет проверенную production-математику;
  его SHA определяется `git log -1` после сохранения отчета.
- Входные dirty report/export/spec перечислены в progress; входной экспорт сохранен
  историей checkpoint-ов. Baseline запускался в отдельной read-only копии.

| Checkpoint | Содержание |
| --- | --- |
| `799bcd8e` | State fixes и глобальная настройка Extension |
| `ff9e2d47` | Независимая Formation, подтвержденные физические границы |
| `92d0a4ae` | Самодостаточный Search snapshot и общий LoadMultiplier |
| `8929d5d1` | Канонические typed results и блочная запись вывода |
| `651e4d91` | Scoped state reuse и допуски приемки |
| `020ff8f9` | Запасы typed results, физические On/Off пары |
| `a0e44989` | Полная матрица, вывод, snapshots, performance и приемка |

Итоговые файлы:

| Артефакт | SHA-256 |
| --- | --- |
| `workbook/output/RC_Section_NDM.xlsm` после обычного расчета сохраненного LC | `AAF5D4FAF06197F01966463DE85BA813EB4216B455AA8504B2DE4F0BA717C012` |
| `workbook/output/VBA_All_Code.txt` | `51C20ABF981EB9AD569F23648E8ABE846E7ECE05637D8043098C155BC8F6F185` |
| `docs/regression/Audit02/RC_Section_NDM_clean_final.xlsm` | `7FBCB89042F0A148D79C88AD1919375B80DE28BD0D5B8DDF2DD15C54CFAFAE92` |

Output-книга обновлялась через `Refresh-VbaModules.ps1`, не полной заменой Config.
После тестов обычный расчет ее текущих данных сохранен через независимую копию:
все значения и формулы Config до/после идентичны. Results и txt-report обновлены
одним расчетом. Код не менялся после полных suites. 103 импортированных модуля
src/tests соответствуют исходникам с учетом штатного представления VBE: ANSI
CP1251, регистр идентификаторов, окончания строк и запись Double literals.

Присланный экспорт из ТЗ не подменяет baseline: реальные размеры source-классов
измерены отдельно. Счетчики строк ниже включают source headers, поэтому отличаются
от `LINES` исходного VBA-export.

## Карта Ответственности

| Владелец | Итоговая ответственность и граница |
| --- | --- |
| `CBatchSectionCalculator` | Порядок LC/этапов, контекст/профили, запуск владельцев, сбор typed results, выбор governing LC. Не содержит width-формулу или bracket/Newton search loop. |
| `CCrackFormationCalculator` | Полная formation-задача, material spec, criterion, load paths/Auto, Pre/Post states, интерпретация и snapshot результата. Не зависит от Width. |
| `CCrackWidthCalculator` | Подготовка tension zone/стержней/Abt/As/ds/ls/напряжений/Psi по готовым физическим states. Не решает равновесие и не ищет Formation. |
| `CCrackWidthFormulaCalculator` | Чистая численная формула a_crc; не получает State/status и не выбирает зону. |
| `CCapacityCalculator` | Постановка Capacity, material spec, путь/стратегия, результат проверки заданного LC. |
| `CCapacitySolver` | Capacity-specific criterion/probe, локальный кэш/warm-start, residual/guess и подтвержденная специальная финализация. Общие численные search loops не дублируются. |
| `CCapacityLimitSearchProblem`, `CCrackLimitSearchProblem` | Явные адаптеры доменной задачи к ILimitSearchProblem. Не являются второй реализацией search. |
| `CLoadMultiplierSearch`, `CUltimateStrainSearch`, coordinator | Общая скобка/recovery/Brent/Secant/Bisection и общий Newton. Физический путь Auto выбирает Formation, не generic Search. |
| `CLimitSearchResult` | Независимые meta/method/diagnostics и найденная точка/state; нет живого domain calculator или LastSolver. |
| `CStateSolutionRunner`, `CSectionSolver` | Единый обычный state solve; низкоуровневый Evaluate используется только для plane/residual/guess, не как доказательство равновесия. |
| `CStateProvider`, `CStateRepository` | Scoped контекст, эквивалентность, reuse конечных states, восстановление snapshot без solve; search probes в repository не записываются. |
| `CSectionStateResult` | Единственный численный снимок НДС, физические признаки и provenance. |
| `CCombinationResult` и typed branches | Каноническое дерево результатов, агрегированные meta/comments. Нет параллельного mutable flat НДС. |
| `CExecutionReport` и writers | Отчет/units/layout/formatting готовых результатов. Writers не назначают инженерный исход и не склеивают ResultComment. |

Политика определения статуса сохранена:

```text
InternalStatus + ResultCode + ResultKind
    -> CResultStatusPolicy
    -> ExternalStatus
    -> summary / detailed block / report / saved output
```

Текст `StopReason` не парсится для назначения machine code или статуса. Разбор
machine-ключа критического элемента для определения вида деформационного критерия
не является разбором пользовательского комментария. Дедупликация текста комментария
в result-tree и удаление повторных точек в display-тексте также не назначают статус.

## Метрики До/После

Источник: `architecture_metrics_2026-10-01.json`; скрипт
`Measure-Audit02Architecture.ps1`. Methods включают процедуры, properties и
тривиальные API. Dependencies означают число упомянутых типов классов.

| Класс | Строки | Методы | Indexed scalar API | Public scalar fields | Зависимости |
| --- | ---: | ---: | ---: | ---: | ---: |
| Batch | 3283 -> 2044 | 334 -> 143 | 180 -> 32 | 0 -> 0 | 32 -> 30 |
| Width | 2842 -> 1270 | 206 -> 91 | 0 -> 0 | 0 -> 0 | 21 -> 10 |
| CapacitySolver | 1817 -> 1910 | 143 -> 146 | 0 -> 0 | 0 -> 0 | 10 -> 11 |
| CombinationResult | 1116 -> 545 | 60 -> 56 | 0 -> 0 | 138 -> 2 | 15 -> 11 |
| Formation | 42 -> 2410 | 1 -> 147 | 0 -> 0 | 0 -> 0 | 7 -> 23 |
| ExecutionReport | 171 -> 648 | 15 -> 42 | 0 -> 0 | 0 -> 0 | 1 -> 7 |
| LoadMultiplierSearch | 783 -> 603 | 34 -> 28 | 0 -> 0 | 0 -> 0 | 4 -> 5 |
| LimitSearchResult | 275 -> 180 | 27 -> 21 | 0 -> 0 | 0 -> 0 | 4 -> 2 |
| DirectStateResult | 154 -> 144 | 6 -> 13 | 0 -> 0 | 20 -> 0 | 4 -> 4 |

CapacitySolver не стал меньше по строкам. Его граница стала чище: независимые
generic search loops удалены, но добавлены реальные проверки физической точки,
sign-specific limits, диагностируемая finalization и комментарии. Не заявляется
фиктивное уменьшение всех четырех классов. Formation вырос потому, что перестал
быть оболочкой: owns formation context и критерий, а не перенесенный helper Width.

У CombinationResult ObjectReadAssignments уменьшились 190 -> 14, у DirectStateResult
36 -> 8. Это диагностический текстовый счетчик, не самостоятельное доказательство
отсутствия дублей. По чтению кода 20 численных public fields DirectState удалены:
getters читают один CSectionStateResult. У CombinationResult остались только
ProfileId и effective DiagramExtension context, не копии физического НДС.

Новых/удаленных production `.cls` нет: 85 -> 85. Общие solver/geometry/plotter/writer
классы не дробились по размеру. Новый код проверки расположен в test sections;
PowerShell scripts воспроизводят приемку, не являются production helper-монолитами.

## Удаленный API И Оставшиеся Границы

- Удалены массовые indexed result-getters Batch; consumers читают `ResultAt(index)`
  и соответствующую typed branch. Источник запасов находится в typed result.
- Удалены flat numerical/meta/status копии CombinationResult и их sync-механизмы;
  Store-методы сохраняют готовый объект, а не второй mutable набор чисел.
- Удалены отдельное НДС DirectStateResult/StoreSolver и cached display-статусы;
  статус вычисляется policy из meta.
- Удалены `DomainContext`, downcasts в generic Search, domain-specific capture
  в Search-result и прежние раздельные ExecuteCapacity/ExecuteCrackFormation loops.
- Formation facade над Width и formation/search state из Width удалены.
- У Batch удалены шесть неиспользуемых private limit/spec helpers; Strength-limit
  API для отображения фактически выбранного material spec сохранен.
- 32 remaining indexed scalar API Batch относятся к входам, requested flags,
  reference transformations и характеристикам выбранной модели. Это не прежняя
  массовая выдача всех результатов LC.
- Сохранен `CCapacityLoadPath`: инженерный разбор пользовательского Capacity path,
  не generic численный механизм. Generic траектория обоих доменов - CLoadPathVector.
- Object material API, симметричный setter SteelStrainLimit для явно симметричных
  низкоуровневых тестов и specialized plane operations сохранены как рабочие
  контракты, не переходные дубли. Production spec передает два независимых предела.
- Старый ключ Extension существует только в reader/physical migration, help об
  обновлении старой книги и baseline/migration tests. Runtime его не читает.
- Плоские сериализованные таблицы Results оставлены: это output-схема, не flat
  дубли доменной модели.

## Extension И Физическая Приемка

Канонический ключ `General.DiagramExtension=Yes/No`, новая книга default Yes.
Migration выполняется до defaults: old-only No сохраняется; существующий new
имеет приоритет; конфликт выдает warning; некорректное явное значение - ошибка
настройки. Проверены отсутствие обоих ключей, повторность, validation/help и
сохранение соседних Config данных.

| Маршрут | Effective модель |
| --- | --- |
| StrengthState / Capacity offset, probes, finalization | Strength spec, общий Extension |
| CrackFormation / BeforeMcrcState | CrackInitiation spec, включая активный tensile concrete |
| AfterMcrcState / Current CrackedState | Cracked spec; Ignore не включает tensile concrete |
| Ordinary retry/guess | Только реально extended материалы; Off не включает скрытое продолжение |
| Stability/аналитические характеристики | Новые state solves не добавлялись; нормативные модули и коэффициенты не менялись |

Использован прежний технический закон порядка 0.01E. Физические узлы, касательная,
плато и eps_ult сохранены. Для физических пределов за +/-10 техническая граница
безопасно находится дальше: directed тест 12/14 -> technical 24/28. Overflow
выдается явно, не превращается в clamp/ложный физический предел.

Converged, WithinPhysicalRange, ExtensionUsed и effective permission независимы.
EvaluateStrainPlane не выставляет Converged. ConfirmEquilibrium проверяет усилия
без нового solve. Конечная точка Capacity/Formation должна удовлетворять исходному
физическому критерию, нагрузке, пути и прежним допускам. Auxiliary point не является
дополнительной несущей способностью. Failed iterate используется максимум как
численная подсказка, не как доказательство физического отказа.

### Проверка NumFail / BaseFail

Реальный VBA batch: 300x200 мм, mesh 30x20 мм, четыре стержня d20, заданный
path lambda*Mxy, постоянная сжимающая N. Два независимых overload-сценария:

| Сценарий | N внутри, Н | Mx, Н*мм | Off | On | On code |
| --- | ---: | ---: | --- | --- | --- |
| OFFSET_2MN | -2000000 | 8000000 | NumFail | BaseFail | rcInitialStateBeyondLimit |
| OFFSET_3MN | -3000000 | 8000000 | NumFail | BaseFail | rcInitialStateBeyondLimit |

Off: равновесие исходной точки не найдено, касательная матрица вырождена,
`rsNumericalFailure/rcNumericalFailure`. On: найдено вспомогательное равновесие,
подтвержден выход за физический предел уже от постоянной части пути;
`rsCheckFailed/rcInitialStateBeyondLimit`, Capacity display BaseFail. Точка несущей
не выдумывается. Проверены typed result, summary, подробный блок, overall и цвет.

Это не массовая замена NumFail на BaseFail: при неполученном равновесии и On остается
численная ошибка. Для точной аналитической Offset-границы классифицируется реально
подтвержденная плоскость при прежних допусках. Округление чуть ниже utilization=1
не объявляется превышением по одному входному N. Сам физический критерий не менялся.

В доступной сохраненной output-книге в конце задачи **одно LC=1**, не набор со старых
скриншотов. Оно пересчитано в двух копиях без замены геометрии/материалов/нагрузок:
Capacity OK при Off и On; summary/detail совпадают. Это отдельный проверенный факт,
а не утверждение о прогоне отсутствующего в текущем файле набора перегрузок.

## Сборка И Тесты

Все Excel/COM прогоны выполнялись последовательно, реальными макросами книги.
Статические проверки не засчитываются как выполнение VBA.

| Suite | Baseline | Final Off | Final On | Output Run-AllTests |
| --- | ---: | ---: | ---: | ---: |
| Geometry | 496/0 | 496/0 | 496/0 | 496/0 |
| Materials | 54/0 | 1359/0 | 1359/0 | 1359/0 |
| SectionSolver | 390/0 | 450/0 | 450/0 | 450/0 |
| Capacity | 934/0 | 1209/0 | 1209/0 | 1209/0 |
| Crack | 328/0 | 383/0 | 383/0 | 383/0 |
| Batch | 659/0 | 2235/0 | 2235/0 | 2235/0 |
| Workbook UI | 308/0 | 574/0 | 574/0 | 574/0 |
| Regression | 39/0 | 39/0 | 39/0 | 39/0 |
| Всего | 3208/0 | 6745/0 | 6745/0 | 6745/0 |

Baseline default и baseline explicit Off сохранены в отдельных immutable audit
логах. Исторические explicit-On setups в Off-прогоне сохранены; каждой suite
открывается свежая Off-копия и изменения setup отбрасываются при закрытии.

Сокращения доказательств в следующих таблицах:

- **B**: `baseline_all_tests_2026-10-01.txt`, `baseline_off_all_tests_2026-10-01.txt`.
- **F**: `final_verified_off_2026-10-01.txt`, `final_verified_on_2026-10-01.txt`.
- **O**: `final_output_all_tests_2026-10-01.txt`.
- **M**: `diagram_extension_migration_2026-10-01.txt` (179 assertions), reader tests
  в F (28 assertions).
- **SN**: `snapshot_save_reopen_accepted_2026-10-01.txt` и UI saved snapshot tests F/O.
- **P**: `performance_confirmed_2026-10-01.txt`.
- **SC**: `final_source_contracts_accepted_2026-10-01.txt` и финальный saved-source log.
- **EP**: `expected_preservation_accepted_2026-10-01.txt`.

Все эти логи находятся в `docs/regression/Audit02/`. Негативные промежуточные
логи сохранены и не являются приемкой: ошибки переноса/компиляции, временный рост
памяти UI, ошибки setup новых граничных тестов и первоначальных PS readers
исправлены до F. Новый exact-offset тест уточнялся по фактическому solver contract;
старые expected/tolerance и production-пределы не менялись.

EP: 247 исходных test-процедур сохранены. 436 AssertClose expected/tolerance
идентичны; 16 имеют только явно проверенную замену прежнего accessor на typed
источник, все tolerance прежние. Полные suites дополнительно подтверждают actual
числа и исторические ветви, а не только наличие методов.

### Численные On/Off Пары

Семь LC (ULS biaxial/x/y/axial, SLS bending/axial/no-crack) в двух фазах Stability.
Одинаковые входы, геометрия, material spec и reference. Все применимые метрики
сравниваются автоматически; полный MAXDIFF по каждой метрике находится в F.

| Метрика | Максимальное абсолютное отклонение |
| --- | ---: |
| Strength/Current CrackedState, reported усилия/плоскости/stress | 0 на точности журнала |
| Capacity lambda | 1.63e-10 |
| Capacity N | 1.2594508e-5 Н |
| Capacity Mx | 6.50160015e-4 Н*мм |
| Capacity My | 6.5332279e-5 Н*мм |
| Capacity max concrete strain | 1e-12 |
| Capacity min steel stress | 3.981e-9 МПа |
| Formation Mcrc | 3.725e-9 Н*мм |
| PostCrackState Mxint | 4.657e-9 Н*мм |
| Width/Psi/geometry/Longitudinal/Stability/reference, reported | 0 на точности журнала |

Нули здесь не означают обязательного bitwise совпадения Double: журнал округляет
малые отклонения. Assertions используют прежние допуски (например, силы 5 Н,
моменты НДС 5000 Н*мм), не точность печати. Все проверки прошли.

### Команды Воспроизведения

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Build-Workbook.ps1 -OutputPath docs/regression/Audit02/RC_Section_NDM_clean_final.xlsm
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-Audit02ModeTests.ps1 -SourceWorkbook docs/regression/Audit02/RC_Section_NDM_clean_final.xlsm -Mode No -ReportPath docs/regression/Audit02/recheck_off.txt
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-Audit02ModeTests.ps1 -SourceWorkbook docs/regression/Audit02/RC_Section_NDM_clean_final.xlsm -Mode Yes -ReportPath docs/regression/Audit02/recheck_on.txt
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-AllTests.ps1 -ReportPath docs/regression/Audit02/recheck_output.txt
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Test-Audit02ExpectedPreservation.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Test-Audit02FinalSourceContracts.ps1
```

Фактические exit codes: full Build 0; clean Off 0; clean On 0; output refresh 0;
Run-AllTests 0; clean/output structural validation 0; migration 0; SN 0;
performance 0; EP 0; SC 0; saved LC pair 0; ordinary final snapshot save 0.
BaselineRoot для двух последних verifier-ов и performance можно передать явным
параметром, используя безопасную копию baseline commit. Не запускать полную
пересборку поверх пользовательской Config; для обновления VBA использовать refresh.

## Производительность

Одна среда, по три повтора baseline/current, одинаковые исходные сценарии,
источники не менялись. Времена ниже - медиана секунд. Solve counts фактические
из modSolverWorkStats; EvaluateStrainPlane не считается новым equilibrium solve.

| Сценарий | Solve до/после | Probes до/после | Медиана времени до/после, с |
| --- | ---: | ---: | ---: |
| bending | 30 -> 28 | 24 -> 24 | 0.2891 -> 0.1484 |
| biaxial | 48 -> 48 | 48 -> 48 | 0.2969 -> 0.2969 |
| asymmetric | 20 -> 94 | 20 -> 23 | 2.4766 -> 15.2813 |
| axial | 30 -> 38 | 32 -> 12 | 0.1172 -> 0.1875 |
| methods | 31 -> 30 | 32 -> 32 | 0.1953 -> 0.1719 |
| crack bending | 73 -> 72 | 66 -> 66 | 0.4141 -> 0.3984 |
| crack paths | 56 -> 56 | 44 -> 44 | 0.3359 -> 0.3750 |
| crack cache | 50 -> 50 | 44 -> 44 | 0.2734 -> 0.3125 |

Рост в asymmetric реальный и явно не скрыт. Baseline в этом кейсе трижды выдавал
physical outcome по несошедшейся probe; current выполняет дополнительные retries
до подтвержденного физического результата (unconfirmedPhysical 3 -> 0).
Это цена исправления R03, не исчезновение кэша. В осевом кейсе unconfirmedPhysical
9 -> 0; ранняя проверенная pure-axial finalization снизила первоначальные 501 solve
до 38. Она подтверждает forces/path/original criterion, а не clamps strains.

Обычные bending/biaxial/methods/crack scenarios не получили роста повторных solve.
Methods probeReuse=5 и crack-cache stateReuse=1 сохранены в обоих вариантах.
Изменение времени малых кейсов не интерпретируется как универсальный процент.
Сравнивались LoadMultiplier representative cases; performance Ultimate не
заявляется отдельно, но его численная приемка входит в полный matrix suite.

## Вывод И Сохраненный Snapshot

- Summary, подробные strength/crack/stability blocks, reserves, governing LC,
  пустые строки исходного LC-range, seven display statuses и comments проверены F/O.
- Block comments приходят из соответствующего result-subtree; batch содержит
  comments всех разделов. Warning не теряется при search/result packaging.
- Успешный отрицательный ответ «трещины нет» не создает fictitious Post; SearchBound
  не становится доказанной недостижимостью. Width/Long blocked при bad Current.
- Saved material snapshot не зависит от будущего Config. Семь настроек материала
  изменены, книга закрыта/открыта; 38950 символов serialized raw/read snapshot
  совпали полностью. Actual export-reader и plot выполнили **0 equilibrium solve**.
- Пользовательский current LC пересчитан при Off/On в копиях; исходник неизменен.
  После suites обычный final calculation сохранен, Config values/formulas неизменны.
- `final_environment_limits_2026-10-01.txt` фиксирует AutoCAD и screen-capture limits.
  Чтение actual export state проверено; запись DWG/визуальные pixels не засчитаны.

## Трассировка Требований

Во всех строках ниже результат **проверено**, если явно не указано ограничение
доступной среды. Q05-01...Q05-28 имеют отдельные реальные test-ID в Directed Coverage
progress; все строки выполнены в F, кроме физических migration/save-reopen с M/SN.

| ID | Файл/метод или решение | Test-ID / доказательство |
| --- | --- | --- |
| R01 | StateProvider.SolverSnapshot / Runner.RestoreSnapshot; Width получает State | audit02.currentCache.*, F/O |
| R02 | LoadMultiplier FindBracket/MaxLambda, Capacity SearchBound | audit02.boundary.*, audit02.genericMultiplier.*, F |
| R03 | Capacity ProbeOnce и подтвержденная finalization | failedProbe/failedPositive/offsetPair, F/P |
| R04 | ILimitSearchProblem, generic Search, независимый SearchResult | genericUltimate/genericMultiplier/searchEngineering.*, F/SC |
| R05 | Formation.CheckFormation owns context/search/states | TestAudit02IndependentFormation, formation.constant.*, F |
| R06 | Runner physical flags, State Initialize/Confirm | audit02.plane.*, stateSnapshot.*, offsetPair.*, F |
| R07 | Canonical Combination/Direct/typed results; migrated writers | audit02.canonical.*, tree display/reset, F/O/метрики |
| R08 | Capacity sign-specific SteelCompression/TensionLimit | audit02.steelSign.*, F |
| R09 | Sequential Nothing guards state/search/optional output/registry | generic contracts, empty state, zero-rebar UI, F/SC/read audit |
| A01 | Existing owners, no new production classes | transfer map, 85 -> 85, SC/метрики |
| A02 | Batch ResultAt/orchestration; report to ExecutionReport; typed reserves | canonical/reset/governing/summary tests, F/O |
| A03 | Formation separate; Width preparation; pure formula | independentFormation/formulaPure, F/SC |
| A04 | Capacity domain callbacks; ordinary probes common runner | steelSign/generic matrix, F, ProbeOnce read audit |
| A05 | Shared FindBracket/RecoverBracket/search loops and RunNewton | genericMultiplier both kinds x methods, genericUltimate, F/SC |
| A06 | SearchResult immutable snapshot, no live solver | searchEngineering.metaSnapshot/capacityState, F |
| A07 | Typed results only; no parallel numerical fields | canonical results/reset, returned snapshot isolation, F/O/метрики |
| A08 | Existing writers preserved; passive statuses/comments | batch writer/reserves/status dictionary, F/O |
| E01 | Reader/provider/SettingsCatalog General key | migration.reader.*, M/F/validation |
| E02 | Physical migration before defaults, No preservation/idempotence | migration.*, M |
| E03 | Effective models all applicable state/search routes | multi-role pair and strategy/path matrix, F |
| E04 | Ignore keeps sigma/Et=0; active tensile/steel extended | material.*.ignored* and diagram pairs, F |
| E05 | Physical nodes/plateau/boundaries unchanged | AssertPhysicalDiagramPair and Q04, F/EP |
| E06 | Existing 0.01E mechanism, safe outer limit/overflow | technical.*, F |
| E07 | Independent converged/physical/used/permission facts | extended start -> physical final, offsetPair, F |
| E08 | Auxiliary not accepted as physical capacity/crack state | offsetPair/formation.physicalBlock/current blocked, F |
| E09 | Same physical criteria for all strategies | capacity/crack method matrix, F |
| E10 | No hidden extension at Off | baseline Off + current Off + preserved explicit-On, B/F/EP |
| S01 | ResultMeta enums/code and ResultStatusPolicy | statusDictionary/tree display + no text mapping read audit, F/SC |
| S02 | Capacity initial Offset classification | initialLambda/offset/offsetPair, F |
| S03 | Initial crack fixed/Auto code/warning/psi1 | formation.constant.*, pathAuto, psi.fallback.*, F |
| S04 | CriterionNotReached/SearchBound/above-current distinct | formation.noCrack/searchBound/aboveCurrent, F |
| S05 | Bad Current primary cause + blocked formula dependencies | aggregate current failure/explicit-On beyondphysical, F/O |
| S06 | Existing Psi User/Auto/AlwaysCalc retained | psi signed/fallback/normal tests, F |
| S07 | Same meta/comments in all visible layers | tree/summary/output comments, offsetPair sheet, F/O |
| C01 | Runner shared, Object linear material API | linear equilibrium/replacement/steelSign, F |
| C02 | Context identity/revision/spec/permission/tolerance | audit02.cache.*, F |
| C03 | Named repository vs local probes; failed retry permitted | cache.retry/Pre/Post and preserved probeReuse, F/P |
| C04 | Converged auxiliary distinct from acceptable | cache reuse semantics + current blocked/offsetPair, F |
| C05 | Saved Results self-contained, no hidden recalc | saved export snapshot + Config/save/reopen, F/O/SN |
| C06 | Clear/Initialize/clone/meta and next solve isolation | canonical/stateSnapshot/generic reinitialize/RepeatedRun, F |
| Q01 | Safe baseline SHA/build/default+Off logs | B, baseline source hash, progress |
| Q02 | Original Off numerical expected retained | F-Off, EP |
| Q03 | Explicit-On setups retained after key migration | progression tests/mode override logs, F/EP |
| Q04 | Identical physical input pairs, MAXDIFF each metric | TestAudit02OnOffPhysicalResults, F |
| Q05 | All 28 rows traced in progress, no untested row | Directed Coverage + F/M/SN |
| Q06 | Old-T/Old-A preserved and extended | table below, F |
| Q07 | All suites + preservation review, negative logs kept | F/O, 247 procedures + 452 expected/tolerance, EP |
| Q08 | Output/cache/snapshots/performance | F/O/SN/P; unavailable DWG/pixel checks explicitly excluded |
| D01 | Architecture/Config/help/current API updated; history marked | Architecture/SettingsCatalog/historical document headers, M/F |
| D02 | General permission vs use vs physical limit explained | SettingsCatalog Extension user-guide/help/validation, M/F |
| W01 | Baseline/Git and user dirty state preserved | baseline SHA/progress/B/checkpoints |
| W02 | Progress stages/responsibility/test-ID/logs | NDM_Audit02_Progress.md |
| W03 | Vertical stages and accepted intermediate suites | checkpoint table and historical logs |
| W04 | Local checkpoints, reconstruction instructions, no destructive Git | Git log/progress; no push |
| W05 | Continued implementation through all applicable DoD | final matrix/report/artifacts, not a partial green suite |
| W06 | This report, metrics/logs/status changes/limits/build/export | all sections of this document |

### Первый Аудит

| ID | Что сохранено/дополнено | Доказательство |
| --- | --- | --- |
| Old-T01 | Secant returns checked root, not arbitrary low endpoint | TestLimitSearchSecantFinalizesCheckedRoot + generic matrix, F |
| Old-T02 | Iteration exhaustion is not success | TestLimitSearchBisectionIterationLimitFails, F |
| Old-T03 | Unconverged initial probe is NumFail, not BaseFail | TestInitialLambdaFailureStatusMapping + failedProbe/offsetPair, F |
| Old-T04 | SearchBound distinct from physical no-criterion | boundary.8/genericMultiplier/searchBound, F |
| Old-T05 | No fictitious Post/formation point | TestCrackFormationNoCrackDoesNotBuildPostState, F |
| Old-T06 | All current consumers work without LastRunner | formation cache + actual Current width cache, F |
| Old-T07 | Failed not reused; latest named state and retry replace properly | TestStateRepositoryReusesOnlyConvergedStates + cache.*, F |
| Old-T08 | Required failed crack stage visible in aggregate/detail | TestCrackAggregateIncludesCurrentStateFailure and output tests, F/O |
| Old-A01 | Generic Search independent even Ultimate; no DomainContext | generic Ultimate/LoadMultiplier and SC |
| Old-A02 | Ordinary probes common Runner, linear API preserved | linear/steelSign tests and ProbeOnce read audit, F |
| Old-A03 | Formation no longer Width wrapper | TestAudit02IndependentFormation, F |
| Old-A04 | Search self-contained and not LC safety check | searchEngineering.*, independent snapshot, F |
| Old-A05 | Writers do not engineer statuses; simple arithmetic allowed | writer/typed reserve/result comments tests, F/O |
| Old-A06 | Typed reasons and meaningful Russian comments preserved | status dictionary/tree display/offsetPair/Auto warnings, F/O |

## Финальный Self-Audit

- [x] R01-R09, A01-A08, E01-E10, S01-S07, C01-C06 закрыты и имеют доказательства.
- [x] Общие численные loops реально объединены; adapters не направляют к двум
  оставленным старым search-реализациям. Только разные численные стратегии и
  обоснованные специальные физические residual/guess/finalization сохранены.
- [x] Formula не получает State/status; Width и Long не создают NumFail сами.
- [x] Reuse key отражает физический/effective контекст, не warm-start/retry.
- [x] Probe-cache локальный, failed state не мешает retry, finished snapshot
  сохраняет meta/числа/diagnostics независимо от дальнейшего рабочего solve.
- [x] Статус из machine data, не из комментария; BaseFail только при подтвержденном
  неприемлемом Capacity Offset. SearchBound и physical no-crack различаются.
- [x] Физические limits/plateau/узлы/знаки/units/reference/Psi/Stability не изменены
  ради сходимости. Source expected/tolerance сохранены и численные пары прошли.
- [x] Workbook build/compile/all suites и доступные output/snapshot checks прошли.
- [x] Измеренные замедления/ограничения среды не скрыты и не засчитаны как green.
- [x] Актуальные architecture/progress, отчет, код/export/book и логи сохранены.

Открытых обязательных дефектов/блокеров в пределах доступной применимости Audit02
не осталось. Отдельные будущие архитектурные задачи и ручная нормативная/AutoCAD
проверка не объявляются выполненными этим отчетом.
