# Audit03: Typed Ошибки Search И Named-State

Дата: 2026-10-04. Baseline этого среза: `dc8fd13c`.
Объем: F03/F05/F07, существующие общие Search и State pipeline.
Это направленная приемка конкретных контрактов, не закрытие всего Audit03.

## Границы Изменений

- LoadMultiplier сохраняет машинную причину отказа на этапах Prepare,
  BuildLoadPath, начальной пробы, построения скобки, уточнения и финализации.
  Ошибка программного callback-а не является численной несходимостью.
- Overflow VBA (`Err.Number = 6`) остается численным отказом. Уже полученная
  terminal-причина не заменяется вторичным исключением и не разрешает
  восстановительную финализацию Capacity.
- Formation проверяет FailureCode после EvaluateStrainPlane. Пустая модель,
  отсутствующая зависимость и действительное арифметическое переполнение
  сохраняют разные typed исходы; готовая terminal-причина запрещает новую оценку.
- Отсутствующий solver-result или не опубликованный обязательный named-state
  дает `rsInternalError + rcSolverDidNotReturnState`, а не неподтвержденный
  `rsNumericalFailure`. Зависимые Width/Longitudinal остаются blocked и `N/A`.
  Успешный StopReason другого solver-а не подменяет причину отсутствия State.
- Критерии, допуски, load paths, материальные пределы и исторические численные
  expected/tolerance не менялись. Новых классов и параллельного Search нет.

## Отрицательные Доказательства

| Проверка | Неизменный Production / Новые Тесты | Итог |
| --- | --- | --- |
| LoadMultiplier, 16 отказов x 2 ResultKind x 3 метода, затем recovery | v180b | `858/174`, watchdog exit=1; exceptions/ошибочная классификация/terminal fallback воспроизведены. |
| Formation residual, 8 случаев x 5 путей, затем реальный recovery | v182 | `380/70`, watchdog exit=1; отсутствующие зависимости превращались в NumFail, после terminal причины выполнялась новая оценка. |
| Отсутствующий runner.ResultSolver | v184b, точная test-only подмена в изолированной книге | `66/5`, watchdog exit=1; ложные NumFail/lifecycle/code и успешный комментарий при отсутствии State. |
| Не опубликованный named-state | v184b, точная test-only подмена до StoreForRequest | `66/5`, watchdog exit=1; те же ошибки границы producer/consumer. |

Исходные журналы v180 и v184 также сохранены: они содержат ошибки новых oracle,
а не только production-дефекты. В v180 неверно ожидался Calculated=True у
InvalidInput/InvalidConfiguration, хотя общий lifecycle сохраняет False.
В v184 внешний статус aggregate ошибочно сравнивался со статусом Width leaf.
Исправлялись только новые oracle; отрицательные артефакты не переписывались.

## Положительные Directed Gates

| Срез / Проверка | Итог | Ограничение |
| --- | --- | --- |
| v181 LoadMultiplier typed matrix | `1326/0` | Все 96 fault-сценариев и последовательные recovery; не вся матрица физических нагрузок. |
| v181 полный Capacity suite | `3074/0` | Не заменяет все восемь suites и остальные требования Config/UI. |
| v183 Formation residual | `450/0` | 40 fixed-plane fault-сценариев и реальный поиск точки после восстановления. |
| v185 LoadMultiplier typed matrix | `1326/0` | Повтор на общей чистой сборке всех production-правок среза. |
| v185 Formation residual | `450/0` | Повтор на той же общей сборке. |
| v185 отсутствующий solver-result | `71/0` | Точная fault-подмена только в отдельной книге; штатный full suite ее не содержит. |
| v185 отсутствующий named-state | `71/0` | Проверены оба потребителя, точные meta/code, blocked leaves и comments четырех writer-блоков. |

Все перечисленные положительные watchdog-прогоны завершились exit=0 и
source unchanged=True. Ошибочные fixture не являются поставляемой книгой.
Для успешных/неуспешных обычных нагрузок сохраняются самостоятельные проверки
трех компонент равновесия и физических пределов.

## Классификация И Lifecycle

| Фактическая Причина | InternalStatus / ResultCode | Display |
| --- | --- | --- |
| Программный callback, неверная внутренняя зависимость | `rsInternalError / rcInternalError` | `CalcErr` |
| Не возвращен обязательный State | `rsInternalError / rcSolverDidNotReturnState` | `CalcErr` |
| Невалидный вход/конфигурация | Соответствующие `rsInvalidInput`/`rsInvalidConfiguration` и их коды | `InputErr` |
| Действительное переполнение численного поиска | `rsNumericalFailure / rcNumericalFailure` | `NumFail` |
| Нет обязательного допустимого State для формулы | `rsBlockedByDependency` с фактической причиной зависимости | `N/A` у формульного leaf |

Внешний статус получает только CResultStatusPolicy. Текст ResultComment не
используется для назначения кода. InvalidInput/InvalidConfiguration имеют
Calculated=False даже при SearchExecuted=True, если ошибка обнаружена после
начала попытки. До Prepare/BuildLoadPath fault SearchExecuted=False; fault
после начала реальных probes сохраняет факт запуска.

## Выполненные Full И Workbook Gates

- Чистый v185, полный Off: восемь suites, `60259/0` assertions, watchdog exit=0,
  source unchanged=True. Это сумма итогов suites, не число независимых случаев.
  UI `38356/0`, elapsedSec=1678.6279296875.
- Numeric actual против полного Off v179: 13641 общая запись, missing=0,
  differences=0. Новые assertions не заменяли старые expected/tolerance.
- Обновленный v185, полный On: восемь suites, `60270/0` assertions,
  watchdog exit=0, source unchanged=True. UI `38356/0`,
  elapsedSec=1636.4638671875. Numeric actual против On v179:
  13646 общих записей, missing=0, differences=0.
- Обновленная книга создана из принятой v179 только штатным VBA import,
  с восстановлением неизменной GEOM_PI. OpenXML подтверждает сохранность
  всех 2681 непустых Config cells, формул, объединений и PrintArea.
- Validate clean/update: `27/27` каждая; formatting `1003/0` каждая,
  1002 адреса, deviations=0. ApplyAlignments не использовался.
- Справка clean/update: 2579 ячеек, одинаковые объединения, failed=0,
  обе книги неизменны. Это equality, не закрытие содержательных D02-кандидатов.
- После последнего сохранения обновленной книги выполнен read-only VBE export:
  source contracts `105/105`, failed=0, production/test classes `83 + 3`.
  В проекте Excel 110 components, включая пять document components.
- Directed status palette на updated v185: `355/0`, exit=0,
  source unchanged=True. Results save/reopen SHA
  `9F6AB303F8493E8761F2694159CB3CC2EB9B0248E466B3EC9221A14018704414`
  и status/style SHA
  `1EABA40A72D5A28392F11FC5D38F0D2B1E9B264FDFAD9F5514ACB9EE1A1E1132`
  совпали до и после повторного открытия. Это направленный gate, не
  замена финальной выпускной проверки после остальных правок Audit03.

| Артефакт | SHA-256 |
| --- | --- |
| Clean v185 | `2E3B3E75ACA83EB8BB11770E22F2E83CD43E46CDB62827A87A40ABA1B1E15B58` |
| Updated v185 | `5899B5DEC96352A8B71D0AD60C0CD3047E47D9C6E25B4F4EC6EC1D2C0289E04A` |
| Фактический VBE export | `F3C74176F1AB0890C14AB8FF2E592E6B3F3740BE6969C10571C39070E0EB0176` |

Статический expected-preservation против исходного Audit03 baseline сохраняет
536 выражений и отмечает шесть прежних изменений: четыре проверки ширины
столбцов теперь требуют точного сохранения пользовательской ширины, а две
условные записи Psi fallback заменены одной согласно согласованному поведению.
Сырой gate exit=1 сохранен как таковой; он не объявляется зеленым. Обоснование
Psi есть в ResultComment_Review, ширин — в пользовательском требовании и v151
Progress. В текущем diff всех четырех тестовых файлов только additions,
deletions=0: эти шесть изменений не внесены срезом v180-v185.

Первый штатный запуск Refresh-VbaModules без escalation завершился
CO_E_SERVER_EXEC_FAILURE и не изменил clone v179. Повтор в разрешенном Excel
COM-контексте завершился успешно. CompareWorkbookHelp нужно запускать PowerShell
7: Windows PowerShell 5 неправильно прочитал BOM-less UTF-8 скрипт; успешный
повтор использует тот же неизменный script. Эти сбои инфраструктуры не являются
успешными проверками либо NumFail расчетного ядра.

## Следующий Обязательный Объем

- Оставшиеся арифметические кандидаты, Config/pairwise, D01/F07 и выпускные
  матрицы/benchmarks: текущими directed assertions они не закрываются.

Clean v185 построен штатным Build-Workbook. Full Off/On и перечисленные workbook
gates выполнены; source/test срез v180-v185 принят для checkpoint.
Основной `workbook/output/RC_Section_NDM.xlsm` не опубликован до полного DoD.
