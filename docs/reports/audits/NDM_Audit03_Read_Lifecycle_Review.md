# Audit03: Повторное Чтение И Ошибочный Snapshot

Дата: 2026-10-04. Предыдущий checkpoint: `37493b25584e0dd53e28c97c4c0b77f6e74656c1`.
Scope: F03/F04/F07, сохранение точной причины и отсутствие частично готового
объекта после ошибки чтения. Это не новый архитектурный этап.

## Подтвержденные Ошибки

1. `CMaterialModelSpec.Initialize` после корректного первого вызова сохранял
   `IsComplete=True`, если повторный разбор одного из четырех selectors
   завершался ошибкой. `Clone/SpecKey` и material provider могли использовать
   прежнюю либо частично измененную спецификацию. Обычный Catalog создает
   новый Spec, поэтому отдельный дефект штатного Config из этого сценария
   не заявляется: воспроизведена ошибка публичного lifecycle API.
2. `CCalculationProfileCatalog.LoadFromRange` публиковал Count до окончания
   чтения. Ошибка PR4 оставляла PR1-PR3 доступными, а PR4 мог иметь ID без
   готового profile. Проверено на том же экземпляре, включая перенесенную
   таблицу и независимость ранее выданного корректного профиля.
3. `CSectionPlotDataReader` читал численные свойства через raw CDbl либо
   SafeDouble. Formula error/Boolean/пустота могли стать числом/нулем,
   а текст/overflow давали техническую ошибку без поля, адреса и действия.
   Аналогичная ошибка воспроизведена в выбранных Stress/Strain строках.
4. Поздняя ошибка геометрии после добавления первого элемента оставляла
   частичный reader доступным. Три публичных режима не очищали данные при
   неуспешном завершении чтения.

## Изменение Контрактов

- Spec сбрасывает complete до разбора, устанавливает True только в конце.
  Provider отвергает незавершенную спецификацию существующим кодом 3251;
  новый статус или класс не вводятся.
- Catalog публикует Count после всех столбцов; handler удаляет частичные
  массивы и повторно передает исходные Err.Number/Source/Description.
- Plot reader использует одну проверку чисел геометрии, свойств и элементов:
  отдельно CVErr/Null/Boolean/Date, пустота, числовой текст и CDbl overflow.
  Существующий код 4719 содержит поле, текущий лист/адрес и предложение
  повторить импорт/расчет. CUnitSystem остается владельцем размерностей.
- Отсутствие optional geometry размеров остается допустимым; поврежденное
  присутствующее численное поле не подменяется нулем. Числовой текст допустим.
- На ошибке начатого чтения три публичных reader-режима очищают элементы,
  annotations, выбранный LC/Profile и плоскость; причина не переписывается.
  Nothing-guards до начала нового чтения не изменялись.
- Неиспользуемый private SafeDouble удален после проверки всех call sites.
  Равновесие, формулы, физические пределы и исторические tolerance не менялись.

## Направленные Доказательства

`RunAudit03ReadLifecycleTests` включен в основную workbook-interface suite.
Fixture: отдельная временная книга с одним бетонным элементом/одним стержнем
и синтетическим сохраненным State; она закрывается без сохранения.
Это проверка snapshot, не настоящий equilibrium case.

| Срез | Результат | Интерпретация |
| --- | --- | --- |
| v205 | остановка до assertions | Ошибка имени нового assertion-helper; исправлена только в тесте. Не production failure. Закрыт только собственный тестовый Excel. |
| v205b | 1021/2556 | Negative прежнего production: Spec, поздний Catalog, properties/частичный reader. |
| v205c | 1091/2838 | Расширенный negative: 468 variants, выбранные element Stress/Strain и три reader-режима. |
| v206 | 3929/0 | Тот же directed matrix на исправленном production, recovery, перенесенные адреса, числовой текст, незатронутые OTHER/другой State rows. |

Варианты: четыре Spec selectors; 24 Catalog variants в двух положениях;
properties со строкой, CVErr, Boolean, пустотой и overflow; выбранные
Stress/Strain; поздний fault геометрии; восстановление на том же объекте.
Число вариантов 468 не равно числу инженерных нагрузок: equilibriumCases=0.
Геометрический oracle дополнен независимой проверкой координаты Y арматуры.

Логи и книги: `docs/regression/Audit03/read_lifecycle_*v205*.log`,
`read_lifecycle_positive_v206.log`, соответствующие отдельные `.xlsm`.
Source unchanged=True у направленных прогонов. Пользовательский Excel PID
23476 и основная output-книга не использовались.

## Окончательные Gates И Ограничения

Все gates этого среза выполнены на замороженном v206:

| Gate | Фактический Результат |
| --- | --- |
| Full Off | 82221/0, все восемь suites; UI 60318/0, 2006.35546875 s; source unchanged=True. |
| Full On | 82232/0, все восемь suites; UI 60318/0, 1973.8359375 s; source unchanged=True. |
| Exact numeric Off, v204 -> v206 | 18422 общих actual-values, missing=0, differences=0. |
| Exact numeric On, v204 -> v206 | 18427 общих actual-values, missing=0, differences=0. |
| Structure | 27/27, read-only; книга не исправлялась. |
| Config formatting | 1003/0; 1002 адреса, deviations=0, без ApplyAlignments. |
| Status palette/save-reopen | 355/0; Results и фактические стили статусов совпали до/после сохранения. |
| Help v204 -> v206 | 2628 непустых ячеек в каждом snapshot, текст и объединения совпали, failed=0. Это сравнение двух принятых срезов, не новая clean-сборка. |
| Actual VBE/source | 105/105, failed=0; 110 компонентов книги. |

Проверенная книга:
`RC_Section_NDM_read_lifecycle_positive_v206.xlsm`, SHA-256
`6AFEEB93764A158D93F0E32A77F8CAACEDE6C1139428BD100101C6278A007464`.
Фактический read-only экспорт после окончательного RestoreGeometryConstants:
`VBA_All_Code_read_lifecycle_v206_2026-10-04.txt`, SHA-256
`9B01E7073D4C95102B6613DDC72FC507884E4B15B10EA8B6D112118F45566B6F`.
Он скопирован в canonical `VBA_All_Code.txt` с тем же хешем. Сверка допускает
только документированное VBE-представление, не новый численный tolerance.

Results save/reopen SHA-256:
`4229D60E3CDE8B06B0568364C1C74FB748E53F1378ACCE794DD62BF6EF73AA98`;
status-style SHA-256:
`1EABA40A72D5A28392F11FC5D38F0D2B1E9B264FDFAD9F5514ACB9EE1A1E1132`.
Источник книги, пользовательский output и исходное ТЗ не изменились.
Первый запуск help-comparison через Windows PowerShell не прочитал UTF-8
скрипт без BOM; повтор через bundled PowerShell 7 завершен failed=0.
Это ошибка запуска инструмента, не дефект текста справки или книги.

Source census: 105 modules, 4390 methods, 1639 guard candidates,
83 production + 3 test classes; semantic acceptance Pending. Число
синтаксических кандидатов не является числом ошибок либо инженерных задач.

Этот срез не закрывает весь F07/D01/K/T. Еще открыты general Plot/annotation
selectors, текстовые/логические metadata snapshot, Unit-error поля,
непереданные зависимости до начала загрузки, Excel guard partial COM failure,
крайняя арифметика и performance reader-а. Реестр Config не получает
дополнительные active-reviewed записи от чисто snapshot/lifecycle теста.
Направленный fixture использует контролируемый PR1 Stress-профиль; при
расширении общей fixture-isolation следует задавать Quantity явно, а не
переносить это условие на произвольную пользовательскую книгу.
