# Audit03: Арматурные Поля Config

## Граница Среза v186-v192

Проверяется настоящий путь Config -> CSystemSettingsReader -> CUnitSystem ->
registry/builder -> CRebarLayout для Circle, RectSet, RoundedRectangle и
HollowRectangle. Новых классов и расчетной методики НДС нет.
Этот документ не является финальным отчетом Audit03.

Срез v185 принят checkpoint `0340524f`. Отрицательные тесты выполнялись на
его production-коде или явно указанной последующей версии; исходные журналы
не переписываются после исправления.

| Прогон | Assertions passed/failed | Сценарии | Результат |
| --- | --- | --- | --- |
| non-circle negative v186 | 7546/1376 | 2552 | exit=1, source unchanged=True |
| non-circle positive v187 | 9570/0 | 2552 | exit=0, source unchanged=True |
| Circle negative v188 | 366/106 | 160 | exit=1, source unchanged=True |
| Circle positive v189 | 524/0 | 160 | exit=0, source unchanged=True |
| inactive-offset negative v190 | 10098/24 | 2744 | exit=1, source unchanged=True |
| expanded non-circle positive v192 | 10146/0 | 2744 | exit=0, source unchanged=True |

Исходные логи находятся в `docs/regression/Audit03/`:
`rebar_fields_negative_v186_2026-10-04.txt`,
`rebar_fields_v187_2026-10-04.txt`,
`circle_rebar_negative_v188_2026-10-04.txt`,
`circle_rebar_v189_2026-10-04.txt`,
`inactive_offsets_negative_v190_2026-10-04.txt`,
`rebar_fields_v192_2026-10-04.txt`.

## Исправленные Контракты

- Активные as, отступы t и селекторы дополнительных рядов обязательны.
  Пустое/неизвестное/ошибочное значение не заменяется параметром шаблона.
- Существующая пустая или нулевая ячейка диаметра/количества выключает ряд.
  Отсутствующая обязательная строка отличается от существующего пустого ввода.
  Отрицательное число, дробный счетчик и ошибка формулы не становятся нулем.
- Reader хранит CVErr геометрии отдельно до обращения активного потребителя.
  Чтение такого поля выдает фактическое место ошибки до любой подстановки;
  другая форма или выключенный ряд не потребляют его.
- В Rectangle нижние размеры/армирование H2/B2 не читаются. Общий RectSet
  loc/bind требуется, если соответствующий ряд есть хотя бы на одной стороне.
- as выключенной наружной грани Rounded/Hollow не читается. as Opening
  сохраняет геометрического потребителя: ограничение соседних проекций;
  он не объявляется неактивным только из-за отсутствия собственного ряда.
- Ошибки активных значений содержат динамический адрес прочитанной ячейки
  и действие: ввести число, выбрать допустимый вариант или исправить формулу.
  Действующий relocation-тест обращается к отложенному geometry CVErr через
  активное GetRawString, а не требует ошибки уже при загрузке другой формы.

Четыре физических failures v186 относятся к дополнительным рядам нижней
прямой грани Opening. В продленном углу обе пробные точки могли находиться
в бетоне, поэтому ContainsPoint не определял нужную нормаль. Для этой прямой
используется уже известная нормаль грани от пустоты в бетон. Остальная
shape-specific геометрия не заменена новым алгоритмом.

## Независимые Проверки

Для прямоугольных fixtures нормаль/касательная заданы тестовыми координатами,
а не возвращены production-helper смещения. Проверяются подписанное движение
as/t, количество и диаметр стержней, расстояние каждого дополнительного ряда
от ближайшего первого ряда и направление относительно пустоты Opening.

Stacked/SideBySide и EachBar/EverySecondBar проверяются в активном сценарии.
Неактивность сравнивается полным layout-snapshot после blank/abc/CVErr;
после каждой попытки исходная Formula восстанавливается. Итоговые таблицы
восстанавливаются целиком, в том числе после ошибки.

Исторические корректные numerical expected/tolerance не ослаблялись.
Изменен только момент обнаружения geometry CVErr в relocation-тесте согласно
новому активному consumer-контракту; адрес, отказ и recovery остаются обязательны.

## Справка И Книги

Чистая v192 собрана штатным Build-Workbook. Обновленная v191 создана из
предыдущего isolated fixture штатным import/help-update с сохранением input.
После последнего сохранения в обе книги восстановлена неизменная GEOM_PI.

- Actual help update v192: failed=0, 1976 строк, 140 прямых ссылок, 118 shapes.
  Значения/формулы/validation/оформление всех input и PrintArea сохранены.
- OpenXML clean/update: 2604 непустые help cells и все merges совпадают,
  failed=0, обе книги неизменны при сравнении.
- OpenXML v185 -> updated v192: все 2681 непустые Config-ячейки и формулы
  неизменны, editable changes=0; merges/PrintArea равны, SHA обеих книг
  сохранены. Отчет `config_preservation_update_v185_v192_2026-10-04.json`.
- NotCracked описан как невызванная формула, N/A и пустой a_crc; вычисленный
  нулевой результат отделен от отсутствия расчета. CrackedState требует
  утвержденную SLS(II)/Ignore модель, а не «обычно» такой выбор.

Три предыдущих help-update v191/v191b/v191c дали по два failures нового
assertion: он искал дословную фразу резервной wildcard-ветви вместо реального
методического раздела. Сохраненные A289/A434/A465 подтверждают нужный смысл;
потеря текста не установлена. Assertion уточнен, негрин журналы сохранены.
Предварительный запуск comparator через Windows PowerShell 5 не прочитал
BOM-less UTF-8 скрипт; неизменный comparator успешно выполнен PowerShell 7.
Это сбой запуска проверки, не NumFail расчетного ядра.

| Книга | SHA-256 |
| --- | --- |
| clean v192 | `7BA6D4D06C868BA9EE54B0CDCAFC46D1AD19064F17A94A0BEA2B98521DF1C865` |
| updated v191 после gates v192 | `FC51F528B17A453995254801CDB788D06233AF78A5D6D6C6F8740CD3773C626E` |

## Итоговые Gates Среза

Полные Off v192 завершены: 70929/0, восемь suites, UI 49026/0,
watchdog exit=0/source unchanged=True. Numeric comparison v185 -> v192 Off:
13641 общих actual-values, missing=0/differences=0, без новых допусков.
Полные On v192 завершены: 70940/0, восемь suites, UI 49026/0,
watchdog exit=0/source unchanged=True. Numeric comparison On: 13646 общих
actual-values, missing=0/differences=0. Validate clean/update: 27/27 каждая,
source unchanged=True. Formatting clean/update: 1003/0 каждая, 1002 адреса,
deviations=0, без ApplyAlignments. Palette/save/reopen updated: 355/0,
Results и status/style hashes равны после reopen, source unchanged=True.
Фактический read-only VBE export получен из обеих книг, по 110 компонентов;
source/VBE equality clean/update: 105/105 каждая, failed=0. Новых классов
нет: 83 production + прежние 3 test. Канонический regression-export взят
из фактически сохраненной updated-книги; SHA
`4BD71131BABD5D3143E26EAA7E03438B8BAD6D4D0BD85B92F52C6299F0FC8E2F`.

После всех gates присоединены 176 адресных entries, из них 156 новых.
Registry `config_behavior_registry_rebar_v192_2026-10-04` содержит
646/1064 active-reviewed, fullAcceptance=False. Из 760 editable адресов
остаются 114: 73 system, 20 profile presentation/description, 21 annotation.
Приемка активного поведения не подменяет полную проверку диапазонов/пар.

Отдельно уточнены восемь ранее принятых Circle contracts: metadata больше не
утверждает runtime-default 300/40/20/Stacked для отсутствующего активного ввода.
Трассировка `NDM_Audit03_Circle_Rebar_Input_Evidence.json` связывает прежние
53 effect assertions с 524 input assertions/160 случаями и обоими full gates.
Эта коррекция не увеличивает accepted count; вместе обновлены 184 entries,
из них 156 новых и 28 ранее принятых. Исторические registry не переписаны.

Ручные source/test/tools/doc правки проходят scoped git whitespace-check.
Общий check новых raw logs/VBE exports отмечает сохраненные пробелы и пустые
строки фактического вывода; эти evidence-файлы намеренно не форматируются
задним числом и сохраняют проверенные SHA. Это не failure расчетных gates.

Полные вещественные границы размещения, curved fixtures и рискованные пары
остаются в расширенной приемке. В частности, generic error размещения слишком
большого дополнительного стержня Circle пока не доказывает адресную диагностику
для конкретного диаметра. Эта граница не скрывается приемкой обычного fixture.

Основная output-книга и входное пользовательское ТЗ сохраняют исходные SHA.
Финальный Audit03 self-audit, матрицы, benchmarks и выпуск не завершены.
