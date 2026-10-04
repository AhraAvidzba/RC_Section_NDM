# Audit03: Обязательные Флаги Отчета И Сообщений

## Scope И Подтвержденный Дефект

Срез после `c0335b35`. Проверяются General.ExecutionReportEnabled и
General.NonCriticalMessagesEnabled через настоящие именованные Config таблицы.
Меняются существующий reporter, workbook/export entrypoints, тесты и справка;
новых классов, инженерных формул и правил solver-а нет.

Frozen v237: 95/57 на прежнем production. Потерянная строка флага отчета
молча выбирала No. Для сообщений обработчик DefaultEnabled скрывал blank,
TODO, неизвестный текст, CVErr, потерянный ключ и отсутствие workbook,
возвращая True вместо исходной ошибки. Проверки повторены в A5 и CH800.

Первый v236 не дошел до проверки поведения: собственный fixture не содержал
обязательного rngPlotAnnotationSettings. Это setup failure/NotRun, не
production defect. Журнал 0/1 сохранен отдельно. V237 добавляет настоящую
таблицу аннотаций и служит отрицательным доказательством исправляемого кода.

## Изменения

- CExecutionReport.Initialize использует GetRequiredBoolean. Повторный
  Initialize по-прежнему сбрасывает Enabled/FilePath/буфер до чтения, а No
  оставляет методы записи безопасными no-op и не удаляет старый отчет.
- NonCriticalMessagesEnabled передает исходную адресную ошибку Config.
  Для Nothing workbook возвращается контролируемая ошибка, а не True.
- Шесть пользовательских кнопок проверяют решение показа информации до
  расчета, очистки, перерисовки, импорта или изменения DWG. Сообщения об
  ошибках сохраняют независимые внешние handlers. AutoCAD-кнопки используют
  уже загруженный снимок settings, без повторного чтения всей таблицы.
- Workbook core без запроса интерактивного сообщения не получает новую
  математическую зависимость от этого UI-флага. Неинициализированный
  необязательный report-объект также не включает запись сам по себе.
- Справка явно описывает обязательность, отсутствие fallback, порядок
  preflight и сохранение ранее записанного txt при No.

## Реальные Проверки

Собственный сохраненный RunControls.xlsx находится в отдельной папке
watchdog-fixture. Вход - реальные System/Unit/Sign/Annotation таблицы.
Yes проверяется по существующему UTF-16 txt и его marker/final. No не
меняет содержимое файла. Recovery и повторный Initialize не возвращают
старый буфер. Selector сообщений проверен на Yes/No, TRUE/FALSE, 1/0,
Да/Нет; это вызов реального API, не ручное создание meta.

Двадцать отрицательных вариантов: четыре неверных значения и удаленный
ключ для двух флагов, в двух положениях System. Проверяются исходный typed
error, ключ, действие и адрес существующей ячейки; адрес удаленной строки
не угадывается. Счетчик equilibrium solve остается прежним.

## Gates

- Directed v238: 152/0, 20 ошибочных вариантов, два положения, noSolve.
- Существующий полный workbook-сценарий с отчетом: 3/0.
- Соседний presentation: 3220/0; адресные input errors: 1973/0.
- Все положительные watchdog: exit=0, source unchanged=True.
- Read-only структура положительной книги: 27/27.
- Actual Help: 2007 строк, 140 ссылок, 118 shapes, failed=0 после
  save/reopen. Строго сохранены все 760 input signatures, включая значения,
  формулы, форматы, выравнивание и validation.
- Три новые контрактные строки дополнительно прочитаны из сохраненного
  OpenXML: Справка!B908, B909, B916. Это не пиксельная/нормативная приемка.
- Реальный VBE export: 111 components; source/export equality 106/106,
  failed=0. Gate хранится в отдельном журнале и не подменяется успешным macro run.

## Артефакты

Frozen negative `RC_Section_NDM_general_run_negative_v237.xlsm`:
SHA-256 `A14D202CB4C39643F19F10ACA7E1FA513E6CF312613EE9221E3E9A69A6447C24`.
Final positive `RC_Section_NDM_general_run_positive_v238.xlsm`:
SHA-256 `1653B401780B2F65C8918FF0D492BBB720DF289CE46837C389D655357E9D59D6`.
Actual export `VBA_All_Code_general_run_v238_2026-10-04.txt`:
SHA-256 `F263C2CEA69741F5615C671CF825870DFA9D9AA3CD33C0099C3D21DFE5EC4663`.
Raw negative/positive/help/report/plot/input/structure журналы сохраняются
отдельно; setup failure не переименован в PASS.

## Ограничения И Следующий Объем

Реестр принимает два дополнительных active-behavior поля: 707/1064,
из 760 editable полей 53 pending, fullAcceptance=False. Направленные
проверки не являются приемкой всех файловых отказов, реальных MsgBox,
внешнего DWG, mutation sensitivity или общего pairwise.

Численная модель не менялась. Полный восьми-suite On/Off не повторен по
разрешению пользователя; финальный release остается обязательным. Main
output, audit specification и пользовательский Excel PID 23476 сохранены.
Оставшиеся Config, metadata/geometry guards, semantic review, финальные
матрицы нагрузок, benchmarks и clean/update/release впереди.
