# Подготовка итоговой книги 12.10.2026

Предыдущий этап 0375b4d4 был прогрессом: реальный сохранённый гайд, таблицы
в колонках Excel, API mapping и native assertions. Этот этап продолжает
подготовку выпуска; основной output ещё не опубликован.

FinalCandidate/RC_Section_NDM.xlsm создана из проверенной NativeGuide-копии.
В неё импортирован полный текущий src/tests (142 source components),
нативная компиляция прошла. Всего 147 компонентов, включая 5 document modules.
VBA_All_Code.txt получен read-only из фактического VBProject.

Независимая Verify-GeometryReleaseData.py сверяет все непустые значения/
формулы вне новой Справки с исходной пользовательской книгой. Реальные LC,
Results и расчётный лист сохранены. 6277 непустых ячеек проверены; ровно
63 разрешённые изменения Config: canonical setbacks, перемещение прежних
единиц/пояснений, append ClosedUniform и список из реестра. Дополнительные
имена только rngSectionRecipeTypes и rngSectionRecipeGuide, прежние не удалены.

Обнаружено преобразование Excel имени _xlnm.Print_Area в Print_Area при
прежних COM-save. В кандидате восстановлен исходный built-in print name
через существующий Restore-WorkbookPrintAreas. Повторная XML-сверка PASS;
printAreaPreserved=True. Главная исходная книга по-прежнему имеет SHA256
769C60CEF6030A009825875B2488B6D1A7C5FC653025F7B6CBAB6B4FCE63EAED.
Кандидат после полного импорта/восстановления: SHA256
86172120584C3A09CBA6EF13A14A17A5ABFF595863E0FEFAFBB1342A1FC814AD.

Чистая сборка FreshBuild создана штатным Build-Workbook из актуальных source,
без legacy. Новый параметр FormulaImageDirectory изолирует картинки формул
в evidence; посторонние текущие workbook/output/formula_images не изменялись.
Сборка действительно вызвала последний раздел Справки. В общей сборке
PowerShell COM binder пытался приводить двумерный Value2 к String; отдельный
типизированный Set-RecipeGuideMatrix устранил проблему. Неуспешная reflection
проба сохранена; её не принимаем за способ записи. BulkWriteProbe.json
подтверждает прямую/одно-колоночную запись и отказ reflection-setter.

Сопоставление дампов кандидата и независимо собранной FreshBuild: 142 source
+ 5 generated document modules MATCH. Первичная строго текстовая сверка
с `.cls` обнаружила только печать Double редактором (exponent/decimal,
15 digits/снятие #) и ANSI-печать µ комментария. SourceEquality.literal-and-
comment-probe.json сохраняет их; core/normative source не менялся ради
формального PASS. Native fresh reference подтверждает одинаковую компиляцию.

FreshBuild/GuideReopen: updater дважды исполнен, одинаковый блок 1705–9291,
все состояния восстановлены при success/edited content/protected sheet,
текст/таблицы совпали после reopen, три перехода раскрыли тела. Изображения
оглавления и длинной сигнатуры просмотрены; таблицы выровнены, текст не обрезан.
Положение на две строки отличается от сохранённой пользовательской книги
1707–9293 из-за её фактического прежнего содержимого; жёсткой строки нет.

Направленные native проверки полного состава:

| Проверка | PASS | FAIL |
|---|---:|---:|
| E09, advanced, operations, diagnostics | 95 | 0 |
| RetiredGeometryCapacity (6 затронутых сценариев) | 74 | 0 |
| CrackGeometry | 41 | 0 |
| 4 нормативные ветви Stability через batch/writer | 100 | 0 |
| PostAudit03SnapshotRepeat | 6059 | 0 |
| BatchWorkStats | 14 | 0 |

Итого 6383 assertions PASS. SnapshotRepeat сохранён/открыт, Results hashes
618D042AF72EA1F9750DA8EE05822FE53B51A454C6155140F497F6DE3A256B7F совпали.
Неверное совместное указание двух macro с VerifyResultsReopen было отклонено
runner до тестов; журнал сохранён, затем проверки запущены отдельно.

NativeCADRegistration.json: BLOCKED. AutoCAD.Application.24.2 зарегистрирован
на SOFiPLUS `sofp.exe /Automation`, а безопасный native runner требует
Autodesk AutoCAD 2023 `acad.exe`. Приложение не запускалось, пользовательские
документы не использовались. D19 допускает явно BLOCKED native integration;
это не PASS round-trip и не отказ геометрии. Fake CAD проверки остаются
обязательными для окончательной матрицы.

Осталось: окончательная направленная матрица требований/семейств и состояний,
fake CAD/passive export, end-to-end performance, финальные отчёты; повторная
проверка данных/source/dump перед публикацией основного output и local commit.
Полные тяжёлые on/off/UI по решению пользователя не выполняются. Не отмечать
всю цель достигнутой на основании этого промежуточного этапа.
