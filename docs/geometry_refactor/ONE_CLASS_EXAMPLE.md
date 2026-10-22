# Проверенный рецепт нового типа: скруглённая трапеция

Полный импортируемый класс: `examples/CSectionRecipeRoundedTrapezoid.cls`.
Он реализует ISectionRecipe и содержит только чтение своего компактного ввода
и декларации. Отдельные geometry/rebar/annotation builders не добавлены.
Общие сервисы и solver не изменены ради этого типа.

В CSectionTypeRegistry.Class_Initialize добавляется одна строка:

```vb
RegisterRecipe "RoundedTrapezoid", New CSectionRecipeRoundedTrapezoid, Array("rngRoundedTrapezoidGeometry")
```

Регистрация включена только в отдельной fixture-книге. Производственный
реестр сохраняет четыре текущих семейства. Из реестра строится выбор типов
при явном вызове modSectionRecipeSetup.UpdateSectionRecipeTypes.

## Собственные исходные данные

Один корневой блок `rngRoundedTrapezoidGeometry` содержит колонки
«Параметр / Значение / Ед. / Комментарий». Название параметра находится
в первой колонке, значение — во второй. Порядок строк допускается менять:
рецепт ищет ключ, а не фиксированный адрес. Длины/диаметры нормализуются
через CSectionInputTable.Length и CUnitSystem один раз. Counts — целые.
Источник единиц — rngUnitSettings; колонки единиц ссылаются на него формулой.

| Ключ | Проверенный ввод, мм/шт | Смысл |
|---|---:|---|
| BottomWidth | 1200 | Ширина исходной нижней грани до скругления |
| TopWidth | 800 | Ширина исходной верхней грани до скругления |
| Height | 800 | Расстояние между горизонтальными исходными гранями |
| CornerRadius | 60 | Радиус всех четырёх скруглений |
| OpeningRadius | 110 | Радиус круглого отверстия в (0,0) |
| OuterAxisDistance | 50 | Положительное расстояние от Outer до осей внутрь бетона |
| InnerAxisDistance | 45 | Положительное расстояние от Opening до осей из пустоты |
| OuterCount | 24 | Стержни по всему наружному замкнутому пути |
| InnerCount | 16 | Стержни по всему внутреннему замкнутому пути |
| OuterDiameter | 20 | Диаметр наружного ряда |
| InnerDiameter | 16 | Диаметр внутреннего ряда |

Защитный слой до поверхности отличается от AxisDistance на половину диаметра.
Рецепт включает CheckFullDisk и RejectOverlap явно; это геометрические проверки,
не нормативный минимальный зазор. Обе траектории аналитические. Вспомогательные
пути осей не становятся бетонными контурами. B на схеме измеряет фактический
наружный габарит после скругления, поэтому он отличается от BottomWidth.

Программный вариант вызывается без Config:

```vb
Dim recipe As CSectionRecipeRoundedTrapezoid, definition As CSectionDefinition
Set recipe = New CSectionRecipeRoundedTrapezoid
Set definition = recipe.ProgramDefinition(1200#, 800#, 800#, 60#, 110#)
```

Все аргументы программного варианта — внутренние мм. Геометрия и бары совпали
с чтением проверенного блока Config. После смены параметра создаётся новое
definition; старый опубликованный снимок остаётся прежним.

## Воспроизведение в отдельной книге

Следующие действия предназначены для fixture, не для исходных данных пользователя.
Native scripts требуют разрешённый COM-доступ к Excel/VBProject. Команды выполняются
из корня репозитория в Windows PowerShell. Сначала копируется проверенная книга
DeclarativeApi в docs/regression/.../OneClassExample/RC_Section_NDM.xlsm.

```powershell
./tools/build_workbook/Update-PrivateWorkbookCode.ps1 `
  -WorkbookPath docs/regression/GeometryDeclarative_2026-10-12/OneClassExample/RC_Section_NDM.xlsm `
  -ModuleNames @('CSectionTypeRegistry','CSectionRecipeRoundedTrapezoid','modTestRoundedTrapezoidExample','modTestDeclarativeExamples') `
  -ExampleFiles @('docs/geometry_refactor/examples/CSectionRecipeRoundedTrapezoid.cls','docs/geometry_refactor/examples/modTestRoundedTrapezoidExample.bas') `
  -RecipeRegistration @('RoundedTrapezoid','CSectionRecipeRoundedTrapezoid','rngRoundedTrapezoidGeometry')
```

RunExampleBuildTests устанавливает только компактный fixture-ввод и проверяет
штатную сборку. RunExampleInputTests проверяет тот же ввод в mm/cm/m,
формулы при Calculation=Manual и ошибочный адресный ввод. RunExampleWorkflow
принимает каталог для PNG и BottomWidth, строит preview, затем выполняет
RunSectionCalculationForWorkbook с реальным DirectState и схемой из Results.
Capacity/ширина трещин/устойчивость в этом smoke выключены; этот контроль не
заменяет обязательную итоговую адресную приёмку остальных расчётных маршрутов.

```powershell
./tools/build_workbook/Run-Audit03Watchdog.ps1 `
  -SourceWorkbook docs/regression/GeometryDeclarative_2026-10-12/OneClassExample/RC_Section_NDM.xlsm `
  -ReportPath docs/regression/GeometryDeclarative_2026-10-12/OneClassExample/InputAndBuild.txt `
  -Macro modTestRoundedTrapezoidExample.RunExampleBuildTests,modTestRoundedTrapezoidExample.RunExampleInputTests,modTestDeclarativeExamples.RunRoundedFrameExampleTests `
  -TimeoutSeconds 120
```

Для workflow используется один Macro=modTestRoundedTrapezoidExample.TestRunExampleWorkflow,
MacroArgument1 — абсолютный путь каталога PNG, MacroArgument2=1200 или 1400,
VerifyResultsReopen включён. Макросы в examples/modTestRoundedTrapezoidExample.bas
— тестовая оркестрация с пометкой ДЛЯ ТЕСТОВ, не алгоритмические компоненты типа.

## Фактические доказательства

94 build + 250 input + 12 E09 + 7 workflow1200 + 7 workflow1400 = 370 assertions PASS.
Native diff: 145 → 148 компонентов, одна новая class, две test standard modules.
Единственный изменённый production body — точная строка регистрации в реестре.
Остальные изменения — только регистр идентификаторов редактора VBE; их сравнение
не игнорирует строки, комментарии, числа или арифметику. Unexpected/sourceMismatch=0.

Smoke 1200: 2.21484375 с; smoke 1400: 1.60546875 с от штатного вызова до записи
Results, отчёта и схемы. Это направленный DirectState, не общий performance gate
всех предельных/трещинных/устойчивых маршрутов. Обе схемы осмотрены, 40 стержней,
наружная/внутренняя подписи отражают 24/16 физических стержней.
Save/reopen сохранил хеши Results в обоих случаях.

E09 находится в tests/modTestDeclarativeExamples.bas: только общий фасад,
без нового типа/Config-режима. Контрольные A=775879.6447372311 мм²,
xc=-11.511227898383616 мм, yc=+4.604491159353446 мм; 40 физических стержней,
реально сошедшийся N+Mx+My, preview и save/reopen. Проверка сетки не требует
равенства аналитической площади её приблизительной механической площади.
Суммы контрольных литералов сохраняют заданные Double при ограничении VBE
15 значащими цифрами; допуски не расширены и baseline не перезаписан.

Пакет фактических результатов: ONE_CLASS_EVIDENCE.json и
GeometryDeclarative_2026-10-12/OneClassExample/. Сохранены failed probes:
reserved Input, неверные ожидания Chart.Series/summary row и округлённые
VBE литералы. Это не результат окончательной приёмки всей задачи.

Основная книга/VBA_All_Code.txt ещё прежние. Полный API-каталог и гайд
в самом конце реального листа Справка остаются следующим этапом.
