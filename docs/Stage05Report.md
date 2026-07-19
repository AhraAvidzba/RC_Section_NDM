# Отчет этапа 5: несущая способность Mx и My

## Статус

Этап 5 доработан в границах поиска несущей способности для направлений `Mx` и `My`. Этап 6 (`Mxy`) не выполнялся.

Реализован `CCapacitySolver`, который использует `CSectionSolver`, ищет предельный множитель `lambda`, отличает физические предельные состояния от численных отказов и записывает результаты в печатные диапазоны через отдельный Excel-адаптер.

## Созданные и измененные файлы

Создано:

- `src/Excel/CCapacityResultWriter.cls`
- `tools/build_workbook/Refresh-Workbook.ps1`

Изменено:

- `src/Solver/CCapacitySolver.cls`
- `src/Solver/CSectionSolver.cls`
- `tests/modTestCapacitySolver.bas`
- `tools/build_workbook/Build-Workbook.ps1`
- `docs/Architecture.md`
- `docs/ImplementationPlan.md`
- `docs/WorkbookLayout.md`
- `docs/ValidationPlan.md`
- `docs/Stage05Report.md`

## Реализованная схема

Для направления `Mx`:

```text
N  = Nspecified
Mx = lambda * MxBase
My = 0
```

Для направления `My`:

```text
N  = Nspecified
Mx = 0
My = lambda * MyBase
```

Определяются:

- `lambdaUltimate`;
- удерживающий момент по выбранному направлению;
- `SafetyFactor = lambdaUltimate`;
- `Utilization = 1 / lambdaUltimate`;
- последнее сошедшееся допустимое состояние `CSectionSolver`;
- физический предельный статус;
- диагностический журнал проб и повторов.

## Численная устойчивость

Несходимость `CSectionSolver` не является физическим разрушением.

Для каждой пробной точки `lambda` выполняется:

- расчет с использованием предыдущего сошедшегося состояния как начального приближения;
- при численной несходимости увеличивается число внутренних ступеней нагрузки;
- дополнительно уменьшается допустимое приращение кривизн;
- выполняется заданное число повторов;
- после исчерпания повторов возвращается `NumericalFailure` или `SingularTangent`.

`NumericalFailure` и `SingularTangent` не используются как верхняя физическая граница несущей способности. При численном отказе до нахождения физической верхней границы уменьшается шаг по `lambda`. При численном отказе внутри уже найденной физической скобки расчет останавливается с численным статусом и сохраняет последнее сошедшееся состояние.

Различаются статусы:

- `ConcreteStrainLimit`;
- `SteelStrainLimit`;
- `NumericalFailure`;
- `SingularTangent`;
- `InvalidInput`.

## Настройки System

На лист `System` добавлены настройки:

- `Capacity.InitialLambda`;
- `Capacity.MaxLambda`;
- `Capacity.ToleranceLambda`;
- `Capacity.MaxRetries`;
- `Capacity.BaseLoadSteps`;
- `Capacity.SolverMaxIterations`;
- `Capacity.ConcreteCompressionLimit`;
- `Capacity.SteelStrainLimit`.

Пределы деформаций на этом этапе имеют статус `PROVISIONAL_FOR_SOLVER_TESTING` и не считаются окончательно верифицированными нормативными параметрами.

## Вывод результатов

Добавлен `CCapacityResultWriter`.

Он записывает результаты:

- `Mx` в `rngResultMx`;
- `My` в `rngResultMy`.

Расчетное ядро не обращается к листам Excel. Доступ к `Workbook`, именованным диапазонам и ячейкам есть только в Excel-адаптере вывода.

## Независимая контрольная проверка

После завершения поиска сохраняется последнее допустимое состояние `CSectionSolver`. Для него вычисляется контрольная невязка удерживающего момента:

```text
MomentEquilibriumResidual = Abs(Mint - MomentUltimate)
```

Для `Mx` используется `Mxint`, для `My` используется `Myint`. Эта величина выводится в результаты и проверяется тестами равновесия.

## Тесты

Расширен `modTestCapacitySolver`.

Проверяются:

- положительный и отрицательный `Mx`;
- положительный и отрицательный `My`;
- `lambdaUltimate < 1`;
- `lambdaUltimate` около `1`;
- случай `N = 0`;
- предел по бетону `ConcreteStrainLimit`;
- предел по арматуре `SteelStrainLimit`;
- численная несходимость без превращения ее в физическую границу;
- несимметричное сечение с `Ixy <> 0` и связанной кривизной `kappaY` при расчете `Mx`;
- запись результатов в `rngResultMx` и `rngResultMy`;
- `InvalidInput` при нулевом базовом моменте.

## Проверки

Команды для проверки:

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Build-Workbook.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Refresh-Workbook.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Validate-Workbook.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-GeometryTests.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-LinearTests.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-MaterialTests.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-SectionSolverTests.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-CapacityTests.ps1
```

Полная пересборка через `Build-Workbook.ps1` в этой сессии была остановлена ошибкой Excel COM на `Workbooks.Add()`:

```text
Microsoft Excel cannot open or save any more documents because there is not enough available memory or disk space.
```

Так как процесс `EXCEL.EXE` не завершался даже через `taskkill /F /IM EXCEL.EXE`, для продолжения работы использован `Refresh-Workbook.ps1`. Он открывает существующую книгу, переимпортирует VBA-модули, обновляет настройки `Capacity.*` на листе `System` и сохраняет книгу без `Workbooks.Add()`.

Фактические результаты проверок после refresh:

```text
Validate-Workbook.ps1: все проверки Passed=True
TOTAL: passed=29; failed=0
TOTAL_LINEAR: passed=49; failed=0
TOTAL_MATERIAL: passed=20; failed=0
TOTAL_SECTION_SOLVER: passed=34; failed=0
TOTAL_CAPACITY: passed=50; failed=0
```

## Известные ограничения

- Используются временные параметры материалов `PROVISIONAL_FOR_SOLVER_TESTING`.
- Окончательная нормативная верификация пределов деформаций и диаграмм материалов остается открытой до закрытия нормативных вопросов.
- Сравнение с ручными расчетами и справочными программами еще не выполнено как инженерная верификация.
- Этап 6 (`Mxy`) не реализован в этой доработке.

## Следующий этап

После успешного прохождения всех регрессионных тестов этап 5 можно считать завершенным. Следующий этап по плану - `Mxy`, но переход к нему в рамках этой задачи не выполнялся.
