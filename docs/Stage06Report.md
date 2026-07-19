# Отчет этапа 6: косой изгиб Mxy

## Статус

Этап 6 реализован в границах поиска несущей способности при `N + Mx + My` с сохранением направления вектора моментов.

Не реализовывались:

- расчет ширины раскрытия трещин;
- пакетный расчет сочетаний;
- новые геометрии.

## Реализованная схема

Для косого изгиба используется тот же одномерный поиск по множителю `lambda`, что и для этапа 5:

```text
N  = Nspecified
Mx = lambda * MxBase
My = lambda * MyBase
```

Направление вектора моментов сохраняется, то есть отношение `Mx / My` остается равным `MxBase / MyBase` для каждой пробной точки.

## Код

Изменено:

- `src/Solver/CCapacitySolver.cls`
- `src/Excel/CCapacityResultWriter.cls`
- `tests/modTestCapacitySolver.bas`
- `docs/Architecture.md`
- `docs/ImplementationPlan.md`
- `docs/WorkbookLayout.md`
- `docs/ValidationPlan.md`

Создано:

- `docs/Stage06Report.md`

## Расчетный слой

В `CCapacitySolver` добавлен метод:

```text
SolveMxy(mesh, rebars, concreteMaterial, steelMaterial, nSpecified, mxBase, myBase)
```

Он использует существующую внутреннюю процедуру поиска `SolveOneDirection` и не дублирует логику этапа 5.

Добавлены результаты:

- `MxUltimate`;
- `MyUltimate`;
- `MomentUltimate` как модуль результирующего момента для `Mxy`;
- `MomentEquilibriumResidual` как результирующая невязка по `Mx` и `My`.

Сохраняются статусы:

- `ConcreteStrainLimit`;
- `SteelStrainLimit`;
- `NumericalFailure`;
- `SingularTangent`;
- `InvalidInput`.

`NumericalFailure` и `SingularTangent` не считаются физической границей несущей способности.

## Вывод в книгу

В `CCapacityResultWriter` добавлен метод:

```text
WriteMxyResult(workbook, capacity)
```

Он записывает результат в `rngResultMxy`.

Расчетный класс `CCapacitySolver` по-прежнему не обращается к Excel-листам.

## Тесты

В `modTestCapacitySolver` добавлены проверки:

- `Mx > 0`, `My > 0`;
- `Mx < 0`, `My > 0`;
- `Mx > 0`, `My < 0`;
- `Mx < 0`, `My < 0`;
- сохранение направления вектора моментов;
- равновесие последнего допустимого состояния;
- несимметричное сечение с `Ixy <> 0`;
- запись результата в `rngResultMxy`.

## Проверки

После обновления книги через:

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Refresh-Workbook.ps1
```

capacity-тесты пройдены:

```text
TOTAL_CAPACITY: passed=92; failed=0
```

Полный регрессионный запуск выполнен:

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Validate-Workbook.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-GeometryTests.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-LinearTests.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-MaterialTests.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-SectionSolverTests.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-CapacityTests.ps1
```

Результаты:

```text
Validate-Workbook.ps1: все проверки Passed=True
TOTAL: passed=29; failed=0
TOTAL_LINEAR: passed=49; failed=0
TOTAL_MATERIAL: passed=20; failed=0
TOTAL_SECTION_SOLVER: passed=34; failed=0
TOTAL_CAPACITY: passed=92; failed=0
```

## Ограничения

- Используются временные параметры материалов `PROVISIONAL_FOR_SOLVER_TESTING`.
- Результаты этапа 6 проверяют численный алгоритм и равновесие, но не являются окончательно верифицированным нормативным расчетом.
- Инженерная сверка с ручными расчетами и внешними программами еще не выполнена.

## Следующий этап

Следующий этап по плану - расчет ширины раскрытия уже образовавшихся нормальных трещин, без расчета образования трещин.
