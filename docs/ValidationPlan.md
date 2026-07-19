# План проверки

Статус: обновлено по результатам этапа 9.

## Структура книги

Команда:

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Validate-Workbook.ps1
```

Проверяет:

- листы `Расчет` и `System`;
- обязательные именованные диапазоны;
- отсутствие дублирующихся имен;
- область печати;
- вертикальные разрывы страниц;
- расположение четырех блоков слева направо;
- отсутствие формул и расчетной логики НДМ.

## Геометрия и волокна

Команда:

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-GeometryTests.ps1
```

Проверяет:

- прямоугольник без скруглений против аналитической площади, центра тяжести, `Ix`, `Iy`, `Ixy`;
- симметричный скругленный прямоугольник против аналитической площади и симметрии;
- асимметричные радиусы и ненулевой `Ixy`;
- недопустимые размеры, радиусы и шаги сетки;
- сходимость площади и моментов инерции при уменьшении шага сетки;
- базовую производительность построения сетки.

## Последний результат

После закрытия зависших процессов Excel и исправления проверки скругленных углов геометрические тесты выполнены успешно:

```text
TOTAL: passed=29; failed=0
```

Если Excel COM снова начнет возвращать ошибки открытия или создания книги, нужно закрыть все процессы Excel в пользовательской сессии и повторить команды сборки и проверки.

## Линейное ядро

Команда:

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-LinearTests.ps1
```

Проверяет:

- решение системы `3 x 3` с выбором главного элемента;
- диагностику вырожденной матрицы;
- центральное сжатие;
- изгиб относительно оси `X`;
- изгиб относительно оси `Y`;
- совместное действие `N + Mx + My`;
- несимметричную геометрию с ненулевой связью кривизн;
- автоматический расчет площади обычного стержня по диаметру;
- эффективный вклад арматуры `Es - Eb`;
- ошибки данных арматуры;
- базовую производительность матрицы, решения и обратного расчета.

Последний результат:

```text
TOTAL_LINEAR: passed=49; failed=0
```

## Нормативные параметры и тестовые нагрузки

Структурная проверка книги должна подтверждать заголовки таблицы `System`:

```text
Ключ | Значение | Значение по умолчанию | Единица | Назначение | Нормативный источник | Изменено пользователем
```

Тестовые нагрузки могут быть произвольными. Для аналитических тестов фиксируются `N`, `Mx`, `My`, единицы внутренних усилий и независимая проверка через классические формулы или другой прозрачный расчет.

Справочные программы не являются обязательным источником тестовых нагрузок; они остаются дополнительными сравнительными материалами.

## Диаграммы материалов

Команда:

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-MaterialTests.ps1
```

Проверяет:

- нулевое напряжение растянутого бетона в бетонных диаграммах;
- линейный участок двухлинейной диаграммы бетона;
- площадку двухлинейной диаграммы бетона;
- трехлинейную интерполяцию бетона по заданным точкам;
- симметрию обычной арматуры при растяжении и сжатии;
- площадку двухлинейной диаграммы арматуры;
- трехлинейную интерполяцию арматуры;
- отбраковку некорректных параметров.

Формулы СП 63, которые не извлекаются текстом из RTF/PDF, не проверяются численно до ручного подтверждения.

## Нелинейный решатель этапа 4

Команда:

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-SectionSolverTests.ps1
```

Проверяет численный алгоритм, а не окончательное соответствие СП:

- чтение временных параметров из `rngSystemSettings`;
- совпадение `CSectionSolver` с линейным решателем на линейных материалах;
- нелинейное замещение бетона арматурой через разность напряжений и касательных модулей;
- сходимость полного Ньютона при центральном сжатии;
- сходимость с временной двухлинейной диаграммой бетона и обычной A400;
- пошаговое приложение нагрузки;
- ограничение приращений;
- наличие диагностического предупреждения `PROVISIONAL_FOR_SOLVER_TESTING`.

Тесты этапа 4 не проверяют:

- поиск предельного множителя;
- удерживающий момент;
- коэффициент запаса;
- коэффициент использования;
- ширину раскрытия трещин;
- окончательную нормативную верификацию диаграмм.

Последний результат:

```text
TOTAL_SECTION_SOLVER: passed=34; failed=0
```

## Несущая способность Mx и My

Команда:

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-CapacityTests.ps1
```

Проверяет:

- поиск предельного множителя для положительного и отрицательного `Mx`;
- поиск предельного множителя для положительного и отрицательного `My`;
- поиск предельного множителя для `N + Mx + My`;
- сохранение направления вектора моментов `Mx = lambda * MxBase`, `My = lambda * MyBase`;
- четыре знаковых сочетания `Mx/My` для `Mxy`;
- случаи `lambdaUltimate < 1` и `lambdaUltimate` около `1`;
- случай `N = 0`;
- физические предельные состояния `ConcreteStrainLimit` и `SteelStrainLimit`;
- численную несходимость как `NumericalFailure` или `SingularTangent`, а не как физическую границу;
- выполнение равновесия для последнего допустимого состояния;
- независимую контрольную невязку удерживающего момента;
- несимметричное сечение с `Ixy <> 0` и связанной кривизной;
- запись результатов в `rngResultMx`, `rngResultMy` и `rngResultMxy`;
- обработку нулевого базового момента как `InvalidInput`.

Последний результат:

```text
TOTAL_CAPACITY: passed=92; failed=0
```

Из-за зависшего процесса Excel полная пересборка через `Workbooks.Add()` была заменена обновлением существующей книги:

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Refresh-Workbook.ps1
```

После refresh также пройдены:

```text
Validate-Workbook.ps1: все проверки Passed=True
TOTAL: passed=29; failed=0
TOTAL_LINEAR: passed=49; failed=0
TOTAL_MATERIAL: passed=20; failed=0
TOTAL_SECTION_SOLVER: passed=34; failed=0
```

## Ширина раскрытия трещин

Команда:

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-CrackTests.ps1
```

Проверяет:

- `Concrete.TensionMode = Ignore`;
- `Concrete.TensionMode = UseDiagram`;
- расчет трещин по уже найденному SLS-состоянию `CSectionSolver`;
- расчет для `Mx`, `My` и `Mxy`;
- определение растянутой арматуры;
- деформации и напряжения арматуры;
- расчет ширины раскрытия и коэффициента использования;
- запись результатов в соответствующий блок книги.

Последний результат:

```text
TOTAL_CRACK: passed=28; failed=0
```
## Stage 8 validation

Batch calculation tests are run with:

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-BatchTests.ps1
```

The test module checks:

- one load combination;
- five load combinations;
- twenty load combinations;
- invalid input row read from `rngLoadCombinations`;
- batch summary writer output on `System`;
- elapsed-time reporting;
- governing-combination detection.

Stage 8 also extends capacity tests with `Concrete.TensionMode` regression checks for the shared `CSectionSolver` in `Mx`, `My`, and `Mxy` strength-state scenarios.

Latest Stage 8 regression results:

```text
Validate-Workbook.ps1: all checks Passed=True
TOTAL: passed=29; failed=0
TOTAL_LINEAR: passed=49; failed=0
TOTAL_MATERIAL: passed=20; failed=0
TOTAL_SECTION_SOLVER: passed=34; failed=0
TOTAL_CAPACITY: passed=104; failed=0
TOTAL_CRACK: passed=28; failed=0
TOTAL_BATCH: passed=17; failed=0
```
## Stage 9 validation

Stage 9 adds tests for the circular section.

Geometry validation checks:

- analytical circle area `A = pi * R^2`;
- centroid coordinates;
- `Ix = Iy = pi * R^4 / 4`;
- `Ixy = 0`;
- invalid radius and diameter inputs.

Calculation validation checks:

- circular-section capacity for `Mx`;
- circular-section capacity for `My`;
- circular-section capacity for `Mxy`;
- symmetry of `Mx` and `My`;
- crack-width calculation for an already cracked circular section.

Latest Stage 9 results:

```text
Validate-Workbook.ps1: all checks Passed=True
TOTAL: passed=39; failed=0
TOTAL_LINEAR: passed=49; failed=0
TOTAL_MATERIAL: passed=20; failed=0
TOTAL_SECTION_SOLVER: passed=34; failed=0
TOTAL_CAPACITY: passed=110; failed=0
TOTAL_CRACK: passed=34; failed=0
```
