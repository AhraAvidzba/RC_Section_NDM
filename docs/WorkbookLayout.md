# Схема книги

Статус: обновлено по результатам этапа 9. Собираемый шаблон находится в `workbook/output/RC_Section_NDM.xlsm`, механизм сборки - в `tools/build_workbook/`.

## Лист `Расчет`

На листе созданы четыре горизонтальных блока слева направо:

1. `Исходные данные`
2. `Расчёт Mx`
3. `Расчёт My`
4. `Расчёт Mxy`

Печать:

- формат А4;
- книжная ориентация;
- каждый блок начинается с новой страницы;
- один блок занимает одну страницу;
- используются вертикальные разрывы страниц;
- область печати `Расчет!$A$1:$BT$60`;
- ручные вертикальные разрывы перед колонками `S`, `AK`, `BC`;
- масштаб печати 80%;
- горизонтальные разрывы отсутствуют.

## Блок исходных данных

На листе `Расчет` размещаются только основные параметры для лаконичного печатного оформления:

- объект или элемент;
- тип сечения;
- основные размеры;
- радиусы углов;
- класс бетона;
- класс арматуры;
- основные данные армирования;
- защитный слой;
- таблица сочетаний нагрузок;
- допустимая ширина раскрытия трещин;
- таблица арматуры.

## Блоки результатов

Блоки `Mx`, `My`, `Mxy` являются областями вывода без расчетных формул на листе.

Результаты поиска несущей способности записываются:

- `Mx` в `rngResultMx`;
- `My` в `rngResultMy`.
- `Mxy` в `rngResultMxy`.

Актуальные адреса диапазонов:

```text
rngRebarInput       = Расчет!$A$25:$G$34
rngLoadCombinations = Расчет!$A$40:$G$60
rngResultMx         = Расчет!$S$17:$AH$35
rngResultMy         = Расчет!$AK$17:$AZ$35
rngResultMxy        = Расчет!$BC$17:$BR$35
```

Запись выполняет Excel-адаптер `CCapacityResultWriter`. Расчетный класс `CCapacitySolver` к листам Excel не обращается.

На этапе 7 в те же блоки добавляются результаты ширины раскрытия уже образовавшихся трещин:

- расчетная ширина раскрытия;
- допустимая ширина;
- коэффициент использования;
- максимальная деформация и напряжение растянутой арматуры;
- количество растянутых стержней.

В блоках результатов выводятся:

- внешние усилия;
- найденные параметры деформационной плоскости;
- внутренние усилия;
- невязки равновесия;
- удерживающий момент или предельный множитель;
- коэффициент запаса;
- коэффициент использования;
- критерий предельного состояния;
- причина остановки;
- краткая диагностика сходимости.

## Лист `System`

Создан лист `System`. В нем размещаются служебные смысловые блоки:

- геометрия и дискретизация;
- материалы;
- нормативные параметры;
- настройки решателя;
- настройки поиска несущей способности;
- настройки расчета ширины раскрытия трещин;
- диагностика.

Технические параметры организованы таблицей:

```text
Ключ | Значение | Единица | Комментарий | Нормативная ссылка
```

Запись журнала в будущем должна выполняться блоками после расчета, а не по одной ячейке на каждой итерации.

Настройки поиска несущей способности включают ключи:

```text
Capacity.InitialLambda
Capacity.MaxLambda
Capacity.ToleranceLambda
Capacity.MaxRetries
Capacity.BaseLoadSteps
Capacity.SolverMaxIterations
Capacity.ConcreteCompressionLimit
Capacity.SteelStrainLimit
```

Предельные деформации в этих ключах на этапе 5 имеют статус `PROVISIONAL_FOR_SOLVER_TESTING` до окончательной нормативной верификации.

Настройки трещин включают ключи:

```text
Concrete.TensionMode
CrackWidth.Enabled
CrackWidth.Allowable
CrackWidth.CrackSpacing
CrackWidth.StrainFactor
CrackWidth.DurationFactor
```

Для расчета уже образовавшихся трещин `Concrete.TensionMode` по умолчанию равен `Ignore`.

## Минимальный набор именованных диапазонов

Созданы только крупные логические диапазоны:

- `rngMainInput`
- `rngRebarInput`
- `rngLoadCombinations`
- `rngResultMx`
- `rngResultMy`
- `rngResultMxy`
- `rngSystemSettings`
- `rngSystemDiagnostics`

Отдельные имена для каждой ячейки на этапе 1 не создаются.

## Таблица арматуры

Диапазон `rngRebarInput` имеет столбцы:

```text
ID | X | Y | Diameter | Area | SteelClass | Comment
```

Проект поддерживает только обычную ненапрягаемую арматуру. Поля для преднапряжения и начальных деформаций в пользовательской таблице не предусматриваются.

## Таблица сочетаний

Диапазон `rngLoadCombinations` имеет столбцы:

```text
CombinationID | N | Mx | My | CalculationType | DurationType | Comment
```

## Проверка

Структура книги проверяется скриптом:

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Validate-Workbook.ps1
```

Проверка подтверждает наличие листов, именованных диапазонов, печатной области, вертикальных разрывов, порядка блоков, заголовков таблиц и отсутствие формул/расчетной логики НДМ.

## Связь с геометрией этапа 2

Лист `Расчет` по-прежнему содержит только лаконичные исходные данные и печатные области. Геометрическая реализация этапа 2 импортируется в VBA-проект книги, но не подключена к формулам листа и не создает фиктивных результатов.

Технические параметры сетки остаются на листе `System` в `rngSystemSettings`, включая ключи `Geometry.Type`, `Mesh.StepX`, `Mesh.StepY` и `Mesh.BoundaryMode`. На этапе 2 эти параметры документируют будущий ввод, но расчет сетки запускается только из VBA-тестов.

## Нормативные параметры на System

Технические параметры, нормативные коэффициенты и параметры материалов организуются таблицей:

```text
Ключ | Значение | Значение по умолчанию | Единица | Назначение | Нормативный источник | Изменено пользователем
```

Колонка `Значение` используется расчетом. Колонка `Значение по умолчанию` хранит подтвержденное нормативное значение из СП. Если пользователь изменил значение, это фиксируется в колонке `Изменено пользователем`, а расчетная диагностика должна показывать, что применено пользовательское значение.

Сборщик книги не должен подставлять неподтвержденные нормативные числа. До трассировки СП такие значения заполняются как `TODO`.

## Параметры диаграмм материалов

На листе `System` размещены точки пользовательских диаграмм материалов:

- `Concrete.Point1.Eps`;
- `Concrete.Point1.Stress`;
- `Concrete.Point2.Eps`;
- `Concrete.Point2.Stress`;
- `Concrete.Point3.Eps`;
- `Concrete.Point3.Stress`;
- `Steel.Point1.Eps`;
- `Steel.Point1.Stress`;
- `Steel.Point2.Eps`;
- `Steel.Point2.Stress`;
- `Steel.Point3.Eps`;
- `Steel.Point3.Stress`.

Отдельных переключателей `Concrete.Diagram` и `Steel.Diagram` больше нет. Расчет всегда получает три пользовательские точки из `System` плюс начало координат. Для формы, эквивалентной двухлинейной диаграмме, одна из точек лежит на той же ветви или площадке; для трехлинейной пользователь меняет координаты точек. Подтвержденные значения по СП 35 и СП 63 заполнены в колонках `Значение` и `Значение по умолчанию`. Параметры, для которых формулы или обозначения в RTF/PDF распознаны неоднозначно, оставлены как `TODO` с источником и причиной.
## Stage 8 update - batch result layout

`rngLoadCombinations` now covers:

```text
Расчет!$A$37:$G$57
```

This gives one header row and up to 20 load-combination rows.

The print area now covers:

```text
Расчет!$A$1:$BT$57
```

The four horizontal A4 blocks remain arranged left to right:

```text
[Исходные данные] [Расчет Mx] [Расчет My] [Расчет Mxy]
```

Batch output is written to the `System` sheet starting at row 92. The summary contains the governing combination ID and name/comment, elapsed calculation time, per-combination statuses, lambda values, crack width, worst utilization, and diagnostics.

The workbook keeps the same minimal named ranges. No per-cell names were added for batch output.
## Stage 9 update - circular geometry settings

The workbook keeps the existing `Расчет` layout and the rounded-rectangle input as the default visible scenario.

Stage 9 added visible circular-section parameters to `System`:

```text
Circle.Diameter
Circle.Radius
Circle.CenterX
Circle.CenterY
```

These parameters document and support the new circular geometry in the calculation core and tests. A separate printed circular-section input form is not added in Stage 9.

## Пользовательский сценарий круглого сечения

После доработки пользовательского интерфейса круглое сечение доступно непосредственно на листе `Расчет`.

В блоке `Исходные данные` добавлен выбор типа сечения:

```text
RoundedRectangle
Circle
```

При выборе `Circle` макрос `RunSectionCalculation` использует поля:

```text
Circle.Diameter
Circle.CenterX
Circle.CenterY
```

Поля ширины, высоты и радиусов скругления в этом режиме не влияют на расчет.

На листе `Расчет` размещены обычные фигуры-кнопки, не ActiveX:

```text
Выполнить расчет -> RunSectionCalculation
Очистить результаты -> ClearSectionResults
```

Кнопки находятся вне печатной области. Пользовательская инструкция для ручного расчета круглого сечения приведена в `docs/UserGuideCircle.md`.

## Быстрый пользовательский режим

На листе `System` добавлены переключатели производительности:

```text
Calculation.Mode
Capacity.CalculateMx
Capacity.CalculateMy
Capacity.CalculateMxy
Solver.LoadSteps
Solver.DiagnosticsEnabled
Capacity.MaxRetries
Capacity.BaseLoadSteps
```

Значения по умолчанию приближены к прямому расчету состояния, как в reference-подходе `ArbitrarySection_NDM`:

```text
Calculation.Mode = DirectState
Solver.LoadSteps = 1
Solver.DiagnosticsEnabled = No
Capacity.MaxRetries = 0
Capacity.BaseLoadSteps = 1
```

В этом режиме книга не ищет `lambdaUltimate` и несущий момент, а выполняет прямой расчет НДС по заданным `N`, `Mx`, `My`. Отключенные результаты прочности выводятся как `NotCalculated`.

Доступные значения `Calculation.Mode`:

```text
DirectState  - НДМ по заданным усилиям без поиска несущей способности
FullCapacity - полный поиск lambdaUltimate и несущего момента
LinearMatrix - один линейно-упругий матричный расчет 3 x 3
```

`LinearMatrix` является быстрым оценочным режимом. Он не использует нелинейные диаграммы материалов, не является расчетом НДМ и не рассчитывает ширину раскрытия трещин.
