# TEST_CASES

Дата фиксации последнего полного прогона: 2026-07-19.

Статус регрессионной базы: `Stage01 TEMPORARY_BASELINE`. Это снимок текущего
поведения расчетного ядра для безопасной архитектурной очистки; он не является
окончательной нормативной верификацией по СП.

## Архитектурное Решение

Целевая постановка проекта: только единый совместный расчет `N + Mx + My`.

Одноосный изгиб сохраняется математически только как частный случай общей
постановки:

```text
N + Mx, My = 0
N + My, Mx = 0
```

Отдельные пользовательские и расчетные ветви `Mx` и `My` удалены на предыдущем
этапе. Оставшиеся упоминания `Mx` и `My` допустимы только как компоненты
вектора моментов, компоненты равновесия, столбцы нагрузок, допуски
`Solver.ToleranceMx/My` и диагностический текст.

## Единый Запуск

Команда:

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-AllTests.ps1
```

Итоговые файлы:

```text
docs/regression/Stage01_AllTests_Report.txt
docs/regression/Stage01_RegressionBaseline_Raw.txt
```

Сводка последнего прогона:

| Блок | Результат |
|---|---:|
| Validate-Workbook.ps1 | passed |
| Geometry | 56 / 0 |
| Linear core | 49 / 0 |
| Materials | 20 / 0 |
| Section solver | 93 / 0 |
| Capacity | 108 / 0 |
| Crack width | 34 / 0 |
| Batch | 16 / 0 |
| Workbook UI | 31 / 0 |
| Regression baseline | 33 / 0 |

## Регрессионные Сценарии

Единицы: `N`, `Nint`, `resN` - Н; `Mx`, `My`, `Mxint`, `Myint`,
`resMx`, `resMy` - Н*мм; деформации безразмерные; кривизны - 1/мм.

Базовая геометрия, если не указано иначе: круг `D = 300 мм`, центр `(0; 0)`,
8 стержней `d20`, `as = 40 мм`, сетка `20 x 20 мм`,
`Mesh.BoundarySubdivisions = 1`, материалы - текущие предварительные параметры.

Обязательные сценарии:

| Сценарий | Что фиксируется |
|---|---|
| Чистое сжатие | `epsilon0`, нулевые кривизны, равновесие по `N` |
| `N + Mx`, `My = 0` | общий метод при одном ненулевом моменте |
| `N + My`, `Mx = 0` | общий метод при одном ненулевом моменте |
| Полный `N + Mx + My` | совместная работа двух кривизн |
| Симметричное круглое сечение | симметрия и равновесие |
| Повторный запуск | воспроизводимость результата |
| Изменение геометрии | изменение НДС при изменении диаметра |
| Перенос начала координат | сохранение кривизн и корректное преобразование моментов |

Для каждого сценария в `Stage01_RegressionBaseline_Raw.txt` сохраняются:
`epsilon0`, `kappaX`, `kappaY`, `Nint`, `Mxint`, `Myint`, невязки, `lambda`,
число итераций, время, число волокон и число стержней.

Примечание по `lambda`: `FAIL:NumericalFailure` фиксирует текущее поведение
поиска несущей способности и не трактуется как физическое разрушение.

## Этап Настроек System

Проверки текущего этапа подтверждают:

| Проверка | Где выполняется |
|---|---|
| Все рабочие ключи читаются из `rngSystemSettings` | `tests/modTestSectionSolver.bas` |
| Дубли рабочих ключей отсутствуют | `CSystemSettingsReader.DuplicateCount`, `Validate-Workbook.ps1` |
| Удаленные ключи отсутствуют в книге | `Validate-Workbook.ps1`, `tests/modTestSectionSolver.bas` |
| Смысловые диапазоны настроек существуют | `SolverSettings`, `CapacitySettings`, `ConcreteDiagram`, `SteelDiagram`, `GeometrySettings`, `OutputSettings`, `AutoCADSettings` |
| `Capacity.MaxLambda` доходит из `System` до расчета и не перекрывается кодом | `tests/modTestBatchCalculation.bas` |
| `Geometry.Type` имеет один источник истины | `System`; лист `Расчет` содержит только связанные пользовательские поля |

Удаленные или объединенные настройки:

| Ключ | Решение |
|---|---|
| `Capacity.Enabled` | удален как дублирующий `Calculation.Mode` |
| `Capacity.CalculateMx` | удален вместе с отдельной ветвью `Mx` |
| `Capacity.CalculateMy` | удален вместе с отдельной ветвью `My` |
| `Capacity.CalculateMxy` | удален; направление всегда задается вектором `MxBase/MyBase` |
| `Mesh.BoundaryMode` | удален; `Mesh.BoundarySubdivisions = 1` задает быстрый режим |
| `Circle.Radius` | удален; для круга используется `Circle.Diameter` |
| `Batch.MaxCombinations` | удален; фактический лимит задается размером `rngLoadCombinations` |
| `Batch.Diagnostics` | удален; не читался расчетом |

Подробная карта настроек хранится в `docs/SETTINGS_REFERENCE.md`.
