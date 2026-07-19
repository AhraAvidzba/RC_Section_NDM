# SETTINGS_REFERENCE

Дата фиксации: 2026-07-19.

Источник настроек пользовательского расчета: лист `System`, диапазон `rngSystemSettings`.
Расчетное ядро не обращается к листам Excel. Excel-слой один раз загружает `System`
через `CSystemSettingsReader` и передает параметры в геометрию, материалы, решатели,
batch-калькулятор и расчет трещин.

## Именованные диапазоны

| Диапазон | Назначение |
|---|---|
| `GeometrySettings` | Тип сечения, параметры круга и волоконной сетки. |
| `ConcreteDiagram` | Класс бетона, статус источника, сопротивления, модуль и параметры диаграммы бетона. |
| `SteelDiagram` | Класс обычной ненапрягаемой арматуры, сопротивления, модуль и параметры диаграммы арматуры. |
| `SolverSettings` | Режим расчета и параметры прямого решателя равновесия `CSectionSolver`. |
| `CapacitySettings` | Параметры поиска `lambdaUltimate` в `CCapacitySolver`. |
| `OutputSettings` | Расчет трещин и параметры прозрачного формульного вывода. |
| `AutoCADSettings` | Зарезервированный именованный блок для уже существующей выгрузки в AutoCAD; рабочих ключей сейчас нет. |

## Рабочие ключи

| Ключ | Где читается | Где используется |
|---|---|---|
| `Geometry.Type` | `modWorkbookCalculation`, `modAutoCADStressExport` | Выбор `RoundedRectangle` или `Circle`; лист `Расчет` только отображает значение. |
| `Circle.Diameter` | `modWorkbookCalculation`, `modAutoCADStressExport` | Построение круглой геометрии и автоматической арматуры. |
| `Circle.CenterX` | `modWorkbookCalculation`, `modAutoCADStressExport` | Центр круга и координаты арматуры. |
| `Circle.CenterY` | `modWorkbookCalculation`, `modAutoCADStressExport` | Центр круга и координаты арматуры. |
| `Mesh.StepX` | `modWorkbookCalculation`, `modAutoCADStressExport` | Шаг волоконной сетки по X. |
| `Mesh.StepY` | `modWorkbookCalculation`, `modAutoCADStressExport` | Шаг волоконной сетки по Y. |
| `Mesh.BoundarySubdivisions` | `modWorkbookCalculation`, `modAutoCADStressExport` | Дробление граничных ячеек; `1` означает расчет по центру базовой ячейки. |
| `Concrete.Eb` | материалы, `LinearMatrix`, формульный блок трещин | Модуль бетона и растянутая ветвь `UseDiagram`. |
| `Concrete.TensionMode` | материалы бетона | `Ignore` или `UseDiagram` для всего `CSectionSolver`. |
| `Concrete.Point*` | `modWorkbookCalculation`, `modAutoCADStressExport` | Пользовательские точки диаграммы бетона. |
| `Steel.Es` | `LinearMatrix`, формульный блок трещин | Модуль арматуры. |
| `Steel.Point*` | `modWorkbookCalculation`, `modAutoCADStressExport` | Пользовательские точки диаграммы арматуры. |
| `Calculation.Mode` | `modWorkbookCalculation`, `CBatchSectionCalculator` | `DirectState`, `FullCapacity` или `LinearMatrix`. |
| `Solver.*` | `CSectionSolver`, `CCapacitySolver`, `CBatchSectionCalculator` | Итерации, ступени, допуски, line search, ограничения приращений и диагностика. |
| `Capacity.*` | `CCapacitySolver`, `CBatchSectionCalculator` | Параметры поиска предельного множителя и предельные деформации. |
| `CrackWidth.*` | `CCrackWidthCalculator`, `CBatchSectionCalculator`, формульный блок трещин | Включение и параметры расчета уже образовавшихся нормальных трещин. |

## Удаленные ключи

| Ключ | Причина |
|---|---|
| `Capacity.Enabled` | Дублировал `Calculation.Mode`. |
| `Capacity.CalculateMx` | Отдельная ветвь `Mx` удалена; одноосный случай считается общей постановкой. |
| `Capacity.CalculateMy` | Отдельная ветвь `My` удалена; одноосный случай считается общей постановкой. |
| `Capacity.CalculateMxy` | Отдельный переключатель направления больше не нужен. |
| `Mesh.BoundaryMode` | Дублировал `Mesh.BoundarySubdivisions`; значение `1` задает быстрый режим без отдельного флага. |
| `Circle.Radius` | Дублировал `Circle.Diameter`; радиус вычисляется как `D / 2`. |
| `Batch.MaxCombinations` | Не читался расчетом; лимит 20 строк задается размером `rngLoadCombinations`. |
| `Batch.Diagnostics` | Не читался расчетом; пакетная диагностика выводится штатным writer-ом без отдельного флага. |

## Выпадающие списки

Списки добавлены только для реально поддерживаемых значений:

| Ключ | Значения |
|---|---|
| `Geometry.Type` | `RoundedRectangle`, `Circle` |
| `Calculation.Mode` | `DirectState`, `FullCapacity`, `LinearMatrix` |
| `Concrete.TensionMode` | `Ignore`, `UseDiagram` |
| `Solver.LineSearchEnabled` | `Yes`, `No` |
| `Solver.DiagnosticsEnabled` | `Yes`, `No` |
| `CrackWidth.Enabled` | `Yes`, `No` |

Фиктивные варианты `Newton/Secant/Brent`, `SolveByUltimateStrain` и альтернативные
алгоритмы поиска `lambda` не добавлялись.
