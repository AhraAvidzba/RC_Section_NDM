# Current Architecture Audit

Дата аудита: 2026-07-19.

Аудит выполнен для текущего Excel/VBA-проекта `RC_Section_NDM`.
Расчетный код на этом этапе не изменялся. Изменены только проектные
документы и одна строка тестовой инфраструктуры `Validate-Workbook.ps1`, где
был найден буквальный фрагмент `` `r`n `` в коде PowerShell.

## 1. Область проверки

Проверены:

- структура каталогов `src`, `tests`, `tools/build_workbook`, `docs`, `workbook`;
- экспортированные VBA-модули `.bas` и `.cls`;
- точки входа пользовательского интерфейса;
- классы геометрии, сетки, материалов, решателей, пакетного расчета и Excel-адаптеров;
- скрипты сборки и запуска тестов;
- именованные диапазоны по скриптам сборки;
- фактические места чтения настроек из `rngSystemSettings`;
- обращения к листам Excel внутри VBA.

Не проверялись:

- файлы из `reference/`, `norms/`, `control_examples/` не изменялись;
- полное чтение книги через COM не завершилось из-за зависшего `EXCEL.EXE`;
- фактическое содержимое `workbook/output/RC_Section_NDM.xlsm` проверено только по скриптам и исходникам, не через успешное COM-чтение.

## 2. Структура проекта

Основные каталоги:

| Каталог | Назначение |
|---|---|
| `src/Batch` | пакетный расчет сочетаний |
| `src/Common` | общие типы геометрии |
| `src/Crack` | расчет ширины раскрытия трещин |
| `src/Excel` | слой Excel, кнопки, чтение диапазонов, вывод результатов, AutoCAD export |
| `src/Geometry` | геометрии и построение бетонной сетки |
| `src/Interfaces` | интерфейс геометрии |
| `src/Materials` | диаграммы бетона и арматуры |
| `src/Section` | арматура и геометрические характеристики |
| `src/Solver` | решатель равновесия, линейная система, несущая способность |
| `tests` | VBA-тесты |
| `tools/build_workbook` | сборка, обновление книги, запуск тестов |
| `workbook/output` | собранная книга `RC_Section_NDM.xlsm` |

## 3. Карта VBA-модулей и классов

| Файл | Модуль | Роль |
|---|---|---|
| `src/Common/modGeometryTypes.bas` | `modGeometryTypes` | константы и типы геометрии, включая `TFiber` |
| `src/Geometry/CFiberMeshBuilder.cls` | `CFiberMeshBuilder` | строит массив бетонных волокон |
| `src/Geometry/CGeometryCircle.cls` | `CGeometryCircle` | параметрическая геометрия круга |
| `src/Geometry/CGeometryRoundedRectangle.cls` | `CGeometryRoundedRectangle` | скругленный прямоугольник |
| `src/Interfaces/ISectionGeometry.cls` | `ISectionGeometry` | интерфейс геометрии |
| `src/Section/CRebarLayout.cls` | `CRebarLayout` | массив стержней арматуры |
| `src/Section/CCircleRebarLayoutBuilder.cls` | `CCircleRebarLayoutBuilder` | автоматическая круговая расстановка арматуры |
| `src/Section/CGeometryPropertiesCalculator.cls` | `CGeometryPropertiesCalculator` | геометрические характеристики по сетке |
| `src/Materials/CConcreteBilinearMaterial.cls` | `CConcreteBilinearMaterial` | двухлинейная диаграмма бетона |
| `src/Materials/CConcreteTrilinearMaterial.cls` | `CConcreteTrilinearMaterial` | трехлинейная диаграмма бетона |
| `src/Materials/CSteelBilinearMaterial.cls` | `CSteelBilinearMaterial` | двухлинейная диаграмма арматуры |
| `src/Materials/CSteelTrilinearMaterial.cls` | `CSteelTrilinearMaterial` | трехлинейная диаграмма арматуры |
| `src/Materials/CLinearConcreteMaterial.cls` | `CLinearConcreteMaterial` | линейный бетон для тестов |
| `src/Materials/CLinearSteelMaterial.cls` | `CLinearSteelMaterial` | линейная арматура для тестов |
| `src/Solver/CLinearSystem3x3.cls` | `CLinearSystem3x3` | решение системы 3x3 |
| `src/Solver/CSectionSolver.cls` | `CSectionSolver` | основной Ньютонов решатель равновесия |
| `src/Solver/CLinearSectionSolver.cls` | `CLinearSectionSolver` | линейный матричный расчет состояния |
| `src/Solver/CCapacitySolver.cls` | `CCapacitySolver` | поиск несущей способности по множителю `lambda` |
| `src/Crack/CCrackWidthCalculator.cls` | `CCrackWidthCalculator` | временный расчет трещин по заданному НДС |
| `src/Batch/CBatchSectionCalculator.cls` | `CBatchSectionCalculator` | пакетный расчет до 20 сочетаний |
| `src/Excel/CSystemSettingsReader.cls` | `CSystemSettingsReader` | чтение `rngSystemSettings` |
| `src/Excel/CLoadCombinationReader.cls` | `CLoadCombinationReader` | чтение `rngLoadCombinations` |
| `src/Excel/CCapacityResultWriter.cls` | `CCapacityResultWriter` | вывод результатов в `rngResultMx/My/Mxy` |
| `src/Excel/CBatchResultWriter.cls` | `CBatchResultWriter` | вывод сводки batch на `System` |
| `src/Excel/modWorkbookCalculation.bas` | `modWorkbookCalculation` | пользовательский расчет книги |
| `src/Excel/modAutoCADStressExport.bas` | `modAutoCADStressExport` | экспорт геометрии и напряжений в AutoCAD |

## 4. Точки входа

Пользовательские макросы:

| Макрос | Файл | Назначение |
|---|---|---|
| `RunSectionCalculation` | `modWorkbookCalculation.bas` | основной запуск расчета по книге |
| `ClearSectionResults` | `modWorkbookCalculation.bas` | очистка результатов и диагностики |
| `ExportSectionStressToAutoCAD` | `modAutoCADStressExport.bas` | выгрузка сетки, арматуры и напряжений в AutoCAD |

Тестовые точки входа:

| Макрос | Скрипт запуска |
|---|---|
| `RunGeometryTests` | `Run-GeometryTests.ps1` |
| `RunLinearCoreTests` | `Run-LinearTests.ps1` |
| `RunMaterialDiagramTests` | `Run-MaterialTests.ps1` |
| `RunSectionSolverTests` | `Run-SectionSolverTests.ps1` |
| `RunCapacitySolverTests` | `Run-CapacityTests.ps1` |
| `RunCrackWidthTests` | `Run-CrackTests.ps1` |
| `RunBatchCalculationTests` | `Run-BatchTests.ps1` |
| `RunWorkbookInterfaceTests` | `Run-WorkbookInterfaceTests.ps1` |

## 5. Листы и именованные диапазоны

По скриптам сборки/обновления используются листы:

- `Расчет`;
- `System`.

Именованные диапазоны:

| Имя | Диапазон по текущим скриптам | Назначение |
|---|---|---|
| `rngMainInput` | `Расчет!A6:Q22` в `Refresh-Stage09.ps1`; `A6:Q20` в старом `Build-Workbook.ps1` | основные исходные данные |
| `rngRebarInput` | `Расчет!A25:G34` | контрольная таблица автоматически рассчитанной арматуры |
| `rngLoadCombinations` | `Расчет!A40:G60` | до 20 сочетаний нагрузок |
| `rngResultMx` | `Расчет!S17:AH35` | блок результатов Mx |
| `rngResultMy` | `Расчет!AK17:AZ35` | блок результатов My |
| `rngResultMxy` | `Расчет!BC17:BR35` | блок результатов Mxy |
| `rngSystemSettings` | формируется на `System` | таблица настроек |
| `rngSystemDiagnostics` | формируется на `System` | диагностика |

Дополнительные технические зоны:

- `System!A92:M125` - batch summary и диагностика;
- `System!A130:G250` - полный список автоматически созданной арматуры;
- `System!I130:Q250` - прозрачный формульный блок расчета трещин.

## 6. Поток данных

Текущий основной поток:

1. `RunSectionCalculationForWorkbook` очищает результаты.
2. `CSystemSettingsReader.LoadFromWorkbook` читает `rngSystemSettings`.
3. `ReadGeometry` читает `rngMainInput` и создает `CGeometryCircle` или `CGeometryRoundedRectangle`.
4. `CFiberMeshBuilder.BuildMesh` строит бетонные волокна.
5. `ReadRebars` читает параметры круга и арматуры из `rngMainInput`, создает `CRebarLayout` через `CCircleRebarLayoutBuilder`.
6. Создаются материалы `CConcreteBilinearMaterial` и `CSteelBilinearMaterial`.
7. `CLoadCombinationReader` читает непустые строки `rngLoadCombinations`.
8. `CBatchSectionCalculator.Execute` рассчитывает все сочетания.
9. `CBatchResultWriter` пишет сводку на `System`.
10. `WriteFirstCombinationResults` дополнительно пишет подробные результаты первого заполненного сочетания в блоки `rngResultMx/My/Mxy`.

Упрощенная схема:

```text
Excel ranges
  -> readers
  -> geometry + mesh + rebar + materials
  -> CBatchSectionCalculator
     -> CCapacitySolver / CSectionSolver / CCrackWidthCalculator
  -> writers
  -> Расчет / System
```

## 7. Геометрия и сетка

Геометрия формируется в:

- `modWorkbookCalculation.ReadGeometry`;
- `modAutoCADStressExport.ReadExportGeometry`;
- тестовых фабриках в `tests/*`.

Поддерживаются:

- `CGeometryRoundedRectangle`;
- `CGeometryCircle`.

Сетка бетонных волокон хранится в `CFiberMeshBuilder.mFibers()` как типизированный массив `TFiber`.

Сетка строится методом:

```vba
CFiberMeshBuilder.BuildMesh geometry, stepX, stepY, materialID, boundarySubdivisions
```

Фактический алгоритм:

- обход прямоугольного bounding box;
- при `boundarySubdivisions = 1` включается ячейка, если центр внутри сечения;
- при `boundarySubdivisions > 1` внутренние базовые ячейки остаются крупными, граничные дробятся на подячейки;
- `FillFactor` всегда равен `1`, частичный коэффициент заполнения не используется.

Кэш сетки:

- отдельного кэша геометрии/сетки нет;
- при каждом `RunSectionCalculationForWorkbook` сетка строится заново;
- пакетный расчет переиспользует одну сетку внутри одного запуска.

## 8. Арматура

Арматура хранится в `CRebarLayout` в отдельных типизированных массивах:

- `ID`;
- `X`;
- `Y`;
- `Diameter`;
- `Area`;
- `SteelClass`;
- `Comment`.

Для круглого сечения пользовательский ввод координат исключен. Координаты
создаются в `CCircleRebarLayoutBuilder` по параметрам:

- `D`;
- `CenterX`;
- `CenterY`;
- `as`;
- `n`;
- `ds`;
- `SteelClass`.

Ручной ввод координат из `rngRebarInput` больше не используется как источник расчета. Таблица является контрольным выводом.

## 9. Координаты, знаки и усилия

Формула плоских сечений фактически записана в `CSectionSolver`:

```text
strain = eps0 + kx * y + ky * x
```

Внутренние усилия:

```text
Nint  = Sum(force)
Mxint = Sum(force * y)
Myint = Sum(force * x)
```

Для бетона:

```text
force = concreteStress(strain) * fiberArea * fillFactor
```

Для арматуры:

```text
effectiveStress = steelStress(strain) - concreteStress(strain)
force = effectiveStress * barArea
```

Это реализует нелинейное замещение бетона арматурой.

Начало координат задается параметрами геометрии. Для круга пользователь вводит
центр на листе `Расчет` в строках, соответствующих `Circle.CenterX` и
`Circle.CenterY`. Отдельной проверки инвариантности результата при переносе
начала координат в текущей регрессии не найдено.

## 10. Решатель равновесия

Фактически реализован один основной метод:

- полный Newton;
- касательная матрица 3x3;
- line search;
- damping;
- ограничение приращений;
- пошаговое приложение нагрузки.

Класс:

```text
CSectionSolver
```

Система 3x3 решается через:

```text
CLinearSystem3x3
```

Не найдено отдельных рабочих реализаций:

- modified Newton;
- secant equilibrium solver;
- Brent;
- safeguarded secant;
- SolveByUltimateStrain.

Упоминания `secant` есть только в методах `GetSecantModulus` материалов, это не метод решения равновесия.

## 11. Поиск несущей способности

Фактический класс:

```text
CCapacitySolver
```

Фактический метод:

- `SolveMx`;
- `SolveMy`;
- `SolveMxy`;
- общий внутренний `SolveOneDirection`;
- внешнее расширение верхней границы `lambda`;
- уточнение через bisection.

Закон масштабирования:

- `N = const`;
- `Mx = lambda * MxBase`;
- `My = lambda * MyBase`.

Для `Mx` и `My` отдельные публичные методы есть, но они являются обертками над общим `SolveOneDirection`.

Несходимость классифицируется отдельно:

- `NumericalFailure`;
- `SingularTangent`;
- `InvalidInput`;
- `ConcreteStrainLimit`;
- `SteelStrainLimit`.

Замечание: в `CBatchSectionCalculator.ConfigureCapacity` часть параметров
несущей способности задается жестко и может перекрывать значения,
прочитанные из `System`.

## 12. Трещины

Фактический класс:

```text
CCrackWidthCalculator
```

Текущий расчет трещин остается временным:

```text
CrackWidth = MaxSteelStrain * CrackSpacing * StrainFactor * DurationFactor
```

Недавно добавлен прозрачный формульный блок на `System!I130:Q250`, где:

- `epsilon0`, `kappaX`, `kappaY` берутся из `CSectionSolver`;
- координаты, деформации и напряжения стержней считаются формулами листа;
- `Ar.Interaction` пока является явной временной формулой-заготовкой.

Для нормативного расчета по СП 35/СП 63 требуется отдельный этап трассировки
формул, особенно для `Ar`.

## 13. Batch и определяющее сочетание

`CBatchSectionCalculator` хранит до 20 сочетаний в фиксированных массивах.

Пустые строки `rngLoadCombinations` игнорируются.
Частично заполненные строки добавляются как `InvalidInput`.

Определяющее сочетание выбирается по максимуму:

```text
max(UtilMx, UtilMy, UtilMxy, CrackUtil)
```

Геометрия, сетка, арматура и материалы создаются один раз на запуск и
переиспользуются для всех сочетаний.

## 14. Excel-слой и производительность

Расчетное ядро `CSectionSolver`, `CCapacitySolver`, материалы, геометрии и
batch-класс не обращаются к `Range`, `Cells`, `Worksheets`.

Обращения к Excel сосредоточены в:

- `modWorkbookCalculation`;
- `modAutoCADStressExport`;
- `CSystemSettingsReader`;
- `CLoadCombinationReader`;
- `CCapacityResultWriter`;
- `CBatchResultWriter`;
- тестовых модулях.

Нарушения новых требований производительности:

- `CSystemSettingsReader.LoadFromRange` читает настройки построчно и использует `ReDim Preserve`;
- `CLoadCombinationReader` читает сочетания построчно;
- `CCapacityResultWriter`, `CBatchResultWriter`, `modWorkbookCalculation` пишут результаты по ячейкам;
- `WriteGeneratedRebarSystemTable` и формульный блок трещин заполняются по ячейкам;
- `CRebarLayout.AddBar` использует `ReDim Preserve` при добавлении каждого стержня.

Эти места не являются внутренними итерациями решателя, но должны быть
рефакторированы на следующих этапах для больших таблиц результата.

## 15. AutoCAD

Текущий AutoCAD-функционал:

- только экспорт;
- файл `modAutoCADStressExport.bas`;
- используется Late Binding через `GetObject(, "AutoCAD.Application")`;
- дополнительных References на AutoCAD Type Library нет;
- если AutoCAD отсутствует, параметрический расчет книги не должен страдать,
  ошибка возникает только при запуске `ExportSectionStressToAutoCAD`.

Импорт готовых областей AutoCAD пока не реализован.

Экспорт строит:

- прямоугольники бетонных волокон фактических размеров `FiberWidth/FiberHeight`;
- окружности арматуры;
- текст напряжений.

Старые объекты на слоях `RC_NDM_*` сейчас не очищаются автоматически.

## 16. Таблица настроек

Столбцы текущей таблицы `rngSystemSettings`:

```text
Ключ | Значение | Значение по умолчанию | Единица | Назначение | Нормативный источник | Изменено пользователем
```

Фактически читаемые настройки:

| Параметр | Где читается | Фактическое назначение | Рекомендация |
|---|---|---|---|
| `Mesh.StepX` | `modWorkbookCalculation`, `modAutoCADStressExport` | шаг сетки по X | оставить |
| `Mesh.StepY` | `modWorkbookCalculation`, `modAutoCADStressExport` | шаг сетки по Y | оставить |
| `Mesh.BoundarySubdivisions` | `modWorkbookCalculation`, `modAutoCADStressExport` | дробление граничных ячеек | оставить |
| `Concrete.Bilinear.Eps1` | Excel-слой создания материала | точка диаграммы бетона | оставить |
| `Concrete.Bilinear.Stress1` | Excel-слой создания материала | напряжение точки диаграммы бетона | оставить |
| `Concrete.Bilinear.EpsU` | Excel-слой создания материала | предельная деформация бетона | оставить |
| `Concrete.Bilinear.StressU` | Excel-слой создания материала | напряжение предельной точки | оставить |
| `Concrete.TensionMode` | материалы бетона | Ignore/UseDiagram | оставить, добавить validation list |
| `Concrete.Eb` | материалы, LinearMatrix, формулы трещин | модуль бетона | оставить |
| `Concrete.Rbt.SLS` | материалы бетона | предел растянутой ветви UseDiagram | оставить, уточнить нормативность |
| `Steel.Bilinear.EpsY` | Excel-слой создания материала | деформация текучести | оставить |
| `Steel.Bilinear.StressY` | Excel-слой создания материала | напряжение текучести | оставить |
| `Steel.Bilinear.EpsU` | Excel-слой создания материала | предельная деформация | оставить |
| `Steel.Es` | LinearMatrix, формулы трещин | модуль арматуры | оставить |
| `Steel.Rs.ULS` | формулы трещин | предел растяжения в формульном блоке | проверить режим ULS/SLS |
| `Steel.Rsc.ULS` | формулы трещин | предел сжатия в формульном блоке | проверить режим ULS/SLS |
| `Materials.SourceStatus` | batch, solver, crack, writer | статус временных параметров | оставить |
| `Calculation.Mode` | batch, first-combination writer | `DirectState`, `FullCapacity`, `LinearMatrix` | оставить, добавить validation list |
| `Solver.MaxIterations` | `CSectionSolver`, batch | максимум итераций равновесия | оставить |
| `Solver.LoadSteps` | `CSectionSolver`, batch | ступени приложения нагрузки | оставить |
| `Solver.ToleranceN` | `CSectionSolver`, batch | допуск N | оставить |
| `Solver.ToleranceMx` | `CSectionSolver`, batch | допуск Mx | оставить |
| `Solver.ToleranceMy` | `CSectionSolver`, batch | допуск My | оставить |
| `Solver.LineSearchEnabled` | `CSectionSolver` | включение line search | оставить |
| `Solver.DampingInitial` | `CSectionSolver` | начальный коэффициент шага Newton | оставить |
| `Solver.MinLineSearchAlpha` | `CSectionSolver` | минимальный alpha line search | оставить |
| `Solver.MaxDeltaEpsilon0` | `CSectionSolver` | ограничение приращения epsilon0 | оставить |
| `Solver.MaxDeltaKappa` | `CSectionSolver` | ограничение приращения кривизн | оставить |
| `Solver.DiagnosticsEnabled` | `CSectionSolver`, batch | журнал итераций | оставить |
| `Capacity.InitialLambda` | `CCapacitySolver` | начальный шаг lambda | оставить |
| `Capacity.MaxLambda` | `CCapacitySolver` | верхняя граница поиска | оставить |
| `Capacity.ToleranceLambda` | `CCapacitySolver` | допуск bisection | оставить |
| `Capacity.MaxRetries` | `CCapacitySolver` | повторы при численном отказе | оставить |
| `Capacity.BaseLoadSteps` | `CCapacitySolver` | базовые ступени внутреннего решателя | оставить |
| `Capacity.SolverMaxIterations` | `CCapacitySolver` | итерации внутреннего решателя | оставить |
| `Capacity.ConcreteCompressionLimit` | `CCapacitySolver` | критерий бетона | оставить |
| `Capacity.SteelStrainLimit` | `CCapacitySolver` | критерий арматуры | оставить |
| `Capacity.CalculateMx` | batch | считать Mx в FullCapacity | оставить |
| `Capacity.CalculateMy` | batch | считать My в FullCapacity | оставить |
| `Capacity.CalculateMxy` | batch | считать Mxy в FullCapacity | оставить |
| `CrackWidth.Enabled` | batch, first-combination writer | включение расчета трещин | оставить |
| `CrackWidth.Allowable` | crack, batch | предельная ширина | оставить |
| `CrackWidth.CrackSpacing` | crack, batch, формульный блок | временный параметр расстояния | заменить нормативной формулой |
| `CrackWidth.StrainFactor` | crack, batch, формульный блок | временный коэффициент | заменить нормативной формулой |
| `CrackWidth.DurationFactor` | crack, batch, формульный блок | временный коэффициент длительности | заменить нормативной формулой |

Отображаемые, но не подтвержденные как читаемые настройки:

| Параметр | Наблюдение | Рекомендация |
|---|---|---|
| `Geometry.Type` | отображается на `System`, но текущий пользовательский расчет читает тип из `rngMainInput` | объединить источник истины |
| `Circle.Diameter`, `Circle.Radius`, `Circle.CenterX`, `Circle.CenterY` | отображаются на `System`, но расчет читает круг из `rngMainInput` | либо удалить дубли, либо сделать явный справочник |
| нормативные `Norms.*` | часть параметров присутствует для трассировки, но не вся читается кодом | отделить нормативную справку от рабочих настроек |
| `Concrete.Trilinear.*`, `Steel.Trilinear.*` | классы есть, но основной Excel-поток создает bilinear-материалы | оставить как будущий режим только после явного выбора диаграммы |

## 17. Самостоятельные ветви Mx/My

В `CCapacitySolver` методы `SolveMx`, `SolveMy`, `SolveMxy` являются
обертками над общим `SolveOneDirection`.

В `CBatchSectionCalculator` есть отдельные статусы и настройки включения:

- `Capacity.CalculateMx`;
- `Capacity.CalculateMy`;
- `Capacity.CalculateMxy`.

Это не отдельные решатели равновесия, но интерфейсно развивает отдельные
ветви результатов. Для этапа 3 нужно решить, сохранять ли их как виды отчета
или свести к единой постановке `N + Mx + My`.

## 18. Сохранение результатов

Сейчас после расчета сохраняются:

- сводка batch на `System!A92:M125`;
- результаты первого заполненного сочетания в `rngResultMx/My/Mxy`;
- контрольная арматура на `Расчет!A25:G34` и `System!A130:G250`;
- формульный блок трещин на `System!I130:Q250`.

Не сохраняются как единый объект/лист:

- полный снимок бетонных волокон с деформациями/напряжениями;
- полный снимок арматуры с силами и напряжениями для критического сочетания;
- `ResultDataVersion`;
- нейтральная линия и оси для повторного графика;
- данные для повторного AutoCAD-экспорта без пересчета.

## 19. Тесты

Текущие тестовые модули:

- `modTestGeometry`;
- `modTestLinearCore`;
- `modTestMaterialDiagrams`;
- `modTestSectionSolver`;
- `modTestCapacitySolver`;
- `modTestCrackWidth`;
- `modTestBatchCalculation`;
- `modTestWorkbookInterface`.

Покрытие уже есть для:

- геометрии скругленного прямоугольника и круга;
- линейного матричного ядра;
- материалов;
- Ньютонова решателя;
- поиска lambda;
- batch на 1/5/20 сочетаний;
- пропуска пустых строк;
- интерфейсного запуска и кнопок.

Недостающее покрытие по новому ТЗ:

- тест координатной инвариантности при переносе начала;
- тест повторного запуска без изменения геометрии с проверкой кэша;
- тест повторного запуска после изменения геометрии;
- тест одинакового результата автоматической и импортированной сетки;
- тест `SolveByUltimateStrain`;
- тест отсутствия AutoCAD;
- тест повторной визуализации по сохраненному листу результата.

## 20. Противоречия и риски

1. `Build-Workbook.ps1` и `Refresh-Stage09.ps1` содержат битую кириллицу в ряде строк. Это не математическая ошибка, но риск для пользовательских подписей.

2. `Validate-Workbook.ps1` использует COM Excel. При текущем аудите COM-открытие книги зависло или не завершилось. Нужна более устойчивая обертка или быстрый отказ с понятной инструкцией.

3. Нет `.git` в текущей рабочей папке, поэтому требование отдельного коммита не может быть выполнено в этой среде.

4. Текущий `CBatchSectionCalculator.ConfigureCapacity` задает часть параметров несущей способности жестко и может расходиться с `CCapacitySolver.ApplySettings`.

5. `CSystemSettingsReader.GetBoolean` содержит поддержку значения `"??"` как `True`, вероятно следствие поврежденной кириллицы. Это нужно исправить после аудита.

6. `CCrackWidthCalculator` и формульный блок трещин пока используют временные коэффициенты, не нормативную формулу СП 35.

7. Нет кэша сетки между запусками.

8. Нет единого технического листа результата.

## 21. Вывод

Текущая архитектура уже имеет важное разделение: расчетное ядро в основном
не зависит от Excel, геометрия отделена от решателя, сетка и арматура
передаются в универсальный `CSectionSolver`, AutoCAD используется через Late
Binding.

Главные следующие задачи перед математическим расширением:

1. зафиксировать регрессионную базу;
2. упорядочить настройки и убрать дубли;
3. устранить жесткие перекрытия настроек в batch-слое;
4. создать единый объект результата и технический лист снимка;
5. перевести крупные writer-операции на пакетную запись массивами;
6. только после этого добавлять второй метод несущей способности и импорт AutoCAD.
