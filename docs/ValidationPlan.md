# ValidationPlan

Дата актуализации: 2026-09-21.

## Цель

Проверки должны подтверждать:

- компиляцию и работоспособность VBA-модулей;
- корректность структуры книги и именованных диапазонов;
- отсутствие устаревших диапазонов и настроек;
- равновесие `N + Mx + My`;
- геометрию, сетку, арматуру, аннотации и контуры;
- материалы по параметрическим TwoLine/ThreeLine-диаграммам;
- прямой НДС, несущую способность, трещины, устойчивость, batch, AutoCAD и Excel-схему.

## Основные Команды

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Build-Workbook.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Validate-Workbook.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-AllTests.ps1
```

Если завис Excel:

```powershell
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
```

## Структурная Проверка Книги

`Validate-Workbook.ps1` проверяет:

- наличие листов `Config`, `Справка`, `Расчет`, `Results`;
- обязательные именованные диапазоны;
- отсутствие дублей имен;
- отсутствие устаревших диапазонов;
- динамическую высоту `rngLoadCombinations`;
- расположение batch summary, подробных блоков прочности, трещин и устойчивости;
- наличие нижних snapshot-таблиц `Results`;
- заголовки таблиц настроек;
- отсутствие дублей ключей `Config`;
- выпадающие списки;
- допустимость формул на листах;
- отсутствие построчного runtime IO в основных writer-ах;
- что после изменений, попадающих в книгу, output-файл действительно пересобран.

## Набор Тестов

| Скрипт | Проверяемая область |
|---|---|
| `Run-GeometryTests.ps1` | `Circle`, `RectSet`, `RoundedRectangle`, `HollowRectangle`, сетка, арматура, аннотации, импорт AutoCAD Region |
| `Run-MaterialTests.ps1` | построение TwoLine/ThreeLine-диаграмм материалов по параметрам I/II ГПС |
| `Run-SectionSolverTests.ps1` | прямой `CSectionSolver`, `Newton`, `Secant`, настройки и счетчик вызовов |
| `Run-CapacityTests.ps1` | λ-траектории, `Auto`, `UltimateStrain`, `LoadMultiplier`, `Bisection`, `Brent`, `Secant` |
| `Run-CrackTests.ps1` | образование нормальной трещины, `CrackedState`, ширина раскрытия и продольные трещины |
| `Run-StabilityTests.ps1` | СП 63, СП 35, случайный эксцентриситет, `phi_l`, ветви `eta` и таблицы 7.21 |
| `Run-BatchTests.ps1` | batch по фактическим строкам `rngLoadCombinations`, worst LC, коэффициенты запаса в summary |
| `Run-WorkbookInterfaceTests.ps1` | чтение книги, кнопки, вывод результатов, схема и настройки интерфейса |
| `Run-RegressionBaselineTests.ps1` | контрольные сценарии и сравнение с baseline |

## Regression-Данные

Отчеты и baseline-файлы хранятся в `docs/regression/`. При изменении расчетной математики baseline обновляется только осознанно: в описании изменения должна быть понятная инженерная причина, а не только “изменились числа”.

## Что Требует Ручной Верификации

- окончательная нормативная трассировка всех формул СП 35 и СП 63;
- независимые ручные контрольные примеры для сложных сечений;
- сравнение с внешними расчетными программами на согласованных исходных данных;
- производительность на очень крупных импортированных сетках;
- поведение AutoCAD import/export на реальных чертежах с разными UCS и слоями.
