# ValidationPlan

Дата актуализации: 2026-08-05.

## Цель

Проверки должны подтверждать:

- компиляцию и работоспособность VBA-модулей;
- корректность структуры книги;
- отсутствие устаревших именованных диапазонов и настроек;
- равновесие `N + Mx + My`;
- геометрию, сетку и арматуру;
- материалы по параметрическим TwoLine/ThreeLine-диаграммам;
- работу прямого решателя, capacity, трещин, batch и UI.

## Основные команды

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Build-Workbook.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Validate-Workbook.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-AllTests.ps1
```

Если завис Excel:

```powershell
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
```

## Структурная проверка книги

`Validate-Workbook.ps1` проверяет:

- наличие листов `Расчет` и `Config`;
- обязательные именованные диапазоны;
- отсутствие дублей имен;
- отсутствие устаревших диапазонов;
- область печати;
- вертикальные разрывы;
- расположение таблицы нагрузок и блока результата;
- заголовки таблицы настроек;
- отсутствие дублей ключей `Config`;
- выпадающие списки;
- новые диапазоны параметров материалов, режимов диаграмм и контрольных таблиц точек;
- заголовки `rngLoadCombinations`;
- допустимость формул на листах;
- отсутствие построчного runtime IO в основных writer-ах.

## Набор тестов

| Скрипт | Проверяемая область |
|---|---|
| `Run-GeometryTests.ps1` | геометрия, сетка, круг, скругленный прямоугольник, Г-сечение, автоматическая арматура, импорт AutoCAD Region в `CSectionModel` на мок-данных |
| `Run-MaterialTests.ps1` | построение TwoLine/ThreeLine-диаграмм материалов по параметрам I/II ГПС |
| `Run-SectionSolverTests.ps1` | прямой `CSectionSolver`, `Newton`, `Secant`, настройки |
| `Run-CapacityTests.ps1` | универсальная λ-траектория `CapacityLoadPath`, `Auto`, `UltimateStrain`, `LoadMultiplier`, `Bisection`, `Brent`, `Secant` |
| `Run-CrackTests.ps1` | ширина раскрытия уже образовавшихся трещин |
| `Run-BatchTests.ps1` | batch до 20 сочетаний, worst LC, summary |
| `Run-WorkbookInterfaceTests.ps1` | кнопки, чтение книги, вывод результатов |
| `Run-RegressionBaselineTests.ps1` | базовые сценарии `Stage01 TEMPORARY_BASELINE` |

## Последний зафиксированный полный прогон

Файл:

```text
docs/regression/Stage01_AllTests_Report.txt
```

Сводка последнего полного прогона:

| Блок | Passed | Failed |
|---|---:|---:|
| Geometry | 56 | 0 |
| Linear | 49 | 0 |
| Material | 20 | 0 |
| Section solver | 175 | 0 |
| Capacity | 167 | 0 |
| Crack | 34 | 0 |
| Batch | 31 | 0 |
| Workbook UI | 34 | 0 |
| Regression baseline | 39 | 0 |

## Что еще требует верификации

- нормативные формулы диаграмм по СП;
- нормативный расчет ширины раскрытия трещин;
- независимые ручные контрольные примеры;
- сравнение с внешними расчетными программами на согласованных исходных данных;
- производительность на крупных и сложных сечениях;
- поведение `UltimateStrain` на несимметричных будущих геометриях.
