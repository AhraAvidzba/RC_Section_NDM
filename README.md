# RC Section NDM

Excel VBA-программа для расчета железобетонных поперечных сечений по нелинейной деформационной модели.

Текущее состояние проекта: рабочий прототип расчетного ядра и книги `workbook/output/RC_Section_NDM.xlsm`.
Основная постановка расчета единая:

```text
N + Mx + My
```

Одноосный изгиб рассматривается только как частный случай общей постановки при `Mx = 0` или `My = 0`.

## Что реализовано

- книга Excel с листами `Расчет` и `System`;
- ввод настроек только на листе `System`;
- таблица сочетаний нагрузок до 20 строк на листе `Расчет`;
- геометрии `Circle` и `RoundedRectangle`;
- автоматическая расстановка продольной ненапрягаемой арматуры для круглого сечения;
- волоконная бетонная сетка с опциональным дроблением граничных ячеек;
- общий нелинейный решатель `CSectionSolver` для `N + Mx + My`;
- поиск несущей способности `CCapacitySolver` методами `LoadMultiplier` и `UltimateStrain`;
- одномерный поиск для `LoadMultiplier`: `Bisection`, `Brent`, `Secant`;
- расчет ширины раскрытия уже образовавшихся нормальных трещин;
- пакетный расчет до 20 сочетаний;
- экспорт расчетной сетки, арматуры и напряжений в текущее окно AutoCAD;
- импорт расчетной сетки из AutoCAD `Region` на заданных слоях;
- автоматические VBA-тесты и PowerShell-сборка книги.

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

## Документация

- `AGENTS.md` - правила разработки и ограничения проекта.
- `docs/PROJECT_STATUS.md` - текущий статус функциональности.
- `docs/Architecture.md` - фактическая архитектура.
- `docs/MathematicalModel.md` - расчетная постановка и алгоритмы.
- `docs/CoordinateSystem.md` - оси, знаки и единицы.
- `docs/WorkbookLayout.md` - структура книги и именованные диапазоны.
- `docs/SETTINGS_REFERENCE.md` - настройки листа `System`.
- `docs/ValidationPlan.md` и `docs/TEST_CASES.md` - проверки и регрессия.
- `docs/NormativeTraceability.md` и `docs/OpenNormativeQuestions.md` - нормативная трассировка.
- `docs/UserGuideCircle.md` - пользовательский сценарий для круглого сечения.

## Ограничения

Проект поддерживает только обычную ненапрягаемую арматуру. Нормативная верификация по СП 35 и СП 63 не завершена: часть параметров материалов и трещин остается пользовательской и должна проверяться инженером.
