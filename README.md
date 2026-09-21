# RC Section NDM

Excel VBA-программа для расчета железобетонных поперечных сечений по нелинейной деформационной модели.

Расчетная постановка в ядре одна:

```text
N + Mx + My
```

Одноосный изгиб считается частным случаем общей постановки при `Mx = 0` или `My = 0`.

## Что реализовано

- книга Excel `workbook/output/RC_Section_NDM.xlsm` с листами `Config`, `Справка`, `Расчет`, `Results`;
- централизованное чтение настроек, единиц и знаков через `CSystemSettingsReader` и `CUnitSystem`;
- динамическое число сочетаний по строкам `rngLoadCombinations`, шаблон по умолчанию содержит 30 строк;
- расчетные профили, где отдельно включаются НДС по модели прочности, несущая способность, трещины и устойчивость;
- геометрии `Circle`, `RectSet`, `RoundedRectangle`, `HollowRectangle`;
- импорт расчетной сетки из AutoCAD `Region` и экспорт результатов обратно в AutoCAD;
- прямой расчет НДС, поиск несущей способности и расчет ширины раскрытия нормальных и продольных трещин;
- расчет продольного изгиба и устойчивости по СП 63 и СП 35;
- Excel-схема сечения по последнему расчетному снимку с листа `Results`;
- человекочитаемый отчет последнего расчета при включенной настройке `General.ExecutionReportEnabled`;
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

- `AGENTS.md` - правила разработки, границы проекта и обязательные проверки.
- `docs/PROJECT_STATUS.md` - текущий статус функциональности.
- `docs/Architecture.md` - фактическая архитектура проекта.
- `docs/MathematicalModel.md` - расчетная постановка и численные алгоритмы.
- `docs/CoordinateSystem.md` - оси, знаки и единицы.
- `docs/WorkbookLayout.md` - структура книги и именованные диапазоны.
- `docs/SETTINGS_REFERENCE.md` - справочник настроек `Config`.
- `docs/ValidationPlan.md` и `docs/TEST_CASES.md` - план проверок и регрессионные сценарии.
- `docs/NormativeTraceability.md` и `docs/OpenNormativeQuestions.md` - нормативная трассировка и открытые вопросы.
- `docs/UserGuideCircle.md`, `docs/UserGuideRectSet.md` - пользовательские сценарии для параметрических сечений.
- `SP63_SP35_SECOND_ORDER_STABILITY_DISCRETE_AUDITED.md` - текущая методика расчета устойчивости.

## Ограничения

Проект поддерживает обычную ненапрягаемую продольную арматуру. Нормативные формулы и коэффициенты по СП 35 и СП 63 должны оставаться трассируемыми по локальным документам `norms/` и по листу `Справка`; неоднозначные вопросы фиксируются в `docs/OpenNormativeQuestions.md`.
