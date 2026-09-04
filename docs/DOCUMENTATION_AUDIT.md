# DOCUMENTATION_AUDIT

Дата ревизии: 2026-08-05.

## Цель

Сверить документацию с фактическим состоянием VBA-кода, книги Excel, сборочных скриптов и тестов.

## Исправленные несоответствия

| Несоответствие | Исправление |
|---|---|
| Документы описывали отдельные пользовательские ветви `Mx`, `My`, `Mxy` | Зафиксирована только общая постановка `N + Mx + My` |
| Упоминались старые диапазоны `rngMainInput`, `rngRebarInput`, `rngResultMx`, `rngResultMy`, `rngResultMxy`, `rngSystemDiagnostics` как рабочие | Они оставлены только в списках удаленных/запрещенных сущностей |
| `ImplementationPlan.md` и `IMPLEMENTATION_PLAN.md` дублировали друг друга | Оставлен один актуальный `docs/IMPLEMENTATION_PLAN.md` |
| Stage-отчеты конфликтовали с текущей архитектурой | Удалены, полезная сводка перенесена в `docs/PROJECT_STATUS.md` |
| Документация говорила, что Brent/Secant для capacity не реализованы | Обновлено: `Bisection`, `Brent`, `Secant` реализованы для `LoadMultiplier` |
| Упоминались типы диаграмм как расчетные режимы | Обновлено: материалы задаются только точками диаграмм |
| Пользовательская инструкция требовала ввод координат арматуры | Обновлено: арматура круга расставляется автоматически по настройкам `Config` |
| Структура `Results` не отражала `rngBatchSummary` | Добавлено описание `Results!A1:CU29` |

## Живой набор документации

- `README.md`;
- `AGENTS.md`;
- `docs/NDM_PROJECT_SPEC.md`;
- `docs/PROJECT_STATUS.md`;
- `docs/Architecture.md`;
- `docs/MathematicalModel.md`;
- `docs/CoordinateSystem.md`;
- `docs/WorkbookLayout.md`;
- `docs/SETTINGS_REFERENCE.md`;
- `docs/ValidationPlan.md`;
- `docs/TEST_CASES.md`;
- `docs/UserGuideCircle.md`;
- `docs/NormativeTraceability.md`;
- `docs/OpenNormativeQuestions.md`;
- `docs/OpenQuestions.md`;
- `docs/IMPLEMENTATION_PLAN.md`.

## Удаленные устаревшие документы

- `docs/CURRENT_ARCHITECTURE_AUDIT.md`;
- `docs/ExistingCodeAnalysis.md`;
- `docs/ImplementationPlan.md`;
- `docs/Stage01Report.md`;
- `docs/Stage02Report.md`;
- `docs/Stage03Report.md`;
- `docs/Stage04Report.md`;
- `docs/Stage04SolverReport.md`;
- `docs/Stage05Report.md`;
- `docs/Stage06Report.md`;
- `docs/Stage07Report.md`;
- `docs/Stage08Report.md`;
- `docs/Stage09Report.md`.

## Не трогалось

Каталоги `reference/`, `norms/`, `control_examples/` не изменялись.
