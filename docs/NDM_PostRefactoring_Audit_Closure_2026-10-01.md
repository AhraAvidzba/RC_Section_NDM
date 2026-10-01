# Закрытие post-refactoring audit

Основание: `docs/NDM_PostRefactoring_Audit_2026-09-30.md`.

Дата финальной проверки: 2026-10-01.

## Что исправлено в рамках закрытия audit

- Устранено regression-падение capacity-тестов после перевода расчетов на общий `CStateSolutionRunner`.
- Причина была не в математике несущей способности, а в слишком узкой типизации материалов в runner-е: часть старых capacity-тестов использует линейные тестовые материалы, а не `CMaterialDiagram`.
- `CStateSolutionRunner` теперь принимает материалы как общий объектный материал solver-а, но специализированные active-set и extension-подсказки включает только для штатных `CMaterialDiagram`.
- Для тестовых линейных материалов сохраняется прежний путь через обычный `CSectionSolver.Solve`, без искусственной проверки границ диаграммы, которых у таких материалов нет.
- Низкоуровневая расчетная математика, критерии Capacity/CrackFormation и правила статусов не менялись.

## Self-audit по обязательным пунктам

| Пункт audit | Итог проверки |
| --- | --- |
| T01 Secant не должен возвращать непроверенную нижнюю границу | Подтверждено тестом `TestLimitSearchSecantFinalizesCheckedRoot`. |
| T02 Bisection не должен считать исчерпание итераций успехом | Подтверждено тестом `TestLimitSearchBisectionIterationLimitFails`. |
| T03 численная ошибка при `lambda = 0` не должна превращаться в `BaseFail` | Подтверждено тестом `TestInitialLambdaFailureStatusMapping`. |
| T04 техническая граница search не должна смешиваться с физическим отсутствием трещины | Подтверждено crack/search tests и `ResultCode = SEARCH_BOUND_REACHED`. |
| T05 `CRITERION_NOT_REACHED` не должен создавать фиктивный PostCrackState | Подтверждено тестом `TestCrackFormationNoCrackDoesNotBuildPostState`. |
| T06 cache-hit без `LastRunner` не должен давать ложный `NumFail` | Подтверждено тестом `TestCrackFormationCacheHitWithoutLastRunner`. |
| T07 repository не должен возвращать неуспешный solve как reusable state | Подтверждено тестом `TestStateRepositoryReusesOnlyConvergedStates`; статическая проверка `CStateRepository` показывает reuse только для `Converged`. |
| T08 подробный блок трещин не должен показывать OK при ошибке обязательного состояния | Подтверждено тестом `TestCrackAggregateIncludesCurrentStateFailure`. |
| A01/A03 общая search-логика Capacity/CrackFormation | Подтверждено наличием общего `CLimitSearch*` и тестового generic problem `CTestLimitSearchProblem`. |
| A02 Capacity probes через общий state-solve pipeline | Подтверждено маршрутом через `CStateSolutionRunner`; прямые specialized-вызовы оставлены только там, где нужны начальные подсказки/низкоуровневая проба. |
| A04 самостоятельный typed result search | Подтверждено `CLimitSearchResult` с `CResultMeta`, `InternalStatus`, `ResultCode`, `ResultKind`. |
| A05 writer-ы не должны собирать инженерские статусы сами | Подтверждено: writer-ы читают `ResultMeta.ResultComment` и внешний статус из policy. |
| A06 централизованные статусы и комментарии | Подтверждено `CResultStatusPolicy` и regression-тестами `TestCombinationResultTreeDrivesDisplayFields`. |

## Финальные проверки

Сборка книги:

```text
tools/build_workbook/Build-Workbook.ps1
OK
```

Полный regression-набор на финальной сборке:

```text
tools/build_workbook/Run-AllTests.ps1
TOTAL: passed=496; failed=0
TOTAL_MATERIAL: passed=54; failed=0
TOTAL_SECTION_SOLVER: passed=390; failed=0
TOTAL_CAPACITY: passed=934; failed=0
TOTAL_CRACK: passed=328; failed=0
TOTAL_BATCH: passed=659; failed=0
TOTAL_WORKBOOK_UI: passed=308; failed=0
TOTAL_REGRESSION_BASELINE: passed=39; failed=0
```

Отчеты:

- `docs/regression/Stage01_AllTests_Report.txt`
- `docs/regression/Stage01_RegressionBaseline_Raw.txt`

## Итоговый вывод

Текущий audit закрыт. Подтвержденные пункты T01–T08 и архитектурные условия A01–A06 проверены тестами и статическим self-audit кода.

Дополнительные архитектурные рефакторинги сверх текущего audit не начинались.
