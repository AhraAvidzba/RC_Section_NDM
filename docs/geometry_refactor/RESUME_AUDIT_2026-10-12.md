# Возобновление рефакторинга 12.10.2026

ТЗ: фактический файл `NDM_Geometry_Declarative_Refactor_TZ_CodexGoal.md`,
версия 1.4. Указанный в сообщении вариант имени с `Codex-Goal` отсутствует.
Начальный HEAD: `4b8f86af695d68c6393cfbbc57c45d776ec10992`.

## Изученное состояние

- Git: изучены последние 12 коммитов, status, полный dirty diff и staged state.
  Последовательность Circle → RectSet → RoundedRectangle → HollowRectangle →
  production input → перенос численных/геометрических fixtures уже выполнена.
- PROGRESS.md, retirement/dependency maps, evidence JSON, журналы и отчёты
  GeometryDeclarative_2026-10-11 прочитаны/инвентаризированы. Полный индекс
  816 текстовых файлов с SHA-256, итоговыми строками и отказами находится в
  `docs/regression/GeometryDeclarative_2026-10-12/Resume/EvidenceInventory.json`.
  138 JSON объявляют PASS, 34 FAIL; копии evidence packs и failed probes не
  считаются независимыми прогонами или итоговой приёмкой.
- Последние записи прежнего чата «Анализ репозитория, этап 0» содержат ошибки
  `context_length_exceeded`, а не отмену задачи или успешно завершённую приёмку.
- Начальный dirty diff сохранён в `Resume/InitialDirty.diff`. Изменения двух
  test modules соответствуют незавершённому этапу WorkbookRetirement.
  Прежние чужие удаления отчёта/рисунков и изменения execution report/formula
  images не откатываются и не включаются в локальные milestone commits.

## Реализовано, но ещё не принято окончательно

Общие paths/RoundedPolygon, snapshot/context/definition/facade, исполнитель
шести режимов армирования, четыре рецепта, единственный зарегистрированный
production-вход и перенос разбора специальных таблиц из settings reader.
Для рецептов сохранены численные и визуальные baseline/actual доказательства.
Исправления аналитических дуг отдельно описаны в BEHAVIOR_CHANGES.md.
НДС/материалы не требуется переписывать.

Локальные результаты этапов не закрывают D01–D24/E01–E14: основной output
содержит прежний VBProject, тесты modTestGeometry ещё используют заменяемые
компоненты, полное retirement и post-retirement запуск отсутствуют.

## Осталось

1. Завершить перенос modTestGeometry, удалить 12 классов, ISectionShape,
   modGeometryPreparation и заменённые helpers; подтвердить чистый VBProject.
2. Довести обязательные anchors/offset/constraints, измерения и диагностику
   общего API по G/R/A/I. Наличие частичных методов не означает полный контракт.
3. Добавить редактируемые StartSetback/EndSetback для применимых открытых
   Rounded/Hollow-путей, сохранить RectSet и явно обозначить ClosedUniform.
4. Подготовить короткий новый рецепт, E06/E09 fixtures, полный API manifest,
   Markdown и реально сохранённый последний раздел листа «Справка».
5. Выполнить итоговые геометрические, расчётные, UI/CAD и performance проверки
   на составе после retirement; обновить основную книгу и VBA_All_Code.txt,
   сохранив пользовательские данные. Составить requirements/test/final reports.

## Что повторять

После удаления: compile/source census, G01–G12/R01–R13/A01–A05/I01–I04,
матрицы четырёх семейств с прежними допусками, СП35/СП63 поэлементно,
направленные capacity/crack/stability/batch/unit-sign/status/Results тесты,
native preview/состояния/повторное открытие и CAD round-trip.
CapacitySolver/ProfileConfig прежних полных прогонов завершились TIMEOUT
180 с; они не считаются PASS. Самые долгие полные on/off и Excel UI наборы
требуют отдельного явного согласия по AGENTS.md; обычные направленные проверки
выполняются в рамках разрешённой задачи.

## Проверка незакоммиченного этапа WorkbookRetirement

Повторены 12.10.2026 в отдельной fixture-книге:

| Набор | Итог |
|---|---|
| RunRectSetIndependentSelectorTests | PASS: 1029, FAIL: 0 |
| RunAudit03InputContractTests | PASS: 52, FAIL: 0 |
| RunAudit03UnitSignStageDiagnosticTests | PASS: 43, FAIL: 0 |
| Read-only source/VBProject comparison | PASS: 155 компонентов, unexpected=0, sourceMismatch=0 |

Файлы: `docs/regression/GeometryDeclarative_2026-10-12/Resume/`.
Excel 16.0 build 10417. Основная книга не изменена. Первый sandbox COM probe
дал `CO_E_SERVER_EXEC_FAILURE`; направленный native повтор вне sandbox PASS.
Это ограничение запуска, а не регрессия расчёта. Ничьи процессы не завершались;
runner закрывает только собственный проверенный PID.
