# Проверки Численных Fixtures

Архив: `docs/regression/GeometryDeclarative_2026-10-11/RetirementFixturesEvidence.zip`.
Хеши исходников, книг и отчетов: `RETIREMENT_FIXTURES_EVIDENCE.json`.
Это промежуточный retirement, не окончательная приемка.

В `RetirementBaseline` сохранены принятая ProductionRecipe-книга с единственной
добавленной snapshot-инструментацией BatchCalculation и соответствующий source.
В `RetirementFixtures` находятся перенесенные fixtures; fake class удален.
Основная пользовательская книга не менялась.

| Проверка | Результат |
|---|---|
| Native compile/source verification | PASS, 11 source bodies, unexpected=0, sourceMismatch=0 |
| Query / ownership / mesh | 1319 / 520 / 34333 PASS |
| НДС / целевые Capacity / CrackWidth | 2834 / 74 / 2134 PASS |
| Batch counters/writer | 14 PASS |
| Numerical baseline | 39 PASS; 11 выводимых записей одинаковы, исключено только время |
| Матрица 12 fixtures | 19632 записи, 0 расхождений с прежними допусками |
| Намеренные подмены | 3 из 3 обнаружены |
| CapacitySolver/ProfileConfig широкие suites | TIMEOUT 180 с, не PASS |

Побитный diff остается FAIL для последних разрядов RoundedTapered; точные
метрики хранятся в LoadMatrixAnalysis.json, исходные 1644 differences не скрыты.
Отдельно сохранены failed probes компиляции, старого каталога ключей и строгой
сверки VBE-написания/Unicode comments. NumericSpellingProof подтверждает
неизменность Double, а не только близость десятичных значений.

Повторяемые команды:

```powershell
python tools/build_workbook/Analyze-RetirementFixtures.py `
  docs/regression/GeometryDeclarative_2026-10-11/RetirementBaseline/LoadMatrixSnapshot.txt `
  docs/regression/GeometryDeclarative_2026-10-11/RetirementFixtures/DumpLoadMatrixGeometrySnapshot.txt `
  --report docs/regression/GeometryDeclarative_2026-10-11/RetirementFixtures/LoadMatrixAnalysis.json

powershell -NoProfile -File tools/build_workbook/Run-Audit03Watchdog.ps1 `
  -SourceWorkbook docs/regression/GeometryDeclarative_2026-10-11/RetirementFixtures/RC_Section_NDM.xlsm `
  -ReportPath docs/regression/GeometryDeclarative_2026-10-11/RetirementFixtures/OwnershipRepeat.txt `
  -Macro RunGeneratedContourOwnershipTests -Mode No -TimeoutSeconds 180
```

Native verification использует принятую ProductionRecipe-книгу как baseline,
11 approved modules и `RetiredModules CFakeMultiOpeningGeometry`. Скрипт не
принимает неоговоренное удаление или retired source, который еще существует.
Проверка последнего запрета выполнена без создания Excel, хеш книги не изменился.
Полный старый on/off набор не запускался.
