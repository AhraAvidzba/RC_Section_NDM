# Приёмка ResolveAnchor

Итог: Acceptance.json, 1792 адресные проверки, девять неизменённых моделей
и 12 неизменённых контрольных НДС. Полные ON/OFF не запускались.
PositionOrder сохранён: тест доказывает влияние на физического владельца
при KeepPublishedOwner и публикации по исходной позиции.

Prepare.ps1 сохраняет фактический baseline и проверяет исходный код.
RunDirected.ps1 выполняет квалифицированные макросы только в собственной
копии. Каждый Excel PID и время создания проверены; другие книги не
закрываются. ControlStates.bas добавляется только в тестовую копию.

Окончательная команда адресного прогона:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File docs/regression/ExcelGuardGuide_2026-10-13/ResolveAnchor_2026-10-15/RunDirected.ps1 -SourceWorkbook docs/regression/ExcelGuardGuide_2026-10-13/ResolveAnchor_2026-10-15/Candidate.xlsm -Stage NewDirectedRun -Macros modTestRecipeOperations.RunResolveAnchorTests,modTestRecipeOperations.RunRecipeOperationTests,modTestRecipeOperations.RunRecipeDiagnosticTests,modTestSectionRecipes.RunRecipeRuleTests,modTestPolygonSpec.RunPolygonSpecTests,modTestAudit02Api.RunAudit02GuideExamplesTests
```

Частные XLSM остаются локальными. ExactSnapshots.zip содержит обе пары
точных снимков; NativeSnapshots.zip — исходный, принятый и фактический
публикуемый экспорт VBA. Для независимой проверки распакуйте native-снимки
сюда и запустите VerifyPreservation.py с параметрами --book/--dump/--report.
Acceptance.py проверяет итоговые receipts, архив точных снимков, ALL CODE
и сохранность пользовательского execution_report.

PreserveBaselinePresentation.py возвращает исходные XML и стили всех
частей книги, оставляя новый VBA и один изменённый абзац A560 с его
высотой. Publish.ps1 проверяет принятого кандидата и исходные hash,
заменяет пару книга/ALL CODE с возможностью восстановления при отказе.
Finalize.ps1 завершает только после независимого чтения основной книги.

Directed02/Run.json фиксирует отказ из-за неверного имени дополнительного
макроса; итоговым является DirectedFinal. Подробности — в отчёте.
