# Новый полный GeometryFinal запуск — подготовка без старта

Сценарии подготовлены, реальная книга/HEAD/SHA пока не зафиксированы. Launch.ps1 следует запускать только основному агенту после успешных направленных/native gates, публикации чистой полной пересборки без переноса пользовательских данных и коммита. Запуск автоматически снимает текущие published book SHA, AllCode SHA, HEAD/branch и хэши всех участвующих входов, включая TestObserveCompleted.py. При незакоммиченных входах, существующем request/launch, evidence или собственной задаче повтор запрещён.

Состав — 178 core suites CoverageVersion6 в12 группах No/Yes и весь FocusedValidationPlan:122 одиночных, batch15 обоих режимов, устойчивость, асимметричная сталь, настоящий CAD, reopen и независимый контроль. Forecast4.5ч, deadline8ч от фактического запуска. Число assertions не заменяет этот состав.

```powershell
& 'C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe' -NoProfile -ExecutionPolicy Bypass -File 'docs/regression/FullSuiteLaunch_2026-10-26_GeometryFinal_Run01/Launch.ps1'
```

Эта команда НЕ выполнялась при подготовке. Она регистрирует ровно одну on-demand задачу RC_NDM_Full_20261026_GeometryFinal_Run01 без повторяющегося или календарного триггера. Task Scheduler запускает независимый host; host удерживает .Handle дочернего driver и сохраняет точные PID+UTC ticks. Новую задачу запускать повторно или создавать дубликат запрещено. Старые задачи и пользовательские Excel/CAD не изменяются. Source fixed path — актуальная опубликованная чистая книга, а не UserBook_BeforeChanges и не пакет прежнего кандидата.

После фактического запуска при каждой плановой проверке:

```powershell
& 'C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe' -NoProfile -ExecutionPolicy Bypass -File 'docs/regression/FullSuiteLaunch_2026-10-26_GeometryFinal_Run01/Read-CurrentStatus.ps1' -CheckAllHashes
& 'C:/Users/avidzba/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe' 'docs/regression/FullSuiteLaunch_2026-10-26_GeometryFinal_Run01/TestObserveCompleted.py'
```

Наблюдатель читает только текущие LaunchRequest/Run/primary. Ошибки всех завершённых core и Focused работ сохраняются после смены этапа, а изменение закрытого primary не стирает прежний FAIL. Результат наблюдения всегда OBSERVATION_NOT_ACCEPTANCE; независимые ожидаемые численные отказы не классифицируются автоматически по полю Failures. Текущие primary/stdout/stderr/Progress и все уже завершённые primary требуют чтения основным агентом. Exit0/Completion.json сами не дают полный PASS. Native PID проверяются вместе со startTicks; старый RUNNING при отсутствующем процессе не считать живым этапом.

ObserverSyntheticSelftest.json содержит только6 синтетических негативных/частичных сценариев, никакие старые результаты не импортированы в PASS. CoveragePreparedProof.json — статический подсчёт плана; PS51ParseProof.json — синтаксис без запуска. LaunchRequest.json/FullLaunch.json/core evidence при подготовке не создаются.

После фактического завершения, сохранения доказательств и проверки отсутствия всех own процессов основной агент удаляет только точную собственную задачу. Повторный долгий полный/целый UI набор требует отдельного согласия; ndm использовать существующий без дубликатов и менять статус только по действующему правилу пользователя.
