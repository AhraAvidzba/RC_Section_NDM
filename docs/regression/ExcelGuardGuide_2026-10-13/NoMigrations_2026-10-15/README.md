# Доказательства приёмки 15.10.2026

Основной итог — Acceptance.json и отчёт
`docs/reports/NoMigrations_2026-10-15.md`. Исходный коммит записан в Before.json.
Пользовательский execution_report не менялся и не включён в коммиты.

* BaselineAudit.json — инвентарь всех 157 фактических VBA-компонентов и
  зависимости до удаления. FinalAudit.py/json — проверка удалённых символов,
  полного производственного инвентаря и границ изменения.
* DirectedCounts.json — 9987 принятых утверждений. Группы FinalSettings,
  Compatibility, FinalPlotAccepted2, AddedRecipe имеют Run.json и отдельные
  текстовые протоколы. AddedRecipe добавляет одну регистрацию только
  в частную копию. Более ранние папки сохраняют промежуточные ошибки.
* AcceptanceModels/AcceptanceNDS — повтор на конечном кандидате;
  ExactSnapshots.zip содержит побайтно равные baseline/acceptance-снимки.
  CollectChecks.py проверяет архив и сводит принятые наборы.
* AcceptanceNativeImages — итоговый повтор с четырьмя PNG, 27 проверок.
  ImportPictures показывает точные/приблизительные размеры и их отключение.
* NativeSnapshots.zip содержит BaselineNativeVBA.txt, FinalNativeVBA.txt,
  PublishedNativeVBA.txt. Это нативные экспорты, включая модули книги.
  Частные XLSM и несжатые многомегабайтные экспорты в Git не включены.
* Preservation/PublishedPreservation — данные, оформление, 29 имён и
  107483 ячейки. PresentationText.json перечисляет шесть пояснений;
  ещё одна строка находится в коротком гайде.
* PublishedSourceEquality и PublishedStandaloneHelp проверяют фактическую
  сохранённую основную книгу. Publication.json подтверждает завершённый
  readback и равенство ALL CODE нативному экспорту.
* BootstrapAndText — текущий начальный список сборщика и семь текстов,
  прочитанные Excel. PowerShellSyntax — синтаксис пяти скриптов.

Prepare.ps1, RunDirected.ps1, Probe-BootstrapAndText.ps1,
PreserveBaselinePresentation.py, Publish.ps1 и Finalize.ps1 отражают
выполненную последовательность. Подготовка и публикация защищены хешами;
они не предназначены для повторного запуска поверх изменённой книги.
RunDirected создаёт отдельную копию и закрывает только свой проверенный PID.
Для COM требуется Windows PowerShell 5.1 и доступ к установленному Excel.
Полные ON/OFF и полная UI-матрица в этой задаче не выполнялись.

Текстовые протоколы сохраняют фактический вывод, включая два конечных
пробела в пустых AUTOCAD_SNAPSHOT. Их не удаляли ради diff --check:
хеши относятся к исходным байтам протоколов.
