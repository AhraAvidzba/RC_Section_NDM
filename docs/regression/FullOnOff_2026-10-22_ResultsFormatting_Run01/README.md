# Полная приёмка оформления Results, 22.10.2026

Итог: INCOMPLETE_OR_FAILED. Все 156 основных наборов завершены, 150 прошли,
6 с ошибками. По 65 FAIL в OFF и ON, одинаковые строки. Обязательный целевой
план выполнен: 122 случая, batch-15, устойчивость, сталь, настоящий AutoCAD,
независимый контроль прошли; OutputPresentation нашёл отсутствие правой рамки
у General.ReformatResultsEnabled. Общий PASS не создавался.

Отчёт: ../../reports/FullOnOff_ResultsFormatting_2026-10-22.md.
Главные доказательства: Run.json, Summary.json, FinalAudit.json,
Finalization.json, Finalizer.log и Focused/Acceptance.json.

RawEvidence.zip содержит полные исходные журналы, снимки выполненных сценариев,
промежуточные JSON, доказательства настоящего AutoCAD и независимого контроля.
EvidenceFiles.json хранит SHA256 архива и каждого вложенного файла.
Для повторного анализа журналов распаковать архив в этот каталог. Большие
проверочные книги и чертежи в Git не включены; исходная книга восстанавливается
из 113ac7a0:workbook/output/RC_Section_NDM.xlsm и сверяется с SHA256 в Run.json.
Проверка защищённых файлов целевого плана использует исходные абсолютные пути;
не подменять отсутствующий файл другим ради положительного результата.

Использованные верификаторы:
- Summarize-FullOnOffAcceptance.py --directory <этот каталог> --check: exit=1.
- Verify-FocusedAcceptance.py --directory <этот каталог>/Focused --check: exit=1.
- Finalize-CompleteAcceptance.py --core-directory <этот каталог>
  --focused-directory <этот каталог>/Focused: exit=1, отказ создавать PASS.

Это завершённый прогон с неуспешными проверками, а не пропущенный набор.
Исходники, книга и исполняемые сценарии не исправлялись во время прогона.
Полный набор автоматически не повторялся. Наблюдение ndm отключено.
