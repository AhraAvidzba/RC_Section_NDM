# Компактная адресная приёмка Snapshot IO

Основная чистая книга: SHA256 `4B12C3C9FF65136B76EBE2F4DD690CC666914AFED68F3B1606583B488BE6013E`.
Исходная опубликованная книга: `91B5C5C6DA655E14B699CBB5C15ADB7CF6BA94555DC0890243A956E79637EF12`, коммит `5be030de3b0c2df5b65530e23fbd3d2ae6c94231`.

Это адресная приёмка, не общий полный ON/OFF. Наблюдение ndm — PAUSED.
Частные книги, пользовательские входные матрицы и полные двоичные снимки
не публикуются. Журналы здесь — явно отмеченные выдержки с SHA256 полного
локального журнала; исходные evidence сохранены в родительском каталоге.

- `DirectedAcceptance.json`: 24 запуска ревизии 08 и восемь затронутых
  повторов ревизии 09; сумма assertions включает повторы.
- `FinalScope.json`: между этими ревизиями изменены только Store,
  modCalculationSnapshot и тесты снимка; остальные 175 компонентов равны.
- `SourceAndStructure.json`, `SourceManifest.json`, `ManualNative.json`:
  настоящий экспорт Excel сопоставлен со всеми исходниками; временного
  профайлера в выдаваемых книгах нет.
- `Final*Acceptance08.json`: пять наблюдений на каждый размер и хранилище,
  равенство результатов/схемы, ноль автоматических десериализаций, реальная
  ручная загрузка и освобождение контекстов.
- `AlternatingExcel08.json`: прямой и обратный A/B; для одного LC десять
  наблюдений версии, для 30 LC — пять. В Raw-файлах, указанных хешами,
  роли A/B обратной серии противоположны прямой; сводка это учитывает.
- `*FullPayload08.json`: независимый разбор всех 16 полных матриц, включая
  все строки состояний. Отличаются только RunID/время разных запусков;
  для остальных чисел эквивалентен только знак точного нуля.
- `CodecAcceptance07.json`, `PackedDictionaryAcceptance.json`: точный
  roundtrip и побайтное равенство шести синтетических файлов. Скорости
  кодека не выдаются за время инженерного расчёта. DirectUTF16 — отвергнутый
  промежуточный вариант, его результат сохранён отдельно.
- `RealAcceptance07.json`, `StandardAcceptance07.json`: отдельный эксперимент
  первого кандидата Memory/Reload. Автоматический Memory существовал уже
  в исходной книге; принудительное повторное чтение — контрольный сценарий.
- `BaselineControlFailure.json`: намеренный отказ старого кода после очистки
  Results в частной книге. Окончательная книга тот же контроль проходит.
- `RuntimeFormatting.json`, `ResultsFormatting08.json`,
  `SettingsHelpStyle.json`: независимые проверки оформления и навигации.

Повторять только в новом evidence-каталоге. Основной шаблон адресного VBA:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/build_workbook/Run-Audit03Watchdog.ps1 -SourceWorkbook workbook/output/RC_Section_NDM.xlsm -ReportPath docs/regression/NewSnapshotCheck/Snapshot_No.txt -Macro modTestCalculationSnapshot.RunCalculationSnapshotTests -Mode No -TimeoutSeconds 300
```

Повторяемые инструменты фаз и кодека создают собственные книги, исходную
не импортируют и не сохраняют. По умолчанию пять наблюдений после прогрева:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/build_workbook/Measure-SnapshotIOPhases.ps1 -SourceWorkbook workbook/output/RC_Section_NDM.xlsm -EvidenceDirectory docs/regression/NewSnapshotPhases
powershell -NoProfile -ExecutionPolicy Bypass -File tools/build_workbook/Measure-SnapshotCodec.ps1 -SourceWorkbook workbook/output/RC_Section_NDM.xlsm -EvidenceDirectory docs/regression/NewSnapshotCodec
```

Реальная фикстура 1232+8/30 остаётся локальной; её SHA256 и происхождение
есть в `ManualTestBook.json`. Тестовая книга выдана пользователю для ручной
проверки. Согласованное следующее изменение — расширение `.ndm`; выданный
выпуск с `.ndms` после выдачи не изменяется.
