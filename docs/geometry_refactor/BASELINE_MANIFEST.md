# Исходное состояние

ТЗ 1.4: `B69D706E458C50D4870EFC18FAC4B8C3AB9E56B4006FF9E2DDCD67B00817F596`.
HEAD: `a138c2732382c826db9ac2f801facd3d533d406c`.

Каталог: `docs/regression/GeometryDeclarative_2026-10-11/Baseline/`.
Сохранен до изменения production-кода. Не перезаписывать.

| Артефакт | SHA-256 |
|---|---|
| RC_Section_NDM.xlsm | 769C60CEF6030A009825875B2488B6D1A7C5FC653025F7B6CBAB6B4FCE63EAED |
| VBA_All_Code.txt | B99E8C182666C01E183D116733F4EDE34B70DCF90ED0D66D457178C1B2AFC389 |
| SourceBaseline.zip | 3F8D385E24FFF42D38CB3EE1BFDAB8B021A1440DB865E9F532F8CB7BD6231FFF |
| DirtyInput.patch | C28167DED5310F1C85C9855C13CB8180C3BF724BCF74CFABCA6A326E7AF31400 |

`Manifest.json` содержит hashes отдельных исходников/tests. `NativeInventory.json`
содержит фактический VBProject, ссылки, имена, объекты листов и проверки данных.
`Config.json`/`Results.json` сохраняют значения и формулы. Excel 16.0 build 10417,
32-bit; 137 компонентов, 25 имен, 284 validation cells, четыре исправные ссылки.

`BaselineInstrumented/` содержит отдельную рабочую копию для расширенного
тестового snapshot. Ее нельзя выдавать за исходный бинарный baseline.
Изменения production-модулей относительно исходного inventory проверяются
отдельно до признания снимков эталоном.

Получены `GeometryRun1.txt`/`GeometryRun2.txt`: 20 124 записи, побитный повтор.
`GeometryRepeat.json`: PASS и три обнаруженные подмены. Получены
`CalculationsSP35_Off.txt`/`CalculationsSP63_Off.txt`: по 108 прямых НДС,
18 batch-сочетаний и полные поэлементные напряжения/деформации.
`Preview.txt` и `Preview/*.png`: девять реальных схем Excel, изображения
просмотрены. Тестовый harness заменяет заголовок geometry-only reader на
«Сгенерированная геометрия: <fixture>», не меняя production-код и схему.

`CodeComparison.json` доказывает неизменность всех нетестовых компонентов
при игнорировании только регистра VBA-идентификаторов вне строк и комментариев.
Это учитывает автоматическую капитализацию проекта VBE, не разрешая менять
литералы, операции, числа или комментарии. Исходный baseline-файл не изменялся.

Отвергнутые первые снимки имеют префикс `RejectedAllSourceImport` и не участвуют
в приемке. Позднее instrumentation ограничена одним тестовым компонентом.
Полная матрица 14.2, остальные режимы схемы, CAD и производительность еще
NOT_RUN; девять fixtures не заменяют их.
