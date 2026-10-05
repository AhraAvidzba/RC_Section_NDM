# Audit03: Диагностика Общего UltimateStrain

Контрольная точка: `4d89bac8`. Основная книга и пользовательские данные
не изменялись. Использованы существующие test class и test module;
новые классы, физические критерии и допуски не вводились.

## Negative Gate

Frozen production v313 с усиленными тестами в изолированной v314:
`674/82`, exit=1, source unchanged=True, Results/style save-reopen=True.
Проверены два kind одного общего Search, 21 сценарий каждого kind, включая
ошибки подготовки старта, оценки Якобиана и повторный успешный запуск
на том же объекте. Отсутствующий context проверяется как контролируемая
ошибка API, без фиктивного результата.

82 неуспешных assertions группируются так:

- 20 rawRuntimeCode и 20 rawRuntimeStage: helper не сохранял структурированную
  исходную диагностику 20 runtime-попыток;
- 10 rawStandardReason и 10 rawCallbackReason: отсутствовала исходная причина
  в отдельном diagnostic-поле;
- 10 localizedReason: стандартные VBA-сообщения попадали в последний SetFailure;
- по четыре typedCause/status/code: деление на ноль ошибочно относилось
  к internal failure вместо numerical failure.

Все recovery assertions прошли. Тест наблюдает последний SetFailure отдельно
от общего DiagnosticLog, чтобы fake-result не скрывал потерю причины.

## Исправление И Границы

Общий UltimateStrain error-boundary сохраняет номер, stage и original
Description в DiagnosticLog. Коды 6/11 дают численную причину, другие
неожиданные ошибки дают internal; ранее установленный terminal outcome
не заменяется вторичным исключением. Стандартная причина объясняется
по-русски, явно заданная причина callback-а сохраняется без изменения.
Сравнение Description используется только для представления, не для статуса.

В рамках A03/A05 удалены AggregateMeta/AggregateExternalStatus и private Merge,
не используемые production. Три assertions сохранили IDs/expected values
и используют WorstResultMeta -> ExternalStatus. Добавлены проверки заголовка
и объединения блока ширины трещин; само оформление writer-а не изменялось.

Positive gate v315 завершен `756/0`: те же scenarios, source unchanged=True,
Results/style save-reopen=True. Result/policy/writer gate `3641/0` также
проверил заголовок/merge, три прежних aggregate assertions и заполненные
Results после сохранения/открытия. Build SHA-256:
`C7A4A7310A5530CD4128351FFEAE5955545BA6722130CB23660F7C0F83462061`.
Negative-книга SHA-256:
`D483C3B7D11A30C2448AD598CE31D3CCFF5A1D3E008960AAE397E7A11A2196FA`.

Actual source/VBE подтвержден на следующей v317: 108/108, failed=0,
113 компонентов. Capacity suite v317 `3562/0`; все 601 прежнее численное
значение полного Off v313 совпало точно, missing=0, differences=0.
Results/style save-reopen=True, source unchanged=True. После build v315
усилен существующий Stability test и добавлен подтвержденный входной guard;
v315 не объявляется точным снимком этих более поздних изменений.
Остальные финальные gates еще не завершены. Audit03 не завершен.
