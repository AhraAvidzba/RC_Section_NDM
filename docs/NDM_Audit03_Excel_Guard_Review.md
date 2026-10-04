# Audit03: Восстановление Excel.Application

Дата: 2026-10-04. Исходный checkpoint: `53761e9b`.

## Подтвержденный Дефект

На собственном пустом Excel.Application установка Calculation=Manual
возвращает реальную COM-ошибку. Старый Enter к этому моменту уже выключал
ScreenUpdating, EnableEvents и DisplayAlerts, но еще не устанавливал mActive.
Поэтому последующий Restore не возвращал ранее измененные настройки.

Negative v224: **120/19**. Проверены восемь исходных сочетаний трех Boolean
при неуспешном Enter и 24 успешных сочетания (три Calculation x восемь
Boolean). Исходная ошибка сохранялась, но немедленное восстановление и
последующий Restore не проходили для измененных настроек. Успешные варианты
прошли. Это 32 варианта интерфейсного контракта, не 139 инженерных задач.

## Исправление И Проверки

- Enter включает признак необходимости восстановления после полного снимка
  и до первого setter-а. При отказе сразу вызывает Restore и повторно
  передает исходные Err.Number/Source/Description.
- Ошибка чтения снимка не объявляет настройки измененными. Публичный Restore
  и деструктор сохраняют прежнюю терпимость к недоступному Excel.
- Fake-классы, новые классы и обращение к пользовательскому экземпляру
  Excel отсутствуют. Тест создает отдельный Application и проверяет Hwnd.
- Directed v225: **139/0**, немедленный rollback, сохранность ошибки,
  повторный Enter, idempotent Restore, деструктор и noSolve.
- Соседний presentation/metadata/AutoUpdate/arc набор: **2355/0**.
- Read-only структура: **27/27**. Palette/save-reopen: **355/0**.
- Results before/after SHA:
  `B30FFE6304424CD524A885D0798A9AA2931C2E6405558BCD5D6A3554460DF9DB`.
  Status styles before/after SHA:
  `1EABA40A72D5A28392F11FC5D38F0D2B1E9B264FDFAD9F5514ACB9EE1A1E1132`.
- Frozen RC_Section_NDM_excel_guard_positive_v225.xlsm SHA:
  `2FB1A9D7399C99FA75CB297703389F22856E5A086F8D844065128309ECC66091`.
- Actual read-only VBE export: **111** компонентов,
  VBA_All_Code_excel_guard_v225_2026-10-04.txt SHA:
  `9B0EA15BE930E5A43CD28EF29DD0CEE45600693C6A06C86F79FD0DEA5C83BD60`.
  Canonical export содержит те же фактические байты.
- Source/export equality: **106/106**, failed=0; состав классов сохранен.
- Census: **106** модулей/**4421** методов/**1721** guard candidates;
  это индекс ревизии, semantic acceptance остается Pending.

Завершенные watchdog source unchanged=True. Пользовательский Excel PID 23476
не использовался. Main output, ТЗ, Solver/Search, expected и tolerance
не менялись. Help/Config formatting не изменялись и не объявляются заново
проверенными данным срезом: их последний адресный gate остается v223.
Полный numerical Off/On остается историческим v206; окончательный release
gate обязателен. Audit03 в целом не завершен.
