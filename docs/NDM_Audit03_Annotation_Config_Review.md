# Audit03: Ячейки Настроек Подписей И Размеров

## Scope

Срез после `0258a89`: только тесты и доказательства K01/K02/K04.
Production, расчетная методика, solver tolerances и прежние expected values
не менялись. Новых классов нет. Основная output-книга не обновляется до
окончательной приемки Audit03.

## Контракт И Проверка

`modTestPlotConfig` копирует четыре реальные таблицы Config в собственную
временную книгу. Все 21 активные ячейки `rngPlotAnnotationSettings` поступают
через настоящий `CSystemSettingsReader.LoadFromWorkbook`, а не через заранее
подготовленные scalar keys. Затем проверяются независимые численные эффекты
`CPlotAnnotationLayout`; тип и размер стрелок дополнительно проверяются у
реальных `Chart.Shape` после `CSectionPlotter.Draw`.

Каждое поле проверено в двух положениях именованного диапазона: H5 и CH800.
Для каждого положения есть измененное допустимое значение и четыре
недопустимых: blank, TODO, INVALID, CVErr. Ошибки называют ключ, Config и
фактическую ячейку нового положения. Счетчик поиска равновесия не изменяется;
тестовая книга закрывается без сохранения даже при ошибке.

Пять неприменимых серых ячеек не объявляются пользовательскими настройками:
LineEnabled размера и ExtensionLineWeight/Color, ArrowType/Size арматуры.
Имена census `Plot.Rebar.*` / `Plot.Dimension.*` явно сопоставлены с runtime
`Plot.RebarLabels.*` / `Plot.Dimensions.*` в per-key evidence.

## Gates

- Directed `annotation_bridge_v230.log`: 425/0, 210 вариантов, 21 поле,
  два положения диапазона, equilibriumCases=0.
- Соседний `annotation_bridge_general_v230.log`: 2780/0. Сохраняются прежние
  проверки общих настроек схемы, snapshot metadata, AutoUpdate и contour arcs;
  также все поддержанные enum-варианты аннотаций и 15 комбинаций стрелок.
- Оба watchdog завершены: exit=0, source unchanged=True.
- Read-only структура: 27/27. Config/help/форматирование не менялись;
  их прежние gates не переименованы в новую выпускную приемку.
- Реальный VBE export: 111 components. Source/export equality 106/106,
  failed=0; экспорт не реконструирован из исходников.

Книга `RC_Section_NDM_annotation_bridge_v230.xlsm`:
SHA-256 `B756C108DC1560B85F85752C3817BAB5313C4B6AADCB69ED54E0634A010CDA3E`.
Экспорт `VBA_All_Code_annotation_v230_2026-10-04.txt`:
SHA-256 `75F671D851A64FC8D39F54BBF66F535F85626B9B8CEDF2D589FC20E08C093C20`.

## Реестр И Ограничения

Per-key evidence присоединен к v199 только после положительных runtime gates.
Новый реестр v230: 687 active-reviewed из 1064 фактических полей; из 760
editable осталось 73 без активной приемки. Coverage остается
`ActiveBehaviorAccepted:FullRangeReviewNotComplete`, fullAcceptance=False.
Изменение каждого поля и ошибки адреса проверены; это не полный непрерывный
диапазон Double, не mutation sensitivity всех call sites и не release pairwise.

Полный восьми-suite On/Off не повторялся: production не менялся, а пользователь
разрешил проверки соразмерно риску с обязательным финальным полным прогоном.
Это не закрывает общую цель, финальные load matrix/benchmarks/self-audit.
Открытые книги пользователя не подключались и не закрывались.
