# Audit03: Настройки AutoCAD И Реальный DWG

## Объем И Среда

Проверены двадцать еще не принятых editable CAD-полей: выбор сохраненного LC,
NeutralLine/PrincipalAxes/LoadPoint/Contour/LabelMode, семь слоев экспорта,
пять индексов цвета и два слоя импорта. AutoCAD.Import.MinArea имеет отдельную
ранее сохраненную приемку. Новых class modules нет: 83 production + 3 test.

Используется именно Autodesk AutoCAD 2023, не SOFiPLUS:
`C:\Program Files\Autodesk\AutoCAD 2023\acad.exe`, COM Version
`24.2s (LMS Tech)`, HWND 25888782. Общий ProgID отсутствует, versioned
регистрация указывает на OEM-приложение, но GetObject подключается к ROT
уже запущенного Autodesk. Регистрации не менялись. Отдельный native тест
не запускает CAD и прекращается, если HWND изменился или открыт чужой DWG.

## Подтвержденные Дефекты

- Import/export подключались только через общий AutoCAD.Application.
  При занятой OEM-программой регистрации настоящее запущенное приложение
  не находилось. Теперь общий connector проверяет versioned ROT и FullName;
  принимает только acad.exe, не использует CreateObject и не запускает OEM.
- Blank/TODO/missing/invalid экспортных настроек заменялись defaults.
  Активные значения теперь обязательны. Reader и CAD consumer указывают
  ключ, действующий адрес Config, причину и способ исправления.
- Индексы цветов округлялись или принимались вне применимого диапазона.
  Layer.Color в AutoCAD 2023 отвергает 0. Поле применяется также к новому
  слою, поэтому утвержденный диапазон 1..255; 0/256, дробь и Long overflow
  отклоняются до экспорта. Все 255 допустимых индексов проверены native API.
- Не проверялись запрещенные символы и длина имен слоев. Проверка общего
  import/export контракта выполняется до изменения DWG; пробелы и проверенные
  русские имена допустимы. Длина не более 255, управляющие и запрещенные
  символы отклоняются, включая запятую. Это согласовано с
  [Autodesk ActiveX: имена объектов](https://help.autodesk.com/cloudhelp/2023/CHS/AutoCAD-ActiveX/files/GUID-7397BDDC-DB5C-484F-863A-30C2A329EF3A.htm).
- EnsureAcadLayer скрывал ошибку Add/Color. Теперь только отсутствие Item
  допускается при lookup; ошибка создания дает понятное сообщение. Цвет
  существующего слоя остается прежним, цвет новых сущностей назначается явно.
- Очистка могла удалить Region на annotation-слое. Region теперь всегда
  сохраняются. Остальные сущности удаляются только с заданных слоев оформления;
  это не обещание сохранности чужого Text/Line на служебном слое.
- Геометрические слои нельзя совместить с annotation/contour/fixed technical
  слоями очистки; importer не принимает одинаковые бетонный и арматурный слои.
  Межполевые ошибки используют текущие Config-адреса.
- Cleanup больше не требует неиспользуемые цвета, label mode и флаги.
  Export при ContourEnabled=No не читает и не создает неактивный contour layer;
  cleanup читает его для удаления оформления предыдущего экспорта.

Для Color API важно различать объект и Layer, а не слепо переносить диапазон
0..256 между ними; ByLayer не допустим для Layer.
[Autodesk ActiveX: Color](https://help.autodesk.com/cloudhelp/2026/ENU/AutoCAD-ActiveX-Reference/files/GUID-A186333C-4857-4DFD-97EC-CDF75523299B.htm).
Отказ Layer.Color=0 дополнительно воспроизведен на установленной версии.

## Направленные Проверки Config

`modTestAutoCADConfig.RunAudit03AutoCADConfigTests`: **1279/0**, 392 вызова
рабочих consumers. Проверены исходное положение rngSystemSettings и перенос
в CR800. Адрес значения после переноса вычисляется через именованный Range,
а не сохраняется как прежний B-адрес. Formulas и привязка имени восстановлены.

Два настоящих равновесных StrengthState сохранены в Results. Выбор CAD_CFG,
CAD_CFG_2 и Worst меняет выбранный snapshot, включая напряжения элементов;
presentation/readback/cleanup не увеличивают счетчик solve. Все PhysicalState
цвета и annotation layers проверены на растяжении, сжатии и NearZero.
Переключение импортных слоев меняет реальный набор Region-fixtures и координаты.

Оба boolean значения, все axes/label modes, ACI-границы, blank/TODO/text/CVErr,
missing/recovery, дробь/overflow, недопустимые имена и конфликты слоев проходят
штатные consumers. Fake entities являются дополнительным доказательством,
а не заменой настоящего DWG.

## Реальный Экспорт И Импорт

`modTestAutoCADConfig.RunAudit03RealAutoCADTests`: **409/0**. Четыре бетонных
Rectangle Region 100x100 мм с центрами (+/-50, +/-50), четыре стержня d16 мм
с центрами (+/-60, +/-60). Нагрузка: N=0, Mx=100000, My=50000 Н*мм,
ZeroMomentPerDepth=0. Проверка использует реально сохраненный Results.

- В native DWG восемь Region на нужных бетонном/арматурном слоях;
  реальные Text/Region имеют ожидаемые PhysicalState цвета и слои.
- Все ACI 1..255, сохранность существующего Layer.Color=7, русское имя
  и длина 255 проверены через настоящий Layer API при EXTNAMES=1.
- Подтверждены наличие axes/loadpoint/neutral line/contour и их отсутствие
  после выключения соответствующей функции; геометрия остается из восьми Region.
- Два label modes проверены по всем восьми реальным Text.
- После изменения текущего INPUT/OUTPUT Length на m, Area на m2 и знаков
  +N/+Mx/+My сохраненный snapshot экспортируется все равно в мм. Импорт дает
  прежние координаты/площади; реальные цвета не меняются. Ожидаемые геометрические
  числа сравниваются с абсолютным допуском 1e-8, указанным в сыром журнале.
- Presentation DWG сохранен с 24 сущностями. Дополнительный Region помещен
  на annotation-слой; cleanup сохраняет все девять Region. Clean DWG сохранен,
  закрыт, снова открыт и импортирован без потери исходных восьми Region.
- Счетчик solve после всех presentation операций равен исходному.

Исторические отказавшие прогоны не удаляются: v267 3/1 выявил Layer.Color=0;
v268 93/1 использовал слишком малый момент под активным фильтром fixture;
v269 80/1 задал недопустимый тестовый вариант знака; v270 260/1 выявил запятую
в имени Layer. Уточнены только неверные тестовые предпосылки и подтвержденные
CAD-дефекты; расчетные expected values/tolerances ядра не менялись.
Frozen Config v266 дал 575/552 после ремонта setup и до production исправлений.
Самый ранний v265 38/3 имел недопустимую одноточечную section fixture.

## Справка И Native Книга

v272 обновлена через существующий source-only refresh и явный update guide.
Help: failed=0, 2078 строк, 140 прямых links, 118 Shapes. Все 760 пользовательских
inputs/Formulas/validation сохранены после save/reopen; signature до и после
FC419EFAF227960B204993E6E70F3897959103F15072892FF91DFE80EE8C0275.
Structure: 27/27. Native CodeModule export: 108/108 source modules, failed=0.
Это фактический экспорт книги, не конкатенация src/tests.

Справка объясняет mm import/export, смену единиц после импорта, required inputs,
ACI 1..255, слои и очистку, неактивный contour и сохранение существующего слоя.
Политика знаков остается внутренней: пользовательские +N/Mx/My не меняют
физический знак strain/stress и PhysicalState сохраненного НДС.

Артефакты находятся в docs/regression/Audit03:

| Артефакт | SHA-256 |
| --- | --- |
| RC_Section_NDM_autocad_help_v272.xlsm | 46C58381CBC609C5562B36CF03517CE58BB9F4A48401CEC937A93A0F339F5A76 |
| VBA_All_Code_autocad_v272_2026-10-04.txt | C5E9CFA311DF3282AB7999FF96E4B930177AD82C79449519F13B270F212ED34F |
| Audit03_AutoCAD_roundtrip_v271_2026-10-04_presentation.dwg | 6CB65D8D21B95AB24037718D6BB770F2CAD85E33E90F8ABF3A717C031F8C4674 |
| Audit03_AutoCAD_roundtrip_v271_2026-10-04.dwg | 57FABF0C858C55EEBA7553F274142F7877BB0C45F807DAA79FFA834C2C412A1B |

Журналы: autocad_config_positive_v271, autocad_real_roundtrip_v271,
autocad_help_v272, autocad_structure_v272, autocad_source_contracts_v272.
Основная пользовательская output-книга и ТЗ остались с исходными хешами.

## Открытые Gates

Текущий восьмисерийный Off regression запущен на независимой копии v272;
его результат пока не присвоен. Registry остается v257, 29 editable полей
не приняты до завершенного общего gate и присоединения поадресной трассировки.
Ранее красный full Off v263 сохраняется как отрицательное доказательство.
Directed green не подменяет final full On/Off, матрицу, benchmark и self-audit.

Не заявляется pixel-проверка: sky capture дважды вернул timeout для окна
AutoCAD, включая открытый read-only presentation DWG. Текстовое UI-дерево
доступно, но пиксели не просмотрены. Более строгий EXTNAMES=0, дополнительные
версии CAD и полный набор произвольных DWG не объявляются проверенными.
AutoCAD.Application.24.2 успешно подключился к реальному acad.exe; OEM apps
не используются. Audit03 в целом остается в работе.
