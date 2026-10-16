# Audit02: карта единого API

Дополнение 16.10.2026: повторяющиеся ветви дополнительных рядов штатных рецептов
объединены циклами. Четыре приватных AdditionalOptions/ExtraOptions заменены одним
RecipeContactOptions в modSectionRecipeRows. Помощник только выбирает направление
контактного смещения и вызывает прежний ConfigureDependentContact; геометрия,
чтение таблиц, ID, комментарии и отдельные политики остаются в рецепте.
Полный прежний API сохранён. Актуальная приёмка описана в
`../../reports/RecipeSimplification_2026-10-16.md`; исторический ход аудита приведён ниже.

Исходная точка: 490f9363c803f5ae5833212887d7ee8674fd125d, книга SHA268d5fd5…,
153 native-компонента, 148 исходных компонентов. Аудит версии 1.0 прочитан
полностью. Эта карта обновляется по реализации, а не выдаёт предложения за
готовый API. Публикация пока не выполнена.

| Текущее средство | Целевой контракт | Потребители / сохраняемый смысл | Удаление |
|---|---|---|---|
| points/radii/names, AddCorner в RoundedRectangle | CPolygonSpec.AddVertex → Seal | 4/5/6 вершин, независимые стороны, точные координаты/радиусы, источник | Ручные массивы и AddCorner |
| Corners + массив радиусов/имён в Hollow | Два sealed CPolygonSpec | Наружная граница и отверстие, повторное использование для offsets | Corners как массивный helper |
| Массивы в учебном RoundedTrapezoid | CPolygonSpec | Один учебный ISectionRecipe, прежняя трапеция и отверстие | Старые вызовы фасада |
| facade/query.RoundedPolygon(pathID, vertices, radii, vertexIDs) | RoundedPolygon(pathID, polygon As CPolygonSpec) | Штатные рецепты, тесты, программные примеры | Публичная массивная сигнатура |
| facade/query.OffsetRoundedPolygon(pathID, vertices, radii, vertexIDs, distance) | OffsetRoundedPolygon(pathID, polygon, distance) | Принятый inward-отступ многоугольника, F02 и коллапс-привязки | Публичная массивная сигнатура |
| Низкоуровневые массивы query | Private kernel и один Friend CopyKernelData описателя | Действующая математика без переписывания | Публичные адаптеры/V2/Ex не создаются |
| Циклы CONTOUR_LINE_i/ARC_i в рецептах | Outer/Opening с политикой SequentialByKind и префиксом как данными | Общий номер сегмента; Hollow: CONTOUR_OUTER/OPENING, остальные CONTOUR | Техническая нумерация в рецептах |
| Настройки публикации options | AnnotationGroup, Publication, Coincidence | Группа подписи, порядок вывода и обработка совпадений задаются явно | Дублирующая команда настройки трёх полей |
| Contact distance в AdditionalOptions/ExtraOptions | ConfigureDependentContact | RowNormal/RowTangent, знак, EverySecond, номинальный промежуточный диаметр | Повторная контактная формула |
| Length + проверка >0/>=0 | PositiveLength/NonNegativeLength в CSectionInputTable | То же чтение и перевод единиц, фактический адрес | Только эквивалентные проверки; GEOM_TOLERANCE остаётся явным |
| ColumnByCaption внутри цикла граней | Однократный ленивый поиск при первом активном ряде | Проверки неоднозначности; выключенные поля не начинают проверяться | Повторный поиск заголовков |
| CirclePath, AdjacentRectangles, Chain, Offset, Transform, ConstrainLine | Сохранены | Геометрические постановки Circle и RectSet не подменяются polygon | Нет |
| Anchor/PathAnchor + Dimension Aligned | Сохранены и проверяются | Виртуальная вершина отличается от точки касания; истинная длина наклонного размера | Нет |

Предыдущие исправления уже присутствуют и повторно не реализуются:
независимые математические примитивы Path/Query, отдельные аннотации,
закрытые Mesh/Rebars/Rules, независимые units и ранняя проверка арматуры.
Cover-обёртка уже удалена. Старый пример RebarRowExtentLabel с axisDistance
адаптируется к текущей пятиаргументной сигнатуре; параметр не возвращается.

План: baseline → описатель и тесты структуры → единые фасад/query и все
потребители → общие options/имена/чтение → адресная матрица и измерения →
гайд/Справка → проверка данных и фактического VBA → FINAL_REPORT C01–C15.
Полные ON/OFF не запускаются. Push разрешён последним прямым заданием
пользователя; запрет push внутри исходного документа этим заменён.

## Фактический перенос: геометрия и общие параметры

RoundedPolygon/OffsetRoundedPolygon у фасада и query принимают только
CPolygonSpec. Прежние массивы доступны исключительно Private-ядру query;
один Friend CopyKernelData находится в описателе. Все 44 прямых тестовых
вызова и десять fillet-fixtures перенесены; FixtureFillets также принимает
CPolygonSpec. В RoundedRectangle удалены AddCorner, три массива и копирующий
цикл; в Hollow два spec повторно используются при offsets. Circle/RectSet
сохранили свои примитивы. Учебная трапеция перенесена на тот же формат.

Технические ID публикует CSectionDefinition по EContourNamingPolicy:
PreserveSegmentIDs=0, SequentialByKind=1. NamingPrefix — данные, по умолчанию
CONTOUR; для Hollow передаются CONTOUR_OUTER/CONTOUR_OPENING. Явные ID нельзя
смешивать с SequentialByKind; одинаковые последовательные префиксы разных
границ отклоняются. Общий номер сегмента и весь прежний порядок сохранены.

ERebarOffsetAxis: RowNormal=1, RowTangent=2. AnnotationGroup, Publication и
Coincidence задаются отдельными присваиваниями; ConfigureDependentContact
меняет только NormalOffset/TangentOffset/EverySecond.
Все четыре семейства используют общую подготовку. Локальные AdditionalOptions
и ExtraOptions сохраняют только конкретную интерпретацию ввода и направления,
не исполняют арматуру. PositiveLength/NonNegativeLength используют прежний
Length и адресный InputError; GEOM_TOLERANCE и составные проверки сохранены.
Поиск заголовков в циклах Rounded/Hollow выполняется лениво один раз.

Доказательства: UnifiedGeometryEquality.json — восемь полных снимков побайтно
прежние; UnifiedCalculationsSP35/SP63 — по восемь сочетаний, числа/статусы и
пояснения совпадают; UnifiedTestsCorrected и Audit02ContractTests — реальные
native-прогоны. Ошибки исходных структурных ожиданий отдельно исправлены:
их владелец теперь CPolygonSpec, топологических — прежний query.

Остаются полный комментарийный гайд/примеры, визуализация и адресная матрица
по всем критериям, измерения после переноса, CAD/export, приёмка и публикация
книги и FINAL_REPORT. Этот список не сокращает согласованный объём.

## Завершённая приёмка Audit02

Карта выше реализована без временных adapters. Private RoundedPolygonKernel
и OffsetRoundedPolygonKernel сохраняют численную математику; единственный
Friend CopyKernelData принадлежит CPolygonSpec. Нет новых публичных синонимов,
OuterPolygon, Custom или Offset группы. Статические gates — ApiRetirementChecks.json.

Гайд modSectionRecipeApiGuide остаётся comment-only и содержит прежние девять
разделов: 136 членов автора, полный минимальный ISectionRecipe и регистрация,
два полных выполненных примера геометрии/правил/измерений. Общий каталог:
292 члена, все с фактическими вызовами и native-журналами. Учебная трапеция
перенесена и выполнена; учебные классы не добавляются в рабочий VBProject.
Справка сохраняет 61 строку и перечисляет четыре штатных рецепта-примера.

RegressionEquality.json: 11 снимков без исключений и девять схем по пикселям.
Три линейных измерения 300/400/500, реальная наклонная грань, виртуальная точка,
смешанная длина и диагностика исчезнувшего/нулевого участка проверены отдельно.
Performance.json содержит по пять повторов десяти форм, без заявления ускорения.
Preservation.json: 107461 ячейка, 27 имён, только пять уточнённых строк Справки.
FinalSourceEquality.json: 152 source, 157 native компонентов, без различий.

V02 имеет статус BLOCKED: AutoCAD 2023 ROT возвращает существующий чужой PID
9344 вместо собственного нового PID18440. Guard остановил работу до документов;
свой процесс закрыт. Это ограничение среды, не PASS. Полная матрица 42 позиций,
измерения каждого рецепта/общего кода и C01–C15 — в FINAL_REPORT.md. Публикацию
и окончательный ALL CODE подтверждает независимый readback Publication.json.

## Закрывающая доработка: точные комментарии и контракт RectSet

ConfigureDependentContact требует оба диаметра >0, ноль является ошибкой;
описание исправлено, алгоритм прежний. Аргументы полного многострочного
вызова разбираются вместе, `_` не попадает в contourIDs. Публичный polygon
описывается CPolygonSpec; массивы относятся к Private-ядру. AxisPath получает
фактический polygon/outerPolygon/openingPolygon, sizes — скалярные параметры.

Для H1/B1/H2/B2 RectSet зафиксировано разрешённое исключение: порядок
«прочитать оба числа → проверить H → проверить B» и прежняя адресная ошибка
не совместимы с немедленной проверкой PositiveLength без расширения API или
повторного чтения. Length + Require оставлены; это проверено 252 адресными
случаями и полным сравнением диагностик. Подробности — в начале FINAL_REPORT.
