# Audit02: карта единого API

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
| Три поля публикации options | CRebarRuleOptions.SetPublication | Только AnnotationGroup/Publication/Coincidence | Повторные присваивания этих полей |
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

ERebarOffsetAxis: RowNormal=1, RowTangent=2. SetPublication меняет строго три
поля; ConfigureDependentContact — только NormalOffset/TangentOffset/EverySecond.
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
