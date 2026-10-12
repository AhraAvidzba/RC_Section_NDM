Attribute VB_Name = "modSectionRecipeApiGuide"
' ========================================================================================================
' ДЕКЛАРАТИВНОЕ ОПИСАНИЕ СЕЧЕНИЯ — ПОЛНЫЙ СПРАВОЧНИК API
' ========================================================================================================
' Модуль содержит только комментарии. Тела классов и методов находятся в соответствующих модулях VBE.
' Поиск: Ctrl+F по имени класса или метода. Краткий порядок подключения — на листе Справка.
' Внутренние единицы: мм, мм^2, Н, Н·мм, МПа; углы в радианах. Положительное N — растяжение.
' Ввод преобразует CUnitSystem; рецепт получает снимки значений, не Excel-объекты.
' ========================================================================================================
' ОСНОВНЫЕ ПОНЯТИЯ
' ========================================================================================================
' Путь (Path) — непрерывная геометрическая линия: отрезок, дуга, окружность или последовательность таких
' участков. Замкнутый путь может задавать границу бетона или отверстия; путь размещения арматуры задаёт
' линию центров стержней.
' Сегмент (Segment) — один отрезок или одна дуга в составе пути. Группа путей (PathGroup) объединяет
' независимые линии без соединений между ними.
' Положение вдоль линии задаётся расстоянием от её начала в мм. В именах API это расстояние может
' называться Station; например, RebarAtStations задаёт места стержней массивом таких расстояний.
' Правило арматуры (Rule) — описание диаметра и мест размещения стержней: по числу стержней, расстоянию
' между ними, отдельным местам или связи с другим рядом. Общий исполнитель по правилу создаёт физические
' стержни.
' ID — уникальное текстовое имя объекта внутри описания сечения. Привязка (Anchor) — именованная точка,
' используемая для размеров и подписей.
' ========================================================================================================
' ЖИЗНЕННЫЙ ЦИКЛ
' ========================================================================================================
' CSectionDefinitionBuilder: New -> Initialize один раз -> декларации -> Finish. Ссылки требуют ранее
' объявленных путей/привязок; правило зависимости разрешается при Finish. После Finish новые декларации
' запрещены.
' CGeometryPath: New -> Initialize -> AddLine/AddArc -> Seal. Методы чтения/slice/transform требуют Seal.
' Они возвращают отдельные опубликованные пути; исходный не меняется.
' CSectionDefinition: Получается из Finish. Геометрия, правила и SourceMap неизменяемы; наружу
' возвращаются скаляры, копии массивов или sealed paths/clones.
' CRebarRuleOptions: New задаёт defaults; автор меняет поля до объявления правила. Фасад копирует options
' в definition; последующее изменение options не меняет правило.
' CGeometryDiagnostic: Создаётся сервисом как immutable snapshot. Читайте через
' LastDiagnostic/DiagnosticAt; диагностический объект не изменяет геометрию или статусы.
' CSectionInputTable: Initialize один раз копирует значения/адреса. Последующее изменение листа/исходного
' массива не меняет снимок; ошибка Excel не скрывается.
' CSectionRecipeContext: Initialize -> AddTable -> Seal на внешней границе. Рецепт только читает
' Table/Units; не сохраняет context в готовом definition.
' CSectionInputSnapshotReader: Внешний workbook+список имён+подготовленные units -> новый sealed context.
' Невыбранные блоки не читаются; live Excel не передаётся геометрическому ядру.
' CSectionTypeRegistry: Новая instance регистрирует текущие типы. RegisterRecipe добавляет один
' экземпляр/metadata; consumer выбирает только зарегистрированный TypeID.
' ISectionRecipe: Вызывается реестром с готовым context; создаёт новое definition. Не выполняет
' mesh/layout/solver/Excel output.
' CSectionBuilder: Один Build выполняет все общие стадии для immutable definition. После успеха доступны
' Mesh/Rebars/Rules; после нового отказа прошлый успешный результат не выдаётся.
' CRebarRuleExecutor: Build начинает новый отчёт/массивы, даже при раннем отказе. Чтение
' происхождения/шагов после исполнения; позиции не являются независимой копией mesh.
' CUnitSystem: InitializeDefaults для программного источника либо LoadFromSettings на внешней границе. В
' рецепте преобразование каждого параметра один раз; сохранённые Results читаются с savedUnit.
' CSectionGeometryQuery: Общие операции над готовыми областями/путями; Initialize нужен для query-session
' по CSectionModel. Новая revision требует актуальной session, исходная область/контуры не меняются.
' CConcreteRegion: Готовый аналитический snapshot общего query. Геометрия read-only; Clone отделяет
' следующий рабочий результат. SetOwner меняет только скалярную инженерную metadata.
' modSectionRecipeTypes: Общие enum/UDT/error contracts, только внутренние единицы. Raw правила/аннотации
' заполняет фасад; автор не реализует координатные циклы баров через UDT.
' modGeometryTypes: Общие математические scalar/UDT/adapters. Внутренние единицы; helpers не выбирают тип
' сечения и не обращаются к Excel.
' ========================================================================================================
' ОШИБКИ И ОГРАНИЧЕНИЯ
' ========================================================================================================
' CSectionDefinitionBuilder: 5530: пустые/повторные ID, неверный режим, отрицательные d/count/setbacks,
' незакрытый фасад. Query/path ошибки связывайте SourceMap; геометрию/направление исправляйте по
' диагностике.
' CGeometryPath: 5520: пустой ID или сегмент, разрыв линии, недопустимый радиус или угол дуги, незамкнутые
' концы, отсутствие сегмента или заданное расстояние за пределами длины линии. Исправьте геометрию, имя
' сегмента или рабочий интервал.
' CSectionDefinition: 5530 либо исходный query/path код: топология Outer/Opening, отсутствующий
' target/anchor/rule, циклы зависимостей. Не выбирать соседний сегмент или fallback тип молча.
' CSectionInputTable: 5525/адресные input codes: несогласованные массивы, пустое обязательное поле, ошибка
' Excel, нечисло/Boolean, повтор ключа/заголовка, неверный Choice/Count. Исправьте реальную ячейку.
' CSectionRecipeContext: 5530: нет units/table, повтор имени, добавление после Seal. Готовьте новый
' context на внешней границе.
' CSectionInputSnapshotReader: Адресные ошибки имени/блока/формулы. Повреждённый выбранный блок не равен
' отсутствию потребителя; восстановите имя/ячейку.
' CSectionTypeRegistry: 5530: неизвестный/повторный TypeID, пустой recipe/context, повреждённые metadata.
' Добавьте одну действительную регистрацию.
' CSectionBuilder: 5530: неопубликованное definition, неверные meshOptions, отсутствие баров/успешной
' сборки. Исправьте ввод; старый результат после отказа не выдаётся.
' CRebarRuleExecutor: 5530: неверные options/длина/step/ClosedUniform, отсутствие hit, конфликт
' владельца/размещения. Читайте Code/ID/числа/source; Created — фактическое число.
' CSectionGeometryQuery: 5401: нарушенная топология/радиусы/ограничения/отсутствующая область. Новые
' offsets/constraints имеют явные Code/ID/числа; GEOMETRY_INVALID используется прочими контрактами.
' CConcreteRegion: 5400: индекс вне SegmentCount/LoopCount. Используйте диапазоны готового snapshot;
' пустая область и Nothing различаются.
' CUnitSystem: Адресные ошибки units/sign settings; пустой/неизвестный активный selector не заменяется
' default. Приводите единицы через CUnitSystem один раз.
' ========================================================================================================
' ISectionRecipe
' ========================================================================================================
' --------------------------------------------------------------------------------------------------------
' ISectionRecipe.CreateDefinition
' Интерфейс единственного класса типа. Возвращает опубликованное definition через общий фасад, без
' сетки/баров/Shapes в рецепте.
' Public Function CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
' ByVal context As CSectionRecipeContext — Подготовленный снимок CSectionRecipeContext, не live Excel.
' Код метода или объявление: ISectionRecipe в окне проекта этой книги (Ctrl+R).
' ========================================================================================================
' CSectionDefinitionBuilder
' ========================================================================================================
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.Initialize
' Открывает один новый фасад с TypeID и версией схемы. Повторная инициализация запрещена.
' Public Sub Initialize(ByVal typeID As String, Optional ByVal schemaVersion As Long = 1)
' ByVal typeID As String — Смысловое имя зарегистрированного типа; непустой текст.
' Optional ByVal schemaVersion As Long = 1 — Положительная целая версия схемы определения.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.CirclePath
' Создаёт замкнутую окружность CCW от точки +X; один сегмент с тем же ID, что путь.
' Public Function CirclePath(ByVal pathID As String, ByVal centerX As Double, ByVal centerY As Double,
' ByVal radius As Double) As CGeometryPath
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal centerX As Double — Модельная координата centerX, мм; знаки соответствуют section XY/WCS.
' ByVal centerY As Double — Модельная координата centerY, мм; знаки соответствуют section XY/WCS.
' ByVal radius As Double — Положительный радиус, мм.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.LinePath
' Создаёт непрерывный открытый путь из одной прямой от первого конца ко второму.
' Public Function LinePath(ByVal pathID As String, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As
' Double, ByVal y2 As Double) As CGeometryPath
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal x1 As Double — Модельная координата x1, мм; знаки соответствуют section XY/WCS.
' ByVal y1 As Double — Модельная координата y1, мм; знаки соответствуют section XY/WCS.
' ByVal x2 As Double — Модельная координата x2, мм; знаки соответствуют section XY/WCS.
' ByVal y2 As Double — Модельная координата y2, мм; знаки соответствуют section XY/WCS.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.ArcPath
' Создаёт точную дугу по центру/радиусу/startAngle/signed sweep; полный оборот замкнут.
' Public Function ArcPath(ByVal pathID As String, ByVal centerX As Double, ByVal centerY As Double, ByVal
' radius As Double, ByVal startAngle As Double, ByVal sweep As Double) As CGeometryPath
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal centerX As Double — Модельная координата centerX, мм; знаки соответствуют section XY/WCS.
' ByVal centerY As Double — Модельная координата centerY, мм; знаки соответствуют section XY/WCS.
' ByVal radius As Double — Положительный радиус, мм.
' ByVal startAngle As Double — Начальный угол, радианы; от +X.
' ByVal sweep As Double — Signed sweep, радианы: CCW положителен, CW отрицателен, 0 запрещён,
' |sweep|<=2*pi.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.RoundedPolygon
' Создаёт согласованную границу по вершинам/радиусам/ID. Edge.ID начинается у указанной вершины, Corner.ID
' существует только при R>0.
' Public Function RoundedPolygon(ByVal pathID As String, ByVal vertices As Variant, ByVal radii As
' Variant, ByVal vertexIDs As Variant) As CGeometryPath
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal vertices As Variant — Двумерные вершины: Array(Array(x,y),...); координаты во внутренних мм.
' ByVal radii As Variant — Массив неотрицательных радиусов вершин в мм, длина равна vertices; clamp
' бетонного радиуса отсутствует.
' ByVal vertexIDs As Variant — Массив уникальных смысловых ID вершин; определяет Edge.ID/Corner.ID.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.OffsetRoundedPolygon
' Смещает весь скруглённый многоугольник с согласованными стыками. Положительное расстояние внутрь
' площади; отрицательное наружу, из Opening в бетон.
' Public Function OffsetRoundedPolygon(ByVal pathID As String, ByVal vertices As Variant, ByVal radii As
' Variant, ByVal vertexIDs As Variant, ByVal distance As Double) As CGeometryPath
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal vertices As Variant — Двумерные вершины: Array(Array(x,y),...); координаты во внутренних мм.
' ByVal radii As Variant — Массив неотрицательных радиусов вершин в мм, длина равна vertices; clamp
' бетонного радиуса отсутствует.
' ByVal vertexIDs As Variant — Массив уникальных смысловых ID вершин; определяет Edge.ID/Corner.ID.
' ByVal distance As Double — Signed offset, мм: положительно внутрь площади, отрицательно наружу; это не
' left-offset обхода.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.AddPath
' Добавляет уже опубликованный непрерывный CGeometryPath в определение; не включает его автоматически в
' бетон.
' Public Sub AddPath(ByVal path As CGeometryPath)
' ByVal path As CGeometryPath — Существующий непрерывный опубликованный CGeometryPath.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.Anchor
' Объявляет именованную модельную точку. Данные копируются как два скаляра.
' Public Sub Anchor(ByVal anchorID As String, ByVal x As Double, ByVal y As Double)
' ByVal anchorID As String — Непустой уникальный ID модельной привязки.
' ByVal x As Double — Модельная координата x, мм; знаки соответствуют section XY/WCS.
' ByVal y As Double — Модельная координата y, мм; знаки соответствуют section XY/WCS.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.PathAnchor
' Получает точку конкретного сегмента по ID и параметру [0;1], объявляет именованную привязку.
' Public Sub PathAnchor(ByVal anchorID As String, ByVal pathID As String, ByVal segmentID As String,
' Optional ByVal parameter As Double = 0#)
' ByVal anchorID As String — Непустой уникальный ID модельной привязки.
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal segmentID As String — Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется
' индексом.
' Optional ByVal parameter As Double = 0# — Нормированный параметр точки сегмента [0;1], не расстояние по
' всему пути.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.Source
' Связывает смысловой ID со строкой источника/адреса без сохранения объекта Excel.
' Public Sub Source(ByVal entityID As String, ByVal sourceReference As String)
' ByVal entityID As String — Смысловой ID сущности для SourceMap.
' ByVal sourceReference As String — Строка исходного адреса/программного параметра; без объекта Excel.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.TransformAnchor
' Создаёт новую привязку из существующей: отражение по локальной X, поворот вокруг (0,0), перенос.
' Public Sub TransformAnchor(ByVal anchorID As String, ByVal sourceAnchorID As String, ByVal shiftX As
' Double, ByVal shiftY As Double, Optional ByVal rotation As Double = 0#, Optional ByVal reflectX As
' Boolean = False)
' ByVal anchorID As String — Непустой уникальный ID модельной привязки.
' ByVal sourceAnchorID As String — ID существующей исходной привязки.
' ByVal shiftX As Double — Перенос по X, мм.
' ByVal shiftY As Double — Перенос по Y, мм.
' Optional ByVal rotation As Double = 0# — Угол поворота вокруг (0,0), радианы, CCW положителен.
' Optional ByVal reflectX As Boolean = False — Отражение относительно локальной оси X перед поворотом и
' переносом.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.Offset
' Смещает линию/аналитическую цепочку слева от обхода при положительном signedLeftDistance. Line/line
' стыки miter; arc стыки должны быть гладкими.
' Public Function Offset(ByVal pathID As String, ByVal sourcePathID As String, ByVal signedLeftDistance As
' Double) As CGeometryPath
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal sourcePathID As String — ID ранее объявленного пути, который читается без изменения.
' ByVal signedLeftDistance As Double — Поперечное смещение, мм; положительно слева от направления пути,
' отрицательно справа.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.ConstrainLine
' Ограничивает или продлевает одну прямую двумя поперечными плоскостями nX*x+nY*y=bound; направление
' исходного пути сохраняется.
' Public Function ConstrainLine(ByVal pathID As String, ByVal sourcePathID As String, ByVal startNX As
' Double, ByVal startNY As Double, ByVal startBound As Double, ByVal endNX As Double, ByVal endNY As
' Double, ByVal endBound As Double) As CGeometryPath
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal sourcePathID As String — ID ранее объявленного пути, который читается без изменения.
' ByVal startNX As Double — X-компонента начальной поперечной нормали; линия не должна быть параллельна
' плоскости.
' ByVal startNY As Double — Y-компонента начальной поперечной нормали.
' ByVal startBound As Double — Значение начальной плоскости startNX*x+startNY*y=startBound; мм при
' единичной нормали.
' ByVal endNX As Double — X-компонента конечной поперечной нормали.
' ByVal endNY As Double — Y-компонента конечной поперечной нормали.
' ByVal endBound As Double — Значение конечной плоскости endNX*x+endNY*y=endBound; мм при единичной
' нормали.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.LastGeometryDiagnostic
' Возвращает последнюю диагностику query фасада с SourceMap. Nothing означает отсутствие записанного
' отказа.
' Public Property Get LastGeometryDiagnostic() As CGeometryDiagnostic
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.Chain
' Выбирает непрерывный участок между двумя именованными сегментными привязками. Направление обхода явно;
' fullLoopIfSame задаёт полный оборот при совпавших концах.
' Public Function Chain(ByVal pathID As String, ByVal sourcePathID As String, ByVal startSegmentID As
' String, ByVal startParameter As Double, ByVal endSegmentID As String, ByVal endParameter As Double,
' Optional ByVal forward As Boolean = True, Optional ByVal fullLoopIfSame As Boolean = False) As
' CGeometryPath
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal sourcePathID As String — ID ранее объявленного пути, который читается без изменения.
' ByVal startSegmentID As String — Сегмент начальной привязки цепочки.
' ByVal startParameter As Double — Параметр начальной точки выбранного сегмента [0;1].
' ByVal endSegmentID As String — Сегмент конечной привязки цепочки.
' ByVal endParameter As Double — Параметр конечной точки выбранного сегмента [0;1].
' Optional ByVal forward As Boolean = True — Явное направление следования исходному обходу; False —
' противоположное.
' Optional ByVal fullLoopIfSame As Boolean = False — Полный оборот при совпавших привязках замкнутого
' пути; иначе нулевая цепочка недопустима.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.Transform
' Создаёт новый путь из существующего изометрией; сохраняет сегментные ID и происхождение, отражение
' меняет знак sweep.
' Public Function Transform(ByVal pathID As String, ByVal sourcePathID As String, ByVal shiftX As Double,
' ByVal shiftY As Double, Optional ByVal rotation As Double = 0#, Optional ByVal reflectX As Boolean =
' False) As CGeometryPath
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal sourcePathID As String — ID ранее объявленного пути, который читается без изменения.
' ByVal shiftX As Double — Перенос по X, мм.
' ByVal shiftY As Double — Перенос по Y, мм.
' Optional ByVal rotation As Double = 0# — Угол поворота вокруг (0,0), радианы, CCW положителен.
' Optional ByVal reflectX As Boolean = False — Отражение относительно локальной оси X перед поворотом и
' переносом.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.RectanglePath
' Создаёт замкнутую наружную границу одного прямоугольника от (x,y) с положительными width/height.
' Public Function RectanglePath(ByVal pathID As String, ByVal x As Double, ByVal y As Double, ByVal width
' As Double, ByVal height As Double) As CGeometryPath
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal x As Double — Модельная координата x, мм; знаки соответствуют section XY/WCS.
' ByVal y As Double — Модельная координата y, мм; знаки соответствуют section XY/WCS.
' ByVal width As Double — Координата/размер width, мм; размеры width/height положительны.
' ByVal height As Double — Координата/размер height, мм; размеры width/height положительны.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.AdjacentRectangles
' Создаёт одну границу смежных неперекрывающихся прямоугольников. Произвольное пересекающееся CAD Union не
' выполняется.
' Public Function AdjacentRectangles(ByVal pathID As String, ByVal rectangles As Variant) As CGeometryPath
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal rectangles As Variant — Array(Array(x,y,width,height),...); мм, размеры положительные,
' прямоугольники смежные без перекрытий.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.PathGroup
' Объявляет массив ID независимых путей; порядок значим, фиктивные соединения отсутствуют.
' Public Sub PathGroup(ByVal groupID As String, ByVal pathIDs As Variant)
' ByVal groupID As String — Смысловой ID группы; путь и арматурная annotation-group — разные контракты.
' ByVal pathIDs As Variant — Массив ID существующих независимых путей; без повторов.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.Outer
' Назначает замкнутому пути роль наружной границы материала. Не требует наличия арматуры.
' Public Sub Outer(ByVal loopID As String, ByVal pathID As String, Optional ByVal comment As String = "",
' Optional ByVal contourIDs As Variant, Optional ByVal arcComment As String = "")
' ByVal loopID As String — Непустой устойчивый ID контура; у CConcreteRegion численный индекс loop
' 1..LoopCount.
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' Optional ByVal comment As String = "" — Пояснение происхождения/инженерного смысла; не управляющий
' статус.
' Optional ByVal contourIDs As Variant — Необязательный массив ID сегментов сохранённого контура, по
' одному на каждый сегмент пути.
' Optional ByVal arcComment As String = "" — Отдельный комментарий аналитических дуг; пустой использует
' обычный comment.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.Opening
' Назначает замкнутому пути роль отверстия с конкретным parentOuterID. Действительное вложение проверяется
' при Finish.
' Public Sub Opening(ByVal loopID As String, ByVal pathID As String, ByVal parentOuterID As String,
' Optional ByVal comment As String = "", Optional ByVal contourIDs As Variant, Optional ByVal arcComment
' As String = "")
' ByVal loopID As String — Непустой устойчивый ID контура; у CConcreteRegion численный индекс loop
' 1..LoopCount.
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal parentOuterID As String — ID ранее объявленной наружной границы для Opening; проверяется реальное
' вложение.
' Optional ByVal comment As String = "" — Пояснение происхождения/инженерного смысла; не управляющий
' статус.
' Optional ByVal contourIDs As Variant — Необязательный массив ID сегментов сохранённого контура, по
' одному на каждый сегмент пути.
' Optional ByVal arcComment As String = "" — Отдельный комментарий аналитических дуг; пустой использует
' обычный comment.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.RebarByCount
' Задаёт места для указанного числа стержней на рабочей части линии размещения. EndMode определяет
' включение концов. В группе это число применяется отдельно к каждой линии; нулевое число стержней или
' нулевой диаметр отключают ряд.
' Public Sub RebarByCount(ByVal ruleID As String, ByVal target As String, ByVal count As Long, ByVal
' diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal steelClass As
' String = "Rebar")
' ByVal ruleID As String — Непустой уникальный ID правила; при диагностике — ID ранее исполненного
' правила.
' ByVal target As String — ID пути/PathGroup, разрешаемый в этом определении.
' ByVal count As Long — Число стержней на одной линии размещения; целое неотрицательное. В группе
' применяется отдельно к каждой линии. Ноль отключает ряд.
' ByVal diameter As Double — Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд.
' Optional ByVal options As CRebarRuleOptions = Nothing — CRebarRuleOptions; поля копируются при
' объявлении. Nothing создаёт реальные default options.
' Optional ByVal steelClass As String = "Rebar" — Строка принадлежности материала арматуры; совпадения
' разных материалов требуют явной политики.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.RebarByMaxStep
' Задаёт места стержней так, чтобы расстояние между соседними стержнями вдоль линии не превышало step.
' Исполнитель подбирает число стержней; фактическое расстояние может быть меньше заданного.
' Public Sub RebarByMaxStep(ByVal ruleID As String, ByVal target As String, ByVal step As Double, ByVal
' diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal steelClass As
' String = "Rebar")
' ByVal ruleID As String — Непустой уникальный ID правила; при диагностике — ID ранее исполненного
' правила.
' ByVal target As String — ID пути/PathGroup, разрешаемый в этом определении.
' ByVal step As Double — Расстояние между соседними стержнями вдоль линии размещения, мм; строго
' положительное.
' ByVal diameter As Double — Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд.
' Optional ByVal options As CRebarRuleOptions = Nothing — CRebarRuleOptions; поля копируются при
' объявлении. Nothing создаёт реальные default options.
' Optional ByVal steelClass As String = "Rebar" — Строка принадлежности материала арматуры; совпадения
' разных материалов требуют явной политики.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.RebarByFixedStep
' Задаёт точное расстояние step между соседними стержнями вдоль линии. RemainderPlacement определяет, где
' оставить неиспользованный остаток длины. При IncludeBoth или ClosedUniform рабочая длина должна
' соответствовать заданному расстоянию.
' Public Sub RebarByFixedStep(ByVal ruleID As String, ByVal target As String, ByVal step As Double, ByVal
' diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal steelClass As
' String = "Rebar")
' ByVal ruleID As String — Непустой уникальный ID правила; при диагностике — ID ранее исполненного
' правила.
' ByVal target As String — ID пути/PathGroup, разрешаемый в этом определении.
' ByVal step As Double — Расстояние между соседними стержнями вдоль линии размещения, мм; строго
' положительное.
' ByVal diameter As Double — Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд.
' Optional ByVal options As CRebarRuleOptions = Nothing — CRebarRuleOptions; поля копируются при
' объявлении. Nothing создаёт реальные default options.
' Optional ByVal steelClass As String = "Rebar" — Строка принадлежности материала арматуры; совпадения
' разных материалов требуют явной политики.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.RebarAtStations
' Задаёт отдельные места стержней расстояниями от начала линии размещения, в мм. Порядок этих расстояний в
' массиве определяет исходную нумерацию стержней.
' Public Sub RebarAtStations(ByVal ruleID As String, ByVal target As String, ByVal stations As Variant,
' ByVal diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal
' steelClass As String = "Rebar")
' ByVal ruleID As String — Непустой уникальный ID правила; при диагностике — ID ранее исполненного
' правила.
' ByVal target As String — ID пути/PathGroup, разрешаемый в этом определении.
' ByVal stations As Variant — Массив расстояний от начала линии размещения до каждого стержня, мм; порядок
' элементов задаёт исходную нумерацию.
' ByVal diameter As Double — Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд.
' Optional ByVal options As CRebarRuleOptions = Nothing — CRebarRuleOptions; поля копируются при
' объявлении. Nothing создаёт реальные default options.
' Optional ByVal steelClass As String = "Rebar" — Строка принадлежности материала арматуры; совпадения
' разных материалов требуют явной политики.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.RebarDependent
' Задаёт дополнительный ряд смещением относительно центров стержней исходного ряда. TangentOffset задаёт
' смещение вдоль касательной, NormalOffset — поперёк линии. Число и места стержней связаны с исходным
' рядом, новое равномерное распределение не выполняется.
' Public Sub RebarDependent(ByVal ruleID As String, ByVal sourceRuleID As String, ByVal diameter As
' Double, ByVal options As CRebarRuleOptions, Optional ByVal steelClass As String = "Rebar")
' ByVal ruleID As String — Непустой уникальный ID правила; при диагностике — ID ранее исполненного
' правила.
' ByVal sourceRuleID As String — ID исходного правила зависимости; циклы запрещены.
' ByVal diameter As Double — Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд.
' ByVal options As CRebarRuleOptions — CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт
' реальные default options.
' Optional ByVal steelClass As String = "Rebar" — Строка принадлежности материала арматуры; совпадения
' разных материалов требуют явной политики.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.RebarProjected
' Задаёт новый ряд проекцией центров стержней исходного ряда на целевую линию target. Направление —
' нормаль исходной линии или общий вектор DirectionX/Y. StartSetback и EndSetback ограничивают рабочий
' участок целевой линии; отсутствие пересечения отклоняется либо пропускается по SkipMissingProjection.
' Public Sub RebarProjected(ByVal ruleID As String, ByVal target As String, ByVal sourceRuleID As String,
' ByVal diameter As Double, ByVal options As CRebarRuleOptions, Optional ByVal steelClass As String =
' "Rebar")
' ByVal ruleID As String — Непустой уникальный ID правила; при диагностике — ID ранее исполненного
' правила.
' ByVal target As String — ID пути/PathGroup, разрешаемый в этом определении.
' ByVal sourceRuleID As String — ID исходного правила зависимости; циклы запрещены.
' ByVal diameter As Double — Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд.
' ByVal options As CRebarRuleOptions — CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт
' реальные default options.
' Optional ByVal steelClass As String = "Rebar" — Строка принадлежности материала арматуры; совпадения
' разных материалов требуют явной политики.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.VerticalExtent
' Измеряет полный вертикальный габарит выбранного пути/группы с точными экстремумами дуг. placementX и
' normalX задают сторону отображения.
' Public Sub VerticalExtent(ByVal annotationID As String, ByVal target As String, ByVal name As String,
' ByVal placementX As Double, Optional ByVal normalX As Double = -1#, Optional ByVal comment As String =
' "")
' ByVal annotationID As String — Непустой уникальный ID аннотации.
' ByVal target As String — ID пути/PathGroup, разрешаемый в этом определении.
' ByVal name As String — Обозначение измерения или имя контейнера.
' ByVal placementX As Double — Модельная X размерной линии, мм.
' Optional ByVal normalX As Double = -1# — X-компонента стороны размерной/подписной линии; layout
' нормирует вектор.
' Optional ByVal comment As String = "" — Пояснение происхождения/инженерного смысла; не управляющий
' статус.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.HorizontalExtent
' Измеряет полный горизонтальный габарит выбранного пути/группы. Это не длина пути и не исходная ширина до
' скругления.
' Public Sub HorizontalExtent(ByVal annotationID As String, ByVal target As String, ByVal name As String,
' ByVal placementY As Double, Optional ByVal normalY As Double = 1#, Optional ByVal comment As String =
' "")
' ByVal annotationID As String — Непустой уникальный ID аннотации.
' ByVal target As String — ID пути/PathGroup, разрешаемый в этом определении.
' ByVal name As String — Обозначение измерения или имя контейнера.
' ByVal placementY As Double — Модельная Y размерной линии, мм.
' Optional ByVal normalY As Double = 1# — Y-компонента стороны размерной/подписной линии.
' Optional ByVal comment As String = "" — Пояснение происхождения/инженерного смысла; не управляющий
' статус.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.HorizontalDistance
' Объявляет размер между двумя заданными X на уровне placementY. DisplayIncrement округляет только
' подпись, не геометрию.
' Public Sub HorizontalDistance(ByVal annotationID As String, ByVal name As String, ByVal startX As
' Double, ByVal endX As Double, ByVal placementY As Double, Optional ByVal normalY As Double = 1#,
' Optional ByVal comment As String = "", Optional ByVal displayIncrement As Double = 0#)
' ByVal annotationID As String — Непустой уникальный ID аннотации.
' ByVal name As String — Обозначение измерения или имя контейнера.
' ByVal startX As Double — Модельная координата startX, мм; знаки соответствуют section XY/WCS.
' ByVal endX As Double — Модельная координата endX, мм; знаки соответствуют section XY/WCS.
' ByVal placementY As Double — Модельная Y размерной линии, мм.
' Optional ByVal normalY As Double = 1# — Y-компонента стороны размерной/подписной линии.
' Optional ByVal comment As String = "" — Пояснение происхождения/инженерного смысла; не управляющий
' статус.
' Optional ByVal displayIncrement As Double = 0# — Неотрицательный шаг округления только видимого
' справочного значения, мм; 0 означает отсутствие такого округления.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.VerticalDistance
' Объявляет вертикальный размер между двумя Y на линии placementX; значение — модуль разности.
' Public Sub VerticalDistance(ByVal annotationID As String, ByVal name As String, ByVal startY As Double,
' ByVal endY As Double, ByVal placementX As Double, Optional ByVal normalX As Double = -1#, Optional ByVal
' comment As String = "")
' ByVal annotationID As String — Непустой уникальный ID аннотации.
' ByVal name As String — Обозначение измерения или имя контейнера.
' ByVal startY As Double — Модельная координата startY, мм; знаки соответствуют section XY/WCS.
' ByVal endY As Double — Модельная координата endY, мм; знаки соответствуют section XY/WCS.
' ByVal placementX As Double — Модельная X размерной линии, мм.
' Optional ByVal normalX As Double = -1# — X-компонента стороны размерной/подписной линии; layout
' нормирует вектор.
' Optional ByVal comment As String = "" — Пояснение происхождения/инженерного смысла; не управляющий
' статус.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.Dimension
' Измеряет между двумя именованными точками: Horizontal/Vertical проекция или Aligned расстояние. Проекция
' размещается через первую точку.
' Public Sub Dimension(ByVal annotationID As String, ByVal startAnchorID As String, ByVal endAnchorID As
' String, ByVal name As String, Optional ByVal measurement As String = "Aligned", Optional ByVal normalX
' As Double = 0#, Optional ByVal normalY As Double = 0#, Optional ByVal comment As String = "")
' ByVal annotationID As String — Непустой уникальный ID аннотации.
' ByVal startAnchorID As String — ID начальной модельной точки измерения.
' ByVal endAnchorID As String — ID конечной модельной точки измерения.
' ByVal name As String — Обозначение измерения или имя контейнера.
' Optional ByVal measurement As String = "Aligned" — Ровно Horizontal, Vertical или Aligned; значение
' чувствительно к указанному контракту.
' Optional ByVal normalX As Double = 0# — X-компонента стороны размерной/подписной линии; layout нормирует
' вектор.
' Optional ByVal normalY As Double = 0# — Y-компонента стороны размерной/подписной линии.
' Optional ByVal comment As String = "" — Пояснение происхождения/инженерного смысла; не управляющий
' статус.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.PathLength
' Объявляет точную суммарную длину target-путей и текстовую привязку; не создаёт фиктивную соединяющую
' линию группы.
' Public Sub PathLength(ByVal annotationID As String, ByVal target As String, ByVal placementAnchorID As
' String, Optional ByVal name As String = "L", Optional ByVal comment As String = "")
' ByVal annotationID As String — Непустой уникальный ID аннотации.
' ByVal target As String — ID пути/PathGroup, разрешаемый в этом определении.
' ByVal placementAnchorID As String — ID точки размещения текстового измерения.
' Optional ByVal name As String = "L" — Обозначение измерения или имя контейнера.
' Optional ByVal comment As String = "" — Пояснение происхождения/инженерного смысла; не управляющий
' статус.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.Radius
' Измеряет радиус конкретной Arc/Circle; параметр выбирает конечную точку радиуса. Прямая — ошибка.
' Public Sub Radius(ByVal annotationID As String, ByVal pathID As String, ByVal segmentID As String,
' Optional ByVal parameter As Double = 0.5, Optional ByVal name As String = "R", Optional ByVal comment As
' String = "")
' ByVal annotationID As String — Непустой уникальный ID аннотации.
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal segmentID As String — Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется
' индексом.
' Optional ByVal parameter As Double = 0.5 — Нормированный параметр точки сегмента [0;1], не расстояние по
' всему пути.
' Optional ByVal name As String = "R" — Обозначение измерения или имя контейнера.
' Optional ByVal comment As String = "" — Пояснение происхождения/инженерного смысла; не управляющий
' статус.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.Diameter
' Измеряет диаметр конкретной Arc/Circle с противоположным концом на той же окружности. Прямая — ошибка.
' Public Sub Diameter(ByVal annotationID As String, ByVal pathID As String, ByVal segmentID As String,
' Optional ByVal parameter As Double = 0.5, Optional ByVal name As String = "D", Optional ByVal comment As
' String = "")
' ByVal annotationID As String — Непустой уникальный ID аннотации.
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal segmentID As String — Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется
' индексом.
' Optional ByVal parameter As Double = 0.5 — Нормированный параметр точки сегмента [0;1], не расстояние по
' всему пути.
' Optional ByVal name As String = "D" — Обозначение измерения или имя контейнера.
' Optional ByVal comment As String = "" — Пояснение происхождения/инженерного смысла; не управляющий
' статус.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.Text
' Объявляет простой текст на именованной модельной привязке; пиксели/шрифт задаёт общий layout.
' Public Sub Text(ByVal annotationID As String, ByVal anchorID As String, ByVal textValue As String,
' Optional ByVal comment As String = "")
' ByVal annotationID As String — Непустой уникальный ID аннотации.
' ByVal anchorID As String — Непустой уникальный ID модельной привязки.
' ByVal textValue As String — Текст подписи; Unicode сохраняется.
' Optional ByVal comment As String = "" — Пояснение происхождения/инженерного смысла; не управляющий
' статус.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.RebarLabel
' Привязывает фактическую подпись группы к габариту target-пути, со стороны normalX/Y.
' Public Sub RebarLabel(ByVal groupID As String, ByVal target As String, ByVal normalX As Double, ByVal
' normalY As Double, Optional ByVal axisDistance As Double = 0#)
' ByVal groupID As String — Смысловой ID группы; путь и арматурная annotation-group — разные контракты.
' ByVal target As String — ID пути/PathGroup, разрешаемый в этом определении.
' ByVal normalX As Double — X-компонента стороны размерной/подписной линии; layout нормирует вектор.
' ByVal normalY As Double — Y-компонента стороны размерной/подписной линии.
' Optional ByVal axisDistance As Double = 0# — Положительное расстояние до оси, мм; не защитный слой до
' поверхности и не продольный setback.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.RebarRowLabel
' Привязывает фактическую подпись группы к концам исполненного ruleID.
' Public Sub RebarRowLabel(ByVal groupID As String, ByVal ruleID As String, ByVal normalX As Double, ByVal
' normalY As Double, Optional ByVal axisDistance As Double = 0#)
' ByVal groupID As String — Смысловой ID группы; путь и арматурная annotation-group — разные контракты.
' ByVal ruleID As String — Непустой уникальный ID правила; при диагностике — ID ранее исполненного
' правила.
' ByVal normalX As Double — X-компонента стороны размерной/подписной линии; layout нормирует вектор.
' ByVal normalY As Double — Y-компонента стороны размерной/подписной линии.
' Optional ByVal axisDistance As Double = 0# — Положительное расстояние до оси, мм; не защитный слой до
' поверхности и не продольный setback.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.RebarRowExtentLabel
' Привязывает подпись к фактическому габариту позиций правила; reverseDirection меняет ориентацию
' anchor-линии.
' Public Sub RebarRowExtentLabel(ByVal groupID As String, ByVal ruleID As String, ByVal normalX As Double,
' ByVal normalY As Double, Optional ByVal axisDistance As Double = 0#, Optional ByVal reverseDirection As
' Boolean = False)
' ByVal groupID As String — Смысловой ID группы; путь и арматурная annotation-group — разные контракты.
' ByVal ruleID As String — Непустой уникальный ID правила; при диагностике — ID ранее исполненного
' правила.
' ByVal normalX As Double — X-компонента стороны размерной/подписной линии; layout нормирует вектор.
' ByVal normalY As Double — Y-компонента стороны размерной/подписной линии.
' Optional ByVal axisDistance As Double = 0# — Положительное расстояние до оси, мм; не защитный слой до
' поверхности и не продольный setback.
' Optional ByVal reverseDirection As Boolean = False — Обращает направление линии привязки подписи;
' физический ряд не меняется.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.Finish
' Проверяет контуры, ссылки, зависимости и аннотации; публикует неизменяемое CSectionDefinition. Фасад
' после этого закрыт.
' Public Function Finish() As CSectionDefinition
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinitionBuilder.RebarRowExtremeLabel
' Выбирает крайние созданные позиции в directionX/Y; сторона текста normalX/Y независима от выбора ряда.
' Public Sub RebarRowExtremeLabel(ByVal groupID As String, ByVal ruleID As String, ByVal directionX As
' Double, ByVal directionY As Double, ByVal normalX As Double, ByVal normalY As Double, Optional ByVal
' axisDistance As Double = 0#)
' ByVal groupID As String — Смысловой ID группы; путь и арматурная annotation-group — разные контракты.
' ByVal ruleID As String — Непустой уникальный ID правила; при диагностике — ID ранее исполненного
' правила.
' ByVal directionX As Double — X-компонента направления проекции/выбора экстремума; вектор не нулевой, у
' RebarRowExtremeLabel единичный.
' ByVal directionY As Double — Y-компонента направления проекции/выбора экстремума.
' ByVal normalX As Double — X-компонента стороны размерной/подписной линии; layout нормирует вектор.
' ByVal normalY As Double — Y-компонента стороны размерной/подписной линии.
' Optional ByVal axisDistance As Double = 0# — Положительное расстояние до оси, мм; не защитный слой до
' поверхности и не продольный setback.
' Код метода или объявление: CSectionDefinitionBuilder в окне проекта этой книги (Ctrl+R).
' ========================================================================================================
' CGeometryPath
' ========================================================================================================
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.LastDiagnostic
' Последняя машинная диагностика; Nothing при отсутствии записанного события.
' Public Property Get LastDiagnostic() As CGeometryDiagnostic
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.Initialize
' Создаёт mutable путь с непустым ID один раз.
' Public Sub Initialize(ByVal pathID As String)
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.AddLine
' Добавляет прямой сегмент с устойчивым ID; начало совпадает с предыдущим концом.
' Public Sub AddLine(ByVal segmentID As String, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As
' Double, ByVal y2 As Double)
' ByVal segmentID As String — Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется
' индексом.
' ByVal x1 As Double — Модельная координата x1, мм; знаки соответствуют section XY/WCS.
' ByVal y1 As Double — Модельная координата y1, мм; знаки соответствуют section XY/WCS.
' ByVal x2 As Double — Модельная координата x2, мм; знаки соответствуют section XY/WCS.
' ByVal y2 As Double — Модельная координата y2, мм; знаки соответствуют section XY/WCS.
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.AddArc
' Добавляет аналитическую дугу по центру/R/startAngle/sweep, без хордовой аппроксимации.
' Public Sub AddArc(ByVal segmentID As String, ByVal centerX As Double, ByVal centerY As Double, ByVal
' radius As Double, ByVal startAngle As Double, ByVal sweep As Double)
' ByVal segmentID As String — Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется
' индексом.
' ByVal centerX As Double — Модельная координата centerX, мм; знаки соответствуют section XY/WCS.
' ByVal centerY As Double — Модельная координата centerY, мм; знаки соответствуют section XY/WCS.
' ByVal radius As Double — Положительный радиус, мм.
' ByVal startAngle As Double — Начальный угол, радианы; от +X.
' ByVal sweep As Double — Signed sweep, радианы: CCW положителен, CW отрицателен, 0 запрещён,
' |sweep|<=2*pi.
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.Seal
' Публикует путь. closed=True проверяет совпадение концов, не достраивает соединение.
' Public Sub Seal(Optional ByVal closed As Boolean = False)
' Optional ByVal closed As Boolean = False — Явный Boolean режима замыкания/касательной; не исправляет
' геометрию автоматически.
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.ID
' Смысловой ID этого пути.
' Public Property Get ID() As String
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.SegmentCount
' Число аналитических line/arc сегментов.
' Public Property Get SegmentCount() As Long
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.IsSealed
' Опубликован ли объект; published геометрия не изменяется.
' Public Property Get IsSealed() As Boolean
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.IsClosed
' Замкнут ли фактический путь, без фиктивного соединения.
' Public Property Get IsClosed() As Boolean
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.Length
' Точная суммарная длина пути, мм.
' Public Property Get Length() As Double
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.SegmentID
' Возвращает смысловой ID сегмента по индексу.
' Public Property Get SegmentID(ByVal index As Long) As String
' ByVal index As Long — 1-based индекс в пределах соответствующего готового контейнера.
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.SourcePathID
' ID исходного пути сегмента после slice/transform/offset.
' Public Property Get SourcePathID(ByVal index As Long) As String
' ByVal index As Long — 1-based индекс в пределах соответствующего готового контейнера.
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.SourceSegmentID
' Смысловой исходный ID сегмента, сохраняемый при распаде/переименовании результата.
' Public Property Get SourceSegmentID(ByVal index As Long) As String
' ByVal index As Long — 1-based индекс в пределах соответствующего готового контейнера.
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.SourceStartParameter
' Начальный параметр исходного сегмента; может выйти за [0;1] при явном продлении прямой.
' Public Property Get SourceStartParameter(ByVal index As Long) As Double
' ByVal index As Long — 1-based индекс в пределах соответствующего готового контейнера.
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.SourceEndParameter
' Конечный параметр исходного сегмента, включая обращённый обход.
' Public Property Get SourceEndParameter(ByVal index As Long) As Double
' ByVal index As Long — 1-based индекс в пределах соответствующего готового контейнера.
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.SegmentIndex
' Разрешает существующий смысловой ID в текущий индекс; исчезнувшая дуга даёт MISSING_SEGMENT.
' Public Function SegmentIndex(ByVal segmentID As String) As Long
' ByVal segmentID As String — Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется
' индексом.
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.AnchorStation
' Находит расстояние от начала всей линии до точки заданного сегмента, мм. Точка задаётся именем сегмента
' и параметром [0;1]. Если при построении участок исчез, обе его границы сохраняют привязку к оставшемуся
' стыку. Неизвестное имя вызывает ошибку.
' Public Function AnchorStation(ByVal segmentID As String, Optional ByVal parameter As Double = 0#) As
' Double
' ByVal segmentID As String — Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется
' индексом.
' Optional ByVal parameter As Double = 0# — Нормированный параметр точки сегмента [0;1], не расстояние по
' всему пути.
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.WithID
' Копирует опубликованный путь с новым ID, сохраняя сегментные ссылки и исходное происхождение.
' Public Function WithID(ByVal pathID As String) As CGeometryPath
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.BetweenAnchors
' Возвращает точную цепочку между привязками с явным направлением и правилом полного обхода.
' Public Function BetweenAnchors(ByVal pathID As String, ByVal startSegmentID As String, ByVal
' startParameter As Double, ByVal endSegmentID As String, ByVal endParameter As Double, Optional ByVal
' forward As Boolean = True, Optional ByVal fullLoopIfSame As Boolean = False) As CGeometryPath
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal startSegmentID As String — Сегмент начальной привязки цепочки.
' ByVal startParameter As Double — Параметр начальной точки выбранного сегмента [0;1].
' ByVal endSegmentID As String — Сегмент конечной привязки цепочки.
' ByVal endParameter As Double — Параметр конечной точки выбранного сегмента [0;1].
' Optional ByVal forward As Boolean = True — Явное направление следования исходному обходу; False —
' противоположное.
' Optional ByVal fullLoopIfSame As Boolean = False — Полный оборот при совпавших привязках замкнутого
' пути; иначе нулевая цепочка недопустима.
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.AtDistance
' Возвращает координаты точки на заданном расстоянии от начала линии и единичный вектор касательной в этой
' точке. На стыке можно выбрать касательную предыдущего сегмента.
' Public Sub AtDistance(ByVal distance As Double, ByRef x As Double, ByRef y As Double, ByRef tangentX As
' Double, ByRef tangentY As Double, Optional ByVal incomingTangentAtJoint As Boolean = False)
' ByVal distance As Double — Расстояние или граница интервала distance, внутренние мм; направление и
' допустимые пределы определяет вызываемый метод.
' ByRef x As Double — Модельная координата x, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые
' скаляры/буфер можно изменять, исходный immutable snapshot не меняется.
' ByRef y As Double — Модельная координата y, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые
' скаляры/буфер можно изменять, исходный immutable snapshot не меняется.
' ByRef tangentX As Double — Компонента единичной касательной, без размерности. ByRef: возвращаемые
' скаляры/буфер можно изменять, исходный immutable snapshot не меняется.
' ByRef tangentY As Double — Компонента единичной касательной, без размерности. ByRef: возвращаемые
' скаляры/буфер можно изменять, исходный immutable snapshot не меняется.
' Optional ByVal incomingTangentAtJoint As Boolean = False — Явный Boolean режима замыкания/касательной;
' не исправляет геометрию автоматически.
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.ProjectionBounds
' Возвращает действительные min/max проекции по нормированному направлению, с экстремумами дуг.
' Public Sub ProjectionBounds(ByVal directionX As Double, ByVal directionY As Double, ByRef minimum As
' Double, ByRef maximum As Double)
' ByVal directionX As Double — X-компонента направления проекции/выбора экстремума; вектор не нулевой, у
' RebarRowExtremeLabel единичный.
' ByVal directionY As Double — Y-компонента направления проекции/выбора экстремума.
' ByRef minimum As Double — Расстояние или граница интервала minimum, внутренние мм; направление и
' допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять,
' исходный immutable snapshot не меняется.
' ByRef maximum As Double — Расстояние или граница интервала maximum, внутренние мм; направление и
' допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять,
' исходный immutable snapshot не меняется.
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.Slice
' Обрезает полный непрерывный путь по длине; точные дуги и source parameters сохраняются.
' Public Function Slice(ByVal startDistance As Double, ByVal endDistance As Double) As CGeometryPath
' ByVal startDistance As Double — Расстояние от начала всей линии, мм, в [0;Length]; endDistance строго
' больше startDistance.
' ByVal endDistance As Double — Расстояние от начала всей линии, мм, в [0;Length]; endDistance строго
' больше startDistance.
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.Reversed
' Копирует путь с обратным порядком сегментов/обходом; signed sweep и source parameters обращаются.
' Public Function Reversed() As CGeometryPath
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.Transformed
' Копирует путь изометрией: отражение по локальной X, поворот, перенос; длина сохраняется.
' Public Function Transformed(ByVal shiftX As Double, ByVal shiftY As Double, Optional ByVal rotation As
' Double = 0#, Optional ByVal reflectX As Boolean = False) As CGeometryPath
' ByVal shiftX As Double — Перенос по X, мм.
' ByVal shiftY As Double — Перенос по Y, мм.
' Optional ByVal rotation As Double = 0# — Угол поворота вокруг (0,0), радианы, CCW положителен.
' Optional ByVal reflectX As Boolean = False — Отражение относительно локальной оси X перед поворотом и
' переносом.
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryPath.RayHit
' Находит ближайшее положительное пересечение луча с конечным путём; False не экстраполирует линию.
' Public Function RayHit(ByVal originX As Double, ByVal originY As Double, ByVal directionX As Double,
' ByVal directionY As Double, ByRef pathDistance As Double, ByRef rayDistance As Double) As Boolean
' ByVal originX As Double — Модельная координата originX, мм; знаки соответствуют section XY/WCS.
' ByVal originY As Double — Модельная координата originY, мм; знаки соответствуют section XY/WCS.
' ByVal directionX As Double — X-компонента направления проекции/выбора экстремума; вектор не нулевой, у
' RebarRowExtremeLabel единичный.
' ByVal directionY As Double — Y-компонента направления проекции/выбора экстремума.
' ByRef pathDistance As Double — Расстояние или граница интервала pathDistance, внутренние мм; направление
' и допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять,
' исходный immutable snapshot не меняется.
' ByRef rayDistance As Double — Расстояние или граница интервала rayDistance, внутренние мм; направление и
' допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять,
' исходный immutable snapshot не меняется.
' Код метода или объявление: CGeometryPath в окне проекта этой книги (Ctrl+R).
' ========================================================================================================
' CSectionDefinition
' ========================================================================================================
' --------------------------------------------------------------------------------------------------------
' CSectionDefinition.LastDiagnostic
' Последняя машинная диагностика; Nothing при отсутствии записанного события.
' Public Property Get LastDiagnostic() As CGeometryDiagnostic
' Код метода или объявление: CSectionDefinition в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinition.AnchorCoordinates
' Возвращает координаты именованной привязки в скалярные переменные вызывающего кода.
' Public Sub AnchorCoordinates(ByVal anchorID As String, ByRef x As Double, ByRef y As Double)
' ByVal anchorID As String — Непустой уникальный ID модельной привязки.
' ByRef x As Double — Модельная координата x, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые
' скаляры/буфер можно изменять, исходный immutable snapshot не меняется.
' ByRef y As Double — Модельная координата y, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые
' скаляры/буфер можно изменять, исходный immutable snapshot не меняется.
' Код метода или объявление: CSectionDefinition в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinition.SourceReference
' Читает скалярный SourceMap по ID; отсутствующая связь возвращает пустую строку.
' Public Function SourceReference(ByVal entityID As String) As String
' ByVal entityID As String — Смысловой ID сущности для SourceMap.
' Код метода или объявление: CSectionDefinition в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinition.TypeID
' Смысловое имя типа опубликованного определения.
' Public Property Get TypeID() As String
' Код метода или объявление: CSectionDefinition в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinition.SchemaVersion
' Положительная версия схемы определения.
' Public Property Get SchemaVersion() As Long
' Код метода или объявление: CSectionDefinition в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinition.IsSealed
' Опубликован ли объект; published геометрия не изменяется.
' Public Property Get IsSealed() As Boolean
' Код метода или объявление: CSectionDefinition в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinition.PathCount
' Число зарегистрированных путей.
' Public Property Get PathCount() As Long
' Код метода или объявление: CSectionDefinition в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinition.LoopCount
' Число подготовленных материальных loops.
' Public Property Get LoopCount() As Long
' Код метода или объявление: CSectionDefinition в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinition.RuleCount
' Число объявленных правил, включая выключенные.
' Public Property Get RuleCount() As Long
' Код метода или объявление: CSectionDefinition в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinition.AnnotationCount
' Число объявленных annotation rules.
' Public Property Get AnnotationCount() As Long
' Код метода или объявление: CSectionDefinition в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinition.Path
' Читает опубликованный неизменяемый путь по ID.
' Public Function Path(ByVal pathID As String) As CGeometryPath
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' Код метода или объявление: CSectionDefinition в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinition.TargetPaths
' Возвращает ID одного пути или массив независимых путей группы.
' Public Function TargetPaths(ByVal targetID As String) As Variant
' ByVal targetID As String — ID пути или группы путей.
' Код метода или объявление: CSectionDefinition в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinition.ConcreteDomain
' Возвращает clone подготовленной бетонной области, без Excel или повторной интерпретации рецепта.
' Public Function ConcreteDomain() As CConcreteRegion
' Код метода или объявление: CSectionDefinition в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinition.MaterialContours
' Возвращает отдельный контейнер аналитических материальных Outer/Opening.
' Public Function MaterialContours() As CSectionContours
' Код метода или объявление: CSectionDefinition в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionDefinition.WriteContours
' Записывает единственный контурный snapshot в переданный контейнер; destination предварительно очищается.
' Public Sub WriteContours(ByVal result As CSectionContours)
' ByVal result As CSectionContours — Подготовленный объект/контейнер выбранного запроса, без повторного
' чтения Excel или назначения типа.
' Код метода или объявление: CSectionDefinition в окне проекта этой книги (Ctrl+R).
' ========================================================================================================
' CRebarRuleOptions
' ========================================================================================================
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.StartSetback
' Начальное отсечение вдоль каждого target-пути, мм>=0; у полного ClosedUniform 0.
' Public StartSetback As Double
' Default: 0.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.EndSetback
' Конечное отсечение вдоль каждого target-пути, мм>=0; ProjectFrom фильтрует hit без перераспределения.
' Public EndSetback As Double
' Default: 0.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.EndMode
' IncludeBoth/ExcludeBoth/IncludeStart/IncludeEnd/ClosedUniform; последние требуют соответствующей
' замкнутости.
' Public EndMode As ERebarEndMode
' Default: IncludeBoth.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.DistributionScope
' Допустимо только PerPath: число стержней, расстояния между ними и отступы от концов применяются отдельно
' к каждой независимой линии группы.
' Public DistributionScope As ERebarDistributionScope
' Default: PerPath.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.RemainderPlacement
' RemainderStart/End/Center: положение остатка для точного шага.
' Public RemainderPlacement As ERemainderPlacement
' Default: RemainderCenter.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.TangentOffset
' Локальное связанное смещение вдоль касательной, мм; signed, не независимый шаг ряда.
' Public TangentOffset As Double
' Default: 0.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.NormalOffset
' Локальное связанное смещение по левой нормали (-ty,tx), мм; signed.
' Public NormalOffset As Double
' Default: 0.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.EverySecond
' Брать локальные исходные позиции 1, 3, ... отдельно на каждом пути; SuccessfulProjectionOrdinal задаёт
' отдельный отсчёт успешных проекций.
' Public EverySecond As Boolean
' Default: False.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.IncomingTangentAtJoint
' В точке соединения двух сегментов брать касательную предыдущего сегмента. По умолчанию берётся
' касательная следующего сегмента.
' Public IncomingTangentAtJoint As Boolean
' Default: False.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.DirectionX
' Глобальная X направления ProjectFrom, используется при SourceNormalSign=0.
' Public DirectionX As Double
' Default: 0.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.DirectionY
' Глобальная Y направления ProjectFrom; направление не нулевое.
' Public DirectionY As Double
' Default: 0.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.SourceNormalSign
' -1/0/+1: локальная нормаль исходного ряда или глобальное DirectionX/Y при 0.
' Public SourceNormalSign As Double
' Default: 0.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.SuccessfulProjectionOrdinal
' True: заново нумеровать только стержни, для которых найдено пересечение с целевой линией, подряд по всей
' группе. False: сохранять исходные номера стержней с возможными пропусками отдельно на каждой линии.
' Public SuccessfulProjectionOrdinal As Boolean
' Default: False.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.NewPublicationRoot
' У ProjectFrom начинает новый корень публикации по успешным позициям.
' Public NewPublicationRoot As Boolean
' Default: False.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.SkipMissingProjection
' Разрешить пропуск без hit с PROJECTION_MISSING; False выдаёт ошибку.
' Public SkipMissingProjection As Boolean
' Default: False.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.Publication
' PublishByRule или PublishPerSourcePosition; влияет на вывод, не геометрический шаг.
' Public Publication As ERebarPublication
' Default: PublishByRule.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.PositionOrder
' Порядок вывода стержней: AlongPath — от начала линии, FromPathMiddle — от её середины. Места стержней и
' их исходные номера сохраняются.
' Public PositionOrder As ERebarPositionOrder
' Default: AlongPath.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.Coincidence
' RejectCoincidence/MergeIdentical/KeepPublishedOwner; последнее сохраняет явно выбранного опубликованного
' владельца.
' Public Coincidence As ERebarCoincidence
' Default: RejectCoincidence.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.PlacementCheck
' CheckCenter или CheckFullDisk относительно Outer/Opening; ужесточение задаётся явно.
' Public PlacementCheck As ERebarPlacementCheck
' Default: CheckCenter.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.OverlapPolicy
' ReportOverlap или RejectOverlap. Контакт допустим; нормативный минимальный зазор не вводится.
' Public OverlapPolicy As ERebarOverlapPolicy
' Default: ReportOverlap.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.SourcePrefix
' Префикс SourcePositionId; пустой использует RuleId+'.' либо связь source+suffix.
' Public SourcePrefix As String
' Default: пустое значение VBA.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.SourceSuffix
' Суффикс SourcePositionId, в том числе для связанных рядов.
' Public SourceSuffix As String
' Default: пустое значение VBA.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.AnnotationGroup
' Смысловая принадлежность подписи физического стержня; не расчётная группа трещин СП35.
' Public AnnotationGroup As String
' Default: пустое значение VBA.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.Comment
' Инженерное пояснение правила; не selector или ResultCode.
' Public Comment As String
' Default: пустое значение VBA.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleOptions.SourceReference
' Реальные адреса исходных параметров правила или программная ссылка.
' Public SourceReference As String
' Default: пустое значение VBA.
' Код метода или объявление: CRebarRuleOptions в окне проекта этой книги (Ctrl+R).
' ========================================================================================================
' CGeometryDiagnostic
' ========================================================================================================
' --------------------------------------------------------------------------------------------------------
' CGeometryDiagnostic.Code
' Машинный код причины; назначается напрямую, не извлекается из русского текста.
' Public Property Get Code() As String
' Код метода или объявление: CGeometryDiagnostic в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryDiagnostic.EntityID
' Смысловой ID сущности отказа, если он известен этому сервису.
' Public Property Get EntityID() As String
' Код метода или объявление: CGeometryDiagnostic в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryDiagnostic.RelatedEntityID
' ID связанного сегмента/позиции/стержня для локализации причины.
' Public Property Get RelatedEntityID() As String
' Код метода или объявление: CGeometryDiagnostic в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryDiagnostic.Reason
' Человекочитаемая причина; не selector/status.
' Public Property Get Reason() As String
' Код метода или объявление: CGeometryDiagnostic в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryDiagnostic.Detail1
' Первое численное поле конкретного Code; единицы определяет Code.
' Public Property Get Detail1() As Double
' Код метода или объявление: CGeometryDiagnostic в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryDiagnostic.Detail2
' Второе численное поле конкретного Code.
' Public Property Get Detail2() As Double
' Код метода или объявление: CGeometryDiagnostic в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CGeometryDiagnostic.SourceReference
' Реальный адрес/программный источник, добавленный SourceMap или правилом.
' Public Property Get SourceReference() As String
' Код метода или объявление: CGeometryDiagnostic в окне проекта этой книги (Ctrl+R).
' ========================================================================================================
' CSectionInputTable
' ========================================================================================================
' --------------------------------------------------------------------------------------------------------
' CSectionInputTable.Initialize
' Копирует двумерные scalar values и согласованные addresses один раз; Excel-объекты не сохраняет.
' Public Sub Initialize(ByVal tableName As String, ByVal values As Variant, Optional ByVal addresses As
' Variant)
' ByVal tableName As String — Имя единственного корневого входного блока этого снимка.
' ByVal values As Variant — Двумерный массив скалярных значений, копируется; вложенные массивы/объекты
' запрещены.
' Optional ByVal addresses As Variant — Необязательный согласованный 2D-массив адресов; без него
' применяется программный tableName[row,col].
' Код метода или объявление: CSectionInputTable в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionInputTable.Name
' Имя корневого блока snapshot.
' Public Property Get Name() As String
' Код метода или объявление: CSectionInputTable в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionInputTable.RowCount
' Число строк неизменяемого input snapshot.
' Public Property Get RowCount() As Long
' Код метода или объявление: CSectionInputTable в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionInputTable.ColumnCount
' Число столбцов input snapshot.
' Public Property Get ColumnCount() As Long
' Код метода или объявление: CSectionInputTable в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionInputTable.SourceReference
' Возвращает исходный адрес ячейки или программный tableName[row,col].
' Public Function SourceReference(ByVal row As Long, ByVal column As Long) As String
' ByVal row As Long — Номер строки снимка 1..RowCount.
' ByVal column As Long — Номер столбца снимка 1..ColumnCount.
' Код метода или объявление: CSectionInputTable в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionInputTable.CellValue
' Читает scalar snapshot; ошибка Excel в ячейке даёт адресное исключение.
' Public Function CellValue(ByVal row As Long, ByVal column As Long) As Variant
' ByVal row As Long — Номер строки снимка 1..RowCount.
' ByVal column As Long — Номер столбца снимка 1..ColumnCount.
' Код метода или объявление: CSectionInputTable в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionInputTable.IsBlank
' Проверяет пустое значение снимка; ошибка Excel не подменяется пустотой.
' Public Function IsBlank(ByVal row As Long, ByVal column As Long) As Boolean
' ByVal row As Long — Номер строки снимка 1..RowCount.
' ByVal column As Long — Номер столбца снимка 1..ColumnCount.
' Код метода или объявление: CSectionInputTable в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionInputTable.Text
' Читает непустой текст либо явно переданный fallback, обрезая крайние пробелы.
' Public Function Text(ByVal row As Long, ByVal column As Long, Optional ByVal defaultValue As Variant) As
' String
' ByVal row As Long — Номер строки снимка 1..RowCount.
' ByVal column As Long — Номер столбца снимка 1..ColumnCount.
' Optional ByVal defaultValue As Variant — Явный fallback конкретного вызова. Пропущенный Optional Variant
' без default означает обязательное поле.
' Код метода или объявление: CSectionInputTable в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionInputTable.Number
' Читает число, исключая Boolean/ошибки/объекты; обязательное пустое поле не заменяется нулём.
' Public Function Number(ByVal row As Long, ByVal column As Long, Optional ByVal defaultValue As Variant)
' As Double
' ByVal row As Long — Номер строки снимка 1..RowCount.
' ByVal column As Long — Номер столбца снимка 1..ColumnCount.
' Optional ByVal defaultValue As Variant — Явный fallback конкретного вызова. Пропущенный Optional Variant
' без default означает обязательное поле.
' Код метода или объявление: CSectionInputTable в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionInputTable.Count
' Читает целое неотрицательное 0..2147483647; отдельный executor проверяет суммарное число позиций.
' Public Function Count(ByVal row As Long, ByVal column As Long, Optional ByVal defaultValue As Long = 0)
' As Long
' ByVal row As Long — Номер строки снимка 1..RowCount.
' ByVal column As Long — Номер столбца снимка 1..ColumnCount.
' Optional ByVal defaultValue As Long = 0 — Явный fallback конкретного вызова. Пропущенный Optional
' Variant без default означает обязательное поле.
' Код метода или объявление: CSectionInputTable в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionInputTable.Choice
' Проверяет вариант по allowed без учёта регистра и возвращает каноническое значение allowed.
' Public Function Choice(ByVal row As Long, ByVal column As Long, ByVal allowed As Variant, Optional ByVal
' defaultValue As String = vbNullString) As String
' ByVal row As Long — Номер строки снимка 1..RowCount.
' ByVal column As Long — Номер столбца снимка 1..ColumnCount.
' ByVal allowed As Variant — Массив разрешённых строк; Choice возвращает каноническую строку из этого
' массива.
' Optional ByVal defaultValue As String = vbNullString — Явный fallback конкретного вызова. Пропущенный
' Optional Variant без default означает обязательное поле.
' Код метода или объявление: CSectionInputTable в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionInputTable.Length
' Читает обязательное число и один раз переводит input length в мм через CUnitSystem.
' Public Function Length(ByVal row As Long, ByVal column As Long, ByVal units As CUnitSystem) As Double
' ByVal row As Long — Номер строки снимка 1..RowCount.
' ByVal column As Long — Номер столбца снимка 1..ColumnCount.
' ByVal units As CUnitSystem — CUnitSystem, уже подготовленная вызывающим workflow; преобразование только
' через её методы.
' Код метода или объявление: CSectionInputTable в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionInputTable.OptionalDiameter
' Пустая ячейка или 0 отключают ряд; положительный диаметр переводится в мм. Текст, включая одиночный
' дефис и TODO, и отрицательный диаметр вызывают адресную ошибку ввода.
' Public Function OptionalDiameter(ByVal row As Long, ByVal column As Long, ByVal units As CUnitSystem) As
' Double
' ByVal row As Long — Номер строки снимка 1..RowCount.
' ByVal column As Long — Номер столбца снимка 1..ColumnCount.
' ByVal units As CUnitSystem — CUnitSystem, уже подготовленная вызывающим workflow; преобразование только
' через её методы.
' Код метода или объявление: CSectionInputTable в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionInputTable.ColumnByCaption
' Находит уникальный столбец заголовка, без фиксированного адреса.
' Public Function ColumnByCaption(ByVal headerRow As Long, ByVal caption As String) As Long
' ByVal headerRow As Long — Строка заголовков, в пределах снимка.
' ByVal caption As String — Смысловой заголовок столбца; повторные заголовки — ошибка.
' Код метода или объявление: CSectionInputTable в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionInputTable.RowByKey
' Находит уникальную строку смыслового ключа в указанном подблоке.
' Public Function RowByKey(ByVal key As String, Optional ByVal keyColumn As Long = 1, Optional ByVal
' firstRow As Long = 1, Optional ByVal lastRow As Long = 0) As Long
' ByVal key As String — Смысловой ключ строки, не Excel-адрес.
' Optional ByVal keyColumn As Long = 1 — Столбец смысловых ключей 1..ColumnCount.
' Optional ByVal firstRow As Long = 1 — Начало диапазона поиска строки, включительно.
' Optional ByVal lastRow As Long = 0 — Конец поиска, включительно; 0 означает RowCount.
' Код метода или объявление: CSectionInputTable в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionInputTable.RowByAnyKey
' Находит уникальную строку по массиву допустимых альтернативных ключей; неоднозначность запрещена.
' Public Function RowByAnyKey(ByVal keys As Variant, Optional ByVal keyColumn As Long = 1, Optional ByVal
' firstRow As Long = 1, Optional ByVal lastRow As Long = 0) As Long
' ByVal keys As Variant — Массив допустимых альтернативных ключей; неоднозначное совпадение отклоняется.
' Optional ByVal keyColumn As Long = 1 — Столбец смысловых ключей 1..ColumnCount.
' Optional ByVal firstRow As Long = 1 — Начало диапазона поиска строки, включительно.
' Optional ByVal lastRow As Long = 0 — Конец поиска, включительно; 0 означает RowCount.
' Код метода или объявление: CSectionInputTable в окне проекта этой книги (Ctrl+R).
' ========================================================================================================
' CSectionRecipeContext
' ========================================================================================================
' --------------------------------------------------------------------------------------------------------
' CSectionRecipeContext.Initialize
' Открывает контекст с подготовленной системой единиц один раз.
' Public Sub Initialize(ByVal units As CUnitSystem)
' ByVal units As CUnitSystem — CUnitSystem, уже подготовленная вызывающим workflow; преобразование только
' через её методы.
' Код метода или объявление: CSectionRecipeContext в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionRecipeContext.AddTable
' Добавляет immutable input snapshot с уникальным именем до Seal.
' Public Sub AddTable(ByVal table As CSectionInputTable)
' ByVal table As CSectionInputTable — Неизменяемый CSectionInputTable.
' Код метода или объявление: CSectionRecipeContext в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionRecipeContext.Seal
' Закрывает состав таблиц; после этого добавление невозможно.
' Public Sub Seal()
' Код метода или объявление: CSectionRecipeContext в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionRecipeContext.Table
' Возвращает снимок зарегистрированного блока по имени; отсутствие — ошибка, не fallback.
' Public Function Table(ByVal tableName As String) As CSectionInputTable
' ByVal tableName As String — Имя единственного корневого входного блока этого снимка.
' Код метода или объявление: CSectionRecipeContext в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionRecipeContext.Units
' Система единиц внешнего workflow; не перезагружайте её из Excel в рецепте.
' Public Property Get Units() As CUnitSystem
' Код метода или объявление: CSectionRecipeContext в окне проекта этой книги (Ctrl+R).
' ========================================================================================================
' CSectionInputSnapshotReader
' ========================================================================================================
' --------------------------------------------------------------------------------------------------------
' CSectionInputSnapshotReader.Capture
' На внешней Excel-границе вычисляет и копирует только выбранные именованные блоки с адресами, затем Seal
' context.
' Public Function Capture(ByVal workbook As Object, ByVal tableNames As Variant, ByVal units As
' CUnitSystem) As CSectionRecipeContext
' ByVal workbook As Object — Excel workbook на внешней границе ввода; рецепт этот объект не хранит.
' ByVal tableNames As Variant — Массив имён блоков выбранного типа; Empty-array допустим для программного
' рецепта.
' ByVal units As CUnitSystem — CUnitSystem, уже подготовленная вызывающим workflow; преобразование только
' через её методы.
' Код метода или объявление: CSectionInputSnapshotReader в окне проекта этой книги (Ctrl+R).
' ========================================================================================================
' CSectionTypeRegistry
' ========================================================================================================
' --------------------------------------------------------------------------------------------------------
' CSectionTypeRegistry.RegisterRecipe
' Регистрирует один ISectionRecipe и массив имён его входных блоков; ID типа уникальный.
' Public Sub RegisterRecipe(ByVal typeID As String, ByVal recipe As ISectionRecipe, ByVal tableNames As
' Variant)
' ByVal typeID As String — Смысловое имя зарегистрированного типа; непустой текст.
' ByVal recipe As ISectionRecipe — Один объект с реализацией ISectionRecipe.
' ByVal tableNames As Variant — Массив имён блоков выбранного типа; Empty-array допустим для программного
' рецепта.
' Код метода или объявление: CSectionTypeRegistry в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionTypeRegistry.CreateRecipe
' Возвращает зарегистрированный объект по TypeID; неизвестный тип — ошибка.
' Public Function CreateRecipe(ByVal typeID As String) As ISectionRecipe
' ByVal typeID As String — Смысловое имя зарегистрированного типа; непустой текст.
' Код метода или объявление: CSectionTypeRegistry в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionTypeRegistry.RecipeTableNames
' Возвращает независимый массив имён ввода, пустой для программного рецепта.
' Public Function RecipeTableNames(ByVal typeID As String) As Variant
' ByVal typeID As String — Смысловое имя зарегистрированного типа; непустой текст.
' Код метода или объявление: CSectionTypeRegistry в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionTypeRegistry.RegisteredRecipeTypes
' Возвращает фактически зарегистрированные TypeID — источник штатного списка Config.
' Public Function RegisteredRecipeTypes() As Variant
' Код метода или объявление: CSectionTypeRegistry в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionTypeRegistry.SelectedRecipeType
' Читает обязательный Geometry.Type через settings и проверяет его по реестру.
' Public Function SelectedRecipeType(ByVal settings As CSystemSettingsReader) As String
' ByVal settings As CSystemSettingsReader — CSystemSettingsReader, уже загрузивший Config.
' Код метода или объявление: CSectionTypeRegistry в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionTypeRegistry.CreateDefinition
' Вызывает ровно один выбранный ISectionRecipe с подготовленным context.
' Public Function CreateDefinition(ByVal typeID As String, ByVal context As CSectionRecipeContext) As
' CSectionDefinition
' ByVal typeID As String — Смысловое имя зарегистрированного типа; непустой текст.
' ByVal context As CSectionRecipeContext — Подготовленный снимок CSectionRecipeContext, не live Excel.
' Код метода или объявление: CSectionTypeRegistry в окне проекта этой книги (Ctrl+R).
' ========================================================================================================
' CSectionBuilder
' ========================================================================================================
' --------------------------------------------------------------------------------------------------------
' CSectionBuilder.Build
' Единая сборка definition -> область -> сетка/правила -> CSectionModel -> аннотации; не ветвится по типу.
' Public Function Build(ByVal definition As CSectionDefinition, ByRef meshOptions As TSectionMeshOptions)
' As CSectionModel
' ByVal definition As CSectionDefinition — Успешно опубликованный неизменяемый CSectionDefinition.
' ByRef meshOptions As TSectionMeshOptions — TSectionMeshOptions: StepX/StepY мм>0, BoundarySubdivisions
' целое>=1, MaterialID>=1. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot
' не меняется.
' Код метода или объявление: CSectionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionBuilder.Mesh
' CFiberMeshBuilder только последней успешно собранной модели; готовую сетку повторно не создавайте.
' Public Property Get Mesh() As CFiberMeshBuilder
' Код метода или объявление: CSectionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionBuilder.Rebars
' Физическая раскладка последней успешной сборки; совпадения уже обработаны явно.
' Public Property Get Rebars() As CRebarLayout
' Код метода или объявление: CSectionBuilder в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionBuilder.Rules
' Исполнитель последней успешной сборки с фактическими счётчиками/шагами/происхождением.
' Public Property Get Rules() As CRebarRuleExecutor
' Код метода или объявление: CSectionBuilder в окне проекта этой книги (Ctrl+R).
' ========================================================================================================
' CRebarRuleExecutor
' ========================================================================================================
' --------------------------------------------------------------------------------------------------------
' CRebarRuleExecutor.RequestedPositions
' Все запрошенные позиции, включая пропущенные проекции; это не число физических стержней.
' Public Property Get RequestedPositions() As Long
' Код метода или объявление: CRebarRuleExecutor в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleExecutor.CreatedPositions
' Действительно добавленные физические стержни текущего Build, даже при отказе в середине.
' Public Property Get CreatedPositions() As Long
' Код метода или объявление: CRebarRuleExecutor в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleExecutor.SkippedProjections
' Число явно разрешённых пропусков target hit.
' Public Property Get SkippedProjections() As Long
' Код метода или объявление: CRebarRuleExecutor в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleExecutor.MergedPositions
' Число заявленных позиций, объединённых с физическим владельцем.
' Public Property Get MergedPositions() As Long
' Код метода или объявление: CRebarRuleExecutor в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleExecutor.Conflicts
' Число несовместимых/отклонённых совпадений или запрещённых пересечений по явной политике.
' Public Property Get Conflicts() As Long
' Код метода или объявление: CRebarRuleExecutor в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleExecutor.BarOverlaps
' Полное число пар с отрицательным геометрическим зазором дисков; касание не включается.
' Public Property Get BarOverlaps() As Long
' Код метода или объявление: CRebarRuleExecutor в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleExecutor.DiagnosticRecordsOmitted
' Число событий сверх 128 сохранённых подробных образцов; полные счётчики не сокращаются.
' Public Property Get DiagnosticRecordsOmitted() As Long
' Код метода или объявление: CRebarRuleExecutor в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleExecutor.LastDiagnostic
' Последняя машинная диагностика; Nothing при отсутствии записанного события.
' Public Property Get LastDiagnostic() As CGeometryDiagnostic
' Код метода или объявление: CRebarRuleExecutor в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleExecutor.DiagnosticCount
' Число сохранённых подробных immutable образцов, 0..128.
' Public Property Get DiagnosticCount() As Long
' Код метода или объявление: CRebarRuleExecutor в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleExecutor.DiagnosticAt
' Читает immutable подробный образец 1..DiagnosticCount; максимум 128 образцов, счётчики полные.
' Public Function DiagnosticAt(ByVal index As Long) As CGeometryDiagnostic
' ByVal index As Long — 1-based индекс в пределах соответствующего готового контейнера.
' Код метода или объявление: CRebarRuleExecutor в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleExecutor.PositionCount
' Число сгенерированных заявленных позиций до объединения физических владельцев.
' Public Property Get PositionCount() As Long
' Код метода или объявление: CRebarRuleExecutor в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleExecutor.PositionRuleID
' ID правила заявленной позиции, включая объединённые позиции.
' Public Function PositionRuleID(ByVal index As Long) As String
' ByVal index As Long — 1-based индекс в пределах соответствующего готового контейнера.
' Код метода или объявление: CRebarRuleExecutor в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleExecutor.PositionPathID
' ID геометрического пути заявленной позиции.
' Public Function PositionPathID(ByVal index As Long) As String
' ByVal index As Long — 1-based индекс в пределах соответствующего готового контейнера.
' Код метода или объявление: CRebarRuleExecutor в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleExecutor.PositionSourceID
' Смысловой SourcePositionId, из которого публиковался физический стержень.
' Public Function PositionSourceID(ByVal index As Long) As String
' ByVal index As Long — 1-based индекс в пределах соответствующего готового контейнера.
' Код метода или объявление: CRebarRuleExecutor в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleExecutor.PositionPhysicalOwner
' Индекс физического владельца в итоговом CRebarLayout; объединённые позиции имеют одного владельца.
' Public Function PositionPhysicalOwner(ByVal index As Long) As Long
' ByVal index As Long — 1-based индекс в пределах соответствующего готового контейнера.
' Код метода или объявление: CRebarRuleExecutor в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleExecutor.RuleTrace
' Текстовый отчёт конкретного правила с фактическими counts и шагами/хордами.
' Public Function RuleTrace(ByVal ruleID As String) As String
' ByVal ruleID As String — Непустой уникальный ID правила; при диагностике — ID ранее исполненного
' правила.
' Код метода или объявление: CRebarRuleExecutor в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleExecutor.RuleStepBounds
' Возвращает число промежутков между стержнями, наименьшее и наибольшее расстояния между ними вдоль линии
' и по прямой, мм. Для замкнутого ряда учитывает промежуток от последнего стержня до первого.
' Public Sub RuleStepBounds(ByVal ruleID As String, ByRef stepCount As Long, ByRef minStep As Double,
' ByRef maxStep As Double, ByRef minChord As Double, ByRef maxChord As Double)
' ByVal ruleID As String — Непустой уникальный ID правила; при диагностике — ID ранее исполненного
' правила.
' ByRef stepCount As Long — Целый неотрицательный счётчик. ByRef: возвращаемые скаляры/буфер можно
' изменять, исходный immutable snapshot не меняется.
' ByRef minStep As Double — Расстояние или граница интервала minStep, внутренние мм; направление и
' допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять,
' исходный immutable snapshot не меняется.
' ByRef maxStep As Double — Расстояние или граница интервала maxStep, внутренние мм; направление и
' допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять,
' исходный immutable snapshot не меняется.
' ByRef minChord As Double — Расстояние или граница интервала minChord, внутренние мм; направление и
' допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять,
' исходный immutable snapshot не меняется.
' ByRef maxChord As Double — Расстояние или граница интервала maxChord, внутренние мм; направление и
' допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять,
' исходный immutable snapshot не меняется.
' Код метода или объявление: CRebarRuleExecutor в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleExecutor.DiagnosticsText
' Полный текстовый отчёт счётчиков и всех RuleTrace; численные длины в мм независимо от output units.
' Public Function DiagnosticsText() As String
' Код метода или объявление: CRebarRuleExecutor в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CRebarRuleExecutor.Build
' Исполняет опубликованные правила в готовой бетонной области и возвращает физическую раскладку; mesh не
' строит.
' Public Function Build(ByVal definition As CSectionDefinition, ByVal geometry As CGeometryRegion) As
' CRebarLayout
' ByVal definition As CSectionDefinition — Успешно опубликованный неизменяемый CSectionDefinition.
' ByVal geometry As CGeometryRegion — Подготовленный CGeometryRegion материала для исполнения правил.
' Код метода или объявление: CRebarRuleExecutor в окне проекта этой книги (Ctrl+R).
' ========================================================================================================
' CUnitSystem
' ========================================================================================================
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InitializeDefaults
' Задаёт программные defaults: input mm/mm^2/N/N*mm/MPa/1/mm и Tension, output m/m^2/N/N*mm/MPa/1/mm. Не
' заменяет обязательный Config ввод.
' Public Sub InitializeDefaults()
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.LoadFromSettings
' Загружает и проверяет все активные единицы и знаки через общий settings reader.
' Public Sub LoadFromSettings(ByVal settings As CSystemSettingsReader)
' ByVal settings As CSystemSettingsReader — CSystemSettingsReader, уже загрузивший Config.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InputLengthToInternal
' Переводит InputLengthToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой
' конвенцией знаков для N/Mx/My.
' Public Function InputLengthToInternal(ByVal value As Double) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InternalLengthToOutput
' Переводит InternalLengthToOutput из внутренних единиц в подготовленные output units; сохранённая единица
' при наличии имеет явный смысл.
' Public Function InternalLengthToOutput(ByVal value As Double, Optional ByVal savedUnit As String =
' vbNullString) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Optional ByVal savedUnit As String = vbNullString — Единица сохранённого Results; пустая использует
' текущую output-единицу CUnitSystem.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.OutputLengthToInternal
' Переводит OutputLengthToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с
' принятой конвенцией знаков для N/Mx/My.
' Public Function OutputLengthToInternal(ByVal value As Double, Optional ByVal savedUnit As String =
' vbNullString) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Optional ByVal savedUnit As String = vbNullString — Единица сохранённого Results; пустая использует
' текущую output-единицу CUnitSystem.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InputAreaToInternal
' Переводит InputAreaToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой
' конвенцией знаков для N/Mx/My.
' Public Function InputAreaToInternal(ByVal value As Double) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InternalAreaToOutput
' Переводит InternalAreaToOutput из внутренних единиц в подготовленные output units; сохранённая единица
' при наличии имеет явный смысл.
' Public Function InternalAreaToOutput(ByVal value As Double) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.OutputAreaToInternal
' Переводит OutputAreaToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой
' конвенцией знаков для N/Mx/My.
' Public Function OutputAreaToInternal(ByVal value As Double, Optional ByVal savedUnit As String =
' vbNullString) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Optional ByVal savedUnit As String = vbNullString — Единица сохранённого Results; пустая использует
' текущую output-единицу CUnitSystem.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InputFourthPowerLengthToInternal
' Переводит InputFourthPowerLengthToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности,
' с принятой конвенцией знаков для N/Mx/My.
' Public Function InputFourthPowerLengthToInternal(ByVal value As Double) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InternalFourthPowerLengthToOutput
' Переводит InternalFourthPowerLengthToOutput из внутренних единиц в подготовленные output units;
' сохранённая единица при наличии имеет явный смысл.
' Public Function InternalFourthPowerLengthToOutput(ByVal value As Double) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.OutputFourthPowerLengthToInternal
' Переводит OutputFourthPowerLengthToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей
' размерности, с принятой конвенцией знаков для N/Mx/My.
' Public Function OutputFourthPowerLengthToInternal(ByVal value As Double, Optional ByVal savedUnit As
' String = vbNullString) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Optional ByVal savedUnit As String = vbNullString — Базовая единица длины сохранённого Results: mm/cm/m;
' степень 4 применяет сам метод. Не передавайте mm4/cm4/m4. Пустая строка использует текущую output-длину.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InputForceToInternal
' Переводит InputForceToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой
' конвенцией знаков для N/Mx/My.
' Public Function InputForceToInternal(ByVal value As Double) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InternalForceToOutput
' Переводит InternalForceToOutput из внутренних единиц в подготовленные output units; сохранённая единица
' при наличии имеет явный смысл.
' Public Function InternalForceToOutput(ByVal value As Double) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InternalForceLengthSquaredToOutput
' Переводит InternalForceLengthSquaredToOutput из внутренних единиц в подготовленные output units;
' сохранённая единица при наличии имеет явный смысл.
' Public Function InternalForceLengthSquaredToOutput(ByVal value As Double) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InputMomentMxToInternal
' Переводит InputMomentMxToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с
' принятой конвенцией знаков для N/Mx/My.
' Public Function InputMomentMxToInternal(ByVal value As Double) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InputMomentMyToInternal
' Переводит InputMomentMyToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с
' принятой конвенцией знаков для N/Mx/My.
' Public Function InputMomentMyToInternal(ByVal value As Double) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InputMomentPerLengthToInternal
' Переводит InputMomentPerLengthToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с
' принятой конвенцией знаков для N/Mx/My.
' Public Function InputMomentPerLengthToInternal(ByVal value As Double) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InternalMomentMxToOutput
' Переводит InternalMomentMxToOutput из внутренних единиц в подготовленные output units; сохранённая
' единица при наличии имеет явный смысл.
' Public Function InternalMomentMxToOutput(ByVal value As Double) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InternalMomentMyToOutput
' Переводит InternalMomentMyToOutput из внутренних единиц в подготовленные output units; сохранённая
' единица при наличии имеет явный смысл.
' Public Function InternalMomentMyToOutput(ByVal value As Double) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InternalMomentMagnitudeToOutput
' Переводит InternalMomentMagnitudeToOutput из внутренних единиц в подготовленные output units;
' сохранённая единица при наличии имеет явный смысл.
' Public Function InternalMomentMagnitudeToOutput(ByVal value As Double) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InputStressToInternal
' Переводит InputStressToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с принятой
' конвенцией знаков для N/Mx/My.
' Public Function InputStressToInternal(ByVal value As Double) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InternalStressToOutput
' Переводит InternalStressToOutput из внутренних единиц в подготовленные output units; сохранённая единица
' при наличии имеет явный смысл.
' Public Function InternalStressToOutput(ByVal value As Double) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.OutputStressToInternal
' Переводит OutputStressToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с
' принятой конвенцией знаков для N/Mx/My.
' Public Function OutputStressToInternal(ByVal value As Double, Optional ByVal savedUnit As String =
' vbNullString) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Optional ByVal savedUnit As String = vbNullString — Единица сохранённого Results; пустая использует
' текущую output-единицу CUnitSystem.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InputCurvatureToInternal
' Переводит InputCurvatureToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с
' принятой конвенцией знаков для N/Mx/My.
' Public Function InputCurvatureToInternal(ByVal value As Double) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InternalCurvatureToOutput
' Переводит InternalCurvatureToOutput из внутренних единиц в подготовленные output units; сохранённая
' единица при наличии имеет явный смысл.
' Public Function InternalCurvatureToOutput(ByVal value As Double, Optional ByVal savedUnit As String =
' vbNullString) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Optional ByVal savedUnit As String = vbNullString — Единица сохранённого Results; пустая использует
' текущую output-единицу CUnitSystem.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.OutputCurvatureToInternal
' Переводит OutputCurvatureToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с
' принятой конвенцией знаков для N/Mx/My.
' Public Function OutputCurvatureToInternal(ByVal value As Double) As Double
' ByVal value As Double — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.OutputLengthUnit
' Возвращает текущий output selector LengthUnit, без преобразования или округления данных.
' Public Property Get OutputLengthUnit() As String
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.OutputAreaUnit
' Возвращает текущий output selector AreaUnit, без преобразования или округления данных.
' Public Property Get OutputAreaUnit() As String
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.OutputMomentUnit
' Возвращает текущий output selector MomentUnit, без преобразования или округления данных.
' Public Property Get OutputMomentUnit() As String
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.InputMomentPerLengthUnit
' Возвращает единицу пользовательского момента на длину по input force/moment/length settings.
' Public Property Get InputMomentPerLengthUnit() As String
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.OutputForceUnit
' Возвращает текущий output selector ForceUnit, без преобразования или округления данных.
' Public Property Get OutputForceUnit() As String
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.OutputForceLengthSquaredUnit
' Возвращает текущий output selector ForceLengthSquaredUnit, без преобразования или округления данных.
' Public Property Get OutputForceLengthSquaredUnit() As String
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.OutputStressUnit
' Возвращает текущий output selector StressUnit, без преобразования или округления данных.
' Public Property Get OutputStressUnit() As String
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.OutputCurvatureUnit
' Возвращает текущий output selector CurvatureUnit, без преобразования или округления данных.
' Public Property Get OutputCurvatureUnit() As String
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.OutputSignN
' Возвращает текущий output selector SignN, без преобразования или округления данных.
' Public Property Get OutputSignN() As String
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.OutputSignMx
' Возвращает текущий output selector SignMx, без преобразования или округления данных.
' Public Property Get OutputSignMx() As String
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CUnitSystem.OutputSignMy
' Возвращает текущий output selector SignMy, без преобразования или округления данных.
' Public Property Get OutputSignMy() As String
' Код метода или объявление: CUnitSystem в окне проекта этой книги (Ctrl+R).
' ========================================================================================================
' CConcreteRegion
' ========================================================================================================
' --------------------------------------------------------------------------------------------------------
' CConcreteRegion.SetOwner
' Назначает скалярное происхождение инженерной области без пересчёта A/контуров; это внешний coordinator
' API.
' Public Sub SetOwner(ByVal combinationID As String, ByVal standard As String, ByVal anchorID As String)
' ByVal combinationID As String — Скалярная metadata владельца/источника; геометрическая классификация от
' строки не зависит.
' ByVal standard As String — Скалярная metadata владельца/источника; геометрическая классификация от
' строки не зависит.
' ByVal anchorID As String — Непустой уникальный ID модельной привязки.
' Код метода или объявление: CConcreteRegion в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CConcreteRegion.GetSegment
' Возвращает точную аналитическую запись одного ребра в скаляры: мм, sweep радианы, loop-индекс.
' Public Sub GetSegment(ByVal index As Long, ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double,
' ByRef y2 As Double, ByRef centerX As Double, ByRef centerY As Double, ByRef sweep As Double, ByRef
' loopID As Long)
' ByVal index As Long — 1-based индекс в пределах соответствующего готового контейнера.
' ByRef x1 As Double — Модельная координата x1, мм; знаки соответствуют section XY/WCS. ByRef:
' возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется.
' ByRef y1 As Double — Модельная координата y1, мм; знаки соответствуют section XY/WCS. ByRef:
' возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется.
' ByRef x2 As Double — Модельная координата x2, мм; знаки соответствуют section XY/WCS. ByRef:
' возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется.
' ByRef y2 As Double — Модельная координата y2, мм; знаки соответствуют section XY/WCS. ByRef:
' возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется.
' ByRef centerX As Double — Модельная координата centerX, мм; знаки соответствуют section XY/WCS. ByRef:
' возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется.
' ByRef centerY As Double — Модельная координата centerY, мм; знаки соответствуют section XY/WCS. ByRef:
' возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется.
' ByRef sweep As Double — Signed sweep, радианы: CCW положителен, CW отрицателен, 0 запрещён,
' |sweep|<=2*pi. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не
' меняется.
' ByRef loopID As Long — Непустой устойчивый ID контура; у CConcreteRegion численный индекс loop
' 1..LoopCount. ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется.
' Код метода или объявление: CConcreteRegion в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CConcreteRegion.SegmentCount
' Число аналитических line/arc сегментов.
' Public Property Get SegmentCount() As Long
' Код метода или объявление: CConcreteRegion в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CConcreteRegion.BoundaryProbes
' Копия сохранённых пассивных пар P/Q; Empty допустим, контуры по ним не восстанавливаются.
' Public Property Get BoundaryProbes() As Variant
' Код метода или объявление: CConcreteRegion в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CConcreteRegion.BoundaryProbeCount
' Число сохранённых диагностических проб P/Q.
' Public Property Get BoundaryProbeCount() As Long
' Код метода или объявление: CConcreteRegion в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CConcreteRegion.LoopCount
' Число подготовленных материальных loops.
' Public Property Get LoopCount() As Long
' Код метода или объявление: CConcreteRegion в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CConcreteRegion.LoopIsOpening
' Читает классификацию конкретной loop по подготовленной топологии.
' Public Property Get LoopIsOpening(ByVal index As Long) As Boolean
' ByVal index As Long — 1-based индекс в пределах соответствующего готового контейнера.
' Код метода или объявление: CConcreteRegion в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CConcreteRegion.Area
' Аналитическая площадь материальной области, мм^2; не механическая A сетки.
' Public Property Get Area() As Double
' Код метода или объявление: CConcreteRegion в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CConcreteRegion.Source
' Скалярное происхождение подготовленной области.
' Public Property Get Source() As String
' Код метода или объявление: CConcreteRegion в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CConcreteRegion.SectionRevision
' Ревизия исходной модели, к которой относится область.
' Public Property Get SectionRevision() As Long
' Код метода или объявление: CConcreteRegion в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CConcreteRegion.OwnerID
' Инженерный владелец области; назначается coordinator, не геометрической классификацией.
' Public Property Get OwnerID() As String
' Код метода или объявление: CConcreteRegion в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CConcreteRegion.Standard
' Скалярная нормативная принадлежность инженерного владельца; geometry не вычисляет нормативные
' коэффициенты.
' Public Property Get Standard() As String
' Код метода или объявление: CConcreteRegion в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CConcreteRegion.AnchorID
' Скалярный ID сохранённой инженерной области.
' Public Property Get AnchorID() As String
' Код метода или объявление: CConcreteRegion в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CConcreteRegion.Clone
' Копирует готовую область и metadata; временные массивы следующей операции не меняют исходный snapshot.
' Public Function Clone() As CConcreteRegion
' Код метода или объявление: CConcreteRegion в окне проекта этой книги (Ctrl+R).
' ========================================================================================================
' CSectionGeometryQuery
' ========================================================================================================
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.LastDiagnostic
' Последняя машинная диагностика; Nothing при отсутствии записанного события.
' Public Property Get LastDiagnostic() As CGeometryDiagnostic
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.TransformCoordinates
' Преобразует скалярную точку: reflectX, rotation, shift; общий порядок совпадает с путями.
' Public Sub TransformCoordinates(ByRef x As Double, ByRef y As Double, ByVal shiftX As Double, ByVal
' shiftY As Double, Optional ByVal rotation As Double = 0#, Optional ByVal reflectX As Boolean = False)
' ByRef x As Double — Модельная координата x, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые
' скаляры/буфер можно изменять, исходный immutable snapshot не меняется.
' ByRef y As Double — Модельная координата y, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые
' скаляры/буфер можно изменять, исходный immutable snapshot не меняется.
' ByVal shiftX As Double — Перенос по X, мм.
' ByVal shiftY As Double — Перенос по Y, мм.
' Optional ByVal rotation As Double = 0# — Угол поворота вокруг (0,0), радианы, CCW положителен.
' Optional ByVal reflectX As Boolean = False — Отражение относительно локальной оси X перед поворотом и
' переносом.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.OffsetPath
' Смещает линию/аналитическую цепочку слева от обхода при положительном signedLeftDistance. Line/line
' стыки miter; arc стыки должны быть гладкими.
' Public Function OffsetPath(ByVal pathID As String, ByVal source As CGeometryPath, ByVal distance As
' Double) As CGeometryPath
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal source As CGeometryPath — Подготовленный объект/контейнер выбранного запроса, без повторного
' чтения Excel или назначения типа.
' ByVal distance As Double — Расстояние или граница интервала distance, внутренние мм; направление и
' допустимые пределы определяет вызываемый метод.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.ConstrainLine
' Ограничивает или продлевает одну прямую двумя поперечными плоскостями nX*x+nY*y=bound; направление
' исходного пути сохраняется.
' Public Function ConstrainLine(ByVal pathID As String, ByVal source As CGeometryPath, ByVal startNX As
' Double, ByVal startNY As Double, ByVal startBound As Double, ByVal endNX As Double, ByVal endNY As
' Double, ByVal endBound As Double) As CGeometryPath
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal source As CGeometryPath — Подготовленный объект/контейнер выбранного запроса, без повторного
' чтения Excel или назначения типа.
' ByVal startNX As Double — X-компонента начальной поперечной нормали; линия не должна быть параллельна
' плоскости.
' ByVal startNY As Double — Y-компонента начальной поперечной нормали.
' ByVal startBound As Double — Значение начальной плоскости startNX*x+startNY*y=startBound; мм при
' единичной нормали.
' ByVal endNX As Double — X-компонента конечной поперечной нормали.
' ByVal endNY As Double — Y-компонента конечной поперечной нормали.
' ByVal endBound As Double — Значение конечной плоскости endNX*x+endNY*y=endBound; мм при единичной
' нормали.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.Initialize
' Привязывает query к CSectionModel/revision. Достоверные контуры приоритетны; иначе используются реальные
' границы КЭ.
' Public Sub Initialize(ByVal section As CSectionModel)
' ByVal section As CSectionModel — CSectionModel на внешней границе query/build; ядро не читает Excel.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.ConcreteDomain
' Возвращает clone области текущей query-session; не подменяет A/I механической сетки.
' Public Function ConcreteDomain() As CConcreteRegion
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.ProjectionBounds
' Находит min/max аналитической области в заданном направлении, включая внутренние экстремумы дуг.
' Public Sub ProjectionBounds(ByVal region As CConcreteRegion, ByVal nx As Double, ByVal ny As Double,
' ByRef minimum As Double, ByRef maximum As Double)
' ByVal region As CConcreteRegion — Готовый CConcreteRegion; Nothing не является пустой областью.
' ByVal nx As Double — Компонента нормали/направления в section XY; вектор не нулевой.
' ByVal ny As Double — Компонента нормали/направления в section XY; вектор не нулевой.
' ByRef minimum As Double — Расстояние или граница интервала minimum, внутренние мм; направление и
' допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять,
' исходный immutable snapshot не меняется.
' ByRef maximum As Double — Расстояние или граница интервала maximum, внутренние мм; направление и
' допустимые пределы определяет вызываемый метод. ByRef: возвращаемые скаляры/буфер можно изменять,
' исходный immutable snapshot не меняется.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.OpeningBounds
' Возвращает габариты всех заданных отверстий из контуров, без ограничителя только первым отверстием.
' Public Sub OpeningBounds(ByVal contours As CSectionContours, ByRef bounds() As Double, ByRef
' openingCount As Long)
' ByVal contours As CSectionContours — CSectionContours с достоверными аналитическими Outer/Opening.
' ByRef bounds() As Double — Выходной массив габаритов, мм; формат определяется конкретным запросом.
' ByRef: возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется.
' ByRef openingCount As Long — Целый неотрицательный счётчик. ByRef: возвращаемые скаляры/буфер можно
' изменять, исходный immutable snapshot не меняется.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.PolygonRegion
' Готовит материальную область по замкнутому многоугольнику и проверяет топологию.
' Public Function PolygonRegion(ByVal coordinates As Variant) As CConcreteRegion
' ByVal coordinates As Variant — Массив координат вершин замкнутого многоугольника, мм.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.CircleRegion
' Создаёт точную круговую область по центру и положительному R.
' Public Function CircleRegion(ByVal centerX As Double, ByVal centerY As Double, ByVal radius As Double)
' As CConcreteRegion
' ByVal centerX As Double — Модельная координата centerX, мм; знаки соответствуют section XY/WCS.
' ByVal centerY As Double — Модельная координата centerY, мм; знаки соответствуют section XY/WCS.
' ByVal radius As Double — Положительный радиус, мм.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.BoundaryRegion
' Готовит область по аналитическим рёбрам, сшивает/проверяет замкнутые loops и роли.
' Public Function BoundaryRegion(ByRef edges() As TRegionEdge, ByVal count As Long, Optional ByVal source
' As String = "ValidatedContour") As CConcreteRegion
' ByRef edges() As TRegionEdge — Массив TRegionEdge аналитических рёбер с loop-данными, мм/радианы. ByRef:
' возвращаемые скаляры/буфер можно изменять, исходный immutable snapshot не меняется.
' ByVal count As Long — Целый неотрицательный счётчик.
' Optional ByVal source As String = "ValidatedContour" — Подготовленный объект/контейнер выбранного
' запроса, без повторного чтения Excel или назначения типа.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.LoopIsOpening
' Классифицирует подготовленную loop геометрически, не по строке имени.
' Public Function LoopIsOpening(ByVal region As CConcreteRegion, ByVal loopID As Long) As Boolean
' ByVal region As CConcreteRegion — Готовый CConcreteRegion; Nothing не является пустой областью.
' ByVal loopID As Long — Непустой устойчивый ID контура; у CConcreteRegion численный индекс loop
' 1..LoopCount.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.ConnectedRegionAtPoint
' Возвращает связанную компоненту материальной области по точке.
' Public Function ConnectedRegionAtPoint(ByVal region As CConcreteRegion, ByVal x As Double, ByVal y As
' Double) As CConcreteRegion
' ByVal region As CConcreteRegion — Готовый CConcreteRegion; Nothing не является пустой областью.
' ByVal x As Double — Модельная координата x, мм; знаки соответствуют section XY/WCS.
' ByVal y As Double — Модельная координата y, мм; знаки соответствуют section XY/WCS.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.MergeDisjointRegions
' Объединяет непересекающиеся области; не является универсальным Union перекрывающихся фигур.
' Public Function MergeDisjointRegions(ByVal regions As Collection) As CConcreteRegion
' ByVal regions As Collection — Collection готовых непересекающихся областей для MergeDisjointRegions.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.SubtractContainedRegion
' Вычитает строго вложенное отверстие, отклоняя нарушенные отношения границ.
' Public Function SubtractContainedRegion(ByVal outer As CConcreteRegion, ByVal opening As
' CConcreteRegion) As CConcreteRegion
' ByVal outer As CConcreteRegion — Исходная область материала/наружная область.
' ByVal opening As CConcreteRegion — Вычитаемая область отверстия.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.SubtractRegion
' Вычитает область по действующему ограниченному контурному контракту query; проверяет допустимость
' результата.
' Public Function SubtractRegion(ByVal outer As CConcreteRegion, ByVal opening As CConcreteRegion) As
' CConcreteRegion
' ByVal outer As CConcreteRegion — Исходная область материала/наружная область.
' ByVal opening As CConcreteRegion — Вычитаемая область отверстия.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.RegionsCoincide
' Проверяет совпадение подготовленных областей по геометрическому контракту.
' Public Function RegionsCoincide(ByVal first As CConcreteRegion, ByVal second As CConcreteRegion) As
' Boolean
' ByVal first As CConcreteRegion — Первая/вторая сравниваемая готовая область.
' ByVal second As CConcreteRegion — Первая/вторая сравниваемая готовая область.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.ClipHalfPlane
' Обрезает материальную область полуплоскостью, сохраняя аналитические рёбра и топологию.
' Public Function ClipHalfPlane(ByVal region As CConcreteRegion, ByVal nx As Double, ByVal ny As Double,
' ByVal bound As Double) As CConcreteRegion
' ByVal region As CConcreteRegion — Готовый CConcreteRegion; Nothing не является пустой областью.
' ByVal nx As Double — Компонента нормали/направления в section XY; вектор не нулевой.
' ByVal ny As Double — Компонента нормали/направления в section XY; вектор не нулевой.
' ByVal bound As Double — Расстояние или граница интервала bound, внутренние мм; направление и допустимые
' пределы определяет вызываемый метод.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.ClipWindow
' Обрезает область конечным окном в направленных n/t координатах.
' Public Function ClipWindow(ByVal region As CConcreteRegion, ByVal nx As Double, ByVal ny As Double,
' ByVal nMin As Double, ByVal nMax As Double, ByVal tMin As Double, ByVal tMax As Double) As
' CConcreteRegion
' ByVal region As CConcreteRegion — Готовый CConcreteRegion; Nothing не является пустой областью.
' ByVal nx As Double — Компонента нормали/направления в section XY; вектор не нулевой.
' ByVal ny As Double — Компонента нормали/направления в section XY; вектор не нулевой.
' ByVal nMin As Double — Расстояние или граница интервала nMin, внутренние мм; направление и допустимые
' пределы определяет вызываемый метод.
' ByVal nMax As Double — Расстояние или граница интервала nMax, внутренние мм; направление и допустимые
' пределы определяет вызываемый метод.
' ByVal tMin As Double — Расстояние или граница интервала tMin, внутренние мм; направление и допустимые
' пределы определяет вызываемый метод.
' ByVal tMax As Double — Расстояние или граница интервала tMax, внутренние мм; направление и допустимые
' пределы определяет вызываемый метод.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.MaterialRegionThroughPlane
' Получает материальную область по заданной плоскости и окну; не нормативный расчёт трещин.
' Public Function MaterialRegionThroughPlane(ByVal region As CConcreteRegion, ByVal nx As Double, ByVal ny
' As Double, ByVal seedPlane As Double) As CConcreteRegion
' ByVal region As CConcreteRegion — Готовый CConcreteRegion; Nothing не является пустой областью.
' ByVal nx As Double — Компонента нормали/направления в section XY; вектор не нулевой.
' ByVal ny As Double — Компонента нормали/направления в section XY; вектор не нулевой.
' ByVal seedPlane As Double — Подготовленный объект/контейнер выбранного запроса, без повторного чтения
' Excel или назначения типа.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.ContainsPoint
' Проверяет действительную принадлежность материальной области с заданной политикой границы.
' Public Function ContainsPoint(ByVal region As CConcreteRegion, ByVal x As Double, ByVal y As Double,
' Optional ByVal includeBoundary As Boolean = True) As Boolean
' ByVal region As CConcreteRegion — Готовый CConcreteRegion; Nothing не является пустой областью.
' ByVal x As Double — Модельная координата x, мм; знаки соответствуют section XY/WCS.
' ByVal y As Double — Модельная координата y, мм; знаки соответствуют section XY/WCS.
' Optional ByVal includeBoundary As Boolean = True — Считать точку на границе принадлежащей области.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.BoundaryDistance
' Возвращает минимальное расстояние до любой внешней/внутренней границы, мм.
' Public Function BoundaryDistance(ByVal region As CConcreteRegion, ByVal x As Double, ByVal y As Double)
' As Double
' ByVal region As CConcreteRegion — Готовый CConcreteRegion; Nothing не является пустой областью.
' ByVal x As Double — Модельная координата x, мм; знаки соответствуют section XY/WCS.
' ByVal y As Double — Модельная координата y, мм; знаки соответствуют section XY/WCS.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.RoundedPolygon
' Создаёт согласованную границу по вершинам/радиусам/ID. Edge.ID начинается у указанной вершины, Corner.ID
' существует только при R>0.
' Public Function RoundedPolygon(ByVal pathID As String, ByVal vertices As Variant, ByVal radii As
' Variant, ByVal vertexIDs As Variant) As CGeometryPath
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal vertices As Variant — Двумерные вершины: Array(Array(x,y),...); координаты во внутренних мм.
' ByVal radii As Variant — Массив неотрицательных радиусов вершин в мм, длина равна vertices; clamp
' бетонного радиуса отсутствует.
' ByVal vertexIDs As Variant — Массив уникальных смысловых ID вершин; определяет Edge.ID/Corner.ID.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.OffsetRoundedPolygon
' Смещает весь скруглённый многоугольник с согласованными стыками. Положительное расстояние внутрь
' площади; отрицательное наружу, из Opening в бетон.
' Public Function OffsetRoundedPolygon(ByVal pathID As String, ByVal vertices As Variant, ByVal radii As
' Variant, ByVal vertexIDs As Variant, ByVal distance As Double) As CGeometryPath
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal vertices As Variant — Двумерные вершины: Array(Array(x,y),...); координаты во внутренних мм.
' ByVal radii As Variant — Массив неотрицательных радиусов вершин в мм, длина равна vertices; clamp
' бетонного радиуса отсутствует.
' ByVal vertexIDs As Variant — Массив уникальных смысловых ID вершин; определяет Edge.ID/Corner.ID.
' ByVal distance As Double — Signed offset, мм: положительно внутрь площади, отрицательно наружу; это не
' left-offset обхода.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.RegionFromPath
' Готовит область по одному замкнутому аналитическому пути; открытый путь не становится бетоном.
' Public Function RegionFromPath(ByVal path As CGeometryPath) As CConcreteRegion
' ByVal path As CGeometryPath — Существующий непрерывный опубликованный CGeometryPath.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.AdjacentRectangles
' Создаёт одну границу смежных неперекрывающихся прямоугольников. Произвольное пересекающееся CAD Union не
' выполняется.
' Public Function AdjacentRectangles(ByVal pathID As String, ByVal rectangles As Variant) As CGeometryPath
' ByVal pathID As String — Непустой уникальный ID пути в определении.
' ByVal rectangles As Variant — Array(Array(x,y,width,height),...); мм, размеры положительные,
' прямоугольники смежные без перекрытий.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.LineIntervals
' Возвращает реальные материальные интервалы пересечения с линией, с отверстиями и всеми компонентами.
' Public Function LineIntervals(ByVal region As CConcreteRegion, ByVal pointX As Double, ByVal pointY As
' Double, ByVal directionX As Double, ByVal directionY As Double, Optional ByVal parameterMin As Double =
' -1E+100, Optional ByVal parameterMax As Double = 1E+100) As Variant
' ByVal region As CConcreteRegion — Готовый CConcreteRegion; Nothing не является пустой областью.
' ByVal pointX As Double — Модельная координата pointX, мм; знаки соответствуют section XY/WCS.
' ByVal pointY As Double — Модельная координата pointY, мм; знаки соответствуют section XY/WCS.
' ByVal directionX As Double — X-компонента направления проекции/выбора экстремума; вектор не нулевой, у
' RebarRowExtremeLabel единичный.
' ByVal directionY As Double — Y-компонента направления проекции/выбора экстремума.
' Optional ByVal parameterMin As Double = -1E+100 — Граница нормированного параметра выбранной
' линии/интервала.
' Optional ByVal parameterMax As Double = 1E+100 — Граница нормированного параметра выбранной
' линии/интервала.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.SampleMaterialIntervals
' Получает материальные интервалы для заданных направлений/уровней; не восстанавливает область по одной
' площади.
' Public Function SampleMaterialIntervals(ByVal region As CConcreteRegion, ByVal nx As Double, ByVal ny As
' Double, ByVal seedPlane As Double, ByVal tMin As Double, ByVal tMax As Double, ByVal probeCount As Long)
' As Variant
' ByVal region As CConcreteRegion — Готовый CConcreteRegion; Nothing не является пустой областью.
' ByVal nx As Double — Компонента нормали/направления в section XY; вектор не нулевой.
' ByVal ny As Double — Компонента нормали/направления в section XY; вектор не нулевой.
' ByVal seedPlane As Double — Signed нормальная координата опорной плоскости, мм, относительно начала
' section XY; это скаляр Double, не объект плоскости.
' ByVal tMin As Double — Расстояние или граница интервала tMin, внутренние мм; направление и допустимые
' пределы определяет вызываемый метод.
' ByVal tMax As Double — Расстояние или граница интервала tMax, внутренние мм; направление и допустимые
' пределы определяет вызываемый метод.
' ByVal probeCount As Long — Целый неотрицательный счётчик.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.ConcreteCoverFromPointAlongDirection
' В текущей query-session находит геометрическое расстояние до бетонной границы по заданному направлению.
' Public Function ConcreteCoverFromPointAlongDirection(ByVal pointX As Double, ByVal pointY As Double,
' ByVal directionX As Double, ByVal directionY As Double) As Double
' ByVal pointX As Double — Модельная координата pointX, мм; знаки соответствуют section XY/WCS.
' ByVal pointY As Double — Модельная координата pointY, мм; знаки соответствуют section XY/WCS.
' ByVal directionX As Double — X-компонента направления проекции/выбора экстремума; вектор не нулевой, у
' RebarRowExtremeLabel единичный.
' ByVal directionY As Double — Y-компонента направления проекции/выбора экстремума.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.RegionFromContours
' Готовит область из единственных достоверных CSectionContours Outer/Opening.
' Public Function RegionFromContours(ByVal contours As CSectionContours) As CConcreteRegion
' ByVal contours As CSectionContours — CSectionContours с достоверными аналитическими Outer/Opening.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' CSectionGeometryQuery.RegionCentroid
' Вычисляет аналитический центр материальной области, мм; не заменяет центр механической сетки НДС.
' Public Sub RegionCentroid(ByVal region As CConcreteRegion, ByRef x As Double, ByRef y As Double)
' ByVal region As CConcreteRegion — Готовый CConcreteRegion; Nothing не является пустой областью.
' ByRef x As Double — Модельная координата x, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые
' скаляры/буфер можно изменять, исходный immutable snapshot не меняется.
' ByRef y As Double — Модельная координата y, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые
' скаляры/буфер можно изменять, исходный immutable snapshot не меняется.
' Код метода или объявление: CSectionGeometryQuery в окне проекта этой книги (Ctrl+R).
' ========================================================================================================
' modSectionRecipeTypes
' ========================================================================================================
' --------------------------------------------------------------------------------------------------------
' modSectionRecipeTypes.ESectionLoopRole
' Outer/Opening — геометрические роли, не выводятся из имени контура.
' Public Enum ESectionLoopRole
' LoopOuter = 1
' LoopOpening = 2
' Код метода или объявление: modSectionRecipeTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modSectionRecipeTypes.ERebarEndMode
' IncludeBoth/ExcludeBoth/IncludeStart/IncludeEnd/ClosedUniform; последние требуют соответствующей
' замкнутости.
' Public Enum ERebarEndMode
' IncludeBoth = 1
' ExcludeBoth = 2
' IncludeStart = 3
' IncludeEnd = 4
' ClosedUniform = 5
' Код метода или объявление: modSectionRecipeTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modSectionRecipeTypes.ERebarDistribution
' ByCount/ByMaxStep/ByFixedStep/ExplicitPositions/ProjectFrom/DependentOffset — режимы общего исполнителя.
' Public Enum ERebarDistribution
' ByCount = 1
' ByMaxStep = 2
' ByFixedStep = 3
' ExplicitPositions = 4
' ProjectFrom = 5
' DependentOffset = 6
' Код метода или объявление: modSectionRecipeTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modSectionRecipeTypes.ERebarDistributionScope
' Допустимо только PerPath: число стержней, расстояния между ними и отступы от концов применяются отдельно
' к каждой независимой линии группы.
' Public Enum ERebarDistributionScope
' PerPath = 1
' Код метода или объявление: modSectionRecipeTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modSectionRecipeTypes.ERemainderPlacement
' RemainderStart/End/Center: положение остатка для точного шага.
' Public Enum ERemainderPlacement
' RemainderStart = 1
' RemainderEnd = 2
' RemainderCenter = 3
' Код метода или объявление: modSectionRecipeTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modSectionRecipeTypes.ERebarPublication
' PublishByRule или PublishPerSourcePosition; влияет на вывод, не геометрический шаг.
' Public Enum ERebarPublication
' PublishByRule = 1
' PublishPerSourcePosition = 2
' Код метода или объявление: modSectionRecipeTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modSectionRecipeTypes.ERebarPositionOrder
' Порядок вывода стержней: AlongPath — от начала линии, FromPathMiddle — от её середины. Места стержней и
' их исходные номера сохраняются.
' Public Enum ERebarPositionOrder
' AlongPath = 1
' FromPathMiddle = 2
' Код метода или объявление: modSectionRecipeTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modSectionRecipeTypes.ERebarCoincidence
' RejectCoincidence/MergeIdentical/KeepPublishedOwner; последнее сохраняет явно выбранного опубликованного
' владельца.
' Public Enum ERebarCoincidence
' RejectCoincidence = 1
' MergeIdentical = 2
' KeepPublishedOwner = 3
' Код метода или объявление: modSectionRecipeTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modSectionRecipeTypes.ERebarPlacementCheck
' CheckCenter или CheckFullDisk относительно Outer/Opening; ужесточение задаётся явно.
' Public Enum ERebarPlacementCheck
' CheckCenter = 1
' CheckFullDisk = 2
' Код метода или объявление: modSectionRecipeTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modSectionRecipeTypes.ERebarOverlapPolicy
' ReportOverlap или RejectOverlap. Контакт допустим; нормативный минимальный зазор не вводится.
' Public Enum ERebarOverlapPolicy
' ReportOverlap = 1
' RejectOverlap = 2
' Код метода или объявление: modSectionRecipeTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modSectionRecipeTypes.TSectionLoop
' Raw скалярное описание роли/ссылок контура. Заполняет фасад Outer/Opening; автору не нужен собственный
' классификатор.
' Public Type TSectionLoop
' PathID As String
' LoopID As String
' Role As ESectionLoopRole
' ParentID As String
' Comment As String
' ArcComment As String
' ContourIDs As Variant
' Код метода или объявление: modSectionRecipeTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modSectionRecipeTypes.TRebarRule
' Raw snapshot правила исполнителя. Фасад копирует options, source references и параметры, координаты
' баров вычисляет общий executor.
' Public Type TRebarRule
' RuleID As String
' Target As String
' Mode As ERebarDistribution
' DistributionScope As ERebarDistributionScope
' Enabled As Boolean
' Diameter As Double
' SteelClass As String
' Count As Long
' Step As Double
' StartSetback As Double
' EndSetback As Double
' EndMode As ERebarEndMode
' RemainderPlacement As ERemainderPlacement
' SourceRuleID As String
' TangentOffset As Double
' NormalOffset As Double
' EverySecond As Boolean
' IncomingTangentAtJoint As Boolean
' DirectionX As Double
' DirectionY As Double
' SourceNormalSign As Double
' SuccessfulProjectionOrdinal As Boolean
' NewPublicationRoot As Boolean
' SkipMissingProjection As Boolean
' Publication As ERebarPublication
' PositionOrder As ERebarPositionOrder
' Coincidence As ERebarCoincidence
' PlacementCheck As ERebarPlacementCheck
' OverlapPolicy As ERebarOverlapPolicy
' SourcePrefix As String
' SourceSuffix As String
' AnnotationGroup As String
' Comment As String
' SourceReference As String
' ExplicitStations As Variant
' Код метода или объявление: modSectionRecipeTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modSectionRecipeTypes.TSectionAnnotation
' Raw annotation rule с модельными привязками/измерением. Фасад объявляет, общий engine измеряет, layout
' оформляет.
' Public Type TSectionAnnotation
' AnnotationID As String
' Kind As String
' Target As String
' SegmentID As String
' StartAnchorID As String
' EndAnchorID As String
' Name As String
' StartX As Double
' StartY As Double
' EndX As Double
' EndY As Double
' NormalX As Double
' NormalY As Double
' DirectionX As Double
' DirectionY As Double
' Value As Double
' DisplayIncrement As Double
' ReverseDirection As Boolean
' Comment As String
' Код метода или объявление: modSectionRecipeTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modSectionRecipeTypes.TSectionMeshOptions
' Настройки общей дискретизации: StepX/StepY в мм>0, subdivisions>=1, MaterialID>=1. Для книги читаются
' общим workflow.
' Public Type TSectionMeshOptions
' StepX As Double
' StepY As Double
' BoundarySubdivisions As Long
' MaterialID As Long
' Код метода или объявление: modSectionRecipeTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modSectionRecipeTypes.TRebarPosition
' Данные места стержня до объединения совпадающих стержней: координаты в мм, касательная, расстояние от
' начала линии, номера внутри правила и исходного ряда, уникальные имена.
' Public Type TRebarPosition
' X As Double
' Y As Double
' TangentX As Double
' TangentY As Double
' Station As Double
' RuleIndex As Long
' RootRuleIndex As Long
' RootOrdinal As Long
' RootPublicationRank As Long
' SourceOrdinal As Long
' SourcePositionID As String
' PathID As String
' Код метода или объявление: modSectionRecipeTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modSectionRecipeTypes.SectionRecipeError
' Выдаёт общий код 5530 с owner/причиной/строкой источника; не назначает нормативный статус.
' Public Sub SectionRecipeError(ByVal owner As String, ByVal reason As String, Optional ByVal
' sourceReference As String = "")
' ByVal owner As String — Имя компонента/владельца ошибки, а не тип сечения для ветвления.
' ByVal reason As String — Понятная причина отказа и исправление; машинный код не извлекается из текста.
' Optional ByVal sourceReference As String = "" — Строка исходного адреса/программного параметра; без
' объекта Excel.
' Код метода или объявление: modSectionRecipeTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modSectionRecipeTypes.SectionRecipeGeometryError
' В обработчике добавляет SourceReference к query/path отказу; неожиданные ошибки сохраняют исходный
' Err.Number/Source.
' Public Sub SectionRecipeGeometryError(ByVal owner As String, ByVal sourceReference As String)
' ByVal owner As String — Имя компонента/владельца ошибки, а не тип сечения для ветвления.
' ByVal sourceReference As String — Строка исходного адреса/программного параметра; без объекта Excel.
' Код метода или объявление: modSectionRecipeTypes в окне проекта этой книги (Ctrl+R).
' ========================================================================================================
' modGeometryTypes
' ========================================================================================================
' --------------------------------------------------------------------------------------------------------
' modGeometryTypes.GEOM_PI
' Число pi как сумма сохраняемых VBE литералов, без изменения Double при компиляции.
' Public Const GEOM_PI As Double = 3.14159265358979 + 3.10862446895044E-15
' Код метода или объявление: modGeometryTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modGeometryTypes.GEOM_TOLERANCE
' Геометрический допуск совпадений исходных координат, мм; не округление ключей результатов.
' Public Const GEOM_TOLERANCE As Double = 0.000000001
' Код метода или объявление: modGeometryTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modGeometryTypes.GEOM_MAX_REBAR_COUNT
' Предел представимости счётчика позиций, не обещание доступности памяти для такого массива.
' Public Const GEOM_MAX_REBAR_COUNT As Double = 2147483646# ' Long-счетчик должен сохранять представимость
' следующего шага обхода.
' Код метода или объявление: modGeometryTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modGeometryTypes.TPoint2D
' Два координатных скаляра X/Y, мм, без Excel или объекта на каждую точку.
' Public Type TPoint2D
' X As Double
' Y As Double
' Код метода или объявление: modGeometryTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modGeometryTypes.TRegionEdge
' Ориентированное аналитическое ребро: endpoints/center мм, Sweep радианы, 0 означает прямую, LoopID —
' численный loop.
' Public Type TRegionEdge
' X1 As Double
' Y1 As Double
' X2 As Double
' Y2 As Double
' CenterX As Double
' CenterY As Double
' Sweep As Double
' LoopID As Long
' Код метода или объявление: modGeometryTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modGeometryTypes.TFiber
' Скалярные данные механического бетонного КЭ. Не заменяются аналитической площадью области.
' Public Type TFiber
' X As Double
' Y As Double
' Width As Double
' Height As Double
' Area As Double
' FillFactor As Double
' MaterialID As Long
' Код метода или объявление: modGeometryTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modGeometryTypes.TGeometryProperties
' Скалярные геометрические свойства во внутренних единицах; источник аналитический/сеточный определяется
' вызывающим алгоритмом.
' Public Type TGeometryProperties
' Area As Double
' StaticMomentX As Double
' StaticMomentY As Double
' CentroidX As Double
' CentroidY As Double
' Ix As Double
' Iy As Double
' Ixy As Double
' Ixc As Double
' Iyc As Double
' Ixyc As Double
' PrincipalI1 As Double
' PrincipalI2 As Double
' PrincipalAngleRad As Double
' RadiusX As Double
' RadiusY As Double
' PrincipalRadius1 As Double
' PrincipalRadius2 As Double
' Код метода или объявление: modGeometryTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modGeometryTypes.GeomMax
' Максимум двух чисел одной размерности, без преобразования единиц.
' Public Function GeomMax(ByVal A As Double, ByVal B As Double) As Double
' ByVal A As Double — Число той же размерности, что второй аргумент; helper единицы не преобразует.
' ByVal B As Double — Число той же размерности, что второй аргумент; helper единицы не преобразует.
' Код метода или объявление: modGeometryTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modGeometryTypes.GeomMin
' Минимум двух чисел одной размерности.
' Public Function GeomMin(ByVal A As Double, ByVal B As Double) As Double
' ByVal A As Double — Число той же размерности, что второй аргумент; helper единицы не преобразует.
' ByVal B As Double — Число той же размерности, что второй аргумент; helper единицы не преобразует.
' Код метода или объявление: modGeometryTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modGeometryTypes.ReadContourArcSweep
' Читает согласованный signed sweep точной контурной дуги, не выбирая произвольные несовместимые входы.
' Public Function ReadContourArcSweep(ByVal value As Variant, ByVal inputLocation As String) As Double
' ByVal value As Variant — Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт
' единственный перевод.
' ByVal inputLocation As String — Строка реального исходного адреса для диагностического сообщения.
' Код метода или объявление: modGeometryTypes в окне проекта этой книги (Ctrl+R).
' --------------------------------------------------------------------------------------------------------
' modGeometryTypes.BuildGeneratedSectionModel
' Общий тонкий adapter mesh/layout -> CSectionModel; автор рецепта использует CSectionBuilder, а не
' отдельный генератор типа.
' Public Function BuildGeneratedSectionModel(ByVal mesh As CFiberMeshBuilder, ByVal rebars As
' CRebarLayout, Optional ByVal sourceType As String = "Generated") As CSectionModel
' ByVal mesh As CFiberMeshBuilder — Подготовленный объект/контейнер выбранного запроса, без повторного
' чтения Excel или назначения типа.
' ByVal rebars As CRebarLayout — Подготовленный объект/контейнер выбранного запроса, без повторного чтения
' Excel или назначения типа.
' Optional ByVal sourceType As String = "Generated" — Скалярная metadata владельца/источника;
' геометрическая классификация от строки не зависит.
' Код метода или объявление: modGeometryTypes в окне проекта этой книги (Ctrl+R).
