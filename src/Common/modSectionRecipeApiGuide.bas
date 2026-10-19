Attribute VB_Name = "modSectionRecipeApiGuide"
'
'
' ====================================================================================================
' СОЗДАНИЕ СЕЧЕНИЯ: ГЕОМЕТРИЯ, АРМАТУРА И АННОТАЦИИ
' ====================================================================================================
'
'   Здесь собраны готовые методы для автора нового типа сечения. Описание задаёт, что нужно получить;
'   программа по нему строит бетонную сетку, размещает стержни и готовит схему.
'
'   Модуль состоит только из комментариев. Точные объявления методов приведены вместе с параметрами; код
'   открывается в соответствующем классе через окно проекта Ctrl+R. Поиск по имени — Ctrl+F.
'
' СОДЕРЖАНИЕ
'     01  Начало и завершение описания
'     02  Входные размеры и единицы
'     03  Границы бетона и отверстий
'     04  Линии центров стержней и преобразования
'     05  Размещение арматуры
'     06  Условия размещения стержней
'     07  Размеры, пояснения и подписи арматуры
'     08  Составные линии и отдельные участки
'     09  Адреса ввода и сообщения об ошибках
'
'
' ====================================================================================================
' КАК ЧИТАТЬ ОПИСАНИЕ
' ====================================================================================================
'
'   Все координаты, длины и диаметры внутри рецепта задаются в мм, площади — в мм^2, углы — в радианах.
'   Значения из таблиц переводит context.Units; повторно переводить уже полученные мм не нужно.
'
'   Path — геометрическая линия из отрезков и дуг. Замкнутая линия может быть границей бетона или
'   отверстия. Линия центров стержней получается, например, смещением контура на as; вычислять
'   координаты каждого стержня вручную не требуется.
'
'   Segment — один отрезок или одна дуга. PathGroup объединяет несколько независимых линий. Station в
'   имени метода означает расстояние от начала линии в мм.
'
'   ID — уникальное текстовое имя линии, ряда, привязки или подписи. Anchor — именованная точка для
'   размеров и пояснений. Rule — описание одного способа размещения стержней.
'
'   Optional означает необязательный параметр; его значение по умолчанию указано после знака =. ByVal
'   передаёт значение, ByRef позволяет методу вернуть результат через параметр. Nothing вместо options
'   создаёт обычные условия размещения.
'
'   Пробел и символ _ в конце строки VBA продолжают команду на следующей строке. Это знак записи кода, а
'   не передаваемый параметр. Именованный аргумент записывается как имя:=значение.
'
'   Initialize открывает одно описание, Finish завершает его. CirclePath, LinePath, RectanglePath и
'   другие создающие команды сами сохраняют линию. Возвращённый path уже зарегистрирован; Outer, Opening
'   и армирование получают его ID. AddPath нужен только для отдельно созданной линии.
'
'   Таблицы и единицы context — независимый снимок ввода. context.Units используется для преобразования
'   чисел; перенастраивать его нельзя. Проверка и размещение стержней выполняются общим Build до
'   создания сетки.
'
'   Команды подписей задают группу, геометрическую привязку и направление. Расстояние от линии арматуры
'   до края бетона программа определяет по фактическому контуру; дополнительный отступ задаётся в
'   настройках плоттера.
'
'   МИНИМАЛЬНЫЙ ПРИМЕР
'     Полный класс CSectionRecipeMySection. Размеры заданы в мм; context требуется интерфейсу, входных
'     таблиц в этом примере нет. Создайте Class Module с этим именем и вставьте текст:
'
'       Option Explicit
'
'       ' ==========================================================================
'       ' Учебный рецепт прямоугольного сечения с пятью верхними стержнями
'       ' ==========================================================================
'       ' Показывает описание бетонной границы, армирования и аннотаций
'       ' через CSectionDefinitionBuilder без отдельной расстановки стержней.
'       Implements ISectionRecipe
'
'       ' Размеры заданы постоянными в мм. Контекст передаётся по интерфейсу рецепта; входные таблицы в
'       ' этом примере не используются.
'       Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
'           Dim facade As CSectionDefinitionBuilder, path As CGeometryPath
'           Dim options As CRebarRuleOptions
'           Set facade = New CSectionDefinitionBuilder
'           facade.Initialize "MySection"  ' "MySection" — имя типа сечения.
'       ' "Boundary" — имя прямоугольника; 0 — X левого нижнего угла, мм; 0 — Y левого нижнего угла,
'       ' мм; 600 — ширина, мм; 1000 — высота, мм.
'           Set path = facade.RectanglePath("Boundary", 0#, 0#, 600#, 1000#)
'           facade.Outer "OUTER", path.ID  ' "OUTER" — имя наружной границы бетона; path.ID — замкнутый контур.
'           Set path = facade.LinePath("TopAxis", 50#, 950#, 550#, 950#)  ' "TopAxis" — имя линии; 50 — X начала, мм; 950 — Y начала, мм; 550 — X конца, мм; 950 — Y конца, мм.
'           Set options = New CRebarRuleOptions
'           options.AnnotationGroup = "Top"
'       ' "Top.Main" — имя нового ряда; path.ID — линия центров стержней; 5 — заданное число
'       ' стержней; 20 — диаметр стержня, мм; options — отступы, нумерация и условия размещения.
'           facade.RebarByCount "Top.Main", path.ID, 5, 20#, options
'       ' "Top" — группа общей подписи всех её рядов; "Top.Main" — ряд, к концам которого привязана
'       ' подпись; 0 — X-компонента направления подписи; 1 — Y-компонента направления подписи.
'           facade.RebarRowLabel "Top", "Top.Main", 0#, 1#
'       ' "DIM_B" — имя размера; "B" — обозначение на схеме; 0 — X первой точки, мм; 600 — X второй
'       ' точки, мм; 1000 — уровень исходной размерной линии Y, мм.
'           facade.HorizontalDistance "DIM_B", "B", 0#, 600#, 1000#
'           Set ISectionRecipe_CreateDefinition = facade.Finish
'       End Function
'
'     В CSectionTypeRegistry.Class_Initialize добавьте RegisterRecipe "MySection", New
'     CSectionRecipeMySection, Array(). Затем выполните UpdateSectionRecipeTypes. Пустой Array()
'     подходит этому примеру с постоянными размерами; для размеров с листа «Сечения» укажите имя входной
'     таблицы.
'     На листе «Сечения» каждому типу соответствует одна таблица: General содержит размеры, Main —
'     основное армирование сторон, Extra — дополнительные ряды. Например: Set reader =
'     context.Reader("rngPointedRectGeometry"), B = reader.PositiveLength("General", "B"), n =
'     reader.Count("Main", "Top", "n"). Ридер проверяет значения и переводит длины в мм. Для своей
'     структуры доступны context.Table и ValueAt.
'
'   УСТРОЙСТВО ШТАТНЫХ РЕЦЕПТОВ
'     В Circle, RectSet, PointedRect и HollowRectangle сначала читаются и проверяются размеры и
'     параметры включённых рядов. Подготовленные значения, адреса исходных ячеек и соответствие названий
'     Config внутренним именам сторон сохраняются в Private Type самого рецепта. Для одного размера
'     отдельный метод чтения не нужен.
'     CreateDefinition связывает этапы: подготовку данных, AddGeometry, AddReinforcement и
'     AddAnnotations. AddGeometry задаёт границы бетона и отверстий. AddReinforcement создаёт линии
'     центров арматуры и регистрирует правила размещения стержней. AddAnnotations задаёт размеры и
'     подписи. Эти методы получают готовые параметры и не читают входные таблицы.
'     Параметры отключённого ряда пропускаются. В HollowRectangle внутренние as нужны и без стержней на
'     соответствующей стороне: они ограничивают линии арматуры у отверстия. Ридер использует короткие
'     ключи независимо от русских подписей и порядка полей.
'
'   ПОДГОТОВКА КОНТАКТНЫХ РЯДОВ
'     ReadRebarInputs читает и проверяет параметры активных рядов. В AddReinforcement создают отдельный
'     CRebarContactBuilder для каждого основного ряда или стороны, задают общие настройки дополнительных
'     рядов и вызывают AppendRow один раз на включённый ряд по порядку. Полученные options передают в
'     facade.RebarDependent. Отключённые ряды пропускают до чтения их необязательных полей; при
'     отрисовке AppendRow не вызывают.
'     Initialize получает новые общие options именно дополнительных рядов и сохраняет их копию. Каждый
'     AppendRow возвращает отдельный объект. Можно изменить SourceSuffix, Comment и другие
'     индивидуальные поля перед передачей фасаду. Построитель прибавляет к сумме радиусов 2,5 мм на
'     выступы арматуры периодического профиля для каждой соседней пары; диаметры и площади в расчёте не
'     меняются. Диаметры промежуточных рядов и эти добавки учитываются отдельно по каждому направлению.
'     EverySecondBar не уменьшает расстояние.
'     Фрагмент для уже объявленного основного ряда Top.Main диаметром 32 мм. Stacked направлен внутрь от
'     верхней линии; SideBySide — вдоль неё. Смещения составят -28,5 и -49 мм по нормали, затем +24,5 мм
'     вдоль линии. Все дополнительные ряды остаются связанными с Top.Main.
'       Dim contact As CRebarContactBuilder, extraOptions As CRebarRuleOptions
'       Dim row2 As CRebarRuleOptions, row3 As CRebarRuleOptions, row4 As CRebarRuleOptions
'       Set extraOptions = New CRebarRuleOptions
'       extraOptions.AnnotationGroup = "Top"  ' Общая подпись основного и дополнительных рядов этой группы.
'       Set contact = New CRebarContactBuilder
'       contact.Initialize 32#, -1#, 1#, extraOptions  ' Основной диаметр, знаки нормали и касательной, общие настройки.
'       Set row2 = contact.AppendRow(20#, "Stacked", "EverySecondBar", "Program.Row2")
'       row2.SourceSuffix = "_R2"  ' Индивидуальное окончание имени позиции второго ряда.
'       Set row3 = contact.AppendRow(16#, "Stacked", "EachBar", "Program.Row3")
'       Set row4 = contact.AppendRow(12#, "SideBySide", "EachBar", "Program.Row4")
'       ' На этапе AddReinforcement используем готовые настройки; ID и связь с Main заданы явно.
'       facade.RebarDependent "Top.Row2", "Top.Main", 20#, row2
'       facade.RebarDependent "Top.Row3", "Top.Main", 16#, row3
'       facade.RebarDependent "Top.Row4", "Top.Main", 12#, row4
'
'   ОПИСАТЕЛЬ И ПРИВЯЗКИ
'     CPolygonSpec объединяет исходные ID, X, Y, FilletRadius и SourceReference каждой вершины. New  ->
'     AddVertex  ->  Seal. Состояние читают Count/IsSealed, значения —
'     VertexID/X/Y/FilletRadius/SourceReference по индексу от 1. После Seal изменение запрещено;
'     повторный Seal безопасен. Первую вершину в конце не повторяют.
'     В конце AddVertex можно задать outgoingEdgeID — имя прямого участка к следующей вершине между
'     касаниями, и cornerID — имя скругления этой вершины. Например: polygon.AddVertex "TR", 400#, 600#,
'     40#, outgoingEdgeID:="Right.UpperSlope", cornerID:="Round.TopRight". VertexID при этом остаётся
'     "TR". Пустые аргументы дают прежние Edge.TR и Corner.TR. OutgoingEdgeID/CornerID возвращают
'     эффективные имена. Все имена сегментов одного многоугольника уникальны, включая исчезнувшие
'     участки; при R=0 фиктивная дуга не создаётся.
'     PathID — имя всей линии; SegmentID — локальное имя её отрезка или дуги. В разных линиях
'     допускаются одинаковые локальные имена. LoopID задаётся в Outer/Opening и обозначает границу
'     бетона или отверстия. ContourID — имя сегмента при сохранении контура: PreserveSegmentIDs берёт
'     реальные имена выбранной линии, SequentialByKind использует прежние префиксы _LINE_/_ARC_ с единым
'     счётчиком и порядком. Явный массив contourIDs меняет только публикацию. Ни роль границы, ни
'     публикация не переименовывают исходную линию.
'     Chain получает имя исходной линии, локальные имена начального и конечного сегментов, параметры от
'     0 до 1 и направление обхода. Параметры задают точки начала и конца внутри этих сегментов; True
'     выбирает обход исходной линии, False — обратный. На замкнутом контуре это два разных маршрута
'     между теми же точками. Например, у нижнего сегмента PointedRect начало справа (0), конец слева
'     (1): Chain от 1 к 0 с False идёт напрямую слева направо, а с True — через остальные стороны
'     контура. При полном обороте от середины сегмента его вторая часть получает свободное имя .Part2,
'     .Part3 и далее. SourcePathID/SourceSegmentID и SourceStartParameter/SourceEndParameter сохраняют
'     исходную линию, сегмент и параметры отдельно; обращаться к частям по угаданному суффиксу не
'     требуется.
'   ЯВНЫЕ ИМЕНА СТОРОН ПРЯМОУГОЛЬНИКА
'     Низ идёт слева направо, правая сторона снизу вверх, верх справа налево, левая сверху вниз. После
'     Transform имена относятся к исходным сторонам и не назначаются заново по глобальным X/Y. Пример
'     выбирает часть контура от середины верха до середины низа по прямому обходу через левую сторону.
'     AdjacentRectangles сохраняет прежние имена: у сложной границы может быть несколько верхних
'     участков.
'       Dim facade As CSectionDefinitionBuilder, outline As CGeometryPath, side As CGeometryPath
'       Set facade = New CSectionDefinitionBuilder: facade.Initialize "NamedRectangle"
'       Set outline = facade.RectanglePath("Boundary", 0#, 0#, 400#, 600#, _
'           bottomEdgeID:="Bottom", topEdgeID:="Top")
'       Set side = facade.Chain("LeftSide", outline.ID, "Top", 0.5, "Bottom", 0.5)
'       facade.Outer "OUTER", outline.ID
'
'     Исходная вершина до скругления может не лежать на контуре. PathAnchor задаёт точку готового
'     сегмента, Anchor — заданную координату. Вершины не превращаются в anchors автоматически.
'     OffsetRoundedPolygon: плюс внутрь исходного многоугольника; Offset: плюс влево от направления
'     готовой линии. У группы независимых линий Offset отсутствует.
'   ВЫБОР УЧАСТКА ПО КООРДИНАТАМ
'     ResolveAnchor находит существующий ID и параметр точки без знания автоматических имён сегментов.
'     Координаты в мм должны лежать на готовой линии. В общем стыке нескольких сегментов возникает
'     ошибка; выберите внутреннюю точку или задайте ID явно. Это поиск принадлежности, а не проекция
'     ближайшей точки.
'     Пример выбирает правую грань прямоугольника и создаёт участок сверху вниз. t — положение середины;
'     в Chain явно задаются концы 1 и 0, False означает обратный обход.
'       Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, rightPath As CGeometryPath
'       Dim rightID As String, t As Double
'       Set facade = New CSectionDefinitionBuilder: facade.Initialize "AnchorExample"
'       Set path = facade.RectanglePath("Rect", 0#, 0#, 400#, 600#)
'       path.ResolveAnchor 400#, 300#, rightID, t
'       Set rightPath = facade.Chain("Right", path.ID, rightID, 1#, rightID, 0#, False)
'
'     Aligned измеряет истинное расстояние, его линия параллельна измеряемому участку. Горизонтальный и
'     вертикальный размеры измеряют проекции; PathLength суммирует прямые и дуги. Поворот текста
'     остаётся прежним, 0/90 градусов.
'
'   ПОЛНЫЙ ПРИМЕР: СКРУГЛЕНИЯ, АРМАТУРА И НАКЛОННЫЙ РАЗМЕР
'     Это самостоятельная функция; b/h/w и все объекты объявлены внутри. Edge.TR — реальная ненулевая
'     наклонная грань между точками касания. Армирование учебное, без проверки достаточности.
'
'       Private Function BuildAudit02DemoDefinition() As CSectionDefinition
'           Dim facade As CSectionDefinitionBuilder, polygon As CPolygonSpec
'           Dim shapePath As CGeometryPath, offsetPath As CGeometryPath, rightAxis As CGeometryPath
'           Dim options As CRebarRuleOptions, b As Double, h As Double, w As Double
'           b = 600#: h = 1000#: w = 200#
'           Set polygon = New CPolygonSpec
'           polygon.AddVertex "TL", -b / 2#, h / 2#, 80#, outgoingEdgeID:="Edge.TL" ' Верхняя грань для привязки бокового ряда, мм.
'           polygon.AddVertex "TR", b / 2#, h / 2#, 120#, outgoingEdgeID:="Edge.TR" ' Наклонная грань для размерной аннотации, мм.
'           polygon.AddVertex "RightNose", b / 2# + w, 0#, 80#
'           polygon.AddVertex "BR", b / 2#, -h / 2#, 120#, outgoingEdgeID:="Edge.BR" ' Нижняя грань для конца бокового ряда, мм.
'           polygon.AddVertex "BL", -b / 2#, -h / 2#, 80#, outgoingEdgeID:="Edge.BL" ' Левая грань от нижнего угла к верхнему, мм.
'           polygon.Seal
'           Set facade = New CSectionDefinitionBuilder: facade.Initialize "Audit02Demo"
'           Set shapePath = facade.RoundedPolygon("Boundary", polygon)
'           facade.Outer "OUTER", shapePath.ID
'           Set offsetPath = facade.OffsetRoundedPolygon("Offset.H.Right", polygon, 50#)
'           Set rightAxis = facade.Chain("H.Right", offsetPath.ID, "Edge.TL", 1#, "Edge.BR", 0#)
'           Set options = New CRebarRuleOptions
'           options.AnnotationGroup = "H.Right" ' В одной подписи группы суммируем стержни каждого диаметра, включая все её ряды.
'           options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
'           options.Coincidence = KeepPublishedOwner ' При совпадении центров сохраняем стержень, который был добавлен раньше.
'           options.EndMode = ExcludeBoth: options.StartSetback = 100#: options.EndSetback = 150#
'           options.IncomingTangentAtJoint = True
'           facade.RebarByCount "H.Right.Main", rightAxis.ID, 15, 20#, options
'           facade.RebarRowExtentLabel "H.Right", "H.Right.Main", 1#, 0#, True
'           facade.PathAnchor "Slope.A", shapePath.ID, "Edge.TR", 0#
'           facade.PathAnchor "Slope.B", shapePath.ID, "Edge.TR", 1#
'           facade.Dimension "DIM_SLOPE", "Slope.A", "Slope.B", "L", "Aligned"
'           facade.HorizontalDistance "DIM_B", "B", -b / 2#, b / 2#, h / 2#, 1#
'           facade.VerticalExtent "DIM_H", shapePath.ID, "H", -b / 2#, -1#
'           Set BuildAudit02DemoDefinition = facade.Finish
'       End Function
'
'   ПОЛНЫЙ ПРИМЕР: ПРИМИТИВЫ, СМЕЩЕНИЯ И ПРОЕКЦИЯ
'     Самостоятельная функция с постоянными данными в мм: прямоугольники смежные, отверстие внутри
'     Outer, линии независимые. RowNormal=1 направлен влево от MainAxis; луч проекции (0,1) пересекает
'     ProjectedAxis. Примеры можно поместить в тот же класс и вызвать вместо прямоугольного описания.
'
'       Private Function BuildAudit02OperationsDefinition() As CSectionDefinition
'           Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, options As CRebarRuleOptions
'           Set facade = New CSectionDefinitionBuilder: facade.Initialize "OperationsExample"
'           Set path = facade.AdjacentRectangles("Boundary", Array(Array(0#, 0#, 1000#, 1000#), Array(1000#, 0#, 200#, 600#)))
'           facade.Outer "Outer", path.ID
'           Set path = facade.CirclePath("Hole", 900#, 800#, 50#): facade.Opening "Opening", path.ID, "Outer"
'           Set path = facade.LinePath("MainAxis", 100#, 100#, 700#, 100#)
'           Set path = facade.ArcPath("Arc", 500#, 600#, 100#, 0#, GEOM_PI / 2#)
'           Set path = facade.Offset("ProjectedAxis", "MainAxis", 100#)
'           Set path = facade.Transform("Rotated", "MainAxis", 900#, 100#, GEOM_PI / 2#, True)
'           Set path = facade.ConstrainLine("Limited", "MainAxis", 1#, 0#, 150#, 1#, 0#, 650#)
'           Set path = facade.LinePath("MaxStepAxis", 100#, 300#, 700#, 300#)
'           Set path = facade.LinePath("FixedStepAxis", 100#, 400#, 700#, 400#)
'           Set path = facade.LinePath("ExplicitAxis", 100#, 500#, 700#, 500#)
'           facade.PathGroup "TwoLines", Array("MaxStepAxis", "FixedStepAxis")
'           Set options = New CRebarRuleOptions
'           options.AnnotationGroup = "Bottom" ' В одной подписи группы суммируем стержни каждого диаметра, включая все её ряды.
'           options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
'           options.Coincidence = KeepPublishedOwner ' При совпадении центров сохраняем стержень, который был добавлен раньше.
'           facade.RebarByCount "Main", "MainAxis", 5, 20#, options
'           Set options = New CRebarRuleOptions
'           options.AnnotationGroup = "Bottom" ' В одной подписи группы суммируем стержни каждого диаметра, включая все её ряды.
'           options.Publication = PublishPerSourcePosition ' В списке стержней нумеруем сначала основной, затем связанные с ним; координаты не меняются.
'           options.Coincidence = KeepPublishedOwner ' При совпадении центров сохраняем стержень, который был добавлен раньше.
'           options.ConfigureDependentContact 20#, 16#, RowNormal, 1#, False
'           facade.RebarDependent "Contact", "Main", 16#, options
'           Set options = New CRebarRuleOptions
'           options.NormalOffset = 40#: options.TangentOffset = 10#
'           facade.RebarDependent "Shifted", "Main", 12#, options
'           Set options = New CRebarRuleOptions: options.DirectionY = 1#
'           facade.RebarProjected "Projected", "ProjectedAxis", "Main", 16#, options
'           facade.RebarByMaxStep "MaxStep", "MaxStepAxis", 150#, 16#
'           facade.RebarByFixedStep "FixedStep", "FixedStepAxis", 150#, 16#
'           facade.RebarAtStations "Explicit", "ExplicitAxis", Array(0#, 150#, 600#), 12#
'           facade.Anchor "A", 100#, 100#: facade.Anchor "B", 700#, 500#
'           facade.TransformAnchor "MovedA", "A", 0#, 700#
'           facade.Dimension "H", "A", "B", "X", "Horizontal"
'           facade.Dimension "V", "A", "B", "Y", "Vertical"
'           facade.Dimension "L", "A", "B", "L", "Aligned"
'           facade.PathLength "Length", "TwoLines", "MovedA", "SumL"
'           facade.Radius "Radius", "Arc", "Arc", 0.5, "R"
'           facade.Diameter "Diameter", "Hole", "Hole", 0.5, "D"
'           facade.Text "Note", "MovedA", "Учебный пример"
'           facade.RebarRowLabel "Bottom", "Main", 0#, -1#
'           Set BuildAudit02OperationsDefinition = facade.Finish
'       End Function
'
'
'
' ====================================================================================================
' 01  НАЧАЛО И ЗАВЕРШЕНИЕ ОПИСАНИЯ
' ====================================================================================================
'
'
'   Класс / модуль: ISectionRecipe
'
'     ----------------------------------------------------------------------------------------------
'     ISectionRecipe.CreateDefinition
'
'       Обязательный метод класса нового типа. Возвращает завершённое описание сечения. Сетку, стержни и
'       рисунок программа строит по этому описанию автоматически.
'
'       Объявление:
'         Public Function CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
'
'       Параметры:
'         ByVal context As CSectionRecipeContext
'           Объект с копиями входных таблиц и системой единиц; ячейки Excel из рецепта не читаются.
'
'
'
'   Класс / модуль: CSectionDefinitionBuilder
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Initialize
'
'       Открывает один новый фасад с TypeID и версией схемы. Повторная инициализация запрещена.
'
'       Объявление:
'         Public Sub Initialize(ByVal typeID As String, Optional ByVal schemaVersion As Long = 1)
'
'       Параметры:
'         ByVal typeID As String
'           Уникальное имя зарегистрированного типа; непустой текст.
'
'         Optional ByVal schemaVersion As Long = 1
'           Положительная целая версия схемы определения.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Finish
'
'       Проверяет контуры, ссылки, зависимости и декларации аннотаций; публикует неизменяемое
'       CSectionDefinition и закрывает фасад. Допустимость параметров и фактическое размещение арматуры
'       проверяет CSectionBuilder до построения сетки.
'
'       Объявление:
'         Public Function Finish() As CSectionDefinition
'
'
'   Класс / модуль: CSectionTypeRegistry
'
'     ----------------------------------------------------------------------------------------------
'     CSectionTypeRegistry.RegisterRecipe
'
'       Регистрирует один ISectionRecipe и массив имён его входных блоков; ID типа уникальный.
'
'       Объявление:
'         Public Sub RegisterRecipe(ByVal typeID As String, ByVal recipe As ISectionRecipe, ByVal
'         tableNames As Variant)
'
'       Параметры:
'         ByVal typeID As String
'           Уникальное имя зарегистрированного типа; непустой текст.
'
'         ByVal recipe As ISectionRecipe
'           Один объект с реализацией ISectionRecipe.
'
'         ByVal tableNames As Variant
'           Массив имён блоков выбранного типа; Empty-array допустим для программного рецепта.
'
'
'
'
' ====================================================================================================
' 02  ВХОДНЫЕ РАЗМЕРЫ И ЕДИНИЦЫ
' ====================================================================================================
'
'
'   Класс / модуль: CSectionRecipeContext
'
'     ----------------------------------------------------------------------------------------------
'     CSectionRecipeContext.Reader
'
'       Создаёт ридер блочной таблицы по её имени. Поля находятся по ключам блока, строки и столбца;
'       длины переводятся в мм по единицам контекста. Для собственной структуры таблицы доступен Table.
'
'       Объявление:
'         Public Function Reader(ByVal tableName As String, Optional ByVal rowPrefix As String =
'         vbNullString) As CSectionInputReader
'
'       Параметры:
'         ByVal tableName As String
'           Имя единственного корневого входного блока этого снимка.
'
'         Optional ByVal rowPrefix As String = vbNullString
'           Начальная часть ключей строк общей таблицы, например Circle.; пустая строка означает чтение
'           полных ключей. Префикс добавляется при поиске строки и удаляется из результата RowKeys.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionRecipeContext.Table
'
'       Возвращает копию зарегистрированной таблицы по имени; если таблицы нет, выдаёт ошибку.
'
'       Объявление:
'         Public Function Table(ByVal tableName As String) As CSectionInputTable
'
'       Параметры:
'         ByVal tableName As String
'           Имя единственного корневого входного блока этого снимка.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionRecipeContext.Units
'
'       Закрытая копия системы единиц этого снимка ввода. Внешняя перенастройка units не меняет
'       контекст; методы преобразования доступны, InitializeDefaults/LoadFromSettings запрещены.
'
'       Объявление:
'         Public Property Get Units() As CUnitSystem
'
'
'   Класс / модуль: CSectionInputReader
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputReader.Initialize
'
'       Сохраняет снимок и независимую копию единиц. Блоки индексируются только при первом обращении по
'       ключам. Блок задаётся за две строки до строки key; столбец key расположен первым. Для прямого
'       чтения ValueAt структура блоков не требуется.
'
'       Объявление:
'         Public Sub Initialize(ByVal table As CSectionInputTable, Optional ByVal units As CUnitSystem =
'         Nothing, Optional ByVal rowPrefix As String = vbNullString)
'
'       Параметры:
'         ByVal table As CSectionInputTable
'           Неизменяемый CSectionInputTable.
'
'         Optional ByVal units As CUnitSystem = Nothing
'           CUnitSystem с уже выбранными единицами; преобразование выполняется её методами.
'
'         Optional ByVal rowPrefix As String = vbNullString
'           Начальная часть ключей строк общей таблицы, например Circle.; пустая строка означает чтение
'           полных ключей. Префикс добавляется при поиске строки и удаляется из результата RowKeys.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputReader.ValueAt
'
'       Возвращает исходное значение по номерам строки и столбца внутри диапазона, начиная с 1. Работает
'       с произвольной структурой без блоков и ключей. Единицы не преобразует; ошибки формул и выход за
'       диапазон отклоняются с указанием источника.
'
'       Объявление:
'         Public Function ValueAt(ByVal row As Long, ByVal column As Long) As Variant
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputReader.RowKeys
'
'       Возвращает массив ключей строк блока в порядке таблицы; пустые строки и строка единиц Units
'       пропускаются.
'
'       Объявление:
'         Public Function RowKeys(ByVal blockKey As String) As Variant
'
'       Параметры:
'         ByVal blockKey As String
'           Короткий ключ блока, например General, Edges, Main или Extra.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputReader.RowIndex
'
'       Возвращает номер строки ключа внутри выбранного блока. Повтор одного ключа в разных блоках
'       допустим.
'
'       Объявление:
'         Public Function RowIndex(ByVal blockKey As String, ByVal rowKey As String) As Long
'
'       Параметры:
'         ByVal blockKey As String
'           Короткий ключ блока, например General, Edges, Main или Extra.
'
'         ByVal rowKey As String
'           Ключ строки внутри выбранного блока, например B или Top.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputReader.ColumnIndex
'
'       Возвращает номер столбца по короткому ключу, независимо от русской подписи и порядка столбцов
'       данных.
'
'       Объявление:
'         Public Function ColumnIndex(ByVal blockKey As String, ByVal columnKey As String) As Long
'
'       Параметры:
'         ByVal blockKey As String
'           Короткий ключ блока, например General, Edges, Main или Extra.
'
'         ByVal columnKey As String
'           Ключ столбца внутри выбранного блока; val означает значение скалярного параметра.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputReader.CellValue
'
'       Находит поле по ключам блока, строки и столбца. Возвращает исходное значение из копии таблицы.
'       При ошибке Excel сообщает адрес ячейки, которую нужно исправить.
'
'       Объявление:
'         Public Function CellValue(ByVal blockKey As String, ByVal rowKey As String, Optional ByVal
'         columnKey As String = "val") As Variant
'
'       Параметры:
'         ByVal blockKey As String
'           Короткий ключ блока, например General, Edges, Main или Extra.
'
'         ByVal rowKey As String
'           Ключ строки внутри выбранного блока, например B или Top.
'
'         Optional ByVal columnKey As String = "val"
'           Ключ столбца внутри выбранного блока; val означает значение скалярного параметра.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputReader.Text
'
'       Находит поле по ключам блока, строки и столбца. Читает непустой текст либо переданное значение
'       по умолчанию, удаляя крайние пробелы.
'
'       Объявление:
'         Public Function Text(ByVal blockKey As String, ByVal rowKey As String, Optional ByVal
'         columnKey As String = "val", Optional ByVal defaultValue As Variant) As String
'
'       Параметры:
'         ByVal blockKey As String
'           Короткий ключ блока, например General, Edges, Main или Extra.
'
'         ByVal rowKey As String
'           Ключ строки внутри выбранного блока, например B или Top.
'
'         Optional ByVal columnKey As String = "val"
'           Ключ столбца внутри выбранного блока; val означает значение скалярного параметра.
'
'         Optional ByVal defaultValue As Variant
'           Значение для пустой ячейки. В Text и Number без этого параметра пустая ячейка считается
'           ошибкой; остальные методы используют указанное в объявлении значение по умолчанию.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputReader.Number
'
'       Находит поле по ключам блока, строки и столбца. Читает число, исключая Boolean/ошибки/объекты;
'       обязательное пустое поле не заменяется нулём.
'
'       Объявление:
'         Public Function Number(ByVal blockKey As String, ByVal rowKey As String, Optional ByVal
'         columnKey As String = "val", Optional ByVal defaultValue As Variant) As Double
'
'       Параметры:
'         ByVal blockKey As String
'           Короткий ключ блока, например General, Edges, Main или Extra.
'
'         ByVal rowKey As String
'           Ключ строки внутри выбранного блока, например B или Top.
'
'         Optional ByVal columnKey As String = "val"
'           Ключ столбца внутри выбранного блока; val означает значение скалярного параметра.
'
'         Optional ByVal defaultValue As Variant
'           Значение для пустой ячейки. В Text и Number без этого параметра пустая ячейка считается
'           ошибкой; остальные методы используют указанное в объявлении значение по умолчанию.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputReader.Count
'
'       Находит поле по ключам блока, строки и столбца. Читает целое неотрицательное 0..2147483647;
'       отдельный executor проверяет суммарное число позиций.
'
'       Объявление:
'         Public Function Count(ByVal blockKey As String, ByVal rowKey As String, Optional ByVal
'         columnKey As String = "val", Optional ByVal defaultValue As Long = 0) As Long
'
'       Параметры:
'         ByVal blockKey As String
'           Короткий ключ блока, например General, Edges, Main или Extra.
'
'         ByVal rowKey As String
'           Ключ строки внутри выбранного блока, например B или Top.
'
'         Optional ByVal columnKey As String = "val"
'           Ключ столбца внутри выбранного блока; val означает значение скалярного параметра.
'
'         Optional ByVal defaultValue As Long = 0
'           Значение для пустой ячейки. В Text и Number без этого параметра пустая ячейка считается
'           ошибкой; остальные методы используют указанное в объявлении значение по умолчанию.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputReader.Choice
'
'       Находит поле по ключам блока, строки и столбца. Проверяет вариант по allowed без учёта регистра
'       и возвращает написание из allowed.
'
'       Объявление:
'         Public Function Choice(ByVal blockKey As String, ByVal rowKey As String, ByVal allowed As
'         Variant, Optional ByVal columnKey As String = "val", Optional ByVal defaultValue As String =
'         vbNullString) As String
'
'       Параметры:
'         ByVal blockKey As String
'           Короткий ключ блока, например General, Edges, Main или Extra.
'
'         ByVal rowKey As String
'           Ключ строки внутри выбранного блока, например B или Top.
'
'         ByVal allowed As Variant
'           Массив разрешённых строк; Choice возвращает строку именно в написании из этого массива.
'
'         Optional ByVal columnKey As String = "val"
'           Ключ столбца внутри выбранного блока; val означает значение скалярного параметра.
'
'         Optional ByVal defaultValue As String = vbNullString
'           Значение для пустой ячейки. В Text и Number без этого параметра пустая ячейка считается
'           ошибкой; остальные методы используют указанное в объявлении значение по умолчанию.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputReader.Length
'
'       Находит поле по ключам блока, строки и столбца. Читает обязательное число и один раз переводит
'       input length в мм через CUnitSystem.
'
'       Объявление:
'         Public Function Length(ByVal blockKey As String, ByVal rowKey As String, Optional ByVal
'         columnKey As String = "val") As Double
'
'       Параметры:
'         ByVal blockKey As String
'           Короткий ключ блока, например General, Edges, Main или Extra.
'
'         ByVal rowKey As String
'           Ключ строки внутри выбранного блока, например B или Top.
'
'         Optional ByVal columnKey As String = "val"
'           Ключ столбца внутри выбранного блока; val означает значение скалярного параметра.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputReader.PositiveLength
'
'       Находит поле по ключам блока, строки и столбца. Читает обязательную длину существующим Length и
'       один раз переводит её в мм; проверяет >0. Пустота/текст/ошибка не становятся нулём; отказ
'       содержит фактический адрес. Не заменяет ограничения GEOM_TOLERANCE/as>d/2.
'
'       Объявление:
'         Public Function PositiveLength(ByVal blockKey As String, ByVal rowKey As String, Optional
'         ByVal columnKey As String = "val") As Double
'
'       Параметры:
'         ByVal blockKey As String
'           Короткий ключ блока, например General, Edges, Main или Extra.
'
'         ByVal rowKey As String
'           Ключ строки внутри выбранного блока, например B или Top.
'
'         Optional ByVal columnKey As String = "val"
'           Ключ столбца внутри выбранного блока; val означает значение скалярного параметра.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputReader.NonNegativeLength
'
'       Находит поле по ключам блока, строки и столбца. Читает обязательную длину существующим Length и
'       один раз переводит её в мм; проверяет >=0. Пустота/текст/ошибка не становятся нулём; отказ
'       содержит фактический адрес. Не заменяет ограничения GEOM_TOLERANCE/as>d/2.
'
'       Объявление:
'         Public Function NonNegativeLength(ByVal blockKey As String, ByVal rowKey As String, Optional
'         ByVal columnKey As String = "val") As Double
'
'       Параметры:
'         ByVal blockKey As String
'           Короткий ключ блока, например General, Edges, Main или Extra.
'
'         ByVal rowKey As String
'           Ключ строки внутри выбранного блока, например B или Top.
'
'         Optional ByVal columnKey As String = "val"
'           Ключ столбца внутри выбранного блока; val означает значение скалярного параметра.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputReader.OptionalDiameter
'
'       Находит поле по ключам блока, строки и столбца. Пустая ячейка или 0 отключают ряд; положительный
'       диаметр переводится в мм. Текст, включая одиночный дефис и TODO, и отрицательный диаметр
'       вызывают адресную ошибку ввода.
'
'       Объявление:
'         Public Function OptionalDiameter(ByVal blockKey As String, ByVal rowKey As String, Optional
'         ByVal columnKey As String = "val") As Double
'
'       Параметры:
'         ByVal blockKey As String
'           Короткий ключ блока, например General, Edges, Main или Extra.
'
'         ByVal rowKey As String
'           Ключ строки внутри выбранного блока, например B или Top.
'
'         Optional ByVal columnKey As String = "val"
'           Ключ столбца внутри выбранного блока; val означает значение скалярного параметра.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputReader.OptionalRowOffset
'
'       Находит поле по ключам блока, строки и столбца. Читает базовый отступ дополнительного ряда:
'       пустая ячейка равна нулю, число неотрицательно и переводится из INPUT Length в мм. Текст,
'       Boolean, ошибка формулы и отрицательное число дают адресную ошибку.
'
'       Объявление:
'         Public Function OptionalRowOffset(ByVal blockKey As String, ByVal rowKey As String, Optional
'         ByVal columnKey As String = "val") As Double
'
'       Параметры:
'         ByVal blockKey As String
'           Короткий ключ блока, например General, Edges, Main или Extra.
'
'         ByVal rowKey As String
'           Ключ строки внутри выбранного блока, например B или Top.
'
'         Optional ByVal columnKey As String = "val"
'           Ключ столбца внутри выбранного блока; val означает значение скалярного параметра.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputReader.SourceReference
'
'       Возвращает фактический адрес исходного поля с именем таблицы, в том числе после перестановки
'       строк и столбцов.
'
'       Объявление:
'         Public Function SourceReference(ByVal blockKey As String, ByVal rowKey As String, Optional
'         ByVal columnKey As String = "val") As String
'
'       Параметры:
'         ByVal blockKey As String
'           Короткий ключ блока, например General, Edges, Main или Extra.
'
'         ByVal rowKey As String
'           Ключ строки внутри выбранного блока, например B или Top.
'
'         Optional ByVal columnKey As String = "val"
'           Ключ столбца внутри выбранного блока; val означает значение скалярного параметра.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputReader.SourceReferences
'
'       Объединяет адреса нескольких полей одной строки для пояснения об ошибке размещения арматуры или
'       геометрии.
'
'       Объявление:
'         Public Function SourceReferences(ByVal blockKey As String, ByVal rowKey As String, ByVal
'         columnKeys As Variant) As String
'
'       Параметры:
'         ByVal blockKey As String
'           Короткий ключ блока, например General, Edges, Main или Extra.
'
'         ByVal rowKey As String
'           Ключ строки внутри выбранного блока, например B или Top.
'
'         ByVal columnKeys As Variant
'           Массив ключей полей, чьи адреса нужно включить в пояснение об ошибке.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputReader.Require
'
'       При нарушении переданного условия сообщает причину и фактический адрес поля. Инженерное условие
'       определяет рецепт.
'
'       Объявление:
'         Public Sub Require(ByVal blockKey As String, ByVal rowKey As String, ByVal columnKey As
'         String, ByVal condition As Boolean, ByVal reason As String)
'
'       Параметры:
'         ByVal blockKey As String
'           Короткий ключ блока, например General, Edges, Main или Extra.
'
'         ByVal rowKey As String
'           Ключ строки внутри выбранного блока, например B или Top.
'
'         ByVal columnKey As String
'           Ключ столбца внутри выбранного блока; val означает значение скалярного параметра.
'
'         ByVal condition As Boolean
'           Проверяемое рецептом инженерное условие; при False выдаётся ошибка ввода.
'
'         ByVal reason As String
'           Понятная причина нарушения условия, без заранее заданного адреса ячейки.
'
'
'
'   Класс / модуль: CSectionInputTable
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.Name
'
'       Имя входной таблицы.
'
'       Объявление:
'         Public Property Get Name() As String
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.RowCount
'
'       Число строк в копии входной таблицы.
'
'       Объявление:
'         Public Property Get RowCount() As Long
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.ColumnCount
'
'       Число столбцов в копии входной таблицы.
'
'       Объявление:
'         Public Property Get ColumnCount() As Long
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.SourceReference
'
'       Возвращает исходный адрес ячейки или программный tableName[row,col].
'
'       Объявление:
'         Public Function SourceReference(ByVal row As Long, ByVal column As Long) As String
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.CellValue
'
'       Возвращает исходное значение из копии таблицы. При ошибке Excel сообщает адрес ячейки, которую
'       нужно исправить.
'
'       Объявление:
'         Public Function CellValue(ByVal row As Long, ByVal column As Long) As Variant
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.IsBlank
'
'       Проверяет пустое значение снимка; ошибка Excel не подменяется пустотой.
'
'       Объявление:
'         Public Function IsBlank(ByVal row As Long, ByVal column As Long) As Boolean
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.Text
'
'       Читает непустой текст либо переданное значение по умолчанию, удаляя крайние пробелы.
'
'       Объявление:
'         Public Function Text(ByVal row As Long, ByVal column As Long, Optional ByVal defaultValue As
'         Variant) As String
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'         Optional ByVal defaultValue As Variant
'           Значение для пустой ячейки. В Text и Number без этого параметра пустая ячейка считается
'           ошибкой; остальные методы используют указанное в объявлении значение по умолчанию.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.Number
'
'       Читает число, исключая Boolean/ошибки/объекты; обязательное пустое поле не заменяется нулём.
'
'       Объявление:
'         Public Function Number(ByVal row As Long, ByVal column As Long, Optional ByVal defaultValue As
'         Variant) As Double
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'         Optional ByVal defaultValue As Variant
'           Значение для пустой ячейки. В Text и Number без этого параметра пустая ячейка считается
'           ошибкой; остальные методы используют указанное в объявлении значение по умолчанию.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.Count
'
'       Читает целое неотрицательное 0..2147483647; отдельный executor проверяет суммарное число
'       позиций.
'
'       Объявление:
'         Public Function Count(ByVal row As Long, ByVal column As Long, Optional ByVal defaultValue As
'         Long = 0) As Long
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'         Optional ByVal defaultValue As Long = 0
'           Значение для пустой ячейки. В Text и Number без этого параметра пустая ячейка считается
'           ошибкой; остальные методы используют указанное в объявлении значение по умолчанию.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.Choice
'
'       Проверяет вариант по allowed без учёта регистра и возвращает написание из allowed.
'
'       Объявление:
'         Public Function Choice(ByVal row As Long, ByVal column As Long, ByVal allowed As Variant,
'         Optional ByVal defaultValue As String = vbNullString) As String
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'         ByVal allowed As Variant
'           Массив разрешённых строк; Choice возвращает строку именно в написании из этого массива.
'
'         Optional ByVal defaultValue As String = vbNullString
'           Значение для пустой ячейки. В Text и Number без этого параметра пустая ячейка считается
'           ошибкой; остальные методы используют указанное в объявлении значение по умолчанию.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.Length
'
'       Читает обязательное число и один раз переводит input length в мм через CUnitSystem.
'
'       Объявление:
'         Public Function Length(ByVal row As Long, ByVal column As Long, ByVal units As CUnitSystem) As
'         Double
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'         ByVal units As CUnitSystem
'           CUnitSystem с уже выбранными единицами; преобразование выполняется её методами.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.PositiveLength
'
'       Читает обязательную длину существующим Length и один раз переводит её в мм; проверяет >0.
'       Пустота/текст/ошибка не становятся нулём; отказ содержит фактический адрес. Не заменяет
'       ограничения GEOM_TOLERANCE/as>d/2.
'
'       Объявление:
'         Public Function PositiveLength(ByVal row As Long, ByVal column As Long, ByVal units As
'         CUnitSystem) As Double
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'         ByVal units As CUnitSystem
'           CUnitSystem с уже выбранными единицами; преобразование выполняется её методами.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.NonNegativeLength
'
'       Читает обязательную длину существующим Length и один раз переводит её в мм; проверяет >=0.
'       Пустота/текст/ошибка не становятся нулём; отказ содержит фактический адрес. Не заменяет
'       ограничения GEOM_TOLERANCE/as>d/2.
'
'       Объявление:
'         Public Function NonNegativeLength(ByVal row As Long, ByVal column As Long, ByVal units As
'         CUnitSystem) As Double
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'         ByVal units As CUnitSystem
'           CUnitSystem с уже выбранными единицами; преобразование выполняется её методами.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.OptionalDiameter
'
'       Пустая ячейка или 0 отключают ряд; положительный диаметр переводится в мм. Текст, включая
'       одиночный дефис и TODO, и отрицательный диаметр вызывают адресную ошибку ввода.
'
'       Объявление:
'         Public Function OptionalDiameter(ByVal row As Long, ByVal column As Long, ByVal units As
'         CUnitSystem) As Double
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'         ByVal units As CUnitSystem
'           CUnitSystem с уже выбранными единицами; преобразование выполняется её методами.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.OptionalRowOffset
'
'       Читает базовый отступ дополнительного ряда: пустая ячейка равна нулю, число неотрицательно и
'       переводится из INPUT Length в мм. Текст, Boolean, ошибка формулы и отрицательное число дают
'       адресную ошибку.
'
'       Объявление:
'         Public Function OptionalRowOffset(ByVal row As Long, ByVal column As Long, ByVal units As
'         CUnitSystem) As Double
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'         ByVal units As CUnitSystem
'           CUnitSystem с уже выбранными единицами; преобразование выполняется её методами.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.ColumnByCaption
'
'       Находит уникальный столбец заголовка, без фиксированного адреса.
'
'       Объявление:
'         Public Function ColumnByCaption(ByVal headerRow As Long, ByVal caption As String) As Long
'
'       Параметры:
'         ByVal headerRow As Long
'           Строка заголовков, в пределах снимка.
'
'         ByVal caption As String
'           Название столбца; повторные заголовки — ошибка.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.RowByKey
'
'       Находит единственную строку с заданным ключом в указанной части таблицы.
'
'       Объявление:
'         Public Function RowByKey(ByVal key As String, Optional ByVal keyColumn As Long = 1, Optional
'         ByVal firstRow As Long = 1, Optional ByVal lastRow As Long = 0) As Long
'
'       Параметры:
'         ByVal key As String
'           Название параметра, по которому ищется строка.
'
'         Optional ByVal keyColumn As Long = 1
'           Номер столбца с названиями параметров: от 1 до ColumnCount.
'
'         Optional ByVal firstRow As Long = 1
'           Начало диапазона поиска строки, включительно.
'
'         Optional ByVal lastRow As Long = 0
'           Конец поиска, включительно; 0 означает RowCount.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.RowByAnyKey
'
'       Находит уникальную строку по массиву допустимых альтернативных ключей; неоднозначность
'       запрещена.
'
'       Объявление:
'         Public Function RowByAnyKey(ByVal keys As Variant, Optional ByVal keyColumn As Long = 1,
'         Optional ByVal firstRow As Long = 1, Optional ByVal lastRow As Long = 0) As Long
'
'       Параметры:
'         ByVal keys As Variant
'           Массив допустимых альтернативных ключей; неоднозначное совпадение отклоняется.
'
'         Optional ByVal keyColumn As Long = 1
'           Номер столбца с названиями параметров: от 1 до ColumnCount.
'
'         Optional ByVal firstRow As Long = 1
'           Начало диапазона поиска строки, включительно.
'
'         Optional ByVal lastRow As Long = 0
'           Конец поиска, включительно; 0 означает RowCount.
'
'
'
'   Класс / модуль: CUnitSystem
'
'     ----------------------------------------------------------------------------------------------
'     CUnitSystem.InputLengthToInternal
'
'       Переводит InputLengthToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с
'       принятой конвенцией знаков для N/Mx/My.
'
'       Объявление:
'         Public Function InputLengthToInternal(ByVal value As Double) As Double
'
'       Параметры:
'         ByVal value As Double
'           Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный
'           перевод.
'
'
'     ----------------------------------------------------------------------------------------------
'     CUnitSystem.InputAreaToInternal
'
'       Переводит InputAreaToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с
'       принятой конвенцией знаков для N/Mx/My.
'
'       Объявление:
'         Public Function InputAreaToInternal(ByVal value As Double) As Double
'
'       Параметры:
'         ByVal value As Double
'           Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный
'           перевод.
'
'
'
'
' ====================================================================================================
' 03  ГРАНИЦЫ БЕТОНА И ОТВЕРСТИЙ
' ====================================================================================================
'
'
'   Класс / модуль: CPolygonSpec
'
'     ----------------------------------------------------------------------------------------------
'     CPolygonSpec.AddVertex
'
'       Добавляет исходную вершину, её координаты, радиус, источник и необязательные имена прямого
'       участка и скругления. Имя вершины независимо от имён сегментов. Пустые имена дают
'       Edge.<vertexID>/Corner.<vertexID>. Повторы эффективных имён отклоняются сразу, включая
'       исчезающие участки.
'
'       Объявление:
'         Public Sub AddVertex(ByVal vertexID As String, ByVal x As Double, ByVal y As Double, Optional
'         ByVal filletRadius As Double = 0#, Optional ByVal sourceReference As String = "", Optional
'         ByVal outgoingEdgeID As String = "", Optional ByVal cornerID As String = "")
'
'       Параметры:
'         ByVal vertexID As String
'           Непустое уникальное имя исходной вершины; регистр различается, повтор первой вершины для
'           замыкания не нужен.
'
'         ByVal x As Double
'           Модельная координата x, мм; знаки соответствуют section XY/WCS.
'
'         ByVal y As Double
'           Модельная координата y, мм; знаки соответствуют section XY/WCS.
'
'         Optional ByVal filletRadius As Double = 0#
'           Неотрицательный конечный радиус сопряжения, мм. Ноль задаёт острый угол; фактические точки
'           касания вычисляет общее ядро.
'
'         Optional ByVal sourceReference As String = ""
'           Строка исходного адреса/программного параметра; без объекта Excel.
'
'         Optional ByVal outgoingEdgeID As String = ""
'           Локальное имя прямого участка от данной вершины к следующей между касаниями скруглений.
'           Пустое даёт Edge.<vertexID>. Имя не должно повторять имя другого участка или скругления
'           этого многоугольника, даже исчезнувшего.
'
'         Optional ByVal cornerID As String = ""
'           Локальное имя скругления данной вершины. Пустое даёт Corner.<vertexID>; при нулевом радиусе
'           имя обозначает только привязку. Имена сегментов чувствительны к регистру и уникальны в
'           пределах линии.
'
'
'     ----------------------------------------------------------------------------------------------
'     CPolygonSpec.Seal
'
'       Завершает набор из минимум трёх вершин в исходном порядке. Повторный Seal безопасен; неудачный
'       оставляет описание открытым.
'
'       Объявление:
'         Public Sub Seal()
'
'     ----------------------------------------------------------------------------------------------
'     CPolygonSpec.Count
'
'       Число исходных вершин, доступное при заполнении и после Seal; не ограничено шестью.
'
'       Объявление:
'         Public Property Get Count() As Long
'
'     ----------------------------------------------------------------------------------------------
'     CPolygonSpec.IsSealed
'
'       True после успешного явного Seal; только такой набор принимает фасад.
'
'       Объявление:
'         Public Property Get IsSealed() As Boolean
'
'     ----------------------------------------------------------------------------------------------
'     CPolygonSpec.VertexID
'
'       Имя исходной вершины по номеру 1..Count, доступное после Seal.
'
'       Объявление:
'         Public Function VertexID(ByVal index As Long) As String
'
'       Параметры:
'         ByVal index As Long
'           1-based индекс в пределах соответствующего готового контейнера.
'
'
'     ----------------------------------------------------------------------------------------------
'     CPolygonSpec.X
'
'       Исходная координата X, мм, по номеру 1..Count; не является автоматически точкой готового
'       контура.
'
'       Объявление:
'         Public Function X(ByVal index As Long) As Double
'
'       Параметры:
'         ByVal index As Long
'           1-based индекс в пределах соответствующего готового контейнера.
'
'
'     ----------------------------------------------------------------------------------------------
'     CPolygonSpec.Y
'
'       Исходная координата Y, мм, по номеру 1..Count.
'
'       Объявление:
'         Public Function Y(ByVal index As Long) As Double
'
'       Параметры:
'         ByVal index As Long
'           1-based индекс в пределах соответствующего готового контейнера.
'
'
'     ----------------------------------------------------------------------------------------------
'     CPolygonSpec.FilletRadius
'
'       Назначенный исходный радиус, мм, по номеру 1..Count; ноль — острый угол.
'
'       Объявление:
'         Public Function FilletRadius(ByVal index As Long) As Double
'
'       Параметры:
'         ByVal index As Long
'           1-based индекс в пределах соответствующего готового контейнера.
'
'
'     ----------------------------------------------------------------------------------------------
'     CPolygonSpec.SourceReference
'
'       Скалярный текст происхождения вершины по номеру 1..Count; ссылки на Excel-объекты не хранятся.
'
'       Объявление:
'         Public Function SourceReference(ByVal index As Long) As String
'
'       Параметры:
'         ByVal index As Long
'           1-based индекс в пределах соответствующего готового контейнера.
'
'
'     ----------------------------------------------------------------------------------------------
'     CPolygonSpec.OutgoingEdgeID
'
'       Эффективное имя прямого участка от этой вершины к следующей, по номеру 1..Count; доступно после
'       Seal и при исчезновении участка.
'
'       Объявление:
'         Public Function OutgoingEdgeID(ByVal index As Long) As String
'
'       Параметры:
'         ByVal index As Long
'           1-based индекс в пределах соответствующего готового контейнера.
'
'
'     ----------------------------------------------------------------------------------------------
'     CPolygonSpec.CornerID
'
'       Эффективное имя скругления этой вершины, по номеру 1..Count; доступно после Seal также при
'       нулевом радиусе.
'
'       Объявление:
'         Public Function CornerID(ByVal index As Long) As String
'
'       Параметры:
'         ByVal index As Long
'           1-based индекс в пределах соответствующего готового контейнера.
'
'
'
'   Класс / модуль: modSectionRecipeTypes
'
'     ----------------------------------------------------------------------------------------------
'     modSectionRecipeTypes.EContourNamingPolicy
'
'       PreserveSegmentIDs=0: исходные/явные ID; SequentialByKind=1: технический префикс_LINE_i/ARC_i с
'       прежним общим номером сегмента.
'
'       Объявление:
'         Public Enum EContourNamingPolicy
'
'       Допустимые значения / состав:
'         PreserveSegmentIDs = 0
'         SequentialByKind = 1
'
'
'   Класс / модуль: CSectionDefinitionBuilder
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.CirclePath
'
'       Создаёт замкнутую окружность CCW от точки +X; один сегмент с тем же ID, что путь.
'
'       Объявление:
'         Public Function CirclePath(ByVal pathID As String, ByVal centerX As Double, ByVal centerY As
'         Double, ByVal radius As Double) As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal centerX As Double
'           Модельная координата centerX, мм; знаки соответствуют section XY/WCS.
'
'         ByVal centerY As Double
'           Модельная координата centerY, мм; знаки соответствуют section XY/WCS.
'
'         ByVal radius As Double
'           Положительный радиус, мм.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.LinePath
'
'       Создаёт непрерывный открытый путь из одной прямой от первого конца ко второму.
'
'       Объявление:
'         Public Function LinePath(ByVal pathID As String, ByVal x1 As Double, ByVal y1 As Double, ByVal
'         x2 As Double, ByVal y2 As Double) As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal x1 As Double
'           Модельная координата x1, мм; знаки соответствуют section XY/WCS.
'
'         ByVal y1 As Double
'           Модельная координата y1, мм; знаки соответствуют section XY/WCS.
'
'         ByVal x2 As Double
'           Модельная координата x2, мм; знаки соответствуют section XY/WCS.
'
'         ByVal y2 As Double
'           Модельная координата y2, мм; знаки соответствуют section XY/WCS.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.ArcPath
'
'       Создаёт точную дугу по центру/радиусу/startAngle/signed sweep; полный оборот замкнут.
'
'       Объявление:
'         Public Function ArcPath(ByVal pathID As String, ByVal centerX As Double, ByVal centerY As
'         Double, ByVal radius As Double, ByVal startAngle As Double, ByVal sweep As Double) As
'         CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal centerX As Double
'           Модельная координата centerX, мм; знаки соответствуют section XY/WCS.
'
'         ByVal centerY As Double
'           Модельная координата centerY, мм; знаки соответствуют section XY/WCS.
'
'         ByVal radius As Double
'           Положительный радиус, мм.
'
'         ByVal startAngle As Double
'           Начальный угол, радианы; от +X.
'
'         ByVal sweep As Double
'           Signed sweep, радианы: CCW положителен, CW отрицателен, 0 запрещён, |sweep|<=2*pi.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RectanglePath
'
'       Создаёт замкнутый прямоугольник от (x,y) с положительными width/height. Необязательные имена
'       относятся к исходным сторонам: низ слева направо, правая снизу вверх, верх справа налево, левая
'       сверху вниз. Имена сопоставляются фактическим сторонам без перестановки сегментов. Пустые
'       аргументы сохраняют прежние pathID.i; после поворота или отражения имена сохраняют исходный
'       смысл.
'
'       Объявление:
'         Public Function RectanglePath(ByVal pathID As String, ByVal x As Double, ByVal y As Double,
'         ByVal width As Double, ByVal height As Double, Optional ByVal bottomEdgeID As String = "",
'         Optional ByVal rightEdgeID As String = "", Optional ByVal topEdgeID As String = "", Optional
'         ByVal leftEdgeID As String = "") As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal x As Double
'           Модельная координата x, мм; знаки соответствуют section XY/WCS.
'
'         ByVal y As Double
'           Модельная координата y, мм; знаки соответствуют section XY/WCS.
'
'         ByVal width As Double
'           Координата/размер width, мм; размеры width/height положительны.
'
'         ByVal height As Double
'           Координата/размер height, мм; размеры width/height положительны.
'
'         Optional ByVal bottomEdgeID As String = ""
'           Имя нижней стороны исходного прямоугольника, направленной слева направо; пустое сохраняет
'           прежнее pathID.i.
'
'         Optional ByVal rightEdgeID As String = ""
'           Имя правой стороны исходного прямоугольника, направленной снизу вверх; пустое сохраняет
'           прежнее pathID.i.
'
'         Optional ByVal topEdgeID As String = ""
'           Имя верхней стороны исходного прямоугольника, направленной справа налево; пустое сохраняет
'           прежнее pathID.i.
'
'         Optional ByVal leftEdgeID As String = ""
'           Имя левой стороны исходного прямоугольника, направленной сверху вниз; пустое сохраняет
'           прежнее pathID.i.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RoundedPolygon
'
'       Создаёт и регистрирует точный контур по завершённому CPolygonSpec с именами из AddVertex. Прямой
'       участок идёт от касания у данной вершины к касанию у следующей; скругление относится к самой
'       вершине. Исчезнувший участок сохраняет только именованную привязку. Исходное описание не
'       меняется.
'
'       Объявление:
'         Public Function RoundedPolygon(ByVal pathID As String, ByVal polygon As CPolygonSpec) As
'         CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal polygon As CPolygonSpec
'           Завершённый CPolygonSpec исходных вершин до скругления. Nothing и не закрытый через Seal
'           набор отклоняются; порядок, радиусы и источники сохраняются.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.AdjacentRectangles
'
'       Создаёт одну границу смежных неперекрывающихся прямоугольников. Произвольное пересекающееся CAD
'       Union не выполняется.
'
'       Объявление:
'         Public Function AdjacentRectangles(ByVal pathID As String, ByVal rectangles As Variant) As
'         CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal rectangles As Variant
'           Array(Array(x,y,width,height),...); мм, размеры положительные, прямоугольники смежные без
'           перекрытий.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.AddPath
'
'       Регистрирует непрерывную линию, созданную вручную или отдельным общим инструментом. Создающие
'       команды фасада уже регистрируют свой результат; повторный AddPath для него является ошибкой.
'       Назначение Outer/Opening выполняется отдельно.
'
'       Объявление:
'         Public Sub AddPath(ByVal path As CGeometryPath)
'
'       Параметры:
'         ByVal path As CGeometryPath
'           Уже добавленная непрерывная линия CGeometryPath.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Outer
'
'       Назначает замкнутому пути роль наружной границы материала. Не требует наличия арматуры.
'
'       Объявление:
'         Public Sub Outer(ByVal loopID As String, ByVal pathID As String, Optional ByVal comment As
'         String = "", Optional ByVal contourIDs As Variant, Optional ByVal arcComment As String = "",
'         Optional ByVal namingPolicy As EContourNamingPolicy = PreserveSegmentIDs, Optional ByVal
'         namingPrefix As String = "")
'
'       Параметры:
'         ByVal loopID As String
'           Непустой устойчивый ID контура; у CConcreteRegion численный индекс loop 1..LoopCount.
'
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'         Optional ByVal contourIDs As Variant
'           Необязательный массив ID сегментов сохранённого контура, по одному на каждый сегмент пути.
'
'         Optional ByVal arcComment As String = ""
'           Отдельный комментарий аналитических дуг; пустой использует обычный comment.
'
'         Optional ByVal namingPolicy As EContourNamingPolicy = PreserveSegmentIDs
'           PreserveSegmentIDs сохраняет имена сегментов/явные contourIDs; SequentialByKind создаёт
'           префикс_LINE_i/ARC_i с общим номером сегмента. Явные ID несовместимы с SequentialByKind.
'
'         Optional ByVal namingPrefix As String = ""
'           Префикс только SequentialByKind; пустой использует CONTOUR. Разным границам нужны разные
'           префиксы; рецепты Hollow явно задают CONTOUR_OUTER/CONTOUR_OPENING.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Opening
'
'       Назначает замкнутому пути роль отверстия с конкретным parentOuterID. Действительное вложение
'       проверяется при Finish.
'
'       Объявление:
'         Public Sub Opening(ByVal loopID As String, ByVal pathID As String, ByVal parentOuterID As
'         String, Optional ByVal comment As String = "", Optional ByVal contourIDs As Variant, Optional
'         ByVal arcComment As String = "", Optional ByVal namingPolicy As EContourNamingPolicy =
'         PreserveSegmentIDs, Optional ByVal namingPrefix As String = "")
'
'       Параметры:
'         ByVal loopID As String
'           Непустой устойчивый ID контура; у CConcreteRegion численный индекс loop 1..LoopCount.
'
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal parentOuterID As String
'           ID ранее объявленной наружной границы для Opening; проверяется реальное вложение.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'         Optional ByVal contourIDs As Variant
'           Необязательный массив ID сегментов сохранённого контура, по одному на каждый сегмент пути.
'
'         Optional ByVal arcComment As String = ""
'           Отдельный комментарий аналитических дуг; пустой использует обычный comment.
'
'         Optional ByVal namingPolicy As EContourNamingPolicy = PreserveSegmentIDs
'           PreserveSegmentIDs сохраняет имена сегментов/явные contourIDs; SequentialByKind создаёт
'           префикс_LINE_i/ARC_i с общим номером сегмента. Явные ID несовместимы с SequentialByKind.
'
'         Optional ByVal namingPrefix As String = ""
'           Префикс только SequentialByKind; пустой использует CONTOUR. Разным границам нужны разные
'           префиксы; рецепты Hollow явно задают CONTOUR_OUTER/CONTOUR_OPENING.
'
'
'
'
' ====================================================================================================
' 04  ЛИНИИ ЦЕНТРОВ СТЕРЖНЕЙ И ПРЕОБРАЗОВАНИЯ
' ====================================================================================================
'
'   Сначала задайте контур, затем объявите смещение на нужное as. Для второго и третьего независимого
'   ряда задайте свои линии и свои as; связь с первым рядом не обязательна.
'
'
'   Класс / модуль: CSectionDefinitionBuilder
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Offset
'
'       Строит смещённую линию: положительное расстояние направлено влево по обходу. Соседние прямые
'       соединяются пересечением продолжений. Скругление между касательными прямыми при исчерпании
'       радиуса исчезает; вместо него остаётся стык прямых. Для отдельной дуги или полного круга такая
'       замена невозможна.
'
'       Объявление:
'         Public Function Offset(ByVal pathID As String, ByVal sourcePathID As String, ByVal
'         signedLeftDistance As Double) As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal sourcePathID As String
'           ID ранее объявленного пути, который читается без изменения.
'
'         ByVal signedLeftDistance As Double
'           Поперечное смещение, мм; положительно слева от направления пути, отрицательно справа.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.OffsetRoundedPolygon
'
'       Создаёт и регистрирует смещённый скруглённый многоугольник по исходному CPolygonSpec, не меняя
'       его. Положительное расстояние внутрь площади; отрицательное наружу, из Opening в бетон.
'       Сохраняет привязки исчезнувших участков; отличается от signed-left Offset готового пути.
'
'       Объявление:
'         Public Function OffsetRoundedPolygon(ByVal pathID As String, ByVal polygon As CPolygonSpec,
'         ByVal distance As Double) As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal polygon As CPolygonSpec
'           Завершённый CPolygonSpec исходных вершин до скругления. Nothing и не закрытый через Seal
'           набор отклоняются; порядок, радиусы и источники сохраняются.
'
'         ByVal distance As Double
'           Signed offset, мм: положительно внутрь площади, отрицательно наружу; это не left-offset
'           обхода.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.ConstrainLine
'
'       Ограничивает или продлевает одну прямую двумя поперечными плоскостями nX*x+nY*y=bound;
'       направление исходного пути сохраняется.
'
'       Объявление:
'         Public Function ConstrainLine(ByVal pathID As String, ByVal sourcePathID As String, ByVal
'         startNX As Double, ByVal startNY As Double, ByVal startBound As Double, ByVal endNX As Double,
'         ByVal endNY As Double, ByVal endBound As Double) As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal sourcePathID As String
'           ID ранее объявленного пути, который читается без изменения.
'
'         ByVal startNX As Double
'           X-компонента начальной поперечной нормали; линия не должна быть параллельна плоскости.
'
'         ByVal startNY As Double
'           Y-компонента начальной поперечной нормали.
'
'         ByVal startBound As Double
'           Значение начальной плоскости startNX*x+startNY*y=startBound; мм при единичной нормали.
'
'         ByVal endNX As Double
'           X-компонента конечной поперечной нормали.
'
'         ByVal endNY As Double
'           Y-компонента конечной поперечной нормали.
'
'         ByVal endBound As Double
'           Значение конечной плоскости endNX*x+endNY*y=endBound; мм при единичной нормали.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Chain
'
'       Выбирает непрерывный участок между двумя именованными сегментными привязками. Направление обхода
'       явно; fullLoopIfSame задаёт полный оборот при совпавших концах.
'
'       Объявление:
'         Public Function Chain(ByVal pathID As String, ByVal sourcePathID As String, ByVal
'         startSegmentID As String, ByVal startParameter As Double, ByVal endSegmentID As String, ByVal
'         endParameter As Double, Optional ByVal forward As Boolean = True, Optional ByVal
'         fullLoopIfSame As Boolean = False) As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal sourcePathID As String
'           ID ранее объявленного пути, который читается без изменения.
'
'         ByVal startSegmentID As String
'           Сегмент начальной привязки цепочки.
'
'         ByVal startParameter As Double
'           Параметр начальной точки выбранного сегмента [0;1].
'
'         ByVal endSegmentID As String
'           Сегмент конечной привязки цепочки.
'
'         ByVal endParameter As Double
'           Параметр конечной точки выбранного сегмента [0;1].
'
'         Optional ByVal forward As Boolean = True
'           Явное направление следования исходному обходу; False — противоположное.
'
'         Optional ByVal fullLoopIfSame As Boolean = False
'           Полный оборот при совпавших привязках замкнутого пути; иначе нулевая цепочка недопустима.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Slice
'
'       Создаёт и регистрирует точный участок ранее объявленной линии между startDistance и endDistance
'       в мм от её начала. Отрезки и дуги сохраняются без приближения; ненулевое отсечение полного круга
'       даёт открытую дугу.
'
'       Объявление:
'         Public Function Slice(ByVal pathID As String, ByVal sourcePathID As String, ByVal
'         startDistance As Double, ByVal endDistance As Double) As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal sourcePathID As String
'           ID ранее объявленного пути, который читается без изменения.
'
'         ByVal startDistance As Double
'           Расстояние от начала всей линии, мм, в [0;Length]; endDistance строго больше startDistance.
'
'         ByVal endDistance As Double
'           Расстояние от начала всей линии, мм, в [0;Length]; endDistance строго больше startDistance.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Transform
'
'       Создаёт новый путь из существующего изометрией; сохраняет сегментные ID и происхождение,
'       отражение меняет знак sweep.
'
'       Объявление:
'         Public Function Transform(ByVal pathID As String, ByVal sourcePathID As String, ByVal shiftX
'         As Double, ByVal shiftY As Double, Optional ByVal rotation As Double = 0#, Optional ByVal
'         reflectX As Boolean = False) As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal sourcePathID As String
'           ID ранее объявленного пути, который читается без изменения.
'
'         ByVal shiftX As Double
'           Перенос по X, мм.
'
'         ByVal shiftY As Double
'           Перенос по Y, мм.
'
'         Optional ByVal rotation As Double = 0#
'           Угол поворота вокруг (0,0), радианы, CCW положителен.
'
'         Optional ByVal reflectX As Boolean = False
'           Отражение относительно локальной оси X перед поворотом и переносом.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.PathGroup
'
'       Объявляет массив ID независимых путей; порядок значим, фиктивные соединения отсутствуют.
'
'       Объявление:
'         Public Sub PathGroup(ByVal groupID As String, ByVal pathIDs As Variant)
'
'       Параметры:
'         ByVal groupID As String
'           Имя группы линий; оно задаётся отдельно от имени группы для подписи арматуры.
'
'         ByVal pathIDs As Variant
'           Массив ID существующих независимых путей; без повторов.
'
'
'
'
' ====================================================================================================
' 05  РАЗМЕЩЕНИЕ АРМАТУРЫ
' ====================================================================================================
'
'   Нулевой диаметр или нулевое число стержней отключают ряд. Пустой диаметр из таблицы читает
'   OptionalDiameter как 0; одиночный дефис недопустим. Для каждого независимого ряда используется свой
'   ID.
'
'
'   Класс / модуль: CSectionDefinitionBuilder
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarByCount
'
'       Задаёт места для указанного числа стержней на рабочей части линии размещения. EndMode определяет
'       включение концов. В группе это число применяется отдельно к каждой линии; нулевое число стержней
'       или нулевой диаметр отключают ряд.
'
'       Объявление:
'         Public Sub RebarByCount(ByVal ruleID As String, ByVal target As String, ByVal count As Long,
'         ByVal diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing, Optional
'         ByVal steelClass As String = "Rebar")
'
'       Параметры:
'         ByVal ruleID As String
'           Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила.
'
'         ByVal target As String
'           ID пути/PathGroup, разрешаемый в этом определении.
'
'         ByVal count As Long
'           Число стержней на одной линии размещения; целое неотрицательное. В группе применяется
'           отдельно к каждой линии. Ноль отключает ряд.
'
'         ByVal diameter As Double
'           Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд.
'
'         Optional ByVal options As CRebarRuleOptions = Nothing
'           CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт условия размещения по
'           умолчанию.
'
'         Optional ByVal steelClass As String = "Rebar"
'           Строка принадлежности материала арматуры; совпадения разных материалов требуют явной
'           политики.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarByMaxStep
'
'       Задаёт места стержней так, чтобы расстояние между соседними стержнями вдоль линии не превышало
'       step. Исполнитель подбирает число стержней; фактическое расстояние может быть меньше заданного.
'
'       Объявление:
'         Public Sub RebarByMaxStep(ByVal ruleID As String, ByVal target As String, ByVal step As
'         Double, ByVal diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing,
'         Optional ByVal steelClass As String = "Rebar")
'
'       Параметры:
'         ByVal ruleID As String
'           Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила.
'
'         ByVal target As String
'           ID пути/PathGroup, разрешаемый в этом определении.
'
'         ByVal step As Double
'           Расстояние между соседними стержнями вдоль линии размещения, мм; строго положительное.
'
'         ByVal diameter As Double
'           Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд.
'
'         Optional ByVal options As CRebarRuleOptions = Nothing
'           CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт условия размещения по
'           умолчанию.
'
'         Optional ByVal steelClass As String = "Rebar"
'           Строка принадлежности материала арматуры; совпадения разных материалов требуют явной
'           политики.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarByFixedStep
'
'       Задаёт точное расстояние step между соседними стержнями вдоль линии. RemainderPlacement
'       определяет, где оставить неиспользованный остаток длины. При IncludeBoth или ClosedUniform
'       рабочая длина должна соответствовать заданному расстоянию.
'
'       Объявление:
'         Public Sub RebarByFixedStep(ByVal ruleID As String, ByVal target As String, ByVal step As
'         Double, ByVal diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing,
'         Optional ByVal steelClass As String = "Rebar")
'
'       Параметры:
'         ByVal ruleID As String
'           Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила.
'
'         ByVal target As String
'           ID пути/PathGroup, разрешаемый в этом определении.
'
'         ByVal step As Double
'           Расстояние между соседними стержнями вдоль линии размещения, мм; строго положительное.
'
'         ByVal diameter As Double
'           Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд.
'
'         Optional ByVal options As CRebarRuleOptions = Nothing
'           CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт условия размещения по
'           умолчанию.
'
'         Optional ByVal steelClass As String = "Rebar"
'           Строка принадлежности материала арматуры; совпадения разных материалов требуют явной
'           политики.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarAtStations
'
'       Задаёт отдельные места стержней расстояниями от начала линии размещения, в мм. Порядок этих
'       расстояний в массиве определяет исходную нумерацию стержней.
'
'       Объявление:
'         Public Sub RebarAtStations(ByVal ruleID As String, ByVal target As String, ByVal stations As
'         Variant, ByVal diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing,
'         Optional ByVal steelClass As String = "Rebar")
'
'       Параметры:
'         ByVal ruleID As String
'           Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила.
'
'         ByVal target As String
'           ID пути/PathGroup, разрешаемый в этом определении.
'
'         ByVal stations As Variant
'           Массив расстояний от начала линии размещения до каждого стержня, мм; порядок элементов
'           задаёт исходную нумерацию.
'
'         ByVal diameter As Double
'           Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд.
'
'         Optional ByVal options As CRebarRuleOptions = Nothing
'           CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт условия размещения по
'           умолчанию.
'
'         Optional ByVal steelClass As String = "Rebar"
'           Строка принадлежности материала арматуры; совпадения разных материалов требуют явной
'           политики.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarDependent
'
'       Задаёт дополнительный ряд смещением относительно центров стержней исходного ряда. TangentOffset
'       задаёт смещение вдоль касательной, NormalOffset — поперёк линии. Места стержней связаны с
'       исходным рядом, новое равномерное распределение не выполняется. Если исходная линия выбрана из
'       замкнутого контура, нормальное смещение проверяется по всему смещённому контуру с учётом
'       соседних сторон. Позиции вне него пропускаются с предупреждением. Контактная добавка вдоль
'       касательной применяется после этой проверки.
'
'       Объявление:
'         Public Sub RebarDependent(ByVal ruleID As String, ByVal sourceRuleID As String, ByVal diameter
'         As Double, ByVal options As CRebarRuleOptions, Optional ByVal steelClass As String = "Rebar")
'
'       Параметры:
'         ByVal ruleID As String
'           Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила.
'
'         ByVal sourceRuleID As String
'           ID исходного правила зависимости; циклы запрещены.
'
'         ByVal diameter As Double
'           Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд.
'
'         ByVal options As CRebarRuleOptions
'           CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт условия размещения по
'           умолчанию.
'
'         Optional ByVal steelClass As String = "Rebar"
'           Строка принадлежности материала арматуры; совпадения разных материалов требуют явной
'           политики.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarProjected
'
'       Задаёт новый ряд проекцией центров стержней исходного ряда на целевую линию target. Направление
'       — нормаль исходной линии или общий вектор DirectionX/Y. StartSetback и EndSetback ограничивают
'       рабочий участок целевой линии; отсутствие пересечения отклоняется либо пропускается по
'       SkipMissingProjection.
'
'       Объявление:
'         Public Sub RebarProjected(ByVal ruleID As String, ByVal target As String, ByVal sourceRuleID
'         As String, ByVal diameter As Double, ByVal options As CRebarRuleOptions, Optional ByVal
'         steelClass As String = "Rebar")
'
'       Параметры:
'         ByVal ruleID As String
'           Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила.
'
'         ByVal target As String
'           ID пути/PathGroup, разрешаемый в этом определении.
'
'         ByVal sourceRuleID As String
'           ID исходного правила зависимости; циклы запрещены.
'
'         ByVal diameter As Double
'           Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд.
'
'         ByVal options As CRebarRuleOptions
'           CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт условия размещения по
'           умолчанию.
'
'         Optional ByVal steelClass As String = "Rebar"
'           Строка принадлежности материала арматуры; совпадения разных материалов требуют явной
'           политики.
'
'
'
'
' ====================================================================================================
' 06  УСЛОВИЯ РАЗМЕЩЕНИЯ СТЕРЖНЕЙ
' ====================================================================================================
'
'   Создайте New CRebarRuleOptions и измените только нужные свойства. Метод размещения копирует их при
'   объявлении ряда, поэтому позднейшая правка options не меняет уже объявленные ряды.
'
'
'   Класс / модуль: CRebarRuleOptions
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.StartSetback
'
'       Начальное отсечение вдоль каждого target-пути, мм>=0; у полного ClosedUniform 0.
'
'       Объявление:
'         Public StartSetback As Double
'       По умолчанию: 0.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.EndSetback
'
'       Конечное отсечение вдоль каждого target-пути, мм>=0; ProjectFrom фильтрует hit без
'       перераспределения.
'
'       Объявление:
'         Public EndSetback As Double
'       По умолчанию: 0.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.EndMode
'
'       IncludeBoth/ExcludeBoth/IncludeStart/IncludeEnd/ClosedUniform; последние требуют соответствующей
'       замкнутости.
'
'       Объявление:
'         Public EndMode As ERebarEndMode
'       По умолчанию: IncludeBoth.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.DistributionScope
'
'       Допустимо только PerPath: число стержней, расстояния между ними и отступы от концов применяются
'       отдельно к каждой независимой линии группы.
'
'       Объявление:
'         Public DistributionScope As ERebarDistributionScope
'       По умолчанию: PerPath.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.RemainderPlacement
'
'       RemainderStart/End/Center: положение остатка для точного шага.
'
'       Объявление:
'         Public RemainderPlacement As ERemainderPlacement
'       По умолчанию: RemainderCenter.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.TangentOffset
'
'       Локальное связанное смещение вдоль касательной, мм; signed, не независимый шаг ряда.
'
'       Объявление:
'         Public TangentOffset As Double
'       По умолчанию: 0.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.NormalOffset
'
'       Локальное связанное смещение по левой нормали (-ty,tx), мм; signed.
'
'       Объявление:
'         Public NormalOffset As Double
'       По умолчанию: 0.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.ValidateRowOffset
'
'       Включает ограниченные проверки семейства с ненулевыми базовыми отступами: весь диск в бетоне,
'       пересечения, самостоятельные совпадения до слияния и вырождение нормального переноса дуги.
'       CRebarContactBuilder задаёт признак автоматически; исполнитель распространяет его на оба
'       дополнительных ряда с общим Main.
'
'       Объявление:
'         Public ValidateRowOffset As Boolean
'       По умолчанию: False.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.EverySecond
'
'       Брать локальные исходные позиции 1, 3, ... отдельно на каждом пути; SuccessfulProjectionOrdinal
'       задаёт отдельный отсчёт успешных проекций.
'
'       Объявление:
'         Public EverySecond As Boolean
'       По умолчанию: False.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.IncomingTangentAtJoint
'
'       В точке соединения двух сегментов брать касательную предыдущего сегмента. По умолчанию берётся
'       касательная следующего сегмента.
'
'       Объявление:
'         Public IncomingTangentAtJoint As Boolean
'       По умолчанию: False.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.DirectionX
'
'       Глобальная X направления ProjectFrom, используется при SourceNormalSign=0.
'
'       Объявление:
'         Public DirectionX As Double
'       По умолчанию: 0.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.DirectionY
'
'       Глобальная Y направления ProjectFrom; направление не нулевое.
'
'       Объявление:
'         Public DirectionY As Double
'       По умолчанию: 0.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.SourceNormalSign
'
'       -1/0/+1: локальная нормаль исходного ряда или глобальное DirectionX/Y при 0.
'
'       Объявление:
'         Public SourceNormalSign As Double
'       По умолчанию: 0.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.SuccessfulProjectionOrdinal
'
'       True: заново нумеровать только стержни, для которых найдено пересечение с целевой линией, подряд
'       по всей группе. False: сохранять исходные номера стержней с возможными пропусками отдельно на
'       каждой линии.
'
'       Объявление:
'         Public SuccessfulProjectionOrdinal As Boolean
'       По умолчанию: False.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.NewPublicationRoot
'
'       У ProjectFrom начинает новую последовательность исходных номеров для успешно спроецированных
'       стержней.
'
'       Объявление:
'         Public NewPublicationRoot As Boolean
'       По умолчанию: False.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.SkipMissingProjection
'
'       Разрешает пропуск стержня без пересечения с целевой линией с диагностикой PROJECTION_MISSING;
'       False выдаёт ошибку.
'
'       Объявление:
'         Public SkipMissingProjection As Boolean
'       По умолчанию: False.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.Publication
'
'       PublishByRule или PublishPerSourcePosition: порядок записи стержней по правилам либо по исходным
'       позициям.
'
'       Объявление:
'         Public Publication As ERebarPublication
'       По умолчанию: PublishByRule.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.PositionOrder
'
'       Порядок вывода стержней: AlongPath — от начала линии, FromPathMiddle — от её середины. Места
'       запрошенных стержней и их исходные номера сохраняются. При Coincidence = KeepPublishedOwner
'       порядок также определяет, какой совпавший стержень сохраняется вместе с его диаметром и
'       материалом. Поэтому менять порядок в таком случае может изменить физическую арматуру.
'
'       Объявление:
'         Public PositionOrder As ERebarPositionOrder
'       По умолчанию: AlongPath.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.Coincidence
'
'       RejectCoincidence/MergeIdentical/KeepPublishedOwner; последнее сохраняет стержень, добавленный в
'       эту позицию первым.
'
'       Объявление:
'         Public Coincidence As ERebarCoincidence
'       По умолчанию: RejectCoincidence.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.PlacementCheck
'
'       CheckCenter или CheckFullDisk относительно Outer/Opening; ужесточение задаётся явно.
'
'       Объявление:
'         Public PlacementCheck As ERebarPlacementCheck
'       По умолчанию: CheckCenter.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.OverlapPolicy
'
'       ReportOverlap или RejectOverlap. Контакт допустим; нормативный минимальный зазор не вводится.
'
'       Объявление:
'         Public OverlapPolicy As ERebarOverlapPolicy
'       По умолчанию: ReportOverlap.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.SourcePrefix
'
'       Префикс SourcePositionId; пустой использует RuleId+'.' либо связь source+suffix.
'
'       Объявление:
'         Public SourcePrefix As String
'       По умолчанию: пустое значение VBA.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.SourceSuffix
'
'       Суффикс SourcePositionId, в том числе для связанных рядов.
'
'       Объявление:
'         Public SourceSuffix As String
'       По умолчанию: пустое значение VBA.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.AnnotationGroup
'
'       Имя группы, объединяющей стержни в общей подписи количества и диаметров на схеме.
'
'       Объявление:
'         Public AnnotationGroup As String
'       По умолчанию: пустое значение VBA.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.Comment
'
'       Инженерное пояснение правила; не selector или ResultCode.
'
'       Объявление:
'         Public Comment As String
'       По умолчанию: пустое значение VBA.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.SourceReference
'
'       Реальные адреса исходных параметров правила или программная ссылка.
'
'       Объявление:
'         Public SourceReference As String
'       По умолчанию: пустое значение VBA.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.ConfigureDependentContact
'
'       Задаёт только NormalOffset/TangentOffset/EverySecond по сумме половин основного и добавляемого
'       диаметров и промежуточных полных диаметров. Оба диаметра строго положительны: ноль вызывает
'       ошибку. Отключённый ряд обрабатывает вызывающий рецепт до вызова. Вторая компонента обнуляется;
'       остальные политики и source не наследуются.
'
'       Объявление:
'         Public Sub ConfigureDependentContact(ByVal baseDiameter As Double, ByVal diameter As Double,
'         ByVal axis As ERebarOffsetAxis, ByVal directionSign As Double, ByVal everySecond As Boolean,
'         Optional ByVal interveningDiameterSum As Double = 0#)
'
'       Параметры:
'         ByVal baseDiameter As Double
'           Строго положительный диаметр основного стержня, мм (>0). Ноль и отрицательное значение
'           вызывают ошибку; отключение ряда выполняет рецепт до вызова помощника.
'
'         ByVal diameter As Double
'           Строго положительный диаметр добавляемого стержня, мм (>0). Ноль и отрицательное значение
'           вызывают ошибку; отключение ряда выполняет рецепт до вызова помощника.
'
'         ByVal axis As ERebarOffsetAxis
'           RowNormal — смещение по левой нормали исходного ряда; RowTangent — вдоль его касательной.
'           Ориентацию задают фактический путь и directionSign.
'
'         ByVal directionSign As Double
'           Только -1 или +1: знак выбранной компоненты смещения; задаётся явно по направлению ряда.
'
'         ByVal everySecond As Boolean
'           True выбирает локальные исходные позиции 1,3,5…; отсчёт успешных проекций остаётся отдельной
'           настройкой.
'
'         Optional ByVal interveningDiameterSum As Double = 0#
'           Неотрицательная добавка к сумме радиусов, мм: промежуточные диаметры и предусмотренные
'           промежутки. CRebarContactBuilder включает сюда по 2,5 мм на каждую соседнюю пару.
'           EverySecond сам по себе не уменьшает расстояние.
'
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.Clone
'
'       Возвращает новую независимую копию всех настроек. Изменения исходного объекта и копии не влияют
'       друг на друга.
'
'       Объявление:
'         Public Function Clone() As CRebarRuleOptions
'
'
'   Класс / модуль: CRebarContactBuilder
'
'     ----------------------------------------------------------------------------------------------
'     CRebarContactBuilder.Initialize
'
'       Сохраняет положительный основной диаметр, два знака смещений и независимый снимок общих настроек
'       именно дополнительных рядов. Вызывается один раз на экземпляр; настройки основного ряда
'       автоматически не наследуются.
'
'       Объявление:
'         Public Sub Initialize(ByVal baseDiameter As Double, ByVal normalSign As Double, ByVal
'         tangentSign As Double, ByVal defaults As CRebarRuleOptions)
'
'       Параметры:
'         ByVal baseDiameter As Double
'           Положительный диаметр стержней основного ряда, мм (>0).
'
'         ByVal normalSign As Double
'           Знак смещения Stacked относительно левой нормали линии: +1 или -1.
'
'         ByVal tangentSign As Double
'           Знак смещения SideBySide вдоль линии: +1 по её направлению, -1 против.
'
'         ByVal defaults As CRebarRuleOptions
'           Отдельные общие CRebarRuleOptions дополнительных рядов. Объект обязателен; сохраняется
'           независимая копия. EndMode, отступы и имена основного ряда автоматически не наследуются.
'
'
'     ----------------------------------------------------------------------------------------------
'     CRebarContactBuilder.AppendRow
'
'       Создаёт независимые CRebarRuleOptions активного ряда. baseOffset задаёт отдельную базовую линию
'       на расстоянии от Main по нормали внутрь бетона. Совпавшие с геометрическим допуском отступы
'       образуют контактную группу: при нуле её первый стержень — Main, при положительном отступе —
'       первый включённый дополнительный ряд. Stacked добавляет контактное расстояние по нормали,
'       SideBySide — по касательной; диаметры и добавки 2,5 мм на выступы арматуры периодического
'       профиля накапливаются отдельно по этим направлениям внутри группы. Расстояние между соседними
'       центрами равно сумме радиусов плюс 2,5 мм; расчётные диаметры и площади не меняются.
'       EverySecondBar меняет только выбор исходных позиций. Ошибка не меняет состояние. Оба связанных
'       правила объявляются от Main.
'
'       Объявление:
'         Public Function AppendRow(ByVal diameter As Double, ByVal location As String, ByVal binding As
'         String, ByVal sourceReference As String, Optional ByVal baseOffset As Double = 0#) As
'         CRebarRuleOptions
'
'       Параметры:
'         ByVal diameter As Double
'           Положительный диаметр добавляемого ряда, мм (>0). Отключённый ряд пропускают до вызова, не
'           увеличивая накопленные диаметры.
'
'         ByVal location As String
'           Только Stacked (поперёк линии) или SideBySide (вдоль линии); неизвестное значение вызывает
'           адресную ошибку.
'
'         ByVal binding As String
'           Только EachBar (все исходные позиции) или EverySecondBar (позиции 1, 3 и следующие
'           нечётные). Накопление номинальных диаметров одинаково для обоих вариантов.
'
'         ByVal sourceReference As String
'           Адреса активных исходных полей этого ряда или программная ссылка. Сохраняются в результате и
'           в сообщении об ошибке.
'
'         Optional ByVal baseOffset As Double = 0#
'           Независимый неотрицательный базовый отступ от Main, мм; необязателен, по умолчанию 0.
'           Умножается на normalSign. Для разных отступов контактные поправки независимы; для одинаковых
'           учитываются только стержни своей группы и промежутки по 2,5 мм между ними.
'
'
'
'   Класс / модуль: modSectionRecipeTypes
'
'     ----------------------------------------------------------------------------------------------
'     modSectionRecipeTypes.ERebarOffsetAxis
'
'       RowNormal=1 или RowTangent=2: выбранная компонента контактного зависимого смещения; знак
'       задаётся отдельно.
'
'       Объявление:
'         Public Enum ERebarOffsetAxis
'
'       Допустимые значения / состав:
'         RowNormal = 1
'         RowTangent = 2
'
'     ----------------------------------------------------------------------------------------------
'     modSectionRecipeTypes.ERebarEndMode
'
'       IncludeBoth/ExcludeBoth/IncludeStart/IncludeEnd/ClosedUniform; последние требуют соответствующей
'       замкнутости.
'
'       Объявление:
'         Public Enum ERebarEndMode
'
'       Допустимые значения / состав:
'         IncludeBoth = 1
'         ExcludeBoth = 2
'         IncludeStart = 3
'         IncludeEnd = 4
'         ClosedUniform = 5
'
'     ----------------------------------------------------------------------------------------------
'     modSectionRecipeTypes.ERebarDistributionScope
'
'       Допустимо только PerPath: число стержней, расстояния между ними и отступы от концов применяются
'       отдельно к каждой независимой линии группы.
'
'       Объявление:
'         Public Enum ERebarDistributionScope
'
'       Допустимые значения / состав:
'         PerPath = 1
'
'     ----------------------------------------------------------------------------------------------
'     modSectionRecipeTypes.ERemainderPlacement
'
'       RemainderStart/End/Center: положение остатка для точного шага.
'
'       Объявление:
'         Public Enum ERemainderPlacement
'
'       Допустимые значения / состав:
'         RemainderStart = 1
'         RemainderEnd = 2
'         RemainderCenter = 3
'
'     ----------------------------------------------------------------------------------------------
'     modSectionRecipeTypes.ERebarPublication
'
'       PublishByRule или PublishPerSourcePosition: порядок записи стержней по правилам либо по исходным
'       позициям.
'
'       Объявление:
'         Public Enum ERebarPublication
'
'       Допустимые значения / состав:
'         PublishByRule = 1
'         PublishPerSourcePosition = 2
'
'     ----------------------------------------------------------------------------------------------
'     modSectionRecipeTypes.ERebarPositionOrder
'
'       Порядок вывода стержней: AlongPath — от начала линии, FromPathMiddle — от её середины. Места
'       запрошенных стержней и их исходные номера сохраняются. При Coincidence = KeepPublishedOwner
'       порядок также определяет, какой совпавший стержень сохраняется вместе с его диаметром и
'       материалом. Поэтому менять порядок в таком случае может изменить физическую арматуру.
'
'       Объявление:
'         Public Enum ERebarPositionOrder
'
'       Допустимые значения / состав:
'         AlongPath = 1
'         FromPathMiddle = 2
'
'     ----------------------------------------------------------------------------------------------
'     modSectionRecipeTypes.ERebarCoincidence
'
'       RejectCoincidence/MergeIdentical/KeepPublishedOwner; последнее сохраняет стержень, добавленный в
'       эту позицию первым.
'
'       Объявление:
'         Public Enum ERebarCoincidence
'
'       Допустимые значения / состав:
'         RejectCoincidence = 1
'         MergeIdentical = 2
'         KeepPublishedOwner = 3
'
'     ----------------------------------------------------------------------------------------------
'     modSectionRecipeTypes.ERebarPlacementCheck
'
'       CheckCenter или CheckFullDisk относительно Outer/Opening; ужесточение задаётся явно.
'
'       Объявление:
'         Public Enum ERebarPlacementCheck
'
'       Допустимые значения / состав:
'         CheckCenter = 1
'         CheckFullDisk = 2
'
'     ----------------------------------------------------------------------------------------------
'     modSectionRecipeTypes.ERebarOverlapPolicy
'
'       ReportOverlap или RejectOverlap. Контакт допустим; нормативный минимальный зазор не вводится.
'
'       Объявление:
'         Public Enum ERebarOverlapPolicy
'
'       Допустимые значения / состав:
'         ReportOverlap = 1
'         RejectOverlap = 2
'
'
'
' ====================================================================================================
' 07  РАЗМЕРЫ, ПОЯСНЕНИЯ И ПОДПИСИ АРМАТУРЫ
' ====================================================================================================
'
'   Для общей подписи стороны задайте одинаковую AnnotationGroup всем её рядам. Подпись считает
'   фактически созданные стержни всей группы и исчезает, если их нет. Смещение подписи в настройках
'   плоттера отсчитывается от границы бетона.
'
'
'   Класс / модуль: CSectionDefinitionBuilder
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Anchor
'
'       Объявляет именованную модельную точку. Данные копируются как два скаляра.
'
'       Объявление:
'         Public Sub Anchor(ByVal anchorID As String, ByVal x As Double, ByVal y As Double)
'
'       Параметры:
'         ByVal anchorID As String
'           Непустой уникальный ID модельной привязки.
'
'         ByVal x As Double
'           Модельная координата x, мм; знаки соответствуют section XY/WCS.
'
'         ByVal y As Double
'           Модельная координата y, мм; знаки соответствуют section XY/WCS.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.PathAnchor
'
'       Получает точку конкретного сегмента по ID и параметру [0;1], объявляет именованную привязку.
'
'       Объявление:
'         Public Sub PathAnchor(ByVal anchorID As String, ByVal pathID As String, ByVal segmentID As
'         String, Optional ByVal parameter As Double = 0#)
'
'       Параметры:
'         ByVal anchorID As String
'           Непустой уникальный ID модельной привязки.
'
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal segmentID As String
'           Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом.
'
'         Optional ByVal parameter As Double = 0#
'           Нормированный параметр точки сегмента [0;1], не расстояние по всему пути.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.TransformAnchor
'
'       Создаёт новую привязку из существующей: отражение по локальной X, поворот вокруг (0,0), перенос.
'
'       Объявление:
'         Public Sub TransformAnchor(ByVal anchorID As String, ByVal sourceAnchorID As String, ByVal
'         shiftX As Double, ByVal shiftY As Double, Optional ByVal rotation As Double = 0#, Optional
'         ByVal reflectX As Boolean = False)
'
'       Параметры:
'         ByVal anchorID As String
'           Непустой уникальный ID модельной привязки.
'
'         ByVal sourceAnchorID As String
'           ID существующей исходной привязки.
'
'         ByVal shiftX As Double
'           Перенос по X, мм.
'
'         ByVal shiftY As Double
'           Перенос по Y, мм.
'
'         Optional ByVal rotation As Double = 0#
'           Угол поворота вокруг (0,0), радианы, CCW положителен.
'
'         Optional ByVal reflectX As Boolean = False
'           Отражение относительно локальной оси X перед поворотом и переносом.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.VerticalExtent
'
'       Измеряет полный вертикальный габарит выбранного пути/группы с точными экстремумами дуг.
'       placementX и normalX задают сторону отображения.
'
'       Объявление:
'         Public Sub VerticalExtent(ByVal annotationID As String, ByVal target As String, ByVal name As
'         String, ByVal placementX As Double, Optional ByVal normalX As Double = -1#, Optional ByVal
'         comment As String = "")
'
'       Параметры:
'         ByVal annotationID As String
'           Непустой уникальный ID аннотации.
'
'         ByVal target As String
'           ID пути/PathGroup, разрешаемый в этом определении.
'
'         ByVal name As String
'           Обозначение измерения или имя контейнера.
'
'         ByVal placementX As Double
'           Модельная X размерной линии, мм.
'
'         Optional ByVal normalX As Double = -1#
'           X-компонента направления от бетона к подписи или выносной линии. Вместе с normalY задаёт
'           любую, в том числе диагональную, нормаль; длина вектора нормируется. У отверстия направление
'           обращено в пустоту.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.HorizontalExtent
'
'       Измеряет полный горизонтальный габарит выбранного пути/группы. Это не длина пути и не исходная
'       ширина до скругления.
'
'       Объявление:
'         Public Sub HorizontalExtent(ByVal annotationID As String, ByVal target As String, ByVal name
'         As String, ByVal placementY As Double, Optional ByVal normalY As Double = 1#, Optional ByVal
'         comment As String = "")
'
'       Параметры:
'         ByVal annotationID As String
'           Непустой уникальный ID аннотации.
'
'         ByVal target As String
'           ID пути/PathGroup, разрешаемый в этом определении.
'
'         ByVal name As String
'           Обозначение измерения или имя контейнера.
'
'         ByVal placementY As Double
'           Модельная Y размерной линии, мм.
'
'         Optional ByVal normalY As Double = 1#
'           Y-компонента той же нормали. Например, (1,1) задаёт диагональ вправо-вверх, (0,1) — вверх.
'           Положение текста отсчитывается от соответствующей бетонной границы, включая границу
'           отверстия.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.HorizontalDistance
'
'       Объявляет размер между двумя заданными X на уровне placementY. DisplayIncrement округляет только
'       подпись, не геометрию.
'
'       Объявление:
'         Public Sub HorizontalDistance(ByVal annotationID As String, ByVal name As String, ByVal startX
'         As Double, ByVal endX As Double, ByVal placementY As Double, Optional ByVal normalY As Double
'         = 1#, Optional ByVal comment As String = "", Optional ByVal displayIncrement As Double = 0#)
'
'       Параметры:
'         ByVal annotationID As String
'           Непустой уникальный ID аннотации.
'
'         ByVal name As String
'           Обозначение измерения или имя контейнера.
'
'         ByVal startX As Double
'           Модельная координата startX, мм; знаки соответствуют section XY/WCS.
'
'         ByVal endX As Double
'           Модельная координата endX, мм; знаки соответствуют section XY/WCS.
'
'         ByVal placementY As Double
'           Модельная Y размерной линии, мм.
'
'         Optional ByVal normalY As Double = 1#
'           Y-компонента той же нормали. Например, (1,1) задаёт диагональ вправо-вверх, (0,1) — вверх.
'           Положение текста отсчитывается от соответствующей бетонной границы, включая границу
'           отверстия.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'         Optional ByVal displayIncrement As Double = 0#
'           Неотрицательный шаг округления только видимого справочного значения, мм; 0 означает
'           отсутствие такого округления.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.VerticalDistance
'
'       Объявляет вертикальный размер между двумя Y на линии placementX; значение — модуль разности.
'
'       Объявление:
'         Public Sub VerticalDistance(ByVal annotationID As String, ByVal name As String, ByVal startY
'         As Double, ByVal endY As Double, ByVal placementX As Double, Optional ByVal normalX As Double
'         = -1#, Optional ByVal comment As String = "")
'
'       Параметры:
'         ByVal annotationID As String
'           Непустой уникальный ID аннотации.
'
'         ByVal name As String
'           Обозначение измерения или имя контейнера.
'
'         ByVal startY As Double
'           Модельная координата startY, мм; знаки соответствуют section XY/WCS.
'
'         ByVal endY As Double
'           Модельная координата endY, мм; знаки соответствуют section XY/WCS.
'
'         ByVal placementX As Double
'           Модельная X размерной линии, мм.
'
'         Optional ByVal normalX As Double = -1#
'           X-компонента направления от бетона к подписи или выносной линии. Вместе с normalY задаёт
'           любую, в том числе диагональную, нормаль; длина вектора нормируется. У отверстия направление
'           обращено в пустоту.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Dimension
'
'       Измеряет между двумя именованными точками: Horizontal/Vertical — проекция, Aligned — истинное
'       расстояние по соединяющей линии под любым углом. Текст параллелен размерной линии и
'       автоматически разворачивается для удобного чтения. Проекция размещается через первую точку.
'
'       Объявление:
'         Public Sub Dimension(ByVal annotationID As String, ByVal startAnchorID As String, ByVal
'         endAnchorID As String, ByVal name As String, Optional ByVal measurement As String = "Aligned",
'         Optional ByVal normalX As Double = 0#, Optional ByVal normalY As Double = 0#, Optional ByVal
'         comment As String = "")
'
'       Параметры:
'         ByVal annotationID As String
'           Непустой уникальный ID аннотации.
'
'         ByVal startAnchorID As String
'           ID начальной модельной точки измерения.
'
'         ByVal endAnchorID As String
'           ID конечной модельной точки измерения.
'
'         ByVal name As String
'           Обозначение измерения или имя контейнера.
'
'         Optional ByVal measurement As String = "Aligned"
'           Ровно Horizontal, Vertical или Aligned; регистр букв должен совпадать с указанным.
'
'         Optional ByVal normalX As Double = 0#
'           X-компонента направления от бетона к подписи или выносной линии. Вместе с normalY задаёт
'           любую, в том числе диагональную, нормаль; длина вектора нормируется. У отверстия направление
'           обращено в пустоту.
'
'         Optional ByVal normalY As Double = 0#
'           Y-компонента той же нормали. Например, (1,1) задаёт диагональ вправо-вверх, (0,1) — вверх.
'           Положение текста отсчитывается от соответствующей бетонной границы, включая границу
'           отверстия.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.PathLength
'
'       Объявляет точную суммарную длину target-путей и текстовую привязку; не создаёт фиктивную
'       соединяющую линию группы.
'
'       Объявление:
'         Public Sub PathLength(ByVal annotationID As String, ByVal target As String, ByVal
'         placementAnchorID As String, Optional ByVal name As String = "L", Optional ByVal comment As
'         String = "")
'
'       Параметры:
'         ByVal annotationID As String
'           Непустой уникальный ID аннотации.
'
'         ByVal target As String
'           ID пути/PathGroup, разрешаемый в этом определении.
'
'         ByVal placementAnchorID As String
'           ID точки размещения текстового измерения.
'
'         Optional ByVal name As String = "L"
'           Обозначение измерения или имя контейнера.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Radius
'
'       Измеряет радиус конкретной Arc/Circle; параметр выбирает конечную точку радиуса. Прямая —
'       ошибка.
'
'       Объявление:
'         Public Sub Radius(ByVal annotationID As String, ByVal pathID As String, ByVal segmentID As
'         String, Optional ByVal parameter As Double = 0.5, Optional ByVal name As String = "R",
'         Optional ByVal comment As String = "")
'
'       Параметры:
'         ByVal annotationID As String
'           Непустой уникальный ID аннотации.
'
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal segmentID As String
'           Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом.
'
'         Optional ByVal parameter As Double = 0.5
'           Нормированный параметр точки сегмента [0;1], не расстояние по всему пути.
'
'         Optional ByVal name As String = "R"
'           Обозначение измерения или имя контейнера.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Diameter
'
'       Измеряет диаметр конкретной Arc/Circle с противоположным концом на той же окружности. Прямая —
'       ошибка.
'
'       Объявление:
'         Public Sub Diameter(ByVal annotationID As String, ByVal pathID As String, ByVal segmentID As
'         String, Optional ByVal parameter As Double = 0.5, Optional ByVal name As String = "D",
'         Optional ByVal comment As String = "")
'
'       Параметры:
'         ByVal annotationID As String
'           Непустой уникальный ID аннотации.
'
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal segmentID As String
'           Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом.
'
'         Optional ByVal parameter As Double = 0.5
'           Нормированный параметр точки сегмента [0;1], не расстояние по всему пути.
'
'         Optional ByVal name As String = "D"
'           Обозначение измерения или имя контейнера.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Text
'
'       Объявляет простой текст на именованной модельной привязке; пиксели/шрифт задаёт общий layout.
'
'       Объявление:
'         Public Sub Text(ByVal annotationID As String, ByVal anchorID As String, ByVal textValue As
'         String, Optional ByVal comment As String = "")
'
'       Параметры:
'         ByVal annotationID As String
'           Непустой уникальный ID аннотации.
'
'         ByVal anchorID As String
'           Непустой уникальный ID модельной привязки.
'
'         ByVal textValue As String
'           Текст подписи; Unicode сохраняется.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarLabel
'
'       Привязывает подпись группы к выбранному пути со стороны normalX/normalY, включая диагональные
'       направления. Опорная линия определяется по точным проекциям отрезков и дуг. Плоттер отсчитывает
'       отступ от ближайшей бетонной границы по нормали; у отверстия сохраняется привязка к его границе.
'       Текст параллелен линии подписи и разворачивается для удобного чтения.
'
'       Объявление:
'         Public Sub RebarLabel(ByVal groupID As String, ByVal target As String, ByVal normalX As
'         Double, ByVal normalY As Double)
'
'       Параметры:
'         ByVal groupID As String
'           Имя группы линий; оно задаётся отдельно от имени группы для подписи арматуры.
'
'         ByVal target As String
'           ID пути/PathGroup, разрешаемый в этом определении.
'
'         ByVal normalX As Double
'           X-компонента направления от бетона к подписи или выносной линии. Вместе с normalY задаёт
'           любую, в том числе диагональную, нормаль; длина вектора нормируется. У отверстия направление
'           обращено в пустоту.
'
'         ByVal normalY As Double
'           Y-компонента той же нормали. Например, (1,1) задаёт диагональ вправо-вверх, (0,1) — вверх.
'           Положение текста отсчитывается от соответствующей бетонной границы, включая границу
'           отверстия.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarRowLabel
'
'       Привязывает фактическую подпись группы к концам исполненного ruleID.
'
'       Объявление:
'         Public Sub RebarRowLabel(ByVal groupID As String, ByVal ruleID As String, ByVal normalX As
'         Double, ByVal normalY As Double)
'
'       Параметры:
'         ByVal groupID As String
'           Имя группы линий; оно задаётся отдельно от имени группы для подписи арматуры.
'
'         ByVal ruleID As String
'           Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила.
'
'         ByVal normalX As Double
'           X-компонента направления от бетона к подписи или выносной линии. Вместе с normalY задаёт
'           любую, в том числе диагональную, нормаль; длина вектора нормируется. У отверстия направление
'           обращено в пустоту.
'
'         ByVal normalY As Double
'           Y-компонента той же нормали. Например, (1,1) задаёт диагональ вправо-вверх, (0,1) — вверх.
'           Положение текста отсчитывается от соответствующей бетонной границы, включая границу
'           отверстия.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarRowExtentLabel
'
'       Привязывает подпись к фактическому габариту позиций правила; reverseDirection меняет ориентацию
'       anchor-линии.
'
'       Объявление:
'         Public Sub RebarRowExtentLabel(ByVal groupID As String, ByVal ruleID As String, ByVal normalX
'         As Double, ByVal normalY As Double, Optional ByVal reverseDirection As Boolean = False)
'
'       Параметры:
'         ByVal groupID As String
'           Имя группы линий; оно задаётся отдельно от имени группы для подписи арматуры.
'
'         ByVal ruleID As String
'           Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила.
'
'         ByVal normalX As Double
'           X-компонента направления от бетона к подписи или выносной линии. Вместе с normalY задаёт
'           любую, в том числе диагональную, нормаль; длина вектора нормируется. У отверстия направление
'           обращено в пустоту.
'
'         ByVal normalY As Double
'           Y-компонента той же нормали. Например, (1,1) задаёт диагональ вправо-вверх, (0,1) — вверх.
'           Положение текста отсчитывается от соответствующей бетонной границы, включая границу
'           отверстия.
'
'         Optional ByVal reverseDirection As Boolean = False
'           Обращает направление линии привязки подписи; физический ряд не меняется.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarRowExtremeLabel
'
'       Опорную прямую выбирает по крайнему положению стержней в directionX/Y; длина и середина вдоль
'       неё учитывают все позиции указанного правила, включая скругления. Сторона текста normalX/Y
'       независима от выбора опорной прямой.
'
'       Объявление:
'         Public Sub RebarRowExtremeLabel(ByVal groupID As String, ByVal ruleID As String, ByVal
'         directionX As Double, ByVal directionY As Double, ByVal normalX As Double, ByVal normalY As
'         Double)
'
'       Параметры:
'         ByVal groupID As String
'           Имя группы линий; оно задаётся отдельно от имени группы для подписи арматуры.
'
'         ByVal ruleID As String
'           Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила.
'
'         ByVal directionX As Double
'           X-компонента направления проекции/выбора экстремума; вектор не нулевой, у
'           RebarRowExtremeLabel единичный.
'
'         ByVal directionY As Double
'           Y-компонента направления проекции/выбора экстремума.
'
'         ByVal normalX As Double
'           X-компонента направления от бетона к подписи или выносной линии. Вместе с normalY задаёт
'           любую, в том числе диагональную, нормаль; длина вектора нормируется. У отверстия направление
'           обращено в пустоту.
'
'         ByVal normalY As Double
'           Y-компонента той же нормали. Например, (1,1) задаёт диагональ вправо-вверх, (0,1) — вверх.
'           Положение текста отсчитывается от соответствующей бетонной границы, включая границу
'           отверстия.
'
'
'
'
' ====================================================================================================
' 08  СОСТАВНЫЕ ЛИНИИ И ОТДЕЛЬНЫЕ УЧАСТКИ
' ====================================================================================================
'
'   Обычно достаточно готовых методов разделов 03–04. Для собственной последовательности отрезков и дуг:
'   Initialize, AddLine/AddArc, Seal, затем facade.AddPath. После Seal менять состав линии нельзя.
'
'
'   Класс / модуль: CGeometryPath
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.Initialize
'
'       Создаёт mutable путь с непустым ID один раз.
'
'       Объявление:
'         Public Sub Initialize(ByVal pathID As String)
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.AddLine
'
'       Добавляет прямой сегмент с устойчивым ID; начало совпадает с предыдущим концом.
'
'       Объявление:
'         Public Sub AddLine(ByVal segmentID As String, ByVal x1 As Double, ByVal y1 As Double, ByVal x2
'         As Double, ByVal y2 As Double)
'
'       Параметры:
'         ByVal segmentID As String
'           Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом.
'
'         ByVal x1 As Double
'           Модельная координата x1, мм; знаки соответствуют section XY/WCS.
'
'         ByVal y1 As Double
'           Модельная координата y1, мм; знаки соответствуют section XY/WCS.
'
'         ByVal x2 As Double
'           Модельная координата x2, мм; знаки соответствуют section XY/WCS.
'
'         ByVal y2 As Double
'           Модельная координата y2, мм; знаки соответствуют section XY/WCS.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.AddArc
'
'       Добавляет аналитическую дугу по центру/R/startAngle/sweep, без хордовой аппроксимации.
'
'       Объявление:
'         Public Sub AddArc(ByVal segmentID As String, ByVal centerX As Double, ByVal centerY As Double,
'         ByVal radius As Double, ByVal startAngle As Double, ByVal sweep As Double)
'
'       Параметры:
'         ByVal segmentID As String
'           Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом.
'
'         ByVal centerX As Double
'           Модельная координата centerX, мм; знаки соответствуют section XY/WCS.
'
'         ByVal centerY As Double
'           Модельная координата centerY, мм; знаки соответствуют section XY/WCS.
'
'         ByVal radius As Double
'           Положительный радиус, мм.
'
'         ByVal startAngle As Double
'           Начальный угол, радианы; от +X.
'
'         ByVal sweep As Double
'           Signed sweep, радианы: CCW положителен, CW отрицателен, 0 запрещён, |sweep|<=2*pi.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.Seal
'
'       Публикует путь. closed=True проверяет совпадение концов, не достраивает соединение.
'
'       Объявление:
'         Public Sub Seal(Optional ByVal closed As Boolean = False)
'
'       Параметры:
'         Optional ByVal closed As Boolean = False
'           Явный Boolean режима замыкания/касательной; не исправляет геометрию автоматически.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.ID
'
'       Уникальное имя этой линии.
'
'       Объявление:
'         Public Property Get ID() As String
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.SegmentCount
'
'       Число аналитических line/arc сегментов.
'
'       Объявление:
'         Public Property Get SegmentCount() As Long
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.IsClosed
'
'       Замкнут ли фактический путь, без фиктивного соединения.
'
'       Объявление:
'         Public Property Get IsClosed() As Boolean
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.Length
'
'       Точная суммарная длина пути, мм.
'
'       Объявление:
'         Public Property Get Length() As Double
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.SegmentID
'
'       Возвращает имя сегмента по его порядковому номеру.
'
'       Объявление:
'         Public Property Get SegmentID(ByVal index As Long) As String
'
'       Параметры:
'         ByVal index As Long
'           1-based индекс в пределах соответствующего готового контейнера.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.SegmentIndex
'
'       Находит текущий номер сегмента по имени; если дуга исчезла при построении, выдаёт
'       MISSING_SEGMENT.
'
'       Объявление:
'         Public Function SegmentIndex(ByVal segmentID As String) As Long
'
'       Параметры:
'         ByVal segmentID As String
'           Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.AnchorStation
'
'       Находит расстояние от начала всей линии до точки заданного сегмента, мм. Точка задаётся именем
'       сегмента и параметром [0;1]. Если при построении участок исчез, обе его границы сохраняют
'       привязку к оставшемуся стыку. Неизвестное имя вызывает ошибку.
'
'       Объявление:
'         Public Function AnchorStation(ByVal segmentID As String, Optional ByVal parameter As Double =
'         0#) As Double
'
'       Параметры:
'         ByVal segmentID As String
'           Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом.
'
'         Optional ByVal parameter As Double = 0#
'           Нормированный параметр точки сегмента [0;1], не расстояние по всему пути.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.ResolveAnchor
'
'       Находит существующий сегмент по координатам точки на завершённой линии. Возвращает SegmentID и
'       локальный параметр [0;1]. Использует общие допуски принадлежности прямым и дугам. Точка вне
'       линии вызывает ANCHOR_NOT_ON_PATH; стык или совпадающие сегменты — ANCHOR_AMBIGUOUS с именами
'       кандидатов. Выходные аргументы при ошибке не меняются. На замкнутой окружности из одного
'       сегмента начало и конец имеют одну геометрическую точку; возвращается параметр 0.
'
'       Объявление:
'         Public Sub ResolveAnchor(ByVal x As Double, ByVal y As Double, ByRef segmentID As String,
'         ByRef parameter As Double)
'
'       Параметры:
'         ByVal x As Double
'           Модельная координата x, мм; знаки соответствуют section XY/WCS.
'
'         ByVal y As Double
'           Модельная координата y, мм; знаки соответствуют section XY/WCS.
'
'         ByRef segmentID As String
'           Возвращаемое существующее имя сегмента этой линии; не новое имя и не обязательно имя
'           исходного контура. ByRef: возвращаемые скаляры/буфер можно изменять, исходная геометрия не
'           меняется.
'
'         ByRef parameter As Double
'           Возвращаемая доля длины найденного сегмента [0;1]: 0 — начало, 1 — конец по направлению его
'           обхода. ByRef: возвращаемые скаляры/буфер можно изменять, исходная геометрия не меняется.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.BetweenAnchors
'
'       Возвращает точную цепочку между привязками с явным направлением и правилом полного обхода.
'
'       Объявление:
'         Public Function BetweenAnchors(ByVal pathID As String, ByVal startSegmentID As String, ByVal
'         startParameter As Double, ByVal endSegmentID As String, ByVal endParameter As Double, Optional
'         ByVal forward As Boolean = True, Optional ByVal fullLoopIfSame As Boolean = False) As
'         CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal startSegmentID As String
'           Сегмент начальной привязки цепочки.
'
'         ByVal startParameter As Double
'           Параметр начальной точки выбранного сегмента [0;1].
'
'         ByVal endSegmentID As String
'           Сегмент конечной привязки цепочки.
'
'         ByVal endParameter As Double
'           Параметр конечной точки выбранного сегмента [0;1].
'
'         Optional ByVal forward As Boolean = True
'           Явное направление следования исходному обходу; False — противоположное.
'
'         Optional ByVal fullLoopIfSame As Boolean = False
'           Полный оборот при совпавших привязках замкнутого пути; иначе нулевая цепочка недопустима.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.AtDistance
'
'       Возвращает координаты точки на заданном расстоянии от начала линии и единичный вектор
'       касательной в этой точке. На стыке можно выбрать касательную предыдущего сегмента.
'
'       Объявление:
'         Public Sub AtDistance(ByVal distance As Double, ByRef x As Double, ByRef y As Double, ByRef
'         tangentX As Double, ByRef tangentY As Double, Optional ByVal incomingTangentAtJoint As Boolean
'         = False)
'
'       Параметры:
'         ByVal distance As Double
'           Расстояние или граница интервала distance, внутренние мм; направление и допустимые пределы
'           определяет вызываемый метод.
'
'         ByRef x As Double
'           Модельная координата x, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые
'           скаляры/буфер можно изменять, исходная геометрия не меняется.
'
'         ByRef y As Double
'           Модельная координата y, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые
'           скаляры/буфер можно изменять, исходная геометрия не меняется.
'
'         ByRef tangentX As Double
'           Компонента единичной касательной, без размерности. ByRef: возвращаемые скаляры/буфер можно
'           изменять, исходная геометрия не меняется.
'
'         ByRef tangentY As Double
'           Компонента единичной касательной, без размерности. ByRef: возвращаемые скаляры/буфер можно
'           изменять, исходная геометрия не меняется.
'
'         Optional ByVal incomingTangentAtJoint As Boolean = False
'           Явный Boolean режима замыкания/касательной; не исправляет геометрию автоматически.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.Slice
'
'       Обрезает полный непрерывный путь по длине; точные дуги и source parameters сохраняются.
'
'       Объявление:
'         Public Function Slice(ByVal startDistance As Double, ByVal endDistance As Double) As
'         CGeometryPath
'
'       Параметры:
'         ByVal startDistance As Double
'           Расстояние от начала всей линии, мм, в [0;Length]; endDistance строго больше startDistance.
'
'         ByVal endDistance As Double
'           Расстояние от начала всей линии, мм, в [0;Length]; endDistance строго больше startDistance.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.Reversed
'
'       Копирует путь с обратным порядком сегментов/обходом; signed sweep и source parameters
'       обращаются.
'
'       Объявление:
'         Public Function Reversed() As CGeometryPath
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.Transformed
'
'       Копирует путь изометрией: отражение по локальной X, поворот, перенос; длина сохраняется.
'
'       Объявление:
'         Public Function Transformed(ByVal shiftX As Double, ByVal shiftY As Double, Optional ByVal
'         rotation As Double = 0#, Optional ByVal reflectX As Boolean = False) As CGeometryPath
'
'       Параметры:
'         ByVal shiftX As Double
'           Перенос по X, мм.
'
'         ByVal shiftY As Double
'           Перенос по Y, мм.
'
'         Optional ByVal rotation As Double = 0#
'           Угол поворота вокруг (0,0), радианы, CCW положителен.
'
'         Optional ByVal reflectX As Boolean = False
'           Отражение относительно локальной оси X перед поворотом и переносом.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.WithID
'
'       Копирует завершённую линию с новым именем, сохраняя соответствие сегментов исходным линиям.
'
'       Объявление:
'         Public Function WithID(ByVal pathID As String) As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'
'
'   Класс / модуль: modGeometryTypes
'
'     ----------------------------------------------------------------------------------------------
'     modGeometryTypes.GEOM_PI
'
'       Число pi как сумма сохраняемых VBE литералов, без изменения Double при компиляции.
'
'       Объявление:
'         Public Const GEOM_PI As Double = 3.14159265358979 + 3.10862446895044E-15
'
'     ----------------------------------------------------------------------------------------------
'     modGeometryTypes.TPoint2D
'
'       Два координатных скаляра X/Y, мм, без Excel или объекта на каждую точку.
'
'       Объявление:
'         Public Type TPoint2D
'
'       Допустимые значения / состав:
'         X As Double
'         Y As Double
'
'
'
' ====================================================================================================
' 09  АДРЕСА ВВОДА И СООБЩЕНИЯ ОБ ОШИБКАХ
' ====================================================================================================
'
'   Source связывает имя геометрического объекта с адресом исходных данных. Ошибка должна указывать,
'   какой размер или ряд нужно исправить. LastGeometryDiagnostic может вернуть Nothing, если сохранённой
'   ошибки нет.
'
'
'   Класс / модуль: CSectionDefinitionBuilder
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Source
'
'       Сохраняет для имени объекта текст его источника или адреса; ссылку на объект Excel не хранит.
'
'       Объявление:
'         Public Sub Source(ByVal entityID As String, ByVal sourceReference As String)
'
'       Параметры:
'         ByVal entityID As String
'           Имя объекта, которому назначается ссылка на исходные данные.
'
'         ByVal sourceReference As String
'           Строка исходного адреса/программного параметра; без объекта Excel.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.LastGeometryDiagnostic
'
'       Возвращает последнюю геометрическую ошибку с адресом исходных данных. Если сохранённой ошибки
'       нет, возвращает Nothing.
'
'       Объявление:
'         Public Property Get LastGeometryDiagnostic() As CGeometryDiagnostic
'
'
'   Класс / модуль: CGeometryPath
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.LastDiagnostic
'
'       Последняя машинная диагностика; Nothing при отсутствии записанного события.
'
'       Объявление:
'         Public Property Get LastDiagnostic() As CGeometryDiagnostic
'
'
'   Класс / модуль: CGeometryDiagnostic
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryDiagnostic.Code
'
'       Машинный код причины; назначается напрямую, не извлекается из русского текста.
'
'       Объявление:
'         Public Property Get Code() As String
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryDiagnostic.EntityID
'
'       Имя объекта, вызвавшего ошибку, если его удалось определить.
'
'       Объявление:
'         Public Property Get EntityID() As String
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryDiagnostic.RelatedEntityID
'
'       ID связанного сегмента/позиции/стержня для локализации причины.
'
'       Объявление:
'         Public Property Get RelatedEntityID() As String
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryDiagnostic.Reason
'
'       Человекочитаемая причина; не selector/status.
'
'       Объявление:
'         Public Property Get Reason() As String
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryDiagnostic.Detail1
'
'       Первое численное поле конкретного Code; единицы определяет Code.
'
'       Объявление:
'         Public Property Get Detail1() As Double
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryDiagnostic.Detail2
'
'       Второе численное поле конкретного Code.
'
'       Объявление:
'         Public Property Get Detail2() As Double
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryDiagnostic.SourceReference
'
'       Адрес исходной ячейки либо пояснение, откуда взят параметр геометрии или арматуры.
'
'       Объявление:
'         Public Property Get SourceReference() As String
'
