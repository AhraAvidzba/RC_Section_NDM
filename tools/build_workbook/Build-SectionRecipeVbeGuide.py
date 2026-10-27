"""Справочник автора геометрии, арматуры и аннотаций: только комментарии VBE."""
import importlib.util
import json
import re
import textwrap
from pathlib import Path
from SectionRecipeApiText import purpose, parameter_meaning
from SectionRecipeAuthorScope import selected_members
from VbeCommentText import readable
from RecipeArgumentText import argument_comment, signatures

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('api', Path(__file__).with_name('Build-SectionRecipeApi.py'))
api = importlib.util.module_from_spec(spec)
spec.loader.exec_module(api)


def author_text(text):
    # Описание для автора сечения; технические имена API сохраняются.
    replacements = {
        'Смысловое имя зарегистрированного типа; непустой текст.': 'Уникальное имя зарегистрированного типа; непустой текст.',
        'Возвращает снимок зарегистрированного блока по имени; отсутствие — ошибка, не fallback.': 'Возвращает копию зарегистрированной таблицы по имени; если таблицы нет, выдаёт ошибку.',
        'Читает непустой текст либо явно переданный fallback, обрезая крайние пробелы.': 'Читает непустой текст либо переданное значение по умолчанию, удаляя крайние пробелы.',
        'Проверяет вариант по allowed без учёта регистра и возвращает каноническое значение allowed.': 'Проверяет вариант по allowed без учёта регистра и возвращает написание из allowed.',
        'Массив разрешённых строк; Choice возвращает каноническую строку из этого массива.': 'Массив разрешённых строк; Choice возвращает строку именно в написании из этого массива.',
        'CUnitSystem, уже подготовленная вызывающим workflow; преобразование только через её методы.': 'CUnitSystem с уже выбранными единицами; преобразование выполняется её методами.',
        'Смысловой заголовок столбца; повторные заголовки — ошибка.': 'Название столбца; повторные заголовки — ошибка.',
        'Находит уникальную строку смыслового ключа в указанном подблоке.': 'Находит единственную строку с заданным ключом в указанной части таблицы.',
        'Смысловой ключ строки, не Excel-адрес.': 'Название параметра, по которому ищется строка.',
        'Столбец смысловых ключей 1..ColumnCount.': 'Номер столбца с названиями параметров: от 1 до ColumnCount.',
        'Существующий непрерывный опубликованный CGeometryPath.': 'Уже добавленная непрерывная линия CGeometryPath.',
        'Смысловой ID группы; путь и арматурная annotation-group — разные контракты.': 'Имя группы линий; оно задаётся отдельно от имени группы для подписи арматуры.',
        'У ProjectFrom начинает новый корень публикации по успешным позициям.': 'У ProjectFrom начинает новую последовательность исходных номеров для успешно спроецированных стержней.',
        'Разрешить пропуск без hit с PROJECTION_MISSING; False выдаёт ошибку.': 'Разрешает пропуск стержня без пересечения с целевой линией с диагностикой PROJECTION_MISSING; False выдаёт ошибку.',
        'PublishByRule или PublishPerSourcePosition; влияет на вывод, не геометрический шаг.': 'PublishByRule или PublishPerSourcePosition: порядок записи стержней по правилам либо по исходным позициям.',
        'порядок также определяет владельца совпавшей позиции: сохраняются его диаметр и материал.': 'порядок также определяет, какой совпавший стержень сохраняется вместе с его диаметром и материалом.',
        'RejectCoincidence/MergeIdentical/KeepPublishedOwner; последнее сохраняет явно выбранного опубликованного владельца.': 'RejectCoincidence/MergeIdentical/KeepPublishedOwner; последнее сохраняет стержень, добавленный в эту позицию первым.',
        'Смысловая принадлежность подписи физического стержня; не расчётная группа трещин СП35.': 'Имя группы, объединяющей стержни в общей подписи количества и диаметров на схеме.',
        'Ровно Horizontal, Vertical или Aligned; значение чувствительно к указанному контракту.': 'Ровно Horizontal, Vertical или Aligned; регистр букв должен совпадать с указанным.',
        'Смысловой ID этого пути.': 'Уникальное имя этой линии.',
        'Возвращает смысловой ID сегмента по индексу.': 'Возвращает имя сегмента по его порядковому номеру.',
        'Разрешает существующий смысловой ID в текущий индекс; исчезнувшая дуга даёт MISSING_SEGMENT.': 'Находит текущий номер сегмента по имени; если дуга исчезла при построении, выдаёт MISSING_SEGMENT.',
        'Копирует опубликованный путь с новым ID, сохраняя сегментные ссылки и исходное происхождение.': 'Копирует завершённую линию с новым именем, сохраняя соответствие сегментов исходным линиям.',
        'Связывает смысловой ID со строкой источника/адреса без сохранения объекта Excel.': 'Сохраняет для имени объекта текст его источника или адреса; ссылку на объект Excel не хранит.',
        'Смысловой ID сущности отказа, если он известен этому сервису.': 'Имя объекта, вызвавшего ошибку, если его удалось определить.',
        'Интерфейс единственного класса типа. Возвращает опубликованное definition через общий фасад, без сетки/баров/Shapes в рецепте.':
            'Обязательный метод класса нового типа. Возвращает завершённое описание сечения. Сетку, стержни и рисунок программа строит по этому описанию автоматически.',
        'Подготовленный снимок CSectionRecipeContext, не live Excel.':
            'Объект с копиями входных таблиц и системой единиц; ячейки Excel из рецепта не читаются.',
        'Имя корневого блока snapshot.': 'Имя входной таблицы.',
        'Число строк неизменяемого input snapshot.': 'Число строк в копии входной таблицы.',
        'Число столбцов input snapshot.': 'Число столбцов в копии входной таблицы.',
        'Читает scalar snapshot; ошибка Excel в ячейке даёт адресное исключение.':
            'Возвращает исходное значение из копии таблицы. При ошибке Excel сообщает адрес ячейки, которую нужно исправить.',
        'Явный fallback конкретного вызова. Пропущенный Optional Variant без default означает обязательное поле.':
            'Значение для пустой ячейки. В Text и Number без этого параметра пустая ячейка считается ошибкой; остальные методы используют указанное в объявлении значение по умолчанию.',
        'Nothing создаёт реальные default options.': 'Nothing создаёт условия размещения по умолчанию.',
        'исходный immutable snapshot не меняется': 'исходная геометрия не меняется',
        'Смысловой ID сущности для SourceMap.': 'Имя объекта, которому назначается ссылка на исходные данные.',
        'Возвращает последнюю диагностику query фасада с SourceMap. Nothing означает отсутствие записанного отказа.':
            'Возвращает последнюю геометрическую ошибку с адресом исходных данных. Если сохранённой ошибки нет, возвращает Nothing.',
        'Реальный адрес/программный источник, добавленный SourceMap или правилом.':
            'Адрес исходной ячейки либо пояснение, откуда взят параметр геометрии или арматуры.',
    }
    for original, replacement in replacements.items():
        text = text.replace(original, replacement)
    return text


def build():
    inventory = api.inventory()
    sections = selected_members(inventory)
    lines = ['Attribute VB_Name = "modSectionRecipeApiGuide"']
    builder_signatures = signatures((ROOT/'src/Section/CSectionDefinitionBuilder.cls').read_text(encoding='utf-8-sig'))

    def comment(text='', indent=0):
        prefix = "' " + ' ' * indent
        wrapped = textwrap.wrap(author_text(readable(text)), 102 - indent, break_long_words=False, break_on_hyphens=False)
        lines.extend(prefix + line for line in wrapped) if text else lines.append("'")

    def heading(title):
        comment(); comment()
        comment('=' * 100)
        comment(title)
        comment('=' * 100)
        comment()

    heading('СОЗДАНИЕ СЕЧЕНИЯ: ГЕОМЕТРИЯ, АРМАТУРА И АННОТАЦИИ')
    comment('Здесь собраны готовые методы для автора нового типа сечения. Описание задаёт, что нужно получить; программа по нему строит бетонную сетку, размещает стержни и готовит схему.', 2)
    comment()
    comment('Модуль состоит только из комментариев. Точные объявления методов приведены вместе с параметрами; код открывается в соответствующем классе через окно проекта Ctrl+R. Поиск по имени — Ctrl+F.', 2)
    comment()
    comment('СОДЕРЖАНИЕ')
    for number, title, members in sections:
        comment(number + '  ' + title, 4)
    heading('КАК ЧИТАТЬ ОПИСАНИЕ')
    for text in [
        'Все координаты, длины и диаметры внутри рецепта задаются в мм, площади — в мм^2, углы — в радианах. Значения из таблиц переводит context.Units; повторно переводить уже полученные мм не нужно.',
        'Path — геометрическая линия из отрезков и дуг. Замкнутая линия может быть границей бетона или отверстия. Линия центров стержней получается, например, смещением контура на as; вычислять координаты каждого стержня вручную не требуется.',
        'Segment — один отрезок или одна дуга. PathGroup объединяет несколько независимых линий. Station в имени метода означает расстояние от начала линии в мм.',
        'ID — уникальное текстовое имя линии, ряда, привязки или подписи. Anchor — именованная точка для размеров и пояснений. Rule — описание одного способа размещения стержней.',
        'Optional означает необязательный параметр; его значение по умолчанию указано после знака =. ByVal передаёт значение, ByRef позволяет методу вернуть результат через параметр. Nothing вместо options создаёт обычные условия размещения.',
        'Пробел и символ _ в конце строки VBA продолжают команду на следующей строке. Это знак записи кода, а не передаваемый параметр. Именованный аргумент записывается как имя:=значение.',
        'Initialize открывает одно описание, Finish завершает его. CirclePath, LinePath, RectanglePath и другие создающие команды сами сохраняют линию. Возвращённый path уже зарегистрирован; Outer, Opening и армирование получают его ID. AddPath нужен только для отдельно созданной линии.',
        'Таблицы и единицы context — независимый снимок ввода. context.Units используется для преобразования чисел; перенастраивать его нельзя. Проверка и размещение стержней выполняются общим Build до создания сетки.',
        'Команды подписей задают группу, геометрическую привязку и направление. Расстояние от линии арматуры до края бетона программа определяет по фактическому контуру; дополнительный отступ задаётся в настройках плоттера.',
    ]:
        comment(text, 2); comment()
    comment('МИНИМАЛЬНЫЙ ПРИМЕР', 2)
    comment('Полный класс CSectionRecipeMySection. Размеры заданы в мм; context требуется интерфейсу, входных таблиц в этом примере нет. Создайте Class Module с этим именем и вставьте текст:', 4)
    comment()
    minimal = (ROOT/'docs/geometry_refactor/examples/CSectionRecipeMySection.cls').read_text(encoding='utf-8-sig')
    minimal = minimal[minimal.index('Option Explicit'):]
    for line in minimal.splitlines():
        explanation = ''
        if re.search(r'\bfacade\.', line) and not line.lstrip().startswith("'"):
            explanation, _ = argument_comment(line.strip(), 'MySection', builder_signatures, {}, '')
        if explanation and len(line) + len(explanation) <= 165:
            lines.append("'       " + line + "  ' " + readable(explanation))
        else:
            if explanation:
                for note in textwrap.wrap(readable(explanation), 90, break_long_words=False, break_on_hyphens=False):
                    lines.append("'       ' " + note)
            lines.append("'       " + line if line else "'")
    comment()
    comment('В CSectionTypeRegistry.Class_Initialize добавьте RegisterRecipe "MySection", New CSectionRecipeMySection, Array(). Затем выполните UpdateSectionRecipeTypes. Пустой Array() подходит этому примеру с постоянными размерами; для размеров с листа «Сечения» укажите имя входной таблицы.', 4)

    comment('На листе «Сечения» каждому типу соответствует одна таблица: General содержит размеры, Main — основное армирование сторон, Extra — дополнительные ряды. Например: Set reader = context.Reader("rngPointedSectionGeometry"), B = reader.PositiveLength("General", "B"), n = reader.Count("Main", "Top", "n"). Ридер проверяет значения и переводит длины в мм. Для своей структуры доступны context.Table и ValueAt.', 4)
    comment()
    comment('УСТРОЙСТВО ШТАТНЫХ РЕЦЕПТОВ', 2)
    comment('В Circle, Rectangles, PointedSection и HollowRectangle сначала читаются и проверяются размеры и параметры включённых рядов. Подготовленные значения, адреса исходных ячеек и соответствие названий Config внутренним именам сторон сохраняются в Private Type самого рецепта. Для одного размера отдельный метод чтения не нужен.', 4)
    comment('CreateDefinition связывает этапы: подготовку данных, AddGeometry, AddReinforcement и AddAnnotations. AddGeometry задаёт границы бетона и отверстий. AddReinforcement создаёт линии центров арматуры и регистрирует правила размещения стержней. AddAnnotations задаёт размеры и подписи. Эти методы получают готовые параметры и не читают входные таблицы.', 4)
    comment('Параметры отключённого ряда пропускаются. В HollowRectangle внутренние as нужны и без стержней на соответствующей стороне: они ограничивают линии арматуры у отверстия. Ридер использует короткие ключи независимо от русских подписей и порядка полей.', 4)
    comment()
    comment('ПОДГОТОВКА КОНТАКТНЫХ РЯДОВ', 2)
    comment('При ненулевом базовом отступе R2 или R3 объявите RebarContour: имя Main, имя Outer/Opening, исходную фигуру до as, заданный as и явные имена R2/R3. Каждый ряд проверяется по своему общему контуру с отступами всех его участков. Координаты сохраняются; несовпадающие концы дуги и прямой соединяются отрезком, стык может быть негладким. Если скругление при отступе исчезло, дуга удаляется, а соседние прямые продолжаются до пересечения. Это само по себе не ошибка. Невозможный или неоднозначный контур даёт адресную ошибку. При a2 = a3 = 0 строятся контактные ряды без индивидуальных проверочных контуров.')
    comment('Например: facade.RebarContour "Top.Main", "OUTER", "Boundary", 50#, "Top.Row2", "Top.Row3". Здесь 50 мм — as основного ряда; Boundary — исходная бетонная фигура. Если направление Main известно, передайте normalSign:=1# или -1#, как в CRebarContactBuilder; это устраняет неоднозначность, когда Main находится посередине сечения. Для неактивной стороны, as которой ограничивает соседнюю линию, предусмотрен RebarContourBaseline.')
    comment('Пересекающиеся дополнительные стержни соседних граней объединяются с предупреждением: остаётся первый добавленный стержень с его координатами, диаметром и материалом. Это относится и к частичному пересечению. В подписи каждой грани учитываются только физически добавленные стержни её основной группы; объединённая позиция не увеличивает площадь или количество. Если стержней группы не осталось, подпись не выводится.')
    comment('ReadRebarInputs читает и проверяет параметры активных рядов. В AddReinforcement создают отдельный CRebarContactBuilder для каждого основного ряда или стороны, задают общие настройки дополнительных рядов и вызывают AppendRow один раз на включённый ряд по порядку. Полученные options передают в facade.RebarDependent. Отключённые ряды пропускают до чтения их необязательных полей; при отрисовке AppendRow не вызывают.', 4)
    comment('Initialize получает новые общие options именно дополнительных рядов и сохраняет их копию. Каждый AppendRow возвращает отдельный объект. Можно изменить SourceSuffix, Comment и другие индивидуальные поля перед передачей фасаду. Построитель прибавляет к сумме радиусов 2,5 мм на выступы арматуры периодического профиля для каждой соседней пары; диаметры и площади в расчёте не меняются. Диаметры промежуточных рядов и эти добавки учитываются отдельно по каждому направлению. EverySecondBar не уменьшает расстояние.', 4)
    comment('Фрагмент для уже объявленного основного ряда Top.Main диаметром 32 мм. Stacked направлен внутрь от верхней линии; SideBySide — вдоль неё. Смещения составят -28,5 и -49 мм по нормали, затем +24,5 мм вдоль линии. Все дополнительные ряды остаются связанными с Top.Main.', 4)
    for line in [
        'Dim contact As CRebarContactBuilder, extraOptions As CRebarRuleOptions',
        'Dim row2 As CRebarRuleOptions, row3 As CRebarRuleOptions, row4 As CRebarRuleOptions',
        'Set extraOptions = New CRebarRuleOptions',
        'extraOptions.AnnotationGroup = "Top"  \' Общая подпись основного и дополнительных рядов этой группы.',
        'Set contact = New CRebarContactBuilder',
        'contact.Initialize 32#, -1#, 1#, extraOptions  \' Основной диаметр, знаки нормали и касательной, общие настройки.',
        'Set row2 = contact.AppendRow(20#, "Stacked", "EverySecondBar", "Program.Row2")',
        'row2.SourceSuffix = "_R2"  \' Индивидуальное окончание имени позиции второго ряда.',
        'Set row3 = contact.AppendRow(16#, "Stacked", "EachBar", "Program.Row3")',
        'Set row4 = contact.AppendRow(12#, "SideBySide", "EachBar", "Program.Row4")',
        "' На этапе AddReinforcement используем готовые настройки; ID и связь с Main заданы явно.",
        'facade.RebarDependent "Top.Row2", "Top.Main", 20#, row2',
        'facade.RebarDependent "Top.Row3", "Top.Main", 16#, row3',
        'facade.RebarDependent "Top.Row4", "Top.Main", 12#, row4',
    ]:
        lines.append("'       " + line)
    comment()
    comment('ОПИСАТЕЛЬ И ПРИВЯЗКИ', 2)
    comment('CPolygonSpec объединяет исходные ID, X, Y, FilletRadius и SourceReference каждой вершины. New → AddVertex → Seal. Состояние читают Count/IsSealed, значения — VertexID/X/Y/FilletRadius/SourceReference по индексу от 1. После Seal изменение запрещено; повторный Seal безопасен. Первую вершину в конце не повторяют.', 4)
    comment('В конце AddVertex можно задать outgoingEdgeID — имя прямого участка к следующей вершине между касаниями, и cornerID — имя скругления этой вершины. Например: polygon.AddVertex "TR", 400#, 600#, 40#, outgoingEdgeID:="Right.UpperSlope", cornerID:="Round.TopRight". VertexID при этом остаётся "TR". Пустые аргументы дают Edge.TR и Corner.TR. OutgoingEdgeID/CornerID возвращают эффективные имена. Все имена сегментов одного многоугольника уникальны, включая исчезнувшие участки; при R=0 фиктивная дуга не создаётся.', 4)
    comment('PathID — имя всей линии; SegmentID — локальное имя её отрезка или дуги. В разных линиях допускаются одинаковые локальные имена. LoopID задаётся в Outer/Opening и обозначает границу бетона или отверстия. ContourID — имя сегмента при сохранении контура: PreserveSegmentIDs берёт реальные имена выбранной линии, SequentialByKind использует префиксы _LINE_/_ARC_ с единым счётчиком и порядком. Явный массив contourIDs меняет только публикацию. Ни роль границы, ни публикация не переименовывают исходную линию.', 4)
    comment('Chain получает имя исходной линии, локальные имена начального и конечного сегментов, параметры от 0 до 1 и направление обхода. Параметры задают точки начала и конца внутри этих сегментов; True выбирает обход исходной линии, False — обратный. На замкнутом контуре это два разных маршрута между теми же точками. Например, у нижнего сегмента PointedSection начало справа (0), конец слева (1): Chain от 1 к 0 с False идёт напрямую слева направо, а с True — через остальные стороны контура. При полном обороте от середины сегмента его вторая часть получает свободное имя .Part2, .Part3 и далее. SourcePathID/SourceSegmentID и SourceStartParameter/SourceEndParameter сохраняют исходную линию, сегмент и параметры отдельно; обращаться к частям по угаданному суффиксу не требуется.', 4)
    comment('ЯВНЫЕ ИМЕНА СТОРОН ПРЯМОУГОЛЬНИКА', 2)
    comment('Низ идёт слева направо, правая сторона снизу вверх, верх справа налево, левая сверху вниз. После Transform имена относятся к исходным сторонам и не назначаются заново по глобальным X/Y. Пример выбирает часть контура от середины верха до середины низа по прямому обходу через левую сторону. AdjacentRectangles сохраняет имена исходных сторон: у сложной границы может быть несколько верхних участков.', 4)
    for line in [
        'Dim facade As CSectionDefinitionBuilder, outline As CGeometryPath, side As CGeometryPath',
        'Set facade = New CSectionDefinitionBuilder: facade.Initialize "NamedRectangle"',
        'Set outline = facade.RectanglePath("Boundary", 0#, 0#, 400#, 600#, _',
        '    bottomEdgeID:="Bottom", topEdgeID:="Top")',
        'Set side = facade.Chain("LeftSide", outline.ID, "Top", 0.5, "Bottom", 0.5)',
        'facade.Outer "OUTER", outline.ID',
    ]:
        lines.append("'       " + line)
    comment()
    comment('Исходная вершина до скругления может не лежать на контуре. PathAnchor задаёт точку готового сегмента, Anchor — заданную координату. Вершины не превращаются в anchors автоматически. OffsetRoundedPolygon: плюс внутрь исходного многоугольника; Offset: плюс влево от направления готовой линии. У группы независимых линий Offset отсутствует.', 4)
    comment('ВЫБОР УЧАСТКА ПО КООРДИНАТАМ', 2)
    comment('ResolveAnchor находит существующий ID и параметр точки без знания автоматических имён сегментов. Координаты в мм должны лежать на готовой линии. В общем стыке нескольких сегментов возникает ошибка; выберите внутреннюю точку или задайте ID явно. Это поиск принадлежности, а не проекция ближайшей точки.', 4)
    comment('Пример выбирает правую грань прямоугольника и создаёт участок сверху вниз. t — положение середины; в Chain явно задаются концы 1 и 0, False означает обратный обход.', 4)
    for line in [
        'Dim facade As CSectionDefinitionBuilder, path As CGeometryPath, rightPath As CGeometryPath',
        'Dim rightID As String, t As Double',
        'Set facade = New CSectionDefinitionBuilder: facade.Initialize "AnchorExample"',
        'Set path = facade.RectanglePath("Rect", 0#, 0#, 400#, 600#)',
        'path.ResolveAnchor 400#, 300#, rightID, t',
        'Set rightPath = facade.Chain("Right", path.ID, rightID, 1#, rightID, 0#, False)',
    ]:
        lines.append("'       " + line)
    comment()
    comment('Aligned измеряет истинное расстояние, его линия параллельна измеряемому участку. Горизонтальный и вертикальный размеры измеряют проекции; PathLength суммирует прямые и дуги. Текст размера параллелен размерной линии и разворачивается для удобного чтения.', 4)
    code = (ROOT/'tests/modTestAudit02Api.bas').read_text(encoding='utf-8-sig')
    for name, title, condition in [
        ('TestBuildAudit02DemoDefinition', 'ПОЛНЫЙ ПРИМЕР: СКРУГЛЕНИЯ, АРМАТУРА И НАКЛОННЫЙ РАЗМЕР',
         'Это самостоятельная функция; b/h/w и все объекты объявлены внутри. Edge.TR — реальная ненулевая наклонная грань между точками касания. Армирование учебное, без проверки достаточности.'),
        ('TestBuildAudit02OperationsDefinition', 'ПОЛНЫЙ ПРИМЕР: ПРИМИТИВЫ, СМЕЩЕНИЯ И ПРОЕКЦИЯ',
         'Самостоятельная функция с постоянными данными в мм: прямоугольники смежные, отверстие внутри Outer, линии независимые. RowNormal=1 направлен влево от MainAxis; луч проекции (0,1) пересекает ProjectedAxis. Примеры можно поместить в тот же класс и вызвать вместо прямоугольного описания.')]:
        comment(); comment(title, 2); comment(condition, 4); comment()
        body = re.search(r'(?ms)^Private Function '+name+r'\(.*?^End Function', code).group(0)
        for line in body.splitlines():
            lines.append("'       " + line if line else "'")
    comment()

    options = (ROOT / 'src/Section/CRebarRuleOptions.cls').read_text(encoding='utf-8-sig')
    initialization = re.search(r'(?s)Private Sub Class_Initialize\(\)(.*?)End Sub', options)[1]
    defaults = dict(re.findall(r'(?m)^\s*(\w+) = ([^\r\n]+)', initialization))
    notes = {
        '04': 'Сначала задайте контур, затем объявите смещение на нужное as. Для второго и третьего независимого ряда задайте свои линии и свои as; связь с первым рядом не обязательна.',
        '05': 'Нулевой диаметр или нулевое число стержней отключают ряд. Пустой диаметр из таблицы читает OptionalDiameter как 0; одиночный дефис недопустим. Для каждого независимого ряда используется свой ID.',
        '06': 'Создайте New CRebarRuleOptions и измените только нужные свойства. Метод размещения копирует их при объявлении ряда, поэтому позднейшая правка options не меняет уже объявленные ряды.',
        '07': 'Для общей подписи стороны задайте одинаковую AnnotationGroup всем её рядам. Подпись считает фактически созданные стержни всей группы и исчезает, если их нет. Смещение подписи в настройках плоттера отсчитывается от границы бетона.',
        '08': 'Обычно достаточно готовых методов разделов 03–04. Для собственной последовательности отрезков и дуг: Initialize, AddLine/AddArc, Seal, затем facade.AddPath. После Seal менять состав линии нельзя.',
        '09': 'Source связывает имя геометрического объекта с адресом исходных данных. Ошибка должна указывать, какой размер или ряд нужно исправить. LastGeometryDiagnostic может вернуть Nothing, если сохранённой ошибки нет.',
    }
    all_members = []
    for number, title, members in sections:
        heading(number + '  ' + title.upper())
        if number in notes:
            comment(notes[number], 2); comment()
        owner = None
        for member in members:
            if member['owner'] != owner:
                owner = member['owner']
                comment(); comment('Класс / модуль: ' + owner, 2); comment()
            comment('-' * 94, 4)
            comment(member['id'], 4)
            comment()
            comment(author_text(purpose(member)), 6)
            comment()
            comment('Объявление:', 6)
            comment(member['signature'], 8)
            if owner == 'CRebarRuleOptions' and member['kind'] == 'Field':
                default = defaults.get(member['name'], 'False' if member.get('resultType') == 'Boolean' else '0' if member.get('resultType') in ('Long', 'Double') else 'пустое значение VBA')
                comment('По умолчанию: ' + default + '.', 6)
            if member.get('fields'):
                comment()
                comment('Допустимые значения / состав:', 6)
                for field in member['fields']:
                    comment(field, 8)
            if member.get('parameters'):
                comment()
                comment('Параметры:', 6)
                for parameter in member['parameters']:
                    comment(parameter['signature'], 8)
                    comment(author_text(parameter_meaning(member, parameter)), 10)
                    comment()
            comment()
            all_members.append(member)
    assert all(not line.strip() or line.lstrip().startswith("'") for line in lines[1:])
    target = ROOT / 'src/Common/modSectionRecipeApiGuide.bas'
    target.write_text('\n'.join(lines) + '\n', encoding='utf-8', newline='\r\n')
    manifest = {'scope': 'Методы автора геометрии, арматуры и аннотаций', 'sections': [{'number': n, 'title': t, 'members': [m['id'] for m in ms]} for n, t, ms in sections], 'members': all_members}
    (ROOT / 'docs/geometry_refactor/SECTION_RECIPE_AUTHOR_MANIFEST.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'status': 'COMMENT_ONLY_VBE_API', 'members': len(all_members), 'sections': len(sections), 'lines': len(lines)}))


if __name__ == '__main__':
    build()
