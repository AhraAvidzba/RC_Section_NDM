"""Редакционная правка четырёх рецептов с проверкой неизменности команд VBA."""
from pathlib import Path
import json
import re
import subprocess

ROOT = Path(__file__).resolve().parents[4]
WIDTH = 200

HEADERS = {
    'Circle': ('Круглое бетонное сечение с арматурой', [
        'Класс читает диаметр сечения и параметры армирования из таблицы круглого сечения.',
        'Основной ряд располагается по окружности на расстоянии as от края бетона. Второй и',
        'третий ряды привязаны к стержням первого ряда и могут находиться рядом с ними или глубже.',
        '',
        'Рецепт возвращает описание контура, рядов и подписей. По этому описанию общий',
        'CSectionBuilder создаёт стержни и бетонную сетку; сам рецепт их не расставляет.']),
    'RectSet': ('Прямоугольное и составное бетонное сечение', [
        'Класс описывает один прямоугольник, Г-образное сечение или два смежных прямоугольника.',
        'В составном сечении верхняя часть может быть смещена относительно нижней.',
        'У каждой грани свои отступы и армирование; дополнительные ряды связаны с основным.',
        '',
        'Размеры читаются из таблицы выбранного варианта. Контуры, ряды и размеры передаются',
        'в CSectionDefinitionBuilder, а стержни и бетонную сетку создаёт общий CSectionBuilder.']),
    'RoundedRectangle': ('Бетонное сечение со скруглёнными боковыми сторонами', [
        'Класс описывает сечение с прямой горизонтальной базой и двумя боковыми сторонами.',
        'Каждая сторона может быть обычной или иметь выступ; радиусы задаются независимо.',
        'Линии центров арматуры получаются смещением того же многоугольника, что задаёт бетон.',
        '',
        'Рецепт читает размеры и армирование, затем возвращает описание контура, рядов и',
        'размеров. Общие геометрические классы строят дуги и смещения, а CSectionBuilder',
        'создаёт стержни и бетонную сетку.']),
    'HollowRectangle': ('Прямоугольное бетонное сечение с отверстием', [
        'Класс описывает наружный прямоугольник и внутреннее отверстие со скруглёнными углами.',
        'Отверстие может быть смещено относительно центра наружного контура.',
        'Наружные ряды задаются по числу стержней. Ряды у отверстия получаются проекцией',
        'наружных рядов; на каждой стороне могут быть ещё два дополнительных ряда.',
        '',
        'Рецепт возвращает описание обеих границ, армирования и размеров. Построение дуг,',
        'проекцию стержней и создание бетонной сетки выполняют общие классы.']),
}

METHODS = {
    'Circle': {
        'ISectionRecipe_CreateDefinition': [
            'Читает размеры круга и параметры включённых рядов арматуры. Возвращает готовое',
            'описание наружной окружности, армирования и размера диаметра для общего сборщика.'],
        'AdditionalOptions': [
            'Возвращает настройки дополнительного ряда: направление и расстояние от основного',
            'стержня, суффикс имени и общую группу подписи. При наличии второго ряда учитывает его диаметр.'],
        'Require': ['Проверяет условие для указанного параметра. При нарушении сообщает причину и адрес его ячейки.'],
    },
    'RectSet': {
        'ISectionRecipe_CreateDefinition': [
            'Читает выбранный вариант сечения и его армирование. Возвращает описание общей',
            'наружной границы, рядов на гранях и размеров для общего сборщика.'],
        'ReadDimensions': [
            'Заполняет sizes размерами выбранного варианта в миллиметрах и проверяет их допустимость.',
            'Для одного прямоугольника не читает размеры нижней части; смещение нужно только двум прямоугольникам.',
            'В каждой паре сначала читаются H и B, затем проверяются их знаки. Так текст в B не',
            'теряется за ошибкой знака H, а сообщение сохраняет прежний адрес исходной ячейки.'],
        'AddGeometry': [
            'Добавляет контуры верхнего и, если он нужен, нижнего прямоугольника. Через upper,',
            'lower и outer отдаёт их линии и общую наружную границу; для одного прямоугольника lower пуст.'],
        'AddDimensions': ['Добавляет на схему общие габариты, размеры отдельных частей и ненулевое смещение верхней части.'],
        'AddRectangleRows': [
            'Задаёт основной и дополнительные ряды на четырёх гранях одного прямоугольника.',
            'Отступы считаются от полной грани; подпись каждой стороны объединяет все её ряды.'],
        'AdditionalOptions': [
            'Возвращает настройки дополнительного ряда: контактное расстояние, направление',
            'и размещение у каждого основного стержня или через один. Все ряды стороны имеют общую подпись.'],
        'Require': ['Проверяет условие для указанной ячейки. При нарушении сообщает причину и её фактический адрес.'],
    },
    'RoundedRectangle': {
        'ISectionRecipe_CreateDefinition': [
            'Читает размеры и армирование сечения со скруглёнными сторонами. Возвращает',
            'описание наружного контура, рядов арматуры и размеров для общего сборщика.'],
        'ReadDimensions': [
            'Заполняет sizes размерами в миллиметрах и проверяет параметры обеих сторон.',
            'Для обычной стороны не читает ширину выступа и его радиус; сохраняет адреса активных полей.'],
        'AddGeometry': [
            'Создаёт наружный контур по заданным вершинам и радиусам, задаёт порядок его обхода',
            'и возвращает зарегистрированную границу бетона. Ошибку связывает с исходными размерами.'],
        'DescribePolygon': [
            'Возвращает завершённое описание вершин и радиусов: четыре основных угла и вершины',
            'включённых выступов. Этот же объект используется для контура бетона и линий арматуры.'],
        'AddReinforcement': [
            'Задаёт ряды по четырём сторонам и объединяет подписи основного и дополнительных рядов.',
            'Порядок сторон сохраняет принадлежность общего углового стержня первому из заданных рядов.'],
        'AxisPath': [
            'Возвращает линию центров стержней выбранной стороны после смещения контура на as.',
            'Для верхней и нижней сторон берёт прямой участок, для боковых — также скругления;',
            'обе боковые линии направлены сверху вниз.'],
        'AdditionalOptions': [
            'Возвращает настройки дополнительного ряда: расстояние между центрами с учётом',
            'диаметров, направление смещения и выбор основных стержней, у которых он появляется.'],
        'AddDimensions': ['Добавляет ширину прямой базы B, полную высоту H и фактические размеры включённых боковых выступов.'],
        'Require': ['Проверяет условие для указанной ячейки. При нарушении сообщает причину и её фактический адрес.'],
    },
    'HollowRectangle': {
        'ISectionRecipe_CreateDefinition': [
            'Читает размеры обеих границ и параметры армирования. Возвращает описание бетона',
            'с отверстием, наружных и внутренних рядов и размеров для общего сборщика.'],
        'ReadDimensions': [
            'Заполняет sizes наружными размерами, размерами и смещением отверстия в миллиметрах.',
            'Проверяет размеры и радиусы, сохраняет адреса полей для сообщений об ошибках.'],
        'ReadFaces': [
            'Заполняет faces параметрами четырёх наружных и четырёх внутренних сторон.',
            'У выключенных рядов не читает ненужные параметры. Внутренние as читает всегда:',
            'они ограничивают линии у отверстия даже при отсутствии арматуры на соседней стороне.'],
        'AddGeometry': ['Добавляет в описание сечения наружную границу и отверстие по двум подготовленным многоугольникам.'],
        'AddBoundary': [
            'Создаёт замкнутую линию по вершинам и радиусам, задаёт порядок обхода и имена её участков.',
            'Назначает её наружной границей или отверстием внутри OUTER, в зависимости от opening.'],
        'DescribePolygon': [
            'Возвращает завершённое описание четырёх углов прямоугольника с заданным радиусом.',
            'B и H задают габариты, x и y — центр; нулевой радиус оставляет углы без скругления.'],
        'AddReinforcement': [
            'Задаёт наружные ряды по числу стержней, а внутренние — проекцией соответствующих наружных.',
            'Добавляет связанные ряды и общую подпись каждой стороны по фактически созданной арматуре.'],
        'AxisPath': [
            'Возвращает линию центров стержней у выбранной наружной стороны или стороны отверстия.',
            'Горизонтальные линии включают скругления, вертикальные используют прямую грань.',
            'При нескруглённом отверстии концы ограничиваются отступами соседних внутренних сторон.'],
        'ExtraOptions': [
            'Возвращает настройки дополнительного ряда: контактное расстояние, направление',
            'и размещение у каждого основного стержня или через один, начиная с первого.'],
        'AddDimensions': ['Добавляет наружные ширину и высоту сечения, а также ширину и высоту внутреннего отверстия.'],
        'Require': ['Проверяет условие для указанной ячейки. При нарушении сообщает причину и её фактический адрес.'],
    },
}


def without_comment(line):
    quoted = False
    i = 0
    while i < len(line):
        if line[i] == '"':
            if quoted and i+1 < len(line) and line[i+1] == '"':
                i += 2
                continue
            quoted = not quoted
        if line[i] == "'" and not quoted:
            return line[:i].rstrip()
        i += 1
    return line.rstrip()


def commands(text):
    joined = re.sub(r' _\r?\n\s*', ' ', '\n'.join(without_comment(s) for s in text.splitlines()))
    return [''.join(token if i % 2 else re.sub(r'\s+', '', token)
                    for i, token in enumerate(re.split(r'("(?:[^"\n]|"")*")', s)))
            for s in joined.splitlines() if s.strip()]


def inline(kind, method, code):
    s = code.strip()
    if s.startswith('Set table = context.Table('): return 'Берём копию входной таблицы и единицы измерения.'
    if s.startswith('ReadDimensions table,'): return 'Читаем и проверяем размеры сечения.'
    if s.startswith('ReadFaces table,'): return 'Читаем армирование наружных и внутренних сторон.'
    if s.startswith('Set facade = New'): return 'Создаём сборщик описания сечения.'
    if s.startswith('facade.Initialize'): return 'Начинаем описание выбранного типа.'
    if s.startswith('Set ISectionRecipe_CreateDefinition ='): return 'Завершаем и возвращаем проверенное описание сечения.'
    if s.startswith('Set options = New'): return 'Создаём настройки размещения этого ряда.'
    if s.startswith('options.EndMode = ClosedUniform:'): return 'Равномерно по кругу без повторного конца; имена начинаются с B.'
    if 'options.EndMode = ExcludeBoth' in s: return 'Исключаем стержни на обоих концах линии.'
    if 'options.StartSetback =' in s: return 'Отступ от начала и, если задан, от конца линии, мм.' if ': options.EndSetback' in s else 'Отступ от начала линии, мм.'
    if s.startswith('options.EndSetback ='): return 'Отступ от конца линии, мм.'
    if s.startswith('options.PositionOrder ='): return 'Нумеруем стержни от середины линии.'
    if s.startswith('options.SetPublication "Circle"'): return 'Одна подпись для всех рядов; совпадение стержней считаем ошибкой.'
    if s.startswith('options.SetPublication'): return 'Общая подпись стороны; совпавший стержень остаётся у первого ряда.'
    if s.startswith('options.SourceReference ='): return 'Адреса полей для сообщения об ошибке размещения.'
    if s.startswith('options.IncomingTangentAtJoint ='): return 'На стыке используем направление предыдущего участка.'
    if s.startswith('options.SourceNormalSign ='): return 'Проекция идёт к отверстию; пропускаем лучи без пересечения.'
    if s.startswith('options.SuccessfulProjectionOrdinal ='): return 'Нумеруем только найденные проекции как самостоятельный ряд.'
    if s.startswith('options.ConfigureDependentContact'): return 'Учитываем диаметры обоих рядов и промежуточные стержни.'
    if s.startswith('options.SourceSuffix ='): return 'Добавляем к имени основного стержня суффикс ряда.'
    if s.startswith('options.PlacementCheck ='): return 'Весь круг стержня должен находиться в бетоне.'
    if s.startswith('everySecond ='): return 'Выбираем все позиции или через одну, начиная с первой.'
    if s.startswith('axis = RowNormal:'): return 'По умолчанию смещаем ряд по нормали к исходной линии.'
    if s.startswith('If location = "SideBySide"'): return 'Располагаем дополнительный стержень вдоль исходного ряда.'
    if 'loc2 = loc3 Then' in s or 'Location2 = faces(i).Location3 Then' in s: return 'Третий ряд в том же направлении должен обойти стержень второго.'
    if s.startswith('Set AdditionalOptions =') or s.startswith('Set ExtraOptions ='): return 'Возвращаем настройки дополнительного ряда.'
    if s.startswith('Set polygon = DescribePolygon'): return 'Общее описание вершин для бетона и всех смещённых линий.'
    if s.startswith('Set outerPolygon = DescribePolygon'): return 'Наружный прямоугольник с центром в начале координат.'
    if s.startswith('Set openingPolygon = DescribePolygon'): return 'Отверстие с собственными размерами, центром и радиусом.'
    if s.startswith('Set original = facade.RoundedPolygon'): return 'Создаём контур по исходным вершинам и радиусам.'
    if s.startswith('Set shape = facade.RoundedPolygon'): return 'Создаём контур по подготовленным вершинам и радиусам.'
    if s.startswith('Set outer = facade.Chain') or s.startswith('Set boundary = facade.Chain'): return 'Полный обход контура от выбранного начального участка.'
    if s.startswith('Set upper = facade.RectanglePath'): return 'Верхняя часть: левый нижний угол, ширина и высота, мм.'
    if s.startswith('Set lower = facade.RectanglePath'): return 'Нижняя часть начинается в (0, 0); B2 и H2 — её габариты, мм.'
    if s.startswith('Set outer = facade.AdjacentRectangles'): return 'Получаем общую границу двух смежных прямоугольников.'
    if s.startswith('Set outer = upper'): return 'У одиночного прямоугольника его контур и есть наружная граница.'
    if s.startswith('Set path = facade.CirclePath("CONTOUR_CIRCLE"'): return 'Круг бетона с центром (0, 0) и радиусом D/2, мм.'
    if s.startswith('Set path = facade.CirclePath("Circle.Axis"'): return 'Линия центров основного ряда: радиус D/2 - as, мм.'
    if s.startswith('facade.Outer'): return 'Назначаем эту замкнутую линию наружной границей бетона.'
    if s.startswith('facade.Opening'): return 'Назначаем отверстие внутри наружной границы OUTER.'
    if s.startswith('AddGeometry facade,'): return 'Добавляем подготовленные границы бетона.'
    if s.startswith('Set outer = AddGeometry'): return 'Получаем наружный контур по подготовленному многоугольнику.'
    if s.startswith('AddReinforcement facade,'): return 'Задаём ряды и их общие подписи по параметрам таблицы.'
    if s.startswith('AddDimensions facade,'): return 'Добавляем габариты и размеры сечения.'
    if s.startswith('AddRectangleRows facade,') and ', 0,' in s: return 'Армирование четырёх граней верхнего прямоугольника.'
    if 'Then AddRectangleRows facade,' in s: return 'Армирование нижней части, если она есть.'
    if s.startswith('Set path = AxisPath('): return 'Получаем линию центров стержней выбранной стороны с отступом as.'
    if s.startswith('Set options = AdditionalOptions(') or s.startswith('Set options = ExtraOptions('): return 'Настраиваем смещение и привязку добавляемого ряда.'
    if s.startswith('facade.RebarByCount'): return 'Задаём основной ряд по числу стержней и диаметру.'
    if s.startswith('facade.RebarDependent') and 'Row2' in s: return 'Второй ряд привязан к позициям основного; задаём его диаметр.'
    if s.startswith('facade.RebarDependent'): return 'Третий ряд привязан к позициям основного; задаём его диаметр.'
    if s.startswith('facade.RebarProjected'): return 'Проецируем стержни наружного ряда на линию у отверстия.'
    if s.startswith('facade.RebarLabel "Circle"'): return 'Общая подпись всех круговых рядов над сечением.'
    if s.startswith('facade.RebarLabel'): return 'Общую подпись рядов привязываем к границе их линии.'
    if s.startswith('facade.RebarRowLabel'): return 'Подпись учитывает все ряды стороны и крайние стержни основного.'
    if s.startswith('facade.RebarRowExtentLabel'): return 'Привязываем подпись к крайним стержням основного ряда.'
    if s.startswith('facade.RebarRowExtremeLabel'): return 'Для горизонтальной стороны подпись берём от крайней линии стержней.'
    if s.startswith('Set offset = facade.OffsetRoundedPolygon'): return 'Смещаем тот же многоугольник внутрь бетона на as, мм.'
    if s.startswith('Set shape = facade.OffsetRoundedPolygon'): return 'Смещаем наружный контур внутрь, а контур отверстия — в бетон.'
    if 'facade.Chain(groupID, offset.ID' in s:
        if '"B.Top"' in s: return 'Верхняя прямая грань слева направо.'
        if '"B.Bottom"' in s: return 'Нижняя прямая грань справа налево.'
        return 'Боковая сторона со скруглениями, сверху вниз.'
    if 'facade.Chain(faces(index).ID' in s:
        if 'Case 3:' in s: return 'Левая вертикальная грань.'
        if 'Case 4:' in s: return 'Правая вертикальная грань с обратным обходом.'
        return 'Горизонтальная сторона вместе с прилегающими скруглениями.'
    if 'facade.ConstrainLine' in s: return 'Ограничиваем линию отступами двух соседних внутренних сторон.'
    if s.startswith('facade.Source faces(index).ID'): return 'Связываем ошибки этой линии с её исходными полями.'
    if 'facade.LinePath("Unlimited.' in s: return 'Прямая у отверстия до ограничения соседними отступами.'
    if 'facade.LinePath(groupID' in s:
        if 'x + width - axisDistance' in s: return 'Правая грань на расстоянии as от бетона, сверху вниз.'
        if 'y + height - axisDistance' in s: return 'Верхняя грань на расстоянии as от бетона, слева направо.'
        if 'x + width, y + axisDistance' in s: return 'Нижняя грань на расстоянии as от бетона, справа налево.'
        return 'Левая грань на расстоянии as от бетона, снизу вверх.'
    if s.startswith('polygon.AddVertex') or 'Then polygon.AddVertex' in s:
        for name, caption in [('"TL"','Верхний левый угол'),('"TR"','Верхний правый угол'),('"BR"','Нижний правый угол'),('"BL"','Нижний левый угол'),('"LeftNose"','Вершина левого выступа'),('"RightNose"','Вершина правого выступа')]:
            if name in s: return caption+' до скругления; координаты и радиус в мм.'
    if s == 'polygon.Seal': return 'Проверяем вершины и завершаем их описание.'
    if s.startswith('Set DescribePolygon ='): return 'Возвращаем подготовленный многоугольник.'
    if s.startswith('facade.VerticalExtent') or 'facade.HorizontalDistance' in s or s.startswith('facade.HorizontalExtent'):
        labels = {
            'DIM_CIRCLE_D':'Диаметр бетонного круга; размер вынесен влево.',
            'DIM_TOP_WIDTH':'Ширина верхнего прямоугольника; размер над сечением.',
            'DIM_RIGHT_HEIGHT':'Высота нижнего прямоугольника; размер справа.',
            'DIM_UPPER_OFFSET_X':'Ненулевое смещение верхней части относительно нижней.',
            'DIM_W_LEFT':'Фактическая ширина левого выступа; подпись округляется до 1 мм.',
            'DIM_W_RIGHT':'Фактическая ширина правого выступа; подпись округляется до 1 мм.',
            'DIM_B_OPENING':'Ширина отверстия; размер под сечением.',
            'DIM_H_OPENING':'Высота отверстия; размер справа от сечения.',
            'DIM_B':'Ширина прямой базы B; размер сверху.' if kind == 'RoundedRectangle' else 'Общая ширина сечения B.',
            'DIM_H':'Полная высота сечения H; размер вынесен влево.',
        }
        for name, caption in labels.items():
            if '"'+name+'"' in s: return caption
    return ''


def rewrite(kind):
    path = ROOT/'src/Geometry'/('CSectionRecipe'+kind+'.cls')
    before = Path(__file__).parent/(path.stem+'.before.cls')
    original = before.read_text(encoding='utf-8-sig') if before.exists() else path.read_text(encoding='utf-8-sig')
    if not before.exists(): before.write_text(original, encoding='utf-8')
    prefix, body = original.split('Option Explicit', 1)
    physical = [without_comment(s) for s in body.splitlines() if not s.lstrip().startswith("'")]
    output = ['Option Explicit', '', "' " + '='*74, "' "+HEADERS[kind][0], "' " + '='*74]
    output += ["'"+(' '+s if s else '') for s in HEADERS[kind][1]]
    output += ['']
    method = ''
    found = []
    i = 0
    while i < len(physical):
        line = physical[i]
        i += 1
        if not line.strip():
            if output[-1] != '': output.append('')
            continue
        chunk = [line]
        while chunk[-1].endswith(' _'):
            chunk.append(physical[i]); i += 1
        logical = ' '.join(s.rstrip()[:-2].strip() if s.endswith(' _') else s.strip() for s in chunk)
        indent = line[:len(line)-len(line.lstrip())]
        if len(indent+logical) <= WIDTH:
            chunk = [indent+logical]
        match = re.match(r'Private (?:Function|Sub) (\w+)\b', logical)
        if match:
            method = match[1]
            found.append(method)
            output += ["' "+s for s in METHODS[kind][method]]
        comment = inline(kind, method, logical)
        if comment:
            if len(chunk[-1])+len(comment)+3 <= WIDTH:
                chunk[-1] += " ' "+comment
            else:
                output.append(indent+"' "+comment)
        output += chunk
    assert set(found) == set(METHODS[kind]), (kind, found)
    result = prefix+'\n'.join(output).rstrip()+'\n'
    assert commands(original) == commands(result), 'Commands changed: '+kind
    path.write_bytes(result.replace('\r\n','\n').replace('\n','\r\n').encode('utf-8'))
    return {'Class':path.stem,'Methods':found,'CommandsUnchanged':True,'BeforeLines':len(original.splitlines()),'AfterLines':len(result.splitlines())}


if __name__ == '__main__':
    report = [rewrite(kind) for kind in HEADERS]
    (Path(__file__).parent/'RecipeCommentRewrite.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False))
