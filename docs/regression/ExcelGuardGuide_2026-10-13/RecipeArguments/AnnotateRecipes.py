"""Добавляет короткие пояснения к строкам декларативных рецептов без изменения команд."""
import hashlib
import json
import re
from pathlib import Path

ROOT = Path.cwd()
EVIDENCE = ROOT/'docs/regression/ExcelGuardGuide_2026-10-13/RecipeArguments'
EVIDENCE.mkdir(exist_ok=True)
FILES = sorted([* (ROOT/'src').rglob('*.cls'), * (ROOT/'docs/geometry_refactor/examples').rglob('*.cls')])
FILES = [p for p in FILES if re.search(r'(?m)^Implements ISectionRecipe\s*$', p.read_text(encoding='utf-8-sig'))]

def split_comment(line):
    quoted = False
    i = 0
    while i < len(line):
        if line[i] == '"':
            if quoted and i+1 < len(line) and line[i+1] == '"':
                i += 2
                continue
            quoted = not quoted
        elif line[i] == "'" and not quoted:
            return line[:i], line[i+1:]
        i += 1
    return line, ''

def statements(text):
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        start = i
        code, _ = split_comment(lines[i])
        joined = code.rstrip()
        while joined.endswith(' _'):
            i += 1
            next_code, _ = split_comment(lines[i])
            joined = joined[:-2].rstrip()+' '+next_code.strip()
        yield start+1, lines[start:i+1], joined
        i += 1

def executable(text):
    # Сохраняем границы команд и строковые литералы; убираем только комментарии
    # и пробелы, в том числе разрешённое VBA продолжение одной команды.
    result = []
    for _, _, code in statements(text):
        if not code.strip():
            continue
        tokens = re.split(r'("(?:[^"\n]|"")*")', code)
        result.append(''.join(token if i % 2 else re.sub(r'\s+', '', token) for i, token in enumerate(tokens)))
    return result

def required(code):
    c = code.strip()
    if not c or c.startswith(('Private ', 'Public ', 'Attribute ', 'Option ', 'Implements ')):
        return False
    return bool(re.search(r'\bfacade\.|\boptions\.\w+\s*=|\bNew C(?:SectionDefinitionBuilder|RebarRuleOptions)\b', c)
                or re.match(r'(?:If .*? Then )?(?:AddGeometry|AddBoundary|AddReinforcement|AddDimensions|AddRectangleRows|AddCorner)\b', c)
                or re.match(r'Set (?:ISectionRecipe_CreateDefinition|ProgramDefinition|Describe|AddGeometry|AxisPath|AdditionalOptions|ExtraOptions|outer|path|options)\s*=', c)
                or (c.startswith('Dim ') and re.search(r'C(?:GeometryPath|SectionDefinitionBuilder|RebarRuleOptions)\b', c))
                or re.match(r'(?:vertices|radii|corners|sizes\.Vertices|points\(count\)|Corners|ReDim (?:vertices|contourIDs|ids)|For i = 0 To count)', c)
                or re.search(r'\b(?:outer|boundary)\.CopySegment\b|contourIDs\(.*?\)\s*=|ids\(i - 1\)\s*=', c))

def explanation(code, owner):
    c = code.strip()
    if c.startswith('Dim '):
        if 'CSectionDefinitionBuilder' in c: return 'Объявляем сборщик описания сечения.'
        if 'CRebarRuleOptions' in c and 'CGeometryPath' not in c: return 'Объявляем условия размещения арматуры.'
        if 'CRebarRuleOptions' in c: return 'Объявляем линии геометрии и условия армирования.'
        return 'Объявляем линии и параметры геометрии.'
    if 'New CSectionDefinitionBuilder' in c:
        names = {'Circle':'круглого', 'RectSet':'составного', 'RoundedRectangle':'скруглённого', 'HollowRectangle':'полого', 'RoundedTrapezoid':'трапециевидного'}
        return 'Создаём описание '+names[owner]+' сечения.'
    if c.startswith('facade.Initialize'): return 'Начинаем описание составного сечения.'
    if 'facade.Finish' in c: return 'Проверяем и возвращаем готовое описание сечения.'
    if re.match(r'Set (?:ISectionRecipe_CreateDefinition|ProgramDefinition) = Describe', c): return 'Возвращаем описание геометрии, арматуры и размеров.'
    if re.match(r'Set AddGeometry =', c): return 'Возвращаем готовую наружную границу.'
    if re.match(r'Set (?:AdditionalOptions|ExtraOptions) =', c): return 'Возвращаем условия размещения дополнительного ряда.'
    if 'Set outer = upper' in c: return 'Принимаем один прямоугольник за наружный контур.'
    if re.search(r'\bfacade\.RectanglePath\b', c): return 'Создаём '+('верхний' if 'Set upper' in c else 'нижний')+' прямоугольник.'
    if 'facade.AdjacentRectangles' in c: return 'Получаем общий наружный контур двух прямоугольников.'
    if 'facade.RoundedPolygon' in c: return 'Создаём замкнутый контур с заданными скруглениями.'
    if 'facade.CirclePath' in c:
        if 'Opening.Circle' in c: return 'Создаём окружность отверстия.'
        if 'Circle.Axis' in c: return 'Создаём окружность центров стержней с отступом as.'
        return 'Создаём наружную окружность бетона.'
    if 'facade.OffsetRoundedPolygon' in c: return 'Смещаем контур к линии центров стержней.'
    if 'facade.Offset(' in c: return 'Смещаем окружность из отверстия в бетон для арматуры.'
    if 'facade.ConstrainLine' in c: return 'Ограничиваем концы ряда отступами соседних граней.'
    if 'facade.LinePath' in c:
        if 'Unlimited.' in c: return 'Создаём исходную прямую ряда у отверстия.'
        if 'x + axisDistance, y,' in c: return 'Левый ряд снизу вверх, с отступом as.'
        if 'x + width - axisDistance' in c: return 'Правый ряд сверху вниз, с отступом as.'
        if 'y + height - axisDistance' in c: return 'Верхний ряд слева направо, с отступом as.'
        return 'Нижний ряд справа налево, с отступом as.'
    if 'facade.Chain' in c:
        if 'Set outer' in c or 'Set boundary' in c: return 'Замыкаем контур с началом на выбранном участке.'
        if owner == 'RoundedRectangle':
            if 'B.Top' in c: return 'Выделяем прямую линию верхнего ряда.'
            if 'B.Bottom' in c: return 'Выделяем прямую линию нижнего ряда.'
            if '"Edge.TL", 0#' in c: return 'Выделяем левую линию ряда со скруглениями.'
            return 'Выделяем правую линию ряда со скруглениями.'
        if 'Case 3:' in c: return 'Выделяем прямую линию левого ряда.'
        if 'Case 4:' in c: return 'Выделяем прямую линию правого ряда.'
        return 'Выделяем верхнюю или нижнюю линию ряда с дугами.'
    if re.match(r'facade.Outer\b', c): return 'Назначаем наружную границу бетона.'
    if re.match(r'facade.Opening\b', c): return 'Объявляем отверстие внутри наружной границы.'
    if 'facade.Source' in c: return 'Привязываем геометрию к ячейкам исходных данных.'
    if 'facade.RebarByCount' in c: return 'Задаём основной ряд по числу и диаметру стержней.'
    if 'facade.RebarProjected' in c: return 'Проецируем наружный ряд на линию у отверстия.'
    if 'facade.RebarDependent' in c: return 'Задаём '+('второй' if '.Row2' in c else 'третий')+' ряд относительно основного.'
    if 'facade.RebarRowExtremeLabel' in c: return 'Привязываем подпись к крайнему ряду выбранной стороны.'
    if 'facade.RebarRowExtentLabel' in c: return 'Привязываем подпись к крайним стержням основного ряда.'
    if 'facade.RebarRowLabel' in c: return 'Привязываем общую подпись стороны к основному ряду.'
    if 'facade.RebarLabel' in c: return 'Объявляем общую подпись количества и диаметров группы.'
    if 'facade.HorizontalDistance' in c:
        if 'DIM_UPPER_OFFSET_X' in c: return 'Добавляем размер смещения верхнего прямоугольника.'
        if 'DIM_B_OPENING' in c: return 'Добавляем размер ширины отверстия.'
        if 'DIM_W_LEFT' in c: return 'Добавляем размер левого выступа.'
        if 'DIM_W_RIGHT' in c: return 'Добавляем размер правого выступа.'
        return 'Добавляем горизонтальный размер ширины B.'
    if 'facade.VerticalExtent' in c:
        if 'DIM_CIRCLE_D' in c: return 'Добавляем размер диаметра бетонного круга.'
        if 'DIM_H_OPENING' in c: return 'Добавляем размер высоты отверстия.'
        if 'DIM_RIGHT_HEIGHT' in c: return 'Добавляем размер высоты нижнего прямоугольника.'
        return 'Добавляем размер полной высоты сечения.'
    if 'facade.HorizontalExtent' in c:
        if 'DIM_TOP_WIDTH' in c: return 'Добавляем размер ширины верхнего прямоугольника.'
        return 'Добавляем размер полной ширины сечения.'
    if 'facade.Radius' in c: return 'Добавляем размер радиуса круглого отверстия.'
    if 'facade.Anchor' in c: return 'Создаём точку привязки названия сечения.'
    if 'facade.Text' in c: return 'Добавляем название сечения на схему.'
    if re.match(r'Set path = AxisPath', c): return 'Получаем линию центров стержней выбранной стороны.'
    if re.match(r'Set options = (?:AdditionalOptions|ExtraOptions)', c): return 'Получаем смещение и привязку дополнительного ряда.'
    if 'New CRebarRuleOptions' in c: return 'Создаём условия размещения арматуры.'
    if 'options.' in c:
        props = re.findall(r'options\.(\w+)\s*=', c)
        keys = set(props)
        if 'NormalOffset' in keys: return 'Выбираем смещение поперёк ряда или вдоль него.'
        if 'EverySecond' in keys:
            if 'Publication' in keys: return 'Все позиции или через одну; сохраняем их имена.'
            return 'Выбираем все исходные позиции или через одну.'
        if {'SourceNormalSign','SkipMissingProjection'} <= keys: return 'Задаём сторону проекции; пропуски допускаем.'
        if {'SuccessfulProjectionOrdinal','NewPublicationRoot'} <= keys: return 'Нумеруем успешные проекции как новый внутренний ряд.'
        if 'EndMode' in keys:
            if 'ExcludeBoth' in c: return 'Исключаем стержни на обоих концах линии.'
            if 'PlacementCheck' in keys: return 'Замкнутый ряд; стержни целиком в бетоне, без пересечений.'
            if len(keys) > 1: return 'Замкнутый ряд; имена и общая группа подписи.'
            return 'Равномерно распределяем стержни по замкнутой линии.'
        if 'StartSetback' in keys and 'EndSetback' in keys: return 'Задаём отступы ряда от обоих концов грани.'
        if keys == {'StartSetback'}: return 'Задаём отступ начала ряда от начала грани.'
        if keys == {'EndSetback'}: return 'Задаём отступ конца ряда от конца грани.'
        if 'PositionOrder' in keys: return 'Нумеруем позиции от середины линии.'
        if 'IncomingTangentAtJoint' in keys: return 'На стыках берём входящую касательную; задаём подпись.'
        if keys == {'SourceReference'}: return 'Привязываем ошибки размещения к ячейкам ввода.'
        if keys == {'Comment'}: return 'Задаём пояснение к стержням этого ряда.'
        if keys == {'SourceSuffix','AnnotationGroup'}: return 'Задаём суффикс ряда и общую группу подписи.'
        if keys == {'Publication','PlacementCheck'}: return 'Сохраняем исходные имена; проверяем весь стержень.'
        if keys == {'Publication','Coincidence'}: return 'Сохраняем имена и владельца совпавшего стержня.'
        if 'Publication' in keys and 'AnnotationGroup' in keys: return 'Задаём имена стержней, владельца и общую подпись.'
        if keys == {'Publication','Comment'}: return 'Задаём имена стержней и пояснение первого ряда.'
        if {'Coincidence','AnnotationGroup','SourceReference'} <= keys: return 'Сохраняем владельца, группу подписи и адрес ввода.'
        if keys == {'AnnotationGroup','SourcePrefix'}: return 'Задаём имена стержней и общую группу подписи.'
        if 'AnnotationGroup' in keys: return 'Задаём общую группу подписи стержней.'
        if keys == {'Publication'}: return 'Сохраняем имена стержней по исходным позициям.'
        if keys == {'Coincidence'}: return 'Совпавший стержень сохраняет первого владельца.'
        raise ValueError('Unexplained options: '+c)
    if re.match(r'(?:If .* Then )?AddCorner\b', c):
        for name, caption in [('TL','левую верхнюю'),('TR','правую верхнюю'),('BR','правую нижнюю'),('BL','левую нижнюю'),('LeftNose','вершину левого выступа'),('RightNose','вершину правого выступа')]:
            if '"'+name+'"' in c: return 'Добавляем '+caption+(' и её радиус.' if 'Nose' in name else ' вершину и её радиус.')
    if c.startswith('AddBoundary facade'): return 'Добавляем '+('границу отверстия.' if '"OPENING"' in c else 'наружную границу бетона.')
    if c.startswith('AddGeometry facade') or 'Set outer = AddGeometry' in c: return 'Описываем наружную границу и геометрию сечения.'
    if c.startswith('AddReinforcement'): return 'Описываем армирование всех активных сторон.'
    if c.startswith('AddDimensions'): return 'Описываем размеры сечения.'
    if 'AddRectangleRows facade' in c: return 'Описываем армирование '+('нижнего' if 'Then AddRectangleRows' in c else 'верхнего')+' прямоугольника.'
    if c.startswith('vertices ='): return 'Задаём координаты четырёх вершин трапеции.'
    if c.startswith('radii ='): return 'Назначаем радиусы скругления вершин.'
    if c.startswith('corners ='): return 'Назначаем имена четырём вершинам трапеции.'
    if c.startswith('sizes.Vertices ='): return 'Сохраняем вершины, радиусы и имена контура.'
    if c.startswith('points(count)'): return 'Записываем координаты, радиус и имя новой вершины.'
    if c.startswith('Corners ='): return 'Задаём четыре вершины прямоугольной границы.'
    if c.startswith('ReDim vertices'): return 'Готовим массивы вершин, радиусов и имён.'
    if c.startswith('For i = 0 To count'): return 'Копируем подготовленные вершины, радиусы и имена.'
    if c.startswith('ReDim '): return 'Готовим имена участков контура.'
    if 'CopySegment' in c: return 'Получаем участок контура для назначения имени.'
    if 'contourIDs(' in c or 'ids(i - 1)' in c: return 'Назначаем имена прямым и дуговым участкам контура.'
    raise ValueError('Missing explanation: '+owner+': '+c)

from RecipeArgumentText import argument_comment,signatures
import textwrap
builder_signatures=signatures((ROOT/'src/Section/CSectionDefinitionBuilder.cls').read_text(encoding='utf-8-sig'))
inventory=[]
for path in FILES:
    before=path.read_text(encoding='utf-8-sig')
    backup=EVIDENCE/'SourceBefore'/path.relative_to(ROOT)
    backup.parent.mkdir(parents=True,exist_ok=True)
    if backup.exists():
        original=backup.read_text(encoding='utf-8-sig')
        assert executable(original)==executable(before)
        before=original
    else:
        backup.write_bytes(path.read_bytes())
    helper_signatures=signatures(before)
    output,entries=[],[]
    for original_line,raw,code in statements(before):
        if required(code):
            fallback=explanation(code,path.stem.removeprefix('CSectionRecipe'))
            description,args=argument_comment(code.strip(),path.stem,builder_signatures,helper_signatures,fallback)
            code=code.rstrip()
            indent=code[:len(code)-len(code.lstrip())]
            if len(code)+len(description)+4 <= 240 and len(description) <= 170:
                placement='inline'
                output.append(code+"  ' "+description)
            else:
                placement='above'
                for line in textwrap.wrap(description,100-len(indent),break_long_words=False,break_on_hyphens=False):
                    output.append(indent+"' "+line)
                output.append(code)
            entries.append({'line':len(output),'originalLine':original_line,'code':code.strip(),'comment':description,'placement':placement,'arguments':args})
        else:
            output.extend(raw)
    after='\n'.join(output)+'\n'
    assert executable(before)==executable(after),path
    assert all(len(line.encode('cp1251'))<900 for line in output),path
    path.write_text(after,encoding='utf-8',newline='\r\n')
    inventory.append({'source':str(path.relative_to(ROOT)).replace('\\','/'),'commentedLines':len(entries),'linesBefore':len(before.splitlines()),'linesAfter':len(output),'entries':entries,'executableCommandsUnchanged':True,'sourceSHA256':hashlib.sha256(path.read_bytes()).hexdigest()})
result={'recipes':inventory,'totalCommentedCommands':sum(x['commentedLines'] for x in inventory),'inline':sum(e['placement']=='inline' for x in inventory for e in x['entries']),'above':sum(e['placement']=='above' for x in inventory for e in x['entries']),'argumentsExplained':sum(len(e['arguments']) for x in inventory for e in x['entries'])}
(EVIDENCE/'SourceComments.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps({k:v for k,v in result.items() if k!='recipes'}))
