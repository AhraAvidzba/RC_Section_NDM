"""Смысл переданных аргументов декларативного API; комментарии для автора рецепта."""
import re

MEANINGS = {
 'Initialize': {'typeID':'имя типа сечения'},
 'CirclePath': {'pathID':'имя новой окружности','centerX':'X центра, мм','centerY':'Y центра, мм','radius':'радиус, мм'},
 'RectanglePath': {'pathID':'имя прямоугольника','x':'X левого нижнего угла, мм','y':'Y левого нижнего угла, мм','width':'ширина, мм','height':'высота, мм'},
 'RoundedPolygon': {'pathID':'имя контура','polygon':'завершённое описание исходных вершин и радиусов, мм'},
 'OffsetRoundedPolygon': {'pathID':'имя линии центров стержней','polygon':'то же завершённое описание исходных вершин','distance':'смещение, мм: положительное внутрь, отрицательное наружу'},
 'Offset': {'pathID':'имя новой линии','sourcePathID':'исходная линия','signedLeftDistance':'смещение влево от обхода, мм; минус означает вправо'},
 'LinePath': {'pathID':'имя линии','x1':'X начала, мм','y1':'Y начала, мм','x2':'X конца, мм','y2':'Y конца, мм'},
 'AdjacentRectangles': {'pathID':'имя общего контура','rectangles':'прямоугольники: X,Y нижнего угла, ширина, высота, мм'},
 'Chain': {'pathID':'имя новой линии','sourcePathID':'исходный контур','startSegmentID':'участок начала','startParameter':'точка начала участка: 0 начало, 1 конец','endSegmentID':'участок конца','endParameter':'точка конца участка: 0 начало, 1 конец','forward':'True по обходу исходной линии, False против','fullLoopIfSame':'True означает полный контур при совпавших концах'},
 'ConstrainLine': {'pathID':'имя ограниченной линии','sourcePathID':'исходная прямая','startNX':'X-компонента нормали первой границы','startNY':'Y-компонента нормали первой границы','startBound':'первая граница: nX*X+nY*Y=bound, мм','endNX':'X-компонента нормали второй границы','endNY':'Y-компонента нормали второй границы','endBound':'вторая граница: nX*X+nY*Y=bound, мм'},
 'Outer': {'loopID':'имя наружной границы бетона','pathID':'замкнутый контур','comment':'пояснение к границе','contourIDs':'имена отдельных участков контура','arcComment':'отдельное пояснение к дугам'},
 'Opening': {'loopID':'имя отверстия','pathID':'замкнутый контур отверстия','parentOuterID':'наружная граница, внутри которой находится отверстие','comment':'пояснение к границе отверстия','contourIDs':'имена отдельных участков','arcComment':'отдельное пояснение к дугам'},
 'Source': {'entityID':'имя геометрии или ряда','sourceReference':'адреса исходных ячеек для сообщений об ошибках'},
 'RebarByCount': {'ruleID':'имя нового ряда','target':'линия центров стержней','count':'заданное число стержней','diameter':'диаметр стержня, мм','options':'отступы, нумерация и условия размещения','steelClass':'группа материала арматуры'},
 'RebarDependent': {'ruleID':'имя дополнительного ряда','sourceRuleID':'исходный ряд, чьи позиции используются','diameter':'диаметр дополнительного стержня, мм','options':'смещение и выбор позиций исходного ряда','steelClass':'группа материала арматуры'},
 'RebarProjected': {'ruleID':'имя внутреннего ряда','target':'линия, на которую проецируются позиции','sourceRuleID':'исходный наружный ряд','diameter':'диаметр внутреннего стержня, мм','options':'сторона проекции, отступы и обработка пропусков','steelClass':'группа материала арматуры'},
 'RebarLabel': {'groupID':'группа общей подписи всех её рядов','target':'линия, по габаритам которой привязывается подпись','normalX':'X-компонента направления подписи','normalY':'Y-компонента направления подписи','axisDistance':'as до осей стержней, мм; отступ подписи задаёт плоттер'},
 'RebarRowLabel': {'groupID':'группа общей подписи всех её рядов','ruleID':'ряд, к концам которого привязана подпись','normalX':'X-компонента направления подписи','normalY':'Y-компонента направления подписи','axisDistance':'as до осей стержней, мм; отступ подписи задаёт плоттер'},
 'RebarRowExtentLabel': {'groupID':'группа общей подписи','ruleID':'ряд для определения крайних стержней','normalX':'X-компонента направления подписи','normalY':'Y-компонента направления подписи','axisDistance':'as до осей стержней, мм; отступ подписи задаёт плоттер','reverseDirection':'True меняет местами концы опорной линии подписи'},
 'RebarRowExtremeLabel': {'groupID':'группа общей подписи','ruleID':'исходный ряд для выбора крайней линии','directionX':'X-компонента направления выбора крайнего ряда','directionY':'Y-компонента направления выбора крайнего ряда','normalX':'X-компонента направления подписи','normalY':'Y-компонента направления подписи','axisDistance':'as до осей стержней, мм; отступ подписи задаёт плоттер'},
 'HorizontalDistance': {'annotationID':'имя размера','name':'обозначение на схеме','startX':'X первой точки, мм','endX':'X второй точки, мм','placementY':'уровень исходной размерной линии Y, мм','normalY':'1 вынос вверх, -1 вниз','comment':'пояснение к размеру','displayIncrement':'шаг округления только подписи, мм'},
 'VerticalExtent': {'annotationID':'имя размера','target':'контур, по которому измеряется высота','name':'обозначение на схеме','placementX':'положение исходной размерной линии X, мм','normalX':'-1 вынос влево, 1 вправо','comment':'пояснение к размеру'},
 'HorizontalExtent': {'annotationID':'имя размера','target':'контур, по которому измеряется ширина','name':'обозначение на схеме','placementY':'положение исходной размерной линии Y, мм','normalY':'1 вынос вверх, -1 вниз','comment':'пояснение к размеру'},
 'Radius': {'annotationID':'имя размера радиуса','pathID':'контур окружности','segmentID':'имя дуги внутри контура','parameter':'относительная точка дуги от 0 до 1','name':'обозначение радиуса','comment':'пояснение к размеру'},
 'Anchor': {'anchorID':'имя точки привязки','x':'X точки, мм','y':'Y точки, мм'},
 'Text': {'annotationID':'имя аннотации','anchorID':'точка размещения текста','textValue':'текст на схеме','comment':'пояснение к аннотации'},
 'Finish': {},
}

PROPERTIES = {
 'StartSetback':'отступ начала ряда от начала грани, мм',
 'EndSetback':'отступ конца ряда от конца грани, мм',
 'EndMode':'учёт концов: ClosedUniform замкнутый ряд; ExcludeBoth без крайних точек',
 'SourcePrefix':'начало имён стержней', 'SourceSuffix':'суффикс имён дополнительного ряда',
 'AnnotationGroup':'имя общей подписи для всех рядов этой группы',
 'Publication':'PublishPerSourcePosition сохраняет имена по исходным позициям',
 'Coincidence':'KeepPublishedOwner сохраняет первого владельца совпавшего стержня',
 'Comment':'пояснение к стержням данного ряда',
 'SourceReference':'адреса ячеек, которые надо исправить при ошибке размещения',
 'PlacementCheck':'CheckFullDisk проверяет нахождение всего стержня в бетоне',
 'OverlapPolicy':'RejectOverlap запрещает пересечение стержней',
 'PositionOrder':'FromPathMiddle нумерует позиции от середины линии',
 'IncomingTangentAtJoint':'True использует входящую касательную на стыке участков',
 'SourceNormalSign':'знак нормали исходного ряда для направления проекции',
 'SkipMissingProjection':'True пропускает позиции без пересечения с целевой линией',
 'SuccessfulProjectionOrdinal':'True нумерует только успешно спроецированные позиции',
 'NewPublicationRoot':'True создаёт отдельные имена стержней внутреннего ряда',
 'EverySecond':'True берёт первый, третий и следующие нечётные стержни; False все',
 'NormalOffset':'знаковое смещение по нормали к исходному ряду, мм',
 'TangentOffset':'знаковое смещение вдоль исходного ряда, мм',
}

HELPERS = {
 'AdditionalOptions': {'table':'снимок таблицы армирования','key':'ключ диаметра дополнительного ряда','suffix':'суффикс имён ряда','location':'Stacked поперёк ряда, SideBySide вдоль','distance':'расстояние между центрами стержней, мм','row':'строка дополнительных рядов','column':'столбец их диаметра','groupID':'общая группа подписи','caption':'название стороны','label':'название стороны','rowNumber':'номер дополнительного ряда','normalSign':'знак смещения по нормали','sideSign':'знак направления на данной стороне'},
 'ExtraOptions': {'face':'параметры данной стороны','rowNumber':'номер дополнительного ряда','location':'Stacked поперёк ряда, SideBySide вдоль','binding':'EachBar все позиции, EverySecondBar через одну','distance':'межосевое расстояние, мм','normalSign':'знак смещения по нормали'},
 'AxisPath': {'facade':'сборщик описания сечения','sizes':'размеры, вершины и радиусы контура','groupID':'имя стороны','axisDistance':'as до центров стержней, мм','sourceReference':'адрес исходной ячейки для диагностики','faces':'параметры всех наружных и внутренних сторон','index':'номер выбранной стороны'},
 'AddCorner': {'points':'массив координат вершин','radii':'массив радиусов','names':'массив имён вершин','count':'число уже заданных вершин; увеличивается на одну','name':'имя добавляемой вершины','x':'X вершины, мм','y':'Y вершины, мм','radius':'радиус скругления, мм'},
 'AddBoundary': {'facade':'сборщик описания сечения','loopID':'имя бетонной границы','shapeID':'имя её исходного контура','B':'ширина, мм','H':'высота, мм','radius':'радиус всех углов, мм','x':'X центра контура, мм','y':'Y центра контура, мм','opening':'True отверстие, False наружная граница'},
 'AddGeometry': {'facade':'сборщик описания сечения','sizes':'подготовленные размеры и радиусы','upper':'возвращаемый контур верхнего прямоугольника','lower':'возвращаемый контур нижнего прямоугольника','outer':'возвращаемая общая наружная граница'},
 'AddReinforcement': {'facade':'сборщик описания сечения','table':'снимок таблицы армирования','units':'перевод исходных размеров в мм','sizes':'размеры и данные формы','faces':'параметры всех наружных и внутренних сторон'},
 'AddDimensions': {'facade':'сборщик описания сечения','sizes':'размеры и данные формы','outer':'наружная граница для габаритов','upper':'верхний прямоугольник','lower':'нижний прямоугольник'},
 'AddRectangleRows': {'facade':'сборщик описания сечения','table':'снимок таблицы армирования','units':'перевод размеров в мм','header':'номер строки шапки таблицы','base':'смещение группы строк: 0 верхний, 4 нижний прямоугольник','x':'X нижнего левого угла, мм','y':'Y нижнего левого угла, мм','width':'ширина прямоугольника, мм','height':'высота прямоугольника, мм'},
 'Describe': {'parameters':'размеры, радиусы, диаметры и число стержней','table':'таблица адресов исходных ячеек; Nothing при программном вводе'},
}

def split_args(text):
    output, start, depth, quoted, i = [], 0, 0, False, 0
    while i < len(text):
        ch = text[i]
        if ch == '"':
            if quoted and i+1 < len(text) and text[i+1] == '"':
                i += 2; continue
            quoted = not quoted
        elif not quoted:
            if ch == '(': depth += 1
            elif ch == ')': depth -= 1
            elif ch == ',' and depth == 0:
                output.append(text[start:i].strip()); start = i+1
        i += 1
    if text[start:].strip(): output.append(text[start:].strip())
    return output

def call_args(code, name, prefix=''):
    match = re.search(re.escape(prefix+name)+r'\b',code)
    assert match, (name,code)
    rest = code[match.end():].strip()
    if not rest: return []
    if rest[0] != '(': return split_args(rest)
    depth, quoted, i = 0, False, 0
    while i < len(rest):
        if rest[i] == '"':
            if quoted and i+1 < len(rest) and rest[i+1] == '"': i += 2; continue
            quoted = not quoted
        elif not quoted:
            if rest[i] == '(': depth += 1
            elif rest[i] == ')':
                depth -= 1
                if depth == 0: return split_args(rest[1:i])
        i += 1
    raise ValueError('Unclosed call: '+code)

def signatures(text):
    logical = re.sub(r'\s+_\r?\n\s*',' ',text)
    result = {}
    for match in re.finditer(r'(?m)^(?:Public|Private) (?:Function|Sub) (\w+)\(',logical):
        result[match[1]] = [re.search(r'(?:ByVal|ByRef)?\s*(\w+)(?:\(\))?\s+As\b',a.replace('Optional ','')).group(1)
                            for a in call_args(logical[match.start():],match[1])]
    return result

def argument_comment(code, owner, builder_signatures, helper_signatures, fallback):
    method = re.search(r'\bfacade\.(\w+)',code)
    label, values, names, meanings = '', [], [], {}
    if method:
        label = method[1]; values = call_args(code,label,'facade.'); names = builder_signatures[label]; meanings = MEANINGS[label]
    else:
        for helper in HELPERS:
            if re.search(r'\b'+helper+r'\s*\(',code) or re.match(r'^(?:If .*? Then )?'+helper+r'\b',code.strip()):
                label=helper; values=call_args(code,helper); names=helper_signatures[label]; meanings=HELPERS[label]; break
    if label:
        assert len(values) <= len(names), (label,values,names)
        args = []
        for p,v in zip(names,values):
            if ':=' in v:
                p,v = v.split(':=',1)
                p = p.strip()
            args.append({'parameter':p,'value':v,'meaning':meanings[p]})
        for a in args:
            v = a['value'].replace('#','').strip()
            if label in ('VerticalExtent','HorizontalExtent','HorizontalDistance'):
                if a['parameter']=='normalX' and v in ('-1','1','1.0','-1.0'):
                    a['meaning']='вынос размера влево' if v.startswith('-') else 'вынос размера вправо'
                if a['parameter']=='normalY' and v in ('-1','1','1.0','-1.0'):
                    a['meaning']='вынос размера вниз' if v.startswith('-') else 'вынос размера вверх'
            if a['parameter']=='forward' and v in ('True','False'):
                a['meaning']='по обходу исходной линии' if v=='True' else 'против обхода исходной линии'
            if a['parameter']=='fullLoopIfSame' and v=='True':
                a['meaning']='полный замкнутый контур при совпавших концах'
            if a['parameter']=='reverseDirection' and v=='True':
                a['meaning']='меняем местами концы опорной линии подписи'
        if not args: return fallback, args
        text = '; '.join((a['parameter'] if a['parameter'] in ('comment','arcComment') or len(a['value'])>55 else a['value'].replace('#',''))+' — '+a['meaning'] for a in args)+'.'
        return text,args
    assignments = re.findall(r'options\.(\w+)\s*=\s*(.*?)(?=:\s*options\.|\s+Else\s+options\.|$)',code)
    if assignments:
        args = [{'parameter':p,'value':v.strip(),'meaning':PROPERTIES[p]} for p,v in assignments]
        for a in args:
            if a['parameter'] == 'EndMode':
                a['meaning'] = 'ClosedUniform распределяет стержни по замкнутой линии без повторного конца' if a['value']=='ClosedUniform' else 'ExcludeBoth исключает точки обоих концов линии'
        return '; '.join(a['parameter']+' — '+a['meaning'] for a in args)+'.',args
    return fallback,[]
