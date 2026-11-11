"""Переносит оформленный лист «Расчет» в новую книгу после штатной сборки.

Шаблон хранит этот лист, его объекты, стили, формулы связи и параметры примера.
Порядок листов, проект VBA и расположение входных таблиц берутся из сборки.
"""

import argparse
from copy import deepcopy
import hashlib
import json
from pathlib import Path
import posixpath
import re
import tempfile
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from lxml import etree as ET


M = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
T = '{http://schemas.openxmlformats.org/package/2006/content-types}'
CALC = 'Расчет'
INPUTS = ('Config', 'Сечения')
SELECTOR = 'rngSectionIllustrationSelector'
STORAGE_DEFAULTS = {'I13', 'B85', *(f'L{row}' for row in range(71, 101))}
EXAMPLE_INPUTS = {
    'Config': ('I106', 'I108', 'I111', 'J111', 'I112', 'J113', 'I115', 'J118'),
    'Сечения': tuple(f'{col}{row}' for rows, columns in ((range(37, 45), 'FG'), (range(50, 58), 'FJ'))
                     for row in rows for col in columns),
}
SAME_SHEET_LINKS = {'I19', 'I20', 'I21', 'I30'}
CALC_REFERENCE = re.compile(r"(?:'Расчет'|Расчет)!")
DIRECT_LINK = re.compile(r"('Расчет'|Расчет)!(\$?)([A-Z]+)(\$?)(\d+)$")
CELL_ADDRESS = re.compile(r'([A-Z]+)(\d+)$')
DRAWING_RELATIONSHIPS = {
    'drawing', 'image', 'chart', 'chartUserShapes', 'chartStyle', 'chartColorStyle',
    'comments', 'vmlDrawing', 'printerSettings',
}


def require(condition, description):
    """Останавливает сборку, если шаблон и актуальная структура разошлись."""
    if not condition:
        raise ValueError(description)


def read_package(path):
    """Читает все части ZIP до изменения; повторяющиеся имена недопустимы."""
    with ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
        require(len(parts) == len(archive.namelist()), f'Повторяющиеся части в {path}')
        return parts


def xml(parts, name):
    """Разбирает XML без обращения к внешним сущностям и ресурсам."""
    return ET.fromstring(parts[name], ET.XMLParser(resolve_entities=False, no_network=True))


def serialize(element):
    """Сохраняет исходные пространства имён Excel и их префиксы."""
    return ET.tostring(element, xml_declaration=True, encoding='UTF-8', standalone=True)


def canonical(element):
    """Сопоставляет XML-стили независимо от наследуемых объявлений пространств имён."""
    return ET.tostring(element, method='c14n', exclusive=True)


def relationship_path(part):
    """Получает имя файла отношений части согласно структуре OPC."""
    folder, name = posixpath.split(part)
    return posixpath.join(folder, '_rels', name + '.rels')


def target_part(owner, target):
    """Нормализует относительный адрес пакета без выхода из каталога xl."""
    path = target.lstrip('/') if target.startswith('/') else posixpath.normpath(
        posixpath.join(posixpath.dirname(owner), target)
    )
    require(path.startswith('xl/'), f'Недопустимый адрес части {target} из {owner}')
    return path


def sheet_parts(parts):
    """Находит листы по книге и её отношениям, а не по номеру XML-файла."""
    workbook = xml(parts, 'xl/workbook.xml')
    rels = xml(parts, 'xl/_rels/workbook.xml.rels')
    targets = {entry.get('Id'): entry.get('Target') for entry in rels}
    return {sheet.get('name'): (index, target_part('xl/workbook.xml', targets[sheet.get(R + 'id')]))
            for index, sheet in enumerate(workbook.find(M + 'sheets'))}


def formula_cells(sheet):
    """Возвращает формулы ячеек; общие прямые ссылки раскрываются относительно ведущей ячейки."""
    cells = {cell.get('r'): cell for cell in sheet.iter(M + 'c')}
    groups = {}
    for address, cell in cells.items():
        formula = cell.find(M + 'f')
        if formula is not None and formula.get('t') == 'shared' and formula.text:
            groups[formula.get('si')] = (address, formula.text)
    result = {}
    for address, cell in cells.items():
        formula = cell.find(M + 'f')
        if formula is None:
            continue
        if formula.text:
            result[address] = formula.text
        elif formula.get('t') == 'shared':
            anchor, expression = groups[formula.get('si')]
            if CALC_REFERENCE.search(expression):
                result[address] = move_direct_link(expression, anchor, address)
    return result


def column_number(letters):
    """Преобразует буквы Excel в номер столбца для раскрытия общей ссылки."""
    number = 0
    for letter in letters:
        number = 26 * number + ord(letter) - ord('A') + 1
    return number


def column_letters(number):
    """Преобразует положительный номер столбца Excel обратно в буквы."""
    require(number > 0, 'Общая формула указывает за пределы листа')
    letters = ''
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(ord('A') + remainder) + letters
    return letters


def move_direct_link(expression, anchor, address):
    """Сдвигает только относительные координаты простой ссылки на «Расчет»."""
    match = DIRECT_LINK.fullmatch(expression)
    require(match is not None, f'Неизвестная общая формула входа {anchor}: {expression}')
    name, abs_col, col, abs_row, row = match.groups()
    from_col, from_row = CELL_ADDRESS.fullmatch(anchor).groups()
    to_col, to_row = CELL_ADDRESS.fullmatch(address).groups()
    if not abs_col:
        col = column_letters(column_number(col) + column_number(to_col) - column_number(from_col))
    if not abs_row:
        row = str(int(row) + int(to_row) - int(from_row))
    require(int(row) > 0, 'Общая формула указывает за пределы листа')
    return f'{name}!{abs_col}{col}{abs_row}{row}'


def set_cell(cell, kind, value):
    """Обновляет только содержимое ячейки, сохраняя стиль и остальные свойства."""
    for tag in ('f', 'v', 'is'):
        previous = cell.find(M + tag)
        if previous is not None:
            cell.remove(previous)
    cell.attrib.pop('t', None)
    if kind == 'formula':
        ET.SubElement(cell, M + 'f').text = value
    elif kind == 'number':
        ET.SubElement(cell, M + 'v').text = value
    elif kind == 'literal':
        cell.set('t', 'inlineStr')
        ET.SubElement(ET.SubElement(cell, M + 'is'), M + 't').text = value
    else:
        require(kind == 'empty' and value is None, f'Неизвестное содержимое ячейки {cell.get("r")}')


def example_input(cell, strings):
    """Сохраняет тип значения параметра примера, в том числе намеренно пустой отступ."""
    kind = cell.get('t', 'n')
    value = cell.findtext(M + 'v')
    require(cell.find(M + 'f') is None and kind in ('s', 'n'),
            f'Ожидалось значение параметра примера в {cell.get("r")}')
    if kind == 's':
        return ['literal', ''.join(strings[int(value)].itertext())]
    return ['empty', None] if value is None else ['number', value]


def prepare_calculation_sheet(sheet, strings):
    """Встраивает подписи листа и очищает кеши формул от сохранённых результатов."""
    stability_reserve_columns = {'AI58': 'AW', 'AJ58': 'AX'}
    for cell in sheet.iter(M + 'c'):
        formula = cell.find(M + 'f')
        value = cell.find(M + 'v')
        if formula is not None:
            reserve_column = stability_reserve_columns.get(cell.get('r'))
            if reserve_column:
                sp35 = f'{reserve_column}42'
                sp63 = f'{reserve_column}56'
                require(formula.text == f'IF($AO$44="СП35",IFERROR({sp35},"-"),IFERROR({sp63},"-"))',
                        f'Изменилась проверенная формула запаса устойчивости Расчет!{cell.get("r")}')
                formula.text = f'IFERROR(ROUND(IF($AO$44="СП35",{sp35},{sp63}),2),"-")'
            if value is not None:
                cell.remove(value)
            cell.attrib.pop('t', None)
        elif cell.get('t') == 's':
            string = strings[int(value.text)]
            cell.remove(value)
            cell.set('t', 'inlineStr')
            inline = ET.SubElement(cell, M + 'is')
            inline.extend(deepcopy(child) for child in string)
    for region in sheet.findall(M + 'conditionalFormatting'):
        if region.get('sqref') == 'S39':
            for rule in region.findall(M + 'cfRule'):
                if rule.findtext(M + 'formula') == '#REF!="да"':
                    region.remove(rule)
            if not len(region):
                sheet.remove(region)


def detach_printer(sheet, relation, path):
    """Отвязывает локальный драйвер, сохраняя поля и параметры страницы."""
    page_setup = sheet.find(M + 'pageSetup')
    require(page_setup is not None and page_setup.get(R + 'id') == relation.get('Id'),
            f'Неизвестная ссылка на принтер в {path}')
    page_setup.attrib.pop(R + 'id')


def collect_objects(parts, root):
    """Берёт граф объектов без расчётного рисунка и настроек драйвера принтера."""
    graph = {}
    ignored = 0

    def collect(owner):
        nonlocal ignored
        if owner in graph:
            return
        graph[owner] = parts[owner]
        rel_path = relationship_path(owner)
        if rel_path not in parts:
            return
        rels = xml(parts, rel_path)
        for relation in list(rels):
            kind = relation.get('Type').rsplit('/', 1)[-1]
            require(kind in DRAWING_RELATIONSHIPS and relation.get('TargetMode') != 'External',
                    f'Неизвестная связь оформления: {kind} в {owner}')
            if kind == 'printerSettings':
                sheet = ET.fromstring(graph[owner])
                detach_printer(sheet, relation, owner)
                graph[owner] = serialize(sheet)
                rels.remove(relation)
                continue
            child = target_part(owner, relation.get('Target'))
            if kind == 'chartUserShapes':
                drawing = xml(parts, child)
                for anchor in list(drawing):
                    properties = [node for node in anchor.iter() if ET.QName(node).localname == 'cNvPr']
                    if properties and all(node.get('name', '').startswith('NDMPlot_') for node in properties):
                        drawing.remove(anchor)
                        ignored += 1
                if not len(drawing):
                    chart = ET.fromstring(graph[owner])
                    reference = next(node for node in chart.iter() if node.get(R + 'id') == relation.get('Id'))
                    reference.getparent().remove(reference)
                    graph[owner] = serialize(chart)
                    rels.remove(relation)
                    continue
                collect(child)
                graph[child] = serialize(drawing)
            else:
                collect(child)
        if len(rels):
            graph[rel_path] = serialize(rels)

    collect(root)
    return graph, ignored


def content_types(parts):
    """Читает индивидуальные типы частей и типы по расширению."""
    types = xml(parts, '[Content_Types].xml')
    overrides = {node.get('PartName').lstrip('/'): node.get('ContentType') for node in types.findall(T + 'Override')}
    defaults = {node.get('Extension'): node.get('ContentType') for node in types.findall(T + 'Default')}
    return types, overrides, defaults


def capture(source, baseline, destination):
    """Создаёт переносимый шаблон из проверенного дооформленного файла."""
    parts = read_package(source)
    base = read_package(baseline)
    locations = sheet_parts(parts)
    base_locations = sheet_parts(base)
    require('xl/theme/theme1.xml' in base and parts['xl/theme/theme1.xml'] == base['xl/theme/theme1.xml'],
            'Оформление и штатная книга используют разные темы')
    calc_index, calc_path = locations[CALC]
    graph, ignored = collect_objects(parts, calc_path)
    calc_sheet = xml(graph, calc_path)
    strings = xml(parts, 'xl/sharedStrings.xml')
    prepare_calculation_sheet(calc_sheet, strings)
    graph[calc_path] = serialize(calc_sheet)
    _, overrides, defaults = content_types(parts)
    types = {part: (overrides[part] if part in overrides else defaults[part.rsplit('.', 1)[-1]])
             for part in graph if not part.endswith('.rels') and part != calc_path}
    names = xml(parts, 'xl/workbook.xml').find(M + 'definedNames')
    names = [{'attributes': dict(node.attrib), 'formula': node.text} for node in names
             if node.get('name') == SELECTOR or node.get('localSheetId') == str(calc_index)]
    links = {}
    for name in INPUTS:
        donor = xml(parts, locations[name][1])
        cells = {cell.get('r'): cell for cell in donor.iter(M + 'c')}
        formulas = formula_cells(donor)
        existing = formula_cells(xml(base, base_locations[name][1]))
        links[name] = {address: ['formula', formula] for address, formula in formulas.items()
                       if existing.get(address) != formula and (
                           CALC_REFERENCE.search(formula) or
                           (name == 'Config' and (address in SAME_SHEET_LINKS or formula == SELECTOR)))}
        if name == 'Config':
            for row in range(71, 101):
                address = f'H{row}'
                link = f'Расчет!{"P" if row <= 85 else "X"}{row - (27 if row <= 85 else 42)}'
                require(links[name].get(address) == ['formula', link],
                        f'Не совпадает проверенная ссылка на ID сочетания Config!{address}')
                links[name][address][1] = f'IF({link}="","",{link})'  # Пустой ID должен остаться пустым в таблице входа.
            for address in sorted(STORAGE_DEFAULTS):
                cell = cells[address]
                require(cell.get('t') == 's' and cell.find(M + 'f') is None,
                        f'Config!{address} должен содержать буквенный выбор')
                value = ''.join(strings[int(cell.findtext(M + 'v'))].itertext())
                links[name][address] = ['literal', value]
        for address in EXAMPLE_INPUTS[name]:
            # Отступы и укладка воспроизводят арматуру примера, а параметры подписей — его схему.
            links[name][address] = example_input(cells[address], strings)
    donor_names = xml(parts, 'xl/workbook.xml').find(M + 'definedNames')
    input_names = {node.get('name'): node.text for node in donor_names
                   if not node.get('localSheetId') and re.match(r"(?:'?(?:Config|Сечения)'?)!", node.text or '')}
    metadata = {
        'schema': 1,
        'sheet': calc_path,
        'themeSHA256': hashlib.sha256(parts['xl/theme/theme1.xml']).hexdigest(),
        'types': types,
        'names': names,
        'inputNames': input_names,
        'inputCells': links,
        'generatedPlotObjectsRemoved': ignored,
    }
    template = {'manifest.json': (json.dumps(metadata, ensure_ascii=False, indent=2) + '\n').encode(),
                'styles.xml': parts['xl/styles.xml'], **graph}
    require(not destination.exists(), f'Шаблон уже существует: {destination}')
    write_zip(destination, template)
    print(f'Шаблон «Расчет»: {len(links["Config"])} входов Config, {len(links["Сечения"])} входов «Сечения», '
          f'{ignored} объектов расчётного снимка исключено')


def merge_styles(base, donor, sheet):
    """Подключает только стили листа «Расчет», оставляя существующие индексы неизменными."""
    sections = ('numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'dxfs')
    indexes = {}
    for tag in sections:
        section = base.find(M + tag)
        if section is None:
            section = ET.Element(M + tag, count='0')
            position = {'numFmts': 0, 'dxfs': base.index(base.find(M + 'cellStyles')) + 1}.get(tag)
            require(position is not None, f'В новой книге отсутствует раздел стилей {tag}')
            base.insert(position, section)
        indexes[tag] = {canonical(item): index for index, item in enumerate(section)}

    number_formats = {item.get('formatCode'): int(item.get('numFmtId')) for item in base.find(M + 'numFmts')}
    donor_formats = {int(item.get('numFmtId')): item.get('formatCode') for item in donor.find(M + 'numFmts')}
    maps = {tag: {} for tag in sections}

    def number_format(index):
        if index < 164:
            return index
        code = donor_formats[index]
        if code not in number_formats:
            new_index = max([163, *number_formats.values()]) + 1
            section = base.find(M + 'numFmts')
            ET.SubElement(section, M + 'numFmt', numFmtId=str(new_index), formatCode=code)
            section.set('count', str(len(section)))
            number_formats[code] = new_index
        return number_formats[code]

    def style(tag, index):
        if index not in maps[tag]:
            item = deepcopy(donor.find(M + tag)[index])
            if tag in ('cellStyleXfs', 'cellXfs'):
                for attribute, section in (('fontId', 'fonts'), ('fillId', 'fills'),
                                           ('borderId', 'borders'), ('xfId', 'cellStyleXfs')):
                    if item.get(attribute) is not None:
                        item.set(attribute, str(style(section, int(item.get(attribute)))))
                if item.get('numFmtId') is not None:
                    item.set('numFmtId', str(number_format(int(item.get('numFmtId')))))
            key = canonical(item)
            if key not in indexes[tag]:
                section = base.find(M + tag)
                indexes[tag][key] = len(section)
                section.append(item)
                section.set('count', str(len(section)))
            maps[tag][index] = indexes[tag][key]
        return maps[tag][index]

    for node in sheet.iter():
        tag = ET.QName(node).localname
        if tag in ('c', 'row') and node.get('s') is not None:
            node.set('s', str(style('cellXfs', int(node.get('s')))))
        elif tag == 'col' and node.get('style') is not None:
            node.set('style', str(style('cellXfs', int(node.get('style')))))
        if node.get('dxfId') is not None:
            node.set('dxfId', str(style('dxfs', int(node.get('dxfId')))))


def copy_objects(parts, template, metadata, base_path, types):
    """Размещает объекты листа под собственными именами, сохраняя их связи и оформление."""
    source_path = metadata['sheet']
    destinations = {source_path: base_path}
    for part, content_type in metadata['types'].items():
        folder, name = posixpath.split(part)
        dest = posixpath.join(folder, 'ndmCalc_' + name)
        require(dest not in parts, f'Часть оформления уже присутствует: {dest}')
        destinations[part] = dest
        parts[dest] = template[part]
        ET.SubElement(types, T + 'Override', PartName='/' + dest, ContentType=content_type)
    for origin, destination in destinations.items():
        old_rels = relationship_path(origin)
        if old_rels not in template:
            continue
        relationships = xml(template, old_rels)
        for relation in relationships:
            child = target_part(origin, relation.get('Target'))
            relation.set('Target', posixpath.relpath(destinations[child], posixpath.dirname(destination)))
        parts[relationship_path(destination)] = serialize(relationships)


def write_zip(destination, parts):
    """Записывает пакет с устойчивым порядком и временем записей."""
    with ZipFile(destination, 'w', ZIP_DEFLATED) as archive:
        for name, payload in parts.items():
            entry = ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            entry.compress_type = ZIP_DEFLATED
            archive.writestr(entry, payload)


def write_verified_package(destination, parts):
    """Проверяет временный пакет и заменяет книгу только после успешной записи."""
    destination.parent.mkdir(parents=True, exist_ok=True)
    # Файл создаётся рядом с книгой и наследует права каталога публикации.
    with tempfile.NamedTemporaryFile(prefix='ndm-calc-', suffix='.tmp', dir=destination.parent, delete=False) as temporary:
        staged = Path(temporary.name)
    try:
        write_zip(staged, parts)
        require(read_package(staged) == parts, 'Проверка записанной книги не прошла')
        staged.replace(destination)
    finally:
        staged.unlink(missing_ok=True)


def make_portable(source):
    """Удаляет созданные Excel бинарные привязки к очередям принтеров из новой книги."""
    parts = read_package(source)
    types, _, _ = content_types(parts)
    removed = set()
    for _, path in sheet_parts(parts).values():
        rel_path = relationship_path(path)
        if rel_path not in parts:
            continue
        relationships = xml(parts, rel_path)
        printers = [item for item in relationships if item.get('Type').endswith('/printerSettings')]
        if not printers:
            continue
        sheet = xml(parts, path)
        for printer in printers:
            detach_printer(sheet, printer, path)
            removed.add(target_part(path, printer.get('Target')))
            relationships.remove(printer)
        parts[path] = serialize(sheet)
        if len(relationships):
            parts[rel_path] = serialize(relationships)
        else:
            parts.pop(rel_path)
    if removed:
        for path in removed:
            parts.pop(path)
        for node in types.findall(T + 'Override'):
            if node.get('PartName').lstrip('/') in removed:
                types.remove(node)
        parts['[Content_Types].xml'] = serialize(types)
        write_verified_package(source, parts)


def apply(source, template_path, destination):
    """Собирает оформленный вариант из новой книги и сохранённого листа."""
    require(source.resolve() != destination.resolve(), 'Обычная и оформленная книги должны быть разными файлами')
    parts = read_package(source)
    template = read_package(template_path)
    metadata = json.loads(template['manifest.json'])
    require(metadata['schema'] == 1, 'Версия шаблона оформления не поддерживается')
    require(hashlib.sha256(parts['xl/theme/theme1.xml']).hexdigest() == metadata['themeSHA256'],
            'Штатная сборка использует другую тему, оформление «Расчет» не применено')
    locations = sheet_parts(parts)
    calc_index, calc_path = locations[CALC]
    workbook = xml(parts, 'xl/workbook.xml')
    names = workbook.find(M + 'definedNames')
    input_names = {node.get('name'): node.text for node in names if not node.get('localSheetId')}
    for name, reference in metadata['inputNames'].items():
        if name == 'rngSystemSettings' and reference == 'Config!$A$3:$E$152':
            # Полный реестр включает строку масштаба стержней ниже всех связанных ячеек.
            reference = 'Config!$A$3:$E$153'
        require(input_names.get(name) == reference, f'Изменилось расположение входного диапазона {name}')
    for name in INPUTS:
        part = locations[name][1]
        sheet = xml(parts, part)
        cells = {cell.get('r'): cell for cell in sheet.iter(M + 'c')}
        for address, (kind, value) in metadata['inputCells'][name].items():
            require(address in cells, f'В новой книге отсутствует {name}!{address}')
            set_cell(cells[address], kind, value)
        parts[part] = serialize(sheet)
    calc = xml(template, metadata['sheet'])
    require(calc.find(M + 'sheetPr').get('codeName') == xml(parts, calc_path).find(M + 'sheetPr').get('codeName'),
            'Изменилось внутреннее имя VBA листа «Расчет»')
    styles = xml(parts, 'xl/styles.xml')
    merge_styles(styles, xml(template, 'styles.xml'), calc)
    parts['xl/styles.xml'] = serialize(styles)
    parts[calc_path] = serialize(calc)
    for node in list(names):
        if node.get('name') == SELECTOR or node.get('localSheetId') == str(calc_index):
            names.remove(node)
    for record in metadata['names']:
        attrs = dict(record['attributes'])
        if 'localSheetId' in attrs:
            attrs['localSheetId'] = str(calc_index)
        ET.SubElement(names, M + 'definedName', attrib=attrs).text = record['formula']
    workbook.find(M + 'calcPr').set('fullCalcOnLoad', '1')
    parts['xl/workbook.xml'] = serialize(workbook)
    types, _, _ = content_types(parts)
    copy_objects(parts, template, metadata, calc_path, types)
    rel_name = 'xl/_rels/workbook.xml.rels'
    rels = xml(parts, rel_name)
    for relation in list(rels):
        if relation.get('Type').endswith('/calcChain'):
            chain = target_part('xl/workbook.xml', relation.get('Target'))
            parts.pop(chain)
            rels.remove(relation)
            for node in types.findall(T + 'Override'):
                if node.get('PartName') == '/' + chain:
                    types.remove(node)
    parts[rel_name] = serialize(rels)
    parts['[Content_Types].xml'] = serialize(types)
    write_verified_package(destination, parts)
    print(f'Workbook built: {destination}')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest='command', required=True)
    for command, params in (('capture', ('source', 'baseline', 'destination')),
                            ('make-portable', ('source',)),
                            ('apply', ('source', 'template', 'destination'))):
        subparser = commands.add_parser(command)
        for param in params:
            subparser.add_argument(param, type=Path)
    args = parser.parse_args()
    if args.command == 'capture':
        capture(args.source, args.baseline, args.destination)
    elif args.command == 'make-portable':
        make_portable(args.source)
    else:
        apply(args.source, args.template, args.destination)
