"""Переносит утверждённое оформление в книгу после штатной сборки.

Шаблон хранит лист «Расчет», его объекты и формулы связи, параметры примера
и оформление входных листов без пользовательских значений. Проект VBA,
настройки и проверки данных создаёт штатный сборщик.
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
XDR = '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}'
CALC = 'Расчет'
INPUTS = ('Config', 'Сечения')
RESULT_LOAD_HEADER = 'AL8:AN9'  # Постоянная объединённая шапка рабочих нагрузок; строки результатов в шаблон не входят.
SELECTOR = 'rngSectionIllustrationSelector'
STORAGE_DEFAULTS = {'I13', *(f'L{row}' for row in range(71, 101))}
EXAMPLE_INPUTS = {
    'Сечения': tuple(f'{col}{row}' for rows, columns in ((range(37, 45), 'FG'), (range(50, 58), 'FJ'))
                     for row in rows for col in columns),
}
ANNOTATION_EXAMPLE_FIELDS = (('Offset', 1), ('LineWeight', 1), ('TextHeight', 1),
                             ('TextHeight', 2), ('TextGap', 1), ('ArrowType', 2),
                             ('Color', 1), ('ShowSymbol', 2))
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
    """Сохраняет формулы утверждённого листа, встраивает подписи и очищает их кеши."""
    for cell in sheet.iter(M + 'c'):
        formula = cell.find(M + 'f')
        value = cell.find(M + 'v')
        if formula is not None:
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


def example_addresses(name, cells, strings, names):
    """Находит параметры подписей по именованной таблице и ключам, без старых адресов Config."""
    if name != 'Config':
        return EXAMPLE_INPUTS[name]
    reference = next(node.text for node in names if node.get('name') == 'rngPlotAnnotationSettings')
    match = re.fullmatch(r"(?:'Config'|Config)!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", reference)
    require(match is not None, 'Не определено положение таблицы подписей схемы')
    column, first, _, last = match.groups()
    rows = {example_input(cells[f'{column}{row}'], strings)[1]: row
            for row in range(int(first) + 1, int(last) + 1)}
    require(all(key in rows for key, _ in ANNOTATION_EXAMPLE_FIELDS), 'Не найдены параметры подписей примера')
    return tuple(f'{column_letters(column_number(column) + offset)}{rows[key]}'
                 for key, offset in ANNOTATION_EXAMPLE_FIELDS)


def capture_input_formatting(parts, path):
    """Сохраняет оформление входного листа без значений, формул и пользовательских комментариев."""
    source = xml(parts, path)
    result = ET.Element(M + 'worksheet', nsmap=source.nsmap)
    columns = source.find(M + 'cols')
    if columns is not None:
        result.append(deepcopy(columns))
    data = ET.SubElement(result, M + 'sheetData')
    for original in source.find(M + 'sheetData'):
        row = ET.SubElement(data, M + 'row', attrib=dict(original.attrib))
        row.attrib.pop('spans', None)
        for cell in original:
            if cell.tag == M + 'c':
                ET.SubElement(row, M + 'c', r=cell.get('r'), s=cell.get('s', '0'))
    merges = source.find(M + 'mergeCells')
    if merges is not None:
        result.append(deepcopy(merges))
    return serialize(result)


def apply_input_formatting(sheet, formatting):
    """Применяет стили и объединения; содержимое ячеек и проверки данных задаёт сборщик."""
    data = sheet.find(M + 'sheetData')
    rows = {int(row.get('r')): row for row in data}
    for original in formatting.find(M + 'sheetData'):
        number = int(original.get('r'))
        if number not in rows:
            row = ET.Element(M + 'row', r=str(number))
            next_row = next((n for n in data if int(n.get('r')) > number), None)
            data.insert(data.index(next_row) if next_row is not None else len(data), row)
            rows[number] = row
        row = rows[number]
        row.attrib.clear()
        row.attrib.update(original.attrib)
        cells = {cell.get('r'): cell for cell in row}
        for source_cell in original:
            address = source_cell.get('r')
            if address not in cells:
                cell = ET.Element(M + 'c', r=address)
                column = column_number(CELL_ADDRESS.fullmatch(address)[1])
                following = next((c for c in row if column_number(CELL_ADDRESS.fullmatch(c.get('r'))[1]) > column), None)
                row.insert(row.index(following) if following is not None else len(row), cell)
                cells[address] = cell
            cells[address].set('s', source_cell.get('s'))
    for tag in ('cols', 'mergeCells'):
        old, new = sheet.find(M + tag), formatting.find(M + tag)
        if old is not None:
            if new is None:
                sheet.remove(old)
            else:
                sheet.replace(old, deepcopy(new))
        elif new is not None:
            # В SpreadsheetML столбцы стоят перед данными, объединения — перед печатными свойствами.
            boundary = sheet.find(M + 'sheetData') if tag == 'cols' else next(
                (n for n in sheet if ET.QName(n).localname in ('phoneticPr', 'conditionalFormatting',
                 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup',
                 'headerFooter', 'drawing', 'legacyDrawing', 'tableParts', 'extLst')), None)
            sheet.insert(sheet.index(boundary) if boundary is not None else len(sheet), deepcopy(new))
    addresses = [CELL_ADDRESS.fullmatch(c.get('r')).groups() for c in data.iter(M + 'c')]
    if addresses:
        last_column = max(column_number(column) for column, _ in addresses)
        last_row = max(int(row) for _, row in addresses)
        sheet.find(M + 'dimension').set('ref', f'A1:{column_letters(last_column)}{last_row}')


def user_input_fill_cells(sheet, styles):
    """Сохраняет назначенные сборщиком поля ввода до наложения оформления шаблона."""
    fills, formats = styles.find(M + 'fills'), styles.find(M + 'cellXfs')
    input_fills = {i for i, fill in enumerate(fills)
                   if fill.find(M + 'patternFill/' + M + 'fgColor') is not None
                   and fill.find(M + 'patternFill/' + M + 'fgColor').get('rgb', '').upper() == 'FFFFF2E6'}
    return {cell.get('r') for cell in sheet.iter(M + 'c')
            if int(formats[int(cell.get('s', '0'))].get('fillId', '0')) in input_fills}


def apply_user_input_fills(sheet, styles, addresses):
    """Меняет только заливку полей ввода; формулы и остальные свойства оформления сохраняются."""
    fills, formats = styles.find(M + 'fills'), styles.find(M + 'cellXfs')
    fill_id = next(i for i, fill in enumerate(fills)
                   if fill.find(M + 'patternFill/' + M + 'fgColor') is not None
                   and fill.find(M + 'patternFill/' + M + 'fgColor').get('rgb', '').upper() == 'FFFFF2E6')
    replacements = {}
    for cell in sheet.iter(M + 'c'):
        if cell.get('r') not in addresses:
            continue
        original = int(cell.get('s', '0'))
        if int(formats[original].get('fillId', '0')) == fill_id:
            continue
        if original not in replacements:
            value = deepcopy(formats[original])
            value.set('fillId', str(fill_id))
            value.set('applyFill', '1')
            replacements[original] = len(formats)
            formats.append(value)
        cell.set('s', str(replacements[original]))
    formats.set('count', str(len(formats)))


def capture_result_header(parts, path):
    """Сохраняет только постоянную шапку нагрузок Results, без результатов и метаданных запуска."""
    original = xml(parts, path)
    result = ET.Element(M + 'worksheet', nsmap=original.nsmap)
    data = ET.SubElement(result, M + 'sheetData')
    for row in original.find(M + 'sheetData'):
        selected = [cell for cell in row if cell.get('r') in {f'{col}{number}' for col in ('AL', 'AM', 'AN') for number in (8, 9)}]
        if selected:
            target = ET.SubElement(data, M + 'row', r=row.get('r'))
            target.extend(deepcopy(cell) for cell in selected)
    merges = ET.SubElement(result, M + 'mergeCells', count='1')
    require(any(node.get('ref') == RESULT_LOAD_HEADER for node in original.find(M + 'mergeCells')),
            'Не найдена объединённая шапка рабочих нагрузок Results')
    ET.SubElement(merges, M + 'mergeCell', ref=RESULT_LOAD_HEADER)
    require(not any(cell.find(M + 'f') is not None for cell in result.iter(M + 'c')), 'В постоянной шапке Results обнаружена формула')
    prepare_calculation_sheet(result, xml(parts, 'xl/sharedStrings.xml'))
    return serialize(result)


def apply_result_header(sheet, header):
    """Переносит шесть ячеек шапки и её объединение; другие области Results остаются штатными."""
    data = sheet.find(M + 'sheetData')
    rows = {int(row.get('r')): row for row in data}
    for original in header.find(M + 'sheetData'):
        number = int(original.get('r'))
        if number not in rows:
            rows[number] = ET.Element(M + 'row', r=str(number))
            following = next((node for node in data if int(node.get('r')) > number), None)
            data.insert(data.index(following) if following is not None else len(data), rows[number])
        row = rows[number]
        for cell in original:
            existing = next((node for node in row if node.get('r') == cell.get('r')), None)
            if existing is not None:
                row.remove(existing)
            column = column_number(CELL_ADDRESS.fullmatch(cell.get('r'))[1])
            following = next((node for node in row if column_number(CELL_ADDRESS.fullmatch(node.get('r'))[1]) > column), None)
            row.insert(row.index(following) if following is not None else len(row), deepcopy(cell))
    merges = sheet.find(M + 'mergeCells')
    if merges is None:
        merges = ET.Element(M + 'mergeCells')
        sheet.insert(sheet.index(data) + 1, merges)
    if not any(node.get('ref') == RESULT_LOAD_HEADER for node in merges):
        merges.append(deepcopy(header.find(M + 'mergeCells')[0]))
    merges.set('count', str(len(merges)))


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
                wrapped = f'IF({link}="","",{link})'
                require(links[name].get(address) in (['formula', link], ['formula', wrapped]),
                        f'Не совпадает проверенная ссылка на ID сочетания Config!{address}')
                links[name][address][1] = wrapped  # Пустой ID должен остаться пустым в таблице входа.
            # Положение настройки меняется вместе с каталогом; адрес B не является контрактом.
            crack_keys = [cell.get('r') for cell in cells.values()
                          if cell.get('t') == 's' and cell.find(M + 'f') is None
                          and ''.join(strings[int(cell.findtext(M + 'v'))].itertext())
                          == 'SLS.Crack.Code']
            require(len(crack_keys) == 1, 'Не найдена единственная настройка SLS.Crack.Code')
            crack_value = 'B' + CELL_ADDRESS.fullmatch(crack_keys[0]).group(2)
            for address in sorted(STORAGE_DEFAULTS | {crack_value}):
                cell = cells[address]
                require(cell.get('t') == 's' and cell.find(M + 'f') is None,
                        f'Config!{address} должен содержать буквенный выбор')
                value = ''.join(strings[int(cell.findtext(M + 'v'))].itertext())
                links[name][address] = ['literal', 'SP35' if address == crack_value else value]  # Штатная методика новой сборки — СП35; в книге-образце может быть выбран СП63.
        for address in example_addresses(name, cells, strings, xml(parts, 'xl/workbook.xml').find(M + 'definedNames')):
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
        'inputFormatting': {name: f'input-formatting/{name}.xml' for name in INPUTS},
        'resultHeader': 'result-load-header.xml',
        'generatedPlotObjectsRemoved': ignored,
    }
    template = {'manifest.json': (json.dumps(metadata, ensure_ascii=False, indent=2) + '\n').encode(),
                'styles.xml': parts['xl/styles.xml'], **graph}
    for name, path in metadata['inputFormatting'].items():
        template[path] = capture_input_formatting(parts, locations[name][1])
    template[metadata['resultHeader']] = capture_result_header(parts, locations['Results'][1])
    require(not destination.exists(), f'Шаблон уже существует: {destination}')
    write_zip(destination, template)
    print(f'Шаблон «Расчет»: {len(links["Config"])} входов Config, {len(links["Сечения"])} входов «Сечения», '
          f'{ignored} объектов расчётного снимка исключено')


def merge_styles(base, donor, sheet):
    """Подключает используемые стили оформляемого листа, сохраняя существующие индексы."""
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

    # Кнопки остаются на своих местах внутри оформления, но не попадают в печать.
    relationships = xml(parts, relationship_path(base_path))
    drawing, = [relation for relation in relationships if relation.get('Type').endswith('/drawing')]
    drawing_path = target_part(base_path, drawing.get('Target'))
    drawings = xml(parts, drawing_path)
    for shape_anchor in drawings:
        shape = shape_anchor.find(XDR + 'sp')
        if shape is not None and shape.get('macro'):
            shape_anchor.find(XDR + 'clientData').set('fPrintsWithSheet', '0')
    parts[drawing_path] = serialize(drawings)


def refresh_section_picture(parts, sheet_path):
    """Берёт стартовый PNG из свежих иллюстраций, сохраняя оформленную рамку CurrentSection."""
    def picture_media(owner, name):
        relationships = xml(parts, relationship_path(owner))
        drawing, = [relation for relation in relationships if relation.get('Type').endswith('/drawing')]
        drawing_path = target_part(owner, drawing.get('Target'))
        anchors = xml(parts, drawing_path)
        picture, = [node for node in anchors.iter() if ET.QName(node).localname == 'pic' and
                    any(child.get('name') == name for child in node.iter())]
        properties = next(node for node in picture.iter() if ET.QName(node).localname == 'cNvPr')
        blip = next(node for node in picture.iter() if ET.QName(node).localname == 'blip')
        relationship, = [node for node in xml(parts, relationship_path(drawing_path))
                         if node.get('Id') == blip.get(R + 'embed') and node.get('Type').endswith('/image')]
        return properties, target_part(drawing_path, relationship.get('Target'))

    picture, target = picture_media(sheet_path, 'CurrentSection')
    require(bool(picture.get('descr')), 'Для стартовой иллюстрации не указан тип сечения')
    source_sheet = sheet_parts(parts)['Сечения'][1]
    _, source = picture_media(source_sheet, picture.get('descr'))
    require(source.endswith('.png') and target.endswith('.png'), 'Стартовая иллюстрация должна быть PNG')
    parts[target] = parts[source]


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
        require(input_names.get(name) == reference, f'Изменилось расположение входного диапазона {name}')
    styles = xml(parts, 'xl/styles.xml')
    editable_fills = {name: user_input_fill_cells(xml(parts, locations[name][1]), styles) for name in INPUTS}
    donor_styles = xml(template, 'styles.xml')
    if metadata.get('resultHeader'):
        header = xml(template, metadata['resultHeader'])
        merge_styles(styles, donor_styles, header)
        result_path = locations['Results'][1]
        results = xml(parts, result_path)
        apply_result_header(results, header)
        parts[result_path] = serialize(results)
    for name in INPUTS:
        part = locations[name][1]
        sheet = xml(parts, part)
        cells = {cell.get('r'): cell for cell in sheet.iter(M + 'c')}
        for address, (kind, value) in metadata['inputCells'][name].items():
            require(address in cells, f'В новой книге отсутствует {name}!{address}')
            set_cell(cells[address], kind, value)
        if name in metadata.get('inputFormatting', {}):
            formatting = xml(template, metadata['inputFormatting'][name])
            merge_styles(styles, donor_styles, formatting)
            apply_input_formatting(sheet, formatting)
        apply_user_input_fills(sheet, styles, editable_fills[name])
        parts[part] = serialize(sheet)
    calc = xml(template, metadata['sheet'])
    require(calc.find(M + 'sheetPr').get('codeName') == xml(parts, calc_path).find(M + 'sheetPr').get('codeName'),
            'Изменилось внутреннее имя VBA листа «Расчет»')
    merge_styles(styles, donor_styles, calc)
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
    refresh_section_picture(parts, calc_path)
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
