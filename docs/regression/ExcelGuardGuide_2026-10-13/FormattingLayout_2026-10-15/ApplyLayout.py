"""Разовая правка готовой книги: разделитель и оформление ссылки заголовка."""
import copy
import hashlib
import json
import re
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as ET

directory = Path(__file__).resolve().parent
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
CELL = re.compile(r'(\$?)([A-Z]+)(\$?)([0-9]+)')
START, END = 85, 182


def shifted(reference):
    def replace(match):
        a, column, b, row = match.groups()
        if len(column) == 1 and 'H' <= column <= 'P' and START <= int(row) <= END:
            row = str(int(row) + 1)
        return a + column + b + row
    return CELL.sub(replace, reference)


with ZipFile(directory / 'Baseline.xlsm') as archive:
    entries = archive.infolist()
    parts = {entry.filename: archive.read(entry.filename) for entry in entries}
strings = ET.fromstring(parts['xl/sharedStrings.xml'])
config = ET.fromstring(parts['xl/worksheets/sheet1.xml'])
book = ET.fromstring(parts['xl/workbook.xml'])
cells = {cell.get('r'): cell for cell in config.iter(S + 'c')}
rows = {int(row.get('r')): row for row in config.find(S + 'sheetData')}
assert cells['H85'].get('t') == 's'
template = copy.deepcopy(strings[int(cells['H85'].find(S + 'v').text)])
assert ''.join(template.itertext()) == 'Круглое сечение (Подробнее)'
template.find(S + 'r/' + S + 't').text = 'Настройки оформления схемы'
template.tag = S + 'is'
heading = cells['H69']
assert ''.join(heading.itertext()) == 'Настройки оформления схемы (Подробнее)'
heading[:] = [template]
heading.set('t', 'inlineStr')

# Перемещаются только правые таблицы геометрии; A:E и служебные списки остаются на месте.
moved = []
for address, cell in cells.items():
    destination = shifted(address)
    if destination != address:
        moved.append((address, destination, cell))
for address, destination, cell in moved:
    rows[int(re.search(r'\d+', address)[0])].remove(cell)
for address, destination, cell in moved:
    number = int(re.search(r'\d+', destination)[0])
    # После последней таблицы допускается заменить лишь пустую оформленную ячейку.
    previous = next((item for item in rows[number] if item.get('r') == destination), None)
    if previous is not None:
        assert len(previous) == 0, destination
        rows[number].remove(previous)
    cell.set('r', destination)
    formula = cell.find(S + 'f')
    if formula is not None and formula.get('ref'):
        formula.set('ref', shifted(formula.get('ref')))
    rows[number].append(cell)


def column_number(cell):
    result = 0
    for char in re.match('[A-Z]+', cell.get('r'))[0]:
        result = result * 26 + ord(char) - 64
    return result


for row in rows.values():
    row[:] = sorted(row, key=column_number)
for tag, attribute in [('mergeCell', 'ref'), ('dataValidation', 'sqref'), ('hyperlink', 'ref')]:
    for item in config.iter(S + tag):
        item.set(attribute, shifted(item.get(attribute)))
link = next(item for item in config.iter(S + 'hyperlink') if item.get('ref') == 'H69')
assert link.get('location') == "'Справка'!A724"
link.set('display', 'Настройки оформления схемы (Подробнее)')

names_changed = {}
for item in book.find(S + 'definedNames'):
    if item.text.startswith('Config!'):
        value = shifted(item.text)
        if value != item.text:
            names_changed[item.get('name')] = [item.text, value]
            item.text = value
assert set(names_changed) == {
    'rngCircleGeometry', 'rngRectSetGeometry',
    'rngRoundedRectangleGeometry', 'rngHollowRectangleGeometry',
}

# Цепочка пересчёта должна указывать на перенесённые формулы единиц.
chain = ET.fromstring(parts['xl/calcChain.xml'])
sheet_id = None
chain_changed = 0
for item in chain:
    sheet_id = item.get('i', sheet_id)
    if sheet_id == '4':
        old = item.get('r')
        new = shifted(old)
        item.set('r', new)
        chain_changed += old != new
assert chain_changed == sum(cell.find(S + 'f') is not None for _, _, cell in moved)
for part, document in [('xl/worksheets/sheet1.xml', config), ('xl/workbook.xml', book), ('xl/calcChain.xml', chain)]:
    parts[part] = ET.tostring(document, encoding='UTF-8', xml_declaration=True, standalone=True)
candidate = directory / 'Candidate.xlsm'
assert not candidate.exists()
with ZipFile(candidate, 'w') as output:
    for entry in entries:
        output.writestr(entry, parts[entry.filename])
receipt = {
    'Status': 'CONFIG_LAYOUT_APPLIED', 'MovedCells': len(moved),
    'EmptySeparator': 'Config!H85:P85', 'MovedRange': 'Config!H85:P182 -> H86:P183',
    'NamesChanged': names_changed, 'CalculationChainAddressesMoved': chain_changed,
    'LinkTarget': link.get('location'), 'LinkFormattingCopiedFrom': 'Config!H85',
    'WorkbookSHA256': hashlib.sha256(candidate.read_bytes()).hexdigest(),
}
(directory / 'Applied.json').write_text(json.dumps(receipt, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps(receipt, ensure_ascii=False))
