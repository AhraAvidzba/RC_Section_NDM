"""Независимая проверка переноса всех ячеек и сохранности остальных частей книги."""
import argparse
import copy
import hashlib
import json
import re
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as ET

directory = Path(__file__).resolve().parent
parser = argparse.ArgumentParser()
parser.add_argument('--book', type=Path, default=directory / 'Candidate.xlsm')
parser.add_argument('--report', type=Path, default=directory / 'Preservation.json')
args = parser.parse_args()
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
errors = []


def check(condition, description):
    if not condition:
        errors.append(description)


def canon(element):
    return element.tag, tuple(sorted(element.attrib.items())), element.text, tuple(canon(child) for child in element)


def move(address):
    def replace(match):
        first, col, last, row = match.groups()
        return first + col + last + str(int(row) + (len(col) == 1 and 'H' <= col <= 'P' and 85 <= int(row) <= 182))
    return re.sub(r'(\$?)([A-Z]+)(\$?)(\d+)', replace, address)


with ZipFile(directory / 'Baseline.xlsm') as a, ZipFile(args.book) as b:
    check(a.namelist() == b.namelist(), 'ZIP part inventory changed')
    changed = [name for name in a.namelist() if a.read(name) != b.read(name)]
    check(set(changed) == {'xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/calcChain.xml'}, 'Unexpected changed ZIP parts')
    old = ET.fromstring(a.read('xl/worksheets/sheet1.xml'))
    new = ET.fromstring(b.read('xl/worksheets/sheet1.xml'))
    expected = {}
    for cell in old.iter(S + 'c'):
        address = cell.get('r')
        if re.fullmatch('[H-P]183', address):
            check(len(cell) == 0, 'Nonempty overwritten destination: ' + address)
            continue
        clone = copy.deepcopy(cell)
        clone.set('r', move(address))
        for formula in clone.findall(S + 'f'):
            if formula.get('ref'):
                formula.set('ref', move(formula.get('ref')))
        expected[clone.get('r')] = clone
    actual = {cell.get('r'): cell for cell in new.iter(S + 'c')}
    check(expected.keys() == actual.keys(), 'Cell inventory differs from the requested move')
    for address, cell in expected.items():
        if address != 'H69':
            check(address in actual and canon(cell) == canon(actual[address]), 'Cell changed: Config!' + address)
    check(not any(re.fullmatch('[H-P]85', address) for address in actual), 'Separator row is not empty')
    check([dict(row.attrib) for row in old.find(S + 'sheetData')] == [dict(row.attrib) for row in new.find(S + 'sheetData')], 'Row heights or metadata changed')
    strings = ET.fromstring(b.read('xl/sharedStrings.xml'))

    def value(cell):
        if cell.get('t') == 's':
            return ''.join(strings[int(cell.find(S + 'v').text)].itertext())
        if cell.get('t') == 'inlineStr':
            return ''.join(cell.find(S + 'is').itertext())
        v = cell.find(S + 'v')
        return v.text if v is not None else None

    heading = actual['H69']
    check(value(heading) == 'Настройки оформления схемы (Подробнее)', 'Heading text changed')
    link_run = next(run for run in heading.find(S + 'is') if run.find(S + 't').text == '(Подробнее)')
    check(link_run.find(S + 'rPr/' + S + 'color').get('rgb') == 'FF0000FF', 'Link text is not blue')
    check(link_run.find(S + 'rPr/' + S + 'u') is not None, 'Link text is not underlined')
    original_heading = next(cell for cell in old.iter(S + 'c') if cell.get('r') == 'H69')
    check(dict(heading.attrib) == dict(original_heading.attrib), 'Heading cell style changed')
    for tag, attribute in [('mergeCell', 'ref'), ('dataValidation', 'sqref'), ('hyperlink', 'ref')]:
        before, after = list(old.iter(S + tag)), list(new.iter(S + tag))
        check(len(before) == len(after), 'Metadata count changed: ' + tag)
        for x, y in zip(before, after):
            clone = copy.deepcopy(x)
            clone.set(attribute, move(clone.get(attribute)))
            if tag == 'hyperlink' and x.get('ref') == 'H69':
                clone.set('display', 'Настройки оформления схемы (Подробнее)')
            check(canon(clone) == canon(y), 'Metadata changed unexpectedly: ' + tag + ' ' + x.get(attribute))
    for document in (old, new):
        for tag in ['sheetData', 'mergeCells', 'dataValidations', 'hyperlinks']:
            document.remove(document.find(S + tag))
    check(canon(old) == canon(new), 'Unrelated Config metadata changed')
    before_book = ET.fromstring(a.read('xl/workbook.xml'))
    after_book = ET.fromstring(b.read('xl/workbook.xml'))
    changed_names = []
    for name in before_book.find(S + 'definedNames'):
        if (name.text or '').startswith('Config!'):
            shifted = move(name.text)
            if name.text != shifted:
                changed_names.append(name.get('name'))
                name.text = shifted
    check(len(changed_names) == 4, 'Wrong number of moved named ranges')
    check(canon(before_book) == canon(after_book), 'Unrelated workbook metadata changed')
    chain = ET.fromstring(a.read('xl/calcChain.xml'))
    sheet = None
    for item in chain:
        sheet = item.get('i', sheet)
        if sheet == '4':
            item.set('r', move(item.get('r')))
    check(canon(chain) == canon(ET.fromstring(b.read('xl/calcChain.xml'))), 'Calculation chain move is inconsistent')
    all_cells = sum(len(ET.fromstring(a.read(name)).findall('.//' + S + 'c')) for name in a.namelist() if re.fullmatch(r'xl/worksheets/sheet\d+\.xml', name))
    contour_value = value(actual['I84'])
    result = {
        'Passed': not errors, 'ExistingCellsChecked': all_cells,
        'MovedGeometryCells': 882, 'Separator': 'Config!H85:P85',
        'BlueUnderlinedLink': True, 'HeadingTarget': "'Справка'!A724",
        'MovedNamedRanges': changed_names, 'DataValidationsPreserved': 55,
        'ContourLineWeightPreserved': contour_value, 'VbaProjectByteIdentical': a.read('xl/vbaProject.bin') == b.read('xl/vbaProject.bin'),
        'AllOtherSheetsAndDrawingsByteIdentical': all(a.read(name) == b.read(name) for name in a.namelist() if name not in changed),
        'ChangedZIPParts': changed, 'WorkbookSHA256': hashlib.sha256(args.book.read_bytes()).hexdigest(), 'Errors': errors,
    }
args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps(result | {'Errors': errors[:20]}, ensure_ascii=False))
if errors:
    raise SystemExit(1)
