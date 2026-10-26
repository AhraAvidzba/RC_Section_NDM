"""Сверяет пользовательские данные/форматирование ZIP/XML без открытия Excel и без записи книг."""
from __future__ import annotations
import argparse
import hashlib
import importlib.util
import json
import posixpath
import re
import zipfile
from decimal import Decimal
from pathlib import Path
from xml.etree import ElementTree as E

ROOT = Path(__file__).resolve().parents[3]
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
SPEC = importlib.util.spec_from_file_location('preservation_reader', ROOT / 'tools/build_workbook/Verify-SectionsSheetPreservation.py')
READER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(READER)


def digest(data):
    return hashlib.sha256(data).hexdigest().upper()


def canon(node):
    """Нормализует XML-компонент стиля независимо от порядка атрибутов и ID."""
    if node is None:
        return None
    return [node.tag, sorted(node.attrib.items()), node.text or '', [canon(c) for c in node]]


def col_name(number):
    text = ''
    while number:
        number, last = divmod(number - 1, 26)
        text = chr(65 + last) + text
    return text


class Book:
    """Читает исходные части и семантические ячейки существующим независимым reader."""
    def __init__(self, path):
        self.path = Path(path)
        self.names, self.sheets, self.calc = READER.read(self.path)
        with zipfile.ZipFile(self.path) as z:
            bad = z.testzip()
            if bad:
                raise ValueError('Broken ZIP member: ' + bad)
            self.parts = {n: z.read(n) for n in z.namelist()}
        styles = E.fromstring(self.parts['xl/styles.xml'])
        fonts = list(styles.find('s:fonts', NS))
        fills = list(styles.find('s:fills', NS))
        borders = list(styles.find('s:borders', NS))
        formats = {int(n.get('numFmtId')): n.get('formatCode') for n in styles.findall('s:numFmts/s:numFmt', NS)}
        bases = list(styles.find('s:cellStyleXfs', NS))

        def xf(node, parent=True):
            ids = ('fontId', 'fillId', 'borderId', 'numFmtId', 'xfId')
            n = int(node.get('numFmtId', '0'))
            result = {'attributes': sorted((k, v) for k, v in node.attrib.items() if k not in ids),
                      'font': canon(fonts[int(node.get('fontId', '0'))]),
                      'fill': canon(fills[int(node.get('fillId', '0'))]),
                      'border': canon(borders[int(node.get('borderId', '0'))]),
                      'numberFormat': formats.get(n, 'BUILTIN:' + str(n)),
                      'children': [canon(c) for c in node]}
            if parent:
                result['base'] = xf(bases[int(node.get('xfId', '0'))], False)
            return result

        self.style = [digest(json.dumps(xf(n), ensure_ascii=False, sort_keys=True).encode())
                      for n in styles.find('s:cellXfs', NS)]
        workbook = E.fromstring(self.parts['xl/workbook.xml'])
        rels = {n.get('Id'): n.get('Target') for n in E.fromstring(self.parts['xl/_rels/workbook.xml.rels'])}
        for item in workbook.find('s:sheets', NS):
            name = item.get('name')
            target = rels[item.get(RID)]
            part = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
            xml = E.fromstring(self.parts[part])
            sheet = self.sheets[name]
            sheet['part'] = part
            sheet['rawcells'] = {n.get('r'): n for n in xml.findall('s:sheetData/s:row/s:c', NS)}
            sheet_format = xml.find('s:sheetFormatPr', NS)
            sheet['defaultRowHeight'] = sheet_format.get('defaultRowHeight', '15') if sheet_format is not None else '15'
            sheet['rowattrs'] = {int(n.get('r')): {k: v for k, v in n.attrib.items() if k in ('ht', 'hidden', 'customHeight', 'outlineLevel')}
                                 for n in xml.findall('s:sheetData/s:row', NS)}
            sheet['merges'] = [n.get('ref') for n in xml.findall('s:mergeCells/s:mergeCell', NS)]

    def style_at(self, sheet_name, ref):
        sheet = self.sheets[sheet_name]
        node = sheet['rawcells'].get(ref)
        return self.style[int(node.get('s', '0'))] if node is not None else self.style[0]

    def cell(self, sheet, ref):
        return self.sheets[sheet]['cells'].get(ref), self.style_at(sheet, ref)


def setting_rows(book):
    """Ключи Config связываются с сохранёнными значениями, а не с прежними адресами."""
    ref = book.names['rngSystemSettings']
    points = re.findall(r'\$([A-Z]+)\$(\d+)', ref)
    first, last = int(points[0][1]), int(points[-1][1])
    rows = {}
    for row in range(first, last + 1):
        value = book.sheets['Config']['cells'].get('A' + str(row))
        if value and value[0] == 'value' and value[1] and not value[1].startswith('['):
            if value[1] in rows:
                raise ValueError('Duplicate setting key: ' + value[1])
            rows[value[1]] = row
    return rows


def verify(before, after):
    old, new = Book(before), Book(after)
    old_keys, new_keys = setting_rows(old), setting_rows(new)
    insertion = old_keys['Mesh.StepMode']
    shift = new_keys['Mesh.StepMode'] - insertion
    errors, allowed, counts = [], [], {}
    if shift not in (0, 1):
        errors.append({'Kind': 'UnexpectedPartialRowShift', 'Value': shift})
    if list(old.sheets) != list(new.sheets):
        errors.append({'Kind': 'SheetOrderOrNames', 'Before': list(old.sheets), 'After': list(new.sheets)})
    if old.calc != new.calc:
        errors.append({'Kind': 'CalculationMode', 'Before': old.calc, 'After': new.calc})

    def mapped(sheet, ref):
        row, col = READER.coord(ref.replace('$', ''))
        if sheet == 'Config' and col <= 5 and row >= insertion:
            row += shift
        return col_name(col) + str(row)

    def excluded(sheet, ref):
        row, col = READER.coord(ref)
        return ((sheet == 'Results' and 155 <= row <= 156 and 96 <= col <= 113)
                or (sheet == 'Config' and shift == 1 and row == insertion and col <= 5))

    descriptions = {'Mesh.CellGeometry', 'Mesh.StepMode', 'Mesh.BoundarySubdivisions'}
    description_cells = {'D' + str(new_keys[k]): k for k in descriptions if k in new_keys}
    protected = [n for n in old.sheets if n != 'Справка']
    for name in protected:
        original = {mapped(name, ref): ref for ref in old.sheets[name]['rawcells']}
        refs = set(original) | set(new.sheets[name]['rawcells'])
        counts[name] = {'Cells': 0, 'FormulaCells': 0, 'Styles': 0}
        for target in sorted(refs, key=READER.coord):
            if excluded(name, target):
                continue
            source = original.get(target, target)
            a, b = old.cell(name, source), new.cell(name, target)
            # У ещё отсутствовавшей ячейки новой координаты нет старого содержимого.
            if target not in original:
                a = (None, old.style[0])
            if a[0] != b[0]:
                if name == 'Config' and target in description_cells and a[0] and b[0] and a[0][0] == b[0][0] == 'value':
                    allowed.append({'Kind': 'CatalogDescription', 'Key': description_cells[target], 'BeforeCell': source, 'AfterCell': target, 'Before': a[0][1], 'After': b[0][1]})
                else:
                    errors.append({'Kind': 'CellDataOrFormula', 'Sheet': name, 'BeforeCell': source, 'AfterCell': target, 'Before': a[0], 'After': b[0]})
            if a[1] != b[1]:
                errors.append({'Kind': 'NormalizedCellStyle', 'Sheet': name, 'BeforeCell': source, 'AfterCell': target, 'Before': a[1], 'After': b[1]})
            counts[name]['Cells'] += 1
            counts[name]['Styles'] += 1
            counts[name]['FormulaCells'] += int(bool(a[0] and a[0][0] == 'formula'))
        # Частичная вставка A:E не перемещает физические высоты строк и H+.
        for row in set(old.sheets[name]['rowattrs']) | set(new.sheets[name]['rowattrs']):
            a = old.sheets[name]['rowattrs'].get(row, {})
            b = new.sheets[name]['rowattrs'].get(row, {})
            def height(book, attrs):
                return {'height': str(Decimal(attrs.get('ht', book.sheets[name]['defaultRowHeight'])).normalize()),
                        'hidden': attrs.get('hidden', '0'), 'outlineLevel': attrs.get('outlineLevel', '0')}
            if height(old, a) != height(new, b):
                errors.append({'Kind': 'RowDimensions', 'Sheet': name, 'Row': row, 'Before': height(old, a), 'After': height(new, b)})
        # Массивы отдельных диапазонов cols могут переупаковываться Excel; сравниваем каждую колонку.
        max_column = max([113] + [READER.coord(ref)[1] for bk in (old, new) for ref in bk.sheets[name]['rawcells']]
                         + [int(x['max']) for bk in (old, new) for x in bk.sheets[name]['cols']])
        for col in range(1, max_column + 1):
            def column(book):
                x = next((v for v in book.sheets[name]['cols'] if int(v['min']) <= col <= int(v['max'])), {})
                return {k: str(Decimal(v).normalize()) if k == 'width' else v for k, v in x.items() if k in ('width', 'hidden', 'customWidth', 'outlineLevel')}
            if column(old) != column(new):
                errors.append({'Kind': 'ColumnDimensions', 'Sheet': name, 'Column': col_name(col), 'Before': column(old), 'After': column(new)})
        old_merges = {':'.join(mapped(name, p) for p in ref.split(':')) for ref in old.sheets[name]['merges']}
        new_merges = set(new.sheets[name]['merges'])
        for ref in sorted(old_merges ^ new_merges):
            if name == 'Results' and all(excluded(name, p) for p in ref.split(':')):
                continue
            errors.append({'Kind': 'MergedRange', 'Sheet': name, 'Range': ref, 'Before': ref in old_merges, 'After': ref in new_merges})

    if set(old_keys) - set(new_keys):
        errors.append({'Kind': 'RemovedSettings', 'Keys': sorted(set(old_keys) - set(new_keys))})
    added_keys = sorted(set(new_keys) - set(old_keys))
    if added_keys not in ([], ['Mesh.CellGeometry']):
        errors.append({'Kind': 'UnexpectedAddedSettings', 'Keys': added_keys})
    if 'Mesh.CellGeometry' in added_keys:
        if new.sheets['Config']['cells'].get('B' + str(new_keys['Mesh.CellGeometry'])) != ('value', 'Adaptive'):
            errors.append({'Kind': 'NewModeDefaultMustBeAdaptive'})
    key_pairs = []
    for key in sorted(set(old_keys) & set(new_keys)):
        for column in ('B', 'C', 'E'):
            a = old.sheets['Config']['cells'].get(column + str(old_keys[key]))
            b = new.sheets['Config']['cells'].get(column + str(new_keys[key]))
            if a != b:
                errors.append({'Kind': 'ConfigKeyValueFormula', 'Key': key, 'Column': column, 'Before': a, 'After': b})
        key_pairs.append({'Key': key, 'BeforeRow': old_keys[key], 'AfterRow': new_keys[key]})

    named = []
    for key, ref in old.names.items():
        if key.startswith('_xlnm.') or '!' not in ref:
            continue
        sheet, address = ref.rsplit('!', 1)
        sheet = sheet.strip("'").replace("''", "'")
        if sheet == 'Справка':
            continue  # Каталог справки действительно пересобран, его старые адреса не утверждаются неизменными.
        expected = re.sub(r'\$[A-Z]+\$\d+', lambda m: '$' + re.sub(r'(\d+)$', r'$\1', mapped(sheet, m.group())), address)
        got = new.names.get(key)
        actual_address = got.rsplit('!', 1)[1] if got and '!' in got else None
        actual_sheet = got.rsplit('!', 1)[0].strip("'").replace("''", "'") if got and '!' in got else None
        if expected != actual_address or actual_sheet != sheet:
            errors.append({'Kind': 'NamedRangeReference', 'Name': key, 'Before': ref, 'ExpectedAddress': expected, 'After': got})
        item = {'Name': key, 'Before': ref, 'After': got}
        points = re.fullmatch(r'\$([A-Z]+)\$(\d+)(?::\$([A-Z]+)\$(\d+))?', address)
        if points and sheet in protected:
            # Именованная матрица сравнивается по каждой старой логической
            # ячейке после частичной вставки, а не по новому физическому адресу.
            r1, c1 = READER.coord(points[1] + points[2])
            r2, c2 = READER.coord((points[3] or points[1]) + (points[4] or points[2]))
            rows_before, rows_after, table_styles = [], [], []
            for row in range(r1, r2 + 1):
                for col in range(c1, c2 + 1):
                    source = col_name(col) + str(row)
                    target = mapped(sheet, source)
                    a, b = old.cell(sheet, source), new.cell(sheet, target)
                    intentional_description = sheet == 'Config' and target in description_cells
                    rows_before.append(None if intentional_description else a[0])
                    rows_after.append(None if intentional_description else b[0])
                    table_styles.append(a[1] == b[1])
            item.update({'RowsBefore': r2 - r1 + 1, 'ColumnsBefore': c2 - c1 + 1,
                         'LogicalCellsCompared': len(rows_before), 'DataAndFormulasPreserved': rows_before == rows_after,
                         'NormalizedStylesPreserved': all(table_styles),
                         'BeforeDataSHA256': digest(json.dumps(rows_before, ensure_ascii=False).encode()),
                         'AfterDataSHA256': digest(json.dumps(rows_after, ensure_ascii=False).encode())})
        named.append(item)
    removed = sorted(set(old.names) - set(new.names))
    if removed:
        errors.append({'Kind': 'RemovedDefinedNames', 'Names': removed})
    added_names = sorted(set(new.names) - set(old.names))
    if set(added_names) - {'rngNDMElementGeometry'}:
        errors.append({'Kind': 'UnexpectedAddedDefinedNames', 'Names': added_names})
    changes = [{'Part': p, 'Change': 'added' if p not in old.parts else 'removed' if p not in new.parts else 'changed',
                'BeforeSHA256': digest(old.parts[p]) if p in old.parts else None,
                'AfterSHA256': digest(new.parts[p]) if p in new.parts else None}
               for p in sorted(set(old.parts) | set(new.parts)) if old.parts.get(p) != new.parts.get(p)]
    return {'Status': 'PASS' if not errors else 'REVIEW_REQUIRED', 'Before': str(old.path), 'After': str(new.path),
            'BeforeSHA256': digest(old.path.read_bytes()), 'AfterSHA256': digest(new.path.read_bytes()),
            'ReadOnlyOpenXML': True, 'ExcelStarted': False, 'ConfigShiftAtoE': shift, 'ConfigHAndLaterShift': 0,
            'ProtectedSheetChecks': counts, 'NamedRanges': named, 'ConfigKeyPairs': key_pairs,
            'AddedNames': added_names, 'AddedSettings': added_keys, 'AllowedCatalogDescriptions': allowed,
            'IntentionalExclusions': ['Справка regenerated from catalogs', 'Results!CR155:DI156 new header', 'Config new Mesh.CellGeometry A:E row'],
            'ZIPParts': {'BeforeCount': len(old.parts), 'AfterCount': len(new.parts), 'UnchangedCount': sum(old.parts.get(p) == new.parts.get(p) for p in set(old.parts) & set(new.parts)), 'Differences': changes},
            'Errors': errors, 'Limitations': ['Drawings, controls, hyperlinks and validation UI require the separate native/layout gates; this check proves stored cell data/formulas, normalized cell styles and dimensions.', 'Formula expressions are compared exactly, so a legitimate changed address is reported for review rather than silently treated as equal.', 'A ZIP-part change is reported, not claimed byte-for-byte preservation of that part.']}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--before', type=Path, required=True)
    parser.add_argument('--after', type=Path, required=True)
    parser.add_argument('--report', type=Path, required=True)
    args = parser.parse_args()
    try:
        report = verify(args.before, args.after)
    except Exception as error:
        report = {'Status': 'ERROR', 'ReadOnlyOpenXML': True, 'Error': repr(error)}
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({k: report[k] for k in ('Status', 'ConfigShiftAtoE', 'ProtectedSheetChecks') if k in report}, ensure_ascii=True))
    print(json.dumps({'Report': str(args.report), 'ErrorCount': len(report.get('Errors', [])), 'ErrorSample': report.get('Errors', [])[:5], 'Error': report.get('Error')}, ensure_ascii=True))
    raise SystemExit(report['Status'] != 'PASS')
