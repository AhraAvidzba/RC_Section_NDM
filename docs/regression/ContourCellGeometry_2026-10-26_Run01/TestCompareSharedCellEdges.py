"""ДЛЯ ТЕСТОВ: read-only проверка общих видимых границ соседних/merged ячеек.

Не изменяет строгий verifier. Сравнивает набор непустых аналитических стилей
на двух сторонах одной физической грани; конфликт разных стилей не скрывается.
Внутренние грани одной MergeArea невидимы и учитываются отдельно.
"""
from pathlib import Path
from xml.etree import ElementTree as E
import collections
import importlib.util
import json

ROOT = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location('check', ROOT / 'TestVerifyOpenXmlPreservation.py')
CHECK = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(CHECK)
NS = CHECK.NS


class Edges:
    def __init__(self, book, normalize_black_index=False):
        self.book = book
        styles = E.fromstring(book.parts['xl/styles.xml'])
        self.xfs = list(styles.find('s:cellXfs', NS))
        self.borders = list(styles.find('s:borders', NS))
        self.normalize_black_index = normalize_black_index
        indexed = styles.findall('s:colors/s:indexedColors/s:rgbColor', NS)
        self.index8_rgb = indexed[8].get('rgb') if len(indexed) > 8 else 'FF000000'
        self.merges = {}
        self.columns = {}
        self.rowstyles = {}
        for name, sheet in book.sheets.items():
            self.merges[name] = []
            for area in sheet['merges']:
                parts = area.split(':')
                r1, c1 = CHECK.READER.coord(parts[0])
                r2, c2 = CHECK.READER.coord(parts[-1])
                self.merges[name].append((area, r1, c1, r2, c2))
            xml = E.fromstring(book.parts[sheet['part']])
            self.columns[name] = [(int(n.get('min')), int(n.get('max')), int(n.get('style', '0')))
                                  for n in xml.findall('s:cols/s:col', NS)]
            self.rowstyles[name] = {int(n.get('r')): int(n.get('s', '0')) for n in xml.findall('s:sheetData/s:row', NS)
                                    if n.get('customFormat') == '1'}

    def merge(self, sheet, row, col):
        return next((area for area, r1, c1, r2, c2 in self.merges[sheet] if r1 <= row <= r2 and c1 <= col <= c2), None)

    def side(self, sheet, row, col, side):
        if row < 1 or col < 1:
            return None
        ref = CHECK.col_name(col) + str(row)
        cell = self.book.sheets[sheet]['rawcells'].get(ref)
        if cell is not None and cell.get('s') is not None:
            index = int(cell.get('s'))
        elif row in self.rowstyles[sheet]:
            index = self.rowstyles[sheet][row]
        else:
            index = next((style for first, last, style in self.columns[sheet] if first <= col <= last), 0)
        border = self.borders[int(self.xfs[index].get('borderId', '0'))]
        node = border.find('s:' + side, NS)
        if node is None or node.get('style', 'none') == 'none':
            return None
        children = [CHECK.canon(child) for child in node]
        if self.normalize_black_index:
            # Только явно известный индекс8; auto64/цвет темы не приравниваются RGB.
            for child in children:
                if child[0].endswith('}color'):
                    attributes = dict(child[1])
                    if attributes.get('indexed') == '8':
                        attributes.pop('indexed')
                        attributes['rgb'] = self.index8_rgb
                        child[1] = sorted(attributes.items())
        return {'Attributes': sorted(node.attrib.items()), 'Children': children}

    def edge(self, sheet, direction, row, col):
        if direction == 'H':
            first, second = (row, col, 'bottom'), (row + 1, col, 'top')
        else:
            first, second = (row, col, 'right'), (row, col + 1, 'left')
        merge_a, merge_b = self.merge(sheet, first[0], first[1]), self.merge(sheet, second[0], second[1])
        internal = merge_a is not None and merge_a == merge_b
        a, b = self.side(sheet, *first), self.side(sheet, *second)
        visible = sorted(set(json.dumps(value, sort_keys=True) for value in (a, b) if value is not None))
        return {'InternalMergedEdge': internal, 'VisibleStyles': [] if internal else visible,
                'First': {'Cell': CHECK.col_name(first[1]) + str(first[0]), 'Side': first[2], 'Style': a},
                'Second': {'Cell': CHECK.col_name(second[1]) + str(second[0]), 'Side': second[2], 'Style': b}}


def compare(before, after, report_path, normalize_black_index=False):
    old, new = CHECK.Book(before), CHECK.Book(after)
    a, b = Edges(old, normalize_black_index), Edges(new, normalize_black_index)
    raw = CHECK.verify(before, after)
    keys, changed_cells = set(), []
    attributes_only = 0
    for error in raw['Errors']:
        if error['Kind'] != 'NormalizedCellStyle':
            continue
        name, ref = error['Sheet'], error['BeforeCell']
        row, col = CHECK.READER.coord(ref)
        changed_cells.append([name, ref])
        keys.update(((name, 'H', row - 1, col), (name, 'H', row, col),
                     (name, 'V', row, col - 1), (name, 'V', row, col)))
    records, differences, ambiguous = [], [], []
    for name, direction, row, col in sorted(keys):
        first, second = a.edge(name, direction, row, col), b.edge(name, direction, row, col)
        item = {'Sheet': name, 'Direction': direction, 'GridRow': row, 'GridColumn': col,
                'Before': first, 'After': second, 'SameVisibleEdge': first['VisibleStyles'] == second['VisibleStyles']}
        records.append(item)
        if not item['SameVisibleEdge']:
            differences.append(item)
        if len(first['VisibleStyles']) > 1 or len(second['VisibleStyles']) > 1:
            ambiguous.append(item)
    proof = {'Status': 'PASS_SHARED_EDGE_APPEARANCE' if not differences and not ambiguous else 'REVIEW_REQUIRED',
             'BeforeSHA256': CHECK.digest(old.path.read_bytes()), 'AfterSHA256': CHECK.digest(new.path.read_bytes()),
             'StrictVerifierUnchanged': True, 'StrictStyleDifferences': len(changed_cells),
             'Index8PaletteResolutionEnabled': normalize_black_index,
             'Index8ResolvedRGB': {'Before': a.index8_rgb, 'After': b.index8_rgb},
             'PhysicalEdgesCompared': len(records), 'ChangedPhysicalEdges': len(differences),
             'ConflictingSideStyleEdges': len(ambiguous), 'Differences': differences, 'Ambiguous': ambiguous,
             'AllEdges': records, 'ExcelStarted': False, 'BooksEdited': False,
             'Limitations': ['This checks shared-border appearance, not strict OOXML cell style identity. It does not waive strict preservation failures.',
                             'If the two sides contain conflicting nonempty line styles, this report does not assume an Excel precedence rule.']}
    Path(report_path).write_text(json.dumps(proof, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    return proof


if __name__ == '__main__':
    proof = compare(ROOT / 'ReleaseCandidate05/RC_Section_NDM.xlsm', ROOT / 'HyperlinkNative04/RC_Section_NDM.xlsm',
                    ROOT / 'HyperlinkNative04/SharedEdgeDiagnosis.json')
    print(json.dumps({key: proof[key] for key in ('Status', 'StrictStyleDifferences', 'PhysicalEdgesCompared',
                                                  'ChangedPhysicalEdges', 'ConflictingSideStyleEdges')}, ensure_ascii=True))
    print(json.dumps({'DifferenceSample': proof['Differences'][:2], 'ConflictSample': proof['Ambiguous'][:2]}, ensure_ascii=True))
