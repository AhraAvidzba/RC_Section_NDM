"""ДЛЯ ТЕСТОВ: независимо сверяет листы и фигуры после переключения снимков.

Не запускает Excel или расчёт. Читает сохранённые нативным тестом копии,
разрешает индексы строк/стилей и проверяет двоичные файлы отдельным читателем.
"""
import argparse
import hashlib
import importlib.util
import json
import posixpath
import struct
from functools import lru_cache
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as E

M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'m': M}


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestCanonical(node):
    """Сравнивает дерево независимо от порядка атрибутов и имён XML-префиксов."""
    if node is None:
        return None
    return node.tag, tuple(sorted(node.attrib.items())), node.text, tuple(TestCanonical(c) for c in node)


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestReadBook(path):
    """Разрешает строки, формулы и смысл стиля каждой ячейки всех шести листов."""
    with ZipFile(path) as archive:
        parts = {n: archive.read(n) for n in archive.namelist()}
    strings = []
    if 'xl/sharedStrings.xml' in parts:
        strings = [''.join(n.itertext()) for n in E.fromstring(parts['xl/sharedStrings.xml'])]
    styles = E.fromstring(parts['xl/styles.xml'])
    groups = {tag: list(styles.find('m:' + tag, NS)) for tag in ('fonts', 'fills', 'borders', 'cellXfs', 'cellStyleXfs')}
    numbers = {n.get('numFmtId'): n.get('formatCode') for n in styles.findall('m:numFmts/m:numFmt', NS)}

    # -------------------- ДЛЯ ТЕСТОВ --------------------
    @lru_cache(None)
    def TestStyle(index):
        """Заменяет индексы зависимостей их содержимым для независимого сравнения."""
        node = groups['cellXfs'][index]
        attrs = dict(node.attrib)
        refs = []
        for attribute, group in (('fontId', 'fonts'), ('fillId', 'fills'), ('borderId', 'borders'), ('xfId', 'cellStyleXfs')):
            refs.append(TestCanonical(groups[group][int(attrs.pop(attribute, '0'))]))
        number = attrs.pop('numFmtId', '0')
        attrs['format'] = numbers.get(number, 'builtin:' + number)
        return tuple(sorted(attrs.items())), tuple(refs), tuple(TestCanonical(c) for c in node)

    # -------------------- ДЛЯ ТЕСТОВ --------------------
    def TestValue(cell):
        """Сохраняет точные биты Double, текст, ошибки и исходную формулу."""
        kind, node = cell.get('t', 'n'), cell.find('m:v', NS)
        value = None if node is None else node.text
        if kind == 's':
            value = strings[int(value)]
        elif kind == 'inlineStr':
            value = ''.join(cell.find('m:is', NS).itertext())
        elif kind == 'n' and value is not None:
            value = struct.pack('<d', float(value)).hex()
        if kind in ('s', 'inlineStr'):
            kind = 'text'
        if value in (None, ''):
            kind, value = 'empty', ''
        return kind, value, TestCanonical(cell.find('m:f', NS))

    book = E.fromstring(parts['xl/workbook.xml'])
    rels = {n.get('Id'): n.get('Target') for n in E.fromstring(parts['xl/_rels/workbook.xml.rels'])}
    sheets = {}
    for sheet in book.findall('m:sheets/m:sheet', NS):
        target = rels[sheet.get('{' + R + '}id')]
        name = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
        root = E.fromstring(parts[name])
        cells, values, rows = {}, {}, {}
        for row in root.findall('m:sheetData/m:row', NS):
            attrs = {k: v for k, v in row.attrib.items() if k not in ('r', 'spans')}
            if 's' in attrs:
                attrs['s'] = TestStyle(int(attrs['s']))
            rows[row.get('r')] = attrs
            for cell in row:
                value = TestValue(cell)
                cells[cell.get('r')] = value, TestStyle(int(cell.get('s', '0')))
                if value != ('empty', '', None):
                    values[cell.get('r')] = value
        columns = []
        for column in root.findall('m:cols/m:col', NS):
            attrs = dict(column.attrib)
            if 'style' in attrs:
                attrs['style'] = TestStyle(int(attrs['style']))
            columns.append(attrs)
        sections = {tag: tuple(TestCanonical(n) for n in root.findall('m:' + tag, NS))
                    for tag in ('mergeCells', 'dataValidations', 'conditionalFormatting', 'sheetFormatPr', 'hyperlinks', 'sheetProtection')}
        sheets[sheet.get('name')] = {'cells': cells, 'values': values, 'rows': rows, 'columns': columns, 'sections': sections}
    drawings = []
    for name, content in parts.items():
        if not name.startswith('xl/drawings/') or not name.endswith('.xml'):
            continue
        root = E.fromstring(content)
        if root.tag.rsplit('}', 1)[-1] != 'userShapes':
            continue
        for node in root.iter():
            if node.tag.rsplit('}', 1)[-1] == 'cNvPr':
                node.attrib.pop('id', None)  # Номера объектов Office меняются при повторном создании фигур.
            if node.tag.rsplit('}', 1)[-1] == 'creationId':
                node.getparent().remove(node)
        drawings.append(TestCanonical(root))
    assert len(drawings) == 1, 'Expected one actual section drawing'
    return sheets, drawings[0]


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestMultipleSnapshots(directory):
    """Проверяет весь опубликованный контракт: данные Results и неизменность остальных листов."""
    native = json.loads((directory / 'NativeMultipleSnapshots.json').read_text(encoding='utf-8-sig'))
    assert len(native['Cases']) == 5 and len({c['DataHash'] for c in native['Cases']}) == 5
    assert len({c['Shape'] for c in native['Cases']}) == 4
    assert sum(bool(r.get('AfterReopen')) for r in native['Restores']) == 5
    reader_path = Path(__file__).with_name('Compare-SnapshotPayloads.py')
    spec = importlib.util.spec_from_file_location('snapshot_reader_test', reader_path)
    reader = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(reader)
    payloads = []
    for case in native['Cases']:
        path = directory / (case['Name'] + '.ndm')
        assert hashlib.sha256(path.read_bytes()).hexdigest().upper() == case['FileSHA256']
        _, _, tables = reader.TestSnapshotReader(path).TestRead()
        assert len(tables) == 16
        if case['Name'] == 'Круг_трещины':
            assert tables['LoadCase.Annotations'][0] > 1, 'The cracked fixture must contain actual interaction regions'
        payloads.append({'Name': case['Name'], 'Tables': len(tables), 'Cells': sum(height * len(columns) for height, columns in tables.values())})
    records = []
    for restore in native['Restores']:
        before, _ = TestReadBook(directory / restore['Before'])
        after, drawing = TestReadBook(directory / restore['After'])
        expected, expected_drawing = TestReadBook(directory / restore['ExpectedBook'])
        assert before.keys() == after.keys() == expected.keys()
        changed = []
        for name in before:
            if name == 'Results':
                continue
            if before[name] != after[name]:
                changed.append(name)
        assert not changed, ('A different sheet changed', restore['Name'], changed)
        a, b = expected['Results']['values'], after['Results']['values']
        differences = [key for key in a.keys() | b.keys() if a.get(key) != b.get(key)]
        assert not differences, ('Restored Results differ', restore['Name'], differences[:20])
        assert drawing == expected_drawing, ('Actual diagram shapes differ', restore['Name'])
        assert restore['SolverCalls'] == 0 and restore['ReadContextReleased']
        records.append({'Name': restore['Name'], 'ResultsNonemptyCells': len(b), 'OtherSheetsUnchanged': list(before.keys() - {'Results'}),
                        'ActualChartShapesExact': True, 'ActualChartShapeCount': len(drawing[3]),
                        'ActualChartSHA256': hashlib.sha256(repr(drawing).encode('utf-8')).hexdigest().upper(),
                        'AfterReopenButton': bool(restore.get('AfterReopen')), 'SolverCalls': 0})
    assert len({r['ActualChartSHA256'] for r in records}) == 5
    result = {'Status': 'PASS', 'SourceSHA256': native['SourceSHA256'], 'Snapshots': payloads, 'Restores': records,
              'InputValuesFormulasStylesDimensionsValidationUnchanged': True, 'VBAImports': 0}
    (directory / 'MultipleSnapshotsAcceptance.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(result, ensure_ascii=True))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('directory', type=Path)
    TestMultipleSnapshots(parser.parse_args().directory)
