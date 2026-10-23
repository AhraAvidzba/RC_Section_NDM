"""Меняет только оформление справки и зависимые адреса в существующей xlsm.

Пакет редактируется без пересериализации VBA, результатов и входных таблиц.
Карта рисунков получена действующим сборщиком в Test-HelpPictures.ps1.
"""
import argparse
import bisect
import copy
import hashlib
import json
import math
import re
import zipfile
from pathlib import Path

from lxml import etree as E

M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
X = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS = {'m': M, 'x': X, 'a': A}


def update_help(source, target, pictures_path, receipt_path, native_checks=None):
    """Сохраняет прочие ZIP-части побайтно; удаляет только пустые низкие строки."""
    with zipfile.ZipFile(source) as archive:
        original = {name: archive.read(name) for name in archive.namelist()}
        infos = archive.infolist()
    parts = dict(original)
    picture_audit = json.loads(Path(pictures_path).read_text(encoding='utf-8-sig'))
    assert picture_audit['WorkbookSHA256'] == hashlib.sha256(Path(source).read_bytes()).hexdigest().upper(), 'Source differs from audited workbook'
    strings = [''.join(item.itertext()) for item in E.fromstring(parts['xl/sharedStrings.xml'])]
    help_part = 'xl/worksheets/sheet3.xml'
    help_xml = E.fromstring(parts[help_part])
    data = help_xml.find('m:sheetData', NS)
    rows = {int(row.get('r')): row for row in data}

    def text(cell):
        """Читает обычный, shared и inline текст без потери Unicode."""
        if cell is None:
            return ''
        if cell.get('t') == 'inlineStr':
            return ''.join(cell.find('m:is', NS).itertext())
        value = cell.find('m:v', NS)
        if value is None:
            return ''
        return strings[int(value.text)] if cell.get('t') == 's' else value.text or ''

    removed = sorted(n for n, row in rows.items() if float(row.get('ht', '15')) < 5)
    assert all(not any(text(cell) or cell.find('m:f', NS) is not None for cell in rows[n]) for n in removed), 'Nonempty low row'

    def moved(row):
        """Переводит адрес после удаления пустых строк, включая концы диапазонов."""
        return max(1, int(row) - bisect.bisect_right(removed, int(row)))

    def local_ref(value):
        """Переводит только A1-ссылки текущего листа, сохраняя абсолютность."""
        return re.sub(r'(\$?[A-Z]+\$?)(\d+)', lambda m: m[1] + str(moved(m[2])), value)

    def help_ref(value):
        """Переводит адрес после имени Справка, не затрагивая числа в тексте."""
        return re.sub(r"('Справка'!|Справка!)(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)", lambda m: m[1] + local_ref(m[2]), value)

    def set_text(cell, value):
        """Записывает literal inlineStr, не изменяя общую таблицу строк книги."""
        for child in list(cell):
            cell.remove(child)
        cell.set('t', 'inlineStr')
        E.SubElement(E.SubElement(cell, f'{{{M}}}is'), f'{{{M}}}t').text = value

    start = next(n for n, row in rows.items() if text(row.find('m:c', NS)).startswith('General.ReformatResultsEnabled'))
    paragraphs = [text(rows[n].find(f'm:c[@r="A{n}"]', NS)) for n in range(start + 1, start + 8)]
    paragraphs[3:5] = [' '.join(paragraphs[3:5])]
    assert len(paragraphs) == 6 and all(paragraphs)
    template_header = rows[start - 6]
    template_body = rows[start - 5]
    for n in range(start, start + 8):
        row = rows[n]
        for child in list(row):
            row.remove(child)
        model = template_header if n == start else template_body
        for cell in model:
            cloned = copy.deepcopy(cell)
            cloned.set('r', re.sub(r'\d+', str(n), cloned.get('r')))
            for child in list(cloned):
                cloned.remove(child)
            cloned.attrib.pop('t', None)
            row.append(cloned)
        if n == start:
            set_text(row[0], 'General.ReformatResultsEnabled')
            row.set('ht', template_header.get('ht', '12.75'))
        elif n <= start + 6:
            set_text(row[0], str(n - start))
            set_text(row[1], paragraphs[n - start - 1])
            row.set('ht', str(max(1, math.ceil(len(paragraphs[n - start - 1]) / 110)) * 12.75))
        else:
            row.set('ht', '12.75')
    merges = help_xml.find('m:mergeCells', NS)
    for merge in list(merges):
        bounds = [int(n) for n in re.findall(r'\d+', merge.get('ref'))]
        if (start < bounds[0] <= start + 7 or all(n in removed for n in range(bounds[0], bounds[-1] + 1))):
            merges.remove(merge)
    links = help_xml.find('m:hyperlinks', NS)
    for link in list(links):
        if 'Config' in link.get('location', ''):
            links.remove(link)

    picture_map = picture_audit['Expected']
    for picture in picture_map:
        assert picture['TargetBlank'] and picture['TargetMerged']
        rows[picture['Row']].set('ht', str(picture['RowHeight']))
        rows[picture['Row']].set('customHeight', '1')
    for n in removed:
        data.remove(rows[n])
    for row in data:
        row.set('r', str(moved(row.get('r'))))
        row.attrib.pop('hidden', None)
        row.attrib.pop('collapsed', None)  # Макрос открытия задаст штатный вид; геометрия рисунков хранится в раскрытых строках.
        for cell in row:
            cell.set('r', local_ref(cell.get('r')))
    if native_checks:
        checks = json.loads(Path(native_checks).read_text(encoding='utf-8-sig'))
        assert checks['Status'] == 'PASS' and checks['ReformatRow'] == moved(start)
        by_row = {int(row.get('r')): row for row in data}
        for height in checks['ReformatHeights']:
            by_row[height['Row']].set('ht', str(height['Height']))
    for element in help_xml.iter():
        for attr in ('ref', 'sqref', 'activeCell', 'topLeftCell'):
            if element.get(attr):
                element.set(attr, local_ref(element.get(attr)))
    # Адреса ячеек sheetData уже переведены; атрибут r не входит в общий проход.
    merges.set('count', str(len(merges)))
    parts[help_part] = E.tostring(help_xml, xml_declaration=True, encoding='UTF-8', standalone=True)

    drawing_part = 'xl/drawings/drawing3.xml'
    rels = E.fromstring(parts['xl/worksheets/_rels/sheet3.xml.rels'])
    import posixpath
    drawing_part = next(posixpath.normpath(posixpath.join('xl/worksheets', r.get('Target'))) for r in rels if r.get('Type').endswith('/drawing'))
    drawing = E.fromstring(parts[drawing_part])
    expected = {p['Name']: p for p in picture_map}
    heights = {int(r.get('r')): float(r.get('ht', '12.75')) for r in data}
    picture_receipt = []
    for anchor in drawing:
        prop = anchor.find('x:pic/x:nvPicPr/x:cNvPr', NS)
        if prop is None:
            continue
        p = expected[prop.get('name')]
        row = moved(p['Row'])
        padding = (p['RowHeight'] - p['Height']) / 2
        for marker, offset in (('from', padding), ('to', padding + p['Height'])):
            anchor.find(f'x:{marker}/x:row', NS).text = str(row - 1)
            anchor.find(f'x:{marker}/x:rowOff', NS).text = str(round(offset * 12700))
        transform = anchor.find('x:pic/x:spPr/a:xfrm', NS)
        transform.find('a:ext', NS).set('cy', str(round(p['Height'] * 12700)))
        transform.find('a:off', NS).set('y', str(round((sum(heights.get(i, 12.75) for i in range(1, row)) + padding) * 12700)))
        picture_receipt.append({**p, 'Row': row})
    parts[drawing_part] = E.tostring(drawing, xml_declaration=True, encoding='UTF-8', standalone=True)

    for name, content in list(parts.items()):
        if name == 'xl/workbook.xml' or re.fullmatch(r'xl/worksheets/sheet\d+\.xml', name):
            root = E.fromstring(content)
            changed = False
            for element in root.iter():
                if element.get('location'):
                    old = element.get('location'); new = help_ref(old)
                    if old != new:
                        element.set('location', new); changed = True
                if element.tag == f'{{{M}}}definedName' and element.text:
                    old = element.text; new = help_ref(old)
                    if old != new:
                        element.text = new; changed = True
            if changed:
                parts[name] = E.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as archive:
        for info in infos:
            archive.writestr(info, parts[info.filename])
    changed_parts = [name for name in parts if parts[name] != original[name]]
    receipt = {'SourceSHA256': hashlib.sha256(Path(source).read_bytes()).hexdigest().upper(), 'ChangedParts': changed_parts, 'RemovedRows': removed, 'ReformatRow': moved(start), 'Pictures': picture_receipt}
    Path(receipt_path).write_text(json.dumps(receipt, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'removedRows': len(removed), 'pictures': len(picture_receipt), 'changedParts': changed_parts, 'reformatRow': moved(start)}, ensure_ascii=False))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source'); parser.add_argument('target'); parser.add_argument('pictures'); parser.add_argument('receipt')
    parser.add_argument('--native-checks')
    args = parser.parse_args()
    update_help(args.source, args.target, args.pictures, args.receipt, args.native_checks)
