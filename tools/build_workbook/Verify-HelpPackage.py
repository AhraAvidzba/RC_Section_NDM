"""-------------------- ДЛЯ ТЕСТОВ --------------------
Проверяет границы изменения xlsm, ссылки, текст и все рисунки справки.
"""
import argparse
import io
import json
import re
import zipfile
from pathlib import Path

from lxml import etree as E
from PIL import Image

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'x': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}


# -------------------- ДЛЯ ТЕСТОВ --------------------
def TestHelpPackage(source_path, candidate_path, receipt_path):
    """Сравнивает исходные данные побайтно, исключая только переведённые ссылки."""
    with zipfile.ZipFile(source_path) as z:
        before = {n: z.read(n) for n in z.namelist()}
    with zipfile.ZipFile(candidate_path) as z:
        after = {n: z.read(n) for n in z.namelist()}
    assert before.keys() == after.keys()
    receipt = json.loads(Path(receipt_path).read_text(encoding='utf-8'))
    changed = [n for n in before if before[n] != after[n]]
    assert changed == receipt['ChangedParts'], changed
    assert before['xl/vbaProject.bin'] == after['xl/vbaProject.bin']
    for name in ('xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'):
        roots = [E.fromstring(p[name]) for p in (before, after)]
        for root in roots:
            for link in root.findall('m:hyperlinks/m:hyperlink', NS):
                if 'Справка!' in link.get('location', '') or "'Справка'!" in link.get('location', ''):
                    link.set('location', 'HELP_TARGET')
        assert E.tostring(roots[0]) == E.tostring(roots[1]), name
    strings = [''.join(t.itertext()) for t in E.fromstring(after['xl/sharedStrings.xml'])]

    # -------------------- ДЛЯ ТЕСТОВ --------------------
    def TestCellText(cell):
        """Разрешает текст ячейки для независимой проверки адресов и содержимого."""
        if cell is None:
            return ''
        if cell.get('t') == 'inlineStr':
            return ''.join(cell.find('m:is', NS).itertext())
        value = cell.find('m:v', NS)
        if value is None:
            return ''
        return strings[int(value.text)] if cell.get('t') == 's' else value.text or ''

    help_xml = E.fromstring(after['xl/worksheets/sheet3.xml'])
    rows = help_xml.findall('m:sheetData/m:row', NS)
    assert not any(float(row.get('ht', '15')) < 5 for row in rows)
    cells = {c.get('r'): c for r in rows for c in r}
    link_count = 0
    for name, content in after.items():
        if re.fullmatch(r'xl/worksheets/sheet\d+\.xml', name):
            root = E.fromstring(content)
            for link in root.findall('m:hyperlinks/m:hyperlink', NS):
                address = link.get('location', '')
                match = re.fullmatch(r"'?Справка'?!\$?([A-Z]+)\$?(\d+)", address)
                if match:
                    assert TestCellText(cells.get(match[1] + match[2])), address
                    link_count += 1
    for link in help_xml.findall('m:hyperlinks/m:hyperlink', NS):
        assert 'Config' not in link.get('location', '')
    start = receipt['ReformatRow']
    assert TestCellText(cells[f'A{start}']) == 'General.ReformatResultsEnabled'
    previous = max(int(re.search(r'\d+', ref)[0]) for ref, cell in cells.items() if ref.startswith('A') and TestCellText(cell) == 'Steel.ThreeLine.Es2')
    body_style = cells[f'B{previous+1}'].get('s')
    for i in range(1, 7):
        assert TestCellText(cells[f'A{start+i}']) == str(i)
        assert TestCellText(cells[f'B{start+i}'])
        assert cells[f'B{start+i}'].get('s') == body_style
    drawing_name = next(n for n in changed if n.startswith('xl/drawings/'))
    drawing = E.fromstring(after[drawing_name])
    pictures = drawing.findall('x:twoCellAnchor/x:pic', NS)
    assert len(pictures) == len(receipt['Pictures']) == 76
    for picture in pictures:
        assert int(picture.find('x:spPr/a:xfrm/a:ext', NS).get('cy')) > 0
    for name in after:
        if name.startswith('xl/media/'):
            assert before[name] == after[name]
            Image.open(io.BytesIO(after[name])).verify()
    result = {'Status': 'PASS', 'ChangedParts': changed, 'HelpLinks': link_count, 'Pictures': len(pictures), 'RemovedEmptyRows': len(receipt['RemovedRows']), 'InputCellsAndResultsUnchanged': True, 'VBAUnchanged': True, 'MediaUnchanged': True}
    print(json.dumps(result, ensure_ascii=False))
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source'); parser.add_argument('candidate'); parser.add_argument('receipt'); parser.add_argument('--output')
    args = parser.parse_args()
    result = TestHelpPackage(args.source, args.candidate, args.receipt)
    if args.output:
        Path(args.output).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
