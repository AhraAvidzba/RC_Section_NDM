"""Проверяет сохранность книги при адресном исправлении оформления Справки."""
import copy
import hashlib
import importlib.util
import json
import re
import sys
from pathlib import Path
from zipfile import ZipFile

spec = importlib.util.spec_from_file_location('reader', Path(__file__).with_name('Verify-SectionRecipeGuide.py'))
reader = importlib.util.module_from_spec(spec)
spec.loader.exec_module(reader)
S = reader.S
directory = Path(sys.argv[1])
old = reader.read(directory / 'Before.xlsm')
new = reader.read(directory / 'RC_Section_NDM.xlsm')
errors = []
checked = 0


def require(condition, message):
    if not condition:
        errors.append(message)


def without_border(style):
    return style[0], style[1][:2] + style[1][3:], style[2]


def edge(style, side):
    return next(item for item in style[1][2][3] if item[0] == S + side)


require(old['names'] == new['names'], 'Named ranges changed')
require(old['drawings'] == new['drawings'], 'Drawings changed')
with ZipFile(directory/'Before.xlsm') as before_zip, ZipFile(directory/'RC_Section_NDM.xlsm') as after_zip:
    vba_unchanged = before_zip.read('xl/vbaProject.bin') == after_zip.read('xl/vbaProject.bin')
require(vba_unchanged, 'Binary VBA project changed')
require(old['sheets'].keys() == new['sheets'].keys(), 'Sheet inventory changed')
for name, before in old['sheets'].items():
    after = new['sheets'][name]
    for address in before['cells'].keys() | after['cells'].keys():
        a, b = before['cells'].get(address), after['cells'].get(address)
        row = reader.row_of(address)
        checked += 1
        if name == 'Справка' and 1749 <= row <= 1765:
            require(a['formula'] == b['formula'], 'Help formula changed: ' + address)
            continue
        if a == b:
            continue
        if name == 'Справка' and row == 552:
            expected = copy.deepcopy(a)
            parts = list(a['style'][1])
            parts[1] = before['cells']['A492']['style'][1][1]
            expected['style'] = a['style'][0], tuple(parts), a['style'][2]
            require(expected == b, 'Recipe heading changed beyond fill: ' + address)
        elif name == 'Справка' and row in (1748, 1766):
            # Excel хранит общую линию двух строк на одной из её сторон.
            require(all(a[k] == b[k] for k in ('value', 'formula', 'attrs')), 'Boundary content changed: ' + address)
            require(without_border(a['style']) == without_border(b['style']), 'Boundary style changed: ' + address)
            side = 'bottom' if row == 1748 else 'top'
            for other in ('top', 'bottom', 'left', 'right', 'diagonal'):
                if other != side:
                    require(edge(a['style'], other) == edge(b['style'], other), 'Unrelated border changed: ' + address)
            actual = edge(b['style'], side)
            if not dict(actual[1]).get('style'):
                neighbor = re.sub(r'\d+', str(row + (1 if row == 1748 else -1)), address)
                actual = edge(after['cells'][neighbor]['style'], 'top' if row == 1748 else 'bottom')
            require(edge(a['style'], side)[1:] == actual[1:], 'Visible boundary changed: ' + address)
        else:
            require(False, 'Unexpected cell change: ' + name + '!' + address)
    for row in before['rows'].keys() | after['rows'].keys():
        a, b = dict(before['rows'].get(row, {})), dict(after['rows'].get(row, {}))
        if name == 'Справка' and 1749 <= row <= 1765:
            for key in ('ht', 'customHeight', '{http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac}dyDescent'):
                a.pop(key, None); b.pop(key, None)
        require(a == b, 'Row metadata changed: ' + name + '!' + str(row))
    for element in before['document']:
        tag = element.tag
        if tag == S + 'sheetData':
            continue
        previous, following = element, after['document'].find(tag)
        if tag == S + 'mergeCells':
            def merges(parent):
                return {x.get('ref') for x in parent}
            expected = merges(previous)
            if name == 'Справка':
                expected -= {f'A{r}:F{r}' for r in range(1750, 1766) if r != 1755}
            require(expected == merges(following), 'Unexpected merged cells: ' + name)
        else:
            require(reader.canonical(previous) == reader.canonical(following), 'Sheet structure changed: ' + name + ':' + tag)

for first, last in ((1749, 1754), (1755, 1765)):
    a, b = old['sheets']['Справка']['cells'], new['sheets']['Справка']['cells']
    before_text = ' '.join(a.get('A'+str(r), {}).get('value') or '' for r in range(first+1, last+1)).strip()
    after_text = ' '.join(b.get('B'+str(r), {}).get('value') or '' for r in range(first+1, last+1)).strip()
    require(before_text == after_text, 'Help prose changed: ' + str(first))

native = Path(sys.argv[2]).read_text(encoding='utf-8-sig')
dump = Path('workbook/output/VBA_All_Code.txt').read_text(encoding='utf-8-sig')
strip_date = lambda text: re.sub(r'^Date:.*$', 'Date:', text, flags=re.M)
require(strip_date(native) == strip_date(dump), 'Native VBA differs from ALL CODE')
components = len(re.findall(r'^COMPONENT:', native, re.M))
require(components == 149, 'Unexpected VBA component count')
result = {'AllPassed': not errors, 'CellsChecked': checked, 'VbaComponents': components,
          'VbaBinaryUnchanged': vba_unchanged,
          'HelpTextPreserved': not errors, 'NamedRanges': len(old['names']),
          'CandidateSHA256': hashlib.sha256((directory/'RC_Section_NDM.xlsm').read_bytes()).hexdigest(),
          'Errors': errors}
(directory/'Preservation.json').write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
print(json.dumps(result, ensure_ascii=False))
raise SystemExit(bool(errors))
