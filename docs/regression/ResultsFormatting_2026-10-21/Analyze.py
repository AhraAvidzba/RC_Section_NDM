"""Compare full-command timings and saved Results cells, styles and rich-text runs."""
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as E
import csv
import hashlib
import json
import posixpath
import re
import statistics
import sys

HERE = Path(__file__).resolve().parent
RUN = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else HERE / 'Run02'
MODES = ['Full','Minimal','FullEnd'] if RUN.name.startswith('ProductionBench') else ['Full','Reuse','FullEnd']
NS = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

def canonical(node):
    if node is None:
        return None
    return (node.tag.split('}')[-1], tuple(sorted(node.attrib.items())), node.text or '', tuple(canonical(c) for c in node))

def digest(value):
    return hashlib.sha256(repr(value).encode('utf-8')).hexdigest()

def column_number(text):
    value = 0
    for char in text:
        value = value * 26 + ord(char) - 64
    return value

def address_parts(ref):
    match = re.fullmatch(r'\$?([A-Z]+)\$?(\d+)', ref)
    return int(match[2]), column_number(match[1])

def workbook(path):
    with ZipFile(path) as z:
        book = E.fromstring(z.read('xl/workbook.xml'))
        rels = {r.get('Id'): r.get('Target') for r in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        sheet = next(s for s in book.find('x:sheets', NS) if s.get('name') == 'Results')
        target = rels[sheet.get(R)]
        target = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
        xml = E.fromstring(z.read(target))
        strings = E.fromstring(z.read('xl/sharedStrings.xml')).findall('x:si', NS)
        styles = E.fromstring(z.read('xl/styles.xml'))
        fonts, fills, borders = [list(styles.find('x:' + name, NS)) for name in ['fonts', 'fills', 'borders']]
        formats = {n.get('numFmtId'): n.get('formatCode') for n in styles.findall('x:numFmts/x:numFmt', NS)}
        style_values = []
        parents = list(styles.find('x:cellStyleXfs', NS))
        for xf in styles.find('x:cellXfs', NS):
            value = dict(xf.attrib)
            # Explicit no-fill and inherited no-fill have identical visible properties.
            if value.get('fillId', '0') == '0' and parents[int(value.get('xfId', '0'))].get('fillId', '0') == '0':
                value.pop('applyFill', None)
            for key, array in [('fontId', fonts), ('fillId', fills), ('borderId', borders)]:
                value[key] = canonical(array[int(value.get(key, '0'))])
            value['numFmtId'] = formats.get(value.get('numFmtId'), 'builtin:' + value.get('numFmtId', '0'))
            style_values.append(digest((tuple(sorted(value.items())), tuple(canonical(c) for c in xf))))
        cells = {}
        default = style_values[0]
        for cell in xml.findall('x:sheetData/x:row/x:c', NS):
            ref = cell.get('r')
            value = cell.find('x:v', NS)
            value = value.text if value is not None else ''
            rich = None
            if cell.get('t') == 's':
                si = strings[int(value)]
                value = ''.join(t.text or '' for t in si.iter('{'+NS['x']+'}t'))
                rich = canonical(si) if si.find('x:r', NS) is not None else None
            formula = cell.find('x:f', NS)
            style = style_values[int(cell.get('s', '0'))]
            if value or formula is not None or style != default:
                cells[ref] = (value, formula.text if formula is not None else None, style, rich)
        names = {n.get('name'): n.text for n in book.findall('x:definedNames/x:definedName', NS)}
        summary_ref = names['rngBatchSummary'].split('!')[-1].split(':')[0]
        sr, sc = address_parts(summary_ref)
        ignored = {(sr+3, sc)}
        indexed = {address_parts(ref): value for ref, value in cells.items()}
        for (row, col), value in indexed.items():
            if value[0] == 'RunID':
                for data_row in range(row+1, max(r for r, c in indexed)+1):
                    if not indexed.get((data_row, col), ('',))[0]:
                        break
                    ignored.add((data_row, col))
        structure = {key: canonical(xml.find('x:'+key, NS)) for key in ['cols','sheetFormatPr']}
        default_height = float(xml.find('x:sheetFormatPr',NS).get('defaultRowHeight','15'))
        effective_rows = []
        for row in xml.findall('x:sheetData/x:row',NS):
            attrs = {k:v for k,v in row.attrib.items() if k not in ['r','spans','customHeight'] and not k.endswith('}dyDescent')}
            if float(attrs.get('ht', default_height)) == default_height: attrs.pop('ht',None)
            if attrs.get('hidden') == '0': attrs.pop('hidden')
            if attrs.get('outlineLevel') == '0': attrs.pop('outlineLevel')
            if 's' in attrs:
                attrs['s'] = style_values[int(attrs['s'])]
                if attrs['s'] == default:
                    attrs.pop('s'); attrs.pop('customFormat',None)
            if attrs: effective_rows.append((row.get('r'),tuple(sorted(attrs.items()))))
        structure['rows'] = effective_rows
        structure['merges'] = sorted(m.get('ref') for m in xml.findall('x:mergeCells/x:mergeCell', NS))
        return {'cells': cells, 'default': default, 'ignored': ignored, 'structure': structure}

def compare(before, after):
    changes = {k: [] for k in ['values','styles','richText','structure']}
    counts = {'cells': 0, 'richTextCells': 0, 'ignoredVolatileCells': 0}
    a, b = workbook(before), workbook(after)
    ignored = a['ignored'] | b['ignored']
    for ref in sorted(set(a['cells']) | set(b['cells'])):
        left = a['cells'].get(ref, ('', None, a['default'], None))
        right = b['cells'].get(ref, ('', None, b['default'], None))
        counts['cells'] += 1
        if left[3] or right[3]: counts['richTextCells'] += 1
        if address_parts(ref) not in ignored:
            if left[:2] != right[:2]: changes['values'].append(ref)
        else:
            counts['ignoredVolatileCells'] += 1
        if left[2] != right[2]: changes['styles'].append(ref)
        if left[3] != right[3]: changes['richText'].append(ref)
    for key in a['structure']:
        if a['structure'][key] != b['structure'][key]: changes['structure'].append(key)
    return {**counts, 'differences': {k: len(v) for k,v in changes.items()}, 'examples': {k: v[:20] for k,v in changes.items() if v}}

results = {}
for path in sorted(RUN.glob('*_Plot*.csv')):
    rows = list(csv.DictReader(path.read_text(encoding='utf-8-sig').splitlines()))
    receipt = path.with_suffix('.txt').read_text(encoding='utf-8-sig')
    assert 'FORMATTING_COMPLETE:' in receipt and 'FAIL:' not in receipt, path.name
    assert all(int(r['mismatch']) == 0 for r in rows), path.name
    groups = {}
    for name in MODES:
        values = [r for r in rows if r['variant'] == name and int(r['repetition']) > 0]
        expected = 3 if path.stem.startswith('Dense') else 5
        assert len(values) == expected
        groups[name] = {'samples': len(values), **{f: statistics.median(float(r[f]) for r in values) for f in ['total','readModel','solve','output','plot','compared']}, 'range': [min(float(r['total']) for r in values), max(float(r['total']) for r in values)]}
    for name, values in groups.items():
        values['speedup'] = groups['Full']['total']/values['total']
        values['savedPercent'] = (1-values['total']/groups['Full']['total'])*100
    comparisons = {}
    for name in MODES[1:]:
        comparisons[name] = compare(RUN / (path.stem+'_Full.xlsm'), RUN / (path.stem+'_'+name+'.xlsm'))
    results[path.stem] = {'variants': groups, 'outputComparison': comparisons}
(RUN / 'Results.json').write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding='utf-8')
for name, run in results.items():
    print(name)
    for variant, result in run['variants'].items():
        print(f'  {variant}: {result["total"]:.3f}s, output {result["output"]:.3f}s, x{result["speedup"]:.2f}')
    print({k: v['differences'] for k,v in run['outputComparison'].items()})
