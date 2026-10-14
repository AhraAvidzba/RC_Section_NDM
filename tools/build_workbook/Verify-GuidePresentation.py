"""Проверяет удаление меток, компактность Справки и сохранность остальных данных книги."""
import argparse
import hashlib
import importlib.util
import json
import re
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('reader', Path(__file__).with_name('Verify-SectionRecipeGuide.py'))
reader = importlib.util.module_from_spec(spec)
spec.loader.exec_module(reader)


def components(path):
    text = path.read_text(encoding='utf-8-sig')
    pattern = r'(?ms)^={100}\nCOMPONENT: ([^\n]+)\nTYPE: (\d+)\nLINES: (\d+)\n={100}\n(.*?)(?=^={100}\nCOMPONENT: |\Z)'
    return {m[1]: (m[2], m[3], m[4]) for m in re.finditer(pattern, text)}


def embedded_markers(book, native):
    # Проверяются все XML-части, включая скрытые ячейки, имена и свойства,
    # и реальные тексты всех VBA-компонентов, а не только видимая Справка.
    pattern = re.compile(r'NDM\.RECIPE\.GUIDE/\d|\b[0-9a-f]{64}\b', re.I)
    found = []
    with ZipFile(book) as archive:
        for name in archive.namelist():
            if name.endswith(('.xml', '.rels')):
                if pattern.search(archive.read(name).decode('utf-8-sig')):
                    found.append(name)
    if pattern.search(native.read_text(encoding='utf-8-sig')):
        found.append('native VBA export')
    return found


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--directory', type=Path, required=True)
    parser.add_argument('--book', type=Path)
    parser.add_argument('--dump', type=Path)
    parser.add_argument('--report', type=Path)
    args = parser.parse_args()
    directory = args.directory
    book = args.book or directory/'Candidate.xlsm'
    dump = args.dump or directory/'FinalNativeVBA.txt'
    old, new = reader.read(directory/'Baseline.xlsm'), reader.read(book)
    layout = json.loads((ROOT/'docs/geometry_refactor/SECTION_RECIPE_GUIDE_LAYOUT.json').read_text(encoding='utf-8'))
    first, last = map(int, re.fullmatch(r"'?Справка'?!\$A\$(\d+):\$F\$(\d+)", old['names']['rngSectionRecipeGuide']).groups())
    expected_text = {'A'+str(first+i): row['text'] for i, row in enumerate(layout['rows'])}
    errors, checked, changed_cells, heights = [], 0, [], []

    def require(value, reason):
        if not value:
            errors.append(reason)

    require(len(layout['rows']) == last-first+1, 'Guide range size changed')
    require(old['names'] == new['names'], 'Named ranges changed')
    require(old['drawings'] == new['drawings'], 'Drawings changed')
    require(old['sheets'].keys() == new['sheets'].keys(), 'Sheet inventory changed')
    for sheet, data in old['sheets'].items():
        actual = new['sheets'][sheet]
        for address in data['cells'].keys() | actual['cells'].keys():
            a, b = data['cells'].get(address), actual['cells'].get(address)
            checked += 1
            if sheet == 'Справка' and address in expected_text:
                require(a is not None and b is not None, 'Guide cell missing: '+address)
                require((b['value'] or '') == expected_text[address], 'Guide text mismatch: '+address)
                require(all(a[k] == b[k] for k in ('formula', 'style', 'attrs')), 'Guide formatting changed: '+address)
                if a != b:
                    changed_cells.append(address)
            else:
                require(a == b, 'Cell changed: '+sheet+'!'+address)
        for row in data['rows'].keys() | actual['rows'].keys():
            a, b = dict(data['rows'].get(row, {})), dict(actual['rows'].get(row, {}))
            if sheet == 'Справка' and first <= row <= last:
                before, after = float(a.get('ht', 0)), float(b.get('ht', 0))
                text = layout['rows'][row-first]['text']
                require(actual['cells'].get('A'+str(row), {}).get('value') in (text, None) if not text else
                        actual['cells'].get('A'+str(row), {}).get('value') == text, 'Guide text mismatch: '+str(row))
                require(0 < after <= 60, 'Oversized or invalid guide row: '+str(row))
                require(bool(text) or after <= 10, 'Oversized empty guide row: '+str(row))
                if before != after:
                    heights.append({'Row': row, 'Before': before, 'After': after})
                for key in ('ht', 'customHeight', '{http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac}dyDescent'):
                    a.pop(key, None); b.pop(key, None)
            require(a == b, 'Row metadata changed: '+sheet+'!'+str(row))
        for tag in {e.tag for e in data['document']} | {e.tag for e in actual['document']}:
            if tag in (reader.S+'sheetData', reader.S+'sheetViews'):
                continue
            a, b = data['document'].find(tag), actual['document'].find(tag)
            if tag == reader.S+'mergeCells':
                require({e.get('ref') for e in a} == {e.get('ref') for e in b}, 'Merges changed: '+sheet)
            else:
                require(reader.canonical(a) == reader.canonical(b), 'Sheet structure changed: '+sheet+':'+tag)
    baseline_native, final_native = components(directory/'BaselineNativeVBA.txt'), components(dump)
    native_unchanged = baseline_native == final_native and len(final_native) == 157
    require(native_unchanged, 'Native VBA code or component inventory changed')
    markers = embedded_markers(book, dump)
    require(not markers, 'Embedded technical markers: '+repr(markers))
    sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
    result = {'Passed': not errors, 'CellsChecked': checked, 'NamedRanges': len(old['names']),
              'HelpCellsChanged': sorted(changed_cells), 'HelpRowHeights': heights,
              'NativeComponents': len(final_native), 'NativeCodeUnchanged': native_unchanged,
              'NoEmbeddedMarkers': not markers, 'CandidateSHA256': sha(book), 'NativeSHA256': sha(dump), 'Errors': errors}
    (args.report or directory/'Preservation.json').write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(json.dumps({k: v for k, v in result.items() if k != 'HelpRowHeights'}, ensure_ascii=False))
    if errors:
        raise SystemExit(1)


if __name__ == '__main__':
    main()
