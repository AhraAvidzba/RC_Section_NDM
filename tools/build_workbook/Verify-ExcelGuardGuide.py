"""Сохранность пользовательских данных и независимая сверка изменений VBA/Справки."""
import hashlib
import importlib.util
import json
import re
import sys
from pathlib import Path

spec = importlib.util.spec_from_file_location('reader', Path(__file__).with_name('Verify-SectionRecipeGuide.py'))
reader = importlib.util.module_from_spec(spec)
spec.loader.exec_module(reader)
directory = Path(sys.argv[1])
before, after = directory / 'LiveBefore.xlsm', directory / 'Candidate.xlsm'
old, new = reader.read(before), reader.read(after)
errors, checked, cache_updates = [], 0, []
layout = json.loads(Path('docs/geometry_refactor/SECTION_RECIPE_GUIDE_LAYOUT.json').read_text(encoding='utf-8'))
start, end = map(int, re.fullmatch(r"'?Справка'?!\$A\$(\d+):\$F\$(\d+)", old['names']['rngSectionRecipeGuide']).groups())
expected_changes = {f'A{start+i}': item['text'] for i, item in enumerate(layout['rows'])
                    if old['sheets']['Справка']['cells'].get(f'A{start+i}', {}).get('value', '') != item['text']}
# Пустое значение ячейки и пустая строка обозначают один пустой абзац.
expected_changes = {k: v for k, v in expected_changes.items() if v}


def require(condition, message):
    if not condition:
        errors.append(message)


require(len(expected_changes) == 3, 'Unexpected scope of Help text edits')
require(old['names'] == new['names'], 'Named ranges changed')
require(old['drawings'] == new['drawings'], 'Drawings changed')
require(old['sheets'].keys() == new['sheets'].keys(), 'Sheet inventory changed')
for name, data in old['sheets'].items():
    actual = new['sheets'][name]
    for address in data['cells'].keys() | actual['cells'].keys():
        a, b = data['cells'].get(address), actual['cells'].get(address)
        checked += 1
        if name == 'Справка' and address in expected_changes:
            require(b['value'] == expected_changes[address], 'Guide text mismatch: ' + address)
            require(all(a[key] == b[key] for key in ('formula', 'style', 'attrs')), 'Guide style/formula changed: ' + address)
        elif a != b:
            if a and b and a['formula'] is not None and all(a[k] == b[k] for k in ('formula', 'style', 'attrs')):
                cache_updates.append(name + '!' + address)
            else:
                errors.append('Unexpected cell change: ' + name + '!' + address)
    for row in data['rows'].keys() | actual['rows'].keys():
        a, b = dict(data['rows'].get(row, {})), dict(actual['rows'].get(row, {}))
        if name == 'Справка' and 'A' + str(row) in expected_changes:
            for key in ('ht', 'customHeight', '{http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac}dyDescent'):
                a.pop(key, None); b.pop(key, None)
        require(a == b, 'Row metadata changed: ' + name + '!' + str(row))
    for element in data['document']:
        if element.tag not in (reader.S + 'sheetData', reader.S + 'sheetViews'):
            following = actual['document'].find(element.tag)
            if element.tag == reader.S + 'mergeCells':
                require({e.get('ref') for e in element} == {e.get('ref') for e in following}, 'Merged cells changed: ' + name)
            else:
                require(reader.canonical(element) == reader.canonical(following), 'Sheet structure changed: ' + name + ':' + element.tag)


def components(path):
    text = path.read_text(encoding='utf-8-sig')
    pattern = r'(?ms)^={100}\nCOMPONENT: ([^\n]+)\nTYPE: (\d+)\nLINES: (\d+)\n={100}\n(.*?)(?=^={100}\nCOMPONENT: |\Z)'
    return {m[1]: m[4].strip() for m in re.finditer(pattern, text)}


def normalize_code_case(body):
    # Регистр VBA-идентификаторов VBE меняет по всему проекту при компиляции.
    # Строковые литералы, комментарии, числа и пробелы сравниваются буквально.
    tokens = re.split(r'("(?:[^"\n]|"")*"|\x27[^\n]*)', body or '')
    return ''.join(t if i % 2 else t.lower() for i, t in enumerate(tokens))


a, b = components(directory / 'LiveNative.txt'), components(directory / 'NativeVBA.txt')
changed = sorted(k for k in a.keys() | b.keys() if normalize_code_case(a.get(k)) != normalize_code_case(b.get(k)))
case_only = sorted(k for k in a.keys() & b.keys() if k not in changed and a[k] != b[k])
expected = sorted(['CExcelAppStateGuard', 'CSectionPlotter', 'modWorkbookCalculation', 'modSectionRecipeSetup', 'modSectionRecipeApiGuide', 'modTestWorkbookInterface', 'CTestExcelApplicationState'])
require(changed == expected, 'Unexpected VBA changes: ' + ', '.join(changed))
require(len(a) == 149 and len(b) == 150, 'Unexpected component count')
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
result = {'AllPassed': not errors, 'CellsChecked': checked, 'HelpCellsChanged': sorted(expected_changes),
          'CachedFormulaValuesUpdated': cache_updates, 'NamedRanges': len(old['names']), 'ChangedComponents': changed,
          'WorkbookModuleUnchanged': a.get('ЭтаКнига') == b.get('ЭтаКнига'),
          'VbeIdentifierCaseOnly': case_only,
          'BeforeSHA256': sha(before), 'CandidateSHA256': sha(after), 'NativeSHA256': sha(directory / 'NativeVBA.txt'), 'Errors': errors}
(directory / 'Preservation.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps(result, ensure_ascii=False))
raise SystemExit(bool(errors))
