"""Проверяет, что пояснения рецепта не изменили данные книги или выполняемый код."""
import hashlib
import importlib.util
import json
import re
import sys
from pathlib import Path

root = Path.cwd()
spec = importlib.util.spec_from_file_location('reader', root/'tools/build_workbook/Verify-SectionRecipeGuide.py')
reader = importlib.util.module_from_spec(spec)
spec.loader.exec_module(reader)
directory = Path(sys.argv[1])
old, new = reader.read(directory/'LiveBefore.xlsm'), reader.read(directory/'Candidate.xlsm')
layout = json.loads((root/'docs/geometry_refactor/SECTION_RECIPE_GUIDE_LAYOUT.json').read_text(encoding='utf-8'))
footer = layout['rows'][-1]['text']
first_row = int(re.search(r'\$A\$(\d+)', old['names']['rngSectionRecipeGuide'])[1])
footer_address = 'A'+str(first_row+len(layout['rows'])-1)
errors, cells, changes = [], 0, []

def require(condition, message):
    if not condition:
        errors.append(message)

require(old['names'] == new['names'], 'Named ranges changed')
require(old['drawings'] == new['drawings'], 'Drawings changed')
require(old['sheets'].keys() == new['sheets'].keys(), 'Sheet inventory changed')
for name, data in old['sheets'].items():
    actual = new['sheets'][name]
    for address in data['cells'].keys() | actual['cells'].keys():
        a, b = data['cells'].get(address), actual['cells'].get(address)
        cells += 1
        if name == 'Справка' and address == footer_address:
            require(b['value'] == footer, 'Guide footer differs from generated layout')
            require(all(a[k] == b[k] for k in ('formula', 'style', 'attrs')), 'Guide footer formatting changed')
            if a != b:
                changes.append(name+'!'+address)
        else:
            require(a == b, 'Cell changed: '+name+'!'+address)
    require(data['rows'] == actual['rows'], 'Row metadata changed: '+name)
    for element in data['document']:
        if element.tag in (reader.S+'sheetData', reader.S+'sheetViews'):
            continue
        following = actual['document'].find(element.tag)
        if element.tag == reader.S+'mergeCells':
            require({e.get('ref') for e in element} == {e.get('ref') for e in following}, 'Merges changed: '+name)
        else:
            require(reader.canonical(element) == reader.canonical(following), 'Sheet structure changed: '+name+':'+element.tag)

def components(path):
    pattern = r'(?ms)^={100}\nCOMPONENT: ([^\n]+)\nTYPE: (\d+)\nLINES: (\d+)\n={100}\n(.*?)(?=^={100}\nCOMPONENT: |\Z)'
    return {m[1]:m[4].strip() for m in re.finditer(pattern, path.read_text(encoding='utf-8-sig'))}

def normalize(body):
    tokens = re.split(r'("(?:[^"\n]|"")*"|\x27[^\n]*)', body)
    return ''.join(t if i % 2 else t.lower() for i,t in enumerate(tokens))

a, b = components(directory/'LiveNative.txt'), components(directory/'NativeVBA.txt')
changed = sorted(k for k in a.keys() | b.keys() if normalize(a.get(k, '')) != normalize(b.get(k, '')))
require(changed == ['modSectionRecipeApiGuide'], 'Unexpected VBA changes: '+str(changed))
require(len(a) == len(b) == 150, 'Component inventory changed')
require(all(line.lstrip().startswith("'") for line in b['modSectionRecipeApiGuide'].splitlines() if line.strip()), 'Guide contains executable code')
guide = b['modSectionRecipeApiGuide'].splitlines()
example = guide[guide.index("'   МИНИМАЛЬНЫЙ ПРИМЕР")+1:]
example = example[:next(i for i,line in enumerate(example) if 'В CSectionTypeRegistry.Class_Initialize' in line)]
commands = [i for i,line in enumerate(example) if line.startswith("'       ") and not line.startswith("'       '")]
require(len(commands) == 13, 'Expected 13 commands in example')
require(all(example[i+1].startswith("'       ' ") for i in commands), 'Missing immediate explanation')
require(changes == ['Справка!'+footer_address], 'Unexpected worksheet text changes')
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
result = {'Passed': not errors, 'CellsChecked': cells, 'NamedRanges': len(old['names']),
          'ChangedCells': changes, 'ChangedComponents': changed, 'NativeComponents': len(b),
          'ExampleCommandsExplained': len(commands), 'ExecutableCodeUnchanged': changed == ['modSectionRecipeApiGuide'],
          'CandidateSHA256': sha(directory/'Candidate.xlsm'), 'NativeSHA256': sha(directory/'NativeVBA.txt'), 'Errors': errors}
(directory/'Preservation.json').write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
print(json.dumps(result, ensure_ascii=False))
raise SystemExit(bool(errors))
