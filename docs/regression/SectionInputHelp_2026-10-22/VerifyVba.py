"""Check actual VBA and build row heights against the accepted native book."""
from pathlib import Path
import json, re, hashlib, sys

sys.stdout.reconfigure(encoding='utf-8')
HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]


def modules(path):
    result = {}
    for part in re.split(r'(?m)^={100}\nCOMPONENT: ', path.read_text(encoding='utf-8-sig'))[1:]:
        header, body = part.split('=' * 100 + '\n', 1)
        result[header.splitlines()[0]] = body.rstrip()
    return result


before = modules(HERE / 'BaselineVBA.txt')
after = modules(HERE / 'NativeCheckVBA.txt')
assert before.keys() == after.keys()
changed = [name for name in before if before[name] != after[name]]
assert len(changed) == 1, changed
name = changed[0]
old = 'Справка по настройкам листа Config'
new = 'Справка по настройкам Config и листа «Сечения»'
assert before[name].replace(old, new) == after[name]
source = (ROOT / 'tools/build_workbook/WorkbookHelpEvents.vba').read_text(encoding='utf-8-sig').strip()
assert source.casefold() == after[name].strip().casefold()
layout = json.loads((ROOT / 'tools/build_workbook/WorkbookLayout.json').read_text(encoding='utf-8-sig'))
heights = {r['Row']: r['Height'] for s in layout['Sheets'] if s['Name'] == 'Справка' for r in s['Rows']}
edit = json.loads((HERE / 'Native02/NativeEdit.json').read_text(encoding='utf-8-sig'))
for block in edit['Blocks']:
    for row, height in block['Heights'].items():
        assert heights.get(int(row), 12.75) == height
result = {'Passed': True, 'Components': len(after), 'UnchangedComponents': len(after) - 1,
          'ChangedComponent': name, 'OnlyChange': 'Renamed help section string literal',
          'SourceEventEqualityIgnoringVBEIdentifierCase': True, 'BuildHelpHeightsMatchAcceptedBook': True,
          'NativeAllCodeSHA256': hashlib.sha256((HERE / 'NativeCheckVBA.txt').read_bytes()).hexdigest().upper()}
(HERE / 'VbaEquality.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps(result, ensure_ascii=False))
