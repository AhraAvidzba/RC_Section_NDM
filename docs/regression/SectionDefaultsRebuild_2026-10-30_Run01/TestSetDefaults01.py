"""ДЛЯ ТЕСТОВ: переносит согласованные значения снимков в каталог сборки, проверяя узкую область изменений."""
import copy
import hashlib
import json
from pathlib import Path

P = Path(__file__).resolve().parent
R = Path('C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM')
source = R / 'tools/build_workbook/GeometryTables.json'
before = json.loads((P / 'GeometryTablesBefore.json').read_text(encoding='utf-8-sig'))
assert json.loads(source.read_text(encoding='utf-8-sig')) == before
after = copy.deepcopy(before)
requested = {
    'PointedSection': {
        'General': {'B': 800, 'H': 250, 'Left.W': 0, 'Left.R1': 100, 'Left.R2': 0,
                    'Right.W': 150, 'Right.R1': 40, 'Right.R2': 50},
        'Main': {'Top': [40, 20, 8, 0, 0, 'IncludeBoth'], 'Bottom': [40, 20, 8, 0, 0, 'IncludeBoth'],
                 'Left': [40, 20, 3, 0, 0, 'ExcludeBoth'], 'Right': [40, 20, 3, 0, 0, 'ExcludeBoth']}},
    'HollowRectangle': {
        'General': {'B': 500, 'H': 800, 'R': 180, 'Bo': 200, 'Ho': 500, 'Ro': 30, 'Xo': 0, 'Yo': 0},
        'Main': {'Top': [40, 20, 5, 0, 0, 'ExcludeBoth'], 'Bottom': [40, 20, 5, 0, 0, 'ExcludeBoth'],
                 'Left': [40, 20, 6, 0, 0, 'IncludeBoth'], 'Right': [40, 20, 6, 0, 0, 'IncludeBoth'],
                 **{k: [40, 20, '—', 0, 0, '—'] for k in ['OpeningTop', 'OpeningBottom', 'OpeningLeft', 'OpeningRight']}}}
}
changes = []
for section in after:
    if section['Type'] not in requested:
        continue
    spec = requested[section['Type']]
    for block in section['Blocks']:
        for row in block['Rows']:
            old = copy.deepcopy(row)
            if block['Key'] == 'General':
                row[2] = spec['General'][row[0]]
            elif block['Key'] == 'Main':
                row[2:8] = spec['Main'][row[0]]
            elif block['Key'] == 'Extra':
                assert row[2:10] == [None, 0, 'EachBar', 'Stacked', None, 0, 'EachBar', 'Stacked']
            for column, (a, b) in enumerate(zip(old, row)):
                if a != b:
                    changes.append({'Type': section['Type'], 'Block': block['Key'], 'Key': row[0],
                                    'Column': block['Keys'][column], 'Before': a, 'After': b})
assert len(changes) == 10, changes
text = json.dumps(after, ensure_ascii=False, indent=2) + '\n'
source.write_text(text, encoding='utf-8-sig')
(P / 'RequestedDefaults.json').write_text(json.dumps(requested, ensure_ascii=False, indent=2), encoding='utf-8')
proof = {'Status': 'BUILD_CATALOG_DEFAULTS_UPDATED', 'ChangedCells': changes,
         'OtherTypesBlocksCommentsAndLayoutsUnchanged': True,
         'CatalogSHA256': hashlib.sha256(source.read_bytes()).hexdigest().upper()}
(P / 'CatalogChanges.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(proof, ensure_ascii=False))
