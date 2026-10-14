"""Создаёт частные скрипты этапа, не изменяя готовые доказательства прошлого этапа."""
from pathlib import Path
import importlib.util
import json

directory=Path(__file__).resolve().parent
previous=directory.parent/'RemovePublication_2026-10-14'
old='docs/regression/ExcelGuardGuide_2026-10-13/RemovePublication_2026-10-14'
new='docs/regression/ExcelGuardGuide_2026-10-13/CircleEverySecond_2026-10-14'
for name in ('Prepare.ps1','RunDirected.ps1'):
    (directory/name).write_text((previous/name).read_text(encoding='utf-8-sig').replace(old,new),encoding='utf-8-sig')
root=directory.parents[3]
spec=importlib.util.spec_from_file_location('reader',root/'tools/build_workbook/Verify-SectionRecipeGuide.py')
reader=importlib.util.module_from_spec(spec);spec.loader.exec_module(reader)
book=reader.read(root/'workbook/output/RC_Section_NDM.xlsm')
print('CIRCLE_RANGE:',book['names']['rngCircleGeometry'])
print('NAMES:',json.dumps(book['names'],ensure_ascii=False))
config=book['sheets']['Config']
import re
first,last=map(int,re.findall(r'\$[A-Z]+\$(\d+):\$[A-Z]+\$(\d+)',book['names']['rngCircleGeometry'])[0])
for row in range(first-3,last+6):
    print(row,[(address,cell['value'],cell['formula']) for address,cell in config['cells'].items() if int(re.search(r'\d+',address)[0])==row and len(re.sub(r'\d+','',address))==1])
