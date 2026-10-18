from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as E
import base64,json,hashlib
root=Path.cwd(); d=root/'docs/regression/SharedRebarMockup_2026-10-18'
m=json.loads((d/'Manifest.json').read_text(encoding='utf-8')); p=Path(m['file'])
a=json.loads((d/'Acceptance.json').read_text(encoding='utf-8-sig')); assert a['Passed'] and a['SHA256']==hashlib.sha256(p.read_bytes()).hexdigest().upper()
n={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with ZipFile(p) as z:
    assert not any(x.endswith('vbaProject.bin') or x.startswith('xl/externalLinks/') for x in z.namelist())
    for part in ['xl/worksheets/sheet1.xml','xl/worksheets/sheet2.xml']:
        sheet=E.fromstring(z.read(part)); assert not sheet.findall('s:conditionalFormatting',n)
        assert not [c for c in sheet.findall('.//s:c',n) if c.get('t')=='e']
    sheet=E.fromstring(z.read('xl/worksheets/sheet1.xml')); cells={c.get('r'):c for c in sheet.findall('.//s:c',n)}
    for row in m['mainRows']+m['extraRows']: assert cells[f'H{row}'].find('s:f',n) is not None
v=Path('C:/Users/avidzba/.codex/visualizations/2026/10/12/01a12ece-8192-7141-89b3-f83f1176184c/ndm-lsection-shared-rebar.html')
text=v.read_text(encoding='utf-8')
for kind in ('Geometry','Rebar'):
    value=base64.b64encode((d/f'{kind}_Native.png').read_bytes()).decode('ascii')
    text=text.replace(kind.upper()+'_IMAGE','data:image/png;base64,'+value)
assert 'GEOMETRY_IMAGE' not in text and 'REBAR_IMAGE' not in text and len(text.encode('utf-8'))<1000000
v.write_text(text,encoding='utf-8')
print('FINAL_MOCKUP: 12 key formulas; no VBA; no external workbook links; no formula errors; native Excel checks passed.')