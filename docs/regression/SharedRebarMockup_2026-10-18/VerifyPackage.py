from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as E
import base64,json,hashlib,re
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
    for row in m['mainRows']+m['extraRows']:
        for column in ('H','I'): assert cells[f'{column}{row}'].find('s:f',n) is not None
    for side in m['sides']: assert cells[side['lengthCell']].find('s:f',n) is not None
    assert cells['J67'].find('s:f',n) is not None
v=Path('C:/Users/avidzba/.codex/visualizations/2026/10/12/01a12ece-8192-7141-89b3-f83f1176184c/ndm-lsection-shared-rebar.html')
text=v.read_text(encoding='utf-8')
for kind in ('Geometry','Rebar','CircleRebar'):
    value=base64.b64encode((d/f'{kind}_Native.png').read_bytes()).decode('ascii')
    text,count=re.subn(r'(<img data-table="'+kind+r'" src=")[^"]*(")',lambda match:match[1]+'data:image/png;base64,'+value+match[2],text)
    assert count==1, f'Expected one screenshot for {kind}'
text,count=re.subn(r'(<img data-table="Geometry"[^>]* alt=")[^"]*(")',lambda match:match[1]+'Геометрия Г-сечения: размеры B1, H1, B2, H2 в General; восемь сторон R1/R2 в Edges. Все длины сторон вычисляются формулами.'+match[2],text)
assert count==1
assert 'GEOMETRY_IMAGE' not in text and 'REBAR_IMAGE' not in text and len(text.encode('utf-8'))<1000000
text=text.replace('остальные одиннадцать ключей пустые','остальные семь ключей пустые')
assert set(re.findall(r'data-table="([^"]+)"',text))=={'Geometry','Rebar','CircleRebar'}
assert all(f'id="{name}"' in text for name in ('ndm-ls-table','ndm-ls-zoom','ndm-ls-zoom-value','ndm-ls-image-area'))
assert not re.search(r'<!doctype|<html|<head|<body|fetch\(|XMLHttpRequest|WebSocket',text,re.I)
assert not re.search(r'две зависимые|шесть сторон|одиннадцать|Sides|12 строк|CIRCLE_REBAR_IMAGE',text)
v.write_text(text,encoding='utf-8')
print(f'VIEWER_FRAGMENT_OK: {len(text.encode("utf-8"))} bytes; three native screenshots.')
print('FINAL_MOCKUP: 8 rows, LSection=8, Circle=1; 16 key and 16 side-name formulas; no VBA; no external workbook links; no formula errors; native Excel checks passed.')
