"""ДЛЯ ТЕСТОВ: проверяет сохранённые подписи Unicode, имена и внешние связи."""
from zipfile import ZipFile
from pathlib import Path
import xml.etree.ElementTree as E,json,hashlib
R=Path(r'C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM')
P=Path(r'C:/Users/avidzba/.codex/visualizations/2026/10/12/01a12ece-8192-7141-89b3-f83f1176184c/TorsionPostprocess_2026-11-17')
path=P/'FinalCandidate12/RC_Section_NDM.xlsm'
N={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with ZipFile(path) as z:
    shared=E.fromstring(z.read('xl/sharedStrings.xml'))
    strings=[''.join(x.itertext()) for x in shared]
    book=E.fromstring(z.read('xl/workbook.xml'))
    rels=E.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rid=next(s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id') for s in book.find('m:sheets',N) if s.get('name')=='Results')
    target=next(r.get('Target') for r in rels if r.get('Id')==rid)
    sheet=E.fromstring(z.read(target.lstrip('/') if target.startswith('/') else 'xl/'+target))
    values={}
    for c in sheet.findall('.//m:c',N):
        v=c.find('m:v',N)
        if v is not None:values[c.get('r')]=strings[int(v.text)] if c.get('t')=='s' else v.text
    expected={'BA47':'T (кручение)','BD47':'cot θ','BJ47':'Δσ_T, max','BP47':'σ_sw','BV47':'σ_diag','CB47':'σ_NM','CC47':'Δσ_T','CE47':'σ_NM+T'}
    labels={a:{'Expected':v,'Actual':values.get(a),'Passed':values.get(a)==v} for a,v in expected.items()}
    links=[x for x in z.namelist() if x.startswith('xl/externalLinks/')]
result={'BookSHA256':hashlib.sha256(path.read_bytes()).hexdigest(),'Labels':labels,'ExternalLinks':links,'Passed':all(x['Passed'] for x in labels.values()) and not links}
(P/'SavedLabelsFinal12.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(result,ensure_ascii=False))
assert result['Passed']
