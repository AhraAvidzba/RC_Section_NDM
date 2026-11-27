"""Проверяет единственную замену текста и собирает копию с неизменными листами."""
from pathlib import Path
import hashlib,importlib.util,json,re
from zipfile import ZipFile
HERE=Path(__file__).resolve().parent
ROOT=next(p for p in HERE.parents if (p/'AGENTS.md').is_file())
spec=importlib.util.spec_from_file_location('package',ROOT/'tools/build_workbook/Package-NormativeCalculators.py')
h=importlib.util.module_from_spec(spec);spec.loader.exec_module(h)
old=h.components(HERE/'BaselineNative.txt');new=h.components(HERE/'VBA_All_Code.txt')
assert old.keys()==new.keys() and len(new)==209
changed=[n for n in old if h.normalized(old[n][1])!=h.normalized(new[n][1])]
assert changed==['modWorkbookCalculation'],changed
pattern=r'(?ms)^Private Function FilteredCombinationMessage\b.*?^End Function$'
source=(ROOT/'src/Excel/modWorkbookCalculation.bas').read_text(encoding='utf-8-sig')
replacement=re.search(pattern,source)[0]
before=old['modWorkbookCalculation'][1]
expected=before.replace(re.search(pattern,before)[0],replacement)
assert h.normalized(expected)==h.normalized(new['modWorkbookCalculation'][1])
assert 'vbCrLf' not in replacement
assert 'N, Mx или My малы и обнулены фильтром, а T не задан:' in replacement
out=HERE/'Accepted/RC_Section_NDM.xlsm';out.parent.mkdir(exist_ok=True)
with ZipFile(HERE/'Source/RC_Section_NDM.xlsm') as base,ZipFile(HERE/'Native/RC_Section_NDM.xlsm') as native,ZipFile(out,'w') as target:
    for entry in base.infolist():target.writestr(entry,native.read(entry.filename) if entry.filename=='xl/vbaProject.bin' else base.read(entry.filename))
with ZipFile(HERE/'Source/RC_Section_NDM.xlsm') as base,ZipFile(out) as target:
    assert base.namelist()==target.namelist() and target.testzip() is None
    changed_parts=[n for n in base.namelist() if base.read(n)!=target.read(n)]
    assert changed_parts==['xl/vbaProject.bin'],changed_parts
result={'Status':'PASS','ChangedComponents':changed,'ChangedPackageParts':changed_parts,'SheetsAndUserDataUnchanged':True,
        'NoExplicitLineBreak':True,'OnlyFormatterTextChanged':True,'Components':209,
        'CandidateSHA256':hashlib.sha256(out.read_bytes()).hexdigest().upper()}
(HERE/'PackageProof.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps(result,ensure_ascii=False))
