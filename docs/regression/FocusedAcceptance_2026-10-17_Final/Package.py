"""Сохраняет точный тестовый вход в архиве и убирает только проверенные копии."""
from pathlib import Path
import hashlib,json,zipfile
D=Path(__file__).resolve().parent
ROOT=D.parents[2]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest().upper()
def read(p):return json.loads(p.read_text(encoding='utf-8-sig'))
probe=D/'Probe/RC_Section_NDM.xlsm';expected=read(D/'Probe/Preparation.json')['ProbeWorkbookSHA256']
archive=D/'ProbeWorkbook.zip'
assert probe.resolve().is_relative_to(D.resolve()) and sha(probe)==expected and not archive.exists()
with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED,compresslevel=9) as z:z.write(probe,'RC_Section_NDM.xlsm')
with zipfile.ZipFile(archive) as z:assert hashlib.sha256(z.read('RC_Section_NDM.xlsm')).hexdigest().upper()==expected
original=probe.stat().st_size;probe.unlink()
source=read(D/'Run.json')['SourceSHA256'];deleted=[]
for path in D.glob('NativeCAD/*/RC_Section_NDM.xlsm'):
    assert path.resolve().is_relative_to(D.resolve()) and sha(path)==source
    deleted.append({'Path':str(path.relative_to(D)),'Bytes':path.stat().st_size,'SHA256':source});path.unlink()
diagnostics=D/'Diagnostics.zip';assert not diagnostics.exists()
records=[]
with zipfile.ZipFile(diagnostics,'w',zipfile.ZIP_DEFLATED,compresslevel=9) as z:
    for name in ('FocusedAcceptance_2026-10-17','FocusedAcceptance_2026-10-17_Run01','FocusedAcceptance_2026-10-17_Run02','FocusedAcceptance_2026-10-17_Run03','FocusedAcceptance_2026-10-17_Run04'):
        folder=D.parent/name
        for path in sorted(folder.rglob('*')):
            if path.is_file() and path.suffix.lower() not in ('.xlsm','.zip'):
                relative=str(path.relative_to(D.parent)).replace('\\','/')
                z.write(path,relative);records.append({'Path':relative,'SHA256':sha(path)})
with zipfile.ZipFile(diagnostics) as z:
    for record in records:assert hashlib.sha256(z.read(record['Path'])).hexdigest().upper()==record['SHA256']
result={'ProbeSHA256':expected,'ProbeOriginalBytes':original,'ProbeArchiveBytes':archive.stat().st_size,
        'ProbeArchiveSHA256':sha(archive),'RemovedNativeDuplicates':deleted,
        'DiagnosticFiles':len(records),'DiagnosticsSHA256':sha(diagnostics),'EveryArchivedEntryVerified':True}
(D/'Archive.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps(result,ensure_ascii=False))
