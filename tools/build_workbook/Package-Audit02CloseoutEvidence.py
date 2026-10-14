"""Сохраняет полные baseline/candidate доказательства без приватных XLSM."""
from pathlib import Path
import hashlib
import json
from zipfile import ZipFile,ZIP_DEFLATED

ROOT=Path(__file__).resolve().parents[2]
S=ROOT/'docs/regression/ExcelGuardGuide_2026-10-13/Audit02Closeout_2026-10-14'

def package():
    groups={
        'ModelAndInputSnapshots.zip':[p for name in ('BaselineSnapshots','CandidateSnapshots') for p in sorted((S/name).glob('modTest*.txt'))],
        'ExcelPlots.zip':[p for name in ('BaselinePreview','CandidatePreview') for p in sorted((S/name).glob('*.png'))]+[
            p for name in ('BaselinePreview','CandidatePreview') for p in sorted((S/name).glob('modTest*.txt'))],
        'NativeSnapshots.zip':[S/name for name in ('BaselineNativeVBA.txt','FinalNativeVBA.txt','TestedNativeVBA.txt','PublishedNativeVBA.txt')],
    }
    archives={}
    for name,files in groups.items():
        path=S/name
        with ZipFile(path,'w',compression=ZIP_DEFLATED,compresslevel=9) as z:
            for p in files:z.write(p,p.relative_to(S).as_posix())
        with ZipFile(path) as z:assert z.testzip() is None
        archives[name]={'SHA256':hashlib.sha256(path.read_bytes()).hexdigest(),'Bytes':path.stat().st_size,
                        'Entries':{p.relative_to(S).as_posix():hashlib.sha256(p.read_bytes()).hexdigest() for p in files}}
    (S/'Archives.json').write_text(json.dumps({'Archives':archives,'PrivateWorkbooksIncluded':False},ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({k:v['Bytes'] for k,v in archives.items()}))

if __name__=='__main__':package()
