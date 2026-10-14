"""Архивирует полные численные/визуальные/native доказательства без тестовых XLSM и DWG."""
from pathlib import Path
import hashlib
import json
from zipfile import ZipFile, ZIP_DEFLATED

ROOT=Path(__file__).resolve().parents[2]
S=ROOT/'docs/regression/ExcelGuardGuide_2026-10-13/Audit02_2026-10-14'

def package():
    groups={
        'UnifiedSnapshots.zip':sorted((S/'UnifiedGeometry').glob('modTest*.txt'))+[S/'UnifiedGeometry/Run.json'],
        'ExcelPlots.zip':sorted((S/'BaselinePreview').glob('*.png'))+sorted((S/'UnifiedPreview').glob('*.png'))+
            sorted((S/'BaselinePreview').glob('modTest*.txt'))+sorted((S/'UnifiedPreview').glob('modTest*.txt'))+
            [S/'BaselinePreview/Run.json',S/'UnifiedPreview/Run.json'],
        'NativeSnapshots.zip':[S/name for name in ('FinalNativeVBA.txt','TestedNativeVBA.txt','PublishedNativeVBA.txt','ExampleNativeVBA.txt')],
    }
    archives={}
    for name,files in groups.items():
        output=S/name
        with ZipFile(output,'w',compression=ZIP_DEFLATED,compresslevel=9) as archive:
            for path in files:archive.write(path,path.relative_to(S).as_posix())
        with ZipFile(output) as archive:assert archive.testzip() is None
        archives[name]={'SHA256':hashlib.sha256(output.read_bytes()).hexdigest(),'Bytes':output.stat().st_size,
                        'Entries':{p.relative_to(S).as_posix():hashlib.sha256(p.read_bytes()).hexdigest() for p in files}}
    result={'Archives':archives,'ContainsTestWorkbooks':False,'ContainsUserDrawings':False}
    (S/'Archives.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({name:info['Bytes'] for name,info in archives.items()}))

if __name__=='__main__':package()
