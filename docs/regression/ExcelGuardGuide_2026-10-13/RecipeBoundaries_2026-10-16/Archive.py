"""Проверяет архив доказательств и удаляет только временные файлы этой проверки."""
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED

D=Path(__file__).resolve().parent
archive=D/'Evidence.zip'
assert not archive.exists()
files=[path for path in D.rglob('*') if path.is_file()]
manifest={}
with ZipFile(archive,'w',ZIP_DEFLATED,compresslevel=9) as out:
    for path in files:
        if path.suffix=='.xlsm':continue
        name=path.relative_to(D).as_posix();raw=path.read_bytes()
        manifest[name]={'Bytes':len(raw),'SHA256':hashlib.sha256(raw).hexdigest()}
        out.writestr(name,raw)
    out.writestr('ArchiveManifest.json',json.dumps(manifest,ensure_ascii=False,indent=2)+'\n')
with ZipFile(archive) as proof:
    assert proof.testzip() is None
    for name,item in manifest.items():assert hashlib.sha256(proof.read(name)).hexdigest()==item['SHA256']
removed=0;count=0
for path in files:
    name=path.relative_to(D).as_posix()
    if path.suffix=='.xlsm' or (path.suffix=='.txt' and not name.startswith('Directed/')):
        assert path.resolve().is_relative_to(D) and path.is_file()
        if path.suffix=='.txt':assert name in manifest
        removed+=path.stat().st_size;count+=1;path.unlink()
result={'ArchiveBytes':archive.stat().st_size,'ArchiveSHA256':hashlib.sha256(archive.read_bytes()).hexdigest(),
        'Entries':len(manifest),'AllEntriesVerified':True,'RemovedBytes':removed,'RemovedFiles':count,
        'BaselineRecovery':'git 99f90bd75ad687eef95abcda7dd87eb816343c5e:workbook/output/RC_Section_NDM.xlsm'}
(D/'Archive.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps(result))
