"""Убирает случайное повторное включение исторических архивов; исходный код baseline остаётся побайтно прежним."""
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED

D=Path(__file__).resolve().parent
source=D/'BaselineSource.zip';compact=D/'BaselineSource.compact.zip'
old=D/'Evidence.zip';new=D/'Evidence.compact.zip'
assert not compact.exists() and not new.exists()
record={'OriginalSourceArchiveBytes':source.stat().st_size,'OriginalSourceArchiveSHA256':hashlib.sha256(source.read_bytes()).hexdigest()}
hashes={}
with ZipFile(source) as original,ZipFile(compact,'w',ZIP_DEFLATED,compresslevel=9) as output:
    output.comment=original.comment
    for entry in original.infolist():
        if not entry.filename.startswith(('src/','tests/','tools/')):continue
        raw=original.read(entry.filename);hashes[entry.filename]=hashlib.sha256(raw).hexdigest()
        output.writestr(entry,raw,compress_type=ZIP_DEFLATED,compresslevel=9)
with ZipFile(compact) as check:
    assert check.testzip() is None
    assert set(check.namelist())==hashes.keys()
    for name,expected in hashes.items():assert hashlib.sha256(check.read(name)).hexdigest()==expected
record.update({'KeptPrefixes':['src/','tests/','tools/'],'KeptEntries':len(hashes),'AllKeptEntryHashesUnchanged':True,
               'BaselineCommit':'99f90bd75ad687eef95abcda7dd87eb816343c5e',
               'SourceArchiveBytes':compact.stat().st_size,'SourceArchiveSHA256':hashlib.sha256(compact.read_bytes()).hexdigest()})
pruning=(json.dumps(record,ensure_ascii=False,indent=2)+'\n').encode('utf-8')
with ZipFile(old) as original,ZipFile(new,'w',ZIP_DEFLATED,compresslevel=9) as output:
    manifest=json.loads(original.read('ArchiveManifest.json'))
    for entry in original.infolist():
        if entry.filename=='ArchiveManifest.json':continue
        raw=compact.read_bytes() if entry.filename=='BaselineSource.zip' else original.read(entry.filename)
        manifest[entry.filename]={'Bytes':len(raw),'SHA256':hashlib.sha256(raw).hexdigest()}
        output.writestr(entry,raw,compress_type=ZIP_DEFLATED,compresslevel=9)
    manifest['SourceArchivePruning.json']={'Bytes':len(pruning),'SHA256':hashlib.sha256(pruning).hexdigest()}
    output.writestr('SourceArchivePruning.json',pruning)
    output.writestr('ArchiveManifest.json',json.dumps(manifest,ensure_ascii=False,indent=2)+'\n')
with ZipFile(new) as check:
    assert check.testzip() is None
    for name,entry in manifest.items():assert hashlib.sha256(check.read(name)).hexdigest()==entry['SHA256']
for path in [source,compact,old,new]:assert path.resolve().is_relative_to(D)
removed=source.stat().st_size+old.stat().st_size-compact.stat().st_size-new.stat().st_size
compact.replace(source);new.replace(old)
(D/'SourceArchivePruning.json').write_bytes(pruning)
receipt=json.loads((D/'Archive.json').read_text(encoding='utf-8'))
receipt.update({'ArchiveBytes':old.stat().st_size,'ArchiveSHA256':hashlib.sha256(old.read_bytes()).hexdigest(),
                'Entries':len(manifest),'AdditionalDuplicateBytesRemoved':removed})
(D/'Archive.json').write_text(json.dumps(receipt,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'ArchiveBytes':old.stat().st_size,'SourceArchiveBytes':source.stat().st_size,'DuplicateBytesRemoved':removed}))
