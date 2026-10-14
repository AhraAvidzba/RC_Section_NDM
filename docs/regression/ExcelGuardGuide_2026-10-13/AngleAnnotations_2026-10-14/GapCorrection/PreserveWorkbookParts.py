"""Исправление только VBA: каждый прочий ZIP-member сохранён побайтно."""
import hashlib
import json
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

directory=Path(__file__).resolve().parent
baseline,candidate=directory/'Baseline.xlsm',directory/'Candidate.xlsm'
with ZipFile(baseline) as old,ZipFile(candidate) as new:
    assert not any('signature' in name.lower() for name in old.namelist())
    code=new.read('xl/vbaProject.bin')
    temporary=candidate.with_suffix('.parts.tmp')
    with ZipFile(temporary,'w',ZIP_DEFLATED) as output:
        for entry in old.infolist():
            output.writestr(entry,code if entry.filename=='xl/vbaProject.bin' else old.read(entry.filename))
temporary.replace(candidate)
with ZipFile(baseline) as old,ZipFile(candidate) as new:
    changed=[name for name in old.namelist() if old.read(name)!=new.read(name)]
assert changed==['xl/vbaProject.bin'],changed
report={'Passed':True,'ChangedParts':changed,'EveryOtherPartByteIdentical':True,'WorkbookSHA256':hashlib.sha256(candidate.read_bytes()).hexdigest()}
(directory/'PartsPreservation.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps(report))
