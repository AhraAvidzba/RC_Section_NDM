"""Архивирует проверочные доказательства и убирает только созданные этой задачей временные книги."""
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parents[2]
D = ROOT/'docs/regression/ExcelGuardGuide_2026-10-13/SegmentNaming_2026-10-16'


def archive():
    target = D/'Evidence.zip'
    assert not target.exists(), 'Preserve the existing archive'
    entries = {}
    files = [path for path in D.rglob('*') if path.is_file()]
    with ZipFile(target,'w',ZIP_DEFLATED,compresslevel=9) as output:
        for path in files:
            name = path.relative_to(D).as_posix()
            # Оставляем одну точную копию последнего пользовательского состояния.
            # Остальные книги восстановимы по исходникам, сохранённым VBA и штатному output.
            if path.suffix == '.xlsm' and name != 'LiveFinalBeforeClose.xlsm':
                continue
            raw = path.read_bytes()
            entries[name] = {'Bytes':len(raw),'SHA256':hashlib.sha256(raw).hexdigest()}
            output.writestr(name,raw)
        output.writestr('ArchiveManifest.json',json.dumps(entries,ensure_ascii=False,indent=2)+'\n')
    with ZipFile(target) as proof:
        assert proof.testzip() is None
        for name,entry in entries.items():
            assert hashlib.sha256(proof.read(name)).hexdigest() == entry['SHA256'], name
    removed, size = [], 0
    for path in files:
        name = path.relative_to(D).as_posix()
        remove = path.suffix == '.xlsm' or (path.suffix == '.txt' and not name.startswith('DirectedAccepted/'))
        if not remove:
            continue
        assert path.resolve().is_relative_to(D.resolve()) and path.is_file()
        if path.suffix == '.txt' or name == 'LiveFinalBeforeClose.xlsm':
            assert name in entries
        size += path.stat().st_size
        path.unlink()
        removed.append(name)
    result = {'Archive':'Evidence.zip','Bytes':target.stat().st_size,'SHA256':hashlib.sha256(target.read_bytes()).hexdigest(),
              'Entries':len(entries),'AllEntryHashesVerified':True,'RemovedFiles':removed,'RemovedBytes':size,
              'UserStateBackup':'Evidence.zip!LiveFinalBeforeClose.xlsm','BaselineRecovery':'git 8f3bd878:workbook/output/RC_Section_NDM.xlsm'}
    (D/'Archive.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({k:v for k,v in result.items() if k != 'RemovedFiles'},ensure_ascii=False))


if __name__ == '__main__':
    archive()
